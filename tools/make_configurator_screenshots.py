"""매뉴얼/스펙시트용 ECblue configurator 화면 캡처 (시뮬레이터에 연결, 예시 기록 사용).

사용법(리눅스, Xvfb/ImageMagick 필요):
  xvfb-run -a -s "-screen 0 1280x800x24" python tools/make_configurator_screenshots.py docs/images
결과: configurator.png, configurator_annotated.png (구성요소 번호 = 매뉴얼 12장 표)
"""
import os
import subprocess
import sys
import tempfile
import time
import tkinter as tk

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ecblue_configurator as cfg  # noqa: E402
from motor_monitor.simulator import FakeGateway  # noqa: E402

OUT = sys.argv[1]
W, H = 1120, 720
os.makedirs(OUT, exist_ok=True)
tmp = tempfile.mkdtemp()

# 예시 장치: 통신 속도 제어(h4=0x0101) 770 RPM, 최고속도 1400, 통신 끊김 시 60초 뒤 770 RPM 자동 운전
gw = FakeGateway().start()
gw.add_ecblue(180, analog_rpm=640, h4=0x0101)
h = gw.ecb[180]["h"]
h[1], h[2], h[5], h[8], h[17] = 1, 770, 770, 1400, (60 << 8) | 2
cfg.PORT = gw.port
cfg.LOG_DIR = os.path.join(tmp, "logs")
cfg.ms.REQUEST_GAP = 0

root = tk.Tk()
root.geometry(f"{W}x{H}+0+0")
app = cfg.ConfiguratorApp(root, settings_file=os.path.join(tmp, "ecblue_configurator.json"))
app.cb_ip.set(gw.host)
app.cb_id.set("180")
app.on_connect()
end = time.time() + 5
while time.time() < end and app.dev is None:
    root.update()
    time.sleep(0.02)

# 지난 1시간 기록 (30초마다): 770 근처, 중간에 700 으로 지정, 잠깐 응답 없음
now = time.time()
app.history.clear()
for k in range(120, 0, -1):
    s = 700 if 40 <= k < 75 else 770
    v = None if k in (90, 91) else s + ((k * 7) % 9) - 4
    app.history.append((now - k * 30, v, None if v is None else s))
app.add_sample(now, 768, 770)
app.tree.selection_set("h17")
app.tree.see("h4")
app._update_now()
app._update_check_label()
root.update()
app.draw_graph()
app._hover_x = int(app.canvas.winfo_width() * 0.55)
app._render_hover()
root.update()
time.sleep(0.3)
root.update()


def box(*widgets):
    """위젯들을 감싸는 사각형 (창 기준 x, y, w, h)."""
    rx, ry = root.winfo_rootx(), root.winfo_rooty()
    xs, ys, xe, ye = [], [], [], []
    for w in widgets:
        w.update_idletasks()
        x, y = w.winfo_rootx() - rx, w.winfo_rooty() - ry
        xs.append(x)
        ys.append(y)
        xe.append(x + w.winfo_width())
        ye.append(y + w.winfo_height())
    return min(xs), min(ys), max(xe) - min(xs), max(ye) - min(ys)


MARKS = [
    ("1", box(app.cb_ip)), ("2", box(app.cb_id)), ("3", box(app.btn_conn)), ("4", box(app.btn_disc)),
    ("5", box(app.dot, app.lbl_status)), ("6", box(app.btn_read)), ("7", box(app.tree)), ("8", box(app.lbl_detail)),
    ("9", box(app.range_buttons["10분"], app.range_buttons["24시간"])), ("10", box(app.btn_fast)),
    ("11", box(app.lbl_check)), ("12", box(app.canvas)), ("13", box(app.lbl_now)),
    ("14", box(app.ent_rpm, app.btn_restore)), ("15", box(app.lbl_target.master)),
    ("16", box(app.lbl_wd)), ("17", box(app.ent_wd_sec.master, app.btn_wd_off)), ("18", box(app.txt)),
]

plain = os.path.join(OUT, "configurator.png")
subprocess.run(["import", "-window", "root", "-crop", f"{W}x{H}+0+0", plain], check=True)
subprocess.run(["convert", plain, "+repage", plain], check=True)
cmd = ["convert", plain, "-fill", "none", "-stroke", "#e0115f", "-strokewidth", "2"]
for n, (x, y, w, hh) in MARKS:
    cmd += ["-draw", f"rectangle {x},{y} {x + w - 1},{y + hh - 1}"]
for n, (x, y, w, hh) in MARKS:
    cx, cy, r = x + 2, y + 2, 9
    cmd += ["-stroke", "none", "-fill", "#e0115f", "-draw", f"circle {cx},{cy} {cx + r},{cy}",
            "-fill", "white", "-font", "DejaVu-Sans-Bold", "-pointsize", "11",
            "-draw", f"text {cx - (3 if len(n) == 1 else 7)},{cy + 4} '{n}'"]
cmd.append(os.path.join(OUT, "configurator_annotated.png"))
subprocess.run(cmd, check=True)
app.shutdown()
root.destroy()
gw.stop()
print("ok", plain)
