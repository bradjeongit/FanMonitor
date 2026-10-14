"""매뉴얼/스펙시트용 화면 캡처 (시뮬레이터에 연결한 상태로 화면을 띄워 캡처, 예시 데이터 사용).

사용법(리눅스, Xvfb/ImageMagick 필요):
  xvfb-run -a -s "-screen 0 1920x1080x24" python tools/make_screenshots.py docs/images
  python tools/annotate_screenshots.py docs/images
"""
import json
import os
import subprocess
import sys
import tempfile
import tkinter as tk

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from motor_monitor import constants as C  # noqa: E402
from motor_monitor.constants import MotorStatus, MotorType  # noqa: E402
from motor_monitor.plant import Plant  # noqa: E402
from motor_monitor.simulator import FakeGateway  # noqa: E402
from motor_monitor.ui import dialogs  # noqa: E402
from motor_monitor.ui.main_window import MainWindow  # noqa: E402

OUT = sys.argv[1]
SCR = tempfile.gettempdir()
os.makedirs(OUT, exist_ok=True)
cfg_path = os.path.join(SCR, "shot_config.dll")
if os.path.exists(cfg_path):
    os.remove(cfg_path)
C.SEND_INTERVAL = 0.05
C.CYCLE_INTERVAL = 0.1

# 예시 구성: 구역마다 기존모터 1~2개(참조용) + 신규모터 여러 개
ZONES = {
    "H1R1": ([(2, 770), (4, 770)], [1]),
    "H1R2": ([(66, 770), (68, 770), (70, 769)], [1, 3, 5, 7]),
    "H1R3": ([], [1, 2, 3, 5, 8, 13]),
    "H1R4": ([], list(range(1, 15))),
    "H1R5": ([(126, 770), (132, 770)], [1, 2, 4, 5, 6, 7]),
    "H1R6": ([(183, 770), (193, 769)], list(range(1, 12))),
    "H1R7": ([(155, 770), (159, 769)], [1, 2, 3, 4, 8, 10, 11, 12, 13]),
    "H1R8": ([(211, 769), (217, 770)], [1, 3, 4, 6, 7, 8]),
}
REF = {"H1R1": ("H1R1", 2), "H1R2": ("H1R2", 66), "H1R3": ("H1R1", 2), "H1R4": ("H1R2", 66),
       "H1R5": ("H1R5", 126), "H1R6": ("H1R6", 193), "H1R7": ("H1R7", 155), "H1R8": ("H1R8", 211)}
LIMITS = {("H1R3", 13): 500, ("H1R4", 12): 450, ("H1R5", 4): 400, ("H1R7", 11): 400}

old_gw, new_gw = {}, {}
plant = Plant(cfg_path)
plant.load()
for code, (olds, news) in ZONES.items():
    og, ng = FakeGateway().start(), FakeGateway().start()
    old_gw[code], new_gw[code] = og, ng
    h = plant.houses[code]
    for mid, rpm in olds:
        og.add_old(mid, rpm)
        h.add_motor(mid, MotorType.OLD_MOTOR, True)
    for mid in news:
        ng.add_new(mid)
        lim = LIMITS.get((code, mid))
        h.add_motor(mid, MotorType.NEW_MOTOR, True, lim is not None, lim or 0)
    line, sid = REF[code]
    plant.configs[code] = plant.configs[code].copy(
        source_ip=og.host, source_port=og.port, target_ip=ng.host, target_port=ng.port,
        source_manual=False, source_line=line, source_id=sid)
plant.houses["H1R8"].add_motor(9, MotorType.NEW_MOTOR, False)     # 중지 예시
plant.save()

W, H = (int(v) for v in os.environ.get("SHOT_SIZE", "1920x1040").split("x"))
root = tk.Tk()
win = MainWindow(root, plant)
root.geometry(f"{W}x{H}+0+0")
MARKS = {}


def bbox(w, ox=0, oy=0):
    w.update_idletasks()
    return [w.winfo_rootx() - ox, w.winfo_rooty() - oy, w.winfo_width(), w.winfo_height()]


def snap(name):
    root.update()
    subprocess.run(["import", "-window", "root", "-crop", f"{W}x{H}+0+0", os.path.join(OUT, name + ".png")],
                   check=True)
    subprocess.run(["convert", os.path.join(OUT, name + ".png"), "+repage", os.path.join(OUT, name + ".png")],
                   check=True)


def crop_dialog(name, dlg):
    root.update()
    dlg.update()
    x, y = dlg.winfo_rootx(), dlg.winfo_rooty()
    w, hgt = dlg.winfo_width(), dlg.winfo_height()
    path = os.path.join(OUT, name + ".png")
    subprocess.run(["import", "-window", "root", "-crop", f"{w}x{hgt}+{x}+{y}", path], check=True)
    subprocess.run(["convert", path, "+repage", path], check=True)


def start():
    for code, zp in win.zone_panels.items():
        if code != "H1R8":
            plant.start_line(code)
            zp.show_running(True)
    win.update_network_status()


def demo_states():
    # 상태 예시: 이상기록(빨간 !), 이상
    plant.update_all_status()
    for code, mid, had_err, status in [("H1R4", 9, True, None), ("H1R6", 5, True, None), ("H1R7", 2, None, MotorStatus.ERR)]:
        h = plant.houses[code]
        with h._lock:
            m = h._motors[f"1_{mid}"]
            if had_err:
                m.had_err = True
            if status is not None:
                m.status = status
            h._touch()
    win.update_network_status()


def main_shot():
    snap("main_running")
    zp = win.zone_panels["H1R1"]
    items = [("1", zp.StopBtn), ("2", zp), ("3", zp.rec_SOURCE_NETSTATUS.master),
             ("4", zp.rec_TARGET_NETSTATUS.master), ("5", zp.rec_SID_NETSTATUS.master),
             ("6", zp.Btn_Settings), ("7", zp.Btn_Add), ("8", zp.LV.header), ("9", zp.LV.canvas)]
    MARKS["main_running"] = [(n, bbox(w)) for n, w in items]
    root.after(100, line_dialog)


def line_dialog():
    d = dialogs.LineSettingsDialog(root, plant, "H1R3", on_saved=win.update_header_box)

    def step():
        crop_dialog("dlg_line_ref", d)
        ox, oy = d.winfo_rootx(), d.winfo_rooty()
        MARKS["dlg_line_ref"] = [(n, bbox(w, ox, oy)) for n, w in [
            ("1", d.header), ("2", d.lbl_house), ("3", d.txtSIP.master), ("4", d.txtSPORT.master),
            ("5", d.txtTIP.master), ("6", d.txtTPORT.master), ("7", d.rb_source_manual), ("8", d.rb_source_ref),
            ("9", d.cb_lines), ("10", d.cb_ids), ("11", d.BtnSave), ("12", d.BtnCancel)]]
        d.source_mode.set("manual")
        d._on_mode()

        def manual():
            crop_dialog("dlg_line_manual", d)
            MARKS["dlg_line_manual"] = [(n, bbox(w, d.winfo_rootx(), d.winfo_rooty()))
                                        for n, w in [("13", d.txtMRPM.master), ("14", d.lbl_rpm_range)]]
            d.destroy()
            root.after(100, motor_add)
        root.after(200, manual)
    root.after(500, step)
    d.show_modal()


def motor_add():
    d = dialogs.MotorDialog(root, plant, win.images, "H1R3", "ADD")
    d.motor_type_var.set(1)
    d._on_type()

    def m():
        crop_dialog("dlg_motor_add", d)
        ox, oy = d.winfo_rootx(), d.winfo_rooty()
        MARKS["dlg_motor_add"] = [(n, bbox(w, ox, oy)) for n, w in [
            ("1", d.header), ("2", d.lbl_house), ("3", d.rb_oldmotor), ("4", d.rb_newmotor), ("5", d.txtID.master),
            ("6", d.rb_use), ("7", d.rb_nouse), ("8", d.cbUSELIMIT), ("9", d.txtLIMITRPM.master),
            ("10", d.BtnSave), ("11", d.BtnCancel)]]
        d.destroy()
        root.after(100, motor_mod)
    root.after(500, m)
    d.show_modal()


def motor_mod():
    d = dialogs.MotorDialog(root, plant, win.images, "H1R3", "MOD", mod_dkey="1_13")

    def m():
        crop_dialog("dlg_motor_mod", d)
        ox, oy = d.winfo_rootx(), d.winfo_rooty()
        MARKS["dlg_motor_mod"] = [(n, bbox(w, ox, oy)) for n, w in [
            ("12", d.mod_type_line), ("13", d.lblID), ("14", d.BtnDelete)]]
        d.destroy()
        root.after(100, finish)
    root.after(500, m)
    d.show_modal()


def finish():
    json.dump(MARKS, open(os.path.join(SCR, "marks.json"), "w"))
    plant.shutdown()
    for g in list(old_gw.values()) + list(new_gw.values()):
        g.stop()
    root.destroy()


root.after(300, start)
root.after(5000, demo_states)
root.after(5600, main_shot)
root.mainloop()
