"""매뉴얼/스펙시트용 화면 캡처 (시뮬레이터 연결 상태로 화면을 띄워 캡처).

사용법(리눅스, Xvfb/ImageMagick 필요):
  xvfb-run -a -s "-screen 0 1920x1080x24" python tools/make_screenshots.py docs/images
  python tools/annotate_screenshots.py docs/images
"""
import json
import os
import subprocess
import sys
import tkinter as tk

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from motor_monitor.constants import MotorType  # noqa: E402
from motor_monitor.plant import Plant  # noqa: E402
from motor_monitor.simulator import FakeGateway  # noqa: E402
from motor_monitor.ui import dialogs  # noqa: E402
from motor_monitor.ui.main_window import MainWindow  # noqa: E402

import tempfile  # noqa: E402

OUT = sys.argv[1]
SCR = tempfile.gettempdir()
os.makedirs(OUT, exist_ok=True)
cfg_path = os.path.join(SCR, "shot_config.ini")
if os.path.exists(cfg_path):
    os.remove(cfg_path)

old = FakeGateway().start()
new = FakeGateway().start()
for uid, rpm in [(5, 950), (6, 1010), (7, 980)]:
    old.add_old(uid, rpm)
for uid in (10, 11, 12, 20):
    new.add_new(uid)

plant = Plant(cfg_path)
plant.load()
h = plant.houses["H1R1E"]
for mid in (5, 6, 7, 8):
    h.add_motor(mid, MotorType.OLD_MOTOR, True)
h.add_motor(9, MotorType.OLD_MOTOR, False)
for mid in (10, 11, 12):
    h.add_motor(mid, MotorType.NEW_MOTOR, True)
plant.configs["H1R1E"] = plant.configs["H1R1E"].copy(
    source_ip=old.host, source_port=old.port, target_ip=new.host, target_port=new.port,
    source_manual=False, source_line="H1R1E", source_id=5)
h2 = plant.houses["H1R1O"]
h2.add_motor(20, MotorType.NEW_MOTOR, True)
h2.add_motor(21, MotorType.NEW_MOTOR, False)
plant.configs["H1R1O"] = plant.configs["H1R1O"].copy(
    source_ip=old.host, source_port=old.port, target_ip=new.host, target_port=new.port,
    source_manual=True, manual_rpm=800)
for code in ("H1R2E", "H1R3O", "H1R4E", "H2R1E", "H2R2O", "H2R4E"):
    plant.houses[code].add_motor(1, MotorType.OLD_MOTOR, True)
    plant.houses[code].add_motor(2, MotorType.NEW_MOTOR, True)
plant.save()

W, H = (int(v) for v in os.environ.get("SHOT_SIZE", "1920x1040").split("x"))
root = tk.Tk()
win = MainWindow(root, plant)
root.geometry(f"{W}x{H}+0+0")


def snap(name, window_id=None):
    root.update()
    target = "root"
    subprocess.run(["import", "-window", target, "-crop", f"{W}x{H}+0+0" if name.startswith("main") else "+0+0",
                    os.path.join(OUT, name + ".png")], check=True)


def start():
    plant.start_house("H1")
    win._show_ctrl("H1", True)
    plant.start_house("H2")
    win._show_ctrl("H2", True)
    win.update_network_status()


MARKS = {}


def bbox(w, ox=0, oy=0):
    w.update_idletasks()
    return [w.winfo_rootx() - ox, w.winfo_rooty() - oy, w.winfo_width(), w.winfo_height()]


def main_shot():
    snap("main_running")
    lp = win.line_panels["H1R1E"]
    zone = lp.master.master
    items = [
        ("1", win.H1_StopBtn), ("2", win.H2_StopBtn),
        ("3", zone),
        ("4", lp), ("5", lp.rec_SOURCE_NETSTATUS.master), ("6", lp.rec_TARGET_NETSTATUS.master),
        ("7", lp.rec_SID_NETSTATUS.master), ("8", lp.Btn_Settings), ("9", lp.Btn_Add),
        ("10", lp.LV.header), ("11", lp.LV.canvas),
    ]
    MARKS["main_running"] = [(n, bbox(w)) for n, w in items]
    root.after(100, line_dialog)


def crop_dialog(name, dlg):
    root.update()
    dlg.update()
    x, y = dlg.winfo_rootx(), dlg.winfo_rooty()
    w, hgt = dlg.winfo_width(), dlg.winfo_height()
    path = os.path.join(OUT, name + ".png")
    subprocess.run(["import", "-window", "root", "-crop", f"{w}x{hgt}+{x}+{y}", path], check=True)
    subprocess.run(["convert", path, "+repage", path], check=True)


def line_dialog():
    d = dialogs.LineSettingsDialog(root, plant, "H1R1E", on_saved=win.update_header_box)

    def step():
        crop_dialog("dlg_line_ref", d)
        ox, oy = d.winfo_rootx(), d.winfo_rooty()
        MARKS["dlg_line_ref"] = [(n, bbox(w, ox, oy)) for n, w in [
            ("1", d.header), ("2", d.lbl_house), ("3", d.txtSIP.master), ("4", d.txtSPORT.master),
            ("5", d.txtTIP.master), ("6", d.txtTPORT.master), ("7", d.rb_source_manual), ("8", d.rb_source_ref),
            ("9", d.cb_lines), ("10", d.cb_ids), ("11", d.BtnSave), ("12", d.BtnCancel)]]
        d.source_mode.set("manual")
        d._on_mode()
        root.after(200, lambda: (crop_dialog("dlg_line_manual", d), MARKS.__setitem__("dlg_line_manual", [(n, bbox(w, d.winfo_rootx(), d.winfo_rooty())) for n, w in [("13", d.txtMRPM.master), ("14", d.lbl_rpm_range)]]), d.destroy(), root.after(100, motor_add)))
    root.after(500, step)
    d.show_modal()


def motor_add():
    d = dialogs.MotorDialog(root, plant, win.images, "H1R1E", "ADD")
    def m():
        ox, oy = d.winfo_rootx(), d.winfo_rooty()
        MARKS["dlg_motor_add"] = [(n, bbox(w, ox, oy)) for n, w in [
            ("1", d.header), ("2", d.lbl_house), ("3", d.rb_oldmotor), ("4", d.rb_newmotor), ("5", d.txtID.master),
            ("6", d.rb_use), ("7", d.rb_nouse), ("8", d.BtnSave), ("9", d.BtnCancel)]]
    root.after(500, lambda: (crop_dialog("dlg_motor_add", d), m(), d.destroy(), root.after(100, motor_mod)))
    d.show_modal()


def motor_mod():
    d = dialogs.MotorDialog(root, plant, win.images, "H1R1E", "MOD", mod_dkey="0_5")
    def m():
        ox, oy = d.winfo_rootx(), d.winfo_rooty()
        MARKS["dlg_motor_mod"] = [(n, bbox(w, ox, oy)) for n, w in [
            ("10", d.mod_type_line), ("11", d.lblID), ("12", d.BtnDelete)]]
    root.after(500, lambda: (crop_dialog("dlg_motor_mod", d), m(), d.destroy(), root.after(100, stopped)))
    d.show_modal()


def stopped():
    plant.stop_house("H1")
    win._show_ctrl("H1", False)
    plant.stop_house("H2")
    win._show_ctrl("H2", False)
    win.update_network_status()
    root.after(300, lambda: (snap("main_stopped"), finish()))


def finish():
    print("RPM written to new motors:", {u: r[0x16] for u, r in new.new_regs.items()})
    json.dump(MARKS, open(os.path.join(SCR, "marks.json"), "w"))
    plant.shutdown()
    old.stop()
    new.stop()
    root.destroy()


root.after(500, start)
root.after(8000, main_shot)
root.mainloop()
