"""운영 프로그램 요청 시험 (2.6.0):
1·2 새 팬 기본값 - 구역은 수동 770 RPM, 기존모터제어는 수동 지정 70%(770 RPM)
4 신규모터 아이콘 개나리색 (색조는 test_zone_requests.test_ui_icons_colors, 여기서는 주황이 남지 않고 L 은 빨강)
5 속도를 보내면 1초 동안 붉은 팬 그림 + 붉은 행 (test_speed_requests.test_sent_motor_icon_and_row_turn_red_for_one_second)
6 상태 칸의 수동/참조 표시 (상태 그림과 숫자 사이)."""
import colorsys
import os

import pytest

from motor_monitor import config as cfgmod
from motor_monitor.constants import (CTRL_MODE_MANUAL, CTRL_MODE_REF, DEFAULT_CTRL_RPM, DEFAULT_MANUAL_RPM,
                                     MotorStatus, MotorType)
from motor_monitor.house import Motor, speed_source
from motor_monitor.plant import Plant


@pytest.fixture
def root():
    tk = pytest.importorskip("tkinter")
    try:
        r = tk.Tk()
    except tk.TclError:
        pytest.skip("디스플레이 없음")
    r.geometry("1000x560+300+200")
    r.update()
    yield r
    r.destroy()


def _plant(tmp_path):
    p = Plant(str(tmp_path / "Config.dll"))
    p.load()
    return p


# ---------------------------------------------------------------- 1·2 새 팬 기본값
def test_default_speeds_are_770():
    assert (DEFAULT_MANUAL_RPM, DEFAULT_CTRL_RPM) == (770, 770)          # 기존모터제어 70% x 1100 = 770


def test_zone_without_values_starts_manual_770():
    configs, _ = cfgmod.config_from_text("")                            # 설정 파일 없음
    assert all(c.source_manual and c.manual_rpm == 770 for c in configs.values())
    text = "[LINE_CONFIG]\r\nH1R1_SOURCE_MANUAL=0\r\nH1R2_MANUAL_RPM=0\r\n"      # 적힌 값은 그대로
    configs, _ = cfgmod.config_from_text(text)
    assert (configs["H1R1"].source_manual, configs["H1R1"].manual_rpm) == (False, 770)
    assert (configs["H1R2"].source_manual, configs["H1R2"].manual_rpm) == (True, 0)


def test_ctrl_add_dialog_defaults_to_manual_70_percent(root, tmp_path, monkeypatch):
    from motor_monitor.ui import dialogs
    from motor_monitor.ui import theme as T
    msgs = []
    monkeypatch.setattr(dialogs.BaseDialog, "error", lambda self, m, t: msgs.append(m))
    monkeypatch.setattr(dialogs.BaseDialog, "ask", lambda self, m, t: True)
    p = _plant(tmp_path)
    d = dialogs.MotorDialog(root, p, T.Images(root), "H1R7", "ADD")
    d.motor_type_var.set(int(MotorType.OLD_CTRL))
    d._on_type()
    assert d.ctrl_mode_var.get() == CTRL_MODE_MANUAL and d.txtSETRPM.get() == "770"
    assert d.lbl_set_pct.cget("text") == "RPM (0~1100) = 70%"
    d.txtID.insert(0, "176")
    d.txtCIP.delete(0, "end")
    d.txtCIP.insert(0, "172.27.0.116")
    d.on_save()                                                         # 속도를 고치지 않고 그대로 추가
    assert msgs == []
    m = next(x for x in p.houses["H1R7"].get_list() if x.id == 176)
    assert (m.set_rpm, m.ctrl_mode, m.speed_src) == (770, CTRL_MODE_MANUAL, "manual")


def test_new_motor_add_dialog_defaults_to_manual_770(root, tmp_path, monkeypatch):
    """2.7.0: RPM 제어 소스는 모터마다 - 신규모터 추가 창의 기본값은 수동 770 RPM."""
    from motor_monitor.ui import dialogs
    from motor_monitor.ui import theme as T
    monkeypatch.setattr(dialogs.BaseDialog, "error", lambda self, m, t: None)
    monkeypatch.setattr(dialogs.BaseDialog, "ask", lambda self, m, t: True)
    p = _plant(tmp_path)
    d = dialogs.MotorDialog(root, p, T.Images(root), "H1R1", "ADD")
    d.motor_type_var.set(int(MotorType.NEW_MOTOR))
    d._on_type()
    assert d.ctrl_mode_var.get() == CTRL_MODE_MANUAL and d.txtSETRPM.get() == "770"
    d.txtID.insert(0, "4")
    d.on_save()
    m = p.houses["H1R1"].get_motor("1_4")
    assert (m.ctrl_mode, m.set_rpm) == (CTRL_MODE_MANUAL, 770)
    assert p.houses["H1R1"].add_motor(5, MotorType.NEW_MOTOR, True)
    assert p.houses["H1R1"].get_motor("1_5").set_rpm == 770                # 코드로 추가해도 기본 770


# ---------------------------------------------------------------- 4 신규모터 아이콘 개나리색
def _pixels(root, name):
    import tkinter as tk
    from motor_monitor.ui import theme as T
    img = tk.PhotoImage(master=root, file=os.path.join(T.ASSET_DIR, f"{name}.png"))
    out = []
    for x in range(img.width()):
        for y in range(img.height()):
            if img.transparency_get(x, y):
                continue
            r, g, b = img.get(x, y)
            out.append(colorsys.rgb_to_hsv(r / 255, g / 255, b / 255))
    return out


def test_new_motor_icons_have_no_orange_and_keep_red_L(root):
    for name in ("newmotor", "newmotor_limit"):
        px = _pixels(root, name)
        orange = [p for p in px if 15 / 360 <= p[0] <= 40 / 360 and p[1] >= 0.3]
        yellow = [p for p in px if 42 / 360 <= p[0] <= 58 / 360 and p[1] >= 0.3]
        assert not orange and len(yellow) > len(px) * 0.8, name
    red = [p for p in _pixels(root, "newmotor_limit") if (p[0] < 10 / 360 or p[0] > 350 / 360) and p[1] > 0.6]
    assert len(red) > 30                                                  # 빨간 L (속도제한) 그대로


# ---------------------------------------------------------------- 6 수동/참조 표시
def test_speed_source_rules():
    new = Motor(1, MotorType.NEW_MOTOR, True)
    assert speed_source(new) == "manual"
    new.ctrl_mode = CTRL_MODE_REF
    assert speed_source(new) == "ref"
    new.mon_only = True
    assert speed_source(new) == "mon"                                    # 모니터링만
    ctrl = Motor(2, MotorType.OLD_CTRL, True, ctrl_mode=CTRL_MODE_MANUAL)
    assert speed_source(ctrl) == "manual"
    ctrl.ctrl_mode = CTRL_MODE_REF
    assert speed_source(ctrl) == "ref"
    assert speed_source(Motor(3, MotorType.OLD_MOTOR, True)) == ""        # 기존모터(읽기만)는 표시 없음


def test_get_list_fills_speed_source_per_motor(tmp_path):
    p = _plant(tmp_path)
    h = p.houses["H1R1"]
    h.add_motor(2, MotorType.OLD_MOTOR, True)
    h.add_motor(5, MotorType.NEW_MOTOR, True)
    h.add_motor(6, MotorType.NEW_MOTOR, True, ctrl_mode=CTRL_MODE_REF, ref_line="H1R1", ref_id=2)
    h.add_motor(9, MotorType.OLD_CTRL, True, ip="172.27.0.84", set_rpm=770, ctrl_mode=CTRL_MODE_REF)
    assert {m.id: m.speed_src for m in h.get_list()} == {2: "", 5: "manual", 6: "ref", 9: "ref"}
    h.mod_motor("1_5", True, mon_only=True)
    assert {m.id: m.speed_src for m in h.get_list()}[5] == "mon"
    assert {m.id: m.disp_ip for m in h.get_list()} == {2: "192.168.0.1", 5: "192.168.0.2", 6: "192.168.0.2",
                                                        9: "172.27.0.84"}


def _list_view(root):
    import tkinter as tk
    from motor_monitor.ui import theme as T
    from motor_monitor.ui.widgets import MotorListView
    frame = tk.Frame(root)
    frame.pack(fill="both", expand=True)
    lv = MotorListView(frame, T.Images(root), on_header_click=lambda k: None, on_double_click=lambda k: None)
    lv.pack(fill="both", expand=True)
    root.update()
    return lv


def test_status_column_shows_manual_or_ref_between_icon_and_number(root):
    lv = _list_view(root)
    rows = [Motor(1, MotorType.NEW_MOTOR, True, status=MotorStatus.OK, rpm=770, target_rpm=770, speed_src="manual"),
            Motor(2, MotorType.NEW_MOTOR, True, status=MotorStatus.OK, rpm=768, target_rpm=769, speed_src="ref"),
            Motor(3, MotorType.OLD_CTRL, True, status=MotorStatus.OK, rpm=770, set_rpm=770, speed_src="manual"),
            Motor(4, MotorType.OLD_CTRL, True, status=MotorStatus.READY, ctrl_mode=CTRL_MODE_REF, speed_src="ref",
                  ref_line="H1R2", ref_id=66),
            Motor(5, MotorType.OLD_MOTOR, True, status=MotorStatus.OK, rpm=770),
            Motor(6, MotorType.NEW_MOTOR, False, status=MotorStatus.NOTUSE, speed_src="manual"),
            Motor(7, MotorType.NEW_MOTOR, True, status=MotorStatus.OK, rpm=500, mon_only=True, speed_src="mon")]
    rows[1].ref_line, rows[1].ref_gw, rows[1].ref_id = "H1R1", 2, 7
    lv.set_rows(rows)
    assert lv.drawn_srcs == ["수동", "참조", "수동", "참조", "", "", "감시"]  # 기존모터·사용안함은 표시 없음
    c = lv.canvas
    badges = [i for i in c.find_withtag("srcbadge") if c.type(i) == "text"]
    assert [c.itemcget(i, "text") for i in badges] == ["수동", "참조", "수동", "참조", "감시"]
    # 첫 행: 상태 그림 오른쪽, RPM 숫자 왼쪽
    _x0, _x1, _xi, x2, _x3 = lv._col_x()
    bl, _bt, br, _bb = c.bbox(badges[0])
    num = next(i for i in c.find_all() if c.type(i) == "text" and c.itemcget(i, "text") == "770")
    assert x2 + lv.STATUS_ICON_W < bl and br < c.bbox(num)[0]
    # 툴팁 (표시 위에 마우스)
    bx = x2 + lv.STATUS_ICON_W + lv.BADGE_X + lv.BADGE_W // 2

    def tip(i):
        return lv._tooltip_text(bx, i * lv.ROW_H + lv.ROW_H // 2)
    assert tip(0) == "수동: 지정 0 RPM (0%)으로 운전\n지금 설정 770 RPM (70%)"
    assert tip(1).startswith("참조(자동): 1구역 기존모터 2 ID 7 회전수를 따라 운전")
    assert tip(2) == "수동: 지정 770 RPM (70%)으로 운전"
    assert tip(3).startswith("참조(자동): 2구역 기존모터 ID 66 회전수를 같은 비율로 따라 운전")
    assert tip(4) == "항목 더블클릭시 수정/삭제"                          # 표시 없는 행
    assert tip(6).startswith("감시(모니터링만)")


def test_zone_settings_save_updates_ip_column_at_once(root, tmp_path, monkeypatch):
    from motor_monitor.ui import dialogs, main_window
    monkeypatch.setattr(dialogs.BaseDialog, "error", lambda self, m, t: None)
    p = _plant(tmp_path)
    h = p.houses["H1R2"]
    h.add_motor(66, MotorType.OLD_MOTOR, True)
    h.add_motor(3, MotorType.NEW_MOTOR, True)
    win = main_window.MainWindow(root, p)
    lv = win.zone_panels["H1R2"].LV
    assert lv.drawn_srcs == ["수동", ""]                                   # 신규모터 3, 기존모터 66 (ID 순)
    ips = lambda: [lv.canvas.itemcget(i, "text") for i in lv.canvas.find_withtag("ip")]
    assert ips() == ["4", "3"]                                            # IP 열: 마지막 숫자만 (192.168.0.4 / .3)
    # 톱니 버튼의 구역설정 창은 저장 후 목록까지 갱신하는 콜백을 씀
    seen = {}

    class FakeDialog:
        def __init__(self, *a, on_saved=None, **k):
            seen["on_saved"] = on_saved

        def show_modal(self):
            pass
    monkeypatch.setattr(main_window, "LineSettingsDialog", FakeDialog)
    win.zone_panels["H1R2"].open_settings()
    assert seen["on_saved"] == win.after_line_saved
    d = dialogs.LineSettingsDialog(root, p, "H1R2", on_saved=win.after_line_saved)
    d.txtSIP.delete(0, "end"); d.txtSIP.insert(0, "172.27.0.84")
    d.txtTIP.delete(0, "end"); d.txtTIP.insert(0, "172.27.0.116")
    d.on_save()
    assert ips() == ["116", "84"]                                         # 저장 즉시 목록 IP 열 갱신
