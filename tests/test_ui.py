"""화면 동작 시험. 디스플레이가 없으면 건너뜀 (리눅스: xvfb-run -a python -m pytest)."""
import pytest

tk = pytest.importorskip("tkinter")

from motor_monitor.constants import MotorStatus, MotorType  # noqa: E402
from motor_monitor.crypto3des import decrypt_config  # noqa: E402
from motor_monitor.plant import Plant  # noqa: E402


@pytest.fixture
def root():
    try:
        r = tk.Tk()
    except tk.TclError:
        pytest.skip("디스플레이 없음")
    r.withdraw()
    yield r
    r.destroy()


@pytest.fixture
def plant(tmp_path):
    p = Plant(str(tmp_path / "Config.dll"))
    p.load()
    return p


@pytest.fixture
def ui(root, plant, monkeypatch):
    from motor_monitor.ui import dialogs, main_window
    msgs = []
    answers = {"yes": True}
    monkeypatch.setattr(dialogs.BaseDialog, "error", lambda self, m, t: msgs.append((t, m)))
    monkeypatch.setattr(dialogs.BaseDialog, "ask", lambda self, m, t: answers["yes"])
    monkeypatch.setattr(main_window.messagebox, "askyesno", lambda *a, **k: answers["yes"])
    win = main_window.MainWindow(root, plant)
    return win, dialogs, msgs, answers


def test_eight_zones_with_own_buttons(ui, plant):
    win, *_ = ui
    assert list(win.zone_panels) == [f"H1R{i}" for i in range(1, 9)]
    zp = win.zone_panels["H1R3"]
    assert zp.StartBtn.text == "3구역 제어 시작" and zp.StopBtn.text == "3구역 제어 종료"
    assert win.mode_banner is None and "모니터링" not in win.root.title()


def test_zone_dialog_validation_and_save(ui, plant):
    win, dialogs, msgs, _ = ui
    d = dialogs.LineSettingsDialog(win.root, plant, "H1R1", on_saved=win.update_header_box)
    assert d.lbl_house.cget("text") == "1구역"
    assert d.title_label.cget("text") == "구역설정"
    assert list(d.cb_lines.cget("values")) == ["구역선택"] + [f"{i}구역" for i in range(1, 9)]
    assert d.source_mode.get() == "manual"

    d.txtSIP.delete(0, "end"); d.txtSIP.insert(0, "300.1.1.1")
    d.on_save()
    assert msgs[-1] == ("에러", "기존모터 접속 IP가 올바른 형식이 아닙니다..")
    d.txtSIP.delete(0, "end"); d.txtSIP.insert(0, "172.27.0.82")
    d.txtMRPM.delete(0, "end"); d.txtMRPM.insert(0, "1200")
    d.on_save()
    assert msgs[-1] == ("에러", "수동 RPM속도를 0~1100범위내로 입력해주세요.")
    d.txtMRPM.delete(0, "end"); d.txtMRPM.insert(0, "800")
    d.on_save()
    c = plant.configs["H1R1"]
    assert (c.source_ip, c.manual_rpm, c.source_manual) == ("172.27.0.82", 800, True)
    assert win.zone_panels["H1R1"].lbl_source.cget("text") == "수동 > 800 RPM"
    assert win.zone_panels["H1R1"].lbl_old_ip.cget("text") == "172.27.0.82/502"


def test_zone_dialog_reference_mode(ui, plant):
    win, dialogs, msgs, _ = ui
    plant.houses["H1R2"].add_motor(66, MotorType.OLD_MOTOR, True)
    d = dialogs.LineSettingsDialog(win.root, plant, "H1R4", on_saved=win.update_header_box)
    d.source_mode.set("ref"); d._on_mode()
    d._select_line(""); d._on_line_changed()
    d.on_save()
    assert msgs[-1] == ("에러", "참조할 모터가 있는 구역을 선택해주세요.")
    d._select_line("R2"); d._on_line_changed()
    assert list(d.cb_ids.cget("values")) == ["ID", "66"]
    d.on_save()
    assert msgs[-1] == ("에러", "참조할 모터 ID를 선택해주세요.")
    d.cb_ids.current(1)
    d.on_save()
    c = plant.configs["H1R4"]
    assert (c.source_manual, c.source_line, c.source_id) == (False, "H1R2", 66)
    assert win.zone_panels["H1R4"].lbl_source.cget("text") == "2구역 > 66"


def test_motor_add_with_limit(ui, plant):
    win, dialogs, msgs, _ = ui
    zp = win.zone_panels["H1R3"]
    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R3", "ADD", on_changed=zp.refresh)
    d.update_idletasks()
    assert d.use_var.get() == 1 and not d.limit_row_visible     # 기본 '사용', 속도제한 줄 숨김
    d.on_save()
    assert msgs[-1] == ("오류", "모터 종류을 선택해 주세요.")
    d.motor_type_var.set(1); d._on_type()
    assert d.limit_row_visible                                    # 신규모터 선택 시 표시
    d.txtID.insert(0, "251")
    d.on_save()
    assert msgs[-1] == ("오류", "모터ID값은 숫자로 1-250범위로 입력하셔야 합니다.")
    d.txtID.delete(0, "end"); d.txtID.insert(0, "13")
    d.cbUSELIMIT.current(0)
    d.txtLIMITRPM.delete(0, "end"); d.txtLIMITRPM.insert(0, "2000")
    d.on_save()
    assert msgs[-1] == ("오류", "제한모드 RPM값은 숫자 0~1100범위로 입력하셔야 합니다.")
    d.txtLIMITRPM.delete(0, "end"); d.txtLIMITRPM.insert(0, "500")
    d.on_save()
    m = plant.houses["H1R3"].get_motor("1_13")
    assert m.use and m.use_limit and m.limit_rpm == 500
    assert [r.key for r in zp.LV.rows] == ["1_13"]
    assert "H1R3_MOTORS=13=1=1=1=500" in decrypt_config(open(plant.config_path, "rb").read())

    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R3", "ADD")
    d.motor_type_var.set(1); d.txtID.insert(0, "13")
    d.on_save()
    assert msgs[-1] == ("오류", "이미 해당 모터 종류와 ID를 갖는 모터가 있습니다. 확인바랍니다.")
    d.destroy()


def test_motor_modify_and_delete(ui, plant):
    win, dialogs, msgs, _ = ui
    zp = win.zone_panels["H1R1"]
    h = plant.houses["H1R1"]
    h.add_motor(2, MotorType.OLD_MOTOR, True)
    h.add_motor(20, MotorType.NEW_MOTOR, True, True, 500)
    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R1", "MOD", mod_dkey="1_20", on_changed=zp.refresh)
    d.update_idletasks()
    assert d.lblID.cget("text") == "20" and d.limit_row_visible
    assert d.cbUSELIMIT.current() == 0 and d.txtLIMITRPM.get() == "500"
    d.cbUSELIMIT.current(1); d.use_var.set(0)
    d.on_save()
    m = h.get_motor("1_20")
    assert not m.use_limit and m.status == MotorStatus.NOTUSE

    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R1", "MOD", mod_dkey="0_2", on_changed=zp.refresh)
    d.update_idletasks()
    assert not d.limit_row_visible                                # 기존모터는 속도제한 없음
    plant.configs["H1R5"] = plant.configs["H1R5"].copy(source_manual=False, source_line="H1R1", source_id=2)
    d.on_delete()
    assert msgs[-1] == ("오류", "해당 모터를 참조로 사용하는 구역이 있습니다. 변경 후, 삭제할 수 있습니다.")
    plant.configs["H1R5"] = plant.configs["H1R5"].copy(source_manual=True)
    d.on_delete()
    assert [r.key for r in zp.LV.rows] == ["1_20"]


def test_zone_start_stop_and_status_icons(ui, plant, monkeypatch):
    win, *_ = ui
    from motor_monitor.ui import theme as T
    from motor_monitor.ui.widgets import MotorListView
    zp = win.zone_panels["H1R2"]
    started = []
    monkeypatch.setattr(plant, "start_line", lambda c: (started.append(c), plant.houses[c]._stop.clear()))
    monkeypatch.setattr(plant, "stop_line", lambda c: plant.houses[c]._stop.set())
    win.update_network_status()
    assert zp.rec_SOURCE_NETSTATUS.color == T.NET_STOP
    zp.on_start()
    assert started == ["H1R2"]
    assert zp.rec_SOURCE_NETSTATUS.color == T.NET_ERR and zp.rec_SID_NETSTATUS.color == T.NET_OK
    assert win.zone_panels["H1R1"].rec_SOURCE_NETSTATUS.color == T.NET_STOP   # 다른 구역은 그대로
    zp.on_stop()
    assert zp.rec_SOURCE_NETSTATUS.color == T.NET_STOP

    h = plant.houses["H1R2"]
    h.add_motor(1, MotorType.NEW_MOTOR, True, True, 400)
    m = h.get_motor("1_1")
    assert MotorListView.type_icon(m) == ("newmotor_limit", "신규모터(제한모드)")
    m.status, m.had_err = MotorStatus.OK, True
    assert MotorListView.status_icon(m) == ("bad_haderr", "정상(이상기록있음)")
    m.status = MotorStatus.ERR
    assert MotorListView.status_icon(m) == ("bad_new", "이상")
    m.status = MotorStatus.READY
    assert MotorListView.status_icon(m) == (None, "")


def test_monitoring_mode_screen(root, tmp_path, monkeypatch):
    from motor_monitor.ui import main_window
    asked = []
    monkeypatch.setattr(main_window.messagebox, "askyesno", lambda t, m, **k: (asked.append((t, m)), True)[1])
    plant = Plant(str(tmp_path / "Config.dll"), monitoring=True)
    plant.load()
    monkeypatch.setattr(plant, "start_line", lambda c: plant.houses[c]._stop.clear())
    win = main_window.MainWindow(root, plant)
    assert "모니터링 모드" in root.title()
    assert win.mode_banner is not None and "제어 신호" in win.mode_banner.cget("text")
    zp = win.zone_panels["H1R2"]
    assert zp.StartBtn.text == "2구역 모니터링 시작" and zp.StopBtn.text == "2구역 모니터링 종료"
    zp.on_start()
    assert asked[-1] == ("모니터링", "2구역 모니터링을 시작하시겠습니까?")


def test_ctrl_motor_add_and_modify(ui, plant):
    win, dialogs, msgs, _ = ui
    zp = win.zone_panels["H1R2"]
    plant.configs["H1R2"] = plant.configs["H1R2"].copy(source_ip="172.27.0.84", source_port=502)
    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R2", "ADD", on_changed=zp.refresh)
    d.update_idletasks()
    assert not d.ctrl_rows_visible
    d.motor_type_var.set(2); d._on_type()
    assert d.ctrl_rows_visible and not d.limit_row_visible
    assert d.txtCIP.get() == "172.27.0.84" and d.txtCPORT.get() == "502"     # 기존모터 게이트웨이 기본값
    assert d.lbl_id_range.cget("text") == "정수(1~247)"
    d.txtID.insert(0, "248")
    d.on_save()
    assert msgs[-1] == ("오류", "모터ID값은 숫자로 1-247범위로 입력하셔야 합니다.")
    d.txtID.delete(0, "end"); d.txtID.insert(0, "66")
    d.on_save()
    assert msgs[-1] == ("오류", "운전속도를 숫자 0~1100범위로 입력해주세요.")       # 운전속도는 꼭 입력
    d.txtSETRPM.insert(0, "700")
    d.txtCIP.delete(0, "end"); d.txtCIP.insert(0, "172.27.0")
    d.on_save()
    assert msgs[-1] == ("오류", "기존모터제어 접속 IP가 올바른 형식이 아닙니다..")
    d.txtCIP.delete(0, "end"); d.txtCIP.insert(0, "172.27.0.84")
    d.on_save()
    key = "2_66@172.27.0.84:502"
    m = plant.houses["H1R2"].get_motor(key)
    assert m.mtype == MotorType.OLD_CTRL and m.set_rpm == 700 and m.use
    assert [r.key for r in zp.LV.rows] == [key]
    from motor_monitor.ui.widgets import MotorListView
    assert MotorListView.type_icon(m)[0] == "oldmotor_ctrl"

    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R2", "MOD", mod_dkey=key, on_changed=zp.refresh)
    d.update_idletasks()
    assert d.ctrl_rows_visible and d.txtSETRPM.get() == "700" and str(d.txtCIP.cget("state")) == "readonly"
    d.txtSETRPM.delete(0, "end"); d.txtSETRPM.insert(0, "550")
    d.on_save()
    assert plant.houses["H1R2"].get_motor(key).set_rpm == 550
