"""화면 동작 시험. 디스플레이가 없으면 건너뜀 (리눅스: xvfb-run -a python -m pytest)."""

import pytest

tk = pytest.importorskip("tkinter")

from motor_monitor.constants import MotorStatus, MotorType  # noqa: E402
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
    p = Plant(str(tmp_path / "Config.ini"))
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


def test_line_dialog_validation_and_save(ui, plant):
    win, dialogs, msgs, _ = ui
    d = dialogs.LineSettingsDialog(win.root, plant, "H1R1E", on_saved=win.update_header_box)
    assert d.lbl_house.cget("text") == "1구역 / 짝"
    assert d.source_mode.get() == "manual"           # 기본값: 수동
    assert d.lbl_rpm_range.cget("text") == "정수(0~1100)"

    d.txtSIP.delete(0, "end"); d.txtSIP.insert(0, "300.1.1.1")
    d.on_save()
    assert msgs[-1] == ("에러", "기존모터 접속 IP가 올바른 형식이 아닙니다..")
    d.txtSIP.delete(0, "end"); d.txtSIP.insert(0, "10.0.0.1")
    d.txtTPORT.delete(0, "end"); d.txtTPORT.insert(0, "70000")
    d.on_save()
    assert msgs[-1] == ("에러", "신규모터 접속 PORT를 숫자로 1~65535범위로 입력해 주세요.")
    d.txtTPORT.delete(0, "end"); d.txtTPORT.insert(0, "503")
    d.txtMRPM.delete(0, "end"); d.txtMRPM.insert(0, "1200")
    d.on_save()
    assert msgs[-1] == ("에러", "수동 RPM속도를 0~1100범위내로 입력해주세요.")
    d.txtMRPM.delete(0, "end"); d.txtMRPM.insert(0, "800")
    d.on_save()
    c = plant.configs["H1R1E"]
    assert (c.source_ip, c.target_port, c.manual_rpm, c.source_manual) == ("10.0.0.1", 503, 800, True)
    lp = win.line_panels["H1R1E"]
    assert lp.lbl_old_ip.cget("text") == "10.0.0.1/502"
    assert lp.lbl_source.cget("text") == "수동 > 800 RPM"


def test_line_dialog_reference_mode(ui, plant):
    win, dialogs, msgs, _ = ui
    plant.houses["H1R2O"].add_motor(7, MotorType.OLD_MOTOR, True)
    d = dialogs.LineSettingsDialog(win.root, plant, "H1R1E", on_saved=win.update_header_box)
    d.source_mode.set("ref"); d._on_mode()
    d._select_line(""); d._on_line_changed()
    d.on_save()
    assert msgs[-1] == ("에러", "참조할 모터가 있는 라인을 선택해주세요.")
    d._select_line("H1R2O"); d._on_line_changed()
    assert list(d.cb_ids.cget("values")) == ["ID", "7"]
    d.on_save()
    assert msgs[-1] == ("에러", "참조할 모터 ID를 선택해주세요.")
    d.cb_ids.current(1)
    d.on_save()
    c = plant.configs["H1R1E"]
    assert (c.source_manual, c.source_line, c.source_id) == (False, "H1R2O", 7)
    assert win.line_panels["H1R1E"].lbl_source.cget("text") == "2구역 > 홀 > 7"


def test_motor_add_modify_delete(ui, plant):
    win, dialogs, msgs, answers = ui
    lp = win.line_panels["H1R1E"]
    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R1E", "ADD", on_changed=lp.refresh)
    d.on_save()
    assert msgs[-1] == ("오류", "모터 종류을 선택해 주세요.")
    d.motor_type_var.set(0)
    d.txtID.insert(0, "100")
    d.on_save()
    assert msgs[-1] == ("오류", "모터ID값은 숫자로 1-99범위로 입력하셔야 합니다.")
    d.txtID.delete(0, "end"); d.txtID.insert(0, "5")
    d.on_save()
    assert msgs[-1] == ("오류", "사용 여부을 선택해 주세요.")
    d.use_var.set(1)
    d.on_save()
    assert [m.key for m in lp.LV.rows] == ["0_5"]
    assert "H1R1E_MOTORS=5=0=1" in open(plant.config_path, encoding="utf-8").read()

    # 중복 추가
    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R1E", "ADD")
    d.motor_type_var.set(0); d.txtID.insert(0, "5"); d.use_var.set(1)
    d.on_save()
    assert msgs[-1] == ("오류", "이미 해당 모터 종류와 ID를 갖는 정보가 있습니다. 확인바랍니다.")
    d.destroy()

    # 수정: 사용 중지
    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R1E", "MOD", mod_dkey="0_5", on_changed=lp.refresh)
    assert d.lblID.cget("text") == "5" and d.use_var.get() == 1
    d.use_var.set(0)
    d.on_save()
    assert plant.houses["H1R1E"].get_motor("0_5").status == MotorStatus.NOTUSE

    # 참조 중인 모터는 삭제 불가
    plant.configs["H1R2E"] = plant.configs["H1R2E"].copy(source_manual=False, source_line="H1R1E", source_id=5)
    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R1E", "MOD", mod_dkey="0_5", on_changed=lp.refresh)
    d.on_delete()
    assert msgs[-1] == ("오류", "해당 모터를 참조로 사용하는 라인이 있습니다. 변경 후, 삭제할 수 있습니다.")
    plant.configs["H1R2E"] = plant.configs["H1R2E"].copy(source_manual=True)
    d.on_delete()
    assert lp.LV.rows == []


def test_eight_zones_on_one_screen(ui, plant):
    win, dialogs, msgs, _ = ui
    assert len(win.line_panels) == 16 and sorted(win.zone_panels) == list(range(1, 9))
    # 5구역(온실2 1구역) 짝 라인을 1구역 짝 기존모터 참조로 설정
    plant.houses["H1R1E"].add_motor(3, MotorType.OLD_MOTOR, True)
    d = dialogs.LineSettingsDialog(win.root, plant, "H2R1E", on_saved=win.update_header_box)
    assert d.lbl_house.cget("text") == "5구역 / 짝"
    assert len(d.cb_lines.cget("values")) == 17
    d.source_mode.set("ref"); d._on_mode()
    d._select_line("H1R1E"); d._on_line_changed()
    d.cb_ids.current(1)
    d.on_save()
    c = plant.configs["H2R1E"]
    assert (c.source_manual, c.source_line, c.source_id) == (False, "H1R1E", 3)
    assert win.line_panels["H2R1E"].lbl_source.cget("text") == "1구역 > 짝 > 3"


def test_main_window_controls_and_sort(ui, plant, monkeypatch):
    win, _, _, answers = ui
    from motor_monitor.ui import theme as T
    h = plant.houses["H1R1E"]
    for mid, t in [(3, 0), (1, 1), (2, 0)]:
        h.add_motor(mid, MotorType(t), True)
    lp = win.line_panels["H1R1E"]
    lp.refresh()
    assert [m.id for m in lp.LV.rows] == [1, 2, 3]
    lp.on_sort("ID"); lp.on_sort("ID")
    assert [m.id for m in lp.LV.rows] == [3, 2, 1]

    win.update_network_status()
    assert lp.rec_SOURCE_NETSTATUS.color == T.NET_STOP

    started = []
    monkeypatch.setattr(plant, "start_house", lambda hp: (started.append(hp), plant.network_run.__setitem__(hp, True)))
    monkeypatch.setattr(plant, "stop_house", lambda hp: plant.network_run.__setitem__(hp, False))
    win.on_h1_start()
    assert started == ["H1"]
    assert lp.rec_SOURCE_NETSTATUS.color == T.NET_ERR          # 실행 중 + 미접속 = 빨강
    assert lp.rec_SID_NETSTATUS.color == T.NET_OK              # 수동 모드 = 항상 정상
    win.on_h2_start()
    assert started == ["H1", "H2"]                             # v1.1.0: 5~8구역 제어 버튼 활성
    assert win.line_panels["H2R1E"].rec_SOURCE_NETSTATUS.color == T.NET_ERR
    win.on_h1_stop()
    assert lp.rec_SOURCE_NETSTATUS.color == T.NET_STOP
