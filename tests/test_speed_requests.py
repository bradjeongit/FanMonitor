"""운영 프로그램 요청 시험 (2.5.0, 2.5.1):
1 목록은 ID 순서가 기본 (test_core 의 정렬 시험), 2 확인 창 키보드(Y/Enter/Space = 예, N/Esc = 아니오),
3 문제 상태 행 음영, 7 팬 속도는 30초에 한 번 점검 (설정을 바꾸면 바로), 모터 최고속도(h8) 대비 % 표시,
속도 보낼 때 아이콘 깜빡임.
2.5.1: 속도 보정(4·5·6 - 실제 속도로 지정값 조절, 무반응 알림, 보정값 기억)은 삭제 -> 목표 속도를 그대로 보냄."""
import os
import sys
import time
from types import SimpleNamespace

import pytest

from motor_monitor import constants as C
from motor_monitor import modbus
from motor_monitor.constants import CTRL_MODE_MANUAL, MotorStatus, MotorType
from motor_monitor.house import Motor
from motor_monitor.plant import Plant
from motor_monitor.simulator import FakeGateway


@pytest.fixture
def fast(monkeypatch):
    for name, val in [("SEND_INTERVAL", 0.01), ("CYCLE_INTERVAL", 0.05), ("RECONNECT_DELAY", 0.2),
                      ("ERROR_RETRY_DELAY", 0.05), ("CHECK_INTERVAL", 0.15)]:
        monkeypatch.setattr(C, name, val)


def _wait(cond, timeout=8.0):
    end = time.time() + timeout
    while time.time() < end:
        if cond():
            return True
        time.sleep(0.02)
    return False


def _plant(path, monitoring=False):
    p = Plant(str(path / "Config.dll"), monitoring=monitoring)
    p.load()
    return p


def _rpm_writes(gw, uid):
    return [v for u, reg, v in list(gw.writes) if u == uid and reg == modbus.NEW_REG_RPM]


# ---------------------------------------------------------------- 2.5.1 속도 보정 삭제: 목표 속도를 그대로 보냄
def test_speed_within_is_display_only():
    assert C.speed_within(792, 800) and C.speed_within(808, 800)          # 화면의 '오차 1% 이내' 표시용
    assert not C.speed_within(791, 800) and not C.speed_within(809, 800)
    assert not hasattr(C, "SPEED_STEP_MIN_PCT") and not hasattr(C, "SPEED_MEMORY_FILE")
    with pytest.raises(ImportError):
        import motor_monitor.speedtrim  # noqa: F401                    # 보정 모듈 없음


def _zone(p, gw, rpm=800):
    p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip=gw.host, source_port=gw.port, target_ip=gw.host,
                                               target_port=gw.port, source_manual=True, manual_rpm=rpm)


def test_new_motor_gets_target_as_is_without_correction(tmp_path, fast):
    """실제 속도가 설정보다 10% 느린 신규모터(배율 0.9)라도 지정값을 올리지 않고 설정 800 만 보냄. 속도 미달 알림·기억 파일 없음."""
    gw = FakeGateway().start()
    gw.add_new(5)
    gw.new_gain[5] = 0.9
    try:
        p = _plant(tmp_path)
        _zone(p, gw)
        h = p.houses["H1R1"]
        h.add_motor(5, MotorType.NEW_MOTOR, True)
        p.start_line("H1R1")
        assert _wait(lambda: h.get_motor("1_5").rpm == 720)               # 실제 = 800 x 0.9
        time.sleep(0.8)                                                    # 점검 여러 번
        assert _rpm_writes(gw, 5) == [800]                                 # 보정 없음: 처음 한 번만
        m = h.get_motor("1_5")
        assert (m.target_rpm, m.cmd_rpm, m.speed_note) == (800, 800, "") and not m.has_problem
        p.shutdown()
        assert not os.path.exists(tmp_path / "speed_memory.json")         # 기억 파일 없음
    finally:
        gw.stop()


def test_target_change_is_sent_as_is(tmp_path, fast):
    """설정 속도가 바뀌면 (1% 이내 변화도) 그 값을 그대로 보냄. 같으면 다시 보내지 않음."""
    gw = FakeGateway().start()
    gw.add_new(5)
    try:
        p = _plant(tmp_path)
        _zone(p, gw)
        h = p.houses["H1R1"]
        h.add_motor(5, MotorType.NEW_MOTOR, True)
        p.start_line("H1R1")
        assert _wait(lambda: _rpm_writes(gw, 5) == [800])
        time.sleep(0.4)
        assert _rpm_writes(gw, 5) == [800]                                 # 같으면 다시 보내지 않음
        p.set_line_config("H1R1", p.configs["H1R1"].copy(manual_rpm=805))
        assert _wait(lambda: _rpm_writes(gw, 5) == [800, 805])
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_new_motor_limit_is_sent(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_new(7)
    gw.new_gain[7] = 0.95
    try:
        p = _plant(tmp_path)
        _zone(p, gw)
        h = p.houses["H1R1"]
        h.add_motor(7, MotorType.NEW_MOTOR, True, True, 500)               # 속도제한 500
        p.start_line("H1R1")
        assert _wait(lambda: _rpm_writes(gw, 7) == [500])
        time.sleep(0.5)
        assert _rpm_writes(gw, 7) == [500] and h.get_motor("1_7").speed_note == ""
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_ctrl_motor_gets_target_as_is_without_correction(tmp_path, fast):
    """기존모터제어 수동 700, 실제 = h2 x 0.9 이어도 h2 는 700 그대로."""
    gw = FakeGateway().start()
    gw.add_ecblue(2, analog_rpm=0, h4=0x0101)
    gw.ecb[2]["gain"] = 0.9
    try:
        p = _plant(tmp_path)
        h = p.houses["H1R1"]
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, set_rpm=700, ctrl_mode=CTRL_MODE_MANUAL)
        key = f"2_2@{gw.host}:{gw.port}"
        p.start_line("H1R1")
        assert _wait(lambda: gw.ecb[2]["h"][2] == 700 and h.get_motor(key).rpm == 630)
        time.sleep(0.8)
        h2s = [v for u, reg, v in gw_writes(gw) if u == 2 and reg == modbus.ECB_H_SPEED]
        assert h2s == [700] and h.get_motor(key).speed_note == ""
        h.service_stop(join=True)
    finally:
        gw.stop()


def gw_writes(gw):
    """시뮬레이터가 받은 ECblue 쓰기 (FC06) 목록 (uid, reg, value)."""
    out = []
    for req in list(gw.requests):
        if len(req) >= 12 and req[7] == modbus.FC_WRITE_SINGLE:
            out.append((req[6], req[8] * 256 + req[9], req[10] * 256 + req[11]))
    return out


# ---------------------------------------------------------------- 7 30초에 한 번 점검, 설정을 바꾸면 바로
def test_speed_is_checked_once_per_interval(tmp_path, fast, monkeypatch):
    monkeypatch.setattr(C, "CHECK_INTERVAL", 1.0)
    gw = FakeGateway().start()
    gw.add_old(2, 770)
    gw.add_old(4, 760)
    try:
        p = _plant(tmp_path, monitoring=True)
        _zone(p, gw)
        h = p.houses["H1R1"]
        h.add_motor(2, MotorType.OLD_MOTOR, True)
        h.add_motor(4, MotorType.OLD_MOTOR, True)
        p.start_line("H1R1")
        time.sleep(2.5)                                             # 0초, 1초, 2초 = 점검 3번
        reads = [r for r in list(gw.requests) if r[7] == modbus.FC_READ_INPUT and r[6] in (2, 4)]
        assert 4 <= len(reads) <= 8                                 # 모터 2대 x 점검 2~4번 (예전: 0.3초마다)
        assert h.get_motor("0_2").status == MotorStatus.OK
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_setting_change_is_applied_without_waiting(tmp_path, fast, monkeypatch):
    monkeypatch.setattr(C, "CHECK_INTERVAL", 30)                    # 실제 주기
    gw = FakeGateway().start()
    gw.add_new(5)
    try:
        p = _plant(tmp_path)
        _zone(p, gw)
        h = p.houses["H1R1"]
        h.add_motor(5, MotorType.NEW_MOTOR, True)
        p.start_line("H1R1")
        assert _wait(lambda: _rpm_writes(gw, 5) == [800], timeout=3)      # 시작하면 바로 첫 점검
        t0 = time.time()
        p.set_line_config("H1R1", p.configs["H1R1"].copy(manual_rpm=900))
        assert _wait(lambda: _rpm_writes(gw, 5) == [800, 900], timeout=3)  # 30초를 기다리지 않음
        assert time.time() - t0 < 3
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_switching_to_control_mode_sends_at_once(tmp_path, fast, monkeypatch):
    monkeypatch.setattr(C, "CHECK_INTERVAL", 30)
    gw = FakeGateway().start()
    gw.add_new(5)
    try:
        p = _plant(tmp_path, monitoring=True)
        _zone(p, gw)
        h = p.houses["H1R1"]
        h.add_motor(5, MotorType.NEW_MOTOR, True)
        p.start_line("H1R1")
        assert _wait(lambda: h.get_motor("1_5").status == MotorStatus.OK, timeout=3)
        assert _rpm_writes(gw, 5) == []                             # 모니터링 모드: 쓰지 않음
        p.set_monitoring(False)
        assert _wait(lambda: _rpm_writes(gw, 5) == [800], timeout=3)
        h.service_stop(join=True)
    finally:
        gw.stop()


# ---------------------------------------------------------------- 2 확인 창 키보드, 3 행 음영 (화면)
@pytest.fixture
def shown_root():
    tk = pytest.importorskip("tkinter")
    try:
        r = tk.Tk()
    except tk.TclError:
        pytest.skip("디스플레이 없음")
    r.geometry("1000x560+300+200")
    r.update()
    yield r
    r.destroy()


@pytest.mark.parametrize("key,expected", [
    ("y", True), ("Y", True), ("space", True), ("Return", True),
    # 숫자패드 Enter: Windows 의 Tk 는 Return 으로 전달하므로(위 Return 으로 확인) X11 에서만 따로 확인
    pytest.param("KP_Enter", True, marks=pytest.mark.skipif(sys.platform.startswith("win"),
                                                            reason="Windows 는 숫자패드 Enter 도 Return")),
    ("n", False), ("Escape", False)])
def test_yes_no_keys(shown_root, key, expected):
    from motor_monitor.ui import dialogs, msgbox
    seen = {}

    def press():
        m = [w for w in shown_root.winfo_children() if isinstance(w, dialogs.MessageDialog)][-1]
        seen["hint"] = m.lbl_keys.cget("text")
        m.focus_force()
        m.update()
        m.event_generate(f"<KeyPress-{key}>")

        def give_up():                       # 키가 먹지 않으면 시험이 멈추지 않고 실패하도록 닫음
            if m.winfo_exists():
                seen["timeout"] = True
                m.destroy()
        shown_root.after(2000, give_up)
    shown_root.after(300, press)
    answer = msgbox.ask_yes_no("확인", "진행할까요?", parent=shown_root)
    assert not seen.get("timeout") and answer is expected
    assert seen["hint"] == "Enter·Space·Y = 예   Esc·N = 아니오"


def test_yes_no_korean_keys_and_ok_dialog(shown_root):
    from motor_monitor.ui import dialogs
    d = dialogs.MessageDialog(shown_root, "확인", "진행할까요?", "question", ("예", "아니오"))
    d._on_key(SimpleNamespace(char="ㅛ", keysym="Hangul_YO"))          # 한글 입력 상태의 Y
    assert d.result == 0
    d = dialogs.MessageDialog(shown_root, "확인", "진행할까요?", "question", ("예", "아니오"))
    d._on_key(SimpleNamespace(char="ㅜ", keysym="Hangul_U"))           # 한글 입력 상태의 N
    assert d.result == 1
    d = dialogs.MessageDialog(shown_root, "오류", "알림", "error", ("확인",))
    assert d.lbl_keys.cget("text") == "Enter·Space·Y = 확인"
    d._on_key(SimpleNamespace(char="n", keysym="n"))                  # 확인 창에서 N 은 무시
    assert d.result is None and d.winfo_exists()
    d._key_default()
    assert d.result == 0


def test_problem_rows_are_shaded(shown_root):
    import tkinter as tk
    from motor_monitor.ui import theme as T
    from motor_monitor.ui.widgets import MotorListView
    ok = Motor(1, MotorType.NEW_MOTOR, True, status=MotorStatus.OK, rpm=800)
    comm = Motor(2, MotorType.NEW_MOTOR, True, status=MotorStatus.OK, had_err=True, err_note="192.168.0.4 접속 실패")
    old = Motor(3, MotorType.OLD_MOTOR, True, status=MotorStatus.ERR)
    slow = Motor(4, MotorType.NEW_MOTOR, True, status=MotorStatus.OK, rpm=700, target_rpm=800, cmd_rpm=824,
                 speed_note="설정 800 RPM 보다 느림")
    off = Motor(5, MotorType.NEW_MOTOR, False, status=MotorStatus.NOTUSE)
    frame = tk.Frame(shown_root)
    frame.pack(fill="both", expand=True)
    lv = MotorListView(frame, T.Images(shown_root), on_header_click=lambda k: None, on_double_click=lambda k: None)
    lv.pack(fill="both", expand=True)
    shown_root.update()
    lv.set_rows([ok, comm, old, slow, off])
    fills = {lv.canvas.gettags(i)[1]: lv.canvas.itemcget(i, "fill") for i in lv.canvas.find_withtag("rowbg")}
    assert fills == {"row1": T.LIST_PROBLEM_BG, "row2": T.LIST_PROBLEM_BG, "row3": T.LIST_PROBLEM_BG}
    assert MotorListView.row_fill(comm, True, False) == T.LIST_PROBLEM_SELECTED
    assert MotorListView.row_fill(comm, False, True) == T.LIST_PROBLEM_HOVER
    assert MotorListView.row_fill(ok, True, False) == T.LIST_SELECTED and MotorListView.row_fill(ok, False, False) is None
    icon, tip = MotorListView.status_icon(slow)
    assert icon == "bad_haderr" and tip.startswith("속도 미달\n설정 800 RPM 보다 느림")
    assert "설정 800 · 지정 824 · 실제 700 RPM" in tip
    assert MotorListView.status_icon(ok) == ("good", "정상")
    ok.target_rpm, ok.cmd_rpm = 800, 800
    assert MotorListView.status_icon(ok)[1] == "정상\n설정 800 · 지정 800 · 실제 800 RPM (오차 1% 이내)"


# ---------------------------------------------------------------- 추가: 모터 최고속도(h8) 기준 %, 속도 보낼 때 깜빡임
def test_ctrl_motor_reads_max_speed_and_shows_percent(tmp_path, fast):
    """ECblue 최고속도(h8) 600 인데 설정 700: 설정 700 을 그대로 보내고(모터가 600 으로 제한) 알림.
    최고속도 대비 % 를 보여줌."""
    gw = FakeGateway().start()
    gw.add_ecblue(2, analog_rpm=0, h4=0x0101)
    gw.ecb[2]["h"][8] = 600
    try:
        p = _plant(tmp_path)
        h = p.houses["H1R1"]
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, set_rpm=700)
        key = f"2_2@{gw.host}:{gw.port}"
        p.start_line("H1R1")
        assert _wait(lambda: "최고속도(h8) 600" in h.get_motor(key).speed_note)
        m = h.get_motor(key)
        assert m.max_rpm == 600 and gw.ecb[2]["h"][2] == 700 and gw.ecb_speed(2) == 600 and m.has_problem
        assert any("최고속도(h8) 600" in e.message for e in p.errors.recent())
        from motor_monitor.ui.widgets import MotorListView
        assert "모터 최고속도(h8) 600 RPM: 설정 700 RPM = 최고속도의 116.7%" in MotorListView.type_icon(m)[1]
        h.mod_motor(key, True, set_rpm=450)                         # 최고속도 안: 알림 해제
        assert _wait(lambda: gw.ecb[2]["h"][2] == 450 and h.get_motor(key).speed_note == "")
        assert "설정 450 RPM = 최고속도의 75%" in MotorListView.type_icon(h.get_motor(key))[1]
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_max_speed_alert_only_in_control_mode(tmp_path, fast):
    """모니터링 모드: h8 을 읽어 % 는 보여 주지만, 속도를 보내지 않으므로 '최고속도로 운전' 알림은 판단하지 않음.
    제어 모드로 바꾸면 700 을 그대로 보내고 알림 (2.5.0 과 같음)."""
    gw = FakeGateway().start()
    gw.add_ecblue(2, analog_rpm=0, h4=0x0101)
    gw.ecb[2]["h"][8] = 600
    try:
        p = _plant(tmp_path, monitoring=True)
        h = p.houses["H1R1"]
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, set_rpm=700)
        key = f"2_2@{gw.host}:{gw.port}"
        p.start_line("H1R1")
        assert _wait(lambda: h.get_motor(key).max_rpm == 600 and h.get_motor(key).target_rpm == 700)
        time.sleep(0.5)                                             # 점검 여러 번
        assert h.get_motor(key).speed_note == ""
        assert not any("최고속도(h8)" in e.message for e in p.errors.recent())
        assert gw_writes(gw) == []                                  # 모니터링 모드: 쓰지 않음
        p.set_monitoring(False)
        assert _wait(lambda: gw.ecb[2]["h"][2] == 700 and "최고속도(h8) 600" in h.get_motor(key).speed_note)
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_speed_send_marks_motor_for_blink(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_new(5)
    try:
        p = _plant(tmp_path)
        _zone(p, gw)
        h = p.houses["H1R1"]
        h.add_motor(5, MotorType.NEW_MOTOR, True)
        assert h.get_motor("1_5").sent_at is None
        p.start_line("H1R1")
        assert _wait(lambda: h.get_motor("1_5").sent_at is not None)      # 속도(0x16)를 보냄
        first = h.get_motor("1_5").sent_at
        time.sleep(0.5)
        assert h.get_motor("1_5").sent_at == first                       # 같은 값 -> 다시 보내지 않음, 깜빡임 없음
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_icon_blinks_red_twice_in_one_second(shown_root):
    import tkinter as tk
    from motor_monitor.ui import theme as T
    from motor_monitor.ui.widgets import MotorListView
    frame = tk.Frame(shown_root)
    frame.pack(fill="both", expand=True)
    lv = MotorListView(frame, T.Images(shown_root), on_header_click=lambda k: None, on_double_click=lambda k: None)
    lv.pack(fill="both", expand=True)
    shown_root.update()
    now = [100.0]
    lv._clock = lambda: now[0]
    m = Motor(5, MotorType.NEW_MOTOR, True, status=MotorStatus.OK, rpm=800, sent_at=time.monotonic())
    other = Motor(6, MotorType.NEW_MOTOR, True, status=MotorStatus.OK, rpm=800)
    lv.set_rows([m, other])                                           # 속도를 보낸 모터: 깜빡임 시작
    seen = []
    for t in (0.0, 0.1, 0.3, 0.6, 0.8, 1.05):                       # 0~0.25 붉게, 0.25~0.5 원래, 0.5~0.75 붉게
        now[0] = 100.0 + t
        lv._redraw()
        seen.append(tuple(lv.drawn_icons))
    assert [s[0] for s in seen] == ["newmotor_red", "newmotor_red", "newmotor", "newmotor_red", "newmotor", "newmotor"]
    assert all(s[1] == "newmotor" for s in seen)                     # 다른 모터는 그대로
    lv.set_rows([m, other])                                           # 같은 송신은 다시 깜빡이지 않음
    now[0] = 101.6
    lv._redraw()
    assert lv.drawn_icons[0] == "newmotor"
    for name in ("oldmotor_red", "newmotor_red", "newmotor_limit_red", "oldmotor_ctrl_red"):
        assert lv.images.get(name, 20) is not None                  # 붉은 그림 파일
    # 실제 시계로 1초 동안 저절로 깜빡였다가 원래대로 돌아옴
    lv._clock = time.monotonic
    m2 = Motor(7, MotorType.NEW_MOTOR, True, status=MotorStatus.OK, rpm=800, sent_at=time.monotonic())
    lv.set_rows([m2])
    shown_root.update()
    assert lv.drawn_icons == ["newmotor_red"]
    end = time.time() + 1.5
    while time.time() < end:
        shown_root.update()
        time.sleep(0.02)
    assert lv.drawn_icons == ["newmotor"] and lv._blink_job is None
