"""운영 프로그램 요청 7건 시험:
1 목록은 ID 순서가 기본 (test_core 의 정렬 시험), 2 확인 창 키보드(Y/Enter/Space = 예, N/Esc = 아니오),
3 문제 상태 행 음영, 4 설정 속도 1% 이내는 오차로 인정, 5 느리면 조금씩 올리고 반응이 없으면 더 보내지 않음,
6 오차 안에 든 지정값을 기억해 다음 실행의 최초 신호로, 7 팬 속도는 30초에 한 번 점검 (설정을 바꾸면 바로)."""
import json
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
from motor_monitor.speedtrim import SpeedMemory, Trim, tolerance, within


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


# ---------------------------------------------------------------- 4·5·6 보정 규칙 (통신 없이)
def test_tolerance_is_one_percent():
    assert tolerance(800) == 8.0 and within(792, 800) and within(808, 800)
    assert not within(791, 800) and not within(809, 800)


def test_trim_within_tolerance_keeps_and_learns():
    t = Trim()
    assert t.decide(800, None, 1100) == 800                  # 처음: 설정 속도 그대로
    assert t.decide(800, 795, 1100) == 800 and t.learned     # 0.6% 차이 = 오차 인정, 기억할 값
    assert t.decide(800, 808, 1100) == 800 and t.learned     # +1% 도 오차 인정
    assert t.note == ""


def test_trim_raises_step_by_step_until_within():
    t = Trim()
    t.decide(800, None, 1100)
    assert t.decide(800, 760, 1100) == 824                    # 5% 느림 -> 한 번에 최대 3%(24) 올림
    assert t.decide(800, 783, 1100) == 841                    # 반응함(+23), 아직 2.1% 느림 -> 2%(16) 이상 올림
    assert t.decide(800, 799, 1100) == 841 and t.learned       # 오차 안 -> 그대로, 기억
    assert (t.target, t.cmd) == (800, 841)


def test_trim_stops_when_speed_does_not_change():
    t = Trim()
    t.decide(800, None, 1100)
    assert t.decide(800, 700, 1100) == 824                    # 올림
    assert t.decide(800, 701, 1100) == 824 and t.note == ""  # 바뀌지 않음 -> 다시 보내지 않고 기다림
    assert t.decide(800, 700, 1100) == 824                    # 60초(2번) 동안 그대로 -> 보정 멈춤, 알림
    assert t.hold == "stall" and "보정 멈춤" in t.note and "824" in t.note
    for _ in range(3):
        assert t.decide(800, 700, 1100) == 824                # 멈춘 뒤에는 더 올리지 않음
    assert t.decide(800, 798, 1100) == 824 and t.note == ""   # 속도가 돌아오면 알림 해제


def test_trim_respects_max():
    t = Trim()
    t.decide(500, None, 500)                                  # 속도제한 500 = 지정값 최대
    assert t.decide(500, 475, 500) == 500
    assert t.hold == "max" and "최대 500 RPM" in t.note
    t2 = Trim()
    t2.decide(1100, None, 1100)
    assert t2.decide(1100, 1000, 1100) == 1100 and t2.hold == "max"


def test_trim_lowers_compensation_when_too_fast_but_not_below_target():
    t = Trim()
    t.decide(800, None, 1100)
    t.decide(800, 760, 1100)                                  # 824 로 올림
    assert t.decide(800, 830, 1100) == 800                    # 너무 빠름 -> 올린 만큼 내림 (설정 아래로는 안 내림)
    assert t.decide(800, 820, 1100) == 800                    # 지정값 = 설정 속도면 더 내리지 않음


def test_trim_ignores_target_changes_within_one_percent():
    t = Trim()
    t.decide(800, None, 1100)
    t.decide(800, 760, 1100)                                  # 824
    assert t.decide(805, 780, 1100) == 844 and t.target == 800  # 805 는 800 의 1% 이내 -> 같은 설정으로 보고
                                                              # 반응(+20) 확인 뒤 남은 차이(20)만큼 더 올림
    assert t.decide(820, None, 1100) == 820 and t.target == 820  # 1% 넘게 바뀜 -> 새로 시작


def test_trim_uses_memory_for_first_signal():
    t = Trim()
    assert t.decide(800, None, 1100, memory=(800, 841)) == 841     # 같은 설정 속도: 기억한 지정값 그대로
    t = Trim()
    assert t.decide(805, None, 1100, memory=(800, 841)) == 841     # 1% 이내면 같은 설정 속도로 봄
    t = Trim()
    assert t.decide(600, None, 1100, memory=(800, 841)) == 631     # 다른 설정 속도: 같은 비율 (600 x 841/800)
    t = Trim()
    assert t.decide(1100, None, 1100, memory=(800, 841)) == 1100   # 최대값은 넘지 않음
    t = Trim()
    assert t.first_signal((800, 841)) == 841                       # 설정 속도를 아직 모를 때 최초 신호
    assert t.first_signal((800, 900)) == 841                       # 한 번만
    assert Trim().first_signal(None) is None
    assert Trim().decide(0, None, 1100, memory=(800, 841)) == 0      # 정지는 보정 없음


def test_speed_memory_file(tmp_path):
    path = tmp_path / "speed_memory.json"
    mem = SpeedMemory(str(path))
    assert mem.get("H1R1/1_5") is None
    mem.remember("H1R1/1_5", 800, 841)                          # 처음 기억은 바로 파일에
    assert json.loads(path.read_text(encoding="utf-8"))["H1R1/1_5"]["cmd"] == 841
    mem.remember("H1R1/1_6", 700, 720)                          # 30초 안의 다음 기억은 모았다가
    assert "H1R1/1_6" not in json.loads(path.read_text(encoding="utf-8"))
    mem.save()                                                  # 종료 때 반드시 저장
    assert SpeedMemory(str(path)).get("H1R1/1_6") == (700, 720)
    mem.remember("H1R1/1_7", 700, 650)                          # 설정보다 작은 지정값은 기억하지 않음
    assert mem.get("H1R1/1_7") is None
    path.write_text("{깨진 파일", encoding="utf-8")
    assert SpeedMemory(str(path)).get("H1R1/1_5") is None       # 읽지 못하면 새로 시작 (프로그램은 계속)


# ---------------------------------------------------------------- 4·5·6 실제 통신 (시뮬레이터)
def _zone(p, gw, rpm=800):
    p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip=gw.host, source_port=gw.port, target_ip=gw.host,
                                               target_port=gw.port, source_manual=True, manual_rpm=rpm)


def test_new_motor_slow_fan_is_corrected_and_remembered(tmp_path, fast):
    """설정 800, 실제 = 지정 x 0.95 인 신규모터: 지정값을 조금씩 올려 1% 안에 들게 하고(841), 그 값을 기억.
    다시 실행하면 최초 신호로 기억한 841 을 보냄."""
    gw = FakeGateway().start()
    gw.add_new(5)
    gw.new_gain[5] = 0.95
    try:
        p = _plant(tmp_path)
        _zone(p, gw)
        h = p.houses["H1R1"]
        h.add_motor(5, MotorType.NEW_MOTOR, True)
        p.start_line("H1R1")
        assert _wait(lambda: within(h.get_motor("1_5").rpm, 800) and h.get_motor("1_5").cmd_rpm == 841)
        assert _rpm_writes(gw, 5) == [800, 824, 841]               # 조금씩 올림
        m = h.get_motor("1_5")
        assert (m.target_rpm, m.speed_note) == (800, "") and not m.has_problem
        time.sleep(0.5)
        assert _rpm_writes(gw, 5) == [800, 824, 841]               # 오차 안 -> 더 보내지 않음
        p.shutdown()
        assert p.memory.get("H1R1/1_5") == (800, 841)
        assert json.loads((tmp_path / "speed_memory.json").read_text(encoding="utf-8"))["H1R1/1_5"]["cmd"] == 841
    finally:
        gw.stop()
    gw2 = FakeGateway().start()                                     # 다시 실행 (모터는 처음 상태)
    gw2.add_new(5)
    gw2.new_gain[5] = 0.95
    try:
        p2 = _plant(tmp_path)
        _zone(p2, gw2)
        p2.start_line("H1R1")
        assert _wait(lambda: _rpm_writes(gw2, 5))
        assert _rpm_writes(gw2, 5)[0] == 841                        # 최초 신호 = 기억한 지정값
        assert _wait(lambda: within(p2.houses["H1R1"].get_motor("1_5").rpm, 800))
        time.sleep(0.4)
        assert _rpm_writes(gw2, 5) == [841]                         # 바로 오차 안 -> 더 보내지 않음
        p2.shutdown()
    finally:
        gw2.stop()


def test_new_motor_no_response_stops_and_shows_problem(tmp_path, fast):
    """700 RPM 이상 돌지 못하는 팬(설정 800): 한 번 올린 뒤 60초(2번) 동안 그대로면 더 보내지 않고 문제 표시."""
    gw = FakeGateway().start()
    gw.add_new(6)
    gw.new_max[6] = 700
    try:
        p = _plant(tmp_path)
        _zone(p, gw)
        h = p.houses["H1R1"]
        h.add_motor(6, MotorType.NEW_MOTOR, True)
        p.start_line("H1R1")
        assert _wait(lambda: "보정 멈춤" in h.get_motor("1_6").speed_note)
        m = h.get_motor("1_6")
        assert m.has_problem and _rpm_writes(gw, 6) == [800, 824]
        assert any(e.target == "신규모터 ID 6" and "보정 멈춤" in e.message for e in p.errors.recent())
        time.sleep(0.6)
        assert _rpm_writes(gw, 6) == [800, 824]                     # 반복해서 보내지 않음
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_new_motor_limit_is_the_maximum(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_new(7)
    gw.new_gain[7] = 0.95
    try:
        p = _plant(tmp_path)
        _zone(p, gw)
        h = p.houses["H1R1"]
        h.add_motor(7, MotorType.NEW_MOTOR, True, True, 500)       # 속도제한 500
        p.start_line("H1R1")
        assert _wait(lambda: "최대 500 RPM" in h.get_motor("1_7").speed_note)
        assert set(_rpm_writes(gw, 7)) == {500}                     # 제한값을 넘겨 보내지 않음
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_target_change_within_one_percent_sends_nothing(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_new(5)
    try:
        p = _plant(tmp_path)
        _zone(p, gw)
        h = p.houses["H1R1"]
        h.add_motor(5, MotorType.NEW_MOTOR, True)
        p.start_line("H1R1")
        assert _wait(lambda: _rpm_writes(gw, 5) == [800] and within(h.get_motor("1_5").rpm, 800))
        p.set_line_config("H1R1", p.configs["H1R1"].copy(manual_rpm=805))   # 0.6% -> 다시 보내지 않음
        time.sleep(0.5)
        assert _rpm_writes(gw, 5) == [800]
        p.set_line_config("H1R1", p.configs["H1R1"].copy(manual_rpm=820))   # 2.5% -> 새 설정
        assert _wait(lambda: _rpm_writes(gw, 5) == [800, 820])
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_ctrl_motor_slow_fan_is_corrected(tmp_path, fast):
    """기존모터제어(ECblue) 수동 700, 실제 = h2 x 0.96: h2 를 조금씩 올려 1% 안에 들게 함."""
    gw = FakeGateway().start()
    gw.add_ecblue(2, analog_rpm=0, h4=0x0101)
    gw.ecb[2]["gain"] = 0.96
    try:
        p = _plant(tmp_path)
        h = p.houses["H1R1"]
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, set_rpm=700, ctrl_mode=CTRL_MODE_MANUAL)
        key = f"2_2@{gw.host}:{gw.port}"
        p.start_line("H1R1")
        assert _wait(lambda: within(gw.ecb_speed(2), 700) and h.get_motor(key).cmd_rpm == gw.ecb[2]["h"][2] > 700)
        h2s = [v for u, reg, v in gw_writes(gw) if u == 2 and reg == modbus.ECB_H_SPEED]
        assert h2s[0] == 700 and h2s == sorted(h2s) and all(b - a <= 21 for a, b in zip(h2s, h2s[1:]))
        assert _wait(lambda: p.memory.get(f"H1R1/{key}") == (700, gw.ecb[2]["h"][2]))   # 다음 점검에서 기억
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
def test_ctrl_motor_reads_max_speed_and_limits_to_it(tmp_path, fast):
    """ECblue 최고속도(h8) 600 인데 설정 700: 최고속도로 운전(h2 = 600)하고 알림. 최고속도 대비 % 를 보여줌."""
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
        assert m.max_rpm == 600 and gw.ecb[2]["h"][2] == 600 and m.has_problem
        from motor_monitor.ui.widgets import MotorListView
        assert "모터 최고속도(h8) 600 RPM: 설정 600 RPM = 최고속도의 100%" in MotorListView.type_icon(m)[1]
        h.mod_motor(key, True, set_rpm=450)                         # 최고속도 안: 알림 해제
        assert _wait(lambda: gw.ecb[2]["h"][2] == 450 and h.get_motor(key).speed_note == "")
        assert "설정 450 RPM = 최고속도의 75%" in MotorListView.type_icon(h.get_motor(key))[1]
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
        assert h.get_motor("1_5").sent_at == first                       # 오차 안 -> 더 보내지 않음, 깜빡임 없음
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
