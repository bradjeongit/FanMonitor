"""운영 프로그램 요청 시험 (2.7.0):
1 구역설정에 기존모터 IP 2개, 2 기존모터 등록 시 IP 1/IP 2 선택, 3 RPM 제어 소스는 모터마다,
4 모터별 '모니터링만', 5 기존모터는 속도를 쓰지 않음(읽기만), 6 작동 안 하는 팬은 행을 붉게,
7 목록 IP 열(마지막 숫자), 8 통신 끊김 시 자동 운전(240초 / 700 RPM) 강제 설정."""
import time

import pytest

from motor_monitor import config as cfgmod
from motor_monitor import constants as C
from motor_monitor import modbus
from motor_monitor.constants import CTRL_MODE_REF, MotorStatus, MotorType
from motor_monitor.crypto3des import decrypt_config
from motor_monitor.house import Motor, ip_tail, motor_key, wd_auto_ok
from motor_monitor.plant import Plant
from motor_monitor.simulator import FakeGateway

WD_ON = (240 << 8) | 2


@pytest.fixture
def fast(monkeypatch):
    for name, val in [("SEND_INTERVAL", 0.01), ("CYCLE_INTERVAL", 0.05), ("RECONNECT_DELAY", 0.2),
                      ("ERROR_RETRY_DELAY", 0.05), ("CHECK_INTERVAL", 0.15)]:
        monkeypatch.setattr(C, name, val)


def _wait(cond, timeout=6.0):
    end = time.time() + timeout
    while time.time() < end:
        if cond():
            return True
        time.sleep(0.02)
    return False


def _plant(tmp_path, monitoring=False):
    p = Plant(str(tmp_path / "Config.dll"), monitoring=monitoring)
    p.load()
    return p


def _writes(gw, uid=None):
    return [(u, r, v) for u, r, v in list(gw.writes) if uid is None or u == uid]


# ---------------------------------------------------------------- 8 자동 운전 상수와 판단
def test_watchdog_constants_and_rules():
    assert (C.WD_SECONDS, C.WD_MODE, C.WD_SPEED) == (240, 2, 700)          # 5분은 h17 최대 255초를 넘어 240초
    assert C.WD_VALUE == WD_ON == 0xF002
    assert (modbus.ECB_H_SETINTERN1, modbus.ECB_H_WATCHDOG) == (5, 17)
    assert wd_auto_ok(700, WD_ON) and wd_auto_ok(500, (60 << 8) | 3)       # 정속 운전이 켜져 있음
    assert wd_auto_ok(0, (60 << 8) | 7)                                     # E1 제어로 전환도 '자동 운전'
    assert not wd_auto_ok(700, 0) and not wd_auto_ok(700, 240 << 8)        # 꺼짐
    assert not wd_auto_ok(700, (60 << 8) | 1)                               # 경보만: 자동 운전 아님
    assert not wd_auto_ok(0, WD_ON)                                         # 정속 0 RPM = 멈춤
    assert not wd_auto_ok(None, None)


# ---------------------------------------------------------------- 1·2·3·4 설정 파일
def test_config_roundtrip_ip2_gateway_reference_and_monitor_only(tmp_path):
    p = _plant(tmp_path)
    p.set_line_config("H1R1", p.configs["H1R1"].copy(source_ip="172.27.0.84", source_ip2="172.27.0.85"))
    h = p.houses["H1R1"]
    assert h.add_motor(7, MotorType.OLD_MOTOR, True)                        # IP 1 의 7
    assert h.add_motor(7, MotorType.OLD_MOTOR, True, gw=2)                  # IP 2 의 같은 ID 7 (다른 모터)
    assert not h.add_motor(7, MotorType.OLD_MOTOR, True, gw=2)              # 같은 통신장치·ID 는 하나만
    assert h.get_motor("0_7").gw == 1 and h.get_motor("0_7#2").gw == 2
    h.add_motor(3, MotorType.NEW_MOTOR, True, ctrl_mode=CTRL_MODE_REF, ref_line="H1R1", ref_gw=2, ref_id=7)
    h.add_motor(4, MotorType.NEW_MOTOR, True, set_rpm=650, mon_only=True)
    h.add_motor(180, MotorType.OLD_CTRL, True, ip="172.27.0.95", set_rpm=700, mon_only=True)
    p.save()
    text = decrypt_config(open(p.config_path, "rb").read())
    assert "H1R1_SOURCE_IP2=172.27.0.85" in text
    assert "H1R1_MOTORS=7=0=1=0=0,7=0=1=0=0,3=1=1=0=0,4=1=1=0=0\r\n" in text      # 운영 형식 목록은 그대로
    assert ("H1R1_MOTORS=0_7/GW:1,0_7/GW:2,1_3/SRC:R/RPM:770/REF:H1R1-2-7/MON:0,1_4/SRC:M/RPM:650/MON:1,"
            "2_180@172.27.0.95/SRC:M/RPM:700/MON:1") in text
    q = _plant(tmp_path)
    hq = q.houses["H1R1"]
    assert q.configs["H1R1"].source_ips() == ["172.27.0.84", "172.27.0.85"]
    assert [m.entry() for m in hq.get_list()] == [m.entry() for m in h.get_list()]
    m = hq.get_motor("1_3")
    assert (m.ctrl_mode, m.ref_line, m.ref_gw, m.ref_id) == (CTRL_MODE_REF, "H1R1", 2, 7)
    assert hq.get_motor("1_4").mon_only and hq.get_motor("2_180@172.27.0.95:502").mon_only


def test_old_zone_reference_moves_to_each_motor():
    """2.6.0 파일: 구역 RPM 제어 소스 -> 신규모터마다, 기존모터제어 '구역 참조' -> 그 구역이 따르던 값."""
    text = ("[LINE_CONFIG]\r\nH1R1_SOURCE_MANUAL=0\r\nH1R1_SOURCE_LINE=H1R2\r\nH1R1_SOURCE_ID=66\r\n"
            "H1R2_SOURCE_MANUAL=1\r\nH1R2_MANUAL_RPM=640\r\n"
            "[MOTOR_CONFIG]\r\nH1R1_MOTORS=5=1=1=0=0\r\nH1R2_MOTORS=66=0=1=0=0,6=1=1=0=0\r\n"
            "[CTRL_MOTOR_CONFIG]\r\nH1R1_MOTORS=180@172.27.0.95:502=1=0=REF\r\n"
            "H1R2_MOTORS=181@172.27.0.95:502=1=0=REF,182@172.27.0.95:502=1=500\r\n")
    _, motors = cfgmod.config_from_text(text)
    by = {(s, m.base_key): m for s, ms in motors.items() for m in ms}
    assert (by["H1R1", "1_5"].ctrl_mode, by["H1R1", "1_5"].ref_line, by["H1R1", "1_5"].ref_id) == ("ref", "H1R2", 66)
    assert (by["H1R2", "1_6"].ctrl_mode, by["H1R2", "1_6"].set_rpm) == ("manual", 640)
    c = by["H1R1", "2_180@172.27.0.95"]
    assert (c.ctrl_mode, c.ref_line, c.ref_id) == ("ref", "H1R2", 66)        # 구역이 따르던 기존모터
    c = by["H1R2", "2_181@172.27.0.95"]
    assert (c.ctrl_mode, c.set_rpm) == ("manual", 640)                       # 구역이 수동 -> 그 RPM 수동
    assert by["H1R2", "2_182@172.27.0.95"].set_rpm == 500                    # 수동 지정은 그대로


# ---------------------------------------------------------------- 1 기존모터 통신장치 2개
def test_two_old_motor_gateways(tmp_path, fast):
    g1, g2, gn = FakeGateway().start(), FakeGateway().start(), FakeGateway().start()
    g1.add_old(7, 770)
    g2.add_old(7, 640)                                                      # 같은 ID 7, 다른 통신장치
    g2.add_old(9, 0)                                                        # 멈춘 팬
    gn.add_new(3)
    try:
        p = _plant(tmp_path)
        p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip=g1.host, source_port=g1.port, source_ip2="127.0.0.1",
                                                   target_ip=gn.host, target_port=gn.port)
        h = p.houses["H1R1"]
        h.add_motor(7, MotorType.OLD_MOTOR, True)
        h.add_motor(7, MotorType.OLD_MOTOR, True, gw=2)
        h.add_motor(9, MotorType.OLD_MOTOR, True, gw=2)
        h.add_motor(3, MotorType.NEW_MOTOR, True, ctrl_mode=CTRL_MODE_REF, ref_line="H1R1", ref_gw=2, ref_id=7)
        # 시뮬레이터는 같은 주소에 PORT 만 달라, IP 2 통신장치의 접속 PORT 를 g2 로 바꿔 시험 (현장은 둘 다 502)
        orig = h._connect

        def connect(ch, ip, port, stop, mtype):
            return orig(ch, ip, g2.port if ch.gw == 2 else port, stop, mtype)
        h._connect = connect
        p.start_line("H1R1")
        assert _wait(lambda: h.get_motor("0_7").rpm == 770 and h.get_motor("0_7#2").rpm == 640)
        assert _wait(lambda: h.gw_net_status(1) and h.gw_net_status(2))
        assert _wait(lambda: gn.new_regs[3][modbus.NEW_REG_RPM] == 640)    # IP 2 의 7 을 참조
        m9 = h.get_motor("0_9#2")
        assert m9.status == MotorStatus.OK and m9.rpm == 0 and m9.not_running   # 6: 회전수 0 = 작동 안 함
        assert not h.get_motor("0_7").not_running
        rows = {m.key: m for m in h.get_list()}
        assert rows["0_7"].disp_ip == g1.host and rows["0_7#2"].disp_ip == "127.0.0.1"
        assert rows["0_7#2"].label == "기존모터 ID 7 (IP 2)"
        # 5: 기존모터에는 속도 쓰기가 없음 (자동 운전 설정 h0·h5·h17 만 - 여기 기존모터는 ECblue 가 아니라 그것도 없음)
        assert _writes(g1) == [] and _writes(g2) == []
        h.service_stop(join=True)
    finally:
        for g in (g1, g2, gn):
            g.stop()


def test_gateway2_down_marks_only_its_motors(tmp_path, fast):
    g1 = FakeGateway().start()
    g1.add_old(7, 770)
    try:
        p = _plant(tmp_path)
        p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip=g1.host, source_port=g1.port, source_ip2="127.0.0.1")
        h = p.houses["H1R1"]
        h.add_motor(7, MotorType.OLD_MOTOR, True)
        h.add_motor(8, MotorType.OLD_MOTOR, True, gw=2)
        orig = h._connect
        h._connect = lambda ch, ip, port, stop, mtype: orig(ch, ip, 1 if ch.gw == 2 else port, stop, mtype)
        p.start_line("H1R1")
        assert _wait(lambda: h.get_motor("0_8#2").had_err)                 # IP 2 접속 실패 -> IP 2 의 모터만
        assert _wait(lambda: h.get_motor("0_7").status == MotorStatus.OK)
        assert not h.get_motor("0_7").had_err and h.gw_net_status(1) and not h.gw_net_status(2)
        assert h.get_motor("0_8#2").not_running
        assert any(e.target == "기존모터 통신장치 2" for e in p.errors.recent())
        h.service_stop(join=True)
    finally:
        g1.stop()


def test_no_ip2_means_no_second_connection(tmp_path, fast):
    p = _plant(tmp_path)
    h = p.houses["H1R1"]
    tried = []
    orig = h._connect
    h._connect = lambda ch, ip, port, stop, mtype: (tried.append(ch.name), orig(ch, ip, 1, stop, mtype))[1]
    p.start_line("H1R1")
    time.sleep(0.4)
    h.service_stop(join=True)
    assert "S2" not in tried and "S1" in tried


# ---------------------------------------------------------------- 8 자동 운전 강제 설정
def test_old_motor_watchdog_is_forced_when_off_even_in_monitoring_mode(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_ecblue(7, analog_rpm=640)                     # 기존모터(ECblue, 아날로그 제어), h17 꺼짐
    gw.add_ecblue(8, analog_rpm=600)
    gw.ecb[8]["h"][5], gw.ecb[8]["h"][17] = 800, (120 << 8) | 2           # 이미 켜져 있음 -> 그대로
    try:
        p = _plant(tmp_path, monitoring=True)
        p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip=gw.host, source_port=gw.port)
        h = p.houses["H1R1"]
        h.add_motor(7, MotorType.OLD_MOTOR, True)
        h.add_motor(8, MotorType.OLD_MOTOR, True)
        p.start_line("H1R1")
        assert _wait(lambda: h.get_motor("0_7").wd_state == "set" and h.get_motor("0_8").wd_state == "ok")
        assert (gw.ecb[7]["h"][5], gw.ecb[7]["h"][17]) == (700, WD_ON)
        assert _writes(gw, 7) == [(7, 0, 10), (7, 5, 700), (7, 17, WD_ON)]   # PIN, 정속 700, 240초
        assert _writes(gw, 8) == []
        assert h.get_motor("0_7").rpm == 640                                # 속도는 쓰지 않음 (아날로그 그대로)
        assert "240초 통신 없으면 정속 700 RPM (이 프로그램이 설정함)" in h.get_motor("0_7").wd_desc
        assert any("자동 운전 설정함" in e.message and e.target == "기존모터 ID 7" for e in p.errors.recent())
        time.sleep(0.5)
        assert len(_writes(gw)) == 3                                        # 다시 쓰지 않음 (영구저장 수명)
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_old_motor_watchdog_gives_up_quietly_on_non_ecblue(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_old(7, 770)                                   # ECblue 아님: FC03 에 응답 없음
    try:
        p = _plant(tmp_path)
        p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip=gw.host, source_port=gw.port)
        h = p.houses["H1R1"]
        h.add_motor(7, MotorType.OLD_MOTOR, True)
        p.start_line("H1R1")
        assert _wait(lambda: h.get_motor("0_7").wd_state == "fail")
        m = h.get_motor("0_7")
        assert m.wd_tries == C.WD_MAX_TRIES and m.status == MotorStatus.OK and not m.had_err   # 상태에는 영향 없음
        errs = [e for e in p.errors.recent() if "자동 운전" in e.message]
        assert len(errs) == 1 and "응답 없음" in errs[0].message
        assert _writes(gw) == []
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_ctrl_watchdog_forced_for_monitor_only_motor(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_ecblue(2, analog_rpm=450, h4=0)
    try:
        p = _plant(tmp_path)                                # 제어 모드
        h = p.houses["H1R1"]
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, set_rpm=900, mon_only=True)
        key = f"2_2@{gw.host}:{gw.port}"
        p.start_line("H1R1")
        assert _wait(lambda: h.get_motor(key).wd_state == "set")
        time.sleep(0.4)
        assert {r for _, r, _ in _writes(gw)} == {0, 5, 17}                # 4: 속도·제어모드는 쓰지 않음
        assert gw.ecb[2]["h"][17] == WD_ON and gw.ecb[2]["h"][5] == 700
        m = h.get_motor(key)
        assert m.rpm == 450 and m.target_rpm is None
        assert [r.speed_src for r in h.get_list()] == ["mon"]
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_simulated_watchdog_runs_700_after_silence(tmp_path, fast, monkeypatch):
    """설정한 자동 운전이 실제로 동작하는지 (시뮬레이터): 통신이 끊긴 뒤 시간이 지나면 700 RPM."""
    gw = FakeGateway().start()
    gw.add_ecblue(2, analog_rpm=450, h4=0x0101)
    try:
        p = _plant(tmp_path)
        h = p.houses["H1R1"]
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, set_rpm=900)
        key = f"2_2@{gw.host}:{gw.port}"
        p.start_line("H1R1")
        assert _wait(lambda: h.get_motor(key).wd_state == "set" and gw.ecb_speed(2) == 900)
        h.service_stop(join=True)
        d = gw.ecb[2]
        d["last_seen"] -= C.WD_SECONDS + 1                  # 240초 넘게 통신 없음
        gw._ecb_watchdog(d)
        assert d["wd_tripped"] and gw.ecb_speed(2) == 700
    finally:
        gw.stop()


# ---------------------------------------------------------------- 4 모니터링만 (신규모터)
def test_monitor_only_new_motor_gets_no_writes(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_new(3, automanual=0, startstop=0, rpm=100)
    gw.add_new(4, automanual=0, startstop=0, rpm=100)
    try:
        p = _plant(tmp_path)
        p.configs["H1R1"] = p.configs["H1R1"].copy(target_ip=gw.host, target_port=gw.port)
        h = p.houses["H1R1"]
        h.add_motor(3, MotorType.NEW_MOTOR, True, set_rpm=800, mon_only=True)
        h.add_motor(4, MotorType.NEW_MOTOR, True, set_rpm=800)
        p.start_line("H1R1")
        assert _wait(lambda: gw.new_regs[4][modbus.NEW_REG_RPM] == 800)
        assert _wait(lambda: h.get_motor("1_3").status == MotorStatus.OK)
        h.mod_motor("1_3", False, mon_only=True)            # 사용 중지해도 RPM 0 을 보내지 않음
        time.sleep(0.5)
        assert _writes(gw, 3) == [] and gw.new_regs[3][modbus.NEW_REG_RPM] == 100
        assert h.get_motor("1_3").target_rpm is None
        # 모니터링만을 끄면 바로 제어
        h.mod_motor("1_3", True, set_rpm=800, mon_only=False)
        assert _wait(lambda: gw.new_regs[3][modbus.NEW_REG_RPM] == 800)
        h.service_stop(join=True)
    finally:
        gw.stop()


# ---------------------------------------------------------------- 5 기존모터는 읽기만 (제어 모드에서도)
def test_old_motor_never_gets_speed_writes_in_control_mode(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_ecblue(7, analog_rpm=640, h4=0)
    gw.ecb[7]["h"][5], gw.ecb[7]["h"][17] = 700, WD_ON
    try:
        p = _plant(tmp_path)
        p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip=gw.host, source_port=gw.port)
        h = p.houses["H1R1"]
        h.add_motor(7, MotorType.OLD_MOTOR, True)
        p.start_line("H1R1")
        assert _wait(lambda: h.get_motor("0_7").rpm == 640 and h.get_motor("0_7").wd_state == "ok")
        h.mod_motor("0_7", False)                           # 사용 중지해도 아무것도 쓰지 않음
        time.sleep(0.5)
        assert _writes(gw) == [] and gw.ecb[7]["h"][4] == 0
        reqs = {(r[7], r[8] * 256 + r[9]) for r in gw.requests}
        assert reqs <= {(4, 14), (3, 5), (3, 17)}            # 회전수(i14) 읽기와 자동 운전 확인 읽기만
        h.service_stop(join=True)
    finally:
        gw.stop()


# ---------------------------------------------------------------- 6·7 화면
def test_not_running_rules():
    ok = Motor(1, MotorType.NEW_MOTOR, True, status=MotorStatus.OK, rpm=770)
    assert not ok.not_running
    assert Motor(1, MotorType.NEW_MOTOR, True, status=MotorStatus.OK, rpm=0).not_running        # 회전수 0
    assert Motor(1, MotorType.OLD_MOTOR, True, status=MotorStatus.ERR, rpm=770).not_running     # 오래 응답 없음
    assert Motor(1, MotorType.OLD_MOTOR, True, status=MotorStatus.READY, had_err=True,
                 err_note="접속 실패").not_running                                               # 통신 에러
    assert not Motor(1, MotorType.OLD_MOTOR, True, status=MotorStatus.READY).not_running       # 아직 응답 전
    assert not Motor(1, MotorType.NEW_MOTOR, False, status=MotorStatus.NOTUSE).not_running     # 사용 안 함
    ctrl = Motor(2, MotorType.OLD_CTRL, True, status=MotorStatus.OK, rpm=500, had_err=True,
                 err_note="운전 허가(h1) 쓰기 실패", ctrl_note="운전 허가(h1) 쓰기 실패")
    assert not ctrl.not_running and ctrl.has_problem                       # 제어 문제만: 붉은 음영
    assert (ip_tail("172.27.0.84"), ip_tail(""), motor_key(MotorType.OLD_MOTOR, 7, gw=2)) == ("84", "-", "0_7#2")


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


def test_list_ip_column_red_rows_and_ip_sort(root, tmp_path):
    from motor_monitor.ui import theme as T
    from motor_monitor.ui import main_window
    p = _plant(tmp_path)
    p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip="172.27.0.84", source_ip2="172.27.0.9",
                                               target_ip="172.27.0.116")
    h = p.houses["H1R1"]
    h.add_motor(5, MotorType.OLD_MOTOR, True)
    h.add_motor(6, MotorType.OLD_MOTOR, True, gw=2)
    h.add_motor(2, MotorType.NEW_MOTOR, True)
    h.add_motor(180, MotorType.OLD_CTRL, True, ip="172.27.0.95")
    with h._lock:
        h._motors["0_5"].status, h._motors["0_5"].rpm = MotorStatus.OK, 0          # 멈춤 -> 붉게
        h._motors["0_6#2"].status, h._motors["0_6#2"].rpm = MotorStatus.OK, 640
        h._motors["1_2"].status, h._motors["1_2"].rpm = MotorStatus.ERR, 770        # 오래 응답 없음 -> 붉게
    win = main_window.MainWindow(root, p)
    lv = win.zone_panels["H1R1"].LV
    c = lv.canvas
    ips = [c.itemcget(i, "text") for i in c.find_withtag("ip")]
    assert [m.id for m in lv.rows] == [2, 5, 6, 180] and ips == ["116", "84", "9", "95"]
    fills = {c.gettags(i)[1]: c.itemcget(i, "fill") for i in c.find_withtag("rowbg")}
    assert fills == {"row0": T.LIST_DEAD_BG, "row1": T.LIST_DEAD_BG}
    assert lv.COLUMNS == ("", "IP", "ID", "상태")
    xs = lv._col_x()
    assert xs[2] - xs[1] == lv.COL_IP
    assert lv._tooltip_text(xs[1] + 3, lv.ROW_H // 2) == "접속 IP 172.27.0.116"
    win.zone_panels["H1R1"].on_sort("IP")                                   # IP 머리글 클릭: 마지막 숫자 순
    assert [m.id for m in lv.rows] == [6, 5, 180, 2]
    zp = win.zone_panels["H1R1"]
    assert (zp.lbl_old_ip.cget("text"), zp.lbl_old_ip2.cget("text"), zp.lbl_new_ip.cget("text")) == (
        "172.27.0.84", "172.27.0.9", "172.27.0.116")


def test_old_motor_add_dialog_picks_ip1_or_ip2(root, tmp_path, monkeypatch):
    from motor_monitor.ui import dialogs
    from motor_monitor.ui import theme as T
    msgs, asked = [], []
    monkeypatch.setattr(dialogs.BaseDialog, "error", lambda self, m, t: msgs.append(m))
    monkeypatch.setattr(dialogs.BaseDialog, "ask", lambda self, m, t: (asked.append(m), True)[1])
    p = _plant(tmp_path)
    p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip="172.27.0.84")
    d = dialogs.MotorDialog(root, p, T.Images(root), "H1R1", "ADD")
    d.motor_type_var.set(int(MotorType.OLD_MOTOR)); d._on_type()
    assert d.gw_row_visible and not d.speed_rows_visible and not d.limit_row_visible
    assert list(d.cb_gw.cget("values")) == ["IP 1  (172.27.0.84)"]          # IP 2 가 없으면 하나만
    d.destroy()
    p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip2="172.27.0.85")
    d = dialogs.MotorDialog(root, p, T.Images(root), "H1R1", "ADD")
    d.motor_type_var.set(int(MotorType.OLD_MOTOR)); d._on_type()
    assert list(d.cb_gw.cget("values")) == ["IP 1  (172.27.0.84)", "IP 2  (172.27.0.85)"]
    d.cb_gw.current(1)
    d.txtID.insert(0, "7")
    d.on_save()
    assert msgs == [] and p.houses["H1R1"].get_motor("0_7#2").gw == 2
    assert "읽기만" in asked[-1] and "240초" in asked[-1] and "700 RPM" in asked[-1]   # 자동 운전 안내


def test_new_motor_dialog_monitor_only_option(root, tmp_path, monkeypatch):
    from motor_monitor.ui import dialogs
    from motor_monitor.ui import theme as T
    asked = []
    monkeypatch.setattr(dialogs.BaseDialog, "error", lambda self, m, t: None)
    monkeypatch.setattr(dialogs.BaseDialog, "ask", lambda self, m, t: (asked.append(m), True)[1])
    p = _plant(tmp_path)
    d = dialogs.MotorDialog(root, p, T.Images(root), "H1R1", "ADD")
    d.motor_type_var.set(int(MotorType.NEW_MOTOR)); d._on_type()
    d.txtID.insert(0, "4")
    d.mon_var.set(1)
    d.on_save()
    m = p.houses["H1R1"].get_motor("1_4")
    assert m.mon_only and "모니터링만" in asked[-1]
    d = dialogs.MotorDialog(root, p, T.Images(root), "H1R1", "MOD", mod_dkey="1_4")
    assert d.mon_var.get() == 1
    d.mon_var.set(0)
    d.on_save()
    assert not p.houses["H1R1"].get_motor("1_4").mon_only
