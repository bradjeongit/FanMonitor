"""운영 프로그램 변경 요청 시험:
1·2 모터 ID 범위, 3·8 기존모터제어 구역 참조(신규모터가 따르는 값을 1100 대비 % 로), 4 아이콘 색,
5 모니터링 모드 쓰기 없음, 6 다시 연결되면 정상 표시, 7 에러 로그, 9 항상 모니터링 모드로 시작,
10 운영 프로그램·미니 제어기 설정 가져오기."""
import json
import os
import socket
import time

import pytest

from motor_monitor import config as cfgmod
from motor_monitor import constants as C
from motor_monitor.constants import CTRL_MODE_MANUAL, CTRL_MODE_REF, MotorStatus, MotorType, id_range
from motor_monitor.crypto3des import encrypt_config
from motor_monitor.errorlog import ErrorEntry, ErrorLog
from motor_monitor.plant import Plant
from motor_monitor.simulator import FakeGateway


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


def _plant(tmp_path, monitoring=False, log_dir=None):
    p = Plant(str(tmp_path / "Config.dll"), monitoring=monitoring, error_log_dir=log_dir)
    p.load()
    return p


def _free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def _gateway(port):
    end = time.monotonic() + 3
    while True:
        try:
            return FakeGateway(port=port).start()
        except OSError:
            if time.monotonic() > end:
                raise
            time.sleep(0.05)


# ---------------------------------------------------------------- 1·2 ID 범위
def test_id_ranges_by_type():
    assert id_range(MotorType.NEW_MOTOR) == (1, 30)
    assert id_range(MotorType.OLD_MOTOR) == (1, 249)
    assert id_range(MotorType.OLD_CTRL) == (1, 249)


# ---------------------------------------------------------------- 3·8 기존모터제어 구역 참조
def test_ctrl_ref_config_roundtrip_and_old_format(tmp_path):
    p = _plant(tmp_path)
    h = p.houses["H1R2"]
    h.add_motor(180, MotorType.OLD_CTRL, True, ip="172.27.0.95", set_rpm=600)
    h.add_motor(176, MotorType.OLD_CTRL, True, ip="172.27.0.95", ctrl_mode=CTRL_MODE_REF)
    p.save()
    text = cfgmod.read_config_text(p.config_path)
    assert "H1R2_MOTORS=180@172.27.0.95:502=1=600,176@172.27.0.95:502=1=0=REF" in text
    q = _plant(tmp_path)
    assert q.houses["H1R2"].get_motor("2_176@172.27.0.95:502").ctrl_mode == CTRL_MODE_REF
    assert q.houses["H1R2"].get_motor("2_180@172.27.0.95:502").ctrl_mode == CTRL_MODE_MANUAL
    # 예전 형식(3항목)은 수동 지정, 참조모터 IP/ID 를 적은 형식은 구역 참조로 읽음. PORT 는 502 고정
    assert cfgmod.parse_ctrl_list("2@1.2.3.4:502=1=700")[0].ctrl_mode == CTRL_MODE_MANUAL
    old = cfgmod.parse_ctrl_list("2@1.2.3.4:5020=1=0=R155@172.27.0.84:502")[0]
    assert (old.ctrl_mode, old.port) == (CTRL_MODE_REF, 502)
    assert cfgmod.parse_ctrl_list("3@1.2.3.4=1=500")[0].port == 502
    assert cfgmod.parse_ctrl_list("3@:502=1=500,4@1.2.3.4:502=1=500=X") == []      # 형식 오류는 무시


def test_reference_endpoint_follows_zone_reference(tmp_path):
    p = _plant(tmp_path)
    p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip="172.27.0.84")
    p.configs["H1R3"] = p.configs["H1R3"].copy(source_manual=False, source_line="H1R1", source_id=155)
    assert p.reference_endpoint("H1R3") == ("172.27.0.84", 502, 155)      # 신규모터가 참조하는 기존모터
    p.configs["H1R4"] = p.configs["H1R4"].copy(source_manual=True)
    assert p.reference_endpoint("H1R4") is None                           # 수동: 참조 기존모터 없음


def test_percent_conversion():
    for v in range(0, C.MAX_RPM + 1):
        assert C.pct_to_rpm(C.rpm_to_pct(v)) == v                          # 1100 대비 % -> 같은 RPM
    assert C.rpm_to_pct(770) == 70.0 and C.rpm_to_pct(1500) == 100.0 and C.rpm_to_pct(-5) == 0.0
    assert (C.pct_text(C.rpm_to_pct(770)), C.pct_text(C.rpm_to_pct(777)), C.pct_text(100)) == ("70%", "70.6%", "100%")


def test_ctrl_zone_ref_target_rules(tmp_path):
    """구역 참조값 -> 1100 대비 % -> 같은 비율. 값이 없으면 이전 값 유지(60초 넘으면 이상표시),
    자기 자신이면 따라가지 않음."""
    p = _plant(tmp_path)
    p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip="172.27.0.84")
    p.configs["H1R2"] = p.configs["H1R2"].copy(source_manual=False, source_line="H1R1", source_id=155)
    src = p.houses["H1R1"]
    src.add_motor(155, MotorType.OLD_MOTOR, True)
    h = p.houses["H1R2"]
    h.add_motor(2, MotorType.OLD_CTRL, True, ip="172.27.0.95", ctrl_mode=CTRL_MODE_REF)
    key = "2_2@172.27.0.95:502"
    now = [1000.0]
    h._clock = lambda: now[0]
    target = lambda: h._ctrl_target(h.get_motor(key))
    # 참조 기존모터를 아직 읽지 못함 -> 속도를 정하지 않고 기다림 (구역을 막 시작한 때는 이상표시 안 함)
    assert target() == (None, "")
    now[0] += C.NO_RESPONSE_WARN_SEC - 1
    assert target() == (None, "")
    now[0] += 1                                                             # 60초 넘게 값 없음 -> 이상표시
    t, note = target()
    assert t is None and "속도를 지정하지 않음" in note and "1구역 기존모터 155" in note
    src._mark_received(src._motors["0_155"], 770)
    assert target() == (770, "")                                            # 770 = 70% -> 1100 x 70%
    m = h.get_motor(key)
    assert (m.ref_rpm, m.ctrl_target, m.ref_desc) == (770, 770, "1구역 기존모터 155")
    src._motors["0_155"].rpm = 1500
    assert target() == (1100, "")                                           # 100% 넘으면 100%
    src._motors["0_155"].status = MotorStatus.ERR                           # 참조 기존모터 이상 -> 이전 값 유지
    assert target() == (1100, "")
    now[0] += C.NO_RESPONSE_WARN_SEC
    t, note = target()
    assert t == 1100 and "이전 지정 100% (1100 RPM) 유지" in note
    # 구역설정이 수동이면 그 수동 RPM 을 같은 비율로
    p.configs["H1R2"] = p.configs["H1R2"].copy(source_manual=True, manual_rpm=600)
    assert target() == (600, "") and h.get_motor(key).ref_desc == "구역 수동 600 RPM"
    # 참조 기존모터가 이 모터 자신 (같은 통신장치·ID) -> 따라가지 않음
    p.configs["H1R2"] = p.configs["H1R2"].copy(source_manual=False, source_line="H1R2", source_id=2,
                                               source_ip="172.27.0.95")
    t, note = target()
    assert t == 600 and "자신" in note
    # 수동 지정은 그대로
    h.mod_motor(key, True, set_rpm=450, ctrl_mode=CTRL_MODE_MANUAL)
    assert target() == (450, "")


def test_ctrl_ref_follows_zone_reference(tmp_path, fast, monkeypatch):
    """구역 참조: 신규모터가 참조하는 기존모터 값(1구역 기존모터 7)을 1100 대비 % 로 같은 비율로 지정.
    시작 직후 참조값을 읽기 전에는 에러 없이 기다림. 참조값이 없어지면 이전 속도 유지(이상표시·에러 기록),
    돌아오면 해제."""
    old = FakeGateway().start()                       # 1구역 기존모터 통신장치 (참조 기존모터 7)
    tgt = FakeGateway().start()                       # 기존모터제어 대상 ECblue
    old.add_old(7, 770)
    tgt.add_ecblue(2, analog_rpm=0, h4=0)
    try:
        p = _plant(tmp_path)
        p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip=old.host, source_port=old.port, target_ip=old.host,
                                                   target_port=old.port, source_manual=False, source_line="H1R1",
                                                   source_id=7)
        h = p.houses["H1R1"]
        h.add_motor(7, MotorType.OLD_MOTOR, True)
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=tgt.host, port=tgt.port, ctrl_mode=CTRL_MODE_REF)
        key = f"2_2@{tgt.host}:{tgt.port}"
        p.start_line("H1R1")
        assert _wait(lambda: tgt.ecb_speed(2) == 770)                    # 770 = 70% -> 770 RPM
        assert _wait(lambda: h.get_motor(key).ref_rpm == 770 and h.get_motor(key).ctrl_target == 770)
        assert not any("구역 참조값" in e.message for e in p.errors.recent())  # 시작 직후 잘못된 에러 없음
        monkeypatch.setattr(C, "NO_RESPONSE_WARN_SEC", 1)                 # 아래 '값 없음' 확인을 빨리
        old.old_rpm[7] = 880
        assert _wait(lambda: tgt.ecb_speed(2) == 880)
        old.old_rpm[7] = 1500                         # 100% 넘음 -> 1100
        assert _wait(lambda: tgt.ecb[2]["h"][2] == 1100)
        # 구역 참조값이 없어짐 (구역설정의 참조기존모터를 없는 모터로) -> 이전 속도 유지, 이상표시, 에러 기록
        p.configs["H1R1"] = p.configs["H1R1"].copy(source_id=99)
        assert _wait(lambda: "구역 참조값" in h.get_motor(key).ctrl_note)
        m = h.get_motor(key)
        assert m.had_err and "이전 지정 100% (1100 RPM) 유지" in m.ctrl_note and tgt.ecb[2]["h"][2] == 1100
        assert any("구역 참조값" in e.message and e.level == "에러" for e in p.errors.recent())
        # 다시 참조 -> 따라가고 이상표시 해제
        old.old_rpm[7] = 550
        p.configs["H1R1"] = p.configs["H1R1"].copy(source_id=7)
        assert _wait(lambda: tgt.ecb_speed(2) == 550)
        assert _wait(lambda: not h.get_motor(key).had_err and h.get_motor(key).ctrl_note == "")
        h.service_stop(join=True)
    finally:
        old.stop()
        tgt.stop()


def test_ctrl_ref_without_any_reading_does_not_write_speed(tmp_path, fast, monkeypatch):
    monkeypatch.setattr(C, "NO_RESPONSE_WARN_SEC", 1)
    gw = FakeGateway().start()
    gw.add_ecblue(2, analog_rpm=0, h4=0x0101)
    gw.ecb[2]["h"][2] = 333
    try:
        p = _plant(tmp_path)
        p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip=gw.host, source_port=gw.port, target_ip=gw.host,
                                                   target_port=gw.port, source_manual=False, source_line="H1R1",
                                                   source_id=99)
        h = p.houses["H1R1"]
        h.add_motor(99, MotorType.OLD_MOTOR, True)    # 응답하지 않는 참조 기존모터
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, ctrl_mode=CTRL_MODE_REF)
        key = f"2_2@{gw.host}:{gw.port}"
        p.start_line("H1R1")
        assert _wait(lambda: "속도를 지정하지 않음" in h.get_motor(key).ctrl_note)
        time.sleep(0.3)
        assert gw.ecb[2]["h"][2] == 333 and not any(w[1] == 2 for w in gw.writes)
        h.service_stop(join=True)
    finally:
        gw.stop()


# ---------------------------------------------------------------- 5 모니터링 모드
def test_monitoring_mode_never_writes_including_ref_mode(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_ecblue(2, analog_rpm=0, h4=0)
    gw.add_new(5, automanual=0, startstop=0, rpm=0)
    gw.add_old(7, 770)
    try:
        p = _plant(tmp_path, monitoring=True)
        p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip=gw.host, source_port=gw.port, target_ip=gw.host,
                                                   target_port=gw.port, source_manual=False, source_line="H1R1",
                                                   source_id=7)
        h = p.houses["H1R1"]
        h.add_motor(7, MotorType.OLD_MOTOR, True)
        h.add_motor(5, MotorType.NEW_MOTOR, True)
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, ctrl_mode=CTRL_MODE_REF)
        p.start_line("H1R1")
        key = f"2_2@{gw.host}:{gw.port}"
        assert _wait(lambda: h.get_motor(key).ref_rpm == 770 and h.get_motor("1_5").status == MotorStatus.OK)
        time.sleep(0.5)
        assert gw.writes == []                          # 어떤 쓰기도 없음
        h.service_stop(join=True)
    finally:
        gw.stop()


# ---------------------------------------------------------------- 6 다시 연결되면 정상 표시
def test_connection_lost_then_recovered_resets_icons(tmp_path, fast):
    port = _free_port()
    p = _plant(tmp_path)
    p.configs["H1R1"] = p.configs["H1R1"].copy(source_ip="127.0.0.1", source_port=port,
                                               target_ip="127.0.0.1", target_port=port)
    h = p.houses["H1R1"]
    h.add_motor(7, MotorType.OLD_MOTOR, True)
    h.add_motor(5, MotorType.NEW_MOTOR, True)
    p.start_line("H1R1")
    try:
        # 통신장치가 꺼져 있음 -> 이상표시 + 에러 기록(장치마다 한 번)
        assert _wait(lambda: h.get_motor("0_7").had_err and h.get_motor("1_5").had_err)
        assert "접속 실패" in h.get_motor("0_7").err_note
        time.sleep(0.6)                                  # 재시도가 여러 번 있어도 기록은 한 번
        errs = [e for e in p.errors.recent() if e.level == "에러"]
        assert sorted(e.target for e in errs) == ["기존모터 통신장치", "신규모터 통신장치"]
        # 통신장치가 켜지고 응답 -> 바로 정상 표시, 복구 기록
        gw = _gateway(port)
        try:
            gw.add_old(7, 770)
            gw.add_new(5)
            assert _wait(lambda: not h.get_motor("0_7").had_err and not h.get_motor("1_5").had_err)
            assert h.get_motor("0_7").status == MotorStatus.OK and h.get_motor("0_7").rpm == 770
            rec = [e for e in p.errors.recent() if e.level == "복구"]
            assert any(e.target == "기존모터 통신장치" and "다시 접속됨" in e.message for e in rec)
            assert any(e.target == "기존모터 ID 7" for e in rec)
            # 실행 중 연결이 끊김 -> 다시 이상표시, 다시 켜지면 정상
            gw.stop()
            assert _wait(lambda: h.get_motor("0_7").had_err)
            gw2 = _gateway(port)
            try:
                gw2.add_old(7, 780)
                gw2.add_new(5)
                assert _wait(lambda: not h.get_motor("0_7").had_err and h.get_motor("0_7").rpm == 780)
            finally:
                gw2.stop()
        finally:
            gw.stop()
    finally:
        h.service_stop(join=True)


def test_ctrl_gateway_down_and_up(tmp_path, fast):
    port = _free_port()
    p = _plant(tmp_path)
    h = p.houses["H1R1"]
    h.add_motor(2, MotorType.OLD_CTRL, True, ip="127.0.0.1", port=port, set_rpm=500)
    key = f"2_2@127.0.0.1:{port}"
    p.start_line("H1R1")
    try:
        assert _wait(lambda: h.get_motor(key).had_err)
        gw = _gateway(port)
        try:
            gw.add_ecblue(2, analog_rpm=0)
            assert _wait(lambda: gw.ecb_speed(2) == 500 and not h.get_motor(key).had_err)
            assert any(e.level == "복구" and e.target == "기존모터제어 통신장치" for e in p.errors.recent())
        finally:
            gw.stop()
    finally:
        h.service_stop(join=True)


# ---------------------------------------------------------------- 7 에러 로그
def test_error_log_file_reload_unseen_and_clear(tmp_path):
    el = ErrorLog(str(tmp_path))
    el.error("1구역", "기존모터 ID 5", "명령 후 60초 동안 응답 없음")
    el.recover("1구역", "기존모터 ID 5", "응답 정상")
    el.error("2구역", "신규모터 통신장치", "192.168.0.4 접속 실패")
    assert el.unseen == 2                                    # 복구는 세지 않음
    assert [e.level for e in el.recent()] == ["에러", "복구", "에러"]   # 최신이 앞
    el.mark_seen()
    assert el.unseen == 0
    # 다시 켜면 파일에서 지난 기록을 읽음
    el2 = ErrorLog(str(tmp_path))
    assert [e.target for e in el2.recent()] == ["신규모터 통신장치", "기존모터 ID 5", "기존모터 ID 5"]
    assert el2.unseen == 0
    el2.clear()
    assert el2.recent() == [] and os.path.getsize(os.path.join(str(tmp_path), "error.log")) > 0
    assert ErrorEntry.parse("깨진 줄") is None


# ---------------------------------------------------------------- 9 항상 모니터링 모드로 시작
def test_program_always_starts_in_monitoring(tmp_path):
    pytest.importorskip("tkinter")
    from motor_monitor.app import create_plant
    p = create_plant(str(tmp_path / "Config.dll"), str(tmp_path / "logs"))
    assert p.monitoring
    p.load()
    assert not os.path.exists(p.config_path)                # 시작만으로는 설정 파일을 쓰지 않음
    assert p.errors.path == os.path.join(str(tmp_path / "logs"), "error.log")


# ---------------------------------------------------------------- 10 설정 가져오기
SAMPLE_INI = ("[LINE_CONFIG]\r\nH1R1_SOURCE_IP=172.27.0.84\r\nH1R1_SOURCE_PORT=502\r\nH1R1_TARGEIP=172.27.0.85\r\n"
              "H1R1_TARGEPORT=502\r\nH1R1_SOURCE_MANUAL=0\r\nH1R1_MANUAL_RPM=0\r\nH1R1_SOURCE_LINE=H1R1\r\n"
              "H1R1_SOURCE_ID=155\r\n[MOTOR_CONFIG]\r\nH1R1_MOTORS=155=0=1=0=0,3=1=1=1=900\r\n")


@pytest.mark.parametrize("kind", ["dll", "ini", "uconfig", "cp949"])
def test_read_any_config_formats(tmp_path, kind):
    path = tmp_path / {"dll": "Config.dll", "ini": "Config.ini", "uconfig": "uconfig.dll", "cp949": "c.ini"}[kind]
    if kind == "dll":
        path.write_bytes(encrypt_config(SAMPLE_INI))
    elif kind == "cp949":
        path.write_bytes(("; 설정\r\n" + SAMPLE_INI).encode("cp949"))
    else:
        path.write_bytes(SAMPLE_INI.encode("utf-8"))
    configs, motors = cfgmod.config_from_text(cfgmod.read_any_config_text(str(path)))
    assert configs["H1R1"].source_ip == "172.27.0.84" and configs["H1R1"].source_id == 155
    assert [(m.id, m.mtype) for m in motors["H1R1"]] == [(155, MotorType.OLD_MOTOR), (3, MotorType.NEW_MOTOR)]


def test_read_any_config_rejects_other_files(tmp_path):
    bad = tmp_path / "x.dll"
    bad.write_bytes(os.urandom(64))
    with pytest.raises(cfgmod.ConfigError):
        cfgmod.read_any_config_text(str(bad))


def test_apply_config_import_backs_up_and_keeps_ctrl(tmp_path):
    p = _plant(tmp_path)
    p.houses["H1R2"].add_motor(9, MotorType.NEW_MOTOR, True)
    p.houses["H1R2"].add_motor(180, MotorType.OLD_CTRL, True, ip="172.27.0.95", set_rpm=600)
    p.save()
    src = tmp_path / "old" / "Config.dll"
    src.parent.mkdir()
    src.write_bytes(encrypt_config(SAMPLE_INI))
    imp = p.read_config_import(str(src))
    assert not imp.has_ctrl and imp.summary_lines()[0].startswith("1구역: 기존 172.27.0.84 / 신규 172.27.0.85 /")
    backup = p.apply_config_import(imp, keep_ctrl=True)
    assert backup and os.path.exists(backup)
    assert p.configs["H1R1"].source_ip == "172.27.0.84"
    assert [m.key for m in p.houses["H1R1"].get_list()] == ["1_3", "0_155"]
    assert p.houses["H1R2"].get_motor("1_9") is None                         # 가져온 내용으로 바뀜
    assert p.houses["H1R2"].get_motor("2_180@172.27.0.95:502") is not None   # 기존모터제어는 유지
    q = _plant(tmp_path)                                                     # 저장됨
    assert q.configs["H1R1"].source_id == 155
    # 백업 파일에는 이전 설정
    old_cfg, old_motors = cfgmod.load_config(backup)
    assert any(m.id == 9 for m in old_motors["H1R2"])


def test_apply_config_import_refused_while_running(tmp_path, fast):
    p = _plant(tmp_path)
    src = tmp_path / "Config.ini"
    src.write_text(SAMPLE_INI)
    imp = p.read_config_import(str(src))
    p.houses["H1R1"]._stop.clear()                    # 실행 중으로 표시
    try:
        with pytest.raises(RuntimeError):
            p.apply_config_import(imp)
    finally:
        p.houses["H1R1"]._stop.set()


def test_import_mini_items(tmp_path):
    path = tmp_path / "mini_control.json"
    path.write_text(json.dumps([
        {"ip": "172.27.0.116", "uid": 3, "mode": "pct", "pct": 50, "ref_ip": "", "ref_uid": 0},
        {"ip": "172.27.0.116", "uid": 4, "mode": "ref", "pct": 0, "ref_ip": "172.27.0.95", "ref_uid": 180},
    ]), encoding="utf-8")
    items = cfgmod.read_mini_items(str(path))
    assert items[0].rpm == 550
    p = _plant(tmp_path)
    p.configs["H1R3"] = p.configs["H1R3"].copy(source_ip="172.27.0.116")
    assert p.zone_for_ip("172.27.0.116") == "H1R3"
    added, skipped = p.import_mini_items(items, "H1R3")
    assert len(added) == 2 and skipped == []
    h = p.houses["H1R3"]
    m3 = h.get_motor("2_3@172.27.0.116:502")
    m4 = h.get_motor("2_4@172.27.0.116:502")
    assert m3.ctrl_mode == CTRL_MODE_MANUAL and m3.set_rpm == 550
    assert m4.ctrl_mode == CTRL_MODE_REF and m4.port == 502                  # 참조팬 -> 구역 참조
    added, skipped = p.import_mini_items(items, "H1R3")                    # 같은 항목은 건너뜀
    assert added == [] and len(skipped) == 2
    bad = tmp_path / "bad.json"
    bad.write_text("{}")
    with pytest.raises(cfgmod.ConfigError):
        cfgmod.read_mini_items(str(bad))


# ---------------------------------------------------------------- 화면
@pytest.fixture
def ui(tmp_path, monkeypatch):
    tk = pytest.importorskip("tkinter")
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("디스플레이 없음")
    root.withdraw()
    from motor_monitor.ui import dialogs, main_window, msgbox, tools
    msgs = []
    monkeypatch.setattr(dialogs.BaseDialog, "error", lambda self, m, t: msgs.append((t, m)))
    monkeypatch.setattr(dialogs.BaseDialog, "ask", lambda self, m, t: True)
    monkeypatch.setattr(msgbox, "ask_yes_no", lambda *a, **k: True)
    monkeypatch.setattr(msgbox, "show_info", lambda *a, **k: msgs.append(("info", a[1])))
    monkeypatch.setattr(msgbox, "show_warning", lambda *a, **k: msgs.append(("warn", a[1])))
    monkeypatch.setattr(msgbox, "show_error", lambda *a, **k: msgs.append(("err", a[1])))
    plant = Plant(str(tmp_path / "Config.dll"), monitoring=True)
    plant.load()
    win = main_window.MainWindow(root, plant)
    yield win, plant, dialogs, tools, msgs
    root.destroy()


def test_ui_starts_monitoring_banner(ui):
    win, plant, *_ = ui
    assert plant.monitoring and "항상 이 모드로 시작" in win.mode_label.cget("text")
    assert win.zone_panels["H1R1"].StartBtn.text == "1구역 모니터링 시작"


def test_ui_icons_colors(ui):
    """기존모터 청색, 신규모터 주황색, 기존모터제어 녹색 (그림 가운데 근처 팬 날개 색)"""
    win, *_ = ui
    import colorsys

    def hue(name):
        img = win.images.get(name, 40)
        hs = []
        for x in range(4, 36, 2):
            for y in range(4, 36, 2):
                try:
                    if img.transparency_get(x, y):
                        continue
                except Exception:
                    pass
                r, g, b = img.get(x, y)
                h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
                if s > 0.4 and 0.2 < l < 0.8:
                    hs.append(h * 360)
        hs.sort()
        return hs[len(hs) // 2]

    assert 190 <= hue("oldmotor") <= 230          # 청색
    assert 15 <= hue("newmotor") <= 40            # 주황색
    assert 110 <= hue("oldmotor_ctrl") <= 160     # 녹색


def test_ui_ctrl_ref_mode_dialog(ui):
    win, plant, dialogs, tools, msgs = ui
    plant.configs["H1R1"] = plant.configs["H1R1"].copy(source_ip="172.27.0.84")
    plant.configs["H1R2"] = plant.configs["H1R2"].copy(source_ip="172.27.0.95", source_manual=False,
                                                       source_line="H1R1", source_id=155)
    src = plant.houses["H1R1"]
    src.add_motor(155, MotorType.OLD_MOTOR, True)
    zp = win.zone_panels["H1R2"]
    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R2", "ADD", on_changed=zp.refresh)
    d.motor_type_var.set(2); d._on_type()
    # 참조 IP/PORT/ID 입력칸 없이, 이 구역 신규모터가 따르는 값(1구역 기존모터 155)을 보여줌
    for name in ("txtCPORT", "txtRIP", "txtRPORT", "txtRID", "BtnRefZone"):
        assert not hasattr(d, name)
    assert d.lbl_zone_ref.cget("text").startswith("1구역 기존모터 155 → 지금 값 없음")
    assert d.ctrl_mode_var.get() == CTRL_MODE_MANUAL and d.rb_ctrl_ref.cget("text") == "구역 참조"
    d.destroy()
    src._mark_received(src._motors["0_155"], 770)
    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R2", "ADD", on_changed=zp.refresh)
    d.motor_type_var.set(2); d._on_type()
    assert d.lbl_zone_ref.cget("text") == "1구역 기존모터 155 → 지금 770 RPM = 70%"
    d.ctrl_mode_var.set(CTRL_MODE_REF); d._on_ctrl_mode()
    assert str(d.txtSETRPM.cget("state")) == "disabled"
    # 참조 기존모터 자신(1구역 통신장치 172.27.0.84 의 ID 155)은 구역 참조로 제어할 수 없음
    d.txtCIP.delete(0, "end"); d.txtCIP.insert(0, "172.27.0.84")
    d.txtID.insert(0, "155")
    d.on_save()
    assert msgs[-1][0] == "오류" and "자신" in msgs[-1][1]
    d.txtCIP.delete(0, "end"); d.txtCIP.insert(0, "172.27.0.95")
    d.txtID.delete(0, "end"); d.txtID.insert(0, "176")
    d.on_save()
    key = "2_176@172.27.0.95:502"
    m = plant.houses["H1R2"].get_motor(key)
    assert m.ctrl_mode == CTRL_MODE_REF and m.port == 502
    from motor_monitor.ui.widgets import MotorListView
    assert "운전속도: 구역 참조" in MotorListView.type_icon(m)[1]
    # 수정: 수동 지정으로 바꿈
    d = dialogs.MotorDialog(win.root, plant, win.images, "H1R2", "MOD", mod_dkey=key, on_changed=zp.refresh)
    assert d.ctrl_mode_var.get() == CTRL_MODE_REF and d.lbl_zone_ref.cget("text").endswith("= 70%")
    d.ctrl_mode_var.set(CTRL_MODE_MANUAL); d._on_ctrl_mode()
    d.txtSETRPM.delete(0, "end"); d.txtSETRPM.insert(0, "650")
    d.on_save()
    m = plant.houses["H1R2"].get_motor(key)
    assert m.ctrl_mode == CTRL_MODE_MANUAL and m.set_rpm == 650
    assert MotorListView.type_icon(m)[1].startswith("기존모터제어 172.27.0.95\n운전속도: 수동 지정 650 RPM (59.1%)")


def test_ui_error_log_button_and_window(ui):
    win, plant, *_ = ui
    win.update_error_button(force=True)
    assert win.BtnErrors.cget("text") == "에러 로그"
    plant.errors.error("2구역", "신규모터 통신장치", "192.168.0.4 접속 실패")
    plant.errors.recover("2구역", "신규모터 통신장치", "다시 접속됨")
    win.update_error_button()
    assert win.BtnErrors.cget("text") == "에러 로그 (1)"
    win.open_error_log()
    w = win.error_window
    rows = w.rows
    assert rows[0][1] == "복구" and rows[1][1] == "에러" and rows[1][4] == "192.168.0.4 접속 실패"
    assert win.BtnErrors.cget("text") == "에러 로그"          # 창을 열면 확인한 것으로
    w.show_recover.set(False); w.refresh_force()
    assert [r[1] for r in w.rows] == ["에러"]
    plant.errors.error("3구역", "기존모터 ID 5", "명령 후 60초 동안 응답 없음")
    w._tick()
    assert w.rows[0][2] == "3구역"
    w.destroy()


def test_ui_import_config_dialog(ui, tmp_path):
    win, plant, dialogs, tools, msgs = ui
    src = tmp_path / "old_Config.dll"
    src.write_bytes(encrypt_config(SAMPLE_INI))
    imp = plant.read_config_import(str(src))
    d = tools.ConfigImportDialog(win.root, plant, imp, on_done=win.after_import)
    assert "1구역: 기존 172.27.0.84 / 신규 172.27.0.85 /" in d.txt.get("1.0", "end")
    assert ":502" not in d.txt.get("1.0", "end")                               # PORT 표시 없음
    d.on_import()
    assert plant.configs["H1R1"].source_ip == "172.27.0.84"
    assert win.zone_panels["H1R1"].lbl_old_ip.cget("text") == "172.27.0.84"
    assert [r.key for r in win.zone_panels["H1R1"].LV.rows] == ["1_3", "0_155"]
    assert msgs[-1][0] == "info" and "가져왔습니다" in msgs[-1][1]


def test_ui_import_mini_dialog(ui, tmp_path):
    win, plant, dialogs, tools, msgs = ui
    plant.configs["H1R4"] = plant.configs["H1R4"].copy(source_ip="172.27.0.116")
    items = [cfgmod.MiniItem("172.27.0.116", 3, "pct", 50, "", 0),
             cfgmod.MiniItem("172.27.0.116", 4, "ref", 0, "172.27.0.95", 180)]
    d = tools.MiniImportDialog(win.root, plant, items, "mini_control.json", on_done=win.after_import)
    assert d.zone_code == "H1R4"                                 # IP 가 같은 구역을 먼저 선택
    specs = [d.tree.item(i, "values")[2] for i in d.tree.get_children()]
    assert specs == ["수동 지정 550 RPM (50%)", "구역 참조 (미니 제어기 참조팬 172.27.0.95 ID 180 대신)"]
    d.on_import()
    keys = sorted(r.key for r in win.zone_panels["H1R4"].LV.rows)
    assert keys == ["2_3@172.27.0.116:502", "2_4@172.27.0.116:502"]
    assert "2개를 추가" in msgs[-1][1]


def test_ui_import_config_refused_while_running(ui, tmp_path):
    win, plant, dialogs, tools, msgs = ui
    plant.houses["H1R1"]._stop.clear()
    try:
        win.import_config_file(str(tmp_path / "none.dll"))
        assert msgs[-1][0] == "warn" and "종료" in msgs[-1][1]
    finally:
        plant.houses["H1R1"]._stop.set()
