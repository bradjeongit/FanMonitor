import os
import time

import pytest

from motor_monitor import config as cfgmod
from motor_monitor import constants as C
from motor_monitor import modbus
from motor_monitor.constants import MotorStatus, MotorType
from motor_monitor.helpers import check_integer, check_valid_ip, get_house_name, get_house_short_name
from motor_monitor.plant import Plant
from motor_monitor.simulator import FakeGateway


# ---------------------------------------------------------------- helpers
def test_house_names():
    assert get_house_name("H1R2E") == "온실1 / 2구역 / 짝"
    assert get_house_short_name("H1R3O", " > ") == "3구역 > 홀"
    assert get_house_name("H1") == ""


@pytest.mark.parametrize("ip,ok", [
    ("192.168.0.1", True), ("1.1.1.1", True), ("255.255.255.255", True),
    ("256.1.1.1", False), ("1.1.1", False), ("a.b.c.d", False), ("", False), ("1.1.1.1.1", False),
])
def test_check_valid_ip(ip, ok):
    assert check_valid_ip(ip) is ok


def test_check_integer():
    assert check_integer("123") and check_integer("") and not check_integer("12a") and not check_integer("-1")


# ---------------------------------------------------------------- modbus (원본과 바이트 단위 동일)
def test_frames_match_original_layout():
    assert modbus.build_old_rpm_read(1, 5) == bytes([0, 1, 0, 0, 0, 6, 5, 4, 0, 14, 0, 1])
    assert modbus.build_new_read(0x0102, 10) == bytes([1, 2, 0, 0, 0, 6, 10, 3, 0, 0x13, 0, 7])
    assert modbus.build_new_rpm_write(3, 10, 1000) == bytes([0, 3, 0, 0, 0, 6, 10, 6, 0, 0x16, 0x03, 0xE8])
    assert modbus.build_new_manual_write(4, 10) == bytes([0, 4, 0, 0, 0, 6, 10, 6, 0, 0x13, 0, 1])
    assert modbus.build_new_run_write(5, 10) == bytes([0, 5, 0, 0, 0, 6, 10, 6, 0, 0x14, 0, 1])


def test_parse_responses():
    old = bytes([0, 1, 0, 0, 0, 5, 5, 4, 2, 0x03, 0xB6])
    r = modbus.parse_old_response(old)
    assert r.unit_id == 5 and r.rpm == 950
    new = bytes([0, 1, 0, 0, 0, 17, 10, 3, 14, 0, 1, 0, 1, 0, 0, 0x03, 0xE8] + [0] * 6)
    n = modbus.parse_new_response(new)
    assert (n.unit_id, n.automanual, n.startstop, n.rpm) == (10, 1, 1, 1000)
    assert modbus.parse_old_response(new) is None


def test_frame_splitter_handles_coalesced_and_split():
    a = modbus.build_old_rpm_read(1, 5)
    b = modbus.build_new_read(2, 6)
    sp = modbus.FrameSplitter()
    assert sp.feed(a + b[:4]) == [a]
    assert sp.feed(b[4:]) == [b]


# ---------------------------------------------------------------- config
def test_config_defaults_and_roundtrip(tmp_path):
    path = str(tmp_path / "Config.ini")
    plant = Plant(path)
    plant.load()
    assert os.path.exists(path)
    c = plant.configs["H1R1E"]
    assert (c.source_ip, c.target_ip, c.source_port, c.source_manual) == ("192.168.0.1", "192.168.0.2", 502, True)
    assert plant.configs["H1R1O"].source_ip == "192.168.0.3"

    plant.houses["H1R1E"].add_motor(5, MotorType.OLD_MOTOR, True)
    plant.houses["H1R1E"].add_motor(10, MotorType.NEW_MOTOR, False)
    plant.set_line_config("H1R1E", c.copy(source_manual=False, source_line="H1R1E", source_id=5))

    text = open(path, encoding="utf-8").read()
    assert "H1R1E_TARGEIP=192.168.0.2" in text      # 원본 키 이름 유지
    assert "H1R1E_MOTORS=10=1=0,5=0=1" in text      # 표시 순서(신규 먼저)로 저장

    p2 = Plant(path)
    p2.load()
    assert p2.configs["H1R1E"].source_manual is False
    assert [(m.id, m.mtype, m.use) for m in p2.houses["H1R1E"].get_list()] == [
        (10, MotorType.NEW_MOTOR, False), (5, MotorType.OLD_MOTOR, True)]


def test_reads_original_windows_ini(tmp_path):
    path = tmp_path / "Config.ini"
    path.write_bytes("[LINE_CONFIG]\r\nH1R1E_SOURCE_IP=10.0.0.5\r\nH1R1E_SOURCE_MANUAL=0\r\n"
                     "H1R1E_SOURCE_LINE=H1R2O\r\nH1R1E_SOURCE_ID=7\r\n"
                     "[MOTOR_CONFIG]\r\nH1R1E_MOTORS=1=0=1,bad,2=1=1\r\n".encode("cp949"))
    configs, motors = cfgmod.load_config(str(path))
    assert configs["H1R1E"].source_ip == "10.0.0.5"
    assert configs["H1R1E"].source_line == "H1R2O" and configs["H1R1E"].source_id == 7
    assert motors["H1R1E"] == [(1, MotorType.OLD_MOTOR, True), (2, MotorType.NEW_MOTOR, True)]


# ---------------------------------------------------------------- 정렬/상태
def _house(tmp_path):
    plant = Plant(str(tmp_path / "Config.ini"))
    plant.load()
    return plant, plant.houses["H1R1E"]


def test_sorting_matches_original(tmp_path):
    _, h = _house(tmp_path)
    for mid, t in [(3, 0), (1, 1), (2, 0), (9, 1)]:
        h.add_motor(mid, MotorType(t), True)
    ids = lambda: [(int(m.mtype), m.id) for m in h.get_list()]
    assert ids() == [(1, 1), (1, 9), (0, 2), (0, 3)]   # 기본: 신규 먼저, 같은 종류는 ID 오름차순
    h.sort("MTYPE")
    assert ids() == [(0, 2), (0, 3), (1, 1), (1, 9)]
    h.sort("ID")
    assert [i for _, i in ids()] == [1, 2, 3, 9]
    h.sort("ID")
    assert [i for _, i in ids()] == [9, 3, 2, 1]


def test_status_transitions(tmp_path):
    plant, h = _house(tmp_path)
    now = [1000.0]
    h._clock = lambda: now[0]
    h.add_motor(5, MotorType.OLD_MOTOR, True)
    key = "0_5"
    h.update_status()
    assert h.get_motor(key).status == MotorStatus.READY
    h._handle_src_frame(bytes([0, 0, 0, 0, 0, 5, 5, 4, 2, 0x03, 0xB6]))
    h.update_status()
    assert h.get_motor(key).status == MotorStatus.OK and h.get_motor(key).rpm == 950
    now[0] += 60
    h.update_status()
    assert h.get_motor(key).status == MotorStatus.ERR
    h.mod_motor(key, False)
    m = h.get_motor(key)
    assert m.status == MotorStatus.NOTUSE and m.must_send_zero and m.rdate is None
    h.mod_motor(key, True)
    assert h.get_motor(key).status == MotorStatus.READY


def test_reference_users(tmp_path):
    plant, h = _house(tmp_path)
    plant.configs["H1R2E"] = plant.configs["H1R2E"].copy(source_manual=False, source_line="H1R1E", source_id=5)
    assert plant.find_reference_users("H1R1E", 5) == ["H1R2E"]
    assert plant.find_reference_users("H1R1E", 6) == []


# ---------------------------------------------------------------- 통신 통합 시험 (시뮬레이터)
@pytest.fixture
def fast(monkeypatch):
    for name, val in [("SEND_INTERVAL", 0.01), ("CYCLE_INTERVAL", 0.05), ("RECONNECT_DELAY", 0.2),
                      ("ERROR_RETRY_DELAY", 0.05)]:
        monkeypatch.setattr(C, name, val)


def _wait(cond, timeout=5.0):
    end = time.time() + timeout
    while time.time() < end:
        if cond():
            return True
        time.sleep(0.02)
    return False


def test_control_loop_follows_reference_motor(tmp_path, fast):
    gw_old = FakeGateway().start()
    gw_new = FakeGateway().start()
    gw_old.add_old(5, 950)
    gw_new.add_new(10, automanual=0, startstop=0, rpm=0)
    try:
        plant = Plant(str(tmp_path / "Config.ini"))
        plant.load()
        code = "H1R1E"
        h = plant.houses[code]
        h.add_motor(5, MotorType.OLD_MOTOR, True)
        h.add_motor(10, MotorType.NEW_MOTOR, True)
        plant.configs[code] = plant.configs[code].copy(
            source_ip=gw_old.host, source_port=gw_old.port,
            target_ip=gw_new.host, target_port=gw_new.port,
            source_manual=False, source_line=code, source_id=5)

        h.service_start()
        assert _wait(lambda: h.get_motor("0_5").rpm == 950)
        assert h.source_net_status and h.target_net_status
        h.update_status()
        assert _wait(lambda: gw_new.new_regs[10][modbus.NEW_REG_RPM] == 950)
        assert gw_new.new_regs[10][modbus.NEW_REG_AUTOMANUAL] == 1
        assert gw_new.new_regs[10][modbus.NEW_REG_STARTSTOP] == 1

        # 수동 모드로 바꾸면 수동 RPM 으로 제어
        plant.configs[code] = plant.configs[code].copy(source_manual=True, manual_rpm=700)
        assert _wait(lambda: gw_new.new_regs[10][modbus.NEW_REG_RPM] == 700)

        # 사용 중지 -> RPM 0 한 번 전송 (사용중 신규모터가 없어도 전송되어야 함)
        h.mod_motor("1_10", False)
        assert _wait(lambda: gw_new.new_regs[10][modbus.NEW_REG_RPM] == 0)
        h.service_stop(join=True)
        assert not h.running
    finally:
        gw_old.stop()
        gw_new.stop()


def test_reference_not_ok_holds_control(tmp_path, fast):
    gw_new = FakeGateway().start()
    gw_new.add_new(10, automanual=1, startstop=1, rpm=500)
    try:
        plant = Plant(str(tmp_path / "Config.ini"))
        plant.load()
        code = "H1R1E"
        h = plant.houses[code]
        h.add_motor(5, MotorType.OLD_MOTOR, True)       # 수신 없음 -> 대기 상태
        h.add_motor(10, MotorType.NEW_MOTOR, True)
        plant.configs[code] = plant.configs[code].copy(
            source_ip="127.0.0.1", source_port=1,
            target_ip=gw_new.host, target_port=gw_new.port,
            source_manual=False, source_line=code, source_id=5)
        h.service_start()
        assert _wait(lambda: len(gw_new.requests) >= 3)
        assert gw_new.writes == []
        assert not h.source_net_status
        h.service_stop(join=True)
    finally:
        gw_new.stop()
