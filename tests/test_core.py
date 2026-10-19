import time

import pytest

from motor_monitor import config as cfgmod
from motor_monitor import constants as C
from motor_monitor import modbus
from motor_monitor.config import MotorEntry
from motor_monitor.constants import MotorStatus, MotorType
from motor_monitor.crypto3des import TripleDES, decrypt_config, encrypt_config  # noqa: F401
from motor_monitor.helpers import check_integer, check_valid_ip, get_house_name, zone_number
from motor_monitor.plant import Plant
from motor_monitor.simulator import FakeGateway


# ---------------------------------------------------------------- helpers
def test_house_names():
    assert get_house_name("H1R3") == "3구역" and zone_number("H1R8") == 8
    assert get_house_name("H1R1E") == "" and get_house_name("") == ""


@pytest.mark.parametrize("ip,ok", [
    ("192.168.0.1", True), ("172.27.0.82", True), ("255.255.255.255", True),
    ("256.1.1.1", False), ("1.1.1", False), ("a.b.c.d", False), ("", False), ("1.1.1.1.1", False),
])
def test_check_valid_ip(ip, ok):
    assert check_valid_ip(ip) is ok


def test_check_integer():
    assert check_integer("123") and check_integer("") and not check_integer("12a") and not check_integer("-1")


# ---------------------------------------------------------------- 암호화 (운영 Config.dll 호환)
def test_des_standard_vector():
    from motor_monitor.crypto3des import _des_block, _subkeys
    assert _des_block(0x0123456789ABCDEF, _subkeys(bytes.fromhex("133457799BBCDFF1"))) == 0x85E813540F0AB405


def test_config_cipher_matches_reference():
    # 독립 구현(cryptography 라이브러리)으로 만든 기준 암호문 - 운영 프로그램과 같은 키/IV
    ref = bytes.fromhex(
        "b3fdb6465fe506ced9a724d8143c11ecc1590e3aa5b648b6cf9b607cdc805e477936555541ef4d5b44f714ba"
        "74833584a3f63418ce0e2df3512e5102127f6c8930f17b5725e575fa26c7e113981799dceeb664b6082d00e4")
    text = "[LINE_CONFIG]\r\nH1R1_SOURCE_IP=10.0.0.1\r\n"
    assert encrypt_config(text) == ref
    assert decrypt_config(ref) == text
    assert decrypt_config(b"") == ""


def test_tripledes_roundtrip_and_bad_key():
    c = TripleDES(bytes(range(24)), bytes(8))
    for n in (0, 1, 7, 8, 9, 100):
        data = bytes(range(n))
        assert c.decrypt(c.encrypt(data)) == data
    with pytest.raises(ValueError):
        TripleDES(bytes(range(1, 25)), bytes(8)).decrypt(c.encrypt(b"hello world"))


# ---------------------------------------------------------------- modbus (운영 프로그램과 바이트 단위 동일)
def test_frames_match_production_layout():
    assert modbus.build_old_rpm_read(1, 5) == bytes([0, 1, 0, 0, 0, 6, 5, 4, 0, 14, 0, 1])
    assert modbus.build_new_read_page1(2, 10) == bytes([0, 2, 0, 0, 0, 6, 10, 3, 0, 0, 0, 9])
    assert modbus.build_new_read(0x0102, 10) == bytes([1, 2, 0, 0, 0, 6, 10, 3, 0, 0x13, 0, 7])
    assert modbus.build_new_rpm_write(3, 10, 1000) == bytes([0, 3, 0, 0, 0, 6, 10, 6, 0, 0x16, 0x03, 0xE8])
    assert modbus.build_new_manual_write(4, 10) == bytes([0, 4, 0, 0, 0, 6, 10, 6, 0, 0x13, 0, 1])
    assert modbus.build_new_run_write(5, 10) == bytes([0, 5, 0, 0, 0, 6, 10, 6, 0, 0x14, 0, 1])
    assert modbus.build_old_rpm_read(0, 250)[6] == 250


def test_parse_responses():
    old = bytes([0, 1, 0, 0, 0, 5, 5, 4, 2, 0x03, 0xB6])
    assert modbus.parse_old_response(old).rpm == 950
    p1 = bytes([0, 1, 0, 0, 0, 21, 10, 3, 18] + [0] * 10 + [0x02, 0x8A] + [0] * 6)
    assert len(p1) == 27 and modbus.parse_new_page1(p1).actual_rpm == 650
    p2 = bytes([0, 1, 0, 0, 0, 17, 10, 3, 14, 0, 1, 0, 1, 0, 0, 0x03, 0xE8] + [0] * 6)
    n = modbus.parse_new_response(p2)
    assert (n.unit_id, n.automanual, n.startstop, n.rpm) == (10, 1, 1, 1000)
    assert modbus.parse_new_page1(p2) is None and modbus.parse_new_response(p1) is None


def test_frame_splitter_handles_coalesced_and_split():
    a = modbus.build_old_rpm_read(1, 5)
    b = modbus.build_new_read(2, 6)
    sp = modbus.FrameSplitter()
    assert sp.feed(a + b[:4]) == [a]
    assert sp.feed(b[4:]) == [b]


# ---------------------------------------------------------------- config
def test_motor_list_formats():
    assert cfgmod.parse_motor_list("2=0=1=0=0,13=1=1=1=500,5=1=0") == [
        MotorEntry(2, MotorType.OLD_MOTOR, True, False, 0),
        MotorEntry(13, MotorType.NEW_MOTOR, True, True, 500),
        MotorEntry(5, MotorType.NEW_MOTOR, False, False, 0)]
    assert cfgmod.format_motor_list([MotorEntry(13, MotorType.NEW_MOTOR, True, True, 500)]) == "13=1=1=1=500"
    assert cfgmod.parse_motor_list("bad,1=9=1") == []


def test_config_defaults_and_roundtrip(tmp_path):
    path = str(tmp_path / "Config.dll")
    plant = Plant(path)
    plant.load()
    raw = open(path, "rb").read()
    text = decrypt_config(raw)
    assert text.startswith("[LINE_CONFIG]\r\nH1R1_SOURCE_IP=192.168.0.1\r\n")
    assert "H1R1_TARGEIP=192.168.0.2" in text and "H1R2_SOURCE_IP=192.168.0.3" in text
    assert "[MOTOR_CONFIG]\r\nH1R1_MOTORS=\r\n" in text and "H1R9" not in text

    h = plant.houses["H1R3"]
    h.add_motor(2, MotorType.OLD_MOTOR, True)
    h.add_motor(13, MotorType.NEW_MOTOR, True, True, 500)
    plant.set_line_config("H1R3", plant.configs["H1R3"].copy(source_manual=False, source_line="H1R1", source_id=2))
    text = decrypt_config(open(path, "rb").read())
    assert "H1R3_MOTORS=2=0=1=0=0,13=1=1=1=500" in text      # 등록 순서로 저장

    p2 = Plant(path)
    p2.load()
    assert p2.configs["H1R3"].source_line == "H1R1"
    m = p2.houses["H1R3"].get_motor("1_13")
    assert m.use_limit and m.limit_rpm == 500


def test_load_then_save_is_byte_identical(tmp_path):
    text = ("[LINE_CONFIG]\r\n" + "".join(
        f"H1R{i}_SOURCE_IP=10.0.0.{i}\r\nH1R{i}_SOURCE_PORT=502\r\nH1R{i}_TARGEIP=10.0.1.{i}\r\n"
        f"H1R{i}_TARGEPORT=502\r\nH1R{i}_SOURCE_MANUAL=0\r\nH1R{i}_MANUAL_RPM=0\r\n"
        f"H1R{i}_SOURCE_LINE=H1R{i}\r\nH1R{i}_SOURCE_ID={i}\r\n" for i in range(1, 9))
        + "[MOTOR_CONFIG]\r\n" + "".join(f"H1R{i}_MOTORS={i}=0=1=0=0,3=1=1=1=450,1=1=1=0=0\r\n" for i in range(1, 9)))
    path = tmp_path / "Config.dll"
    path.write_bytes(encrypt_config(text))
    original = path.read_bytes()
    Plant(str(path)).load()
    assert path.read_bytes() == original


def test_corrupt_config_is_not_overwritten(tmp_path):
    path = tmp_path / "Config.dll"
    path.write_bytes(b"not an encrypted file!!")
    with pytest.raises(cfgmod.ConfigError):
        Plant(str(path)).load()
    assert path.read_bytes() == b"not an encrypted file!!"


# ---------------------------------------------------------------- 정렬/상태
def _house(tmp_path, code="H1R1"):
    plant = Plant(str(tmp_path / "Config.dll"))
    plant.load()
    return plant, plant.houses[code]


def test_sorting_matches_production(tmp_path):
    _, h = _house(tmp_path)
    for mid, t in [(126, 0), (1, 1), (2, 0), (9, 1)]:
        h.add_motor(mid, MotorType(t), True)
    ids = lambda: [(int(m.mtype), m.id) for m in h.get_list()]
    assert ids() == [(1, 1), (1, 9), (0, 2), (0, 126)]   # 기본: 신규 먼저, 같은 종류는 ID 오름차순
    h.sort("MTYPE")
    assert ids() == [(0, 2), (0, 126), (1, 1), (1, 9)]
    h.sort("ID")
    assert [i for _, i in ids()] == [1, 2, 9, 126]
    h.sort("ID")
    assert [i for _, i in ids()] == [126, 9, 2, 1]


def test_status_had_err_and_timeout(tmp_path):
    plant, h = _house(tmp_path)
    now = [1000.0]
    h._clock = lambda: now[0]
    h.add_motor(5, MotorType.OLD_MOTOR, True)
    key = "0_5"
    h.update_status()
    assert h.get_motor(key).status == MotorStatus.READY
    h._handle_src_frame(bytes([0, 0, 0, 0, 0, 5, 5, 4, 2, 0x03, 0xB6]))
    m = h.get_motor(key)
    assert m.status == MotorStatus.OK and m.rpm == 950 and not m.had_err
    # 명령 전송 후 60초 무응답 -> 이상기록(빨간 !), 상태는 정상 유지
    with h._lock:
        h._motors[key].check_since = now[0]
    now[0] += 60
    h.update_status()
    m = h.get_motor(key)
    assert m.had_err and m.status == MotorStatus.OK and m.check_since is None
    # 응답이 다시 와도 이상기록은 유지 (제어 재시작 시 해제)
    h._handle_src_frame(bytes([0, 0, 0, 0, 0, 5, 5, 4, 2, 0x03, 0xB6]))
    assert h.get_motor(key).had_err
    # 마지막 응답 후 1시간 -> 이상
    now[0] += 3600
    h.update_status()
    assert h.get_motor(key).status == MotorStatus.ERR
    h.mod_motor(key, False)
    m = h.get_motor(key)
    assert m.status == MotorStatus.NOTUSE and m.must_send_zero and m.rdate is None


def test_status_only_updates_while_running(tmp_path):
    plant, h = _house(tmp_path)
    h.add_motor(5, MotorType.OLD_MOTOR, False)
    with h._lock:
        h._motors["0_5"].status = MotorStatus.OK
    plant.update_all_status()                 # 제어 중이 아니면 갱신 안 함 (운영 동일)
    assert h.get_motor("0_5").status == MotorStatus.OK


def test_limit_and_reference(tmp_path):
    plant, h = _house(tmp_path)
    h.add_motor(2, MotorType.OLD_MOTOR, True)
    h.add_motor(13, MotorType.NEW_MOTOR, True, True, 500)
    plant.configs["H1R1"] = plant.configs["H1R1"].copy(source_manual=False, source_line="H1R1", source_id=2)
    h._handle_src_frame(bytes([0, 0, 0, 0, 0, 5, 2, 4, 2, 0x03, 0x02]))   # 기존모터 2 = 770
    assert h.resolve_setting_rpm() == 770
    p2 = bytes([0, 1, 0, 0, 0, 17, 13, 3, 14, 0, 1, 0, 1, 0, 0, 0x01, 0xF4] + [0] * 6)  # 이미 500
    h._handle_tgt_frame(p2)
    assert list(h._tgt_write_q) == []          # 제한 500 적용 -> 쓸 것 없음
    assert plant.find_reference_users("H1R1", 2) == ["H1R1"]
    plant.configs["H1R1"] = plant.configs["H1R1"].copy(source_id=99)    # 없는 모터 참조
    assert h.resolve_setting_rpm() is None and not plant.reference_ok("H1R1")


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


def test_control_loop(tmp_path, fast):
    gw_old = FakeGateway().start()
    gw_new = FakeGateway().start()
    gw_old.add_old(126, 770)
    gw_new.add_new(1)
    gw_new.add_new(4)
    try:
        plant = Plant(str(tmp_path / "Config.dll"))
        plant.load()
        code = "H1R5"
        h = plant.houses[code]
        h.add_motor(126, MotorType.OLD_MOTOR, True)
        h.add_motor(1, MotorType.NEW_MOTOR, True)
        h.add_motor(4, MotorType.NEW_MOTOR, True, True, 400)        # 속도제한 400
        plant.configs[code] = plant.configs[code].copy(
            source_ip=gw_old.host, source_port=gw_old.port, target_ip=gw_new.host, target_port=gw_new.port,
            source_manual=False, source_line=code, source_id=126)

        plant.start_line(code)
        assert _wait(lambda: h.get_motor("0_126").rpm == 770)
        assert h.source_net_status and h.target_net_status
        assert _wait(lambda: gw_new.new_regs[1][modbus.NEW_REG_RPM] == 770)
        assert _wait(lambda: gw_new.new_regs[4][modbus.NEW_REG_RPM] == 400)
        assert gw_new.new_regs[1][modbus.NEW_REG_AUTOMANUAL] == 1 and gw_new.new_regs[1][modbus.NEW_REG_STARTSTOP] == 1
        # 신규모터 표시 RPM = 레지스터 0x05 (실제 RPM)
        assert _wait(lambda: h.get_motor("1_1").rpm == 770 and h.get_motor("1_4").rpm == 400)
        assert h.get_motor("1_1").status == MotorStatus.OK

        plant.configs[code] = plant.configs[code].copy(source_manual=True, manual_rpm=300)
        assert _wait(lambda: gw_new.new_regs[1][modbus.NEW_REG_RPM] == 300)
        assert _wait(lambda: gw_new.new_regs[4][modbus.NEW_REG_RPM] == 300)   # 제한보다 낮으면 그대로

        h.mod_motor("1_1", False)
        h.mod_motor("1_4", False)
        assert _wait(lambda: gw_new.new_regs[1][modbus.NEW_REG_RPM] == 0 and gw_new.new_regs[4][modbus.NEW_REG_RPM] == 0)
        plant.stop_line(code)
        h.service_stop(join=True)
        assert not h.running
    finally:
        gw_old.stop()
        gw_new.stop()


def test_reference_not_ok_holds_control(tmp_path, fast):
    gw_new = FakeGateway().start()
    gw_new.add_new(10, automanual=1, startstop=1, rpm=500)
    try:
        plant = Plant(str(tmp_path / "Config.dll"))
        plant.load()
        code = "H1R1"
        h = plant.houses[code]
        h.add_motor(5, MotorType.OLD_MOTOR, True)       # 수신 없음 -> 대기
        h.add_motor(10, MotorType.NEW_MOTOR, True)
        plant.configs[code] = plant.configs[code].copy(
            source_ip="127.0.0.1", source_port=1, target_ip=gw_new.host, target_port=gw_new.port,
            source_manual=False, source_line=code, source_id=5)
        h.service_start()
        assert _wait(lambda: len(gw_new.requests) >= 4)
        assert gw_new.writes == []
        assert not h.source_net_status
        assert _wait(lambda: h.get_motor("1_10").status == MotorStatus.OK)   # 신규모터 응답은 정상 표시
        h.service_stop(join=True)
    finally:
        gw_new.stop()


# ---------------------------------------------------------------- 모니터링 모드 (읽기 전용)
def test_monitoring_mode_never_writes(tmp_path, fast):
    gw_old = FakeGateway().start()
    gw_new = FakeGateway().start()
    gw_old.add_old(2, 770)
    gw_new.add_new(1, automanual=0, startstop=0, rpm=100)        # 제어 모드라면 수동·동작·RPM 쓰기 3건이 나갈 상태
    gw_new.add_new(3, automanual=1, startstop=1, rpm=900)        # 속도제한 모터 (제어 모드라면 500 으로 씀)
    try:
        path = tmp_path / "Config.dll"
        Plant(str(path)).load()                                  # 설정 파일 준비 (제어 모드)
        before = path.read_bytes()
        plant = Plant(str(path), monitoring=True)
        plant.load()
        code = "H1R1"
        h = plant.houses[code]
        h.add_motor(2, MotorType.OLD_MOTOR, True)
        h.add_motor(1, MotorType.NEW_MOTOR, True)
        h.add_motor(3, MotorType.NEW_MOTOR, True, True, 500)
        plant.configs[code] = plant.configs[code].copy(
            source_ip=gw_old.host, source_port=gw_old.port, target_ip=gw_new.host, target_port=gw_new.port,
            source_manual=False, source_line=code, source_id=2)
        plant.start_line(code)
        # 읽기는 정상 동작
        assert _wait(lambda: h.get_motor("0_2").rpm == 770)
        assert _wait(lambda: h.get_motor("1_3").status == MotorStatus.OK and h.get_motor("1_1").status == MotorStatus.OK)
        assert _wait(lambda: sum(1 for r in gw_new.requests if r[7] == modbus.FC_READ_HOLDING) >= 8)
        h.mod_motor("1_3", False)                                 # 제어 모드라면 RPM 0 쓰기
        time.sleep(0.5)
        # 쓰기(FC06) 요청은 단 한 건도 없어야 함
        assert gw_new.writes == []
        assert all(r[7] != modbus.FC_WRITE_SINGLE for r in gw_new.requests + gw_old.requests)
        assert list(h._tgt_write_q) == []
        plant.shutdown()
        h.service_stop(join=True)
        assert path.read_bytes() == before                        # 시작/종료 시 설정 파일 자동 저장 안 함
    finally:
        gw_old.stop()
        gw_new.stop()


def test_monitoring_mode_switch_at_runtime(tmp_path, fast):
    """실행 중 모드 전환: 모니터링 -> 제어(쓰기 시작) -> 모니터링(쓰기 즉시 중단)"""
    gw_old = FakeGateway().start()
    gw_new = FakeGateway().start()
    gw_old.add_old(2, 770)
    gw_new.add_new(1, automanual=1, startstop=1, rpm=100)
    gw_new.add_new(5, automanual=1, startstop=1, rpm=600)
    try:
        plant = Plant(str(tmp_path / "Config.dll"), monitoring=True)
        plant.load()
        code = "H1R1"
        h = plant.houses[code]
        h.add_motor(2, MotorType.OLD_MOTOR, True)
        h.add_motor(1, MotorType.NEW_MOTOR, True)
        h.add_motor(5, MotorType.NEW_MOTOR, True)
        plant.configs[code] = plant.configs[code].copy(
            source_ip=gw_old.host, source_port=gw_old.port, target_ip=gw_new.host, target_port=gw_new.port,
            source_manual=False, source_line=code, source_id=2)
        plant.start_line(code)
        assert _wait(lambda: h.get_motor("1_1").status == MotorStatus.OK)
        h.mod_motor("1_5", False)                                 # 모니터링 중 사용중지 -> RPM 0 은 보류
        time.sleep(0.3)
        assert gw_new.writes == []
        # 제어 모드로 전환: 다음 주기부터 RPM 쓰기, 보류했던 RPM 0 도 전송
        plant.set_monitoring(False)
        assert not plant.monitoring
        assert _wait(lambda: gw_new.new_regs[1][modbus.NEW_REG_RPM] == 770)
        assert _wait(lambda: gw_new.new_regs[5][modbus.NEW_REG_RPM] == 0)
        # 다시 모니터링 모드: 이후 쓰기 0건
        plant.set_monitoring(True)
        time.sleep(0.2)
        n = len(gw_new.writes)
        gw_new.new_regs[1][modbus.NEW_REG_RPM] = 100              # 제어 모드라면 770 으로 다시 씀
        time.sleep(0.5)
        assert len(gw_new.writes) == n and gw_new.new_regs[1][modbus.NEW_REG_RPM] == 100
        assert list(h._tgt_write_q) == []
        h.service_stop(join=True)
    finally:
        gw_old.stop()
        gw_new.stop()


# ---------------------------------------------------------------- 기존모터제어 (ECblue)
def test_ecblue_frames_and_mode_bits():
    assert modbus.build_read_holding(1, 2, modbus.ECB_H_CONTROL, 4) == bytes([0, 1, 0, 0, 0, 6, 2, 3, 0, 1, 0, 4])
    assert modbus.build_read_input(1, 2, 14, 1) == bytes([0, 1, 0, 0, 0, 6, 2, 4, 0, 14, 0, 1])
    assert modbus.ecb_mode_value(0x0000) == 0x0101 and modbus.ecb_mode_value(0x00F3) == 0x01F1
    assert modbus.ecb_mode_ok(0x0101) and not modbus.ecb_mode_ok(0x0001) and not modbus.ecb_mode_ok(0x0100)
    exc = modbus.parse_response(bytes([0, 7, 0, 0, 0, 3, 2, 0x86, 4]))
    assert exc.tid == 7 and exc.fc == 6 and exc.exception == 4


def test_ctrl_motor_config_roundtrip(tmp_path):
    path = tmp_path / "Config.dll"
    plant = Plant(str(path))
    plant.load()
    text = decrypt_config(path.read_bytes())
    assert "CTRL_MOTOR_CONFIG" not in text                        # 없으면 섹션도 없음 (운영 파일과 동일)
    h = plant.houses["H1R2"]
    h.add_motor(66, MotorType.OLD_MOTOR, True)
    h.add_motor(66, MotorType.OLD_CTRL, True, ip="172.27.0.84", port=502, set_rpm=700)
    h.add_motor(68, MotorType.OLD_CTRL, False, ip="172.27.0.84", port=502, set_rpm=650)
    plant.save()
    text = decrypt_config(path.read_bytes())
    assert "H1R2_MOTORS=66=0=1=0=0\r\n" in text                    # 운영 형식 목록에는 기존모터제어 제외
    assert "[CTRL_MOTOR_CONFIG]" in text and "H1R2_MOTORS=66@172.27.0.84:502=1=700,68@172.27.0.84:502=0=650" in text
    p2 = Plant(str(path))
    p2.load()
    m = p2.houses["H1R2"].get_motor("2_66@172.27.0.84:502")
    assert m.mtype == MotorType.OLD_CTRL and m.set_rpm == 700 and m.port == 502 and m.use
    assert p2.houses["H1R2"].get_motor("0_66") is not None       # 같은 ID 의 읽기용 기존모터와 별개
    assert p2.houses["H1R2"].get_old_motor_ids() == ["66"]


def test_ctrl_motor_sets_speed_and_keeps_it(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_ecblue(2, analog_rpm=0, h4=0, pin_level=0)              # 아날로그 신호 끊김(0 RPM), PIN 보호 0단계
    try:
        plant = Plant(str(tmp_path / "Config.dll"))
        plant.load()
        h = plant.houses["H1R1"]
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, set_rpm=700)
        key = f"2_2@{gw.host}:{gw.port}"
        plant.start_line("H1R1")
        assert _wait(lambda: gw.ecb_speed(2) == 700)
        regs = gw.ecb[2]["h"]
        assert regs[4] == 0x0101 and regs[1] & 1 and regs[2] == 700   # 제어모드 1 + Enable + 운전속도
        assert (2, 0, 10) in gw.writes                                 # PIN 0010 입력 후 h4 변경
        assert _wait(lambda: h.get_motor(key).rpm == 700 and h.get_motor(key).status == MotorStatus.OK)
        # 모터가 리셋되어 속도가 0 이 되어도 다시 지정 속도로 맞춤
        regs[2] = 0
        assert _wait(lambda: gw.ecb_speed(2) == 700)
        # 운전속도 변경
        h.mod_motor(key, True, set_rpm=550)
        assert _wait(lambda: gw.ecb_speed(2) == 550)
        assert gw.ecb[2]["h4_writes"] == 1                              # 영구저장 h4 는 한 번만 씀
        # 사용 중지 -> 운전속도 0 한 번
        h.mod_motor(key, False)
        assert _wait(lambda: gw.ecb[2]["h"][2] == 0)
        plant.stop_line("H1R1")
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_ctrl_motor_mode_write_refused_is_limited(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_ecblue(5, analog_rpm=300, h4=0)
    gw.ecb[5]["refuse_h4"] = True
    try:
        plant = Plant(str(tmp_path / "Config.dll"))
        plant.load()
        h = plant.houses["H1R1"]
        h.add_motor(5, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, set_rpm=700)
        key = f"2_5@{gw.host}:{gw.port}"
        plant.start_line("H1R1")
        assert _wait(lambda: h.get_motor(key).ctrl_note.startswith("제어모드 설정 실패 (PIN"), timeout=8)
        n_tries = sum(1 for r in gw.requests if r[7] == 6 and r[9] == modbus.ECB_H_CONTROLMODE)
        time.sleep(0.5)
        assert n_tries == C.CTRL_MODE_WRITE_MAX_TRIES                     # 영구저장 레지스터 반복 쓰기 방지
        assert sum(1 for r in gw.requests if r[7] == 6 and r[9] == modbus.ECB_H_CONTROLMODE) == n_tries
        assert h.get_motor(key).had_err and h.get_motor(key).rpm == 300   # 실제 속도는 계속 표시
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_ctrl_motor_survives_unexpected_error(tmp_path, fast, monkeypatch, caplog):
    """처리 중 예상 밖 오류가 나도 로그를 남기고 다음 주기에 계속 제어함."""
    from motor_monitor.house import HouseLine
    gw = FakeGateway().start()
    gw.add_ecblue(2, analog_rpm=0, h4=0x0101)
    orig = HouseLine._ctrl_cycle
    calls = {"n": 0}

    def flaky(self, *a, **kw):
        calls["n"] += 1
        if calls["n"] == 1:
            raise ValueError("test")
        return orig(self, *a, **kw)

    monkeypatch.setattr(HouseLine, "_ctrl_cycle", flaky)
    try:
        plant = Plant(str(tmp_path / "Config.dll"))
        plant.load()
        h = plant.houses["H1R1"]
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, set_rpm=600)
        plant.start_line("H1R1")
        assert _wait(lambda: gw.ecb_speed(2) == 600)
        assert calls["n"] >= 2 and "처리 오류" in caplog.text
        h.service_stop(join=True)
    finally:
        gw.stop()


def test_ctrl_motor_monitoring_never_writes(tmp_path, fast):
    gw = FakeGateway().start()
    gw.add_ecblue(2, analog_rpm=450, h4=0)
    try:
        plant = Plant(str(tmp_path / "Config.dll"), monitoring=True)
        plant.load()
        h = plant.houses["H1R1"]
        h.add_motor(2, MotorType.OLD_CTRL, True, ip=gw.host, port=gw.port, set_rpm=700)
        key = f"2_2@{gw.host}:{gw.port}"
        plant.start_line("H1R1")
        assert _wait(lambda: h.get_motor(key).rpm == 450)
        h.mod_motor(key, False)
        time.sleep(0.5)
        assert gw.writes == [] and all(r[7] != 6 for r in gw.requests)
        h.service_stop(join=True)
    finally:
        gw.stop()


# ---------------------------------------------------------------- 현장 점검 도구 (tools/ecblue_check.py)
def test_ecblue_check_tool(fast, capsys):
    import importlib.util
    import os
    spec = importlib.util.spec_from_file_location(
        "ecblue_check", os.path.join(os.path.dirname(__file__), "..", "tools", "ecblue_check.py"))
    tool = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(tool)
    gw = FakeGateway().start()
    gw.add_ecblue(180, analog_rpm=640, pin_level=0)       # h4 쓰기 전에 PIN 필요
    gw.add_ecblue(176, analog_rpm=0)
    try:
        args = ["--ip", gw.host, "--port", str(gw.port), "--ids", "180", "176", "--no-log"]
        # 1단계: 읽기만, 쓰기 0건
        assert tool.main(args) == 0
        out = capsys.readouterr().out
        assert "TCP 접속 성공" in out and "-> 640 RPM" in out and "버스 속도제어 준비: 안 됨" in out
        assert gw.writes == []
        before = {u: list(gw.ecb[u]["h"]) for u in (180, 176)}
        # 2단계: 실제 제어 코드로 500 RPM 운전 -> 판정 -> 복원
        assert tool.main(args + ["--rpm", "500", "--seconds", "3", "--yes", "--restore", "yes"]) == 0
        out = capsys.readouterr().out
        assert "ID 180: 지정 500 / 실제 500 RPM (차이 0) -> 정상" in out
        assert "ID 176: 지정 500 / 실제 500 RPM (차이 0) -> 정상" in out
        assert out.count("복원 확인") == 2 and "복원 확인 안 됨" not in out
        assert any(w[1] == modbus.ECB_H_CONTROLMODE for w in gw.writes)
        for u in (180, 176):
            assert gw.ecb[u]["h"][1:5] == before[u][1:5]           # h1~h4 원래대로
        assert gw.ecb_speed(180) == 640                             # 아날로그 속도로 복귀
    finally:
        gw.stop()
