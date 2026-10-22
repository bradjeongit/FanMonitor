"""tools/motor_tester.py 시험 (시뮬레이터 사용).

통신 동작(원본 MOTOR_TESTER 와 같은지) + 원본이 죽거나 멈추던 상황(연결 끊김, 깨진 데이터,
시작/정지 반복, 작업 쓰레드 이상)에서 계속 동작하는지 확인합니다. 화면 시험은 tkinter/디스플레이가 없으면 건너뜀.
"""
import importlib.util
import os
import socket
import struct
import sys
import threading
import time

import pytest

from motor_monitor import modbus  # noqa: E402
from motor_monitor.simulator import FakeGateway  # noqa: E402

TOOLS = os.path.join(os.path.dirname(__file__), "..", "tools")


def _load():
    spec = importlib.util.spec_from_file_location("motor_tester", os.path.join(TOOLS, "motor_tester.py"))
    mod = importlib.util.module_from_spec(spec)
    sys.modules["motor_tester"] = mod          # dataclass 처리에 필요
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture
def M(monkeypatch, tmp_path):
    m = _load()
    # 시간을 줄여 빠르게 시험 (순서와 동작은 같음)
    for name, v in (("CONNECT_TIMEOUT", 0.5), ("RETRY_WAIT", 0.1), ("SOURCE_REPLY_WAIT", 0.3),
                    ("SOURCE_PAUSE", 0.1), ("TARGET_READ_GAP", 0.3), ("TARGET_WRITE_GAP", 0.1),
                    ("TARGET_PAUSE", 0.1), ("RECV_POLL", 0.05)):
        monkeypatch.setattr(m, name, v)
    monkeypatch.setattr(m, "LOG_DIR", str(tmp_path / "log"))
    monkeypatch.setattr(m, "CONFIG_FILE", str(tmp_path / "Motor_Config.ini"))
    monkeypatch.setattr(m, "LEGACY_CONFIG_FILE", None)
    return m


@pytest.fixture
def gw():
    g = FakeGateway().start()
    yield g
    g.stop()


class Lines:
    """LogSink 대신 쓰는 기록기"""

    def __init__(self):
        self.items = []
        self.lock = threading.Lock()

    def __call__(self, target, msg):
        with self.lock:
            self.items.append((target, msg))

    def text(self, target=None):
        with self.lock:
            return "\n".join(m for t, m in self.items if target is None or t == target)


def wait_for(cond, timeout=10.0):
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        if cond():
            return True
        time.sleep(0.02)
    return False


def _settings(M, gw, **kw):
    s = M.Settings(src_ip=gw.host, src_port=gw.port, src_id=1, tgt_ip=gw.host, tgt_port=gw.port,
                   tgt_ids=[2, 3, 0, 0, 0])
    for k, v in kw.items():
        setattr(s, k, v)
    return s


def _gateway(port):
    """같은 포트로 시뮬레이터를 다시 열기 (앞 시뮬레이터의 대기 소켓이 풀릴 때까지 잠깐 재시도)"""
    end = time.monotonic() + 3
    while True:
        try:
            return FakeGateway(port=port).start()
        except OSError:
            if time.monotonic() > end:
                raise
            time.sleep(0.05)


def _free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    p = s.getsockname()[1]
    s.close()
    return p


# ---------------------------------------------------------------- 설정
def test_settings_roundtrip(M, tmp_path):
    p = str(tmp_path / "a.ini")
    s = M.Settings(manual=True, src_ip="10.0.0.1", src_port=5020, src_id=7, src_rpm=650,
                   tgt_ip="10.0.0.2", tgt_port=503, tgt_ids=[1, 2, 0, 4, 5])
    assert s.save(p)
    text = open(p, encoding="utf-8").read()
    assert "[SOURCE_CONFIG]" in text and "SOURCE_TYPE=1" in text and "TARGET_ID5=5" in text
    assert M.Settings.load(p, legacy="") == s


def test_settings_bad_values_do_not_crash(M, tmp_path):
    """원본 Get_Local_Config 와 같이 범위 밖·숫자 아닌 값은 기본값. 깨진 줄은 건너뜀"""
    p = tmp_path / "bad.ini"
    p.write_bytes("[SOURCE_CONFIG]\r\nSOURCE_PORT=99999\r\nSOURCE_ID=abc\r\nSOURCE_RPM=5000\r\n"
                  "[TARGET_CONFIG]\r\nTARGET_ID1=-3\r\nTARGET_ID2=150\r\n깨진줄\r\n".encode("cp949"))
    s = M.Settings.load(str(p), legacy="")
    assert (s.src_port, s.src_id, s.src_rpm) == (502, 1, 0)
    assert s.tgt_ids[:2] == [1, 0]


def test_settings_missing_file_and_legacy(M, tmp_path):
    assert M.Settings.load(str(tmp_path / "none.ini"), legacy="") == M.Settings()
    legacy = tmp_path / "legacy.ini"
    M.Settings(src_ip="1.2.3.4").save(str(legacy))
    assert M.Settings.load(str(tmp_path / "none.ini"), legacy=str(legacy)).src_ip == "1.2.3.4"


def test_motor_ids_skip_zero_and_duplicates(M):
    assert M.Settings(tgt_ids=[0, 3, 3, 0, 5]).motor_ids() == [3, 5]
    assert M.Settings(tgt_ids=[0, 0, 0, 0, 0]).motor_ids() == []


# ---------------------------------------------------------------- 통신 동작
def test_auto_mode_copies_source_rpm(M, gw):
    gw.add_old(1, 770)
    gw.add_new(2, automanual=0, startstop=0, rpm=0)
    gw.add_new(3, automanual=1, startstop=1, rpm=500)
    out = Lines()
    t = M.Tester(_settings(M, gw), out)
    t.start()
    try:
        assert wait_for(lambda: gw.new_regs[2][modbus.NEW_REG_RPM] == 770 and gw.new_regs[3][modbus.NEW_REG_RPM] == 770)
        assert gw.new_regs[2][modbus.NEW_REG_AUTOMANUAL] == 1 and gw.new_regs[2][modbus.NEW_REG_STARTSTOP] == 1
        # 이미 수동·동작인 3번에는 RPM 만 씀
        assert (3, modbus.NEW_REG_AUTOMANUAL, 1) not in gw.writes
        gw.add_old(1, 600)                       # 원본 회전수가 바뀌면 따라감
        assert wait_for(lambda: gw.new_regs[2][modbus.NEW_REG_RPM] == 600)
    finally:
        assert t.stop()
    src, tgt = out.text("SOURCE"), out.text("TARGET")
    assert "[ID:01][S]" in src and "(읽기결과-회전수:770)" in src
    assert "(읽기결과-자동수동:0/동작정지:0/회전수:0)" in tgt and "(쓰기요청-회전수:770)" in tgt and "(쓰기성공)" in tgt


def test_auto_mode_target_waits_for_first_source_reply(M, gw):
    gw.add_new(2)
    out = Lines()
    t = M.Tester(_settings(M, gw), out)
    t.start()
    try:
        assert wait_for(lambda: "[ID:01][E]응답없음" in out.text("SOURCE"))
        time.sleep(0.3)
        assert not any(r[7] == modbus.FC_READ_HOLDING for r in gw.requests)   # 원본: 원본 응답 전에는 신규모터 안 건드림
    finally:
        t.stop()


def test_manual_mode_and_apply(M, gw):
    gw.add_old(1, 770)
    gw.add_new(2)
    out = Lines()
    t = M.Tester(_settings(M, gw, manual=True, src_rpm=500), out)
    t.start()
    try:
        assert wait_for(lambda: gw.new_regs[2][modbus.NEW_REG_RPM] == 500)
        t.set_rpm(300)
        assert wait_for(lambda: gw.new_regs[2][modbus.NEW_REG_RPM] == 300)
        assert not any(r[7] == modbus.FC_READ_INPUT for r in gw.requests)   # 수동이면 원본 안 읽음
    finally:
        t.stop()


def test_missing_target_motor_reported(M, gw):
    gw.add_old(1, 700)
    gw.add_new(2)
    out = Lines()
    t = M.Tester(_settings(M, gw, tgt_ids=[2, 9, 0, 0, 0]), out)
    t.start()
    try:
        assert wait_for(lambda: "[ID:09][E]응답없음" in out.text("TARGET"))
        assert wait_for(lambda: gw.new_regs[2][modbus.NEW_REG_RPM] == 700)
    finally:
        t.stop()


def test_log_sink_files_and_queue(M, tmp_path):
    sink = M.LogSink(str(tmp_path / "log"))
    sink("SOURCE", "가")
    sink("TARGET", "나")
    files = sorted(os.listdir(tmp_path / "log"))
    assert files[0].startswith("source-") and files[1].startswith("target-")
    assert open(tmp_path / "log" / files[0], encoding="utf-8").read().rstrip().endswith("]가")
    assert [t for t, _ in sink.ui] == ["SOURCE", "TARGET"]


def test_log_sink_survives_unwritable_dir(M, tmp_path):
    f = tmp_path / "file"
    f.write_text("x")
    sink = M.LogSink(str(f / "log"))   # 폴더를 만들 수 없는 경로
    sink("SOURCE", "a")
    assert sink.file_errors == 1 and len(sink.ui) == 1


# ---------------------------------------------------------------- 안정성 (원본 다운 원인)
def test_connect_fail_then_recover(M):
    """통신장치가 꺼져 있다가 켜져도 계속 동작 (원본: 접속 시험 소켓을 닫지 않아 연결·핸들이 쌓임)"""
    port = _free_port()
    out = Lines()
    s = M.Settings(src_ip="127.0.0.1", src_port=port, tgt_ip="127.0.0.1", tgt_port=port, tgt_ids=[2, 0, 0, 0, 0])
    t = M.Tester(s, out)
    t.start()
    try:
        assert wait_for(lambda: out.text("SOURCE").count("접속실패!") >= 3)
        g = _gateway(port)
        try:
            g.add_old(1, 640)
            g.add_new(2)
            assert wait_for(lambda: g.new_regs[2][modbus.NEW_REG_RPM] == 640)
        finally:
            g.stop()
    finally:
        assert t.stop()


def test_gateway_drop_and_reconnect(M):
    port = _free_port()
    g = _gateway(port)
    g.add_old(1, 700)
    g.add_new(2)
    out = Lines()
    t = M.Tester(M.Settings(src_ip="127.0.0.1", src_port=port, tgt_ip="127.0.0.1", tgt_port=port,
                            tgt_ids=[2, 0, 0, 0, 0]), out)
    t.start()
    try:
        assert wait_for(lambda: g.new_regs[2][modbus.NEW_REG_RPM] == 700)
        g.stop()
        assert wait_for(lambda: "[E]통신 끊김" in out.text() or "접속실패!" in out.text())
        g2 = _gateway(port)
        try:
            g2.add_old(1, 710)
            g2.add_new(2)
            assert wait_for(lambda: g2.new_regs[2][modbus.NEW_REG_RPM] == 710)
        finally:
            g2.stop()
    finally:
        assert t.stop()
    assert t.restarts == 0 and t.errors == 0


class NastyServer:
    """깨진 데이터, 나뉜 프레임, 붙은 프레임, 예외 응답을 섞어 보내는 장치"""

    def __init__(self):
        self.srv = socket.socket()
        self.srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.srv.bind(("127.0.0.1", 0))
        self.srv.listen(5)
        self.port = self.srv.getsockname()[1]
        self.stop_ev = threading.Event()
        threading.Thread(target=self._accept, daemon=True).start()

    def _accept(self):
        self.srv.settimeout(0.1)
        while not self.stop_ev.is_set():
            try:
                c, _ = self.srv.accept()
            except (socket.timeout, OSError):
                continue
            threading.Thread(target=self._client, args=(c,), daemon=True).start()

    def _client(self, c):
        c.settimeout(0.1)
        n = 0
        while not self.stop_ev.is_set():
            try:
                req = c.recv(512)
            except socket.timeout:
                continue
            except OSError:
                return
            if not req:
                return
            if len(req) < 8:                     # 요청이 나뉘어 온 경우 (시험 장치는 무시)
                continue
            n += 1
            tid, uid, fc = req[0:2], req[6], req[7]
            if fc == modbus.FC_READ_INPUT:
                good = tid + struct.pack(">HHBBBH", 0, 5, uid, 4, 2, 555)
            else:
                good = tid + struct.pack(">HHBBB7H", 0, 17, uid, 3, 14, 0, 0, 0, 0, 0, 0, 0)
            try:
                kind = n % 5
                if kind == 0:
                    c.sendall(os.urandom(37))                       # 쓰레기
                elif kind == 1:
                    c.sendall(good[:5]); time.sleep(0.02); c.sendall(good[5:])   # 나뉨
                elif kind == 2:
                    c.sendall(good + good)                          # 붙음
                elif kind == 3:
                    c.sendall(tid + struct.pack(">HHBBB", 0, 3, uid, fc | 0x80, 11))   # 예외 응답
                else:
                    c.close()                                       # 갑자기 끊음
                    return
            except OSError:
                return

    def close(self):
        self.stop_ev.set()
        self.srv.close()


def test_garbage_and_broken_frames_do_not_kill_threads(M):
    srv = NastyServer()
    out = Lines()
    t = M.Tester(M.Settings(src_ip="127.0.0.1", src_port=srv.port, tgt_ip="127.0.0.1", tgt_port=srv.port,
                            tgt_ids=[2, 3, 0, 0, 0]), out)
    t.start()
    try:
        assert wait_for(lambda: "(읽기결과-회전수:555)" in out.text("SOURCE"))
        assert wait_for(lambda: out.text("TARGET").count("[ID:0") > 20, timeout=30)
        assert all(th.is_alive() for th in t._threads.values())
    finally:
        assert t.stop()
        srv.close()
    assert t.restarts == 0 and t.errors == 0
    assert "(예외응답:11)" in out.text()


def test_handler_error_is_logged_and_recovered(M, gw, monkeypatch):
    """응답 처리 중 예상 못한 오류가 나도 쓰레드가 죽지 않고 다시 접속해 계속함"""
    gw.add_old(1, 720)
    gw.add_new(2)
    real = M.modbus.parse_new_response
    calls = {"n": 0}

    def flaky(f):
        calls["n"] += 1
        if calls["n"] == 1:
            raise ValueError("시험용 오류")
        return real(f)

    monkeypatch.setattr(M.modbus, "parse_new_response", flaky)
    out = Lines()
    t = M.Tester(_settings(M, gw, tgt_ids=[2, 0, 0, 0, 0]), out)
    t.start()
    try:
        assert wait_for(lambda: gw.new_regs[2][modbus.NEW_REG_RPM] == 720)
    finally:
        t.stop()
    assert t.errors == 1 and "[E]처리 오류 (ValueError: 시험용 오류)" in out.text("TARGET")


def test_supervisor_restarts_dead_thread(M, gw, monkeypatch):
    gw.add_old(1, 730)
    gw.add_new(2)
    t = M.Tester(_settings(M, gw, tgt_ids=[2, 0, 0, 0, 0]), Lines())
    real = t._flush_writes
    state = {"killed": False}

    def killer(*a):
        if not state["killed"]:
            state["killed"] = True
            raise M._Restart()
        return real(*a)

    monkeypatch.setattr(t, "_flush_writes", killer)
    t.start()
    try:
        assert wait_for(lambda: not t._threads["target"].is_alive())
        assert t.supervise() and t.restarts == 1
        assert wait_for(lambda: gw.new_regs[2][modbus.NEW_REG_RPM] == 730)
    finally:
        t.stop()


def test_supervisor_unblocks_stalled_thread(M, gw, monkeypatch):
    gw.add_old(1, 700)
    t = M.Tester(_settings(M, gw, tgt_ids=[2, 0, 0, 0, 0]), Lines())
    t.start()
    try:
        assert wait_for(lambda: t._src.connected)
        monkeypatch.setattr(M, "STALL_LIMIT", 0.0)
        t._beats["source"] = time.monotonic() - 1
        notes = t.supervise()
        assert any("진행이 없어" in n for n in notes)
    finally:
        t.stop()


def _fd_count():
    try:
        return len(os.listdir("/proc/self/fd"))
    except OSError:
        return None


def test_start_stop_cycles_no_leak(M):
    """시작/정지를 반복해도 쓰레드·소켓이 남지 않고, 정지가 바로 됨 (원본: Thread.Abort, 소켓 미정리)"""
    base_threads, base_fd = threading.active_count(), _fd_count()
    gw = FakeGateway().start()
    gw.add_old(1, 700)
    gw.add_new(2)
    for _ in range(25):
        t = M.Tester(_settings(M, gw), Lines())
        t.start()
        time.sleep(0.05)
        t0 = time.monotonic()
        assert t.stop()
        assert time.monotonic() - t0 < 1.5
    gw.stop()                                   # 시뮬레이터 쪽 소켓도 정리한 뒤 비교
    # 정리가 끝날 때까지 기다림 (누수라면 아무리 기다려도 줄지 않음)
    assert wait_for(lambda: threading.active_count() <= base_threads + 1, timeout=5)
    if base_fd is not None:
        assert wait_for(lambda: _fd_count() <= base_fd + 2, timeout=5), _fd_count()


def test_soak_with_flapping_gateway(M):
    """통신장치가 계속 켜졌다 꺼지는 상황에서 일정 시간 실행 (짧은 장시간 운전 시험)"""
    port = _free_port()
    out = Lines()
    t = M.Tester(M.Settings(src_ip="127.0.0.1", src_port=port, tgt_ip="127.0.0.1", tgt_port=port,
                            tgt_ids=[2, 3, 4, 0, 0]), out)
    base_fd = _fd_count()
    t.start()
    try:
        end = time.monotonic() + 6
        rpm = 500
        while time.monotonic() < end:
            g = _gateway(port)
            g.add_old(1, rpm)
            for u in (2, 3, 4):
                g.add_new(u)
            time.sleep(0.8)
            g.stop()
            time.sleep(0.2)
            rpm += 10
            assert all(th.is_alive() for th in t._threads.values())
    finally:
        assert t.stop()
    assert t.restarts == 0 and t.errors == 0
    assert "(쓰기성공)" in out.text("TARGET")
    if base_fd is not None:
        assert wait_for(lambda: _fd_count() <= base_fd + 4, timeout=5), _fd_count()


# ---------------------------------------------------------------- 속도를 계속 바꿀 때 (현장 다운 상황)
def test_auto_mode_rapid_source_rpm_changes(M, gw):
    """기존모터 회전수가 계속 바뀌어도 쓰기가 밀리지 않고, 마지막 값으로 맞춰지며, 오류·재시작 없음"""
    gw.add_old(1, 500)
    for u in (2, 3, 4):
        gw.add_new(u)
    out = Lines()
    t = M.Tester(_settings(M, gw, tgt_ids=[2, 3, 4, 0, 0]), out)
    t.start()
    try:
        end = time.monotonic() + 6
        rpm = 100
        while time.monotonic() < end:
            rpm = 100 + (rpm + 137) % 700
            gw.add_old(1, rpm)
            time.sleep(0.15)
            assert len(t._writes) <= 3 * 3          # 대기 중인 쓰기는 모터×레지스터 수를 넘지 않음
        assert wait_for(lambda: all(gw.new_regs[u][modbus.NEW_REG_RPM] == rpm for u in (2, 3, 4)))
        assert all(th.is_alive() for th in t._threads.values())
    finally:
        assert t.stop()
    assert t.restarts == 0 and t.errors == 0
    assert "[E]" not in out.text("TARGET")


def test_manual_mode_apply_hammered(M, gw):
    """[적용]을 아주 빠르게 계속 눌러도(다른 쓰레드에서 값 변경) 마지막 값이 반영되고 오류 없음"""
    for u in (2, 3):
        gw.add_new(u)
    out = Lines()
    t = M.Tester(_settings(M, gw, manual=True, src_rpm=400, tgt_ids=[2, 3, 0, 0, 0]), out)
    t.start()
    stop = threading.Event()
    last = {"v": 400}

    def clicker():
        v = 0
        while not stop.is_set():
            v = (v + 37) % 801
            t.set_rpm(v)
            last["v"] = v
            time.sleep(0.002)

    th = threading.Thread(target=clicker, daemon=True)
    th.start()
    try:
        time.sleep(5)
        stop.set()
        th.join()
        assert wait_for(lambda: all(gw.new_regs[u][modbus.NEW_REG_RPM] == last["v"] for u in (2, 3)))
    finally:
        assert t.stop()
    assert t.restarts == 0 and t.errors == 0


class WriteHostileGateway(FakeGateway):
    """쓰기를 받으면 가끔 연결을 끊거나, 늦게 답하거나, 예외로 답하는 통신장치"""

    def __init__(self):
        super().__init__()
        self.n = 0

    def handle(self, req):
        if len(req) >= 8 and req[7] == modbus.FC_WRITE_SINGLE:
            self.n += 1
            k = self.n % 4
            if k == 1:
                raise ConnectionAbortedError("시험: 쓰기 중 끊음")
            if k == 2:
                time.sleep(0.4)                      # 다음 요청 시간까지 늦게 답함
            if k == 3:
                tid, uid = req[0:2], req[6]
                return tid + struct.pack(">HHBBB", 0, 3, uid, 0x86, 4)   # 장치 오류 예외
        return super().handle(req)

    def _client_loop(self, c):
        try:
            super()._client_loop(c)
        except ConnectionAbortedError:
            pass
        finally:
            try:
                c.close()
            except OSError:
                pass


def test_speed_changes_with_hostile_write_responses(M):
    """속도 변경(쓰기) 때 통신장치가 끊거나 늦거나 예외로 답해도 프로그램은 계속 동작하고,
    신규모터 쪽 문제로 기존모터 연결이 끊기지 않음 (원본: 신규모터 오류 때 기존모터 연결을 닫음)"""
    src = FakeGateway().start()
    src.add_old(1, 500)
    tgt = WriteHostileGateway().start()
    for u in (2, 3, 4):
        tgt.add_new(u)
    out = Lines()
    s = M.Settings(src_ip=src.host, src_port=src.port, tgt_ip=tgt.host, tgt_port=tgt.port, tgt_ids=[2, 3, 4, 0, 0])
    t = M.Tester(s, out)
    t.start()
    try:
        end = time.monotonic() + 8
        rpm = 500
        while time.monotonic() < end:
            rpm = 300 + (rpm + 53) % 500
            src.add_old(1, rpm)
            time.sleep(0.5)
            assert all(th.is_alive() for th in t._threads.values())
        src.add_old(1, 640)
        # 오류가 반복돼도 뒤쪽 모터(4번)까지 차례가 돌아와야 함 (재접속 후 첫 모터부터 다시 시작하면 굶음)
        assert wait_for(lambda: all(tgt.new_regs[u][modbus.NEW_REG_RPM] == 640 for u in (2, 3, 4)), timeout=20), \
            {u: tgt.new_regs[u][modbus.NEW_REG_RPM] for u in (2, 3, 4)}
    finally:
        assert t.stop()
        src.stop()
        tgt.stop()
    assert t.restarts == 0 and t.errors == 0
    assert "[E]통신 끊김" in out.text("TARGET") and "(예외응답:4)" in out.text("TARGET")
    assert "[E]" not in out.text("SOURCE")            # 신규모터 문제가 기존모터 통신에 영향 없음


class DropSecondWriteGateway(FakeGateway):
    """연결마다 두 번째 쓰기를 받으면 연결을 끊는 통신장치 (쓰기는 반영 안 됨)"""

    def __init__(self):
        super().__init__()
        self._local = threading.local()        # 연결(접속 처리 쓰레드)마다 쓰기 횟수

    def _client_loop(self, c):
        self._local.writes = 0
        try:
            super()._client_loop(c)
        except ConnectionAbortedError:
            pass
        finally:
            try:
                c.close()
            except OSError:
                pass

    def handle(self, req):
        if len(req) >= 8 and req[7] == modbus.FC_WRITE_SINGLE:
            self._local.writes += 1
            if self._local.writes == 2:
                raise ConnectionAbortedError("시험: 두 번째 쓰기에서 끊음")
        return super().handle(req)


def test_errors_every_cycle_do_not_starve_later_motors(M):
    """오류로 다시 접속할 때마다 첫 모터부터 다시 하면, 오류가 매번 나는 경우 뒤쪽 모터는 영영 차례가 오지 않음.
    다시 접속해도 멈춘 모터부터 이어서 해야 모든 모터에 속도가 써짐."""
    g = DropSecondWriteGateway().start()
    for u in (2, 3, 4):
        g.add_new(u, automanual=1, startstop=1, rpm=0)
    t = M.Tester(M.Settings(manual=True, src_rpm=100, tgt_ip=g.host, tgt_port=g.port, tgt_ids=[2, 3, 4, 0, 0]), Lines())
    stop = threading.Event()

    def changer():                                   # 속도를 계속 바꿈 -> 매 바퀴 모든 모터에 쓰기 발생
        v = 100
        while not stop.is_set():
            v = 100 + (v + 7) % 700
            t.set_rpm(v)
            time.sleep(0.02)

    th = threading.Thread(target=changer, daemon=True)
    t.start()
    th.start()
    try:
        assert wait_for(lambda: {u for u, reg, _ in list(g.writes) if reg == modbus.NEW_REG_RPM} >= {2, 3, 4}, timeout=15), \
            sorted({u for u, reg, _ in g.writes})
    finally:
        stop.set()
        th.join()
        t.stop()
        g.stop()


class BusyFirstWriteGateway(FakeGateway):
    """쓰기를 번갈아 '장치 사용 중'(예외 6)으로 거절하는 통신장치"""

    def __init__(self):
        super().__init__()
        self.rejects = 0
        self._flip = False

    def handle(self, req):
        if len(req) >= 8 and req[7] == modbus.FC_WRITE_SINGLE:
            self._flip = not self._flip
            if self._flip:
                self.rejects += 1
                return req[0:2] + struct.pack(">HHBBB", 0, 3, req[6], 0x86, 6)
        return super().handle(req)


def test_rejected_write_is_retried_once(M, monkeypatch):
    """속도 쓰기가 거절되면 다음 바퀴까지 기다리지 않고 바로 1번 다시 보냄"""
    monkeypatch.setattr(M, "TARGET_PAUSE", 30.0)     # 다음 바퀴가 오기 전에 반영돼야 함
    g = BusyFirstWriteGateway().start()
    g.add_new(2, automanual=1, startstop=1, rpm=0)
    out = Lines()
    t = M.Tester(M.Settings(manual=True, src_rpm=450, tgt_ip=g.host, tgt_port=g.port, tgt_ids=[2, 0, 0, 0, 0]), out)
    t.start()
    try:
        assert wait_for(lambda: g.new_regs[2][modbus.NEW_REG_RPM] == 450, timeout=5)
    finally:
        t.stop()
        g.stop()
    assert g.rejects == 1 and "(쓰기요청-회전수:450, 재시도)" in out.text("TARGET")


# ---------------------------------------------------------------- 화면
@pytest.fixture
def app(M, tmp_path, monkeypatch):
    tk = pytest.importorskip("tkinter")
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("디스플레이 없음")
    root.withdraw()
    warnings = []
    monkeypatch.setattr(M.messagebox, "showwarning", lambda *a, **k: warnings.append(a[1]))
    a = M.App(root, sink=M.LogSink(str(tmp_path / "log")), config_file=str(tmp_path / "Motor_Config.ini"))
    a.warnings = warnings
    yield a
    try:
        if a.tester is not None:
            a.stop()
        root.destroy()
    except tk.TclError:
        pass


def _pump(app, seconds):
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        app.root.update()
        time.sleep(0.02)


def test_ui_requires_target_id(M, app):
    for sp in app.sp_tgt_ids:
        sp.var.set("0")
    app.on_start()
    assert app.tester is None and app.warnings == ["신규모터 ID값을 설정하세요."]


def test_ui_rejects_bad_ip(M, app):
    app.ent_tgt_ip.delete(0, "end")
    app.ent_tgt_ip.insert(0, "1.2.3")
    app.on_start()
    assert app.tester is None and "IP 형식" in app.warnings[0]


def test_ui_start_stop_and_logs(M, app, gw):
    gw.add_old(1, 760)
    gw.add_new(2)
    app.load_settings(_settings(M, gw, tgt_ids=[2, 0, 0, 0, 0]))
    app.on_start()
    assert app.tester is not None and app.btn_start.cget("text") == "■"
    assert str(app.ent_src_ip.cget("state")) == "disabled"
    end = time.monotonic() + 10
    while time.monotonic() < end and gw.new_regs[2][modbus.NEW_REG_RPM] != 760:
        _pump(app, 0.1)
    assert gw.new_regs[2][modbus.NEW_REG_RPM] == 760
    _pump(app, 0.3)
    assert "(읽기결과-회전수:760)" in app.txt_src.get("1.0", "end")
    assert "쓰기요청-회전수:760" in app.txt_tgt.get("1.0", "end")
    assert "설정 회전수 760" in app.status.cget("text")
    app.on_start()                                   # ■ -> 정지
    assert app.tester is None and app.btn_start.cget("text") == "▶"
    assert str(app.ent_src_ip.cget("state")) == "normal"
    assert M.Settings.load(app.config_file, legacy="").tgt_ids[0] == 2   # 시작할 때 설정 저장


def test_ui_manual_apply_button(M, app, gw):
    gw.add_new(2)
    app.load_settings(_settings(M, gw, manual=True, src_rpm=400, tgt_ids=[2, 0, 0, 0, 0]))
    app.on_start()
    assert app.btn_apply.winfo_manager() == "place" and str(app.sp_src_rpm.cget("state")) == "normal"
    app.sp_src_rpm.var.set("250")
    app.on_apply()
    end = time.monotonic() + 10
    while time.monotonic() < end and gw.new_regs[2][modbus.NEW_REG_RPM] != 250:
        _pump(app, 0.1)
    assert gw.new_regs[2][modbus.NEW_REG_RPM] == 250
    app.stop()
    assert app.btn_apply.winfo_manager() == ""


def test_ui_log_window_is_capped(M, app, monkeypatch):
    """로그창이 끝없이 커지지 않음 (원본: TextBox 에 계속 추가 -> 장시간 운전 시 느려지고 멈춤)"""
    monkeypatch.setattr(M, "UI_MAX_LINES", 200)
    for i in range(1500):
        app.sink("TARGET", f"줄 {i}")
    _pump(app, 0.5)
    text = app.txt_tgt.get("1.0", "end").strip().splitlines()
    assert len(text) <= 200 and text[-1].endswith("줄 1499")


def test_ui_apply_clicked_rapidly(M, app, gw):
    gw.add_new(2)
    app.load_settings(_settings(M, gw, manual=True, src_rpm=400, tgt_ids=[2, 0, 0, 0, 0]))
    app.on_start()
    for i in range(300):
        app.sp_src_rpm.var.set(str(i * 7 % 801))
        app.btn_apply.invoke()
        if i % 20 == 0:
            app.root.update()
    final = 299 * 7 % 801
    end = time.monotonic() + 10
    while time.monotonic() < end and gw.new_regs[2][modbus.NEW_REG_RPM] != final:
        _pump(app, 0.1)
    assert gw.new_regs[2][modbus.NEW_REG_RPM] == final
    assert app.tester.restarts == 0 and app.tester.errors == 0
    app.stop()


def test_ui_close_saves_settings(M, app, gw):
    app.load_settings(_settings(M, gw, src_rpm=123))
    path = app.config_file
    app.on_close()
    assert M.Settings.load(path, legacy="").src_rpm == 123
