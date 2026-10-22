"""tools/motor_scanner.py 시험 (시뮬레이터 사용). 화면 시험은 tkinter/디스플레이가 없으면 건너뜀."""
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
    spec = importlib.util.spec_from_file_location("motor_scanner", os.path.join(TOOLS, "motor_scanner.py"))
    mod = importlib.util.module_from_spec(spec)
    sys.modules["motor_scanner"] = mod
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture
def S(monkeypatch, tmp_path):
    m = _load()
    monkeypatch.setattr(m, "REQUEST_GAP", 0)
    monkeypatch.setattr(m, "WRITE_VERIFY_GAP", 0)
    monkeypatch.setattr(m, "LOG_DIR", str(tmp_path))
    monkeypatch.setattr(m, "SETTINGS_FILE", str(tmp_path / "motor_scanner.json"))
    return m


@pytest.fixture
def gw():
    g = FakeGateway().start()
    yield g
    g.stop()


def client(S, gw, timeout=0.3):
    return S.Client(gw.host, gw.port, timeout)


# ---------------------------------------------------------------- 판별·해석
def test_classify_product_codes(S):
    assert S.classify(0x0100)[0] == S.KIND_ECBLUE and S.classify(0x0032)[0] == S.KIND_ECBLUE
    assert S.classify(0x0103) == (S.KIND_ECBLUE, "ECblue (EC90)")
    assert S.classify(0x0401)[0] == S.KIND_AM and S.is_ecblue(S.KIND_AM)
    assert S.classify(0x0201)[0] == S.KIND_ZA and not S.is_ecblue(S.KIND_ZA)
    assert S.fw_text(1430) == "14.30" and S.fw_text(None) == "-"


def test_watchdog_encode_decode(S):
    assert S.wd_encode(30, 2) == 0x1E02 and S.wd_decode(0x1E02) == (30, 2)
    assert S.wd_text(0) == "꺼짐" and S.wd_text(0x0002) == "꺼짐"           # 시간 0 = 꺼짐
    assert S.wd_text(0x1E02, 450) == "켜짐: 30초 끊기면 정속 운전 (Set Intern1) 450 rpm"
    with pytest.raises(ValueError):
        S.wd_encode(256, 2)
    # 속도(h5)는 바뀔 때만 씀 (영구저장 수명 보호)
    assert S.watchdog_changes(450, 30, 2, 450) == [(17, 0x1E02)]
    assert S.watchdog_changes(0, 30, 2, 450) == [(5, 450), (17, 0x1E02)]
    assert S.watchdog_changes(0, 10, 7, None) == [(17, 0x0A07)]


def test_bits_text(S):
    assert S.bits_text(0, S.ERR_BITS, "정상") == "0x0000 정상"
    assert "통신 워치독" in S.bits_text(0x8000, S.ERR_BITS)


# ---------------------------------------------------------------- 스캔
def test_scan_finds_devices_and_kinds(S, gw):
    gw.add_ecblue(3, analog_rpm=640)
    gw.add_ecblue(7, analog_rpm=0, product=0x0032, firmware=1210)
    gw.add_new(5, automanual=1, startstop=1, rpm=550)
    c = client(S, gw, timeout=0.15)
    found = {uid: S.scan_id(c, uid) for uid in range(1, 9)}
    assert {u for u, d in found.items() if d} == {3, 5, 7}
    assert found[3]["kind"] == S.KIND_ECBLUE and found[3]["speed"] == 640 and found[3]["fw"] == 1430
    assert found[7]["name"] == "ECblue Basic Standard (구형)"
    assert found[5]["kind"] == S.KIND_NEW and "설정 550 rpm" in found[5]["name"]
    assert not any(r[7] == modbus.FC_WRITE_SINGLE for r in gw.requests)   # 스캔은 읽기만


def test_scan_without_new_motor_probe(S, gw):
    gw.add_new(5)
    c = client(S, gw, timeout=0.15)
    assert S.scan_id(c, 5, find_new=False) is None


class ExceptionGateway(FakeGateway):
    """없는 장치에는 게이트웨이 예외 0x0B, ID 9 는 모르는 장치(예외 0x01)로 응답"""

    def handle(self, req):
        uid, fc = req[6], req[7]
        r = super().handle(req)
        if r is not None:
            return r
        code = 0x01 if uid == 9 else 0x0B
        return req[0:2] + struct.pack(">HHBBB", 0, 3, uid, fc | 0x80, code)


def test_scan_gateway_exceptions():
    S = _load()
    S.REQUEST_GAP = 0
    g = ExceptionGateway().start()
    try:
        g.add_ecblue(2)
        c = S.Client(g.host, g.port, 2.0)
        t0 = time.monotonic()
        assert S.scan_id(c, 4) is None                           # 0x0B -> 장치 없음, 기다리지 않음
        assert time.monotonic() - t0 < 1.0
        d = S.scan_id(c, 9)
        assert d["kind"] == S.KIND_OTHER                         # 예외라도 응답이 있으면 '있음'
        assert S.scan_id(c, 2)["kind"] == S.KIND_ECBLUE
    finally:
        g.stop()


def _run(w, timeout=30):
    """작업 하나가 끝날 때까지 결과 큐를 비움 (진행 알림은 화면 쓰레드처럼 실행). 마지막 결과를 돌려줌."""
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        done, res = w.results.get(timeout=timeout)
        if done is None:
            return res
        done(res)
    raise TimeoutError


def test_scan_job_range_progress_and_cancel(S, gw):
    for u in (2, 4, 6):
        gw.add_ecblue(u)
    w = S.Worker()
    w.start()
    seen, asked = [], []

    def prog(p):
        asked.append(p["uid"])
        if "device" in p:
            seen.append(p["device"]["uid"])
    w.submit(S.job_scan, gw.host, gw.port, 2, 6, 0.1, False, prog)
    res = _run(w)
    assert res.get("done") and res["found"] == 3 and seen == [2, 4, 6]
    assert min(asked) == 2 and max(asked) == 6                     # 범위 밖 ID 는 묻지 않음
    assert {r[6] for r in gw.requests} <= {2, 3, 4, 5, 6}
    # 중지: ID 3 을 확인할 때 중지 요청
    w.submit(S.job_scan, gw.host, gw.port, 1, 240, 0.1, True, lambda p: w.cancel.set() if p["uid"] >= 3 else None)
    res = _run(w)
    assert res.get("cancelled") and res["last"] <= 4
    w.stop()


def test_scan_connect_error(S):
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    w = S.Worker()
    w.start()
    w.submit(S.job_scan, "127.0.0.1", port, 1, 3, 0.1, False, lambda p: None)
    _, res = w.results.get(timeout=10)
    assert "접속 실패" in res["error"]
    w.stop()


# ---------------------------------------------------------------- 읽기·쓰기
def test_read_all_falls_back_when_register_missing(S, gw):
    gw.add_ecblue(3, firmware=1310)          # FW14 미만: i13 없음 -> 구간 읽기 실패 후 하나씩
    data = S.read_all(client(S, gw), 3)
    assert data["i"][0] == 1310 and data["i"][13] is None and data["i"][14] == 770
    assert data["h"][16] == 513 and data["h"][18] == 75 and data["h"][25] == 0x1414


def test_write_needs_pin_and_verifies(S, gw):
    gw.add_ecblue(3, pin_level=0)
    c = client(S, gw)
    steps = S.write_regs(c, 3, [(8, 1000), (7, 200)])
    assert [lv for lv, _ in steps] == ["info", "info", "info"] and "PIN 보호 해제" in steps[0][1]
    assert (3, 0, 10) in gw.writes and gw.ecb[3]["h"][8] == 1000 and gw.ecb[3]["h"][7] == 200


def test_write_rejected_is_reported(S, gw):
    gw.add_ecblue(3)
    steps = S.write_regs(client(S, gw), 3, [(8, 7000)])            # 범위 밖 -> 장치가 거부
    assert steps[-1][0] == "error" and "예외 3" in steps[-1][1]
    assert gw.ecb[3]["h"][8] == 0


def test_max_speed_limits_actual_speed(S, gw):
    gw.add_ecblue(3, analog_rpm=900)
    S.write_regs(client(S, gw), 3, [(8, 700)])
    assert gw.ecb_speed(3) == 700


def test_watchdog_on_off_and_trip(S, gw):
    gw.add_ecblue(3, analog_rpm=800, pin_level=1)
    c = client(S, gw)
    steps = S.write_regs(c, 3, S.watchdog_changes(0, 1, 2, 450))
    assert all(lv == "info" for lv, _ in steps)
    assert gw.ecb[3]["h"][17] == 0x0102 and gw.ecb[3]["h"][5] == 450
    assert gw.ecb[3]["nv_writes"] == {5: 1, 17: 1}
    # 통신이 1초 넘게 없으면 h5 속도로, 오류 bit15
    c.close()
    time.sleep(1.3)
    i = S.read_regs(c, "i", [(12, 1), (14, 1)], 3)
    assert i[14] == 450 and i[12] & 0x8000
    # 끄기
    S.write_regs(c, 3, [(17, 0)])
    assert gw.ecb[3]["h"][17] == 0
    c.close()
    time.sleep(1.3)
    i = S.read_regs(c, "i", [(12, 1), (14, 1)], 3)
    assert i[14] == 800 and not i[12] & 0x8000


def test_watchdog_test_job(S, gw):
    gw.add_ecblue(3, analog_rpm=800)
    gw.ecb[3]["h"][5], gw.ecb[3]["h"][17] = 450, S.wd_encode(1, 2)
    w = S.Worker()
    w.start()
    prog = []
    w.submit(S.job_wd_test, gw.host, gw.port, 0.3, 3, 1.5, lambda p: prog.append(p))
    res = _run(w)
    w.stop()
    assert prog
    assert res["first"][14] == 450 and res["first"][12] & 0x8000
    assert res["after"][14] == 800                                  # 통신이 다시 오면 원래대로 (시뮬레이터)


# ---------------------------------------------------------------- 설정 파일
def test_settings_roundtrip(S, tmp_path):
    p = str(tmp_path / "s.json")
    assert S.load_settings(p)["first"] == 1 and S.load_settings(p)["last"] == 240
    S.save_settings({"ip": "10.0.0.5", "port": 503, "first": 10, "last": 20, "timeout": 0.3, "find_new": False}, p)
    s = S.load_settings(p)
    assert (s["ip"], s["port"], s["first"], s["last"], s["find_new"]) == ("10.0.0.5", 503, 10, 20, False)


# ---------------------------------------------------------------- 화면
@pytest.fixture
def app(S, tmp_path, monkeypatch):
    tk = pytest.importorskip("tkinter")
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("디스플레이 없음")
    root.withdraw()
    msgs = []
    monkeypatch.setattr(S.messagebox, "showwarning", lambda *a, **k: msgs.append(("warn", a[1])))
    monkeypatch.setattr(S.messagebox, "showinfo", lambda *a, **k: msgs.append(("info", a[1])))
    monkeypatch.setattr(S.messagebox, "askyesno", lambda *a, **k: True)
    a = S.App(root, settings_file=str(tmp_path / "s.json"))
    a.msgs = msgs
    yield a
    try:
        a.on_close()
    except tk.TclError:
        pass


def pump(app, cond, timeout=15.0):
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        app.root.update()
        if cond():
            return True
        time.sleep(0.02)
    return False


def _setup(app, gw, first=1, last=8):
    app.v_ip.set(gw.host)
    app.v_port.set(str(gw.port))
    app.v_first.set(str(first))
    app.v_last.set(str(last))
    app.v_timeout.set("0.1")


def test_ui_defaults_and_bad_range(S, app):
    assert (app.v_first.get(), app.v_last.get()) == ("1", "240")
    app.v_first.set("50")
    app.v_last.set("10")
    app.on_scan()
    assert not app.scanning and "ID 범위" in app.msgs[-1][1]


def test_ui_scan_list_and_detail_edit(S, app, gw):
    gw.add_ecblue(3, analog_rpm=640, pin_level=0)
    gw.add_new(5)
    _setup(app, gw)
    app.on_scan()
    assert app.scanning
    assert pump(app, lambda: not app.scanning)
    assert sorted(int(i) for i in app.tree.get_children()) == [3, 5]
    assert app.tree.item("3")["values"][1] == S.KIND_ECBLUE and app.tree.item("5")["values"][1] == S.KIND_NEW
    # 신규모터는 설정 열기 불가
    app.tree.selection_set("5")
    app.root.update()
    assert str(app.btn_open.cget("state")) == "disabled"
    app.tree.selection_set("3")
    app.root.update()
    assert str(app.btn_open.cget("state")) == "normal"
    app.on_open()
    win = app.details[3]
    assert pump(app, lambda: win.data is not None)
    assert win.cur_labels[8].cget("text") == "0"
    # 최고속도·최저속도 변경
    win.new_vars[8].set("1100")
    win.new_vars[7].set("150")
    win.on_apply()
    assert pump(app, lambda: gw.ecb[3]["h"][8] == 1100 and gw.ecb[3]["h"][7] == 150)
    assert pump(app, lambda: win.cur_labels[8].cget("text") == "1100")
    # 최저 > 최고 는 거부
    win.new_vars[7].set("2000")
    win.on_apply()
    assert "최저속도" in app.msgs[-1][1] and gw.ecb[3]["h"][7] == 150


def test_ui_watchdog_on_off(S, app, gw):
    gw.add_ecblue(3, analog_rpm=640)
    _setup(app, gw, 3, 3)
    app.on_scan()
    assert pump(app, lambda: not app.scanning)
    app.tree.selection_set("3")
    app.on_open()
    win = app.details[3]
    assert pump(app, lambda: win.data is not None)
    win.v_wd_sec.set("20")
    win.v_wd_mode.set(S.WD_MODES[2])
    win.v_wd_speed.set("500")
    win.on_wd_on()
    assert pump(app, lambda: gw.ecb[3]["h"][17] == 0x1402 and gw.ecb[3]["h"][5] == 500)
    assert pump(app, lambda: "켜짐: 20초" in win.lb_wd.cget("text"))
    assert "켜짐: 20초" in app.tree.item("3")["values"][5]
    win.on_wd_off()
    assert pump(app, lambda: gw.ecb[3]["h"][17] == 0)
    assert pump(app, lambda: win.lb_wd.cget("text") == "현재: 꺼짐")


def test_ui_watchdog_speed_above_max_rejected(S, app, gw):
    gw.add_ecblue(3)
    gw.ecb[3]["h"][8] = 600
    _setup(app, gw, 3, 3)
    app.on_scan()
    assert pump(app, lambda: not app.scanning)
    app.tree.selection_set("3")
    app.on_open()
    win = app.details[3]
    assert pump(app, lambda: win.data is not None)
    win.v_wd_speed.set("900")
    win.on_wd_on()
    assert "최고속도" in app.msgs[-1][1] and gw.ecb[3]["h"][17] == 0
