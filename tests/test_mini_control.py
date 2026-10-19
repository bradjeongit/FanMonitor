"""tools/mini_control.py 시험 (시뮬레이터 사용). 화면 시험은 디스플레이가 없으면 건너뜀."""
import importlib.util
import json
import os
import sys
import time

import pytest

tk = pytest.importorskip("tkinter")

from motor_monitor import modbus  # noqa: E402
from motor_monitor.simulator import FakeGateway  # noqa: E402

TOOLS = os.path.join(os.path.dirname(__file__), "..", "tools")


def _load():
    sys.path.insert(0, TOOLS)
    spec = importlib.util.spec_from_file_location("mini_control", os.path.join(TOOLS, "mini_control.py"))
    mod = importlib.util.module_from_spec(spec)
    sys.modules["mini_control"] = mod          # dataclass 처리에 필요
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture
def M(monkeypatch, tmp_path):
    m = _load()
    monkeypatch.setattr(m._bus_mod, "TIMEOUT", 0.4)
    monkeypatch.setattr(m._bus_mod, "REQUEST_GAP", 0)
    monkeypatch.setattr(m, "LOG_DIR", str(tmp_path))
    return m


@pytest.fixture
def gw(M, monkeypatch):
    g = FakeGateway().start()
    monkeypatch.setattr(M, "PORT", g.port)
    yield g
    g.stop()


def _snap(gw, **kw):
    s = {"ip": gw.host, "uid": 2, "mode": "pct", "pct": 50, "ref_ip": "", "ref_uid": 0, "first": True,
         "last_target": None, "mode_allowed": False, "mode_tries": 0}
    s.update(kw)
    return s


def test_pct_to_rpm(M):
    assert M.pct_to_rpm(100) == 1100 and M.pct_to_rpm(50) == 550 and M.pct_to_rpm(0) == 0
    assert M.pct_to_rpm(33.3) == 366 and M.pct_to_rpm(150) == 1100


def test_pct_control_asks_mode_then_sets_speed(M, gw):
    gw.add_ecblue(2, analog_rpm=300, h4=0, pin_level=0)
    b = M.Buses()
    res = M.control_job(b, _snap(gw))
    assert res["need_mode"] and not gw.writes                        # 확인 전에는 아무것도 쓰지 않음
    res = M.control_job(b, _snap(gw, mode_allowed=True))
    assert res["ok"] and res["target"] == 550 and res["rpm"] == 550 and gw.ecb_speed(2) == 550
    assert (2, 0, 10) in gw.writes and gw.ecb[2]["h"][4] == 0x0101 and res["mode_tries"] == 1
    # 정기 점검: 정상이면 쓰기 없음, 정보 단계는 기록하지 않음
    n = len(gw.writes)
    res = M.control_job(b, _snap(gw, first=False, last_target=550, mode_allowed=True, mode_tries=1))
    assert res["ok"] and len(gw.writes) == n and res["steps"] == []
    # 모터 쪽 h2 가 풀리면(정전 등) 다시 지정하고 주의로 기록
    gw.ecb[2]["h"][2] = 0
    res = M.control_job(b, _snap(gw, first=False, last_target=550, mode_allowed=True, mode_tries=1))
    assert res["ok"] and gw.ecb_speed(2) == 550 and any(lv == M.WARN for lv, _ in res["steps"])
    b.close()


def test_ref_control_follows_reference(M, gw):
    gw.add_ecblue(2, analog_rpm=0, h4=0x0101)
    gw.add_ecblue(7, analog_rpm=640)                                  # 참조팬 (아날로그로 640 RPM)
    b = M.Buses()
    s = _snap(gw, mode="ref", ref_ip=gw.host, ref_uid=7)
    res = M.control_job(b, s)
    assert res["ok"] and res["ref_state"] == "정상" and res["ref_rpm"] == 640 and gw.ecb_speed(2) == 640
    gw.ecb[7]["analog"] = 820
    res = M.control_job(b, dict(s, first=False, last_target=640))
    assert res["target"] == 820 and gw.ecb_speed(2) == 820
    # 참조팬 응답 없음 -> 에러 기록, 이전 지정 유지
    del gw.ecb[7]
    res = M.control_job(b, dict(s, first=False, last_target=820))
    assert res["ref_state"] == "응답 없음" and res["target"] == 820 and gw.ecb_speed(2) == 820
    assert any(lv == M.ERROR and "이전 지정 820" in m for lv, m in res["steps"])
    # 처음 적용인데 참조팬이 없으면 쓰지 않음
    n = len(gw.writes)
    res = M.control_job(b, dict(s, last_target=None))
    assert not res["ok"] and len(gw.writes) == n
    b.close()


def test_ref_speed_limited_to_max(M, gw):
    gw.add_ecblue(2, analog_rpm=0, h4=0x0101)
    gw.add_ecblue(7, analog_rpm=1300)
    res = M.control_job(M.Buses(), _snap(gw, mode="ref", ref_ip=gw.host, ref_uid=7))
    assert res["target"] == 1100 and gw.ecb_speed(2) == 1100


def test_mode_write_tries_limited(M, gw):
    gw.add_ecblue(2, analog_rpm=0, h4=0)
    gw.ecb[2]["refuse_h4"] = True
    b = M.Buses()
    tries = 0
    for _ in range(5):
        res = M.control_job(b, _snap(gw, mode_allowed=True, mode_tries=tries))
        tries = res["mode_tries"]
    n = sum(1 for r in gw.requests if r[7] == 6 and r[9] == modbus.ECB_H_CONTROLMODE)
    assert n == M.MODE_WRITE_MAX_TRIES and not res["ok"]
    assert any("더 시도하지 않음" in m for lv, m in res["steps"] if lv == M.ERROR)


def test_connect_failure_is_error(M, monkeypatch):
    monkeypatch.setattr(M, "PORT", 1)                                # 닫힌 포트
    res = M.control_job(M.Buses(), {"ip": "127.0.0.1", "uid": 2, "mode": "pct", "pct": 10, "ref_ip": "",
                                    "ref_uid": 0, "first": True, "last_target": None,
                                    "mode_allowed": False, "mode_tries": 0})
    assert not res["ok"] and any(lv == M.ERROR and "접속 실패" in m for lv, m in res["steps"])


def test_items_file_roundtrip(M, tmp_path):
    p = str(tmp_path / "items.json")
    items = [M.Item("172.27.0.116", 180, "pct", 45.5), M.Item("172.27.0.116", 176, "ref", 0, "172.27.0.116", 180)]
    M.save_items(p, items)
    back = M.load_items(p)
    assert [i.to_json() for i in back] == [i.to_json() for i in items]
    assert set(json.load(open(p, encoding="utf-8"))[0]) == {"ip", "uid", "mode", "pct", "ref_ip", "ref_uid"}
    assert M.load_items(str(tmp_path / "none.json")) == []


# ---------------------------------------------------------------- 화면
@pytest.fixture
def app(M, gw, tmp_path, monkeypatch):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("디스플레이 없음")
    root.withdraw()
    monkeypatch.setattr(M, "CHECK_INTERVAL_MS", 300)
    monkeypatch.setattr(M.messagebox, "askyesno", lambda *a, **k: True)
    a = M.MiniControlApp(root, items_file=str(tmp_path / "items.json"))
    yield a, root
    a.shutdown()
    root.destroy()


def pump(root, cond, timeout=6.0):
    end = time.time() + timeout
    while time.time() < end:
        root.update()
        if cond():
            return True
        time.sleep(0.02)
    return False


def _log(a):
    return a.logbox.get("1.0", "end")


def test_ui_flow(M, gw, app):
    a, root = app
    gw.add_ecblue(180, analog_rpm=0, h4=0, pin_level=0)
    gw.add_ecblue(176, analog_rpm=700)
    a.items = [M.Item(gw.host, 180, "ref", 0, gw.host, 176), M.Item(gw.host, 99, "pct", 50)]
    a.refresh()
    a.tree.selection_set(("0", "1"))
    a.on_apply()
    it, bad = a.items
    assert pump(root, lambda: it.state == "정상 제어 중" and bad.level == M.ERROR)
    assert gw.ecb_speed(180) == 700 and it.ref_rpm == 700 and it.target == 700
    # 참조팬 속도 변경 -> 정기 점검으로 따라감
    gw.ecb[176]["analog"] = 500
    assert pump(root, lambda: gw.ecb_speed(180) == 500 and it.target == 500)
    text = _log(a)
    assert "[에러] [%s ID99]" % gw.host in text                    # 없는 모터 -> 빨간 [에러]
    assert "적용 완료" in text and "점검:" not in text                   # 정상 점검 결과는 로그에 없음
    assert it.checked != "-"                                              # 목록에는 마지막 점검 시각 표시
    first = text.splitlines()[0]
    assert first[:4].isdigit() and first[10] == " " and first[13] == ":"   # 맨 앞 시간
    ranges = a.logbox.tag_ranges("error")
    assert ranges and "[에러]" in a.logbox.get(ranges[0], ranges[1])
    vals = a.tree.item("1", "values")
    assert "[에러]" in vals[5] and "error" in a.tree.item("1", "tags")
    # 정지 -> h2 = 0, 점검 멈춤
    a.tree.selection_set("0")
    a.on_stop()
    assert pump(root, lambda: gw.ecb[180]["h"][2] == 0 and not it.active)
    n = len(gw.writes)
    pump(root, lambda: False, timeout=1.0)
    assert len(gw.writes) == n
    assert os.path.exists(os.path.join(M.LOG_DIR, time.strftime("mini_control_%Y%m%d.log")))


def test_ui_recovery_logged_once(M, gw, app):
    a, root = app
    gw.add_ecblue(180, analog_rpm=0, h4=0x0101)
    gw.add_ecblue(176, analog_rpm=600)
    a.items = [M.Item(gw.host, 180, "ref", 0, gw.host, 176)]
    a.refresh()
    a.tree.selection_set("0")
    a.on_apply()
    it = a.items[0]
    assert pump(root, lambda: it.state == "정상 제어 중")
    saved = gw.ecb.pop(176)                                               # 참조팬 응답 없음 -> 에러
    assert pump(root, lambda: it.level == M.ERROR)
    gw.ecb[176] = saved                                                   # 복구 -> '정상으로 돌아옴' 한 번
    assert pump(root, lambda: it.level == M.INFO)
    pump(root, lambda: False, timeout=1.0)                                # 이후 정상 점검 여러 번
    assert _log(a).count("정상으로 돌아옴") == 1 and "점검:" not in _log(a)


@pytest.fixture
def topmost_spy(monkeypatch):
    """창마다 설정된 -topmost 값을 기록. (X11 가상화면에는 창 관리자가 없어 Tk 가 이 값을 되읽지 못함.
    Windows 에서는 Tk 가 직접 처리)"""
    state = {}
    orig = tk.Wm.wm_attributes

    def spy(self, *args, **kw):
        if len(args) == 2 and args[0] == "-topmost":
            state[str(self)] = bool(args[1])
        elif len(args) == 1 and args[0] == "-topmost":
            return int(state.get(str(self), False))
        return orig(self, *args, **kw)
    monkeypatch.setattr(tk.Wm, "wm_attributes", spy)
    monkeypatch.setattr(tk.Wm, "attributes", spy)
    return state


def test_ui_topmost_checkbox(M, topmost_spy, gw, tmp_path, monkeypatch):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("디스플레이 없음")
    root.withdraw()
    items_file = str(tmp_path / "items.json")
    a = M.MiniControlApp(root, items_file=items_file)
    try:
        assert not a.v_topmost.get() and topmost_spy[str(root)] is False      # 기본: 꺼짐
        a.chk_topmost.invoke()                                                 # 체크
        assert topmost_spy[str(root)] is True
        path = M.settings_path(items_file)
        assert M.load_settings(path) == {"topmost": True} and "항상 위에 표시" in _log(a)
        d = M.ItemDialog(root, "항목 추가")                                   # 입력 창도 위에 표시
        assert topmost_spy.get(str(d)) is True
        d.destroy()
        # 다시 실행하면 저장된 설정으로 시작
        root2 = tk.Toplevel(root)
        a2 = M.MiniControlApp(root2, items_file=items_file)
        assert a2.v_topmost.get() and topmost_spy[str(root2)] is True
        a2.shutdown()
        root2.destroy()
        a.chk_topmost.invoke()                                                 # 해제
        assert topmost_spy[str(root)] is False and M.load_settings(path) == {"topmost": False}
        d = M.ItemDialog(root, "항목 추가")
        assert str(d) not in topmost_spy                                       # 해제 상태면 입력 창은 보통
        d.destroy()
    finally:
        a.shutdown()
        root.destroy()
