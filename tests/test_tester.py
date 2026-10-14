"""tools/ecblue_tester.py 화면 시험 (시뮬레이터 사용). 디스플레이가 없으면 건너뜀."""
import importlib.util
import os
import datetime
import threading
import time

import pytest

tk = pytest.importorskip("tkinter")

from motor_monitor.simulator import FakeGateway  # noqa: E402


def _load_tool():
    spec = importlib.util.spec_from_file_location(
        "ecblue_tester", os.path.join(os.path.dirname(__file__), "..", "tools", "ecblue_tester.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture
def env(tmp_path, monkeypatch):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("디스플레이 없음")
    root.withdraw()
    T = _load_tool()
    gw = FakeGateway().start()
    gw.add_ecblue(180, analog_rpm=640, pin_level=0)
    gw.add_ecblue(176, analog_rpm=300)
    monkeypatch.setattr(T, "HOST", gw.host)
    monkeypatch.setattr(T, "PORT", gw.port)
    monkeypatch.setattr(T, "LOG_DIR", str(tmp_path))
    monkeypatch.setattr(T, "REQUEST_GAP", 0)
    monkeypatch.setattr(T, "TIMEOUT", 0.5)
    monkeypatch.setattr(T, "MONITOR_INTERVAL_MS", 200)
    answers = []
    monkeypatch.setattr(T.messagebox, "askyesno", lambda *a, **k: (answers.append(a[0]), True)[1])
    app = T.TesterApp(root)
    yield T, app, gw, root, answers, tmp_path
    app.shutdown()
    root.destroy()
    gw.stop()


def pump(root, cond, timeout=5.0):
    end = time.time() + timeout
    while time.time() < end:
        root.update()
        if cond():
            return True
        time.sleep(0.02)
    return False


def test_tester_full_flow(env):
    T, app, gw, root, answers, tmp = env
    p180, p176 = app.panels[180], app.panels[176]
    assert T.MOTOR_IDS == (180, 176) and list(app.panels) == [180, 176]
    # 1) 접속 테스트: 읽기만, 원래 설정 기억
    p180.on_connect_test()
    p176.on_connect_test()
    assert pump(root, lambda: p180.orig is not None and p176.orig is not None)
    assert gw.writes == []
    assert p180.lbl_rpm.cget("text").startswith("640 RPM") and "아날로그" in p180.info.cget("text")
    # 2) 현재 속도
    p176.on_read_speed()
    assert pump(root, lambda: "300 RPM" in p176.lbl_rpm.cget("text") and p176.last_ok)
    # 3) 속도 지정: 제어모드 변경 확인 -> PIN -> h4 -> h1 -> h2, 모니터링 자동 시작
    p180.ent_rpm.delete(0, "end"); p180.ent_rpm.insert(0, "500")
    p180.on_set_speed()
    assert pump(root, lambda: p180.target == 500)
    assert answers[:2] == ["속도 지정", "제어모드 변경"]
    assert gw.ecb_speed(180) == 500 and p180.monitoring
    assert pump(root, lambda: p180.lbl_rpm.cget("text").startswith("500 RPM"))
    pump(root, lambda: False, 0.5)                               # 모니터링 기록(CSV) 몇 번
    # 4) 지정 속도 유지: 모터 쪽 h2 가 바뀌면(정전 등) 모니터링이 다시 지정
    gw.ecb[180]["h"][2] = 0
    assert pump(root, lambda: gw.ecb[180]["h"][2] == 500)
    # 5) 정지: 유지 쓰기(500)가 대기 중일 때 정지를 눌러도 최종 값은 0 (작업은 순서대로 처리)
    p180.set_monitoring(False)
    pump(root, lambda: not p180.busy_monitor and not p180.keep_pending)
    gate = threading.Event()
    app.worker.submit(lambda bus: (gate.wait(3), {})[1])          # 작업 쓰레드를 잠시 막음
    p180.monitoring = True
    p180._done_monitor({"uid": 180, "ok": True, "time": datetime.datetime.now(), "h1": 1, "h2": 0, "h3": 0,
                        "h4": 0x0101})                            # h2 가 바뀐 것을 본 모니터링 -> 유지 쓰기 대기열
    p180.monitoring = False
    assert p180.keep_pending
    p180.on_stop()
    gate.set()
    assert pump(root, lambda: not p180.keep_pending)
    pump(root, lambda: False, 0.3)
    assert gw.ecb_speed(180) == 0 and gw.ecb[180]["h"][2] == 0    # 500 으로 되돌아가지 않음
    # 6) 원래대로 -> 아날로그 640
    p180.on_restore()
    assert pump(root, lambda: gw.ecb_speed(180) == 640 and gw.ecb[180]["h"][4] == 0)
    # 로그 파일과 CSV
    pump(root, lambda: False, 0.5)
    logtxt = open(app.log_path, encoding="utf-8").read()
    for s in ("접속 테스트 결과: 정상", "현재 속도: 300 RPM", "속도 지정 완료: 500 RPM", "지정 속도 유지",
              "정지 완료", "원래대로 완료"):
        assert s in logtxt, s
    csv = open(app.csv_path, encoding="utf-8-sig").read().splitlines()
    assert csv[0].startswith("시각,ID,응답,실제RPM") and any(",180,1,500," in r for r in csv)


def test_tester_no_response(env):
    T, app, gw, root, _answers, _tmp = env
    del gw.ecb[176]                       # 버스에서 응답 없는 모터
    p = app.panels[176]
    p.on_connect_test()
    assert pump(root, lambda: p.last_ok is False)
    assert "응답 없음" in p.lbl_state.cget("text") and p.orig is None
    T.messagebox.showinfo = lambda *a, **k: None
    p.on_restore()                        # 원래 설정을 모르면 안내만
    assert gw.writes == []
