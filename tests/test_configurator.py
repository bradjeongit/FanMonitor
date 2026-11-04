"""tools/ecblue_configurator.py 시험 (시뮬레이터 사용). 디스플레이가 없으면 건너뜀.
IP·ID 선택과 장치 1대 연결, 레지스터 의미 표시, 속도 그래프(보기·확인 주기·24시간 보관), 속도 지정·원래대로,
통신 끊김 시 자동 운전(h17·h5) 확인·설정."""
import importlib.util
import json
import os
import time

import pytest

tk = pytest.importorskip("tkinter")

from motor_monitor import modbus  # noqa: E402
from motor_monitor.simulator import FakeGateway  # noqa: E402


def _load_tool():
    spec = importlib.util.spec_from_file_location(
        "ecblue_configurator", os.path.join(os.path.dirname(__file__), "..", "tools", "ecblue_configurator.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class Env:
    def __init__(self, m, app, gw, root, tmp):
        self.m, self.app, self.gw, self.root, self.tmp = m, app, gw, root, tmp
        self.asked = []          # (제목, 내용)
        self.replies = []        # askyesno 답 (비어 있으면 '예')
        self.errors = []

    def pump(self, cond, timeout=5.0):
        end = time.time() + timeout
        while time.time() < end:
            self.root.update()
            if cond():
                return True
            time.sleep(0.02)
        return False

    def connect(self, uid, host=None):
        app = self.app
        app.cb_ip.set(host or self.gw.host)
        app.cb_id.set(str(uid))
        app.on_connect()
        assert self.pump(lambda: not app.connecting)
        return app.dev is not None

    def row(self, iid):
        return self.app.tree.item(iid, "values")


@pytest.fixture
def env(tmp_path, monkeypatch):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("디스플레이 없음")
    m = _load_tool()
    gw = FakeGateway().start()
    gw.add_ecblue(180, analog_rpm=640, pin_level=0)            # 아날로그 제어(h4=0), PIN 보호 0단계
    gw.add_ecblue(176, analog_rpm=0, h4=0x0101)                 # 통신 속도 제어
    gw.ecb[176]["h"][2] = 700
    gw.ecb[176]["h"][1] = 1
    monkeypatch.setattr(m, "PORT", gw.port)
    monkeypatch.setattr(m, "LOG_DIR", str(tmp_path / "logs"))
    monkeypatch.setattr(m, "TIMEOUT", 0.5)
    monkeypatch.setattr(m.ms, "REQUEST_GAP", 0)
    monkeypatch.setattr(m.ms, "WRITE_VERIFY_GAP", 0)
    e = None

    def ask(title, msg, **_k):
        e.asked.append((title, msg))
        return e.replies.pop(0) if e.replies else True
    monkeypatch.setattr(m.messagebox, "askyesno", ask)
    monkeypatch.setattr(m.messagebox, "showerror", lambda t, msg, **k: e.errors.append(msg))
    monkeypatch.setattr(m.messagebox, "showinfo", lambda t, msg, **k: e.errors.append(msg))
    root.geometry("1120x720+0+0")
    app = m.ConfiguratorApp(root, settings_file=str(tmp_path / "ecblue_configurator.json"))
    e = Env(m, app, gw, root, tmp_path)
    root.update()
    yield e
    app.shutdown()
    root.destroy()
    gw.stop()


# ---------------------------------------------------------------- 1·2·3 장치 1대 연결, 레지스터 의미
def test_connect_reads_registers_with_meanings(env):
    app = env.app
    assert env.connect(180)
    assert app.dev["ecblue"] and "ECblue" in app.lbl_status.cget("text") and "응답 정상" in app.lbl_status.cget("text")
    assert env.row("i14")[1:] == ("640", "640 rpm")                  # 아날로그로 640 rpm
    assert "ECblue Basic Standard" in env.row("i1")[2]
    assert env.row("i0")[2] == "14.30"                               # 펌웨어
    assert env.row("h1")[2] == "0x0000 운전 허가 꺼짐"
    assert env.row("h4")[2].startswith("0x0000 E1 아날로그")
    assert env.row("h17")[2] == "꺼짐"
    assert "writable" in app.tree.item("h2", "tags") and "writable" not in app.tree.item("h3", "tags")
    app.tree.selection_set("h3")
    env.root.update()
    assert app.lbl_detail.cget("text").startswith("h3 통신 설정 (COM) = 63234: 0xF702 주소 247")
    # 자동 운전: 꺼짐이라 입력칸은 기본값, 속도 기록 첫 점
    assert app.lbl_wd.cget("text").startswith("지금: 꺼짐")
    assert (app.ent_wd_sec.get(), app.ent_wd_rpm.get(), app.cb_wd_mode.get()) == ("60", "770", "정속 운전")
    assert list(app.history)[-1][1:] == (640, 0)
    assert all(str(b.state()) == "()" for b in (app.btn_set, app.btn_stop, app.btn_restore, app.btn_wd_set))
    s = json.load(open(env.tmp / "ecblue_configurator.json", encoding="utf-8"))
    assert (s["ip"], s["uid"], s["ids"][0]) == (env.gw.host, 180, 180)
    logs = os.listdir(env.tmp / "logs")
    assert len(logs) == 1 and logs[0].startswith("ecblue_configurator_") and logs[0].endswith(".log")  # CSV 없음


def test_one_device_at_a_time(env):
    app = env.app
    assert env.connect(180)
    app.add_sample(app._now(), 641, 0)
    n = len(app.history)
    assert n >= 2
    env.replies = [False]                                            # 다른 장치로 바꾸기 취소
    app.cb_id.set("176")
    app.on_connect()
    assert "속도 기록" in env.asked[-1][1] and app.dev["uid"] == 180 and len(app.history) == n
    assert env.connect(176)                                          # 이번에는 '예' -> 기록 지움
    assert "지워집니다" in env.asked[-1][1]
    assert app.dev["uid"] == 176 and app.hist_dev == (env.gw.host, 176) and len(app.history) == 1
    assert app.orig == {1: 1, 2: 700, 4: 0x0101}
    asked = len(env.asked)
    assert env.connect(176)                                          # 같은 장치 다시 연결: 기록 유지, 묻지 않음
    assert len(env.asked) == asked and len(app.history) == 2
    app.on_disconnect()
    assert app.dev is None and len(app.history) == 2
    assert str(app.btn_set.state()) == "('disabled',)" and "연결 끊음" in app.lbl_status.cget("text")
    assert not env.connect(99)                                       # 없는 ID: 응답 없음
    assert app.lbl_status.cget("text").startswith("연결 실패") and "응답 없음" in app.lbl_status.cget("text")


def test_non_ecblue_device_is_shown_but_not_controlled(env):
    env.gw.add_ecblue(30, analog_rpm=500, product=0x0200)            # FU Basic Icontrol (ZIEHL-ABEGG, ECblue 아님)
    app = env.app
    assert env.connect(30)
    assert not app.dev["ecblue"] and "ECblue 가 아닙니다" in app.txt.get("1.0", "end")
    assert str(app.btn_set.state()) == "('disabled',)" and str(app.btn_wd_set.state()) == "('disabled',)"
    assert app._sample_job is None and not app.history                 # 속도 확인·기록 안 함


# ---------------------------------------------------------------- 4~8 속도 그래프, 확인 주기, 24시간 보관
def test_check_interval_and_10s_button_only_in_10min_view(env, monkeypatch):
    m, app = env.m, env.app
    assert app.range_var.get() == "1시간" and app.interval_sec() == 30
    assert str(app.btn_fast.state()) == "('disabled',)"              # 10분 보기에서만 누를 수 있음
    app.fast_var.set(True)
    assert app.interval_sec() == 30                                  # 켜져 있어도 10분 보기가 아니면 30초
    app.on_fast()                                                    # 막혀 있어도 눌린 것처럼 와도 무시
    assert not app.fast_var.get() and app.interval_sec() == 30
    app.set_range("10분")
    assert str(app.btn_fast.state()) == "()"
    app.fast_var.set(True)
    app.on_fast()
    assert app.interval_sec() == 10 and "10초 단위 체크 켬" in app.txt.get("1.0", "end")
    assert app.lbl_check.cget("text").startswith("10초마다 확인")
    app.set_range("6시간")                                           # 다른 보기: 30초로 돌아감
    assert not app.fast_var.get() and app.interval_sec() == 30
    assert str(app.btn_fast.state()) == "('disabled',)" and "10초 단위 체크 끔" in app.txt.get("1.0", "end")
    # 실제로 그 주기로 확인하는지 (시험에서는 30초 -> 0.6초, 10초 -> 0.1초)
    monkeypatch.setattr(m, "CHECK_SEC", 0.6)
    monkeypatch.setattr(m, "FAST_CHECK_SEC", 0.1)
    assert env.connect(176)
    n0 = len(app.history)
    env.pump(lambda: False, 1.3)
    slow = len(app.history) - n0
    app.set_range("10분")
    app.fast_var.set(True)
    app.on_fast()
    n1 = len(app.history)
    env.pump(lambda: False, 1.3)
    fast = len(app.history) - n1
    assert 1 <= slow <= 3 and fast >= 5, (slow, fast)
    assert list(app.history)[-1][1:] == (700, 700)


def test_history_kept_up_to_24h_and_deleted_on_close(env):
    app = env.app
    T = [1_800_000_000.0]
    app._now = lambda: T[0]
    for k in range(26):                                              # 1시간마다 26개 (0~25시간)
        app.add_sample(T[0] + k * 3600, 700 + k, 700)
    T[0] += 25 * 3600
    times = [p[0] for p in app.history]
    assert times[0] == 1_800_000_000.0 + 3600                        # 24시간보다 오래된 첫 기록은 지움
    assert len(app.history) == 25
    env.replies = [False]
    app.on_close()                                                   # 닫기 확인에 '기록이 지워짐' 안내
    assert "속도 기록 25개는 지워집니다" in env.asked[-1][1]
    app.shutdown()
    assert not app.history and app.worker.is_alive() is False
    with open(app.log_path, encoding="utf-8") as f:
        assert "속도 기록 25개 지움" in f.read()


def test_graph_views_and_hover(env):
    m, app = env.m, env.app
    T = 1_800_000_000.0
    app._now = lambda: T
    for k in range(24 * 60 * 2, -1, -1):                             # 24시간 동안 30초마다
        t = T - k * 30
        v = None if k in (100, 101) else 700 + (k % 7)
        app.history.append((t, v, None if v is None else 700))
    env.root.update()
    seen = {}
    for key, sec in m.RANGES:
        app.set_range(key)
        env.root.update()
        c = app.canvas
        n = len(app._plot)
        assert abs(n - (sec // 30 + 1)) <= 1, (key, n)                 # 그 보기의 기록만 그림
        assert c.find_withtag("actual") and c.find_withtag("set")
        seen[key] = len(c.find_withtag("xlabel"))
        assert 3 <= seen[key] <= 13, (key, seen[key])
    app.set_range("1시간")
    env.root.update()
    x0, x1 = app._geom[0], app._geom[1]
    app._hover_x = int(x0 + (x1 - x0) * 0.5)
    app._render_hover()
    assert "실제" in app.hover_text and "RPM" in app.hover_text and app.canvas.find_withtag("hover")
    app._on_leave()
    assert not app.canvas.find_withtag("hover")
    # 끊김: 응답 없음(None) 이나 GAP_SEC 보다 긴 간격에서 선이 끊김
    segs = m.ConfiguratorApp.segments([(0, 700, 700), (30, None, None), (60, 710, 700), (300, 720, 720)], 1, False)
    assert [len(s) for s in segs] == [1, 1, 1]
    steps = m.ConfiguratorApp.segments([(0, 700, 700), (30, 710, 720)], 2, True)
    assert steps == [[(0, 700), (30, 700), (30, 720)]]               # 지정 속도는 계단선


# ---------------------------------------------------------------- 속도 지정·유지·원래대로
def test_set_speed_changes_mode_then_restore(env, monkeypatch):
    m, app, gw = env.m, env.app, env.gw
    monkeypatch.setattr(m, "CHECK_SEC", 0.2)
    assert env.connect(180)
    app.ent_rpm.set("500")
    app.on_set_speed()                                               # 확인 -> 제어모드 변경 확인 -> 지정
    assert env.pump(lambda: app.target == 500)
    assert [a[0] for a in env.asked[-2:]] == ["속도 지정", "제어모드 변경"]
    assert gw.ecb[180]["h"][4] == 0x0101 and gw.ecb[180]["h"][2] == 500 and gw.ecb_speed(180) == 500
    assert env.pump(lambda: env.row("h2")[1] == "500" and env.row("i14")[1] == "500")
    assert app.lbl_target.cget("text") == "지정 속도: 500 RPM"
    gw.ecb[180]["h"][2] = 0                                          # 정전 후 재시작처럼 모터 값이 바뀜
    assert env.pump(lambda: gw.ecb[180]["h"][2] == 500)              # 지정 속도 유지로 다시 지정
    assert "지정 속도 유지" in app.txt.get("1.0", "end")
    app.on_restore()                                                 # 처음 연결 때 값(h1=0, h2=0, h4=0)으로
    assert env.pump(lambda: gw.ecb[180]["h"][4] == 0 and gw.ecb[180]["h"][2] == 0 and gw.ecb[180]["h"][1] == 0)
    assert env.pump(lambda: "원래대로 완료" in app.txt.get("1.0", "end"))
    app.ent_rpm.set("1200")
    app.on_set_speed()
    assert "0~1100" in env.errors[-1]


# ---------------------------------------------------------------- 9 통신 끊김 시 자동 운전 (h17 + h5)
def test_auto_run_when_no_control_signal(env):
    app, gw = env.app, env.gw
    assert env.connect(180)                                          # PIN 보호 0단계 -> 쓰기 전에 PIN 해제 필요
    assert app.lbl_wd.cget("text").startswith("지금: 꺼짐")
    app.ent_wd_sec.set("90")
    app.ent_wd_rpm.set("760")
    app.cb_wd_mode.current(1)                                        # 경보 + 정속 운전
    app.on_wd_set()
    assert "90초 동안 없으면 760 RPM" in env.asked[-1][1] and "h5 = 760" in env.asked[-1][1]
    assert env.pump(lambda: gw.ecb[180]["h"][17] == (90 << 8) | 3 and gw.ecb[180]["h"][5] == 760)
    assert env.pump(lambda: app.lbl_wd.cget("text").startswith("지금: 켜짐 - 통신이 90초 동안 없으면 760 RPM"))
    assert "켜짐: 90초 끊기면 경보 + 정속 운전 760 rpm" == env.row("h17")[2]
    assert "PIN 보호 해제" in app.txt.get("1.0", "end")
    # 다시 연결하면 장치에 있는 시작 시간·속도·동작을 읽어 입력칸을 채움
    app.ent_wd_sec.set("5")
    assert env.connect(180)
    assert (app.ent_wd_sec.get(), app.ent_wd_rpm.get(), app.cb_wd_mode.get()) == ("90", "760", "경보 + 정속 운전")
    # 짧은 시간은 경고 (30초마다 통신) - '아니오' 면 쓰지 않음
    app.ent_wd_sec.set("20")
    env.replies = [False]
    app.on_wd_set()
    assert "20초는 짧습니다" in env.asked[-1][1]
    env.pump(lambda: False, 0.3)
    assert gw.ecb[180]["h"][17] == (90 << 8) | 3
    # 모터 최고속도(h8)보다 빠른 속도는 쓰지 않음
    gw.ecb[180]["h"][8] = 600
    app.on_reread()
    assert env.pump(lambda: env.row("h8")[1] == "600")
    app.ent_wd_sec.set("60")
    app.ent_wd_rpm.set("700")
    app.on_wd_set()
    assert "최고속도(h8) 600" in env.errors[-1]
    # 끄기
    app.on_wd_off()
    assert env.pump(lambda: gw.ecb[180]["h"][17] == 0)
    assert env.pump(lambda: app.lbl_wd.cget("text").startswith("지금: 꺼짐 - 통신이 끊겨도"))
    assert gw.ecb[180]["h"][5] == 760                                # 속도(h5)는 남음


def test_auto_run_registers_unreadable_disables_setting(env, monkeypatch):
    app, gw = env.app, env.gw
    orig = gw._handle_ecb

    def no_h17(tid, uid, fc, addr, val, req):                      # h17 을 지원하지 않는 장치 흉내
        if fc == modbus.FC_READ_HOLDING and addr <= 17 < addr + val:
            return orig(tid, uid, fc, 40, val, req)                 # 없는 주소 -> 예외 응답
        return orig(tid, uid, fc, addr, val, req)
    monkeypatch.setattr(gw, "_handle_ecb", no_h17)
    assert env.connect(176)
    assert env.row("h17")[1:] == ("-", "읽지 못함")
    assert app.lbl_wd.cget("text").startswith("지금: 읽지 못함")
    assert str(app.btn_wd_set.state()) == "('disabled',)" and str(app.btn_wd_off.state()) == "('disabled',)"
