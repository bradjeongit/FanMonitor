"""운영 프로그램 후속 요청 시험:
1·2 PORT 는 502 고정 (입력칸·표시 없음), 3 기존모터제어 구역 참조 % (test_zone_requests 에도 있음),
4 새 창은 기존 창 안에 뜸 (모니터 기준 아님)."""
import logging

import pytest

from motor_monitor import config as cfgmod
from motor_monitor.constants import MotorType
from motor_monitor.plant import Plant

TEXT_5020 = ("[LINE_CONFIG]\r\nH1R1_SOURCE_IP=172.27.0.84\r\nH1R1_SOURCE_PORT=5020\r\nH1R1_TARGEIP=172.27.0.85\r\n"
             "H1R1_TARGEPORT=503\r\n[MOTOR_CONFIG]\r\nH1R1_MOTORS=7=0=1=0=0\r\n"
             "[CTRL_MOTOR_CONFIG]\r\nH1R1_MOTORS=2@172.27.0.95:5020=1=700\r\n")


# ---------------------------------------------------------------- 1·2 PORT 502 고정
def test_config_port_is_fixed_502(caplog):
    with caplog.at_level(logging.WARNING):
        configs, motors = cfgmod.config_from_text(TEXT_5020)
    c = configs["H1R1"]
    assert (c.source_ip, c.source_port, c.target_ip, c.target_port) == ("172.27.0.84", 502, "172.27.0.85", 502)
    assert "H1R1_SOURCE_PORT=5020" in caplog.text and "H1R1_TARGEPORT=503" in caplog.text   # 다른 값은 기록만
    ctrl = [m for m in motors["H1R1"] if m.mtype == MotorType.OLD_CTRL]
    assert [(m.ip, m.port) for m in ctrl] == [("172.27.0.95", 502)]
    # 저장도 항상 502 (메모리에 다른 값이 있어도)
    configs["H1R1"] = c.copy(source_port=1234)
    text = cfgmod.build_config_text(configs, motors)
    assert "H1R1_SOURCE_PORT=502\r\n" in text and "H1R1_TARGEPORT=502\r\n" in text
    assert "H1R1_MOTORS=2@172.27.0.95:502=1=700" in text and "1234" not in text


def test_plant_load_uses_502(tmp_path):
    path = tmp_path / "Config.ini"
    path.write_text(TEXT_5020, encoding="utf-8")
    p = Plant(str(tmp_path / "Config.dll"))
    p.load()
    imp = p.read_config_import(str(path))
    p.apply_config_import(imp)
    assert (p.configs["H1R1"].source_port, p.configs["H1R1"].target_port) == (502, 502)
    assert p.houses["H1R1"].get_motor("2_2@172.27.0.95:502") is not None
    assert all(":5020" not in line and ":502" not in line for line in imp.summary_lines())


# ---------------------------------------------------------------- 4 창 위치 (계산)
def _dialogs():
    pytest.importorskip("tkinter")
    from motor_monitor.ui import dialogs
    return dialogs


def test_inside_position_rules():
    d = _dialogs()
    pos = d.inside_position
    assert pos((100, 50, 1000, 700), (400, 300)) == (400, 250)                 # 기존 창 가운데
    assert pos((-1920, 0, 1920, 1080), (400, 300)) == (-1160, 390)            # 왼쪽 모니터 (음수 좌표 그대로)
    assert pos((2000, -1080, 1600, 900), (400, 300)) == (2600, -780)          # 위쪽 모니터
    assert pos((100, 50, 300, 200), (400, 300)) == (100, 50)                  # 기존 창보다 크면 왼쪽 위
    # 대화상자 위의 메시지 창: 대화상자 가운데, 프로그램 창 밖으로는 나가지 않음
    assert pos((500, 300, 400, 300), (460, 170), (0, 0, 1600, 900)) == (470, 365)
    assert pos((1300, 700, 400, 300), (460, 170), (0, 0, 1600, 900)) == (1140, 730)
    assert d.geometry_text(400, 300, -1160, 390) == "400x300+-1160+390"


# ---------------------------------------------------------------- 4 창 위치 (실제 창)
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


@pytest.fixture
def plant(tmp_path):
    p = Plant(str(tmp_path / "Config.dll"))
    p.load()
    return p


def _inside(rect, outer):
    x, y, w, h = rect
    ox, oy, ow, oh = outer
    return ox <= x and x + w <= ox + ow and oy <= y and y + h <= oy + oh


def _rect(w):
    w.update_idletasks()
    return w.winfo_rootx(), w.winfo_rooty(), w.winfo_width(), w.winfo_height()


def _messages(parent, dialogs):
    return [w for w in parent.winfo_children() if isinstance(w, dialogs.MessageDialog)]


def test_dialog_opens_inside_main_window(shown_root, plant):
    from motor_monitor.ui import dialogs
    outer = dialogs.owner_rect(shown_root)
    d = dialogs.LineSettingsDialog(shown_root, plant, "H1R1")
    seen = {}

    def check():
        seen["rect"] = _rect(d)
        d.on_cancel()
    shown_root.after(300, check)
    d.show_modal()
    assert seen["rect"][2:] == (400, 300) and _inside(seen["rect"], outer)
    assert seen["rect"][:2] == dialogs.inside_position(outer, (400, 300))       # 가운데


def test_motor_dialog_stays_inside_when_it_grows(shown_root, plant):
    """기존모터제어를 고르면 창이 커짐(340 -> 460): 프로그램 창 아래로 넘치지 않게 위로 올림."""
    from motor_monitor.ui import dialogs
    from motor_monitor.ui import theme as T
    images = T.Images(shown_root)
    outer = dialogs.owner_rect(shown_root)
    d = dialogs.MotorDialog(shown_root, plant, images, "H1R1", "ADD")
    seen = {}

    def check():
        seen["before"] = _rect(d)
        d.motor_type_var.set(int(MotorType.OLD_CTRL)); d._on_type()
        seen["after"] = _rect(d)
        d.on_cancel()
    shown_root.after(300, check)
    d.show_modal()
    assert seen["before"][3] == 340 and seen["after"][3] == 460
    assert _inside(seen["before"], outer) and _inside(seen["after"], outer)


def test_message_window_inside_and_answers(shown_root):
    from motor_monitor.ui import dialogs, msgbox
    outer = dialogs.owner_rect(shown_root)
    seen = []

    def answer(action, *args):
        def f():
            m = _messages(shown_root, dialogs)[-1]
            seen.append((_rect(m), m.kind, [b._text for b in m.buttons]))
            getattr(m, action)(*args)
        return f
    shown_root.after(300, answer("on_default"))                 # Enter = 예
    assert msgbox.ask_yes_no("확인", "진행할까요?", parent=shown_root) is True
    shown_root.after(300, answer("on_cancel"))                  # Esc / 닫기 = 아니오
    assert msgbox.ask_yes_no("확인", "진행할까요?", parent=shown_root) is False
    shown_root.after(300, answer("press", 1))                   # [아니오]
    assert msgbox.ask_yes_no("확인", "진행할까요?", parent=shown_root, icon="warning") is False
    shown_root.after(300, answer("on_default"))
    msgbox.show_error("오류", "아주 긴 안내 문구입니다. " * 12, parent=shown_root)
    assert [(k, b) for _, k, b in seen] == [("question", ["예", "아니오"]), ("question", ["예", "아니오"]),
                                            ("warning", ["예", "아니오"]), ("error", ["확인"])]
    for rect, *_ in seen:
        assert _inside(rect, outer)
    assert seen[3][0][3] > seen[0][0][3]                       # 긴 문구는 창이 높아짐


def test_message_over_dialog_returns_to_dialog(shown_root, plant):
    """대화상자에서 띄운 메시지 창: 대화상자 가운데(프로그램 창 안)에 뜨고, 닫으면 대화상자로 돌아감."""
    from motor_monitor.ui import dialogs
    outer = dialogs.owner_rect(shown_root)
    d = dialogs.LineSettingsDialog(shown_root, plant, "H1R1")
    state = {}

    def close_message():
        m = _messages(d, dialogs)[0]
        state["msg"] = m.message
        state["rect"] = _rect(m)
        state["dlg"] = _rect(d)
        m.on_default()

    def step():
        d.txtSIP.delete(0, "end"); d.txtSIP.insert(0, "300.1.1.1")
        shown_root.after(300, close_message)
        d.on_save()                                            # 실제 메시지 창이 닫힐 때까지 기다림
        state["grab"] = d.grab_current()
        d.on_cancel()
    shown_root.after(300, step)
    d.show_modal()
    assert state["msg"] == "기존모터 접속 IP가 올바른 형식이 아닙니다.."
    assert _inside(state["rect"], outer)
    mx, my, mw, mh = state["rect"]
    dx, dy, dw, dh = state["dlg"]
    assert abs((mx + mw // 2) - (dx + dw // 2)) <= 1 and abs((my + mh // 2) - (dy + dh // 2)) <= 1
    assert state["grab"] is d                                   # 메시지 창을 닫으면 대화상자가 다시 입력을 받음


def test_error_log_window_inside_main_window(shown_root, plant):
    from motor_monitor.ui import dialogs
    from motor_monitor.ui.tools import ErrorLogWindow
    w = ErrorLogWindow(shown_root, plant)
    try:
        rect = _rect(w)
        assert rect[2:] == ErrorLogWindow.SIZE and _inside(rect, dialogs.owner_rect(shown_root))
    finally:
        w.destroy()
