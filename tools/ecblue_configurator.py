"""ECblue configurator - ECblue 팬 1대를 골라 레지스터 확인, 속도 그래프, 속도 지정, 통신 끊김 시 자동 운전 설정.

예전 '기존모터 시험기'(ECBLUE_TESTER, 172.27.0.95 의 ID 180·176 고정)를 바꾼 프로그램입니다.
1. IP 와 ID 를 골라 [연결] -> 그 장치의 상태(i)·설정(h) 레지스터를 읽어 값과 의미를 표시. 한 번에 장치 1대
2. 속도 그래프: 30초마다 실제 속도(i14)와 지정 속도(h2)를 확인해 그림. 보기 10분/1시간/6시간/12시간/24시간 (기본 1시간)
   [10초 단위 체크]: 10분 보기에서만 잠시 10초마다 확인 (다른 보기로 바꾸면 30초로 돌아감)
   기록은 메모리에만 최대 24시간 보관하고, 프로그램을 닫거나 다른 장치에 연결하면 지움
3. 속도 지정(h2, 필요하면 제어모드 h4 를 통신 속도로)·정지·원래대로, 지정 속도 유지
4. 통신 끊김 시 자동 운전 (통신 워치독 h17 + Set Intern1 h5): 정한 시간 동안 통신(제어 신호)이 없으면 팬이
   스스로 정한 속도로 운전. 연결하면 지금 시작 시간·속도가 있는지 읽어 보여 주고, 두 값을 함께 설정하거나 끔
레지스터 정의·읽기·쓰기(PIN 보호 해제, 쓴 뒤 확인)는 모터 스캐너(tools/motor_scanner.py)와 같은 코드를 씀.
근거: ZIEHL-ABEGG "ECblue MODBUS Description" L-BAL-E076-GB 1517 Index 013
실행: ECBLUE_CONFIGURATOR.exe 또는 ecblue_configurator.bat 더블클릭, 또는 python tools/ecblue_configurator.py
로그: logs/ecblue_configurator_YYYYMMDD.log (모든 동작). 최근 IP·ID: ecblue_configurator.json (exe/프로그램 폴더)
"""
from __future__ import annotations

import bisect
import collections
import datetime
import ipaddress
import json
import logging
import math
import os
import queue
import sys
import time
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Deque, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from motor_monitor import modbus  # noqa: E402

try:  # 같은 폴더의 모터 스캐너(레지스터 정의·통신) 재사용 (python tools/ecblue_configurator.py 실행 시)
    import motor_scanner as ms  # noqa: E402
except ImportError:  # 패키지/exe 로 묶였을 때
    from tools import motor_scanner as ms  # type: ignore  # noqa: E402

TITLE = "ECblue configurator"
PORT = 502
ID_MIN, ID_MAX = 1, 247
DEFAULT_IP, DEFAULT_ID = "172.27.0.95", 180      # 예전 시험기의 대상
TIMEOUT = 1.0                     # 응답 대기 (초)
MAX_RPM = 1100                    # 속도 지정 범위 (운영 프로그램과 같은 최대값)
MODE_WRITE_MAX_TRIES = 3          # 제어모드 h4(영구저장) 쓰기 시도 한도 (실행 1회당)
CHECK_SEC = 30                    # 속도 확인 주기 (초)
FAST_CHECK_SEC = 10               # [10초 단위 체크] 주기 (10분 보기에서만)
KEEP_SEC = 24 * 3600              # 속도 기록 보관 (메모리, 최대 24시간)
GAP_SEC = 75                      # 기록 사이가 이보다 길면 그래프 선을 끊음 (통신·연결 끊김)
RANGES = (("10분", 600), ("1시간", 3600), ("6시간", 6 * 3600), ("12시간", 12 * 3600), ("24시간", 24 * 3600))
RANGE_SEC = dict(RANGES)
DEFAULT_RANGE = "1시간"
FAST_RANGE = "10분"
TICK_SEC = {"10분": 120, "1시간": 600, "6시간": 3600, "12시간": 2 * 3600, "24시간": 4 * 3600}   # 시간 눈금
REDRAW_MS = 5000                  # 그래프 시간축을 움직이는 주기
WD_DEFAULT_SEC = 60               # 자동 운전 시작 시간 기본값
WD_DEFAULT_RPM = 770              # 자동 운전·속도 지정 기본값 (운영 프로그램 새 팬 기본값 770 RPM = 70%)
WD_RECOMMENDED_SEC = 45           # 운영 프로그램·이 프로그램은 30초마다 통신 -> 이보다 짧으면 통신 중에도 켜질 수 있음
WD_MODE_CHOICES = ((2, "정속 운전"), (3, "경보 + 정속 운전"))    # 통신 끊김 시 Set Intern1(h5) 속도로 운전
RECENT_MAX = 10                   # 기억할 최근 IP·ID 개수
WRITABLE = (1, 2, 4, 5, 17)       # 이 프로그램이 바꾸는 설정 (목록에서 노란 줄)
H1_BITS = ((1, "Set Intern2 선택"), (2, "Set Intern3 선택"), (3, "속도 제한 켬"), (4, "역회전"))
HOLD_ROWS = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 25, 30)
ACTUAL_COLOR, SET_COLOR = "#1f6fb2", "#e67e22"

_BASE = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
         else os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_DIR = os.path.join(_BASE, "logs")
SETTINGS_FILE = os.path.join(_BASE, "ecblue_configurator.json")
FONT = ("맑은 고딕", 9)
SMALL = ("맑은 고딕", 8)

log = logging.getLogger("ecblue_configurator")


# ====================================================================== 레지스터 의미
def _ok(r: Optional[modbus.Response], st: str) -> bool:
    return r is not None and st == "ok"


def _st_text(r: Optional[modbus.Response], st: str) -> str:
    if r is not None and r.exception is not None:
        return f"예외 {r.exception}: {ms.EXCEPTIONS.get(r.exception, '알 수 없음')}"
    return {"timeout": "응답 없음"}.get(st, st)


def h1_text(v: Optional[int]) -> str:
    """h1 운전 제어 비트."""
    if v is None:
        return "-"
    parts = ["운전 허가 켬" if v & modbus.ECB_ENABLE_BIT else "운전 허가 꺼짐"]
    parts += [name for bit, name in H1_BITS if v >> bit & 1]
    return f"0x{v:04X} " + ", ".join(parts)


def hold_name(addr: int) -> str:
    if addr == 17:
        return "통신 워치독 (자동 운전)"
    p = ms.PARAM_BY_ADDR.get(addr)
    return p.name if p else f"h{addr}"


def hold_meaning(addr: int, v: Optional[int], h: Dict[int, Optional[int]]) -> str:
    """설정(h) 값의 의미. 모터 스캐너의 정의(EDIT_PARAMS·VIEW_PARAMS·워치독)를 그대로 씀."""
    if v is None:
        return "읽지 못함"
    if addr == 1:
        return h1_text(v)
    if addr == 17:
        return ms.wd_text(v, h.get(5))
    if addr == 8 and v == 0:
        return "0 = 최고속도 제한 없음"
    p = ms.PARAM_BY_ADDR.get(addr)
    if p is not None and p.unit:
        main = f"{v} {p.unit}"
    else:
        main = ms.view_text(addr, v)
    note = p.note if p is not None else ""
    return f"{main} - {note}" if note else main


def wd_summary(h17: Optional[int], h5: Optional[int]) -> str:
    """통신 끊김 시 자동 운전의 지금 상태 (시작 시간·속도가 있는지)."""
    if h17 is None:
        return "읽지 못함 - 이 장치에서 통신 워치독(h17)을 읽을 수 없어 설정할 수 없습니다"
    sec, mode = ms.wd_decode(h17)
    if sec == 0 or mode == 0:
        saved = "" if h5 is None else f" (저장된 정속 속도 h5 = {h5} RPM)"
        return "꺼짐 - 통신이 끊겨도 자동 운전하지 않음" + saved
    if mode in (2, 3):
        speed = "속도 모름 (h5 읽지 못함)" if h5 is None else f"{h5} RPM"
        return f"켜짐 - 통신이 {sec}초 동안 없으면 {speed} 으로 자동 운전 ({ms.WD_MODES[mode]})"
    return ms.wd_text(h17, h5)


def remember(values: list, value) -> list:
    """최근 기록 맨 앞에 넣기 (중복 없이 RECENT_MAX 개)."""
    return ([value] + [v for v in values if v != value])[:RECENT_MAX]


def load_settings(path: str) -> dict:
    s = {"ip": DEFAULT_IP, "uid": DEFAULT_ID, "ips": [DEFAULT_IP], "ids": [DEFAULT_ID]}
    try:
        with open(path, encoding="utf-8") as f:
            d = json.load(f)
        if isinstance(d, dict):
            if isinstance(d.get("ip"), str):
                s["ip"] = d["ip"]
            if isinstance(d.get("uid"), int):
                s["uid"] = d["uid"]
            if isinstance(d.get("ips"), list):
                s["ips"] = [x for x in d["ips"] if isinstance(x, str)][:RECENT_MAX] or s["ips"]
            if isinstance(d.get("ids"), list):
                s["ids"] = [x for x in d["ids"] if isinstance(x, int)][:RECENT_MAX] or s["ids"]
    except (OSError, ValueError):
        pass
    return s


def save_settings(s: dict, path: str) -> None:
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=1)
    except OSError as e:
        log.warning("설정 저장 실패: %s", e)


def utc_offset(t: float) -> float:
    off = datetime.datetime.fromtimestamp(t).astimezone().utcoffset()
    return off.total_seconds() if off else 0.0


def dur_text(sec: float) -> str:
    sec = int(sec)
    if sec < 60:
        return f"{sec}초"
    h, m = divmod(sec // 60, 60)
    return f"{h}시간 {m}분" if h else f"{m}분"


# ====================================================================== 통신 작업 (작업 쓰레드에서 실행)
def _write1(c: "ms.Client", uid: int, reg: int, value: int, steps: List[Tuple[str, str]]) -> bool:
    r, st = c.request(modbus.build_new_write_single, uid, reg, value)
    ok = _ok(r, st)
    steps.append(("info" if ok else "error", f"h{reg} = {value} 쓰기 " + ("완료" if ok else f"실패 ({_st_text(r, st)})")))
    return ok


def job_connect(w: "ms.Worker", host: str, port: int, timeout: float, uid: int) -> dict:
    """접속 -> 제품코드(i1)로 종류 확인 -> 전체 레지스터 읽기."""
    res = {"host": host, "port": port, "uid": uid}
    c = w.use(host, port, timeout)
    try:
        c.connect()
    except OSError as e:
        res["error"] = f"{host}:{port} 접속 실패 ({e})"
        return res
    r, st = c.request(modbus.build_read_input, uid, 1, 1)
    if not (_ok(r, st) and r.values):
        res["error"] = f"{host} ID {uid} 응답 없음 ({_st_text(r, st)})"
        return res
    kind, name = ms.classify(r.values[0])
    data = ms.read_all(c, uid)
    res.update(kind=kind, name=name, product=r.values[0], ecblue=ms.is_ecblue(kind), i=data["i"], h=data["h"])
    return res


def job_sample(w: "ms.Worker", host: str, port: int, timeout: float, uid: int) -> dict:
    """속도 확인: i12~i15 (오류·실제 속도·전류), h1~h4 (운전 허가·지정 속도·제어모드)."""
    c = w.use(host, port, timeout)
    return {"host": host, "uid": uid, "i": ms.read_regs(c, "i", [(12, 4)], uid),
            "h": ms.read_regs(c, "h", [(1, 4)], uid)}


def job_read_all(w: "ms.Worker", host: str, port: int, timeout: float, uid: int) -> dict:
    data = ms.read_all(w.use(host, port, timeout), uid)
    return {"host": host, "uid": uid, "i": data["i"], "h": data["h"]}


def job_set_speed(w: "ms.Worker", host: str, port: int, timeout: float, uid: int, rpm: int, allow_mode: bool,
                  mode_tries: int) -> dict:
    """속도 지정: (필요하면 PIN -> h4 통신 속도 제어) h1 운전 허가, h2 지정 속도 -> 다시 읽어 확인."""
    c = w.use(host, port, timeout)
    steps: List[Tuple[str, str]] = []
    res = {"host": host, "uid": uid, "steps": steps, "rpm": rpm, "mode_written": False, "ok": False}
    h = ms.read_regs(c, "h", [(1, 4)], uid)
    if h.get(1) is None or h.get(4) is None:
        steps.append(("error", "현재 설정(h1~h4)을 읽지 못함 - 쓰지 않음"))
        return res
    if not modbus.ecb_mode_ok(h[4]):
        if not allow_mode:
            res.update(need_mode=True, h4=h[4])
            return res
        if mode_tries >= MODE_WRITE_MAX_TRIES:
            steps.append(("error", f"제어모드(h4) 쓰기 {MODE_WRITE_MAX_TRIES}번 실패 - 더 시도하지 않음 (PIN 보호/통신 확인)"))
            return res
        res["mode_written"] = True
        _write1(c, uid, modbus.ECB_H_PIN, modbus.ECB_PIN_LEVEL2, steps)
        if not _write1(c, uid, modbus.ECB_H_CONTROLMODE, modbus.ecb_mode_value(h[4]), steps):
            return res
    if not h[1] & modbus.ECB_ENABLE_BIT:
        _write1(c, uid, modbus.ECB_H_CONTROL, h[1] | modbus.ECB_ENABLE_BIT, steps)
    _write1(c, uid, modbus.ECB_H_SPEED, rpm, steps)
    data = ms.read_all(c, uid)
    hh = data["h"]
    res["data"] = data
    res["ok"] = (hh.get(2) == rpm and modbus.ecb_mode_ok(hh.get(4) or 0)
                 and bool((hh.get(1) or 0) & modbus.ECB_ENABLE_BIT))
    return res


def job_write_speed(w: "ms.Worker", host: str, port: int, timeout: float, uid: int, rpm: int,
                    still_valid: Callable[[], bool] = lambda: True) -> dict:
    """h2 만 쓰기 (정지, 지정 속도 유지). still_valid 가 False 면(그 사이 다른 지정이 있었음) 쓰지 않음."""
    if not still_valid():
        return {"host": host, "uid": uid, "steps": [], "ok": True, "skipped": True}
    steps: List[Tuple[str, str]] = []
    ok = _write1(w.use(host, port, timeout), uid, modbus.ECB_H_SPEED, rpm, steps)
    return {"host": host, "uid": uid, "steps": steps, "ok": ok}


def job_restore(w: "ms.Worker", host: str, port: int, timeout: float, uid: int, orig: Dict[int, int]) -> dict:
    """처음 연결 때 읽은 h1/h2/h4 로 되돌림: h2 -> h1 -> (PIN -> h4)."""
    c = w.use(host, port, timeout)
    steps: List[Tuple[str, str]] = []
    cur = ms.read_regs(c, "h", [(1, 4)], uid)
    if cur.get(4) is None:
        steps.append(("error", "현재 설정(h1~h4)을 읽지 못함 - 쓰지 않음"))
        return {"host": host, "uid": uid, "steps": steps, "ok": False}
    for reg in (modbus.ECB_H_SPEED, modbus.ECB_H_CONTROL):
        if cur.get(reg) != orig[reg]:
            _write1(c, uid, reg, orig[reg], steps)
    if cur[4] != orig[4]:
        _write1(c, uid, modbus.ECB_H_PIN, modbus.ECB_PIN_LEVEL2, steps)
        _write1(c, uid, modbus.ECB_H_CONTROLMODE, orig[4], steps)
    data = ms.read_all(c, uid)
    if not steps:
        steps.append(("info", "바뀐 값 없음"))
    ok = all(data["h"].get(k) == orig[k] for k in (1, 2, 4))
    return {"host": host, "uid": uid, "steps": steps, "ok": ok, "data": data}


def job_write(w: "ms.Worker", host: str, port: int, timeout: float, uid: int, changes: List[Tuple[int, int]]) -> dict:
    """[(주소, 값)] 쓰기 (모터 스캐너와 같은 방법: 필요하면 PIN 해제, 쓴 뒤 다시 읽어 확인)."""
    c = w.use(host, port, timeout)
    steps = ms.write_regs(c, uid, changes)
    data = ms.read_all(c, uid)
    return {"host": host, "uid": uid, "steps": steps, "data": data,
            "ok": all(lvl == "info" for lvl, _msg in steps)}


def job_close(w: "ms.Worker") -> dict:
    if w.client is not None:
        w.client.close()
    return {}


# ====================================================================== 화면
class ConfiguratorApp:
    _now = staticmethod(time.time)        # 시험에서 바꿀 수 있는 시계

    def __init__(self, root: tk.Tk, settings_file: str = "") -> None:
        self.root = root
        self.settings_file = settings_file or SETTINGS_FILE
        self.settings = load_settings(self.settings_file)
        os.makedirs(LOG_DIR, exist_ok=True)
        self.log_path = os.path.join(LOG_DIR, f"ecblue_configurator_{datetime.date.today():%Y%m%d}.log")
        self._fh = logging.FileHandler(self.log_path, encoding="utf-8")
        self._fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        log.addHandler(self._fh)
        log.setLevel(logging.INFO)

        self.worker = ms.Worker()
        self.worker.start()
        self.dev: Optional[dict] = None              # 연결된 장치 (host, port, uid, kind, name, product, ecblue)
        self.data: Dict[str, Dict[int, Optional[int]]] = {"i": {}, "h": {}}   # 마지막으로 읽은 레지스터
        self.orig: Optional[Dict[int, int]] = None   # 이 장치에 처음 연결했을 때 h1/h2/h4 ([원래대로] 기준)
        # 속도 기록 (시각, 실제 속도 i14, 지정 속도 h2). 메모리에만, 최대 24시간
        self.history: Deque[Tuple[float, Optional[int], Optional[int]]] = collections.deque()
        self.hist_dev: Optional[Tuple[str, int]] = None   # 기록이 속한 장치 (IP, ID)
        self.target: Optional[int] = None            # 이 프로그램이 지정한 속도 (유지 대상)
        self.gen = 0                                 # 지정·정지·원래대로마다 증가 -> 대기 중인 유지 쓰기 무효화
        self.keep_pending = False
        self.mode_tries = 0
        self.busy_sample = False
        self.connecting = False
        self.last_ok: Optional[bool] = None
        self.hover_text = ""                         # 그래프에 마우스를 올렸을 때 보이는 글 (시험용)
        self._alive = True
        self._sample_job = None
        self._redraw_job = None
        self._hover_x: Optional[int] = None
        self._plot: List[Tuple[float, Optional[int], Optional[int]]] = []    # 마지막으로 그린 기록
        self._geom: Optional[tuple] = None           # 마지막 그래프 좌표 (x0, x1, y0, y1, t0, span, ymax)
        self.range_var = tk.StringVar(value=DEFAULT_RANGE)
        self.fast_var = tk.BooleanVar(value=False)
        self.keep_var = tk.BooleanVar(value=True)
        self._build()
        self._update_controls()
        self._update_check_label()
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.log("info", f"{TITLE} 시작")
        root.after(100, self._poll)
        self._redraw_job = root.after(REDRAW_MS, self._redraw_tick)

    # ------------------------------------------------------------------ 화면 구성
    def _build(self) -> None:
        r = self.root
        r.title(TITLE)
        r.geometry("1120x720")
        r.minsize(900, 600)
        style = ttk.Style(r)
        style.configure("Treeview", rowheight=20)
        r.columnconfigure(0, weight=1)
        r.rowconfigure(1, weight=1)

        top = ttk.Frame(r, padding=(8, 6))
        top.grid(row=0, column=0, sticky="ew")
        ttk.Label(top, text="IP", font=FONT).pack(side="left")
        self.cb_ip = ttk.Combobox(top, width=15, values=self.settings["ips"], font=FONT)
        self.cb_ip.set(self.settings["ip"])
        self.cb_ip.pack(side="left", padx=(4, 10))
        ttk.Label(top, text="ID", font=FONT).pack(side="left")
        self.cb_id = ttk.Combobox(top, width=5, values=[str(x) for x in self.settings["ids"]], font=FONT)
        self.cb_id.set(str(self.settings["uid"]))
        self.cb_id.pack(side="left", padx=(4, 10))
        self.btn_conn = ttk.Button(top, text="연결", command=self.on_connect)
        self.btn_conn.pack(side="left")
        self.btn_disc = ttk.Button(top, text="연결 끊기", command=self.on_disconnect)
        self.btn_disc.pack(side="left", padx=(4, 0))
        self.dot = tk.Label(top, text="●", fg="#808080", font=("TkDefaultFont", 12))
        self.dot.pack(side="left", padx=(12, 2))
        self.lbl_status = ttk.Label(top, text="연결 안 됨 - IP 와 ID 를 고르고 [연결] 을 누르세요", font=FONT)
        self.lbl_status.pack(side="left", fill="x", expand=True)
        self.cb_ip.bind("<Return>", lambda e: self.on_connect())
        self.cb_id.bind("<Return>", lambda e: self.on_connect())

        self._pw = pw = ttk.PanedWindow(r, orient="horizontal")
        pw.grid(row=1, column=0, sticky="nsew", padx=8)
        left = ttk.LabelFrame(pw, text=" 레지스터 (값과 의미) ", padding=4)
        right = ttk.Frame(pw)
        pw.add(left, weight=1)
        pw.add(right, weight=1)
        self._sash_done = False
        pw.bind("<Configure>", self._place_sash)
        self._build_registers(left)
        self._build_graph(right)
        self._build_speed(right)
        self._build_watchdog(right)

        lf = ttk.LabelFrame(r, text=f" 로그 (logs 폴더 {os.path.basename(self.log_path)}) ", padding=2)
        lf.grid(row=2, column=0, sticky="nsew", padx=8, pady=(4, 8))
        self.txt = tk.Text(lf, height=5, font=("TkFixedFont", 8), wrap="none", state="disabled")
        sb = ttk.Scrollbar(lf, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt.pack(fill="both", expand=True)
        self.txt.tag_configure("warn", foreground="#d35400")
        self.txt.tag_configure("error", foreground="#c0392b")

    def _build_registers(self, f: ttk.Frame) -> None:
        bar = ttk.Frame(f)
        bar.pack(fill="x")
        self.btn_read = ttk.Button(bar, text="새로 읽기", command=self.on_reread)
        self.btn_read.pack(side="left")
        self.lbl_read = ttk.Label(bar, text="", font=FONT)
        self.lbl_read.pack(side="left", padx=6)
        tk.Label(bar, text=" 노란 줄 = 이 프로그램이 바꾸는 값 ", bg="#fff3c4", font=SMALL).pack(side="right")
        self.tree = ttk.Treeview(f, columns=("name", "value", "meaning"), show="tree headings", height=20)
        for col, text, width, stretch in (("#0", "주소", 64, False), ("name", "이름", 138, False),
                                          ("value", "값", 56, False), ("meaning", "의미", 240, True)):
            self.tree.heading(col, text=text, anchor="w")
            self.tree.column(col, width=width, stretch=stretch, anchor="e" if col == "value" else "w")
        self.lbl_detail = tk.Label(f, text="줄을 누르면 전체 의미가 여기에 보입니다", fg="#555555", font=SMALL,
                                   anchor="w", justify="left")
        self.lbl_detail.pack(side="bottom", fill="x", pady=(4, 0))
        f.bind("<Configure>", lambda e: self.lbl_detail.configure(wraplength=max(e.width - 12, 200)))
        sb = ttk.Scrollbar(f, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True, pady=(4, 0))
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._show_detail())
        self.tree.tag_configure("writable", background="#fff3c4")
        self.tree.tag_configure("missing", foreground="#999999")
        self.tree.tag_configure("alarm", foreground="#c0392b")
        self.tree.tag_configure("group", font=("맑은 고딕", 9, "bold"))
        self.tree.insert("", "end", iid="grp_i", text="상태", values=("입력 레지스터 i (읽기만)", "", ""),
                         open=True, tags=("group",))
        for addr, name, _fmt in ms.INPUT_ROWS:
            self.tree.insert("grp_i", "end", iid=f"i{addr}", text=f"i{addr}", values=(name, "", ""))
        self.tree.insert("", "end", iid="grp_h", text="설정", values=("홀딩 레지스터 h", "", ""),
                         open=True, tags=("group",))
        for addr in HOLD_ROWS:
            self.tree.insert("grp_h", "end", iid=f"h{addr}", text=f"h{addr}", values=(hold_name(addr), "", ""),
                             tags=("writable",) if addr in WRITABLE else ())

    def _build_graph(self, parent: ttk.Frame) -> None:
        f = ttk.LabelFrame(parent, text=" 속도 그래프 (RPM) ", padding=4)
        f.pack(fill="both", expand=True)
        bar = ttk.Frame(f)
        bar.pack(fill="x")
        ttk.Label(bar, text="보기", font=FONT).pack(side="left", padx=(0, 4))
        self.range_buttons: Dict[str, ttk.Radiobutton] = {}
        for key, _sec in RANGES:
            b = ttk.Radiobutton(bar, text=key, value=key, variable=self.range_var, style="Toolbutton",
                                command=lambda k=key: self.set_range(k))
            b.pack(side="left")
            self.range_buttons[key] = b
        self.btn_fast = ttk.Checkbutton(bar, text="10초 단위 체크", variable=self.fast_var, style="Toolbutton",
                                        command=self.on_fast)
        self.btn_fast.pack(side="left", padx=(12, 0))
        self.lbl_check = ttk.Label(f, text="", font=SMALL, foreground="#555555")
        self.lbl_check.pack(fill="x", pady=(2, 0))
        self.canvas = tk.Canvas(f, bg="white", height=200, highlightthickness=1, highlightbackground="#d0d0d0")
        self.canvas.pack(fill="both", expand=True, pady=(4, 2))
        self.canvas.bind("<Configure>", lambda e: self.draw_graph())
        self.canvas.bind("<Motion>", self._on_hover)
        self.canvas.bind("<Leave>", self._on_leave)
        self.lbl_now = tk.Label(f, text="기록 없음", font=("맑은 고딕", 10, "bold"), anchor="w")
        self.lbl_now.pack(fill="x")

    def _build_speed(self, parent: ttk.Frame) -> None:
        f = ttk.LabelFrame(parent, text=" 속도 지정 (h2) ", padding=4)
        f.pack(fill="x", pady=(6, 0))
        row = ttk.Frame(f)
        row.pack(fill="x")
        ttk.Label(row, text="속도", font=FONT).pack(side="left")
        self.ent_rpm = ttk.Spinbox(row, from_=0, to=MAX_RPM, increment=10, width=6, font=FONT)
        self.ent_rpm.set(str(WD_DEFAULT_RPM))
        self.ent_rpm.pack(side="left", padx=4)
        ttk.Label(row, text=f"RPM (0~{MAX_RPM})", font=FONT).pack(side="left")
        self.btn_set = ttk.Button(row, text="지정", command=self.on_set_speed)
        self.btn_stop = ttk.Button(row, text="정지(0)", command=self.on_stop)
        self.btn_restore = ttk.Button(row, text="원래대로", command=self.on_restore)
        for b in (self.btn_set, self.btn_stop, self.btn_restore):
            b.pack(side="left", padx=(6, 0))
        row2 = ttk.Frame(f)
        row2.pack(fill="x", pady=(4, 0))
        ttk.Checkbutton(row2, text="지정 속도 유지 (모터 값이 바뀌어 있으면 다시 지정)",
                        variable=self.keep_var).pack(side="left")
        self.lbl_target = ttk.Label(row2, text="지정 속도: 없음", font=FONT)
        self.lbl_target.pack(side="right")

    def _build_watchdog(self, parent: ttk.Frame) -> None:
        f = ttk.LabelFrame(parent, text=" 통신 끊김 시 자동 운전 - 제어 신호가 없으면 정한 속도로 (h17 + h5) ",
                           padding=4)
        f.pack(fill="x", pady=(6, 0))
        self.lbl_wd = tk.Label(f, text="지금: -", font=("맑은 고딕", 9, "bold"), anchor="w", justify="left")
        self.lbl_wd.pack(fill="x")
        row = ttk.Frame(f)
        row.pack(fill="x", pady=(4, 0))
        ttk.Label(row, text="시작 시간", font=FONT).pack(side="left")
        self.ent_wd_sec = ttk.Spinbox(row, from_=1, to=255, increment=5, width=5, font=FONT)
        self.ent_wd_sec.set(str(WD_DEFAULT_SEC))
        self.ent_wd_sec.pack(side="left", padx=4)
        ttk.Label(row, text="초 (1~255)", font=FONT).pack(side="left")
        ttk.Label(row, text="속도", font=FONT).pack(side="left", padx=(12, 0))
        self.ent_wd_rpm = ttk.Spinbox(row, from_=0, to=MAX_RPM, increment=10, width=6, font=FONT)
        self.ent_wd_rpm.set(str(WD_DEFAULT_RPM))
        self.ent_wd_rpm.pack(side="left", padx=4)
        ttk.Label(row, text="RPM", font=FONT).pack(side="left")
        row = ttk.Frame(f)
        row.pack(fill="x", pady=(4, 0))
        ttk.Label(row, text="동작", font=FONT).pack(side="left")
        self.cb_wd_mode = ttk.Combobox(row, state="readonly", width=14, values=[t for _m, t in WD_MODE_CHOICES],
                                       font=FONT)
        self.cb_wd_mode.current(0)
        self.cb_wd_mode.pack(side="left", padx=(4, 0))
        self.btn_wd_set = ttk.Button(row, text="설정", command=self.on_wd_set)
        self.btn_wd_off = ttk.Button(row, text="끄기", command=self.on_wd_off)
        self.btn_wd_set.pack(side="left", padx=(8, 0))
        self.btn_wd_off.pack(side="left", padx=(4, 0))
        note = tk.Label(f, fg="#666666", font=SMALL, anchor="w", justify="left",
                        text=("이 팬에 Modbus 통신이 '시작 시간' 동안 한 번도 없으면 팬이 스스로 '속도'(Set Intern1, h5)로 "
                              f"돕니다. 운영 프로그램과 이 프로그램은 {CHECK_SEC}초마다 통신하므로 시작 시간은 60초 이상을 "
                              "권장합니다. 두 값은 전원을 꺼도 유지됩니다(영구저장)."))
        note.pack(fill="x", pady=(4, 0))
        f.bind("<Configure>", lambda e: note.configure(wraplength=max(e.width - 16, 200)))

    def _place_sash(self, e) -> None:
        """처음 한 번: 레지스터 표와 그래프를 반씩 (표의 '의미' 열이 보이도록)."""
        if not self._sash_done and e.width > 300:
            self._sash_done = True
            self._pw.sashpos(0, int(e.width * 0.48))

    # ------------------------------------------------------------------ 공통
    def log(self, level: str, msg: str) -> None:
        {"info": log.info, "warn": log.warning, "error": log.error}.get(level, log.info)(msg)
        self.txt.configure(state="normal")
        self.txt.insert("end", f"{datetime.datetime.now():%H:%M:%S} {msg}\n", (level,) if level != "info" else ())
        lines = int(self.txt.index("end-1c").split(".")[0])
        if lines > 2000:
            self.txt.delete("1.0", f"{lines - 2000}.0")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _log_steps(self, steps: List[Tuple[str, str]]) -> None:
        for lvl, msg in steps:
            self.log(lvl, f"  {msg}")

    def _poll(self) -> None:
        try:
            while True:
                done, res = self.worker.results.get_nowait()
                if done is not None:
                    try:
                        done(res)
                    except Exception:
                        log.exception("결과 처리 오류")
        except queue.Empty:
            pass
        if self._alive:
            self.root.after(100, self._poll)

    def _stale(self, res: dict) -> bool:
        """그 사이 연결이 끊겼거나 다른 장치로 바뀌어 화면에 쓰지 않을 결과인지."""
        d = self.dev
        return d is None or res.get("uid") != d["uid"] or res.get("host") != d["host"]

    def _int_entry(self, w, lo: int, hi: int, what: str) -> Optional[int]:
        try:
            v = int(str(w.get()).strip())
        except ValueError:
            v = None
        if v is None or not lo <= v <= hi:
            messagebox.showerror(TITLE, f"{what} {lo}~{hi} 사이 숫자로 입력하세요.", parent=self.root)
            return None
        return v

    def interval_sec(self) -> int:
        """지금 속도 확인 주기: 10분 보기에서 [10초 단위 체크]를 켰을 때만 10초, 아니면 30초."""
        return FAST_CHECK_SEC if self.fast_var.get() and self.range_var.get() == FAST_RANGE else CHECK_SEC

    def _update_controls(self) -> None:
        connected = self.dev is not None and not self.connecting
        ecb = connected and bool(self.dev.get("ecblue"))
        h = self.data.get("h", {})

        def st(on: bool) -> List[str]:
            return ["!disabled"] if on else ["disabled"]
        self.btn_conn.state(st(not self.connecting))
        self.btn_disc.state(st(connected))
        self.btn_read.state(st(connected))
        self.btn_set.state(st(ecb))
        self.btn_stop.state(st(ecb))
        self.btn_restore.state(st(ecb and self.orig is not None))
        wd_ok = ecb and h.get(17) is not None and h.get(5) is not None
        self.btn_wd_set.state(st(wd_ok))
        self.btn_wd_off.state(st(wd_ok))
        self.btn_fast.state(st(self.range_var.get() == FAST_RANGE))

    def _update_status(self) -> None:
        d = self.dev
        if d is None:
            return
        fw = ms.fw_text(self.data["i"].get(0))
        state = "응답 정상" if self.last_ok else "응답 없음"
        self.lbl_status.configure(text=f"{d['host']} ID {d['uid']} · {d['name']} · 펌웨어 {fw} · {state} "
                                       f"({datetime.datetime.now():%H:%M:%S})")

    def _set_ok(self, ok: bool) -> None:
        if ok != self.last_ok:
            if self.last_ok is not None:
                self.log("info" if ok else "warn", "응답이 다시 정상" if ok else "응답 없음 - 통신이 끊겼습니다")
            self.last_ok = ok
        self.dot.configure(fg="#27ae60" if ok else "#c0392b")
        self._update_status()

    def _show_registers(self) -> None:
        i, h = self.data.get("i", {}), self.data.get("h", {})
        for addr, name, fmt in ms.INPUT_ROWS:
            v = i.get(addr)
            tags = ("missing",) if v is None else (("alarm",) if addr == 12 and v else ())
            self.tree.item(f"i{addr}", values=(name, "-" if v is None else v, "읽지 못함" if v is None else fmt(v)),
                           tags=tags)
        for addr in HOLD_ROWS:
            v = h.get(addr)
            tags = (("writable",) if addr in WRITABLE else ()) + (("missing",) if v is None else ())
            self.tree.item(f"h{addr}", values=(hold_name(addr), "-" if v is None else v, hold_meaning(addr, v, h)),
                           tags=tags)
        self._show_detail()

    def _show_detail(self) -> None:
        """표에서 고른 줄의 전체 의미 (칸이 좁아 잘린 글도 보이게)."""
        sel = self.tree.selection()
        if not sel or sel[0].startswith("grp_"):
            return
        name, value, meaning = self.tree.item(sel[0], "values")
        self.lbl_detail.configure(text=f"{sel[0]} {name} = {value}: {meaning}")

    def _show_wd(self, fill: bool) -> None:
        """자동 운전 지금 상태. fill: 입력칸을 장치의 지금 값(시작 시간·속도·동작)으로 채움."""
        h = self.data.get("h", {})
        h17, h5 = h.get(17), h.get(5)
        text = wd_summary(h17, h5)
        on = h17 is not None and all(ms.wd_decode(h17))
        self.lbl_wd.configure(text="지금: " + text, fg="#1e8449" if on else ("#c0392b" if h17 is None else "#555555"))
        if fill and h17 is not None:
            sec, mode = ms.wd_decode(h17)
            self.ent_wd_sec.set(str(sec if on else WD_DEFAULT_SEC))
            self.ent_wd_rpm.set(str(h5 if h5 else WD_DEFAULT_RPM))
            idx = next((k for k, (m, _t) in enumerate(WD_MODE_CHOICES) if m == mode), 0)
            self.cb_wd_mode.current(idx)

    def _apply_data(self, data: dict, sample: bool = True) -> None:
        """새로 읽은 전체 레지스터를 화면에 (설정을 바꾼 뒤 등). sample: 속도 기록에도 한 점 추가."""
        self.data = {"i": dict(data["i"]), "h": dict(data["h"])}
        ok = any(v is not None for v in data["i"].values())
        self._set_ok(ok)
        self._show_registers()
        self._show_wd(fill=False)
        self.lbl_read.configure(text=f"읽은 시각 {datetime.datetime.now():%H:%M:%S}")
        if sample and self.dev is not None and self.dev.get("ecblue"):
            self.add_sample(self._now(), data["i"].get(14), data["h"].get(2))
            self.draw_graph()
        self._update_now()
        self._update_controls()

    # ------------------------------------------------------------------ 연결
    def on_connect(self) -> None:
        ip = self.cb_ip.get().strip()
        try:
            ipaddress.IPv4Address(ip)
        except ValueError:
            messagebox.showerror(TITLE, "IP 형식이 올바르지 않습니다. 예) 172.27.0.95", parent=self.root)
            return
        try:
            uid = int(self.cb_id.get().strip())
        except ValueError:
            uid = -1
        if not ID_MIN <= uid <= ID_MAX:
            messagebox.showerror(TITLE, f"ID 는 {ID_MIN}~{ID_MAX} 사이 숫자로 입력하세요.", parent=self.root)
            return
        if self.hist_dev is not None and self.hist_dev != (ip, uid) and self.history:
            old = f"{self.hist_dev[0]} ID {self.hist_dev[1]}"
            if not messagebox.askyesno(TITLE, f"다른 장치({ip} ID {uid})로 바꾸면 지금 장치({old})의 속도 기록 "
                                              f"{len(self.history)}개가 지워집니다.\n계속할까요?", parent=self.root):
                return
        self._stop_sampling()
        self.dev = None
        self.connecting = True
        self.dot.configure(fg="#f39c12")
        self.lbl_status.configure(text=f"{ip} ID {uid} 연결 중...")
        self._update_controls()
        s = self.settings
        s["ip"], s["uid"] = ip, uid
        s["ips"], s["ids"] = remember(s["ips"], ip), remember(s["ids"], uid)
        save_settings(s, self.settings_file)
        self.cb_ip.configure(values=s["ips"])
        self.cb_id.configure(values=[str(x) for x in s["ids"]])
        self.log("info", f"{ip}:{PORT} ID {uid} 연결")
        self.worker.submit(job_connect, ip, PORT, TIMEOUT, uid, done=self._connected)

    def _connected(self, res: dict) -> None:
        self.connecting = False
        if "error" in res:
            self.dev = None
            self.dot.configure(fg="#c0392b")
            self.lbl_status.configure(text=f"연결 실패: {res['error']}")
            self.log("error", f"연결 실패: {res['error']}")
            self._update_controls()
            return
        key = (res["host"], res["uid"])
        h = res["h"]
        if key != self.hist_dev:                     # 새 장치: 이전 장치의 기록·지정 정보 지움
            if self.history:
                self.log("info", f"이전 장치({self.hist_dev[0]} ID {self.hist_dev[1]})의 속도 기록 "
                                 f"{len(self.history)}개 지움")
            self.history.clear()
            self.hist_dev = key
            self.target = None
            self.gen += 1
            self.mode_tries = 0
            self.orig = None
            self.lbl_target.configure(text="지정 속도: 없음")
        if self.orig is None and all(h.get(k) is not None for k in (1, 2, 4)):
            self.orig = {k: h[k] for k in (1, 2, 4)}
        self.dev = {k: res[k] for k in ("host", "port", "uid", "kind", "name", "product", "ecblue")}
        self.data = {"i": dict(res["i"]), "h": dict(h)}
        self.last_ok = None
        self._set_ok(True)
        self.log("info", f"연결됨: {res['host']} ID {res['uid']} - {res['name']} (제품코드 0x{res['product']:04X}, "
                         f"펌웨어 {ms.fw_text(res['i'].get(0))})")
        self.log("info", "  자동 운전: " + wd_summary(h.get(17), h.get(5)))
        self._show_registers()
        self._show_wd(fill=True)
        self.lbl_read.configure(text=f"읽은 시각 {datetime.datetime.now():%H:%M:%S}")
        if res["ecblue"]:
            self.add_sample(self._now(), res["i"].get(14), h.get(2))
            self._schedule_sample()
        else:
            self.log("warn", "ECblue 가 아닙니다 - 값은 보여 주지만 속도 확인·지정·자동 운전 설정은 하지 않습니다")
        self._update_controls()
        self._update_now()
        self._update_check_label()
        self.draw_graph()

    def on_disconnect(self) -> None:
        d = self.dev
        if d is None:
            return
        self._stop_sampling()
        self.dev = None
        self.worker.submit(job_close)
        self.dot.configure(fg="#808080")
        self.lbl_status.configure(text=f"연결 끊음 ({d['host']} ID {d['uid']}) - 속도 기록은 다른 장치에 연결하거나 "
                                       "프로그램을 닫을 때까지 남아 있습니다")
        self.log("info", f"연결 끊음: {d['host']} ID {d['uid']}")
        self._update_controls()

    def on_reread(self) -> None:
        d = self.dev
        if d is None:
            return
        self.log("info", "레지스터 새로 읽기")
        self.worker.submit(job_read_all, d["host"], d["port"], TIMEOUT, d["uid"], done=self._reread_done)

    def _reread_done(self, res: dict) -> None:
        if "error" in res or self._stale(res):
            if "error" in res:
                self.log("error", f"읽기 오류: {res['error']}")
            return
        self._apply_data(res, sample=False)

    # ------------------------------------------------------------------ 속도 확인·기록
    def _stop_sampling(self) -> None:
        if self._sample_job is not None:
            try:
                self.root.after_cancel(self._sample_job)
            except tk.TclError:
                pass
            self._sample_job = None

    def _schedule_sample(self) -> None:
        """다음 속도 확인 예약 (지금 주기로)."""
        self._stop_sampling()
        if self.dev is not None and self.dev.get("ecblue") and self._alive:
            self._sample_job = self.root.after(int(self.interval_sec() * 1000), self._sample_tick)

    def _sample_tick(self) -> None:
        self._sample_job = None
        d = self.dev
        if d is None:
            return
        if not self.busy_sample:
            self.busy_sample = True
            self.worker.submit(job_sample, d["host"], d["port"], TIMEOUT, d["uid"], done=self._sampled)
        self._schedule_sample()

    def _sampled(self, res: dict) -> None:
        self.busy_sample = False
        if "error" in res:
            self.log("error", f"속도 확인 오류: {res['error']}")
            return
        if self._stale(res):
            return
        i, h = res["i"], res["h"]
        rpm, h2 = i.get(14), h.get(2)
        self.data["i"].update(i)
        self.data["h"].update(h)
        self.add_sample(self._now(), rpm, h2)
        self._set_ok(rpm is not None)
        self._show_registers()
        self._update_now()
        self._update_check_label()
        self.draw_graph()
        self._maybe_keep(h)

    def add_sample(self, t: float, rpm: Optional[int], h2: Optional[int]) -> None:
        """기록 추가. 24시간(KEEP_SEC)보다 오래된 기록은 지움."""
        self.history.append((t, rpm, h2))
        cut = t - KEEP_SEC
        while self.history and self.history[0][0] < cut:
            self.history.popleft()

    def _maybe_keep(self, h: Dict[int, Optional[int]]) -> None:
        """지정 속도 유지: 모터의 h2 가 바뀌어 있으면(정전 후 재시작 등) 다시 지정."""
        d, h2 = self.dev, h.get(2)
        if (not self.keep_var.get() or self.target is None or h2 is None or h2 == self.target
                or not modbus.ecb_mode_ok(h.get(4) or 0) or self.keep_pending or d is None):
            return
        self.log("warn", f"지정 속도 유지: 모터 값 h2={h2} -> {self.target} 다시 지정")
        gen = self.gen
        self.keep_pending = True

        def done(res: dict) -> None:
            self.keep_pending = False
            if not res.get("skipped"):
                self._log_steps(res.get("steps", []))
        self.worker.submit(job_write_speed, d["host"], d["port"], TIMEOUT, d["uid"], self.target,
                           lambda: self.gen == gen, done=done)

    def set_range(self, key: str) -> None:
        """그래프 보기 바꾸기. 10분이 아니면 [10초 단위 체크]는 꺼지고 30초 주기로."""
        self.range_var.set(key)
        if key != FAST_RANGE and self.fast_var.get():
            self.fast_var.set(False)
            self.log("info", f"10초 단위 체크 끔 ({FAST_RANGE} 보기에서만 동작) - {CHECK_SEC}초마다 확인")
            self._schedule_sample()
        self._update_controls()
        self._update_check_label()
        self.draw_graph()

    def on_fast(self) -> None:
        if self.range_var.get() != FAST_RANGE:       # 버튼이 막혀 있어 보통은 오지 않음
            self.fast_var.set(False)
            return
        self.log("info", f"10초 단위 체크 {'켬' if self.fast_var.get() else '끔'} - {self.interval_sec()}초마다 확인")
        self._schedule_sample()
        self._update_check_label()

    def _update_check_label(self) -> None:
        n = len(self.history)
        span = self.history[-1][0] - self.history[0][0] if n >= 2 else 0
        self.lbl_check.configure(text=f"{self.interval_sec()}초마다 확인 · 기록 {n}개 ({dur_text(span)}) · "
                                      "최대 24시간, 닫으면 지워짐")

    def _update_now(self) -> None:
        if not self.history:
            self.lbl_now.configure(text="기록 없음")
            return
        t, v, s = self.history[-1]
        text = (f"현재 {'-' if v is None else v} RPM   지정 {'-' if s is None else s} RPM   "
                f"마지막 확인 {time.strftime('%H:%M:%S', time.localtime(t))}")
        h8 = self.data.get("h", {}).get(8)
        if h8:
            text += f"   최고속도(h8) {h8} RPM"
        self.lbl_now.configure(text=text, fg="#c0392b" if v is None else "#222222")

    def _redraw_tick(self) -> None:
        self._redraw_job = None
        if not self._alive:
            return
        self.draw_graph()
        self._update_check_label()
        self._redraw_job = self.root.after(REDRAW_MS, self._redraw_tick)

    # ------------------------------------------------------------------ 그래프
    @staticmethod
    def segments(pts: List[tuple], idx: int, step: bool) -> List[List[Tuple[float, int]]]:
        """선 조각들. 값이 없거나(통신 끊김) 기록 사이가 GAP_SEC 보다 길면 끊음. step: 계단선(지정 속도)."""
        segs: List[List[Tuple[float, int]]] = []
        cur: List[Tuple[float, int]] = []
        last_t: Optional[float] = None
        for p in pts:
            t, v = p[0], p[idx]
            if v is None or (last_t is not None and t - last_t > GAP_SEC):
                if cur:
                    segs.append(cur)
                cur = []
            if v is not None:
                if step and cur:
                    cur.append((t, cur[-1][1]))
                cur.append((t, v))
            last_t = t
        if cur:
            segs.append(cur)
        return segs

    def draw_graph(self) -> None:
        c = self.canvas
        c.delete("all")
        self._plot, self._geom = [], None
        W, H = c.winfo_width(), c.winfo_height()
        if W < 120 or H < 80:
            return
        x0, x1, y0, y1 = 52, W - 14, 26, H - 22
        key = self.range_var.get()
        span = RANGE_SEC[key]
        now = self._now()
        t0 = now - span
        pts = [p for p in self.history if p[0] >= t0]
        vals = [p[1] for p in pts if p[1] is not None] + [p[2] for p in pts if p[2] is not None]
        vmax = max(vals + [100])
        step = next(st for st in (50, 100, 200, 250, 500, 1000, 2000, 5000, 10000) if vmax * 1.1 / st <= 6)
        ymax = math.ceil(vmax * 1.1 / step) * step

        def X(t: float) -> float:
            return x0 + (t - t0) / span * (x1 - x0)

        def Y(v: float) -> float:
            return y1 - v / ymax * (y1 - y0)
        for k in range(int(ymax // step) + 1):                    # 가로 눈금 (RPM)
            y = Y(k * step)
            c.create_line(x0, y, x1, y, fill="#ececec" if k else "#b0b0b0")
            c.create_text(x0 - 6, y, text=str(k * step), anchor="e", fill="#555555", font=SMALL)
        tick = TICK_SEC[key]                                       # 세로 눈금 (시각, 이 PC 의 시간대 기준)
        off = utc_offset(now)
        t = math.ceil((t0 + off) / tick) * tick - off
        while t <= now:
            x = X(t)
            c.create_line(x, y0, x, y1, fill="#f2f2f2")
            c.create_text(x, y1 + 3, text=time.strftime("%H:%M", time.localtime(t)), anchor="n", fill="#555555",
                          font=SMALL, tags=("xlabel",))
            t += tick
        c.create_rectangle(x0, y0, x1, y1, outline="#b0b0b0")
        c.create_line(x0 + 4, 12, x0 + 24, 12, fill=ACTUAL_COLOR, width=2)
        c.create_text(x0 + 28, 12, text="실제 속도 (i14)", anchor="w", font=SMALL, fill="#333333")
        c.create_line(x0 + 132, 12, x0 + 152, 12, fill=SET_COLOR, width=2, dash=(4, 3))
        c.create_text(x0 + 156, 12, text="지정 속도 (h2)", anchor="w", font=SMALL, fill="#333333")
        self._geom = (x0, x1, y0, y1, t0, span, ymax)
        if not pts:
            msg = (f"{key} 동안 기록 없음 - 연결하면 {self.interval_sec()}초마다 속도를 기록합니다"
                   if self.dev is None or not self.history else f"{key} 동안 기록 없음")
            c.create_text((x0 + x1) / 2, (y0 + y1) / 2, text=msg, fill="#888888", font=FONT, tags=("empty",))
            return
        self._plot = pts
        for seg in self.segments(pts, 2, True):
            if len(seg) >= 2:
                c.create_line(*[q for t_, v in seg for q in (X(t_), Y(v))], fill=SET_COLOR, width=2, dash=(4, 3),
                              tags=("set",))
        dots = len(pts) <= 120
        for seg in self.segments(pts, 1, False):
            if len(seg) >= 2:
                c.create_line(*[q for t_, v in seg for q in (X(t_), Y(v))], fill=ACTUAL_COLOR, width=2,
                              tags=("actual",))
            if dots or len(seg) == 1:
                for t_, v in seg:
                    c.create_oval(X(t_) - 2, Y(v) - 2, X(t_) + 2, Y(v) + 2, fill=ACTUAL_COLOR, outline="",
                                  tags=("dot",))
        if self._hover_x is not None:
            self._render_hover()

    def _on_hover(self, e) -> None:
        self._hover_x = e.x
        self._render_hover()

    def _on_leave(self, _e=None) -> None:
        self._hover_x = None
        self.hover_text = ""
        self.canvas.delete("hover")

    def _render_hover(self) -> None:
        """마우스 위치에서 가장 가까운 기록의 시각과 속도를 보여 줌."""
        c = self.canvas
        c.delete("hover")
        self.hover_text = ""
        g, x = self._geom, self._hover_x
        if g is None or not self._plot or x is None:
            return
        x0, x1, y0, y1, t0, span, ymax = g
        if not x0 <= x <= x1:
            return
        t = t0 + (x - x0) / (x1 - x0) * span
        times = [p[0] for p in self._plot]
        k = bisect.bisect_left(times, t)
        j = min((j for j in (k - 1, k) if 0 <= j < len(times)), key=lambda j: abs(times[j] - t))
        pt, v, s = self._plot[j]
        px = x0 + (pt - t0) / span * (x1 - x0)
        c.create_line(px, y0, px, y1, fill="#888888", dash=(2, 2), tags=("hover",))
        if v is not None:
            py = y1 - v / ymax * (y1 - y0)
            c.create_oval(px - 4, py - 4, px + 4, py + 4, outline=ACTUAL_COLOR, width=2, tags=("hover",))
        self.hover_text = (f"{time.strftime('%H:%M:%S', time.localtime(pt))}  실제 {'응답 없음' if v is None else f'{v} RPM'}"
                           f"  지정 {'-' if s is None else f'{s} RPM'}")
        right = px + 8 + 260 < x1
        item = c.create_text(px + 8 if right else px - 8, y0 + 6, text=self.hover_text, anchor="nw" if right else "ne",
                             font=SMALL, fill="#222222", tags=("hover",))
        bb = c.bbox(item)
        if bb:
            box = c.create_rectangle(bb[0] - 3, bb[1] - 2, bb[2] + 3, bb[3] + 2, fill="#ffffe8", outline="#bbbbbb",
                                     tags=("hover",))
            c.tag_raise(item, box)

    # ------------------------------------------------------------------ 속도 지정
    def on_set_speed(self, allow_mode: bool = False) -> None:
        d = self.dev
        if d is None:
            return
        rpm = self._int_entry(self.ent_rpm, 0, MAX_RPM, "속도는")
        if rpm is None:
            return
        if not allow_mode:
            h8 = self.data["h"].get(8) or 0
            extra = f"\n(모터 최고속도 h8 = {h8} RPM 보다 높아 {h8} RPM 으로 제한됩니다)" if 0 < h8 < rpm else ""
            if not messagebox.askyesno("속도 지정", f"{d['host']} ID {d['uid']} 를 {rpm} RPM 으로 운전합니다.{extra}\n"
                                                  "계속할까요?", parent=self.root):
                return
        self.gen += 1
        self.log("info", f"속도 지정 {rpm} RPM 시작")
        self.worker.submit(job_set_speed, d["host"], d["port"], TIMEOUT, d["uid"], rpm, allow_mode, self.mode_tries,
                           done=self._set_done)

    def _set_done(self, res: dict) -> None:
        if "error" in res:
            self.log("error", f"속도 지정 오류: {res['error']}")
            return
        if res.get("need_mode"):
            self.log("warn", f"제어모드가 통신 속도 제어가 아님 (h4 = {ms.view_text(4, res['h4'])})")
            if not self._stale(res) and messagebox.askyesno(
                    "제어모드 변경",
                    f"ID {res['uid']} 의 제어모드(h4)를 '통신 속도 제어 + 버스 운전허가' 로 바꿔야 합니다.\n"
                    f"지금: {ms.view_text(4, res['h4'])}\n\n"
                    "h4 는 전원을 꺼도 유지되는 값이며, 바뀐 동안 모터는 아날로그(0-10V) 신호 대신 통신으로 받은 속도로 "
                    "돕니다. [원래대로] 버튼으로 되돌릴 수 있습니다.\n\n바꿀까요?", icon="warning", parent=self.root):
                self.on_set_speed(allow_mode=True)
            else:
                self.log("info", "제어모드 변경 취소 - 속도 지정 안 함")
            return
        if res.get("mode_written"):
            self.mode_tries += 1
        self._log_steps(res["steps"])
        if res.get("ok"):
            self.target = res["rpm"]
            self.lbl_target.configure(text=f"지정 속도: {self.target} RPM")
            self.log("info", f"속도 지정 완료: {self.target} RPM (실제 회전수는 가속 후 따라옴)")
        else:
            self.log("error", "속도 지정 실패")
        if "data" in res and not self._stale(res):
            self._apply_data(res["data"])

    def on_stop(self) -> None:
        d = self.dev
        if d is None or not messagebox.askyesno("정지", f"{d['host']} ID {d['uid']} 의 지정 속도(h2)를 0 으로 "
                                                       "합니다.\n계속할까요?", parent=self.root):
            return
        self.gen += 1
        self.target = 0
        self.lbl_target.configure(text="지정 속도: 0 RPM (정지)")
        self.log("info", "정지(지정 속도 0)")
        self.worker.submit(job_write_speed, d["host"], d["port"], TIMEOUT, d["uid"], 0,
                           done=lambda res: self._log_steps(res.get("steps", [])))

    def on_restore(self) -> None:
        d, o = self.dev, self.orig
        if d is None:
            return
        if o is None:
            messagebox.showinfo("원래대로", "처음 연결할 때 설정을 읽지 못해 되돌릴 기준이 없습니다.", parent=self.root)
            return
        if not messagebox.askyesno("원래대로", f"{d['host']} ID {d['uid']} 를 처음 연결할 때 읽은 설정으로 되돌립니다.\n"
                                              f"h1 = {o[1]}, h2 = {o[2]}, h4 = {ms.view_text(4, o[4])}\n계속할까요?",
                                   parent=self.root):
            return
        self.gen += 1
        self.target = None
        self.lbl_target.configure(text="지정 속도: 없음")
        self.log("info", "원래대로 되돌리기")
        self.worker.submit(job_restore, d["host"], d["port"], TIMEOUT, d["uid"], dict(o), done=self._restore_done)

    def _restore_done(self, res: dict) -> None:
        if "error" in res:
            self.log("error", f"원래대로 오류: {res['error']}")
            return
        self._log_steps(res["steps"])
        self.log("info" if res["ok"] else "error", "원래대로 완료" if res["ok"] else "원래대로 확인 안 됨")
        if "data" in res and not self._stale(res):
            self._apply_data(res["data"])

    # ------------------------------------------------------------------ 통신 끊김 시 자동 운전
    def on_wd_set(self) -> None:
        d = self.dev
        if d is None:
            return
        sec = self._int_entry(self.ent_wd_sec, 1, 255, "시작 시간은")
        if sec is None:
            return
        rpm = self._int_entry(self.ent_wd_rpm, 0, MAX_RPM, "자동 운전 속도는")
        if rpm is None:
            return
        h = self.data["h"]
        h8 = h.get(8) or 0
        if 0 < h8 < rpm:
            messagebox.showerror(TITLE, f"자동 운전 속도 {rpm} RPM 이 모터 최고속도(h8) {h8} RPM 보다 높습니다.",
                                 parent=self.root)
            return
        mode, mode_name = WD_MODE_CHOICES[max(self.cb_wd_mode.current(), 0)]
        changes = ms.watchdog_changes(h.get(5), sec, mode, rpm)
        warn = ""
        if sec < WD_RECOMMENDED_SEC:
            warn = (f"\n\n주의: 시작 시간 {sec}초는 짧습니다. 운영 프로그램과 이 프로그램은 {CHECK_SEC}초마다 통신하므로 "
                    "통신 중에도 자동 운전으로 바뀔 수 있습니다 (60초 이상 권장).")
        msg = (f"{d['host']} ID {d['uid']} 에 '통신 끊김 시 자동 운전' 을 설정합니다.\n\n"
               f"· 통신(제어 신호)이 {sec}초 동안 없으면 {rpm} RPM 으로 운전 ({mode_name})\n"
               f"· 쓰는 값: {', '.join(f'h{a} = {v}' for a, v in changes)} (전원을 꺼도 유지)" + warn + "\n\n계속할까요?")
        if not messagebox.askyesno("자동 운전 설정", msg, icon="warning" if warn else "question", parent=self.root):
            return
        self.log("info", f"자동 운전 설정: {sec}초 동안 통신이 없으면 {rpm} RPM ({mode_name})")
        self.worker.submit(job_write, d["host"], d["port"], TIMEOUT, d["uid"], changes,
                           done=lambda res: self._wd_done("자동 운전 설정", res))

    def on_wd_off(self) -> None:
        d = self.dev
        if d is None or not messagebox.askyesno("자동 운전 끄기", f"{d['host']} ID {d['uid']} 의 '통신 끊김 시 자동 운전' "
                                                             "을 끕니다 (h17 = 0).\n계속할까요?", parent=self.root):
            return
        self.log("info", "자동 운전 끄기")
        self.worker.submit(job_write, d["host"], d["port"], TIMEOUT, d["uid"], [(17, 0)],
                           done=lambda res: self._wd_done("자동 운전 끄기", res))

    def _wd_done(self, title: str, res: dict) -> None:
        if "error" in res:
            self.log("error", f"{title} 오류: {res['error']}")
            return
        self._log_steps(res["steps"])
        if res["ok"]:
            h = res["data"]["h"]
            self.log("info", f"{title} 완료 - 지금: {wd_summary(h.get(17), h.get(5))}")
        else:
            self.log("error", f"{title} 실패 - 위 로그 확인")
        if not self._stale(res):
            self._apply_data(res["data"], sample=False)

    # ------------------------------------------------------------------ 종료
    def on_close(self) -> None:
        msg = "프로그램을 닫을까요?\n모터는 마지막으로 지정한 상태로 계속 운전합니다."
        if self.history:
            msg += f"\n속도 기록 {len(self.history)}개는 지워집니다."
        if not messagebox.askyesno(TITLE, msg, parent=self.root):
            return
        self.shutdown()
        self.root.destroy()

    def shutdown(self) -> None:
        if not self._alive:
            return
        self._alive = False
        self._stop_sampling()
        if self._redraw_job is not None:
            try:
                self.root.after_cancel(self._redraw_job)
            except tk.TclError:
                pass
            self._redraw_job = None
        n = len(self.history)
        self.history.clear()                          # 속도 기록은 메모리에만 있으며 닫으면 지움
        self.log("info", f"{TITLE} 종료 - 속도 기록 {n}개 지움, 모터 설정은 바꾸지 않음")
        self.worker.stop()
        self.worker.join(timeout=5)
        log.removeHandler(self._fh)
        self._fh.close()


def main() -> None:
    try:
        root = tk.Tk()
        ConfiguratorApp(root)
        root.mainloop()
    except Exception:
        # 콘솔 없는 exe 에서도 원인을 알 수 있도록 파일과 메시지 창으로 알림
        import traceback
        os.makedirs(LOG_DIR, exist_ok=True)
        path = os.path.join(LOG_DIR, "ecblue_configurator_error.txt")
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now()}\n{traceback.format_exc()}\n")
        try:
            messagebox.showerror(TITLE, f"시작 중 오류가 발생했습니다.\n{path}")
        except tk.TclError:
            pass
        raise


if __name__ == "__main__":
    main()
