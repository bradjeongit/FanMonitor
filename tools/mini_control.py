"""ECblue 미니 제어기.

항목(제어할 팬)마다 IP·ID 를 입력하고, 속도(%) 또는 참조팬(IP·ID)을 지정합니다.
- 속도 %: 1100 RPM 에 대한 비율 (100% = 1100 RPM)
- 참조팬: 참조팬의 실제 회전수(i14)를 읽어 같은 속도로 맞춤 (최대 1100)
[적용] 후에는 다른 지정을 할 때까지 정기적으로 속도를 점검하고, 지정값이 풀려 있으면 다시 지정합니다.
진행 내용은 아래 로그창과 logs/mini_control_YYYYMMDD.log 에 시간과 함께 남습니다.

실행: MINI_CONTROL.exe 또는 mini_control.bat 더블클릭, 또는 python tools/mini_control.py
항목 목록은 실행 폴더의 mini_control.json 에 저장됩니다. 운영 Config.dll 은 사용하지 않습니다.
"""
from __future__ import annotations

import datetime
import ipaddress
import json
import logging
import os
import queue
import sys
import threading
import tkinter as tk
from dataclasses import asdict, dataclass, field
from tkinter import messagebox, ttk
from typing import Callable, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from motor_monitor import modbus  # noqa: E402

try:  # 같은 폴더의 시험기 통신부 재사용 (python tools/mini_control.py 실행 시)
    import ecblue_tester as _bus_mod  # noqa: E402
except ImportError:  # 패키지/exe 로 묶였을 때
    from tools import ecblue_tester as _bus_mod  # type: ignore  # noqa: E402

DEFAULT_IP = "172.27.0.116"
PORT = 502
MAX_RPM = 1100
ID_MIN, ID_MAX = 1, 247
CHECK_INTERVAL_MS = 30000         # 정기 점검 주기 (30초)
MODE_WRITE_MAX_TRIES = 3          # h4(영구저장) 쓰기 시도 한도 (실행 1회당, 항목별)
SPEED_TOLERANCE = 30              # 실제 속도가 지정과 이만큼 넘게 차이 나면 '주의' (RPM)

_BASE = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
         else os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_DIR = os.path.join(_BASE, "logs")
ITEMS_FILE = os.path.join(_BASE, "mini_control.json")

log = logging.getLogger("mini_control")

INFO, WARN, ERROR = "info", "warn", "error"


def pct_to_rpm(pct: float) -> int:
    return max(0, min(MAX_RPM, round(MAX_RPM * pct / 100)))


def rpm_to_pct(rpm: int) -> float:
    return round(rpm * 100 / MAX_RPM, 1)


# ====================================================================== 항목
@dataclass(eq=False)
class Item:
    ip: str
    uid: int
    mode: str = "pct"            # "pct" 속도 %, "ref" 참조팬
    pct: float = 0.0
    ref_ip: str = ""
    ref_uid: int = 0
    # ---- 실행 중 상태 (저장 안 함) ----
    active: bool = field(default=False)
    gen: int = field(default=0)            # 지정이 바뀔 때마다 증가 -> 이전 작업 결과 무시
    busy: bool = field(default=False)
    mode_allowed: bool = field(default=False)
    mode_tries: int = field(default=0)
    target: Optional[int] = field(default=None)
    rpm: Optional[int] = field(default=None)
    h2: Optional[int] = field(default=None)
    ref_rpm: Optional[int] = field(default=None)
    ref_state: str = field(default="-")
    state: str = field(default="대기 (적용 전)")
    level: str = field(default=INFO)
    checked: str = field(default="-")

    @property
    def name(self) -> str:
        return f"{self.ip} ID{self.uid}"

    def spec_text(self) -> str:
        if self.mode == "ref":
            return f"참조팬 {self.ref_ip} ID{self.ref_uid}"
        return f"{self.pct:g}% ({pct_to_rpm(self.pct)} RPM)"

    def to_json(self) -> dict:
        return {k: v for k, v in asdict(self).items() if k in ("ip", "uid", "mode", "pct", "ref_ip", "ref_uid")}


def load_items(path: str) -> List[Item]:
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return [Item(str(d["ip"]), int(d["uid"]), d.get("mode", "pct"), float(d.get("pct", 0)),
                     str(d.get("ref_ip", "")), int(d.get("ref_uid", 0))) for d in data]
    except FileNotFoundError:
        return []


def save_items(path: str, items: List[Item]) -> None:
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump([it.to_json() for it in items], f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


# ====================================================================== 통신 (작업 쓰레드)
class Buses:
    """IP 별 TCP 연결. 작업 쓰레드에서만 사용."""

    def __init__(self) -> None:
        self._buses: Dict[str, "_bus_mod.Bus"] = {}

    def get(self, ip: str) -> "_bus_mod.Bus":
        b = self._buses.get(ip)
        if b is None:
            b = self._buses[ip] = _bus_mod.Bus(ip, PORT)
        return b

    def close(self) -> None:
        for b in self._buses.values():
            b.close()


class Worker(threading.Thread):
    """통신 작업을 차례대로 하나씩 처리하고 결과를 UI 쪽 큐로 돌려줌."""

    def __init__(self) -> None:
        super().__init__(daemon=True, name="mini-control-worker")
        self.buses = Buses()
        self.jobs: "queue.Queue[Optional[tuple]]" = queue.Queue()
        self.results: "queue.Queue[tuple]" = queue.Queue()

    def submit(self, fn: Callable, *args, done: Callable[[dict], None]) -> None:
        self.jobs.put((fn, args, done))

    def run(self) -> None:
        while True:
            job = self.jobs.get()
            if job is None:
                break
            fn, args, done = job
            try:
                res = fn(self.buses, *args)
            except Exception as e:  # 작업 오류로 쓰레드가 죽지 않도록
                log.exception("작업 오류")
                res = {"steps": [(ERROR, f"프로그램 내부 오류: {e}")], "ok": False}
            self.results.put((done, res))
        self.buses.close()

    def stop(self) -> None:
        self.jobs.put(None)


def control_job(buses: Buses, snap: dict) -> dict:
    """참조팬 읽기 -> 목표 계산 -> 대상팬 읽기 -> (필요 시) 제어모드·Enable·속도 쓰기.

    snap: ip, uid, mode, pct, ref_ip, ref_uid, first(적용 직후 여부), last_target,
          mode_allowed, mode_tries
    결과: steps [(level, msg)], ok, need_mode, target, rpm, h2, ref_rpm, ref_state, mode_tries
    """
    steps: List[Tuple[str, str]] = []
    res: dict = {"steps": steps, "ok": False, "mode_tries": snap["mode_tries"]}
    first = snap["first"]

    def say(lv: str, msg: str) -> None:
        """적용 직후에는 모든 단계를, 정기 점검 때는 주의/에러만 기록."""
        if first or lv != INFO:
            steps.append((lv, msg))

    def connect(b) -> bool:
        if b.connected:
            return True
        ok, msg = b.connect()
        steps.append((INFO if ok else ERROR, msg))
        return ok

    # 1) 목표 속도
    target = snap["last_target"]
    if snap["mode"] == "ref":
        rb = buses.get(snap["ref_ip"])
        say(INFO, f"참조팬 {snap['ref_ip']} ID{snap['ref_uid']} 속도 읽기")
        r, d = (rb.request(modbus.build_read_input, snap["ref_uid"], modbus.ECB_I_SPEED, 1)
                if connect(rb) else (None, "접속 안 됨"))
        if r is not None and r.exception is None and r.values:
            res["ref_rpm"] = r.values[0]
            res["ref_state"] = "정상"
            target = min(r.values[0], MAX_RPM)
            say(INFO, f"참조팬 속도 {r.values[0]} RPM ({d})"
                + (f" -> 최대 {MAX_RPM} 로 제한" if r.values[0] > MAX_RPM else ""))
        else:
            res["ref_state"] = "응답 없음" if r is None else "예외 응답"
            if target is None:
                steps.append((ERROR, f"참조팬 {snap['ref_ip']} ID{snap['ref_uid']} 읽기 실패: {d} - 속도를 지정하지 않음"))
                return res
            steps.append((ERROR, f"참조팬 {snap['ref_ip']} ID{snap['ref_uid']} 읽기 실패: {d} - 이전 지정 {target} RPM 유지"))
    else:
        target = pct_to_rpm(snap["pct"])
    res["target"] = target

    # 2) 대상팬 현재 상태
    bus = buses.get(snap["ip"])
    uid = snap["uid"]
    say(INFO, "대상팬 설정(h1~h4) 읽기")
    if not connect(bus):
        steps.append((ERROR, "대상팬에 접속하지 못해 제어하지 못함"))
        return res
    r, d = bus.request(modbus.build_read_holding, uid, modbus.ECB_H_CONTROL, 4)
    if r is None or r.exception is not None or len(r.values) < 4:
        steps.append((ERROR, f"대상팬 설정 읽기 실패: {d}"))
        return res
    h1, h2, _h3, h4 = r.values[:4]
    say(INFO, f"h1={h1} h2={h2} h4={_bus_mod.mode_text(h4)} ({d})")

    # 3) 제어모드 (영구저장 h4 - 확인 후 제한된 횟수만)
    if not modbus.ecb_mode_ok(h4):
        if not snap["mode_allowed"]:
            res["need_mode"] = True
            res["h4"] = h4
            steps.append((WARN, f"제어모드가 통신 속도제어가 아님 (h4={_bus_mod.mode_text(h4)}) - 변경 확인 필요"))
            return res
        if res["mode_tries"] >= MODE_WRITE_MAX_TRIES:
            steps.append((ERROR, f"제어모드(h4) 쓰기 {MODE_WRITE_MAX_TRIES}회 실패 - 더 시도하지 않음 (PIN 보호/통신 확인)"))
            return res
        res["mode_tries"] += 1
        steps.append((INFO if first else WARN, f"제어모드 변경: PIN 입력 후 h4 {_bus_mod.mode_text(h4)} -> "
                      f"{_bus_mod.mode_text(modbus.ecb_mode_value(h4))}"))
        _write(bus, uid, modbus.ECB_H_PIN, modbus.ECB_PIN_LEVEL2, steps, "PIN 입력")  # 실패해도 h4 는 시도
        if not _write(bus, uid, modbus.ECB_H_CONTROLMODE, modbus.ecb_mode_value(h4), steps, "제어모드(h4)"):
            return res

    # 4) Enable, 5) 운전속도
    if not h1 & modbus.ECB_ENABLE_BIT:
        if not _write(bus, uid, modbus.ECB_H_CONTROL, h1 | modbus.ECB_ENABLE_BIT, steps, "운전 허가(h1 Enable)"):
            return res
    if h2 != target:
        if not first:
            steps.append((WARN, f"지정 속도가 바뀌어 있음 (h2={h2}) -> {target} RPM 다시 지정"))
        if not _write(bus, uid, modbus.ECB_H_SPEED, target, steps, f"운전속도(h2) {target} RPM"):
            return res
        h2 = target
    else:
        say(INFO, f"운전속도 이미 {target} RPM")
    res["h2"] = h2

    # 6) 실제 회전수
    r, d = bus.request(modbus.build_read_input, uid, modbus.ECB_I_SPEED, 1)
    if r is None or r.exception is not None or not r.values:
        steps.append((ERROR, f"대상팬 실제 속도 읽기 실패: {d}"))
        return res
    res["rpm"] = r.values[0]
    res["ok"] = True
    return res


def _write(bus, uid: int, reg: int, value: int, steps: List[Tuple[str, str]], what: str) -> bool:
    r, d = bus.request(modbus.build_new_write_single, uid, reg, value)
    if r is None or r.exception is not None:
        steps.append((ERROR, f"{what} 쓰기 실패: {d}"))
        return False
    steps.append((INFO, f"{what} 쓰기 완료 ({d})"))
    return True


def stop_job(buses: Buses, snap: dict) -> dict:
    """운전속도 0 지정 (제어모드는 그대로)."""
    steps: List[Tuple[str, str]] = []
    bus = buses.get(snap["ip"])
    if not bus.connected:
        ok, msg = bus.connect()
        steps.append((INFO if ok else ERROR, msg))
        if not ok:
            return {"steps": steps, "ok": False}
    ok = _write(bus, snap["uid"], modbus.ECB_H_SPEED, 0, steps, "운전속도(h2) 0 RPM (정지)")
    return {"steps": steps, "ok": ok}


# ====================================================================== 화면
class ItemDialog(tk.Toplevel):
    """항목 추가/수정 창."""

    def __init__(self, master, title: str, item: Optional[Item] = None) -> None:
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.transient(master)
        self.result: Optional[Item] = None
        it = item or Item(DEFAULT_IP, 1)
        self.v_ip = tk.StringVar(value=it.ip)
        self.v_id = tk.StringVar(value=str(it.uid) if item else "")
        self.v_mode = tk.StringVar(value=it.mode)
        self.v_pct = tk.StringVar(value=f"{it.pct:g}" if item else "")
        self.v_rip = tk.StringVar(value=it.ref_ip or DEFAULT_IP)
        self.v_rid = tk.StringVar(value=str(it.ref_uid) if it.ref_uid else "")
        self.v_rpm = tk.StringVar()

        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="제어할 팬", font=("", 10, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(f, text="IP").grid(row=1, column=0, sticky="e", padx=4, pady=3)
        self.e_ip = ttk.Entry(f, textvariable=self.v_ip, width=18)
        self.e_ip.grid(row=1, column=1, sticky="w")
        ttk.Label(f, text="ID").grid(row=2, column=0, sticky="e", padx=4, pady=3)
        self.e_id = ttk.Entry(f, textvariable=self.v_id, width=8)
        self.e_id.grid(row=2, column=1, sticky="w")
        ttk.Label(f, text=f"({ID_MIN}~{ID_MAX})").grid(row=2, column=2, sticky="w")

        ttk.Separator(f).grid(row=3, column=0, columnspan=3, sticky="ew", pady=8)
        ttk.Label(f, text="속도 지정 방법", font=("", 10, "bold")).grid(row=4, column=0, columnspan=3, sticky="w")
        ttk.Radiobutton(f, text="속도 %", variable=self.v_mode, value="pct",
                        command=self._on_mode).grid(row=5, column=0, sticky="w", pady=3)
        self.e_pct = ttk.Entry(f, textvariable=self.v_pct, width=8)
        self.e_pct.grid(row=5, column=1, sticky="w")
        ttk.Label(f, textvariable=self.v_rpm, foreground="#555").grid(row=5, column=2, sticky="w")
        ttk.Radiobutton(f, text="참조팬", variable=self.v_mode, value="ref",
                        command=self._on_mode).grid(row=6, column=0, sticky="w", pady=3)
        ttk.Label(f, text="IP").grid(row=7, column=0, sticky="e", padx=4)
        self.e_rip = ttk.Entry(f, textvariable=self.v_rip, width=18)
        self.e_rip.grid(row=7, column=1, sticky="w")
        ttk.Label(f, text="ID").grid(row=8, column=0, sticky="e", padx=4, pady=3)
        self.e_rid = ttk.Entry(f, textvariable=self.v_rid, width=8)
        self.e_rid.grid(row=8, column=1, sticky="w")
        ttk.Label(f, text="참조팬의 실제 속도를 그대로 따라갑니다", foreground="#555").grid(
            row=9, column=0, columnspan=3, sticky="w")

        b = ttk.Frame(f)
        b.grid(row=10, column=0, columnspan=3, sticky="e", pady=(12, 0))
        ttk.Button(b, text="확인", command=self._ok).pack(side="left", padx=4)
        ttk.Button(b, text="취소", command=self.destroy).pack(side="left")
        self.v_pct.trace_add("write", lambda *_: self._show_rpm())
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self._on_mode()
        self._show_rpm()
        (self.e_id if not item else self.e_pct).focus_set()
        self.grab_set()

    def _on_mode(self) -> None:
        ref = self.v_mode.get() == "ref"
        self.e_pct.configure(state="disabled" if ref else "normal")
        for e in (self.e_rip, self.e_rid):
            e.configure(state="normal" if ref else "disabled")

    def _show_rpm(self) -> None:
        try:
            p = float(self.v_pct.get())
            self.v_rpm.set(f"= {pct_to_rpm(p)} RPM (100% = {MAX_RPM})" if 0 <= p <= 100 else "0~100 사이")
        except ValueError:
            self.v_rpm.set(f"(100% = {MAX_RPM} RPM)")

    def _err(self, msg: str, widget) -> None:
        messagebox.showerror("입력 오류", msg, parent=self)
        widget.focus_set()

    def _ok(self) -> None:
        ip = self.v_ip.get().strip()
        if not _valid_ip(ip):
            return self._err("제어할 팬의 IP 형식이 올바르지 않습니다.", self.e_ip)
        uid = _int(self.v_id.get())
        if uid is None or not ID_MIN <= uid <= ID_MAX:
            return self._err(f"ID 는 {ID_MIN}~{ID_MAX} 사이 정수로 입력하세요.", self.e_id)
        mode = self.v_mode.get()
        pct, rip, rid = 0.0, "", 0
        if mode == "pct":
            try:
                pct = float(self.v_pct.get())
            except ValueError:
                pct = -1
            if not 0 <= pct <= 100:
                return self._err("속도 % 는 0~100 사이 숫자로 입력하세요.", self.e_pct)
        else:
            rip = self.v_rip.get().strip()
            if not _valid_ip(rip):
                return self._err("참조팬 IP 형식이 올바르지 않습니다.", self.e_rip)
            rid = _int(self.v_rid.get())
            if rid is None or not ID_MIN <= rid <= ID_MAX:
                return self._err(f"참조팬 ID 는 {ID_MIN}~{ID_MAX} 사이 정수로 입력하세요.", self.e_rid)
            if (rip, rid) == (ip, uid):
                return self._err("참조팬을 자기 자신으로 지정할 수 없습니다.", self.e_rid)
        self.result = Item(ip, uid, mode, pct, rip, rid)
        self.destroy()


def _valid_ip(s: str) -> bool:
    try:
        ipaddress.IPv4Address(s)
        return True
    except ValueError:
        return False


def _int(s: str) -> Optional[int]:
    try:
        return int(s.strip())
    except ValueError:
        return None


class MiniControlApp:
    COLS = (("name", "제어할 팬", 140), ("spec", "지정", 175), ("ref", "참조팬 상태 / 속도", 140),
            ("target", "지정 RPM", 100), ("rpm", "실제 RPM", 70), ("state", "상태", 270), ("checked", "마지막 점검", 80))

    def __init__(self, root: tk.Tk, items_file: Optional[str] = None) -> None:
        self.root = root
        self.items_file = items_file or ITEMS_FILE
        root.title("ECblue 미니 제어기")
        root.geometry("1040x620")
        root.minsize(820, 480)
        self.items: List[Item] = []
        self.worker = Worker()
        self.worker.start()
        self._closing = False
        self._setup_file_log()

        bar = ttk.Frame(root, padding=(8, 6))
        bar.pack(fill="x")
        for text, cmd in (("항목 추가", self.on_add), ("수정", self.on_edit), ("삭제", self.on_delete)):
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=2)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        for text, cmd in (("적용", self.on_apply), ("정지 (0 RPM)", self.on_stop), ("점검 중지", self.on_pause)):
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=2)
        ttk.Label(bar, text=f"점검 주기 {CHECK_INTERVAL_MS // 1000}초 · 100% = {MAX_RPM} RPM",
                  foreground="#555").pack(side="right")

        pw = ttk.PanedWindow(root, orient="vertical")
        pw.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        tf = ttk.Frame(pw)
        self.tree = ttk.Treeview(tf, columns=[c[0] for c in self.COLS], show="headings", selectmode="extended",
                                 height=8)
        for key, text, w in self.COLS:
            self.tree.heading(key, text=text)
            self.tree.column(key, width=w, anchor="w" if key in ("name", "spec", "state", "ref") else "center")
        self.tree.tag_configure("error", foreground="#C00000")
        self.tree.tag_configure("warn", foreground="#B36B00")
        self.tree.tag_configure("idle", foreground="#777777")
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", lambda e: self.on_edit())
        pw.add(tf, weight=2)

        lf = ttk.LabelFrame(pw, text="로그")
        self.logbox = tk.Text(lf, height=12, wrap="none", font=("Consolas", 9), state="disabled")
        self.logbox.tag_configure("time", foreground="#666666")
        self.logbox.tag_configure("error", foreground="#D00000")
        self.logbox.tag_configure("warn", foreground="#B36B00")
        lsb = ttk.Scrollbar(lf, orient="vertical", command=self.logbox.yview)
        self.logbox.configure(yscrollcommand=lsb.set)
        self.logbox.pack(side="left", fill="both", expand=True)
        lsb.pack(side="right", fill="y")
        pw.add(lf, weight=3)

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        try:
            self.items = load_items(self.items_file)
        except (OSError, ValueError, KeyError, TypeError) as e:
            self.log(ERROR, f"항목 파일을 읽지 못함 ({self.items_file}): {e}")
        self.refresh()
        self.log(INFO, f"시작 - 항목 {len(self.items)}개 불러옴. 항목을 선택하고 [적용]을 누르면 제어를 시작합니다")
        root.after(100, self._poll)
        self._check_after = root.after(CHECK_INTERVAL_MS, self._periodic)

    # ---------------------------------------------------------------- 로그
    def _setup_file_log(self) -> None:
        self._file_log: Optional[logging.Handler] = None
        try:
            os.makedirs(LOG_DIR, exist_ok=True)
            h = logging.FileHandler(os.path.join(LOG_DIR, datetime.date.today().strftime("mini_control_%Y%m%d.log")),
                                    encoding="utf-8")
            h.setFormatter(logging.Formatter("%(asctime)s %(message)s"))
            log.addHandler(h)
            log.setLevel(logging.INFO)
            self._file_log = h
        except OSError:
            pass

    def log(self, level: str, msg: str, item: Optional[Item] = None) -> None:
        text = f"[{item.name}] {msg}" if item else msg
        prefix = {ERROR: "[에러] ", WARN: "[주의] "}.get(level, "")
        log.log({ERROR: logging.ERROR, WARN: logging.WARNING}.get(level, logging.INFO), prefix + text)
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        box = self.logbox
        box.configure(state="normal")
        box.insert("end", now + "  ", "time")
        tag = level if level in (ERROR, WARN) else ()
        box.insert("end", ("✖ " if level == ERROR else "⚠ " if level == WARN else "") + prefix + text + "\n", tag)
        if int(box.index("end-1c").split(".")[0]) > 3000:
            box.delete("1.0", "500.0")
        box.configure(state="disabled")
        box.see("end")

    # ---------------------------------------------------------------- 목록
    def refresh(self) -> None:
        sel = set(self.tree.selection())
        self.tree.delete(*self.tree.get_children())
        for i, it in enumerate(self.items):
            ref = "-"
            if it.mode == "ref":
                ref = it.ref_state + (f" / {it.ref_rpm} RPM" if it.ref_rpm is not None else "")
            vals = (it.name, it.spec_text(), ref,
                    "-" if it.target is None else f"{it.target} ({rpm_to_pct(it.target):g}%)",
                    "-" if it.rpm is None else it.rpm,
                    ("✖ [에러] " if it.level == ERROR else "⚠ " if it.level == WARN else "") + it.state,
                    it.checked)
            tag = it.level if it.level in (ERROR, WARN) else ("idle" if not it.active else "")
            self.tree.insert("", "end", iid=str(i), values=vals, tags=(tag,) if tag else ())
        for iid in sel:
            if self.tree.exists(iid):
                self.tree.selection_add(iid)

    def selected(self) -> List[Item]:
        return [self.items[int(i)] for i in self.tree.selection()]

    def _save(self) -> None:
        try:
            save_items(self.items_file, self.items)
        except OSError as e:
            self.log(ERROR, f"항목 저장 실패 ({self.items_file}): {e}")

    def _find(self, ip: str, uid: int, skip: Optional[Item] = None) -> Optional[Item]:
        return next((x for x in self.items if x is not skip and (x.ip, x.uid) == (ip, uid)), None)

    def on_add(self) -> None:
        d = ItemDialog(self.root, "항목 추가")
        self.root.wait_window(d)
        it = d.result
        if it is None:
            return
        if self._find(it.ip, it.uid):
            messagebox.showerror("항목 추가", f"{it.name} 은 이미 목록에 있습니다.", parent=self.root)
            return
        self.items.append(it)
        self._save()
        self.refresh()
        self.tree.selection_set(str(len(self.items) - 1))
        self.log(INFO, f"항목 추가: 지정 {it.spec_text()} - [적용]을 눌러야 제어를 시작합니다", it)

    def on_edit(self) -> None:
        sel = self.selected()
        if len(sel) != 1:
            messagebox.showinfo("수정", "수정할 항목을 하나 선택하세요.", parent=self.root)
            return
        old = sel[0]
        d = ItemDialog(self.root, "항목 수정", old)
        self.root.wait_window(d)
        new = d.result
        if new is None:
            return
        if self._find(new.ip, new.uid, skip=old):
            messagebox.showerror("수정", f"{new.name} 은 이미 목록에 있습니다.", parent=self.root)
            return
        same_fan = (new.ip, new.uid) == (old.ip, old.uid)
        was_active = old.active and same_fan
        old_spec = old.spec_text()
        idx = self.items.index(old)
        old.gen += 1                    # 진행 중인 작업 결과 무시
        if same_fan:                    # 같은 팬이면 제어모드 허가·시도 횟수 유지
            new.mode_allowed, new.mode_tries = old.mode_allowed, old.mode_tries
        self.items[idx] = new
        self._save()
        self.log(INFO, f"지정 변경: {old_spec} -> {new.spec_text()}", new)
        if old.active and not same_fan:
            self.log(WARN, f"제어 대상이 바뀌어 {old.name} 점검을 멈춤 (그 팬은 마지막 속도로 계속 운전)")
        self.refresh()
        if was_active:
            self.apply(new)

    def on_delete(self) -> None:
        sel = self.selected()
        if not sel:
            return
        names = ", ".join(it.name for it in sel)
        if not messagebox.askyesno("삭제", f"{names}\n\n목록에서 삭제할까요?\n"
                                   "(삭제해도 팬은 마지막으로 지정된 속도로 계속 운전합니다)", parent=self.root):
            return
        for it in sel:
            it.gen += 1
            it.active = False
            self.items.remove(it)
            self.log(INFO, "항목 삭제 (팬은 마지막 지정 속도로 계속 운전)", it)
        self._save()
        self.refresh()

    # ---------------------------------------------------------------- 제어
    def on_apply(self) -> None:
        sel = self.selected()
        if not sel:
            messagebox.showinfo("적용", "적용할 항목을 선택하세요.", parent=self.root)
            return
        for it in sel:
            self.apply(it)

    def apply(self, it: Item) -> None:
        it.gen += 1
        it.active = True
        it.target = None
        it.state, it.level = "적용 중...", INFO
        self.log(INFO, f"적용 시작: {it.spec_text()}", it)
        self.refresh()
        self._submit(it, first=True)

    def _submit(self, it: Item, first: bool) -> None:
        snap = {"ip": it.ip, "uid": it.uid, "mode": it.mode, "pct": it.pct, "ref_ip": it.ref_ip,
                "ref_uid": it.ref_uid, "first": first, "last_target": it.target,
                "mode_allowed": it.mode_allowed, "mode_tries": it.mode_tries}
        gen = it.gen
        it.busy = True
        self.worker.submit(control_job, snap, done=lambda res: self._done(it, gen, first, res))

    def _done(self, it: Item, gen: int, first: bool, res: dict) -> None:
        it.busy = False
        if gen != it.gen or it not in self.items:
            return                      # 그 사이 지정이 바뀌었거나 삭제됨
        it.mode_tries = res.get("mode_tries", it.mode_tries)
        for lv, msg in res.get("steps", []):
            self.log(lv, msg, it)
        if "ref_state" in res:
            it.ref_state = res["ref_state"]
            it.ref_rpm = res.get("ref_rpm", it.ref_rpm if res["ref_state"] == "정상" else None)
        if res.get("target") is not None:
            it.target = res["target"]
        if "h2" in res:
            it.h2 = res["h2"]
        it.rpm = res.get("rpm")
        it.checked = datetime.datetime.now().strftime("%H:%M:%S")
        if res.get("need_mode"):
            self.refresh()
            if self._ask_mode(it, res.get("h4", 0)):
                it.mode_allowed = True
                self.log(INFO, "제어모드 변경 허가 - 다시 적용합니다", it)
                self._submit(it, first=True)
            else:
                it.active = False
                it.state, it.level = "제어모드 변경 거부 - 제어 안 함", WARN
                self.log(WARN, "제어모드 변경을 거부해 제어하지 않음", it)
                self.refresh()
            return
        errors = [m for lv, m in res.get("steps", []) if lv == ERROR]
        if res.get("ok"):
            diff = abs((it.rpm or 0) - (it.target or 0))
            if errors:          # 참조팬 읽기 실패 등, 제어는 이전 값으로 계속
                it.state, it.level = errors[-1], ERROR
            elif diff > SPEED_TOLERANCE:
                it.state, it.level = f"속도 맞추는 중 (차이 {diff} RPM)", WARN
                if not first:
                    self.log(WARN, f"점검: 실제 {it.rpm} RPM, 지정 {it.target} RPM (차이 {diff})", it)
            else:
                # 정상 점검 결과는 목록에만 표시. 에러/주의 상태에서 돌아온 때만 한 번 기록
                if not first and it.level in (ERROR, WARN):
                    self.log(INFO, f"정상으로 돌아옴: 실제 {it.rpm} RPM / 지정 {it.target} RPM", it)
                it.state, it.level = "정상 제어 중", INFO
            if first:
                self.log(INFO, f"적용 완료: 지정 {it.target} RPM, 현재 실제 {it.rpm} RPM. "
                               f"{CHECK_INTERVAL_MS // 1000}초마다 점검합니다", it)
        else:
            it.state, it.level = (errors[-1] if errors else "제어 실패"), ERROR
            if first and it.target is None:
                self.log(ERROR, "적용 실패 - 원인을 확인한 뒤 다시 [적용]하세요. 점검은 계속 시도합니다", it)
        self.refresh()

    def _ask_mode(self, it: Item, h4: int) -> bool:
        return messagebox.askyesno(
            "제어모드 변경",
            f"{it.name} 의 제어모드(h4={_bus_mod.mode_text(h4)})를 '통신 속도제어 + 버스 Enable' 로 바꿔야 합니다.\n\n"
            "h4 는 전원을 꺼도 유지되는 값이며, 바뀐 뒤에는 팬이 아날로그(0-10V) 신호 대신\n"
            "이 프로그램이 지정한 속도로 돕니다. 바꿀까요?", parent=self.root)

    def on_stop(self) -> None:
        sel = self.selected()
        if not sel:
            return
        if not messagebox.askyesno("정지", "선택한 팬의 운전속도를 0 으로 지정하고 점검을 멈출까요?", parent=self.root):
            return
        for it in sel:
            it.gen += 1
            it.active = False
            it.state, it.level = "정지 요청 중...", INFO
            self.log(INFO, "정지: 운전속도 0 지정", it)
            gen = it.gen

            def done(res: dict, it=it, gen=gen) -> None:
                if gen != it.gen:
                    return
                for lv, msg in res["steps"]:
                    self.log(lv, msg, it)
                it.target = 0 if res["ok"] else it.target
                it.state, it.level = ("정지 (0 RPM 지정, 점검 안 함)", INFO) if res["ok"] else (
                    res["steps"][-1][1], ERROR)
                self.refresh()
            self.worker.submit(stop_job, {"ip": it.ip, "uid": it.uid}, done=done)
        self.refresh()

    def on_pause(self) -> None:
        for it in self.selected():
            if it.active:
                it.gen += 1
                it.active = False
                it.state, it.level = "점검 중지 (팬은 마지막 지정 속도로 운전)", INFO
                self.log(INFO, "점검 중지 - 팬은 마지막 지정 속도로 계속 운전", it)
        self.refresh()

    def _periodic(self) -> None:
        for it in self.items:
            if it.active and not it.busy:
                self._submit(it, first=False)
        self._check_after = self.root.after(CHECK_INTERVAL_MS, self._periodic)

    def _poll(self) -> None:
        try:
            while True:
                done, res = self.worker.results.get_nowait()
                try:
                    done(res)
                except Exception:
                    log.exception("결과 처리 오류")
        except queue.Empty:
            pass
        if not self._closing:
            self.root.after(100, self._poll)

    # ---------------------------------------------------------------- 종료
    def on_close(self) -> None:
        if any(it.active for it in self.items) and not messagebox.askyesno(
                "종료", "제어 중인 항목이 있습니다. 종료하면 점검을 멈추고 팬은 마지막 지정 속도로 계속 운전합니다.\n"
                "종료할까요?", parent=self.root):
            return
        self.shutdown()
        self.root.destroy()

    def shutdown(self) -> None:
        self._closing = True
        try:
            self.root.after_cancel(self._check_after)
        except (tk.TclError, ValueError):
            pass
        self.log(INFO, "종료")
        self.worker.stop()
        if self._file_log is not None:
            log.removeHandler(self._file_log)
            self._file_log.close()


def main() -> None:
    root = tk.Tk()
    try:
        ico = os.path.join(_BASE, "motor_monitor", "assets", "motor.ico")
        if os.path.exists(ico):
            root.iconbitmap(ico)
    except tk.TclError:
        pass
    MiniControlApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        os.makedirs(LOG_DIR, exist_ok=True)
        with open(os.path.join(LOG_DIR, "mini_control_crash.log"), "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now()}\n{traceback.format_exc()}\n")
        raise
