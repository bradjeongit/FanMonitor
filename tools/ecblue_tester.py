"""ECblue 기존모터 시험기 (작은 화면용).

정해진 통신장치·모터 2대(아래 HOST/PORT/MOTOR_IDS)를 직접 시험하고 모든 결과를 로그로 남깁니다.
모터마다: 접속 테스트 / 현재 속도 / 모니터링(2초마다, 지정 속도 유지) / 속도 지정 / 정지 / 원래대로.

실행: ECBLUE_TESTER.exe 또는 ecblue_tester.bat 더블클릭, 또는 python tools/ecblue_tester.py
로그: (exe 옆 또는 프로그램 폴더) logs/ecblue_tester_YYYYMMDD.log (모든 동작), logs/ecblue_monitor_YYYYMMDD.csv (모니터링 측정값)
운영 Config.dll 은 사용하지 않습니다.
"""
from __future__ import annotations

import datetime
import logging
import os
import queue
import socket
import sys
import threading
import time
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from motor_monitor import modbus  # noqa: E402

# ---- 시험 대상 (고정) ----
HOST = "172.27.0.95"
PORT = 502
MOTOR_IDS = (180, 176)

TIMEOUT = 2.0               # 응답 대기 (초)
REQUEST_GAP = 0.1           # 요청 사이 간격 (초)
MONITOR_INTERVAL_MS = 2000  # 모니터링 주기
MAX_RPM = 1100
MODE_WRITE_MAX_TRIES = 3    # h4(영구저장) 쓰기 시도 한도 (실행 1회당, 모터별)
# 로그 폴더: exe 로 실행하면 exe 옆, 파이썬으로 실행하면 프로그램 폴더(tools 의 상위) 아래 logs
_BASE = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
         else os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_DIR = os.path.join(_BASE, "logs")

I_ERROR = 12                # i12 오류 상태
I_CURRENT = 15              # i15 모터 전류 (0.01 A)

EXCEPTIONS = {1: "지원하지 않는 기능", 2: "없는 레지스터", 3: "허용되지 않는 값", 4: "장치 오류",
              6: "장치 사용 중", 10: "게이트웨이 경로 없음", 11: "게이트웨이: 모터 응답 없음"}
ERROR_BITS = {15: "통신 워치독", 14: "기동 오류", 13: "온도 오류", 12: "안전 정지", 11: "사인필터",
              10: "피크 전류", 9: "모터 구속", 8: "홀센서", 7: "TB", 6: "전원 결상/정전",
              5: "입력전압 낮음", 4: "입력전압 높음", 3: "DC링크 전압 낮음", 2: "DC링크 전압 높음",
              1: "지락", 0: "IGBT 오류"}

log = logging.getLogger("ecblue_tester")


def error_text(code: Optional[int]) -> str:
    if code is None:
        return "-"
    if code == 0:
        return "정상 (0)"
    names = [ERROR_BITS[b] for b in range(15, -1, -1) if code >> b & 1]
    return f"0x{code:04X} " + ", ".join(names)


def mode_text(h4: Optional[int]) -> str:
    if h4 is None:
        return "-"
    m = h4 & modbus.ECB_MODE_MASK
    kind = {0: "아날로그(E1)", 1: "통신속도"}.get(m, f"방식{m}")
    return f"0x{h4:04X} {kind}{'+버스EN' if h4 & modbus.ECB_BUSCON_ENABLE else ''}"


# ====================================================================== 통신
class Bus:
    """통신장치와 TCP 연결 1개. 작업 쓰레드에서만 사용."""

    def __init__(self, host: str, port: int) -> None:
        self.host, self.port = host, port
        self.sock: Optional[socket.socket] = None
        self.tid = 0

    @property
    def connected(self) -> bool:
        return self.sock is not None

    def connect(self) -> Tuple[bool, str]:
        self.close()
        t0 = time.monotonic()
        try:
            self.sock = socket.create_connection((self.host, self.port), timeout=3.0)
            self.sock.settimeout(TIMEOUT)
        except OSError as e:
            self.sock = None
            return False, f"TCP 접속 실패 {self.host}:{self.port} ({e})"
        return True, f"TCP 접속 성공 {self.host}:{self.port} ({(time.monotonic() - t0) * 1000:.0f} ms)"

    def close(self) -> None:
        if self.sock is not None:
            try:
                self.sock.close()
            except OSError:
                pass
        self.sock = None

    def request(self, frame_builder, *args) -> Tuple[Optional[modbus.Response], str]:
        """(응답, 설명). 연결이 없으면 먼저 접속."""
        if self.sock is None:
            ok, msg = self.connect()
            log.info(msg)
            if not ok:
                return None, "접속 안 됨"
        self.tid = (self.tid + 1) % 65536
        frame = frame_builder(self.tid, *args)
        t0 = time.monotonic()
        raw = b""
        try:
            self.sock.sendall(frame)
            splitter = modbus.FrameSplitter()
            while time.monotonic() - t0 < TIMEOUT:
                try:
                    data = self.sock.recv(513)
                except socket.timeout:
                    break
                if not data:
                    raise OSError("통신장치가 연결을 끊음")
                raw += data
                for f in splitter.feed(data):
                    r = modbus.parse_response(f)
                    if r is not None and r.tid == self.tid:
                        ms = (time.monotonic() - t0) * 1000
                        time.sleep(REQUEST_GAP)
                        if r.exception is not None:
                            return r, f"예외 {r.exception}: {EXCEPTIONS.get(r.exception, '알 수 없음')} ({ms:.0f} ms)"
                        return r, f"정상 ({ms:.0f} ms)"
        except OSError as e:
            self.close()
            return None, f"통신 오류 ({e})"
        time.sleep(REQUEST_GAP)
        return None, f"응답 없음 ({TIMEOUT:g}초" + (f", 다른 데이터: {raw[:24].hex(' ')})" if raw else ")")


class Worker(threading.Thread):
    """통신 작업을 한 번에 하나씩 처리. 결과는 UI 쪽 큐로 돌려줌."""

    def __init__(self, bus: Bus) -> None:
        super().__init__(daemon=True, name="ecblue-worker")
        self.bus = bus
        self.jobs: "queue.Queue[Optional[tuple]]" = queue.Queue()
        self.results: "queue.Queue[tuple]" = queue.Queue()

    def submit(self, fn: Callable, *args, done: Optional[Callable] = None) -> None:
        self.jobs.put((fn, args, done))

    def run(self) -> None:
        while True:
            job = self.jobs.get()
            if job is None:
                break
            fn, args, done = job
            try:
                res = fn(self.bus, *args)
            except Exception as e:  # 작업 오류로 쓰레드가 죽지 않도록
                log.exception("작업 오류")
                res = {"error": str(e)}
            self.results.put((done, res))
        self.bus.close()

    def stop(self) -> None:
        self.jobs.put(None)


# ====================================================================== 작업 (작업 쓰레드에서 실행)
def read_all(bus: Bus, uid: int) -> dict:
    """h1~h4, i14(+i15), i12 읽기"""
    st: dict = {"uid": uid, "time": datetime.datetime.now()}
    r, d = bus.request(modbus.build_read_holding, uid, modbus.ECB_H_CONTROL, 4)
    st["hold_msg"] = d
    if r is not None and r.exception is None and len(r.values) >= 4:
        st["h1"], st["h2"], st["h3"], st["h4"] = r.values[:4]
    r, d = bus.request(modbus.build_read_input, uid, modbus.ECB_I_SPEED, 2)
    if r is not None and r.exception is not None:          # i15 를 지원하지 않으면 i14 만
        r, d = bus.request(modbus.build_read_input, uid, modbus.ECB_I_SPEED, 1)
    st["speed_msg"] = d
    if r is not None and r.exception is None and r.values:
        st["rpm"] = r.values[0]
        if len(r.values) > 1:
            st["current"] = r.values[1] / 100
    if "h1" in st or "rpm" in st:
        r, _d = bus.request(modbus.build_read_input, uid, I_ERROR, 1)
        if r is not None and r.exception is None and r.values:
            st["err"] = r.values[0]
    st["ok"] = "h1" in st or "rpm" in st
    return st


def job_connect_test(bus: Bus, uid: int) -> dict:
    ok, msg = bus.connect()
    st = {"uid": uid, "connect_msg": msg, "ok": False}
    if ok:
        st = read_all(bus, uid)
        st["connect_msg"] = msg
    return st


def job_read_speed(bus: Bus, uid: int) -> dict:
    r, d = bus.request(modbus.build_read_input, uid, modbus.ECB_I_SPEED, 1)
    st = {"uid": uid, "speed_msg": d, "time": datetime.datetime.now()}
    if r is not None and r.exception is None and r.values:
        st["rpm"] = r.values[0]
    st["ok"] = "rpm" in st
    return st


def _write(bus: Bus, uid: int, reg: int, value: int, steps: List[str]) -> bool:
    r, d = bus.request(modbus.build_new_write_single, uid, reg, value)
    ok = r is not None and r.exception is None
    steps.append(f"h{reg}={value} 쓰기: {d}")
    return ok


def job_set_speed(bus: Bus, uid: int, rpm: int, allow_mode: bool, mode_tries: int) -> dict:
    """속도 지정: (필요 시) PIN -> h4 제어모드, h1 Enable, h2 운전속도 -> 다시 읽어 확인"""
    steps: List[str] = []
    st = read_all(bus, uid)
    res = {"uid": uid, "steps": steps, "rpm_target": rpm, "mode_written": False}
    if "h4" not in st:
        steps.append(f"현재 설정 읽기 실패: {st['hold_msg']}")
        res["ok"] = False
        return res
    if not modbus.ecb_mode_ok(st["h4"]):
        if not allow_mode:
            res.update(ok=False, need_mode=True, h4=st["h4"])
            return res
        if mode_tries >= MODE_WRITE_MAX_TRIES:
            steps.append(f"제어모드(h4) 쓰기 {MODE_WRITE_MAX_TRIES}회 실패 - 더 시도하지 않음 (PIN 보호/통신 확인)")
            res["ok"] = False
            return res
        res["mode_written"] = True
        _write(bus, uid, modbus.ECB_H_PIN, modbus.ECB_PIN_LEVEL2, steps)
        if not _write(bus, uid, modbus.ECB_H_CONTROLMODE, modbus.ecb_mode_value(st["h4"]), steps):
            res["ok"] = False
            return res
    if not st["h1"] & modbus.ECB_ENABLE_BIT:
        _write(bus, uid, modbus.ECB_H_CONTROL, st["h1"] | modbus.ECB_ENABLE_BIT, steps)
    _write(bus, uid, modbus.ECB_H_SPEED, rpm, steps)
    after = read_all(bus, uid)
    res["after"] = after
    res["ok"] = (after.get("h2") == rpm and modbus.ecb_mode_ok(after.get("h4", 0))
                 and bool(after.get("h1", 0) & modbus.ECB_ENABLE_BIT))
    steps.append(f"확인: h1={after.get('h1')} h2={after.get('h2')} h4={mode_text(after.get('h4'))} "
                 f"실제 {after.get('rpm', '-')} RPM")
    return res


def job_write_speed_only(bus: Bus, uid: int, rpm: int, still_valid: Callable[[], bool] = lambda: True) -> dict:
    """h2 만 쓰기. still_valid 가 False 면(대기 중 사용자가 정지/원래대로/새 속도를 지정) 쓰지 않음."""
    steps: List[str] = []
    if not still_valid():
        return {"uid": uid, "steps": ["취소됨 (그 사이 다른 지정이 있었음)"], "ok": True, "skipped": True}
    ok = _write(bus, uid, modbus.ECB_H_SPEED, rpm, steps)
    return {"uid": uid, "steps": steps, "ok": ok}


def job_restore(bus: Bus, uid: int, orig: dict) -> dict:
    """처음 읽은 값으로 복원: h2 -> h1 -> PIN -> h4"""
    steps: List[str] = []
    cur = read_all(bus, uid)
    if "h4" not in cur:
        return {"uid": uid, "steps": [f"현재 설정 읽기 실패: {cur['hold_msg']}"], "ok": False}
    for reg, key in ((modbus.ECB_H_SPEED, "h2"), (modbus.ECB_H_CONTROL, "h1")):
        if cur[key] != orig[key]:
            _write(bus, uid, reg, orig[key], steps)
    if cur["h4"] != orig["h4"]:
        _write(bus, uid, modbus.ECB_H_PIN, modbus.ECB_PIN_LEVEL2, steps)
        _write(bus, uid, modbus.ECB_H_CONTROLMODE, orig["h4"], steps)
    after = read_all(bus, uid)
    ok = all(after.get(k) == orig[k] for k in ("h1", "h2", "h4"))
    steps.append(f"확인: h1={after.get('h1')} h2={after.get('h2')} h4={mode_text(after.get('h4'))} "
                 f"실제 {after.get('rpm', '-')} RPM")
    return {"uid": uid, "steps": steps if steps else ["바뀐 값 없음"], "ok": ok, "after": after}


# ====================================================================== 화면
class MotorPanel(ttk.LabelFrame):
    def __init__(self, master, app: "TesterApp", uid: int) -> None:
        super().__init__(master, text=f" 모터 ID {uid} ", padding=4)
        self.app, self.uid = app, uid
        self.orig: Optional[dict] = None       # 이번 실행에서 처음 읽은 h1/h2/h4 (원래대로 기준)
        self.target: Optional[int] = None      # 지정 속도 (유지 대상)
        self.gen = 0                           # 사용자 지정(속도/정지/원래대로)마다 증가 -> 대기 중인 유지 쓰기 무효화
        self.keep_pending = False
        self.monitoring = False
        self.busy_monitor = False
        self.mode_tries = 0
        self.last_ok: Optional[bool] = None

        top = ttk.Frame(self)
        top.pack(fill="x")
        self.dot = tk.Label(top, text="●", fg="#808080", font=("TkDefaultFont", 12))
        self.dot.pack(side="left")
        self.lbl_state = ttk.Label(top, text="시험 전")
        self.lbl_state.pack(side="left", padx=(2, 0))

        self.lbl_rpm = tk.Label(self, text="- RPM", font=("TkDefaultFont", 16, "bold"), anchor="w")
        self.lbl_rpm.pack(fill="x")
        self.info = ttk.Label(self, text="\n\n", justify="left", font=("TkDefaultFont", 8))
        self.info.pack(fill="x")
        self.bind("<Configure>", lambda e: self.info.configure(wraplength=max(e.width - 12, 100)))

        b1 = ttk.Frame(self)
        b1.pack(fill="x", pady=(4, 0))
        self.btn_conn = ttk.Button(b1, text="접속 테스트", command=self.on_connect_test)
        self.btn_speed = ttk.Button(b1, text="현재 속도", command=self.on_read_speed)
        self.btn_conn.pack(side="left", fill="x", expand=True)
        self.btn_speed.pack(side="left", fill="x", expand=True, padx=(2, 0))
        self.btn_mon = ttk.Button(self, text="모니터링 시작", command=self.on_toggle_monitor)
        self.btn_mon.pack(fill="x", pady=(2, 0))

        b2 = ttk.Frame(self)
        b2.pack(fill="x", pady=(4, 0))
        ttk.Label(b2, text="속도").pack(side="left")
        self.ent_rpm = ttk.Spinbox(b2, from_=0, to=MAX_RPM, increment=50, width=5)
        self.ent_rpm.pack(side="left", padx=2)
        ttk.Label(b2, text="RPM").pack(side="left")
        self.btn_set = ttk.Button(b2, text="지정", width=5, command=self.on_set_speed)
        self.btn_set.pack(side="left", fill="x", expand=True, padx=(4, 0))
        b3 = ttk.Frame(self)
        b3.pack(fill="x", pady=(2, 0))
        self.btn_stop = ttk.Button(b3, text="정지(0)", command=self.on_stop)
        self.btn_restore = ttk.Button(b3, text="원래대로", command=self.on_restore)
        self.btn_stop.pack(side="left", fill="x", expand=True)
        self.btn_restore.pack(side="left", fill="x", expand=True, padx=(2, 0))
        self.keep = tk.BooleanVar(value=True)
        ttk.Checkbutton(self, text="모니터링 중 지정 속도 유지", variable=self.keep).pack(anchor="w", pady=(2, 0))
        self.lbl_target = ttk.Label(self, text="지정 속도: 없음", font=("TkDefaultFont", 8))
        self.lbl_target.pack(anchor="w")

    def L(self, msg: str, level=logging.INFO) -> None:
        self.app.log(f"[ID{self.uid}] {msg}", level)

    # ---- 표시 ----
    def show(self, st: dict) -> None:
        ok = st.get("ok", False)
        self.dot.configure(fg="#27ae60" if ok else "#c0392b")
        when = st.get("time", datetime.datetime.now()).strftime("%H:%M:%S")
        self.lbl_state.configure(text=f"{'응답 정상' if ok else '응답 없음'}  {when}")
        if "rpm" in st:
            self.lbl_rpm.configure(text=f"{st['rpm']} RPM")
        elif not ok:
            self.lbl_rpm.configure(text="- RPM")
        if "h4" in st:
            cur = f"   전류 {st['current']:.2f} A" if "current" in st else ""
            self.info.configure(text=(f"운전 h1: {'Enable' if st['h1'] & 1 else '꺼짐'}   지정 h2: {st['h2']}{cur}\n"
                                      f"제어모드 h4: {mode_text(st['h4'])}\n"
                                      f"오류 i12: {error_text(st.get('err'))}"))
            if self.orig is None:
                self.orig = {k: st[k] for k in ("h1", "h2", "h4")}
                self.L(f"원래 설정 기억: h1={st['h1']} h2={st['h2']} h4={mode_text(st['h4'])}")
        if self.last_ok is not None and ok != self.last_ok:
            self.L("응답이 다시 정상" if ok else "응답 끊김", logging.INFO if ok else logging.WARNING)
        self.last_ok = ok

    def summary(self, st: dict) -> str:
        parts = []
        if "h4" in st:
            parts.append(f"h1={st['h1']} h2={st['h2']} h3=0x{st['h3']:04X} h4={mode_text(st['h4'])}")
        else:
            parts.append(f"h1~h4 {st.get('hold_msg', '-')}")
        parts.append(f"실제 {st['rpm']} RPM" if "rpm" in st else f"i14 {st.get('speed_msg', '-')}")
        if "current" in st:
            parts.append(f"{st['current']:.2f} A")
        if "err" in st:
            parts.append(f"오류 {error_text(st['err'])}")
        return " | ".join(parts)

    # ---- 버튼 ----
    def on_connect_test(self) -> None:
        self.L("접속 테스트 시작")
        self.app.worker.submit(job_connect_test, self.uid, done=self._done_connect)

    def _done_connect(self, st: dict) -> None:
        self.L(st.get("connect_msg", ""))
        if "hold_msg" in st:
            self.L(f"h1~h4 읽기: {st['hold_msg']}")
            self.L(f"실제 회전수 읽기: {st['speed_msg']}")
            self.L("접속 테스트 결과: " + ("정상 - " if st["ok"] else "모터 응답 없음 - ") + self.summary(st),
                   logging.INFO if st["ok"] else logging.WARNING)
        self.show(st)

    def on_read_speed(self) -> None:
        self.app.worker.submit(job_read_speed, self.uid, done=self._done_speed)

    def _done_speed(self, st: dict) -> None:
        self.L(f"현재 속도: {st['rpm']} RPM" if st["ok"] else f"현재 속도 읽기 실패: {st['speed_msg']}",
               logging.INFO if st["ok"] else logging.WARNING)
        self.show(st)

    def on_toggle_monitor(self) -> None:
        self.set_monitoring(not self.monitoring)

    def set_monitoring(self, on: bool) -> None:
        if on == self.monitoring:
            return
        self.monitoring = on
        self.btn_mon.configure(text="모니터링 중지" if on else "모니터링 시작")
        self.L(f"모니터링 {'시작' if on else '중지'} ({MONITOR_INTERVAL_MS / 1000:g}초 주기, CSV 기록)")

    def monitor_tick(self) -> None:
        if self.monitoring and not self.busy_monitor:
            self.busy_monitor = True
            self.app.worker.submit(read_all, self.uid, done=self._done_monitor)

    def _done_monitor(self, st: dict) -> None:
        self.busy_monitor = False
        self.show(st)
        self.app.csv_row(st, self.target)
        if not self.monitoring:
            return
        # 지정 속도 유지: 모터 쪽 h2 가 바뀌어 있으면(정전 후 재시작 등) 다시 지정
        if (self.keep.get() and self.target is not None and "h2" in st and st["h2"] != self.target
                and modbus.ecb_mode_ok(st.get("h4", 0)) and not self.keep_pending):
            self.L(f"지정 속도 유지: h2={st['h2']} -> {self.target} 다시 지정", logging.WARNING)
            gen = self.gen
            self.keep_pending = True

            def done(res: dict) -> None:
                self.keep_pending = False
                if not res.get("skipped"):
                    self._done_steps("속도 재지정")(res)
            self.app.worker.submit(job_write_speed_only, self.uid, self.target, lambda: self.gen == gen, done=done)

    def _rpm_value(self) -> Optional[int]:
        try:
            v = int(self.ent_rpm.get().strip())
        except ValueError:
            v = -1
        if not 0 <= v <= MAX_RPM:
            messagebox.showerror("속도 지정", f"속도는 0~{MAX_RPM} 사이 숫자로 입력하세요.", parent=self.app.root)
            return None
        return v

    def on_set_speed(self, allow_mode: bool = False) -> None:
        rpm = self._rpm_value()
        if rpm is None:
            return
        if not allow_mode and not messagebox.askyesno(
                "속도 지정", f"모터 ID {self.uid} 를 {rpm} RPM 으로 운전합니다.\n계속할까요?", parent=self.app.root):
            return
        self.gen += 1
        self.L(f"속도 지정 {rpm} RPM 시작")
        self.app.worker.submit(job_set_speed, self.uid, rpm, allow_mode, self.mode_tries, done=self._done_set)

    def _done_set(self, res: dict) -> None:
        if res.get("need_mode"):
            self.L(f"제어모드가 통신 속도제어가 아님 (h4={mode_text(res['h4'])})", logging.WARNING)
            if messagebox.askyesno(
                    "제어모드 변경",
                    f"모터 ID {self.uid} 의 제어모드(h4)를 '통신 속도제어 + 버스 Enable' 로 바꿔야 합니다.\n\n"
                    "h4 는 전원을 꺼도 유지되는 값이며, 바뀐 동안 모터는 아날로그(0-10V) 신호 대신\n"
                    "통신으로 받은 속도로 돕니다. [원래대로] 버튼으로 되돌릴 수 있습니다.\n\n바꿀까요?",
                    icon="warning", parent=self.app.root):
                self.on_set_speed(allow_mode=True)
            else:
                self.L("제어모드 변경 취소 - 속도 지정 안 함")
            return
        if res.get("mode_written"):
            self.mode_tries += 1
        for s in res["steps"]:
            self.L(f"  {s}")
        if res.get("ok"):
            self.target = res["rpm_target"]
            self.lbl_target.configure(text=f"지정 속도: {self.target} RPM")
            self.L(f"속도 지정 완료: {self.target} RPM (실제 회전수는 가속 후 따라옴)")
            self.set_monitoring(True)
        else:
            self.L("속도 지정 실패", logging.ERROR)
        if "after" in res:
            self.show(res["after"])

    def on_stop(self) -> None:
        if not messagebox.askyesno("정지", f"모터 ID {self.uid} 의 운전속도를 0 으로 지정합니다.\n계속할까요?",
                                   parent=self.app.root):
            return
        self.gen += 1
        self.target = 0
        self.lbl_target.configure(text="지정 속도: 0 RPM (정지)")
        self.L("정지(운전속도 0) 지정")
        self.app.worker.submit(job_write_speed_only, self.uid, 0, done=self._done_steps("정지"))

    def on_restore(self) -> None:
        if self.orig is None:
            messagebox.showinfo("원래대로", "아직 원래 설정을 읽지 못했습니다. 먼저 [접속 테스트] 를 하세요.",
                                parent=self.app.root)
            return
        o = self.orig
        if not messagebox.askyesno(
                "원래대로", f"모터 ID {self.uid} 를 처음 읽은 설정으로 되돌립니다.\n"
                f"h1={o['h1']}, h2={o['h2']}, h4={mode_text(o['h4'])}\n계속할까요?", parent=self.app.root):
            return
        self.gen += 1
        self.target = None
        self.lbl_target.configure(text="지정 속도: 없음")
        self.L("원래대로 되돌리기 시작")
        self.app.worker.submit(job_restore, self.uid, dict(o), done=self._done_restore)

    def _done_restore(self, res: dict) -> None:
        for s in res["steps"]:
            self.L(f"  {s}")
        self.L("원래대로 완료" if res["ok"] else "원래대로 확인 안 됨",
               logging.INFO if res["ok"] else logging.ERROR)
        if "after" in res:
            self.show(res["after"])

    def _done_steps(self, title: str) -> Callable[[dict], None]:
        def done(res: dict) -> None:
            for s in res["steps"]:
                self.L(f"  {s}")
            self.L(f"{title} {'완료' if res['ok'] else '실패'}", logging.INFO if res["ok"] else logging.ERROR)
        return done


class TesterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        os.makedirs(LOG_DIR, exist_ok=True)
        day = datetime.date.today().strftime("%Y%m%d")
        self.log_path = os.path.join(LOG_DIR, f"ecblue_tester_{day}.log")
        self.csv_path = os.path.join(LOG_DIR, f"ecblue_monitor_{day}.csv")
        self._fh = logging.FileHandler(self.log_path, encoding="utf-8")
        self._fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        log.addHandler(self._fh)
        log.setLevel(logging.INFO)

        root.title(f"ECblue 기존모터 시험기 - {HOST}:{PORT}")
        root.geometry("620x520")
        root.minsize(460, 470)
        head = ttk.Frame(root, padding=(6, 4))
        head.pack(fill="x")
        self.lbl_dev = ttk.Label(head, text=f"통신장치 {HOST}:{PORT}   모터 ID {', '.join(map(str, MOTOR_IDS))}",
                                 font=("TkDefaultFont", 9, "bold"))
        self.lbl_dev.pack(side="left")
        self.btn_all = ttk.Button(head, text="모두 모니터링", command=self.toggle_all)
        self.btn_all.pack(side="right")

        body = ttk.Frame(root, padding=(4, 0))
        body.pack(fill="x")
        self.panels: Dict[int, MotorPanel] = {}
        for i, uid in enumerate(MOTOR_IDS):
            body.columnconfigure(i, weight=1, uniform="m")
            p = MotorPanel(body, self, uid)
            p.grid(row=0, column=i, sticky="nsew", padx=2)
            self.panels[uid] = p

        self.log_frame = lf = ttk.LabelFrame(root, text=f" 로그 (logs 폴더 {os.path.basename(self.log_path)}) ", padding=2)
        lf.pack(fill="both", expand=True, padx=6, pady=4)
        self.txt = tk.Text(lf, height=5, font=("TkFixedFont", 8), wrap="none", state="disabled")
        sb = ttk.Scrollbar(lf, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt.pack(fill="both", expand=True)
        self.txt.tag_configure("WARNING", foreground="#d35400")
        self.txt.tag_configure("ERROR", foreground="#c0392b")

        self.worker = Worker(Bus(HOST, PORT))
        self.worker.start()
        self.log(f"시험기 시작 - 대상 {HOST}:{PORT}, 모터 ID {MOTOR_IDS}")
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.after(100, self._poll)
        root.after(MONITOR_INTERVAL_MS, self._monitor)

    def log(self, msg: str, level=logging.INFO) -> None:
        log.log(level, msg)
        line = f"{datetime.datetime.now():%H:%M:%S} {msg}\n"
        self.txt.configure(state="normal")
        self.txt.insert("end", line, logging.getLevelName(level))
        if int(self.txt.index("end-1c").split(".")[0]) > 2000:
            self.txt.delete("1.0", "500.0")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def csv_row(self, st: dict, target: Optional[int]) -> None:
        new = not os.path.exists(self.csv_path)
        try:
            with open(self.csv_path, "a", encoding="utf-8-sig") as f:
                if new:
                    f.write("시각,ID,응답,실제RPM,전류A,h1,h2(지정),h4,오류i12,유지목표RPM\n")
                f.write(",".join(str(v) for v in (
                    st["time"].strftime("%Y-%m-%d %H:%M:%S"), st["uid"], 1 if st["ok"] else 0,
                    st.get("rpm", ""), f"{st['current']:.2f}" if "current" in st else "",
                    st.get("h1", ""), st.get("h2", ""), st.get("h4", ""), st.get("err", ""),
                    "" if target is None else target)) + "\n")
        except OSError as e:
            log.warning("CSV 기록 실패: %s", e)

    def toggle_all(self) -> None:
        on = not all(p.monitoring for p in self.panels.values())
        for p in self.panels.values():
            p.set_monitoring(on)

    def _poll(self) -> None:
        try:
            while True:
                done, res = self.worker.results.get_nowait()
                if "error" in res:
                    self.log(f"작업 오류: {res['error']}", logging.ERROR)
                    for p in self.panels.values():
                        p.busy_monitor = False
                elif done is not None:
                    done(res)
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _monitor(self) -> None:
        for p in self.panels.values():
            p.monitor_tick()
        self.root.after(MONITOR_INTERVAL_MS, self._monitor)

    def on_close(self) -> None:
        if not messagebox.askyesno("종료", "시험기를 종료할까요?\n모터는 마지막으로 지정한 상태로 계속 운전합니다.",
                                   parent=self.root):
            return
        self.shutdown()
        self.root.destroy()

    def shutdown(self) -> None:
        self.log("시험기 종료 - 모터 설정은 바꾸지 않음")
        self.worker.stop()
        self.worker.join(timeout=5)
        log.removeHandler(self._fh)
        self._fh.close()


def main() -> None:
    try:
        root = tk.Tk()
        TesterApp(root)
        root.mainloop()
    except Exception:
        # 콘솔 없는 exe 에서도 원인을 알 수 있도록 파일과 메시지 창으로 알림
        import traceback
        os.makedirs(LOG_DIR, exist_ok=True)
        path = os.path.join(LOG_DIR, "ecblue_tester_error.txt")
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now()}\n{traceback.format_exc()}\n")
        try:
            messagebox.showerror("ECblue 시험기", f"시작 중 오류가 발생했습니다.\n{path}")
        except tk.TclError:
            pass
        raise


if __name__ == "__main__":
    main()
