"""모터 스캐너 - 통신장치(IP)에 연결된 모터 ID 를 찾고, ECblue 팬의 설정을 확인·변경.

1. IP·PORT 와 ID 범위(기본 1~240)를 정해 [스캔 시작] -> 응답하는 ID 를 목록으로 표시
2. 제품코드(i1)로 ECblue 인지 판별. ECblue 는 [설정 열기]로 상태·설정(최고/최저 속도 등)을 확인·변경
3. "통신 끊김 시 자동 운전"(통신 워치독 h17): 정해진 시간 동안 통신이 없으면 지정 속도로 스스로 운전. 켜기/끄기/시험

근거: ZIEHL-ABEGG "ECblue MODBUS Description" L-BAL-E076-GB 1517 Index 013 (ECblue FW 1.03 ~ 14)
  h5 Set Intern1, h7 최저속도, h8 최고속도 (1/min, 영구저장, PIN 1단계)
  h17 통신 워치독: bit15~8 시간(초, 0=꺼짐), bit7~0 동작(2 = 통신 끊김 시 정속 1). 영구저장, PIN 2단계
  i1 제품코드: 상위 바이트 0x01 = ECblue 계열 (구형 ECblue Basic = 0x0032)
  문서의 "정속 1(constant speed 1)" 은 Set Intern1(h5) 로 판단함 -> [시험] 으로 현장에서 확인

실행: MOTOR_SCANNER.exe 또는 motor_scanner.bat 더블클릭, 또는 python tools/motor_scanner.py
로그: logs/motor_scanner_YYYYMMDD.log, 마지막 입력값: motor_scanner.json (exe/프로그램 폴더)
"""
from __future__ import annotations

import datetime
import ipaddress
import json
import logging
import os
import queue
import socket
import sys
import threading
import time
import traceback
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from motor_monitor import modbus  # noqa: E402

try:
    import tkinter as tk
    from tkinter import messagebox, ttk
except ImportError:  # pragma: no cover
    tk = None  # type: ignore

TITLE = "모터 스캐너"
DEFAULT_PORT = 502
ID_MIN, ID_MAX = 1, 247
DEFAULT_RANGE = (1, 240)
DEFAULT_TIMEOUT = 0.5      # 응답 대기 (초)
CONNECT_TIMEOUT = 3.0
REQUEST_GAP = 0.03         # 요청 사이 간격 (초)
WRITE_VERIFY_GAP = 0.2     # 쓰기 후 확인 읽기 전 대기
REFRESH_MS = 2000          # 상세 창 자동 갱신

_BASE = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
         else os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_DIR = os.path.join(_BASE, "logs")
SETTINGS_FILE = os.path.join(_BASE, "motor_scanner.json")

log = logging.getLogger("motor_scanner")

GATEWAY_NO_DEVICE = (0x0A, 0x0B)   # 게이트웨이 경로 없음 / 대상 장치 응답 없음
EXCEPTIONS = {1: "지원하지 않는 기능", 2: "없는 레지스터", 3: "허용되지 않는 값", 4: "장치 오류(쓰기 거부)",
              6: "장치 사용 중", 10: "게이트웨이 경로 없음", 11: "게이트웨이: 장치 응답 없음"}

# ====================================================================== ECblue 레지스터 정보 (L-BAL-E076-GB)
PIN_LEVEL1, PIN_LEVEL2 = 1234, 10      # h0 에 써서 PIN 보호 1단계 / 2단계 해제 (약 15분 유지)

PRODUCTS = {
    0x0032: "ECblue Basic Standard (구형)", 0x0100: "ECblue Basic Standard", 0x0101: "ECblue MODBUS 02 (EC116)",
    0x0102: "ECblue Lite (EC116)", 0x0103: "ECblue (EC90)",
    0x0232: "FU Basic Icontrol (구형)", 0x0200: "FU Basic Icontrol", 0x0201: "FU Basic Fcontrol",
    0x0300: "PMcontrol Basic", 0x0401: "AM-MODBUS", 0x0402: "AM-MODBUS-W", 0x0403: "AM-PREMIUM",
    0x0404: "AM-PREMIUM-W", 0x0405: "AM-CONFIG", 0x0407: "AM-UNICON-IO01", 0x0500: "UNIcon (MCRE17)",
    0x0600: "A-G-247NW AZUN27", 0x0601: "A-G-247NW AZUN30", 0x0701: "B-G-028NE (MODBUS IO)",
}

KIND_ECBLUE = "ECblue"
KIND_AM = "ECblue(AM 모듈)"
KIND_ZA = "ZIEHL-ABEGG 기타"
KIND_NEW = "신규모터(추정)"
KIND_OTHER = "기타 장치"


def classify(product: int) -> Tuple[str, str]:
    """제품코드(i1) -> (종류, 제품 이름)"""
    name = PRODUCTS.get(product, f"제품코드 0x{product:04X}")
    if product == 0x0032 or product >> 8 == 0x01:
        return KIND_ECBLUE, name
    if product in (0x0401, 0x0402, 0x0403, 0x0404):
        return KIND_AM, name          # ECblue 팬에 붙는 통신 모듈 (같은 레지스터 사용)
    return KIND_ZA, name


def is_ecblue(kind: str) -> bool:
    return kind in (KIND_ECBLUE, KIND_AM)


def fw_text(fw: Optional[int]) -> str:
    return "-" if fw is None else f"{fw // 100}.{fw % 100:02d}"


WD_MODES = {
    0: "꺼짐",
    1: "경보만 (K1 릴레이)",
    2: "정속 운전 (Set Intern1)",
    3: "경보 + 정속 운전",
    4: "E1 신호 이상 시 경보",
    5: "E1 신호 이상 시 정속 운전",
    6: "E1 신호 이상 시 경보 + 정속",
    7: "E1(아날로그) 제어로 전환 (FW 13.30 이상)",
    8: "경보 + E1 제어로 전환 (FW 13.30 이상)",
}
WD_CHOICES = (2, 3, 7, 8)        # 화면에서 고를 수 있는 '통신 끊김 시' 동작
WD_SPEED_MODES = (2, 3, 5, 6)    # Set Intern1 속도로 도는 동작
WD_FW_E1 = 1330                  # 모드 7, 8 은 FW 13.30 부터


def wd_decode(h17: int) -> Tuple[int, int]:
    """h17 -> (시간 초, 동작)"""
    return (h17 >> 8) & 0xFF, h17 & 0xFF


def wd_encode(seconds: int, mode: int) -> int:
    if not 0 <= seconds <= 255 or not 0 <= mode <= 255:
        raise ValueError("워치독 시간 0~255초, 동작 0~255")
    return (seconds << 8) | mode


def wd_text(h17: Optional[int], set1: Optional[int] = None) -> str:
    if h17 is None:
        return "-"
    sec, mode = wd_decode(h17)
    if sec == 0 or mode == 0:
        return "꺼짐"
    s = f"켜짐: {sec}초 끊기면 {WD_MODES.get(mode, f'동작 {mode}')}"
    if mode in WD_SPEED_MODES and set1 is not None:
        s += f" {set1} rpm"
    return s


@dataclass(frozen=True)
class Param:
    addr: int
    name: str
    unit: str = ""
    lo: int = 0
    hi: int = 65535
    pin: int = 0
    nv: bool = False
    editable: bool = False
    note: str = ""


# 변경 가능한 설정 (화면에서 값 입력)
EDIT_PARAMS = [
    Param(2, "운전속도 (통신 지정)", "rpm", 0, 6000, 0, False, True, "제어모드 1(통신 속도)일 때"),
    Param(8, "최고속도 (Max. Speed)", "rpm", 0, 6000, 1, True, True, "이보다 빠르게 안 돎 (0=제한 없음)"),
    Param(7, "최저속도 (Min. Speed)", "rpm", 0, 6000, 1, True, True, "제어모드 2·3 에서 하한"),
    Param(5, "Set Intern1 (정속 1)", "rpm", 0, 6000, 1, True, True, "통신 끊김 시 자동 운전 속도"),
    Param(6, "Set Intern2", "rpm", 0, 6000, 1, True, True, "제어모드 5 / h1 bit1"),
    Param(9, "Set Intern3", "rpm", 0, 6000, 1, True, True, "제어모드 6 / h1 bit2"),
    Param(18, "속도 제한 (Limit)", "%", 0, 100, 2, True, True, "h1 bit3 로 켰을 때"),
    Param(11, "E1 최소 (E1 Min)", "%", 0, 100, 2, True, True, "아날로그 시작점 (기본 5)"),
    Param(12, "E1 최대 (E1 Max)", "%", 0, 100, 2, True, True, "아날로그 끝점 (기본 100)"),
]
# 읽기 전용으로 보여주는 설정 (바꾸면 통신·운전 방식이 바뀌어 이 프로그램에서는 바꾸지 않음)
VIEW_PARAMS = [
    Param(1, "운전 제어 (Control)", note="bit0 운전허가, bit1 SetIntern2, bit2 SetIntern3, bit3 제한, bit4 역회전"),
    Param(3, "통신 설정 (COM)", note="주소·속도·패리티. 바꾸면 통신이 끊길 수 있음"),
    Param(4, "제어 모드 (Controlmode)"),
    Param(10, "입력 반전 (Inverting)"),
    Param(13, "E1 기능"), Param(14, "D1 기능"), Param(15, "K1 릴레이 기능"),
    Param(16, "설정 플래그 (Setup Flags)", note="bit9~8 기본 PIN 보호 단계"),
    Param(25, "가감속 시간 (Ramp)", note="공장 설정 - 바꾸지 말 것"),
    Param(30, "속도 회피 구간 사용"),
]
PARAM_BY_ADDR = {p.addr: p for p in EDIT_PARAMS + VIEW_PARAMS}

CONTROL_MODES = {0: "E1 아날로그(0-10V/PWM)", 1: "통신 속도 h2 (rpm)", 2: "통신 h2 (0-32767 = 0-100%)",
                 3: "통신 h2 (0-100%)", 4: "Set Intern1", 5: "Set Intern2", 6: "Set Intern2(문서 표기)"}
OP_BITS = {15: "Fan Bad", 14: "역회전 중", 13: "내부 온도 경보", 12: "IGBT 온도 경보", 11: "DC 과전압",
           10: "K1 상태", 9: "E1 디지털", 8: "D1 상태", 7: "전류 제한", 6: "약계자", 5: "화재 경보",
           4: "회전 방향 이상", 3: "내부 시스템 오류", 2: "IGBT 점검", 1: "온도 관리 중", 0: "정지(STOP)"}
ERR_BITS = {15: "통신 워치독(COM)", 14: "기동 오류", 13: "온도 오류", 12: "안전 정지", 11: "사인필터",
            10: "피크 전류", 9: "모터 구속", 8: "홀센서", 7: "TB", 6: "전원 결상/정전",
            5: "입력전압 낮음", 4: "입력전압 높음", 3: "DC링크 전압 낮음", 2: "DC링크 전압 높음",
            1: "지락", 0: "IGBT 오류"}


def bits_text(v: Optional[int], names: Dict[int, str], zero: str = "없음") -> str:
    if v is None:
        return "-"
    on = [names[b] for b in range(15, -1, -1) if v >> b & 1 and b in names]
    return f"0x{v:04X} " + (", ".join(on) if on else zero)


def view_text(addr: int, v: Optional[int]) -> str:
    if v is None:
        return "-"
    if addr == 3:
        return f"0x{v:04X} 주소 {v >> 8}, 속도코드 {(v >> 4) & 0xF}, 모드코드 {v & 0xF}"
    if addr == 4:
        extra = []
        if v & 0x0100:
            extra.append("버스 운전허가")
        return f"0x{v:04X} {CONTROL_MODES.get(v & 0xF, f'방식 {v & 0xF}')}" + (f" + {', '.join(extra)}" if extra else "")
    if addr == 16:
        return f"0x{v:04X} 기본 PIN 보호 {(v >> 8) & 3}단계"
    if addr == 25:
        return f"감속 {v >> 8}초 / 가속 {v & 0xFF}초"
    return f"{v} (0x{v:04X})"


def _signed(v: Optional[int]) -> Optional[int]:
    return None if v is None else (v - 65536 if v >= 32768 else v)


INPUT_ROWS = [  # (주소, 이름, 표시 함수)
    (0, "펌웨어", fw_text),
    (1, "제품코드", lambda v: "-" if v is None else f"0x{v:04X} {classify(v)[1]}"),
    (14, "현재 속도", lambda v: "-" if v is None else f"{v} rpm"),
    (15, "모터 전류", lambda v: "-" if v is None else f"{v / 100:.2f} A"),
    (33, "입력 전력", lambda v: "-" if v is None else f"{v} W"),
    (27, "출력(변조율)", lambda v: "-" if v is None else f"{v * 100 / 32767:.1f} %"),
    (26, "E1 입력", lambda v: "-" if v is None else f"{v * 10 / 32767:.2f} V ({v * 100 / 32767:.1f} %)"),
    (20, "DC 링크 전압", lambda v: "-" if v is None else f"{v} V"),
    (21, "전원 전압(최대값)", lambda v: "-" if v is None else f"{v} V"),
    (22, "IGBT 온도", lambda v: "-" if v is None else f"{_signed(v) / 10:.1f} ℃"),
    (23, "내부 온도", lambda v: "-" if v is None else f"{_signed(v) / 10:.1f} ℃"),
    (24, "MCU 온도", lambda v: "-" if v is None else f"{_signed(v) / 10:.1f} ℃"),
    (10, "운전 상태", lambda v: bits_text(v, OP_BITS, "정상 운전")),
    (12, "오류 상태", lambda v: bits_text(v, ERR_BITS, "정상")),
    (13, "오류 상태 2 (FW14)", lambda v: "-" if v is None else f"0x{v:04X}"),
    (49, "현재 PIN 보호 단계", lambda v: "-" if v is None else f"{v}단계"),
]
INPUT_BLOCKS = [(0, 16), (20, 8), (33, 1), (49, 1)]       # 한 번에 읽을 구간 (실패하면 하나씩)
HOLD_BLOCKS = [(0, 20), (25, 1), (30, 8)]


# ====================================================================== 통신
class Client:
    """통신장치와 TCP 연결 1개. 작업 쓰레드에서만 사용."""

    def __init__(self, host: str, port: int, timeout: float = DEFAULT_TIMEOUT) -> None:
        self.host, self.port, self.timeout = host, port, timeout
        self.sock: Optional[socket.socket] = None
        self.tid = 0

    def connect(self) -> None:
        self.close()
        s = socket.create_connection((self.host, self.port), timeout=CONNECT_TIMEOUT)
        s.settimeout(0.1)
        self.sock = s

    def close(self) -> None:
        if self.sock is not None:
            try:
                self.sock.close()
            except OSError:
                pass
        self.sock = None

    def request(self, build: Callable[..., bytes], *args, timeout: Optional[float] = None
                ) -> Tuple[Optional[modbus.Response], str]:
        """(응답, 상태). 상태: ok / exception / timeout / error. 연결이 없으면 먼저 접속."""
        if self.sock is None:
            try:
                self.connect()
            except OSError as e:
                return None, f"error: 접속 실패 ({e})"
        self.tid = (self.tid + 1) % 65536
        frame = build(self.tid, *args)
        wait = self.timeout if timeout is None else timeout
        splitter = modbus.FrameSplitter()
        try:
            self.sock.sendall(frame)
            end = time.monotonic() + wait
            while time.monotonic() < end:
                try:
                    data = self.sock.recv(512)
                except socket.timeout:
                    continue
                if not data:
                    raise ConnectionError("통신장치가 연결을 끊음")
                for f in splitter.feed(data):
                    r = modbus.parse_response(f)
                    if r is not None and r.tid == self.tid:     # 늦게 온 이전 응답은 무시
                        time.sleep(REQUEST_GAP)
                        return r, ("exception" if r.exception is not None else "ok")
        except OSError as e:
            self.close()
            return None, f"error: {e}"
        time.sleep(REQUEST_GAP)
        return None, "timeout"


def _ok(r: Optional[modbus.Response], st: str) -> bool:
    return r is not None and st == "ok"


def read_regs(c: Client, kind: str, blocks: List[Tuple[int, int]], uid: int) -> Dict[int, Optional[int]]:
    """구간별로 읽고, 구간에 없는 레지스터가 섞여 실패하면 하나씩 읽음. 못 읽은 값은 None."""
    build = modbus.build_read_holding if kind == "h" else modbus.build_read_input
    out: Dict[int, Optional[int]] = {}
    for start, count in blocks:
        r, st = c.request(build, uid, start, count)
        if _ok(r, st) and len(r.values) == count:
            out.update({start + i: v for i, v in enumerate(r.values)})
            continue
        if st.startswith("error") or st == "timeout":
            out.update({start + i: None for i in range(count)})
            continue
        for a in range(start, start + count):
            r, st = c.request(build, uid, a, 1)
            out[a] = r.values[0] if _ok(r, st) and r.values else None
    return out


def scan_id(c: Client, uid: int, find_new: bool = True) -> Optional[dict]:
    """ID 하나 확인. 응답이 없으면 None."""
    r, st = c.request(modbus.build_read_input, uid, 1, 1)          # i1 제품코드
    if st.startswith("error"):
        raise ConnectionError(st)
    if _ok(r, st) and r.values:
        kind, name = classify(r.values[0])
        res = {"uid": uid, "kind": kind, "product": r.values[0], "name": name}
        if is_ecblue(kind):
            res.update(read_summary(c, uid))
        return res
    if r is not None and r.exception in GATEWAY_NO_DEVICE:
        return None                                                  # 게이트웨이가 '장치 없음'이라고 알려줌
    present = r is not None                                          # 예외 응답 = 장치는 있음
    if find_new or present:
        r2, st2 = c.request(modbus.build_new_read, uid)              # 신규모터(SUNINUS) 0x13 부터 7개
        if st2.startswith("error"):
            raise ConnectionError(st2)
        if _ok(r2, st2) and len(r2.values) == 7:
            return {"uid": uid, "kind": KIND_NEW, "product": None,
                    "name": f"자동수동 {r2.values[0]} / 동작 {r2.values[1]} / 설정 {r2.values[3]} rpm"}
        if r2 is not None and r2.exception not in GATEWAY_NO_DEVICE:
            present = True
    if present:
        return {"uid": uid, "kind": KIND_OTHER, "product": None,
                "name": f"응답 있음 (예외 {r.exception if r else '-'})"}
    return None


def read_summary(c: Client, uid: int) -> dict:
    """목록 표시용: 펌웨어, 속도, 워치독, Set Intern1"""
    i = read_regs(c, "i", [(0, 1), (14, 1)], uid)
    h = read_regs(c, "h", [(5, 1), (17, 1)], uid)
    return {"fw": i.get(0), "speed": i.get(14), "h17": h.get(17), "h5": h.get(5)}


def read_all(c: Client, uid: int) -> dict:
    return {"uid": uid, "i": read_regs(c, "i", INPUT_BLOCKS, uid), "h": read_regs(c, "h", HOLD_BLOCKS, uid),
            "time": datetime.datetime.now()}


def write_regs(c: Client, uid: int, changes: List[Tuple[int, int]]) -> List[Tuple[str, str]]:
    """[(주소, 값)] 을 차례로 쓰고 다시 읽어 확인. 필요하면 PIN 보호를 먼저 풂. 결과 [(수준, 메시지)]."""
    steps: List[Tuple[str, str]] = []
    need = max((PARAM_BY_ADDR[a].pin if a in PARAM_BY_ADDR else 2 if a == 17 else 0) for a, _ in changes)
    if need:
        r, st = c.request(modbus.build_read_input, uid, 49, 1)
        level = r.values[0] if _ok(r, st) and r.values else 0
        if level < need:
            r, st = c.request(modbus.build_new_write_single, uid, 0, PIN_LEVEL2)
            if not _ok(r, st):
                steps.append(("error", f"PIN 보호 해제 실패 ({_st_text(r, st)}) - 쓰지 않음"))
                return steps
            steps.append(("info", f"PIN 보호 해제 (현재 {level}단계 → 2단계, 약 15분 유지)"))
    for addr, value in changes:
        name = PARAM_BY_ADDR[addr].name if addr in PARAM_BY_ADDR else ("통신 워치독" if addr == 17 else f"h{addr}")
        r, st = c.request(modbus.build_new_write_single, uid, addr, value)
        if not _ok(r, st):
            steps.append(("error", f"h{addr} {name} = {value} 쓰기 실패 ({_st_text(r, st)})"))
            continue
        time.sleep(WRITE_VERIFY_GAP)
        r, st = c.request(modbus.build_read_holding, uid, addr, 1)
        got = r.values[0] if _ok(r, st) and r.values else None
        if got == value:
            steps.append(("info", f"h{addr} {name} = {value} 쓰기 완료 (확인됨)"))
        else:
            steps.append(("warn", f"h{addr} {name} = {value} 썼으나 다시 읽은 값 {got}"))
    return steps


def _st_text(r: Optional[modbus.Response], st: str) -> str:
    if r is not None and r.exception is not None:
        return f"예외 {r.exception}: {EXCEPTIONS.get(r.exception, '알 수 없음')}"
    return {"timeout": "응답 없음"}.get(st, st)


def watchdog_changes(cur_h5: Optional[int], seconds: int, mode: int, speed: Optional[int]) -> List[Tuple[int, int]]:
    """자동 운전 켜기에 필요한 쓰기. 속도(h5)는 바뀔 때만 (영구저장 수명 보호)."""
    out: List[Tuple[int, int]] = []
    if mode in WD_SPEED_MODES and speed is not None and speed != cur_h5:
        out.append((5, speed))
    out.append((17, wd_encode(seconds, mode)))
    return out


# ====================================================================== 작업 쓰레드
class Worker(threading.Thread):
    """통신 작업을 한 번에 하나씩. 결과는 화면 쪽 큐로."""

    def __init__(self) -> None:
        super().__init__(daemon=True, name="scanner-worker")
        self.jobs: "queue.Queue[Optional[tuple]]" = queue.Queue()
        self.results: "queue.Queue[tuple]" = queue.Queue()
        self.client: Optional[Client] = None
        self.cancel = threading.Event()

    def submit(self, fn: Callable, *args, done: Optional[Callable] = None) -> None:
        self.jobs.put((fn, args, done))

    def post(self, done: Optional[Callable], res) -> None:
        self.results.put((done, res))

    def run(self) -> None:
        while True:
            job = self.jobs.get()
            if job is None:
                break
            fn, args, done = job
            try:
                res = fn(self, *args)
            except Exception as e:
                log.exception("작업 오류")
                res = {"error": f"{type(e).__name__}: {e}"}
            self.post(done, res)
        if self.client:
            self.client.close()

    def use(self, host: str, port: int, timeout: float) -> Client:
        if self.client is None or (self.client.host, self.client.port) != (host, port):
            if self.client:
                self.client.close()
            self.client = Client(host, port, timeout)
        self.client.timeout = timeout
        return self.client

    def stop(self) -> None:
        self.cancel.set()
        self.jobs.put(None)


def job_scan(w: Worker, host: str, port: int, first: int, last: int, timeout: float, find_new: bool,
             on_progress: Callable) -> dict:
    w.cancel.clear()
    c = w.use(host, port, timeout)
    try:
        c.connect()
    except OSError as e:
        return {"error": f"{host}:{port} 접속 실패 ({e})"}
    found = 0
    for uid in range(first, last + 1):
        if w.cancel.is_set():
            return {"cancelled": True, "found": found, "last": uid - 1}
        w.post(on_progress, {"uid": uid, "found": found})
        try:
            res = scan_id(c, uid, find_new)
        except ConnectionError as e:
            try:          # 연결이 끊기면 한 번 다시 접속하고 같은 ID 를 다시 확인
                c.connect()
                res = scan_id(c, uid, find_new)
            except (OSError, ConnectionError):
                return {"error": f"ID {uid} 확인 중 통신 끊김 ({e})", "found": found}
        if res:
            found += 1
            w.post(on_progress, {"uid": uid, "found": found, "device": res})
    return {"done": True, "found": found}


def job_read_all(w: Worker, host: str, port: int, timeout: float, uid: int) -> dict:
    return read_all(w.use(host, port, timeout), uid)


def job_summary(w: Worker, host: str, port: int, timeout: float, uid: int) -> dict:
    c = w.use(host, port, timeout)
    r, st = c.request(modbus.build_read_input, uid, 1, 1)
    if not _ok(r, st):
        return {"uid": uid, "error": _st_text(r, st)}
    res = {"uid": uid}
    res.update(read_summary(c, uid))
    return res


def job_write(w: Worker, host: str, port: int, timeout: float, uid: int, changes: List[Tuple[int, int]]) -> dict:
    c = w.use(host, port, timeout)
    steps = write_regs(c, uid, changes)
    data = read_all(c, uid)
    data["steps"] = steps
    return data


def job_wd_test(w: Worker, host: str, port: int, timeout: float, uid: int, wait_s: float,
                on_progress: Callable) -> dict:
    """통신을 wait_s 초 멈춘 뒤 속도·오류를 읽어 자동 운전이 됐는지 확인 (이 프로그램의 모든 통신을 멈춤)."""
    c = w.use(host, port, timeout)
    c.close()                                   # 연결도 끊어 통신장치 쪽 요청이 전혀 없게 함
    w.cancel.clear()
    end = time.monotonic() + wait_s
    while time.monotonic() < end:
        if w.cancel.is_set():
            return {"uid": uid, "cancelled": True}
        w.post(on_progress, {"left": end - time.monotonic()})
        time.sleep(0.5)
    first = read_regs(c, "i", [(12, 1), (14, 1)], uid)
    end = time.monotonic() + 2.0           # 2초 동안 0.5초마다 통신을 이어가며 원래대로 돌아오는지 확인
    after = first
    while time.monotonic() < end:
        time.sleep(0.5)
        after = read_regs(c, "i", [(12, 1), (14, 1)], uid)
    h = read_regs(c, "h", [(2, 1), (5, 1), (17, 1)], uid)
    return {"uid": uid, "first": first, "after": after, "h": h}


# ====================================================================== 설정 저장
def load_settings(path: str = "") -> dict:
    s = {"ip": "172.27.0.95", "port": DEFAULT_PORT, "first": DEFAULT_RANGE[0], "last": DEFAULT_RANGE[1],
         "timeout": DEFAULT_TIMEOUT, "find_new": True}
    try:
        with open(path or SETTINGS_FILE, encoding="utf-8") as f:
            s.update({k: v for k, v in json.load(f).items() if k in s})
    except (OSError, ValueError):
        pass
    return s


def save_settings(s: dict, path: str = "") -> None:
    try:
        with open(path or SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=1)
    except OSError:
        log.exception("설정 저장 실패")


# ====================================================================== 화면
FONT = ("맑은 고딕", 9)


def _int(text, default: int, lo: int, hi: int) -> Optional[int]:
    try:
        v = int(str(text).strip())
    except ValueError:
        return None
    return v if lo <= v <= hi else None


class App:
    def __init__(self, root: "tk.Tk", settings_file: str = "") -> None:
        self.root = root
        self.settings_file = settings_file or SETTINGS_FILE
        self.worker = Worker()
        self.worker.start()
        self.devices: Dict[int, dict] = {}
        self.details: Dict[int, "DetailWindow"] = {}
        self.scanning = False
        self.busy_test = False
        root.title(TITLE)
        root.geometry("980x620")
        root.minsize(760, 520)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.report_callback_exception = lambda *a: log.error("화면 오류\n%s", "".join(traceback.format_exception(*a)))
        self._build()
        s = load_settings(self.settings_file)
        self.v_ip.set(s["ip"]); self.v_port.set(str(s["port"]))
        self.v_first.set(str(s["first"])); self.v_last.set(str(s["last"]))
        self.v_timeout.set(str(s["timeout"])); self.v_new.set(bool(s["find_new"]))
        self._poll_id = root.after(100, self._poll)

    # ---------------------------------------------------------------- 구성
    def _build(self) -> None:
        r = self.root
        top = ttk.Frame(r, padding=(8, 8, 8, 2))
        top.pack(fill="x")
        self.v_ip, self.v_port = tk.StringVar(), tk.StringVar()
        self.v_first, self.v_last, self.v_timeout = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.v_new = tk.BooleanVar(value=True)
        ttk.Label(top, text="IP").pack(side="left")
        self.ent_ip = ttk.Entry(top, textvariable=self.v_ip, width=15)
        self.ent_ip.pack(side="left", padx=(4, 10))
        ttk.Label(top, text="PORT").pack(side="left")
        self.sp_port = ttk.Spinbox(top, from_=1, to=65535, textvariable=self.v_port, width=6)
        self.sp_port.pack(side="left", padx=(4, 10))
        ttk.Label(top, text="ID 범위").pack(side="left")
        self.sp_first = ttk.Spinbox(top, from_=ID_MIN, to=ID_MAX, textvariable=self.v_first, width=4)
        self.sp_first.pack(side="left", padx=(4, 2))
        ttk.Label(top, text="~").pack(side="left")
        self.sp_last = ttk.Spinbox(top, from_=ID_MIN, to=ID_MAX, textvariable=self.v_last, width=4)
        self.sp_last.pack(side="left", padx=(2, 10))
        ttk.Label(top, text="응답 대기(초)").pack(side="left")
        self.sp_timeout = ttk.Spinbox(top, from_=0.1, to=3.0, increment=0.1, textvariable=self.v_timeout, width=4)
        self.sp_timeout.pack(side="left", padx=(4, 10))
        self.chk_new = ttk.Checkbutton(top, text="신규모터도 찾기", variable=self.v_new)
        self.chk_new.pack(side="left")
        self.btn_scan = ttk.Button(top, text="스캔 시작", command=self.on_scan)
        self.btn_scan.pack(side="right")

        bar = ttk.Frame(r, padding=(8, 2))
        bar.pack(fill="x")
        self.progress = ttk.Progressbar(bar, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True)
        self.lb_status = ttk.Label(bar, text="IP 와 ID 범위를 정하고 [스캔 시작]을 누르세요.", width=46)
        self.lb_status.pack(side="left", padx=(8, 0))

        mid = ttk.Frame(r, padding=(8, 2))
        mid.pack(fill="both", expand=True)
        cols = ("id", "kind", "name", "fw", "speed", "wd")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=12, selectmode="browse")
        for c, t, w in (("id", "ID", 50), ("kind", "종류", 110), ("name", "제품 / 내용", 230), ("fw", "펌웨어", 60),
                        ("speed", "현재 속도", 80), ("wd", "통신 끊김 시 자동 운전", 360)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center" if c in ("id", "fw", "speed") else "w")
        sb = ttk.Scrollbar(mid, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.tag_configure("ecb", foreground="#0050A0")
        self.tree.tag_configure("other", foreground="#555555")
        self.tree.bind("<Double-1>", lambda e: self.on_open())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._update_buttons())

        btns = ttk.Frame(r, padding=(8, 2))
        btns.pack(fill="x")
        self.btn_open = ttk.Button(btns, text="ECblue 상태·설정 열기", command=self.on_open, state="disabled")
        self.btn_open.pack(side="left")
        self.btn_reread = ttk.Button(btns, text="선택 다시 읽기", command=self.on_reread, state="disabled")
        self.btn_reread.pack(side="left", padx=6)
        self.btn_clear = ttk.Button(btns, text="목록 지우기", command=self.on_clear)
        self.btn_clear.pack(side="left")

        lf = ttk.Frame(r, padding=(8, 2, 8, 8))
        lf.pack(fill="both")
        self.logbox = tk.Text(lf, height=8, font=FONT, state="disabled", wrap="word")
        lsb = ttk.Scrollbar(lf, command=self.logbox.yview)
        self.logbox.configure(yscrollcommand=lsb.set)
        self.logbox.pack(side="left", fill="both", expand=True)
        lsb.pack(side="left", fill="y")
        self.logbox.tag_configure("error", foreground="#D00000")
        self.logbox.tag_configure("warn", foreground="#B36B00")

    # ---------------------------------------------------------------- 공통
    def log(self, level: str, msg: str) -> None:
        line = f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
        getattr(log, {"error": "error", "warn": "warning"}.get(level, "info"))(msg)
        self.logbox.configure(state="normal")
        self.logbox.insert("end", line + "\n", (level,))
        n = int(self.logbox.index("end-1c").split(".")[0])
        if n > 2000:
            self.logbox.delete("1.0", f"{n - 2000}.0")
        self.logbox.configure(state="disabled")
        self.logbox.see("end")

    def conn(self) -> Optional[Tuple[str, int, float]]:
        ip = self.v_ip.get().strip()
        try:
            ipaddress.ip_address(ip)
        except ValueError:
            messagebox.showwarning(TITLE, "IP 형식이 올바르지 않습니다.", parent=self.root)
            return None
        port = _int(self.v_port.get(), DEFAULT_PORT, 1, 65535)
        try:
            timeout = float(self.v_timeout.get())
        except ValueError:
            timeout = -1
        if port is None or not 0.1 <= timeout <= 3.0:
            messagebox.showwarning(TITLE, "PORT(1~65535)와 응답 대기(0.1~3초)를 확인하세요.", parent=self.root)
            return None
        return ip, port, timeout

    def submit(self, fn, *args, done=None) -> None:
        self.worker.submit(fn, *args, done=done)

    def _poll(self) -> None:
        try:
            for _ in range(200):
                try:
                    done, res = self.worker.results.get_nowait()
                except queue.Empty:
                    break
                if done:
                    done(res)
        except Exception:
            log.exception("화면 갱신 오류")
        finally:
            self._poll_id = self.root.after(100, self._poll)

    def _update_buttons(self) -> None:
        d = self.selected()
        can = d is not None and is_ecblue(d["kind"]) and not self.scanning
        self.btn_open.configure(state="normal" if can else "disabled")
        self.btn_reread.configure(state="normal" if can else "disabled")

    def selected(self) -> Optional[dict]:
        sel = self.tree.selection()
        return self.devices.get(int(sel[0])) if sel else None

    # ---------------------------------------------------------------- 스캔
    def on_scan(self) -> None:
        if self.scanning:
            self.worker.cancel.set()
            self.btn_scan.configure(state="disabled", text="중지 중…")
            return
        if self.busy_test:
            return
        c = self.conn()
        if c is None:
            return
        first = _int(self.v_first.get(), 1, ID_MIN, ID_MAX)
        last = _int(self.v_last.get(), 240, ID_MIN, ID_MAX)
        if first is None or last is None or first > last:
            messagebox.showwarning(TITLE, f"ID 범위를 {ID_MIN}~{ID_MAX} 사이에서, 시작 ≤ 끝 으로 정하세요.",
                                   parent=self.root)
            return
        ip, port, timeout = c
        save_settings({"ip": ip, "port": port, "first": first, "last": last, "timeout": timeout,
                       "find_new": bool(self.v_new.get())}, self.settings_file)
        self.on_clear()
        self.scan_conn = c
        self.scanning = True
        self.scan_range = (first, last)
        self.progress.configure(maximum=last - first + 1, value=0)
        self.btn_scan.configure(text="스캔 중지")
        for w in (self.ent_ip, self.sp_port, self.sp_first, self.sp_last, self.sp_timeout, self.chk_new):
            w.configure(state="disabled")
        self._update_buttons()
        est = (last - first + 1) * timeout * (2 if self.v_new.get() else 1)
        self.log("info", f"스캔 시작 {ip}:{port} ID {first}~{last} (응답 없는 ID 가 많으면 최대 약 {est / 60:.1f}분)")
        self.submit(job_scan, ip, port, first, last, timeout, bool(self.v_new.get()), self._scan_progress,
                    done=self._scan_done)

    def _scan_progress(self, p: dict) -> None:
        first, last = self.scan_range
        self.progress.configure(value=p["uid"] - first + (1 if "device" in p else 0))
        self.lb_status.configure(text=f"ID {p['uid']} 확인 중 ({p['uid'] - first + 1}/{last - first + 1}), "
                                      f"찾음 {p['found']}개")
        if "device" in p:
            d = p["device"]
            self.devices[d["uid"]] = d
            self._show(d)
            self.log("info", f"ID {d['uid']}: {d['kind']} - {d['name']}")

    def _scan_done(self, res: dict) -> None:
        self.scanning = False
        self.btn_scan.configure(text="스캔 시작", state="normal")
        for w in (self.ent_ip, self.sp_port, self.sp_first, self.sp_last, self.sp_timeout, self.chk_new):
            w.configure(state="normal")
        n = len(self.devices)
        if "error" in res:
            self.log("error", f"스캔 중단: {res['error']}")
            self.lb_status.configure(text=f"오류로 중단 - 찾음 {n}개")
        elif res.get("cancelled"):
            self.log("warn", f"스캔 중지 (ID {res['last']} 까지 확인) - 찾음 {n}개")
            self.lb_status.configure(text=f"중지됨 - 찾음 {n}개")
        else:
            self.progress.configure(value=self.progress.cget("maximum"))
            ecb = sum(1 for d in self.devices.values() if is_ecblue(d["kind"]))
            self.log("info", f"스캔 완료 - 찾음 {n}개 (ECblue {ecb}개)")
            self.lb_status.configure(text=f"완료 - 찾음 {n}개 (ECblue {ecb}개)")
        self._update_buttons()

    def _show(self, d: dict) -> None:
        ecb = is_ecblue(d["kind"])
        speed = d.get("speed")
        vals = (d["uid"], d["kind"], d["name"], fw_text(d.get("fw")) if ecb else "-",
                "-" if speed is None else f"{speed} rpm", wd_text(d.get("h17"), d.get("h5")) if ecb else "-")
        iid = str(d["uid"])
        if self.tree.exists(iid):
            self.tree.item(iid, values=vals)
        else:
            after = [int(i) for i in self.tree.get_children() if int(i) < d["uid"]]
            self.tree.insert("", len(after), iid=iid, values=vals, tags=("ecb" if ecb else "other",))

    def on_clear(self) -> None:
        if self.scanning:
            return
        self.devices.clear()
        self.tree.delete(*self.tree.get_children())
        self._update_buttons()

    def on_reread(self) -> None:
        d = self.selected()
        c = self.conn()
        if d and c:
            self.submit(job_summary, *c, d["uid"], done=self._summary_done)

    def _summary_done(self, res: dict) -> None:
        d = self.devices.get(res.get("uid"))
        if d is None:
            return
        if "error" in res:
            self.log("warn", f"ID {res['uid']} 다시 읽기 실패: {res['error']}")
            return
        d.update({k: res[k] for k in ("fw", "speed", "h17", "h5")})
        self._show(d)

    # ---------------------------------------------------------------- 상세
    def on_open(self) -> None:
        d = self.selected()
        if d is None or not is_ecblue(d["kind"]) or self.scanning:
            return
        c = self.conn()
        if c is None:
            return
        win = self.details.get(d["uid"])
        if win and win.alive():
            win.top.lift()
            return
        self.details[d["uid"]] = DetailWindow(self, d, c)

    def on_close(self) -> None:
        self.worker.cancel.set()
        self.worker.stop()
        try:
            self.root.after_cancel(self._poll_id)
        except (tk.TclError, ValueError):
            pass
        self.root.destroy()


class DetailWindow:
    """ECblue 한 대의 상태 / 설정 변경 / 통신 끊김 시 자동 운전"""

    def __init__(self, app: App, dev: dict, conn: Tuple[str, int, float]) -> None:
        self.app, self.dev, self.conn = app, dev, conn
        self.uid = dev["uid"]
        self.data: Optional[dict] = None
        self.top = tk.Toplevel(app.root)
        self.top.title(f"ECblue ID {self.uid} - {conn[0]}  ({dev['name']})")
        self.top.geometry("900x580")
        self.top.protocol("WM_DELETE_WINDOW", self.close)
        self.auto = tk.BooleanVar(value=False)
        self._after: Optional[str] = None
        self._pending = False
        self._build()
        self.refresh()

    def alive(self) -> bool:
        try:
            return bool(self.top.winfo_exists())
        except tk.TclError:
            return False

    def close(self) -> None:
        if self._after:
            self.top.after_cancel(self._after)
        self.top.destroy()
        self.app.details.pop(self.uid, None)

    # ---------------------------------------------------------------- 구성
    def _build(self) -> None:
        nb = ttk.Notebook(self.top)
        nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.nb = nb
        # 상태
        f1 = ttk.Frame(nb, padding=6)
        nb.add(f1, text="상태")
        self.tv_info = ttk.Treeview(f1, columns=("name", "value"), show="headings", height=16)
        self.tv_info.heading("name", text="항목")
        self.tv_info.heading("value", text="값")
        self.tv_info.column("name", width=170)
        self.tv_info.column("value", width=660)
        self.tv_info.pack(fill="both", expand=True)
        row = ttk.Frame(f1)
        row.pack(fill="x", pady=(6, 0))
        ttk.Button(row, text="새로 읽기", command=self.refresh).pack(side="left")
        ttk.Checkbutton(row, text="2초마다 자동 갱신", variable=self.auto, command=self._auto_changed).pack(side="left", padx=8)
        self.lb_time = ttk.Label(row, text="")
        self.lb_time.pack(side="right")
        # 설정 변경
        f2 = ttk.Frame(nb, padding=6)
        nb.add(f2, text="설정 변경")
        grid = ttk.Frame(f2)
        grid.pack(fill="x")
        for col, t in enumerate(("항목", "레지스터", "현재값", "새 값", "단위·범위", "비고")):
            ttk.Label(grid, text=t, font=(FONT[0], 9, "bold")).grid(row=0, column=col, sticky="w", padx=3)
        self.cur_labels: Dict[int, ttk.Label] = {}
        self.new_vars: Dict[int, tk.StringVar] = {}
        for i, p in enumerate(EDIT_PARAMS, start=1):
            ttk.Label(grid, text=p.name).grid(row=i, column=0, sticky="w", padx=3, pady=1)
            ttk.Label(grid, text=f"h{p.addr}").grid(row=i, column=1, sticky="w", padx=3)
            lb = ttk.Label(grid, text="-", width=8)
            lb.grid(row=i, column=2, sticky="w", padx=3)
            self.cur_labels[p.addr] = lb
            v = tk.StringVar()
            ttk.Entry(grid, textvariable=v, width=8).grid(row=i, column=3, sticky="w", padx=3)
            self.new_vars[p.addr] = v
            ttk.Label(grid, text=f"{p.unit} {p.lo}~{p.hi}").grid(row=i, column=4, sticky="w", padx=3)
            flag = ("영구저장·" if p.nv else "") + (f"PIN{p.pin}" if p.pin else "PIN 없음")
            ttk.Label(grid, text=f"{flag} - {p.note}", foreground="#555555").grid(row=i, column=5, sticky="w", padx=3)
        row2 = ttk.Frame(f2)
        row2.pack(fill="x", pady=6)
        self.btn_apply = ttk.Button(row2, text="변경 적용", command=self.on_apply)
        self.btn_apply.pack(side="left")
        ttk.Button(row2, text="입력 되돌리기", command=self._fill_inputs).pack(side="left", padx=6)
        ttk.Label(row2, text="값을 바꾼 항목만 씁니다. 영구저장 항목은 쓰기 수명(약 1만 회)이 있어 자주 바꾸지 마세요.",
                  foreground="#555555").pack(side="left", padx=6)
        ttk.Label(f2, text="읽기 전용 (이 프로그램에서는 바꾸지 않음)", font=(FONT[0], 9, "bold")).pack(anchor="w", pady=(6, 2))
        self.tv_view = ttk.Treeview(f2, columns=("reg", "name", "value", "note"), show="headings", height=8)
        for c, t, w in (("reg", "레지스터", 60), ("name", "항목", 170), ("value", "값", 300), ("note", "비고", 330)):
            self.tv_view.heading(c, text=t)
            self.tv_view.column(c, width=w)
        self.tv_view.pack(fill="both", expand=True)
        # 통신 끊김 시 자동 운전
        f3 = ttk.Frame(nb, padding=10)
        nb.add(f3, text="통신 끊김 시 자동 운전")
        self.lb_wd = ttk.Label(f3, text="-", font=(FONT[0], 11, "bold"))
        self.lb_wd.pack(anchor="w")
        ttk.Label(f3, wraplength=840, justify="left", foreground="#333333", text=(
            "통신 워치독(h17): 정한 시간 동안 이 팬에 Modbus 통신이 한 번도 오지 않으면 팬이 스스로 지정 동작을 합니다. "
            "'정속 운전'은 Set Intern1(h5) 속도로 돕니다. 통신이 계속 오는 동안에는 평소대로 동작합니다.")).pack(anchor="w", pady=(4, 8))
        g = ttk.Frame(f3)
        g.pack(anchor="w")
        self.v_wd_sec, self.v_wd_speed = tk.StringVar(value="30"), tk.StringVar(value="")
        self.v_wd_mode = tk.StringVar(value=WD_MODES[2])
        ttk.Label(g, text="통신이 끊긴 지").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(g, from_=1, to=255, textvariable=self.v_wd_sec, width=5).grid(row=0, column=1, padx=4)
        ttk.Label(g, text="초 (1~255) 지나면").grid(row=0, column=2, sticky="w")
        ttk.Label(g, text="동작").grid(row=1, column=0, sticky="w", pady=4)
        self.cb_mode = ttk.Combobox(g, textvariable=self.v_wd_mode, state="readonly", width=34,
                                    values=[WD_MODES[m] for m in WD_CHOICES])
        self.cb_mode.grid(row=1, column=1, columnspan=2, sticky="w", padx=4)
        ttk.Label(g, text="운전 속도").grid(row=2, column=0, sticky="w")
        ttk.Entry(g, textvariable=self.v_wd_speed, width=7).grid(row=2, column=1, padx=4, sticky="w")
        self.lb_wd_speed = ttk.Label(g, text="rpm (Set Intern1, h5)")
        self.lb_wd_speed.grid(row=2, column=2, sticky="w")
        b = ttk.Frame(f3)
        b.pack(anchor="w", pady=10)
        self.btn_wd_on = ttk.Button(b, text="켜기 / 적용", command=self.on_wd_on)
        self.btn_wd_on.pack(side="left")
        self.btn_wd_off = ttk.Button(b, text="끄기", command=self.on_wd_off)
        self.btn_wd_off.pack(side="left", padx=6)
        self.btn_wd_test = ttk.Button(b, text="동작 시험", command=self.on_wd_test)
        self.btn_wd_test.pack(side="left")
        self.lb_wd_test = ttk.Label(f3, text="", foreground="#0050A0", wraplength=840, justify="left")
        self.lb_wd_test.pack(anchor="w")
        ttk.Label(f3, wraplength=840, justify="left", foreground="#7A4A00", text=(
            "주의\n"
            "• 같은 통신선에 다른 프로그램(운영 프로그램, 미니 제어기 등)이 이 팬과 통신 중이면 워치독이 동작하지 않습니다.\n"
            "• 반대로 이 팬을 평소 통신으로 제어하지 않는 경우(아날로그 E1 제어 등) 켜 두면 정한 시간 뒤 정속 운전으로 바뀝니다.\n"
            "• h17·h5 는 영구저장 설정입니다(쓰기 수명 약 1만 회). 문서의 '정속 1'을 Set Intern1(h5)로 판단했으므로 "
            "처음에는 [동작 시험]으로 확인하세요.\n"
            "• 통신이 다시 왔을 때 원래 동작으로 돌아오는 방식은 제조사 문서에 없습니다. 시험 결과의 '통신 2초 계속 뒤' 값을 확인하세요."
        )).pack(anchor="w", pady=(8, 0))

    # ---------------------------------------------------------------- 읽기
    def _auto_changed(self) -> None:
        if self.auto.get():
            self._schedule()
        elif self._after:
            self.top.after_cancel(self._after)
            self._after = None

    def _schedule(self) -> None:
        if self._after:
            self.top.after_cancel(self._after)
        self._after = self.top.after(REFRESH_MS, self._tick)

    def _tick(self) -> None:
        self._after = None
        if self.auto.get() and self.alive():
            self.refresh()
            self._schedule()

    def refresh(self) -> None:
        if self._pending or self.app.busy_test:
            return
        self._pending = True
        self.app.submit(job_read_all, *self.conn, self.uid, done=self._loaded)

    def _loaded(self, data: dict) -> None:
        self._pending = False
        if not self.alive():
            return
        if "error" in data:
            self.app.log("error", f"ID {self.uid} 읽기 오류: {data['error']}")
            return
        first = self.data is None
        self.data = data
        i, h = data["i"], data["h"]
        self.tv_info.delete(*self.tv_info.get_children())
        for addr, name, fn in INPUT_ROWS:
            self.tv_info.insert("", "end", values=(f"{name} (i{addr})", fn(i.get(addr))))
        self.tv_info.insert("", "end", values=("통신 끊김 시 자동 운전 (h17)", wd_text(h.get(17), h.get(5))))
        self.lb_time.configure(text=f"읽은 시각 {data['time']:%H:%M:%S}")
        if all(v is None for v in i.values()):
            self.lb_time.configure(text=f"응답 없음 ({data['time']:%H:%M:%S})")
        for p in EDIT_PARAMS:
            v = h.get(p.addr)
            self.cur_labels[p.addr].configure(text="-" if v is None else str(v))
        self.tv_view.delete(*self.tv_view.get_children())
        for p in VIEW_PARAMS:
            self.tv_view.insert("", "end", values=(f"h{p.addr}", p.name, view_text(p.addr, h.get(p.addr)), p.note))
        self.lb_wd.configure(text="현재: " + wd_text(h.get(17), h.get(5)))
        fw = i.get(0) or 0
        self.cb_mode.configure(values=[WD_MODES[m] for m in WD_CHOICES if m in (2, 3) or fw >= WD_FW_E1])
        if first:
            self._fill_inputs()
            sec, mode = wd_decode(h.get(17) or 0)
            if sec:
                self.v_wd_sec.set(str(sec))
            if mode in WD_CHOICES:
                self.v_wd_mode.set(WD_MODES[mode])
            if h.get(5) is not None:
                self.v_wd_speed.set(str(h[5]))
        d = self.app.devices.get(self.uid)
        if d is not None:
            d.update({"fw": i.get(0), "speed": i.get(14), "h17": h.get(17), "h5": h.get(5)})
            self.app._show(d)

    def _fill_inputs(self) -> None:
        if self.data:
            for p in EDIT_PARAMS:
                v = self.data["h"].get(p.addr)
                self.new_vars[p.addr].set("" if v is None else str(v))

    # ---------------------------------------------------------------- 설정 변경
    def on_apply(self) -> None:
        if not self.data:
            return
        h = self.data["h"]
        changes, lines, errs = [], [], []
        new_vals = {}
        for p in EDIT_PARAMS:
            text = self.new_vars[p.addr].get().strip()
            if text == "" or h.get(p.addr) is None:
                continue
            v = _int(text, 0, p.lo, p.hi)
            if v is None:
                errs.append(f"{p.name}: {p.lo}~{p.hi} 사이 정수")
                continue
            new_vals[p.addr] = v
            if v != h[p.addr]:
                changes.append((p.addr, v))
                lines.append(f"• h{p.addr} {p.name}: {h[p.addr]} → {v} {p.unit}" + (" (영구저장)" if p.nv else ""))
        mx, mn = new_vals.get(8, h.get(8)), new_vals.get(7, h.get(7))
        if mx and mn is not None and mn > mx:
            errs.append(f"최저속도({mn})가 최고속도({mx})보다 큽니다")
        if errs:
            messagebox.showwarning(TITLE, "입력값을 확인하세요.\n" + "\n".join(errs), parent=self.top)
            return
        if not changes:
            messagebox.showinfo(TITLE, "바뀐 값이 없습니다.", parent=self.top)
            return
        if not messagebox.askyesno(TITLE, f"ECblue ID {self.uid} 에 다음 값을 씁니다.\n\n" + "\n".join(lines)
                                   + "\n\n진행할까요?", parent=self.top):
            return
        self._write(changes, "설정 변경")

    def _write(self, changes: List[Tuple[int, int]], what: str) -> None:
        self.app.log("info", f"ID {self.uid} {what}: " + ", ".join(f"h{a}={v}" for a, v in changes))
        self._set_busy(True)
        self.app.submit(job_write, *self.conn, self.uid, changes, done=self._written)

    def _written(self, data: dict) -> None:
        self._set_busy(False)
        if "error" in data:
            self.app.log("error", f"ID {self.uid} 쓰기 오류: {data['error']}")
            return
        for level, msg in data.get("steps", []):
            self.app.log(level, f"ID {self.uid} {msg}")
        bad = [m for lv, m in data.get("steps", []) if lv != "info"]
        if self.alive():
            self.data = None          # 새로 읽은 값으로 입력칸도 채움
            self._loaded(data)
            if bad:
                messagebox.showwarning(TITLE, "\n".join(bad), parent=self.top)

    def _set_busy(self, busy: bool) -> None:
        if self.alive():
            st = "disabled" if busy else "normal"
            for b in (self.btn_apply, self.btn_wd_on, self.btn_wd_off, self.btn_wd_test):
                b.configure(state=st)

    # ---------------------------------------------------------------- 통신 끊김 시 자동 운전
    def _wd_mode(self) -> int:
        return next(m for m in WD_CHOICES if WD_MODES[m] == self.v_wd_mode.get())

    def on_wd_on(self) -> None:
        if not self.data:
            return
        h = self.data["h"]
        sec = _int(self.v_wd_sec.get(), 30, 1, 255)
        mode = self._wd_mode()
        speed = None
        if mode in WD_SPEED_MODES:
            speed = _int(self.v_wd_speed.get(), 0, 0, 6000)
            if speed is None:
                messagebox.showwarning(TITLE, "운전 속도를 0~6000 rpm 으로 입력하세요.", parent=self.top)
                return
            if h.get(8) and speed > h[8]:
                messagebox.showwarning(TITLE, f"운전 속도가 최고속도(h8 = {h[8]} rpm)보다 큽니다. "
                                              "최고속도로 제한되므로 값을 낮추거나 최고속도를 먼저 바꾸세요.", parent=self.top)
                return
        if sec is None:
            messagebox.showwarning(TITLE, "시간을 1~255초로 입력하세요.", parent=self.top)
            return
        changes = watchdog_changes(h.get(5), sec, mode, speed)
        if changes == [(17, h.get(17))]:
            messagebox.showinfo(TITLE, "이미 같은 설정입니다.", parent=self.top)
            return
        desc = f"통신이 {sec}초 끊기면 {WD_MODES[mode]}" + (f" {speed} rpm" if speed is not None else "")
        if not messagebox.askyesno(TITLE, f"ECblue ID {self.uid}\n\n{desc}\n\n"
                                          "영구저장 설정(h17" + (", h5" if any(a == 5 for a, _ in changes) else "")
                                   + ")을 씁니다. 진행할까요?", parent=self.top):
            return
        self._write(changes, f"자동 운전 켜기 ({desc})")

    def on_wd_off(self) -> None:
        if not self.data:
            return
        if (self.data["h"].get(17) or 0) == 0:
            messagebox.showinfo(TITLE, "이미 꺼져 있습니다.", parent=self.top)
            return
        if messagebox.askyesno(TITLE, f"ECblue ID {self.uid} 의 통신 끊김 시 자동 운전을 끕니다(h17 = 0). 진행할까요?",
                               parent=self.top):
            self._write([(17, 0)], "자동 운전 끄기")

    def on_wd_test(self) -> None:
        if not self.data:
            return
        sec, mode = wd_decode(self.data["h"].get(17) or 0)
        if not sec or not mode:
            messagebox.showinfo(TITLE, "먼저 자동 운전을 켜세요.", parent=self.top)
            return
        if self.app.scanning:
            return
        wait = sec + 3
        if not messagebox.askyesno(TITLE, (
                f"이 프로그램의 모든 통신을 {wait}초 동안 멈춘 뒤 ID {self.uid} 의 속도와 오류 상태를 읽습니다.\n\n"
                "다른 프로그램이 같은 통신선으로 이 팬과 통신 중이면 시험 결과가 맞지 않습니다.\n"
                f"시험 중 팬이 {WD_MODES[mode]}"
                + (f" ({self.data['h'].get(5)} rpm)" if mode in WD_SPEED_MODES else "") + " 상태가 될 수 있습니다.\n\n진행할까요?"),
                parent=self.top):
            return
        self.app.busy_test = True
        self._set_busy(True)
        self.app.log("info", f"ID {self.uid} 자동 운전 시험: {wait}초 동안 통신 중지")
        self.app.submit(job_wd_test, *self.conn, self.uid, wait, self._test_progress, done=self._test_done)

    def _test_progress(self, p: dict) -> None:
        if self.alive():
            self.lb_wd_test.configure(text=f"시험 중: 통신 중지 … {p['left']:.0f}초 남음")

    def _test_done(self, res: dict) -> None:
        self.app.busy_test = False
        self._set_busy(False)
        if "error" in res or res.get("cancelled"):
            self.app.log("error", f"ID {self.uid} 자동 운전 시험 실패: {res.get('error', '취소')}")
            return
        f, a, h = res["first"], res["after"], res["h"]
        sec, mode = wd_decode(h.get(17) or 0)
        tripped = bool((f.get(12) or 0) & 0x8000)
        want = h.get(5)
        msg = (f"통신 재개 직후: 속도 {f.get(14)} rpm, 통신 워치독 오류 {'있음' if tripped else '없음'} / "
               f"통신 2초 계속 뒤: 속도 {a.get(14)} rpm, 오류 {'있음' if (a.get(12) or 0) & 0x8000 else '없음'}")
        if mode in WD_SPEED_MODES and f.get(14) is not None and want is not None:
            target = min(want, h[8]) if h.get(8) else want          # 최고속도(h8)로 제한됨
            ok = abs(f[14] - target) <= max(30, target * 0.05)
            msg += f"  →  지정 속도 {target} rpm 와 {'일치 (정상 동작)' if ok else '다름 - 현장 확인 필요'}"
        level = "info" if tripped else "warn"
        if not tripped:
            msg += "  (워치독 오류가 보이지 않음: 다른 통신이 있었거나, 시간이 짧았거나, 이 펌웨어에서 표시하지 않을 수 있음)"
        self.app.log(level, f"ID {self.uid} 자동 운전 시험 결과 - {msg}")
        if self.alive():
            self.lb_wd_test.configure(text="시험 결과 - " + msg)
            self.refresh()


def main() -> None:
    os.makedirs(LOG_DIR, exist_ok=True)
    h = logging.FileHandler(os.path.join(LOG_DIR, f"motor_scanner_{datetime.date.today():%Y%m%d}.log"), encoding="utf-8")
    h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logging.getLogger().addHandler(h)
    logging.getLogger().setLevel(logging.INFO)
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
