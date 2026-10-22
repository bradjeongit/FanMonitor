"""SUNINUS MOTOR TEST (모터 시험기) - MOTOR_TESTER.exe(VB.NET) 를 파이썬으로 새로 구현.

원본 동작
- [정상모터] 모드: 기존모터(RPM원본) 1대의 회전수(FC04 레지스터 14)를 읽어, 신규모터(최대 5대)에
  같은 회전수를 씀. 신규모터가 자동이면 수동(0x13=1), 정지면 동작(0x14=1)으로 바꾼 뒤 RPM(0x16)을 씀.
- [수동수치] 모드: 기존모터를 읽지 않고 화면의 RPM(0~800) 값을 신규모터에 씀. 실행 중 [적용] 으로 변경.
- 통신 내용은 화면의 두 탭(원본통신LOG / 신규모터통신LOG)과 log\\source-날짜.log, log\\target-날짜.log 에 남김.
- 설정은 Motor_Config.ini (SOURCE_CONFIG / TARGET_CONFIG) 에 저장.

원본이 실행 중 죽거나 멈추던 원인과 이 구현에서 바꾼 점은 docs/USER_MANUAL.md 14장 참고.
요약: 작업 쓰레드의 예외 처리 누락, 끝난 쓰레드 재시작(ThreadStateException), 동기 Invoke,
끝없이 커지는 로그창, 닫지 않는 소켓, 잠금 없는 공유 큐, 수동 모드의 쉬지 않는 반복(CPU 100%).

실행: MOTOR_TESTER.exe 또는 motor_tester.bat 더블클릭, 또는 python tools/motor_tester.py
"""
from __future__ import annotations

import collections
import datetime
import faulthandler
import ipaddress
import logging
import os
import socket
import sys
import threading
import time
import traceback
from dataclasses import dataclass, field
from typing import Callable, Deque, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from motor_monitor import modbus  # noqa: E402

try:  # 화면 없이 통신부만 시험할 때도 가져올 수 있도록
    import tkinter as tk
    from tkinter import messagebox, ttk
except ImportError:  # pragma: no cover
    tk = None  # type: ignore

TITLE = "SUNINUS MOTOR TEST"

# ---- 시간 (초). 원본 값 그대로 ----
CONNECT_TIMEOUT = 1.0     # 원본 CHECK_CONNECT_POSSIBLE: 1초 안에 접속 안 되면 실패
RETRY_WAIT = 1.0          # 접속 실패 후 다시 시도까지
SOURCE_REPLY_WAIT = 1.0   # 기존모터 읽기 요청 후 응답 확인까지
SOURCE_PAUSE = 5.0        # 기존모터 응답 확인 후 다음 읽기까지
TARGET_READ_GAP = 1.0     # 신규모터 읽기 요청 사이
TARGET_WRITE_GAP = 0.5    # 신규모터 쓰기 요청 사이
TARGET_PAUSE = 5.0        # 신규모터 한 바퀴 읽은 뒤 쉬는 시간
# ---- 안정성 (새로 추가) ----
RECV_POLL = 0.2           # 수신 대기 단위. 종료 요청에 이 시간 안에 반응
STALL_LIMIT = 30.0        # 작업 쓰레드가 이 시간 넘게 진행이 없으면 연결을 끊어 풀어줌
UI_POLL_MS = 100
UI_MAX_LINES = 3000       # 로그창 최대 줄 수 (넘으면 오래된 줄부터 지움. 파일에는 모두 남음)
UI_QUEUE_MAX = 5000       # 화면으로 보낼 로그 대기 최대 개수

RPM_MIN, RPM_MAX = 0, 800
ID_MAX = 99
TARGET_SLOTS = 5

_BASE = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
         else os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_DIR = os.path.join(_BASE, "log")                 # 원본과 같은 폴더 이름
CONFIG_FILE = os.path.join(_BASE, "Motor_Config.ini")
# 원본은 "\Motor_Config.ini" (현재 드라이브 맨 위, 예: C:\Motor_Config.ini) 를 썼음. 처음 실행 때 가져옴
LEGACY_CONFIG_FILE: Optional[str] = os.path.abspath("\\Motor_Config.ini") if os.name == "nt" else None

log = logging.getLogger("motor_tester")


# ====================================================================== 설정
def _clamp_int(text, default: int, lo: int, hi: int) -> int:
    """화면 입력값: 숫자가 아니면 기본값, 범위 밖이면 가까운 끝값"""
    try:
        v = int(float(str(text).strip()))
    except (TypeError, ValueError):
        return default
    return max(lo, min(hi, v))


def _ini_int(text, default: int, lo: int, hi: int) -> int:
    """설정 파일 값: 원본 Get_Local_Config 와 같이 숫자가 아니거나 범위 밖이면 기본값"""
    try:
        v = int(str(text).strip())
    except (TypeError, ValueError):
        return default
    return v if lo <= v <= hi else default


def parse_ini(text: str) -> Dict[str, Dict[str, str]]:
    """Windows GetPrivateProfileString 처럼 너그럽게 읽음: 깨진 줄은 건너뛰고, 이름은 대소문자 구분 없음."""
    out: Dict[str, Dict[str, str]] = {}
    sec: Optional[Dict[str, str]] = None
    for line in text.splitlines():
        line = line.strip()
        if not line or line[0] in ";#":
            continue
        if line.startswith("[") and "]" in line:
            sec = out.setdefault(line[1:line.index("]")].strip().upper(), {})
        elif "=" in line and sec is not None:
            k, v = line.split("=", 1)
            sec.setdefault(k.strip().upper(), v.strip())
    return out


@dataclass
class Settings:
    manual: bool = False                 # SOURCE_TYPE 0=정상모터, 1=수동수치
    src_ip: str = "127.0.0.1"
    src_port: int = 502
    src_id: int = 1
    src_rpm: int = 0
    tgt_ip: str = "127.0.0.1"
    tgt_port: int = 502
    tgt_ids: List[int] = field(default_factory=lambda: [1, 0, 0, 0, 0])

    def motor_ids(self) -> List[int]:
        """0(사용안함)과 중복을 뺀 신규모터 ID (입력 순서 유지). 원본 CMotor_Runner 생성자와 같음."""
        out: List[int] = []
        for i in self.tgt_ids:
            if i > 0 and i not in out:
                out.append(i)
        return out

    @classmethod
    def load(cls, path: str = "", legacy: Optional[str] = None) -> "Settings":
        path = path or CONFIG_FILE
        legacy = LEGACY_CONFIG_FILE if legacy is None else legacy
        src = path if os.path.exists(path) else (legacy if legacy and os.path.exists(legacy) else None)
        s = cls()
        if not src:
            return s
        try:
            with open(src, "rb") as f:
                raw = f.read()
        except OSError:
            log.exception("설정 파일 읽기 실패: %s", src)
            return s
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("cp949", errors="replace")
        ini = parse_ini(text)

        def get(sec: str, key: str, default: str) -> str:
            return ini.get(sec.upper(), {}).get(key.upper(), default).strip()

        # 원본과 같이 숫자가 아니거나 범위를 벗어난 값은 기본값으로
        s.manual = get("SOURCE_CONFIG", "SOURCE_TYPE", "0") != "0"
        s.src_ip = get("SOURCE_CONFIG", "SOURCE_IP", "127.0.0.1") or "127.0.0.1"
        s.src_port = _ini_int(get("SOURCE_CONFIG", "SOURCE_PORT", "502"), 502, 1, 65535)
        s.src_id = _ini_int(get("SOURCE_CONFIG", "SOURCE_ID", "1"), 1, 1, ID_MAX)
        s.src_rpm = _ini_int(get("SOURCE_CONFIG", "SOURCE_RPM", "0"), 0, RPM_MIN, RPM_MAX)
        s.tgt_ip = get("TARGET_CONFIG", "TARGET_IP", "127.0.0.1") or "127.0.0.1"
        s.tgt_port = _ini_int(get("TARGET_CONFIG", "TARGET_PORT", "502"), 502, 1, 65535)
        s.tgt_ids = [_ini_int(get("TARGET_CONFIG", f"TARGET_ID{n}", "1" if n == 1 else "0"),
                                1 if n == 1 else 0, 0, ID_MAX) for n in range(1, TARGET_SLOTS + 1)]
        return s

    def save(self, path: str = "") -> bool:
        path = path or CONFIG_FILE
        lines = [
            "[SOURCE_CONFIG]",
            f"SOURCE_TYPE={1 if self.manual else 0}",
            f"SOURCE_IP={self.src_ip}",
            f"SOURCE_PORT={self.src_port}",
            f"SOURCE_ID={self.src_id}",
            f"SOURCE_RPM={self.src_rpm}",
            "[TARGET_CONFIG]",
            f"TARGET_IP={self.tgt_ip}",
            f"TARGET_PORT={self.tgt_port}",
        ] + [f"TARGET_ID{n}={v}" for n, v in enumerate(self.tgt_ids, start=1)]
        tmp = path + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8", newline="\r\n") as f:
                f.write("\n".join(lines) + "\n")
            os.replace(tmp, path)   # 쓰는 도중 꺼져도 기존 설정이 깨지지 않도록
            return True
        except OSError:
            log.exception("설정 파일 저장 실패: %s", path)
            return False


# ====================================================================== 로그
class LogSink:
    """통신 로그: 파일(log/source-날짜.log, log/target-날짜.log)에 쓰고, 화면용 대기열에 넣음.
    작업 쓰레드는 화면을 기다리지 않음 (원본은 Control.Invoke 로 화면이 처리할 때까지 멈춤)."""

    def __init__(self, log_dir: Optional[str] = None) -> None:
        self.log_dir = log_dir or LOG_DIR
        self.ui: Deque[Tuple[str, str]] = collections.deque(maxlen=UI_QUEUE_MAX)
        self._lock = threading.Lock()
        self.file_errors = 0

    def __call__(self, target: str, msg: str) -> None:
        now = datetime.datetime.now()
        self.ui.append((target, f"[{now:%H:%M:%S}]{msg}"))
        name = "source" if target == "SOURCE" else "target"
        path = os.path.join(self.log_dir, f"{name}-{now:%Y-%m-%d}.log")
        with self._lock:
            try:
                os.makedirs(self.log_dir, exist_ok=True)
                with open(path, "a", encoding="utf-8") as f:
                    f.write(f"[{now:%Y-%m-%d %H:%M:%S}]{msg}\n")
            except OSError:
                self.file_errors += 1   # 디스크 문제로 프로그램이 멈추지 않도록 무시하고 계속


def setup_program_log(log_dir: Optional[str] = None) -> None:
    """프로그램 오류 기록(log/program-날짜.log)과 예기치 않은 예외/강제 종료 기록."""
    log_dir = log_dir or LOG_DIR
    try:
        os.makedirs(log_dir, exist_ok=True)
        h = logging.FileHandler(os.path.join(log_dir, f"program-{datetime.date.today():%Y-%m-%d}.log"),
                                encoding="utf-8")
        h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(threadName)s: %(message)s"))
        root = logging.getLogger()
        root.addHandler(h)
        root.setLevel(logging.INFO)
        # 파이썬 자체가 비정상 종료될 때 위치 기록
        faulthandler.enable(open(os.path.join(log_dir, "fault.log"), "a"))  # noqa: SIM115
    except OSError:
        pass

    def _hook(exc_type, exc, tb):
        log.critical("처리되지 않은 예외\n%s", "".join(traceback.format_exception(exc_type, exc, tb)))

    def _thread_hook(args):
        if args.exc_type is SystemExit:
            return
        log.critical("쓰레드 %s 처리되지 않은 예외\n%s", getattr(args.thread, "name", "?"),
                     "".join(traceback.format_exception(args.exc_type, args.exc_value, args.exc_traceback)))

    sys.excepthook = _hook
    threading.excepthook = _thread_hook


# ====================================================================== 통신
class Link:
    """TCP 연결 1개. 쓰는 쓰레드는 하나뿐이고, 감시(close)만 다른 쓰레드에서 할 수 있음."""

    def __init__(self, host: str, port: int) -> None:
        self.host, self.port = host, port
        self._sock: Optional[socket.socket] = None
        self._lock = threading.Lock()
        self._splitter = modbus.FrameSplitter()
        self.tid = 0

    @property
    def connected(self) -> bool:
        return self._sock is not None

    def connect(self) -> None:
        self.close()   # 이전 소켓은 반드시 닫음 (원본은 버려두어 연결이 쌓임)
        s = socket.create_connection((self.host, self.port), timeout=CONNECT_TIMEOUT)
        s.settimeout(RECV_POLL)
        with self._lock:
            self._sock = s
            self._splitter = modbus.FrameSplitter()

    def close(self) -> None:
        with self._lock:
            s, self._sock = self._sock, None
        if s is not None:
            try:
                s.shutdown(socket.SHUT_RDWR)
            except OSError:
                pass
            try:
                s.close()
            except OSError:
                pass

    def next_tid(self) -> int:
        self.tid = (self.tid + 1) % 10000   # 원본: 0~9999
        return self.tid

    def send(self, frame: bytes) -> None:
        s = self._sock
        if s is None:
            raise ConnectionError("연결 없음")
        s.sendall(frame)

    def receive(self, seconds: float, handler: Callable[[bytes], None], stop: threading.Event,
                beat: Callable[[], None] = lambda: None) -> None:
        """seconds 동안 받은 응답을 handler 로 넘김. 연결이 끊기면 OSError."""
        deadline = time.monotonic() + seconds
        while not stop.is_set():
            beat()
            if time.monotonic() >= deadline:
                return
            s = self._sock
            if s is None:
                raise ConnectionError("연결 없음")
            try:
                data = s.recv(513)
            except socket.timeout:
                continue
            if not data:
                raise ConnectionError("상대가 연결을 끊음")
            for f in self._splitter.feed(data):
                handler(f)


@dataclass
class MotorStatus:
    uid: int
    reply: bool = True
    automanual: Optional[int] = None
    startstop: Optional[int] = None
    rpm: Optional[int] = None
    last_reply: Optional[datetime.datetime] = None


class _Restart(BaseException):
    """시험용: 쓰레드를 강제로 죽이는 예외"""


class Tester:
    """원본 CMotor_Runner. start()/stop() 은 화면 쓰레드에서, 나머지는 작업 쓰레드에서 실행."""

    def __init__(self, settings: Settings, emit: Callable[[str, str], None]) -> None:
        self.settings = settings
        self._emit_cb = emit
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._target_go = threading.Event()   # 자동 모드: 기존모터 첫 응답 뒤 신규모터 시작 (원본과 같음)
        self._setting_rpm = settings.src_rpm
        self.source_rpm: Optional[int] = None
        self.source_time: Optional[datetime.datetime] = None
        self.motors: Dict[int, MotorStatus] = {i: MotorStatus(i) for i in settings.motor_ids()}
        self._writes: "collections.OrderedDict[Tuple[int, int], Tuple[int, str, int]]" = collections.OrderedDict()
        self._last_write: Optional[Tuple[int, int]] = None   # 방금 보낸 쓰기 (모터, 레지스터)
        self._last_write_rejected = False                    # 그 쓰기에 예외 응답이 왔는지
        self._src = Link(settings.src_ip, settings.src_port)
        self._tgt = Link(settings.tgt_ip, settings.tgt_port)
        self._threads: Dict[str, threading.Thread] = {}
        self._beats: Dict[str, float] = {}
        self._got_source = False
        self.restarts = 0
        self.errors = 0

    # ---------------------------------------------------------------- 공개
    @property
    def running(self) -> bool:
        return bool(self._threads) and not self._stop.is_set()

    @property
    def setting_rpm(self) -> int:
        with self._lock:
            return self._setting_rpm

    def set_rpm(self, rpm: int) -> None:
        """수동수치 [적용] (원본 btn_rpm_apply_Click)"""
        with self._lock:
            self._setting_rpm = max(0, min(0xFFFF, int(rpm)))

    def start(self) -> None:
        self._stop.clear()
        if self.settings.manual:
            self._target_go.set()
        else:
            self._start_thread("source")
        self._start_thread("target")

    def stop(self, timeout: float = 3.0) -> bool:
        """쓰레드를 강제로 끊지 않고(원본 Thread.Abort) 스스로 끝나게 한 뒤 연결을 닫음."""
        self._stop.set()
        self._target_go.set()
        self._src.close()
        self._tgt.close()
        ok = True
        end = time.monotonic() + timeout
        for t in list(self._threads.values()):
            t.join(max(0.0, end - time.monotonic()))
            ok = ok and not t.is_alive()
        self._threads.clear()
        return ok

    def supervise(self) -> List[str]:
        """화면 쓰레드에서 주기적으로 호출. 죽은 쓰레드는 다시 시작, 멈춘 쓰레드는 연결을 끊어 풀어줌."""
        notes: List[str] = []
        if self._stop.is_set():
            return notes
        for name, t in list(self._threads.items()):
            if not t.is_alive():
                self.restarts += 1
                notes.append(f"{name} 작업이 멈춰 다시 시작합니다")
                log.error("%s 쓰레드 종료됨 -> 다시 시작 (%d회)", name, self.restarts)
                self._start_thread(name)
            elif time.monotonic() - self._beats.get(name, time.monotonic()) > STALL_LIMIT:
                notes.append(f"{name} 작업이 {STALL_LIMIT:.0f}초 넘게 진행이 없어 연결을 다시 맺습니다")
                log.error("%s 쓰레드 진행 없음 -> 연결 끊음", name)
                (self._src if name == "source" else self._tgt).close()
                self._beats[name] = time.monotonic()
        return notes

    def snapshot(self) -> dict:
        with self._lock:
            return {
                "setting_rpm": self._setting_rpm,
                "source_rpm": self.source_rpm,
                "source_time": self.source_time,
                "motors": {k: MotorStatus(**vars(v)) for k, v in self.motors.items()},
                "src_connected": self._src.connected,
                "tgt_connected": self._tgt.connected,
            }

    # ---------------------------------------------------------------- 내부
    def _emit(self, target: str, msg: str) -> None:
        try:
            self._emit_cb(target, msg)
        except Exception:  # 로그 출력 문제로 통신이 멈추지 않도록
            log.exception("로그 출력 실패")

    def _start_thread(self, name: str) -> None:
        fn = self._source_loop if name == "source" else self._target_loop
        t = threading.Thread(target=self._guard, args=(name, fn), name=f"motor-{name}", daemon=True)
        self._beats[name] = time.monotonic()
        self._threads[name] = t
        t.start()

    def _guard(self, name: str, fn: Callable[[], None]) -> None:
        try:
            fn()
        except _Restart:
            log.error("%s 쓰레드 강제 종료 (시험)", name)
        except BaseException:  # pragma: no cover - 마지막 안전장치
            log.critical("%s 쓰레드 예외로 종료\n%s", name, traceback.format_exc())

    def _beat(self, name: str) -> Callable[[], None]:
        def f() -> None:
            self._beats[name] = time.monotonic()
        return f

    def _connect(self, link: Link, target: str) -> bool:
        try:
            link.connect()
            log.info("%s 접속 %s:%s", target, link.host, link.port)
            return True
        except (OSError, ValueError, UnicodeError):  # 잘못된 IP 문자열도 여기서 처리
            link.close()
            self._emit(target, "접속실패!")
            self._stop.wait(RETRY_WAIT)
            return False

    def _on_error(self, link: Link, target: str, e: BaseException) -> None:
        link.close()
        if isinstance(e, OSError):
            self._emit(target, f"[E]통신 끊김 ({e})")
        else:
            self.errors += 1
            log.error("%s 처리 중 오류\n%s", target, traceback.format_exc())
            self._emit(target, f"[E]처리 오류 ({type(e).__name__}: {e}) - 다시 접속합니다")
        self._stop.wait(RETRY_WAIT)

    # ---- 기존모터 (원본 src_send_job + src_recv_job)
    def _source_loop(self) -> None:
        link, beat, sid = self._src, self._beat("source"), self.settings.src_id
        while not self._stop.is_set():
            beat()
            try:
                if not link.connected and not self._connect(link, "SOURCE"):
                    continue
                frame = modbus.build_old_rpm_read(link.next_tid(), sid)
                self._got_source = False
                self._emit("SOURCE", f"[ID:{sid:02d}][S]{modbus.to_hex(frame)} (읽기요청)")
                link.send(frame)
                link.receive(SOURCE_REPLY_WAIT, self._on_source_frame, self._stop, beat)
                if self._stop.is_set():
                    break
                if not self._got_source:
                    self._emit("SOURCE", f"[ID:{sid:02d}][E]응답없음")
                link.receive(SOURCE_PAUSE, self._on_source_frame, self._stop, beat)  # 늦은 응답도 처리
            except Exception as e:
                if self._stop.is_set():
                    break
                self._on_error(link, "SOURCE", e)
        link.close()

    def _on_source_frame(self, f: bytes) -> None:
        sid = self.settings.src_id
        r = modbus.parse_old_response(f)
        if r is not None and r.unit_id == sid:
            with self._lock:
                self._setting_rpm = r.rpm
                self.source_rpm = r.rpm
                self.source_time = datetime.datetime.now()
            self._got_source = True
            self._emit("SOURCE", f"[ID:{sid:02d}][R]{modbus.to_hex(f)} (읽기결과-회전수:{r.rpm})")
            self._target_go.set()
        else:
            self._emit("SOURCE", f"[ID:{sid:02d}][R]{modbus.to_hex(f)}")

    # ---- 신규모터 (원본 target_send_job + target_recv_job)
    def _target_loop(self) -> None:
        link, beat = self._tgt, self._beat("target")
        while not self._stop.is_set() and not self._target_go.is_set():
            beat()
            self._target_go.wait(RECV_POLL)
        ids = list(self.motors)
        pos = 0   # 다음에 읽을 모터. 오류로 다시 접속해도 처음부터가 아니라 이어서 진행 (뒤쪽 모터가 굶지 않도록)
        while not self._stop.is_set():
            beat()
            try:
                if not link.connected:
                    if not self._connect(link, "TARGET"):
                        continue
                    with self._lock:
                        self._writes.clear()   # 원본: 접속하면 쓰기 대기열 비움 (해당 모터를 다시 읽어 다시 만듦)
                while pos < len(ids) and not self._stop.is_set():
                    uid = ids[pos]
                    frame = modbus.build_new_read(link.next_tid(), uid)
                    with self._lock:
                        self.motors[uid].reply = False
                    self._emit("TARGET", f"[ID:{uid:02d}][S]{modbus.to_hex(frame)} (읽기요청)")
                    link.send(frame)
                    link.receive(TARGET_READ_GAP, self._on_target_frame, self._stop, beat)
                    self._flush_writes(link, beat)   # 이 모터의 쓰기를 끝낸 뒤 다음 모터로
                    pos += 1
                if self._stop.is_set():
                    break
                pos = 0
                for uid in ids:
                    if not self.motors[uid].reply:
                        self._emit("TARGET", f"[ID:{uid:02d}][E]응답없음")
                link.receive(TARGET_PAUSE, self._on_target_frame, self._stop, beat)
                self._flush_writes(link, beat)       # 쉬는 동안 늦게 온 읽기 응답의 쓰기
            except Exception as e:
                if self._stop.is_set():
                    break
                self._on_error(link, "TARGET", e)
        link.close()

    def _queue_write(self, uid: int, reg: int, value: int, comment: str, attempt: int = 0) -> None:
        with self._lock:
            self._writes[(uid, reg)] = (value, comment, attempt)   # 같은 레지스터는 최신 값 하나만

    def _flush_writes(self, link: Link, beat: Callable[[], None]) -> None:
        while not self._stop.is_set():
            with self._lock:
                if not self._writes:
                    return
                (uid, reg), (value, comment, attempt) = self._writes.popitem(last=False)
            if uid not in self.motors:
                continue
            frame = modbus.build_new_write_single(link.next_tid(), uid, reg, value)
            self._last_write, self._last_write_rejected = (uid, reg), False
            self._emit("TARGET", f"[ID:{uid:02d}][S]{modbus.to_hex(frame)} ({comment}"
                                 + (", 재시도)" if attempt else ")"))
            link.send(frame)
            link.receive(TARGET_WRITE_GAP, self._on_target_frame, self._stop, beat)
            with self._lock:
                retry = self._last_write_rejected and attempt == 0 and (uid, reg) not in self._writes
            if retry:   # 장치가 바쁨 등으로 거절하면 다음 바퀴(약 10초 뒤)까지 기다리지 않고 1번 다시 보냄
                self._queue_write(uid, reg, value, comment, attempt=1)

    def _on_target_frame(self, f: bytes) -> None:
        r = modbus.parse_new_response(f)
        if r is not None:
            uid = r.unit_id
            self._emit("TARGET", f"[ID:{uid:02d}][R]{modbus.to_hex(f)} "
                                 f"(읽기결과-자동수동:{r.automanual}/동작정지:{r.startstop}/회전수:{r.rpm})")
            if uid not in self.motors:
                return
            with self._lock:
                m = self.motors[uid]
                m.reply, m.automanual, m.startstop, m.rpm = True, r.automanual, r.startstop, r.rpm
                m.last_reply = datetime.datetime.now()
                want = self._setting_rpm
            if r.automanual != 1:
                self._queue_write(uid, modbus.NEW_REG_AUTOMANUAL, 1, "쓰기요청-자동수동:1")
            if r.startstop != 1:
                self._queue_write(uid, modbus.NEW_REG_STARTSTOP, 1, "쓰기요청-동작정지:1")
            if r.rpm != want:
                self._queue_write(uid, modbus.NEW_REG_RPM, want, f"쓰기요청-회전수:{want}")
            return
        g = modbus.parse_response(f)
        if g is not None and g.fc == modbus.FC_WRITE_SINGLE and g.exception is None:
            if g.unit_id in self.motors:
                with self._lock:
                    self.motors[g.unit_id].reply = True
            self._emit("TARGET", f"[ID:{g.unit_id:02d}][R]{modbus.to_hex(f)} (쓰기성공)")
        elif g is not None and g.exception is not None:
            # 예외 응답에는 레지스터 주소가 없으므로 방금 보낸 쓰기(같은 모터)에 대한 거절로 봄
            if g.fc == modbus.FC_WRITE_SINGLE and self._last_write and self._last_write[0] == g.unit_id:
                self._last_write_rejected = True
            self._emit("TARGET", f"[ID:{g.unit_id:02d}][R]{modbus.to_hex(f)} (예외응답:{g.exception})")
        else:
            self._emit("TARGET", f"[R]{modbus.to_hex(f)}")


# ====================================================================== 화면 (원본 Form1)
WHITE, PANEL = "#FFFFFF", "#404040"
FONT = ("굴림", 9)


class App:
    def __init__(self, root: "tk.Tk", sink: Optional[LogSink] = None, config_file: str = "") -> None:
        self.root = root
        self.sink = sink or LogSink()
        self.config_file = config_file or CONFIG_FILE
        self.tester: Optional[Tester] = None
        root.title(TITLE)
        root.geometry("856x545")
        root.resizable(False, False)
        root.configure(bg=PANEL)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.report_callback_exception = self._ui_error
        self._build()
        self.load_settings(Settings.load(self.config_file))
        self._poll_id = root.after(UI_POLL_MS, self._poll)

    # ---------------------------------------------------------------- 구성
    def _label(self, parent, text, x, y, bg=PANEL, **kw):
        lb = tk.Label(parent, text=text, fg=WHITE, bg=bg, font=FONT, **kw)
        lb.place(x=x, y=y)
        return lb

    def _spin(self, parent, x, y, w, lo, hi, value):
        var = tk.StringVar(value=str(value))
        sp = tk.Spinbox(parent, from_=lo, to=hi, textvariable=var, font=FONT, justify="right")
        sp.place(x=x, y=y, width=w, height=21)
        sp.var, sp.lo, sp.hi = var, lo, hi  # type: ignore[attr-defined]
        return sp

    @staticmethod
    def spin_value(sp, default: int = 0) -> int:
        v = _clamp_int(sp.var.get(), default, sp.lo, sp.hi)
        sp.var.set(str(v))
        return v

    def _build(self) -> None:
        r = self.root
        top = tk.Frame(r, bg=PANEL, height=108, width=856)
        top.place(x=0, y=0)
        # ---- RPM원본 (GroupBox1)
        g1 = tk.Frame(top, bg=PANEL, highlightthickness=1, highlightbackground="#8A8A8A")
        g1.place(x=9, y=4, width=409, height=98)
        tk.Label(g1, text="RPM\n원본", fg=WHITE, bg="#0000FF", font=FONT).place(x=2, y=5, width=37, height=86)
        self.mode = tk.IntVar(value=0)
        radio = dict(fg=WHITE, bg=PANEL, activebackground=PANEL, activeforeground=WHITE, selectcolor=PANEL,
                     disabledforeground="#B0B0B0", font=FONT, padx=0, pady=0)
        # 원본은 라디오버튼 옆 글자(Label5)를 눌러도 선택됨 -> 글자를 라디오버튼에 포함
        self.rb_auto = tk.Radiobutton(g1, text="정상모터 -", variable=self.mode, value=0, **radio)
        self.rb_auto.place(x=42, y=20)
        self._label(g1, "IP:", 127, 24)
        self.ent_src_ip = tk.Entry(g1, font=FONT)
        self.ent_src_ip.place(x=150, y=19, width=80, height=21)
        self._label(g1, "PORT:", 234, 24)
        self.sp_src_port = self._spin(g1, 278, 19, 54, 1, 65535, 502)
        self._label(g1, "ID:", 335, 24)
        self.sp_src_id = self._spin(g1, 356, 19, 46, 1, ID_MAX, 1)
        self.rb_manual = tk.Radiobutton(g1, text="수동수치 -", variable=self.mode, value=1, **radio)
        self.rb_manual.place(x=42, y=57)
        self._label(g1, "RPM:", 126, 61)
        self.sp_src_rpm = self._spin(g1, 163, 57, 67, RPM_MIN, RPM_MAX, 0)
        self.btn_apply = tk.Button(g1, text="적용", font=FONT, command=self.on_apply)
        self._label(g1, "(0~800, 0:정지)", 287, 62)
        # ---- 신규모터 (GroupBox2)
        g2 = tk.Frame(top, bg=PANEL, highlightthickness=1, highlightbackground="#8A8A8A")
        g2.place(x=423, y=4, width=329, height=98)
        tk.Label(g2, text="신규\n모터", fg=WHITE, bg="#C00000", font=FONT).place(x=3, y=5, width=40, height=86)
        self._label(g2, "IP:", 51, 20)
        self.ent_tgt_ip = tk.Entry(g2, font=FONT)
        self.ent_tgt_ip.place(x=77, y=15, width=108, height=21)
        self._label(g2, "PORT:", 201, 19)
        self.sp_tgt_port = self._spin(g2, 250, 15, 57, 1, 65535, 502)
        self._label(g2, "ID:", 51, 54)
        self.sp_tgt_ids = [self._spin(g2, x, 50, 46, 0, ID_MAX, 1 if n == 0 else 0)
                           for n, x in enumerate((77, 125, 173, 221, 269))]
        self._label(g2, "( 0 : 사용안함 )", 226, 76)
        # ---- 시작/정지
        self.btn_start = tk.Button(top, text="▶", font=("굴림", 20), command=self.on_start)
        self.btn_start.place(x=758, y=8, width=87, height=90)
        # ---- 로그 탭
        nb = ttk.Notebook(r)
        nb.place(x=0, y=109, width=856, height=414)
        self.txt_src = self._log_tab(nb, "원본통신LOG")
        self.txt_tgt = self._log_tab(nb, "신규모터통신LOG")
        # ---- 상태 줄 (새로 추가)
        self.status = tk.Label(r, text="정지", anchor="w", font=FONT, bg="#E8E8E8")
        self.status.place(x=0, y=523, width=856, height=22)
        self._inputs = [self.rb_auto, self.rb_manual, self.ent_src_ip, self.sp_src_port, self.sp_src_id,
                        self.sp_src_rpm, self.ent_tgt_ip, self.sp_tgt_port] + self.sp_tgt_ids

    def _log_tab(self, nb, title):
        fr = tk.Frame(nb)
        txt = tk.Text(fr, font=FONT, wrap="none", state="disabled", undo=False)
        sb = tk.Scrollbar(fr, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)
        nb.add(fr, text=title)
        return txt

    # ---------------------------------------------------------------- 설정 <-> 화면
    def load_settings(self, s: Settings) -> None:
        self.mode.set(1 if s.manual else 0)
        for ent, v in ((self.ent_src_ip, s.src_ip), (self.ent_tgt_ip, s.tgt_ip)):
            ent.delete(0, "end")
            ent.insert(0, v)
        self.sp_src_port.var.set(str(s.src_port))
        self.sp_src_id.var.set(str(s.src_id))
        self.sp_src_rpm.var.set(str(s.src_rpm))
        self.sp_tgt_port.var.set(str(s.tgt_port))
        for sp, v in zip(self.sp_tgt_ids, s.tgt_ids):
            sp.var.set(str(v))

    def read_settings(self) -> Settings:
        return Settings(
            manual=self.mode.get() == 1,
            src_ip=self.ent_src_ip.get().strip(), src_port=self.spin_value(self.sp_src_port, 502),
            src_id=self.spin_value(self.sp_src_id, 1), src_rpm=self.spin_value(self.sp_src_rpm, 0),
            tgt_ip=self.ent_tgt_ip.get().strip(), tgt_port=self.spin_value(self.sp_tgt_port, 502),
            tgt_ids=[self.spin_value(sp, 0) for sp in self.sp_tgt_ids],
        )

    # ---------------------------------------------------------------- 동작
    @staticmethod
    def _ip_ok(text: str) -> bool:
        try:
            ipaddress.ip_address(text)
            return True
        except ValueError:
            return False

    def on_start(self) -> None:
        if self.tester is not None:
            self.stop()
            return
        s = self.read_settings()
        if not s.motor_ids():
            messagebox.showwarning(TITLE, "신규모터 ID값을 설정하세요.", parent=self.root)
            return
        checks = [("신규모터", s.tgt_ip)] + ([] if s.manual else [("RPM원본", s.src_ip)])
        bad = [name for name, ip in checks if not self._ip_ok(ip)]
        if bad:
            messagebox.showwarning(TITLE, f"{', '.join(bad)} IP 형식이 올바르지 않습니다.", parent=self.root)
            return
        s.save(self.config_file)
        for txt in (self.txt_src, self.txt_tgt):
            txt.configure(state="normal")
            txt.delete("1.0", "end")
            txt.configure(state="disabled")
        self.tester = Tester(s, self.sink)
        log.info("시작 %s", s)
        self.tester.start()
        self._set_running(True, s.manual)

    def stop(self) -> None:
        t, self.tester = self.tester, None
        if t is not None:
            if not t.stop():
                log.error("작업 쓰레드가 제때 끝나지 않음")
            log.info("정지")
        self._set_running(False, self.mode.get() == 1)

    def on_apply(self) -> None:   # 원본 btn_rpm_apply_Click
        if self.tester is not None:
            v = self.spin_value(self.sp_src_rpm, 0)
            self.tester.set_rpm(v)
            self.sink("TARGET", f"[수동] 설정 회전수 변경: {v}")

    def _set_running(self, running: bool, manual: bool) -> None:   # 원본 make_disable / make_enable
        for w in self._inputs:
            w.configure(state="disabled" if running else "normal")
        if running and manual:
            self.sp_src_rpm.configure(state="normal")
            self.btn_apply.place(x=237, y=55, width=44, height=23)
        else:
            self.btn_apply.place_forget()
        self.btn_start.configure(text="■" if running else "▶", fg="red" if running else "black")

    def on_close(self) -> None:   # 원본 Form1_FormClosing
        try:
            if self.tester is not None:
                self.stop()
            self.read_settings().save(self.config_file)
        finally:
            try:
                self.root.after_cancel(self._poll_id)
            except (tk.TclError, ValueError):
                pass
            self.root.destroy()

    # ---------------------------------------------------------------- 주기 처리
    def _append(self, txt, lines: List[str]) -> None:
        txt.configure(state="normal")
        at_end = txt.yview()[1] >= 0.999
        txt.insert("end", "\n".join(lines) + "\n")
        n = int(txt.index("end-1c").split(".")[0])
        if n > UI_MAX_LINES:
            txt.delete("1.0", f"{n - UI_MAX_LINES}.0")
        txt.configure(state="disabled")
        if at_end:
            txt.see("end")

    def _poll(self) -> None:
        try:
            src: List[str] = []
            tgt: List[str] = []
            for _ in range(1000):
                try:
                    target, line = self.sink.ui.popleft()
                except IndexError:
                    break
                (src if target == "SOURCE" else tgt).append(line)
            if src:
                self._append(self.txt_src, src)
            if tgt:
                self._append(self.txt_tgt, tgt)
            if self.tester is not None:
                for note in self.tester.supervise():
                    self.sink("TARGET", f"[감시] {note}")
            self._update_status()
        except Exception:
            log.exception("화면 갱신 오류")
        finally:
            self._poll_id = self.root.after(UI_POLL_MS, self._poll)

    def _update_status(self) -> None:
        t = self.tester
        if t is None:
            self.status.configure(text="정지")
            return
        snap = t.snapshot()
        parts = [f"설정 회전수 {snap['setting_rpm']}"]
        if not t.settings.manual:
            st = snap["source_time"]
            parts.append("원본 " + (f"{snap['source_rpm']} RPM ({st:%H:%M:%S})" if st else "응답 대기"))
        ok = sum(1 for m in snap["motors"].values() if m.reply and m.last_reply)
        parts.append(f"신규모터 응답 {ok}/{len(snap['motors'])}")
        if t.restarts:
            parts.append(f"작업 재시작 {t.restarts}회")
        self.status.configure(text="  |  ".join(parts))

    def _ui_error(self, exc_type, exc, tb) -> None:
        log.error("화면 처리 오류\n%s", "".join(traceback.format_exception(exc_type, exc, tb)))


def main() -> None:
    setup_program_log()
    log.info("프로그램 시작")
    root = tk.Tk()
    App(root)
    root.mainloop()
    log.info("프로그램 종료")


if __name__ == "__main__":
    main()
