"""ECblue 통신 공용부 (미니 제어기가 사용). 예전 기존모터 시험기(ecblue_tester.py)의 통신부를 그대로 옮김.

- Bus: 통신장치와 TCP 연결 1개 (작업 쓰레드에서만 사용). request() -> (응답, 설명 문구)
- mode_text / error_text: 제어모드(h4)·오류(i12) 값을 사람이 읽는 문구로
"""
from __future__ import annotations

import logging
import os
import socket
import sys
import time
from typing import Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from motor_monitor import modbus  # noqa: E402

TIMEOUT = 2.0               # 응답 대기 (초)
REQUEST_GAP = 0.1           # 요청 사이 간격 (초)

EXCEPTIONS = {1: "지원하지 않는 기능", 2: "없는 레지스터", 3: "허용되지 않는 값", 4: "장치 오류",
              6: "장치 사용 중", 10: "게이트웨이 경로 없음", 11: "게이트웨이: 모터 응답 없음"}
ERROR_BITS = {15: "통신 워치독", 14: "기동 오류", 13: "온도 오류", 12: "안전 정지", 11: "사인필터",
              10: "피크 전류", 9: "모터 구속", 8: "홀센서", 7: "TB", 6: "전원 결상/정전",
              5: "입력전압 낮음", 4: "입력전압 높음", 3: "DC링크 전압 낮음", 2: "DC링크 전압 높음",
              1: "지락", 0: "IGBT 오류"}

log = logging.getLogger("ecblue_bus")


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
