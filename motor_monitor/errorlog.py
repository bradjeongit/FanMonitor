"""에러 기록: 통신 끊김, 응답 없음, 쓰기 실패 등을 모아 화면의 [에러 로그] 창에 보여줌.

- 최근 기록은 메모리에(최대 MAX_ENTRIES 건), 모든 기록은 logs/error.log 파일에 남김 (2MB x 5개 순환)
- 프로그램을 다시 켜면 파일의 최근 기록을 다시 읽어 창에 보여줌
- 같은 문제가 계속되어도 처음 한 번과 복구 한 번만 기록하도록 호출하는 쪽에서 상태 변화 때만 add() 함
"""
from __future__ import annotations

import collections
import datetime
import logging
import logging.handlers
import os
import threading
from dataclasses import dataclass
from typing import Deque, List, Optional

MAX_ENTRIES = 1000
FILE_NAME = "error.log"
_SEP = " | "

ERROR = "에러"
RECOVER = "복구"


@dataclass(frozen=True)
class ErrorEntry:
    time: datetime.datetime
    level: str          # 에러 / 복구
    zone: str           # '1구역' 등 (없으면 '-')
    target: str         # '기존모터 ID 5', '신규모터 통신장치 192.168.0.2:502' 등
    message: str

    def line(self) -> str:
        return _SEP.join((self.time.strftime("%Y-%m-%d %H:%M:%S"), self.level, self.zone, self.target,
                          self.message.replace("\n", " ")))

    @classmethod
    def parse(cls, line: str) -> Optional["ErrorEntry"]:
        parts = line.rstrip("\r\n").split(_SEP, 4)
        if len(parts) != 5:
            return None
        try:
            t = datetime.datetime.strptime(parts[0], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None
        return cls(t, parts[1], parts[2], parts[3], parts[4])


class ErrorLog:
    def __init__(self, log_dir: Optional[str] = None, max_entries: int = MAX_ENTRIES) -> None:
        self._lock = threading.Lock()
        self._entries: Deque[ErrorEntry] = collections.deque(maxlen=max_entries)
        self.version = 0                  # 기록이 늘 때마다 증가 (화면 갱신용)
        self.seen_version = 0             # [에러 로그] 창에서 확인한 시점
        self.path: Optional[str] = None
        self._file: Optional[logging.Logger] = None
        if log_dir:
            self._open_file(log_dir)

    def _open_file(self, log_dir: str) -> None:
        try:
            os.makedirs(log_dir, exist_ok=True)
            self.path = os.path.join(log_dir, FILE_NAME)
            self._load_previous()
            logger = logging.getLogger(f"motor_monitor.errorfile.{id(self)}")
            logger.propagate = False
            logger.setLevel(logging.INFO)
            h = logging.handlers.RotatingFileHandler(self.path, maxBytes=2 * 1024 * 1024, backupCount=5,
                                                     encoding="utf-8")
            h.setFormatter(logging.Formatter("%(message)s"))
            logger.addHandler(h)
            self._file = logger
        except OSError:
            self._file = None

    def _load_previous(self) -> None:
        """지난 실행의 기록을 다시 읽음 (파일 끝부분만)."""
        if not self.path or not os.path.exists(self.path):
            return
        try:
            with open(self.path, encoding="utf-8", errors="replace") as f:
                lines = collections.deque(f, maxlen=self._entries.maxlen)
        except OSError:
            return
        for line in lines:
            e = ErrorEntry.parse(line)
            if e is not None:
                self._entries.append(e)

    def add(self, level: str, zone: str, target: str, message: str) -> ErrorEntry:
        e = ErrorEntry(datetime.datetime.now().replace(microsecond=0), level, zone or "-", target or "-", message)
        with self._lock:
            self._entries.append(e)
            self.version += 1
        if self._file is not None:
            try:
                self._file.info(e.line())
            except Exception:  # 기록 실패로 통신이 멈추지 않도록
                pass
        return e

    def error(self, zone: str, target: str, message: str) -> ErrorEntry:
        return self.add(ERROR, zone, target, message)

    def recover(self, zone: str, target: str, message: str) -> ErrorEntry:
        return self.add(RECOVER, zone, target, message)

    def recent(self) -> List[ErrorEntry]:
        """최근 기록 (최신이 앞)."""
        with self._lock:
            return list(reversed(self._entries))

    @property
    def unseen(self) -> int:
        """창에서 확인하지 않은 새 에러 수 (복구 기록 제외)."""
        with self._lock:
            n = self.version - self.seen_version
            if n <= 0:
                return 0
            return sum(1 for e in list(self._entries)[-n:] if e.level == ERROR)

    def mark_seen(self) -> None:
        with self._lock:
            self.seen_version = self.version

    def clear(self) -> None:
        """화면 목록 비우기 (파일 기록은 그대로)."""
        with self._lock:
            self._entries.clear()
            self.version += 1
            self.seen_version = self.version
