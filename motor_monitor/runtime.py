"""실행 환경: 중복 실행 방지, 로그 설정."""
from __future__ import annotations

import logging
import logging.handlers
import os
import sys
import tempfile
from typing import Optional


class SingleInstance:
    """잠금 파일로 중복 실행을 막음 (원본: 같은 이름의 프로세스가 있으면 종료)."""

    def __init__(self, name: str = "bldc_motor_monitor") -> None:
        self.path = os.path.join(tempfile.gettempdir(), f"{name}.lock")
        self._fh = None

    def acquire(self) -> bool:
        try:
            fh = open(self.path, "a+")
        except OSError:
            return True  # 잠금 파일을 못 만들면 실행은 허용
        try:
            if sys.platform.startswith("win"):
                import msvcrt
                fh.seek(0)
                msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            fh.close()
            return False
        self._fh = fh
        return True

    def release(self) -> None:
        if self._fh is not None:
            try:
                self._fh.close()
            finally:
                self._fh = None


def setup_logging(log_dir: Optional[str], level: int = logging.INFO) -> None:
    root = logging.getLogger()
    root.setLevel(level)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    console = logging.StreamHandler()
    console.setFormatter(fmt)
    root.addHandler(console)
    if log_dir:
        try:
            os.makedirs(log_dir, exist_ok=True)
            fh = logging.handlers.RotatingFileHandler(
                os.path.join(log_dir, "motor_monitor.log"),
                maxBytes=2 * 1024 * 1024, backupCount=5, encoding="utf-8")
            fh.setFormatter(fmt)
            root.addHandler(fh)
        except OSError:
            root.warning("로그 파일을 만들 수 없습니다: %s", log_dir)
