"""프로그램 시작점."""
from __future__ import annotations

import argparse
import logging
import os
import sys
import tkinter as tk
from tkinter import messagebox

from . import __version__
from .config import app_base_dir
from .plant import Plant
from .runtime import SingleInstance, setup_logging


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="BLDC MOTOR MONITOR")
    ap.add_argument("--config", help="설정 파일 경로 (기본: 실행 폴더의 Config.ini)")
    ap.add_argument("--debug", action="store_true", help="통신 상세 로그 출력")
    args = ap.parse_args(argv)

    setup_logging(os.path.join(app_base_dir(), "logs"), logging.DEBUG if args.debug else logging.INFO)
    log = logging.getLogger("motor_monitor")
    log.info("BLDC MOTOR MONITOR v%s 시작", __version__)

    root = tk.Tk()
    root.withdraw()

    instance = SingleInstance()
    if not instance.acquire():
        messagebox.showinfo("BLDC MOTOR MONITOR", "모니터가 이미 실행중입니다.", parent=root)
        root.destroy()
        return 1

    try:
        plant = Plant(args.config)
        plant.load()
        from .ui.main_window import MainWindow
        MainWindow(root, plant)
        root.deiconify()
        root.mainloop()
    finally:
        instance.release()
    return 0


if __name__ == "__main__":
    sys.exit(main())
