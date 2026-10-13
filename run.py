"""BLDC MOTOR MONITOR 실행 파일.  사용법: python run.py  [--config 경로] [--debug]"""
import sys

from motor_monitor.app import main

if __name__ == "__main__":
    sys.exit(main())
