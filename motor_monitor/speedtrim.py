"""팬 속도 보정과 기억. 신규모터(0x16 지정 / 0x05 실제)와 기존모터제어(h2 지정 / i14 실제)에 같이 씀.

용어
- 설정 속도(target): 모터가 돌아야 할 속도 (신규모터: 구역 목표 RPM, 속도제한 적용 / 기존모터제어: 운전속도)
- 지정값(cmd): 모터에 보내는 값. 보정하면 설정 속도보다 클 수 있음 (최대값 이하)
- 실제 속도(actual): 모터가 보고하는 회전수

점검(CHECK_INTERVAL, 30초)마다 한 번 decide() 를 부름:
1. 실제 속도가 설정 속도의 ±1%(SPEED_TOLERANCE_PCT) 이내면 오차로 인정 - 그대로 두고 그때의 지정값을 기억(learned)
2. 1% 넘게 느리면 지정값을 조금씩(설정 속도의 2~3%) 올림. 최대값(cmd_max)은 넘지 않음
3. 올린 뒤 실제 속도가 1% 넘게 바뀌지 않으면 더 올리지 않고 기다림. SPEED_STALL_CHECKS 번(60초) 동안 그대로면
   보정을 멈추고(더 보내지 않음) 알림(note). 최대값까지 올렸는데도 느려도 멈추고 알림
4. 1% 넘게 빠르고 지정값을 올려 둔 상태면 조금씩 내림 (설정 속도 아래로는 내리지 않음)
5. 설정 속도가 1% 넘게 바뀌면 새로 시작. 기억한 값이 있으면 같은 비율(지정값/설정 속도)로 첫 지정값을 정함
   (설정 속도가 기억할 때와 1% 이내로 같으면 기억한 지정값 그대로). 1% 이내로 바뀌면 다시 보내지 않음
6. 설정 속도를 아직 모를 때(참조값 없음)의 최초 신호는 기억한 지정값 (first_signal)
"""
from __future__ import annotations

import datetime
import json
import logging
import os
import tempfile
import threading
import time
from dataclasses import dataclass
from typing import Dict, Optional, Tuple

from . import constants as C

log = logging.getLogger(__name__)

Memory = Tuple[int, int]      # 기억한 (설정 속도, 지정값)


def tolerance(target: int) -> float:
    """오차로 인정하는 차이 (RPM): 설정 속도의 SPEED_TOLERANCE_PCT %."""
    return abs(target) * C.SPEED_TOLERANCE_PCT / 100.0


def within(actual: int, target: int) -> bool:
    return abs(actual - target) <= tolerance(target)


@dataclass
class Trim:
    """모터 1대의 보정 상태 (그 모터를 맡은 통신 쓰레드 하나만 바꿈)."""
    target: Optional[int] = None     # 지금 맞추는 설정 속도 (1% 이내 변화는 그대로 둠)
    cmd: Optional[int] = None        # 지정값
    base: Optional[int] = None       # 마지막으로 올리기 직전의 실제 속도 (반응 확인용)
    waiting: int = 0                 # 올린 뒤 반응을 기다린 점검 횟수
    hold: str = ""                   # 보정을 멈춘 이유: "stall" 반응 없음 / "max" 최대값 (있으면 더 보내지 않음)
    note: str = ""                   # 속도 문제 알림 (없으면 빈 문자열)
    learned: bool = False            # 이번 점검에서 오차 안에 들었는지 (기억할 값)

    def first_signal(self, memory: Optional[Memory]) -> Optional[int]:
        """설정 속도를 아직 모를 때: 처음 한 번만 기억한 지정값을 보냄. 그 뒤로는 그 값을 유지."""
        if self.cmd is None and memory is not None:
            self.cmd = max(0, min(int(memory[1]), C.MAX_RPM))
        return self.cmd

    def decide(self, target: int, actual: Optional[int], cmd_max: int,
               memory: Optional[Memory] = None) -> int:
        """점검 1번. 보낼 지정값을 돌려줌 (모터의 지금 지정값과 다를 때만 쓰면 됨).
        actual 이 None 이면(못 읽음, 방금 운전 방식을 바꿈 등) 판단하지 않고 지금 지정값 유지."""
        cmd_max = max(0, int(cmd_max))
        target = max(0, min(int(target), cmd_max))
        self.learned = False
        if self.target is None or abs(target - self.target) > tolerance(self.target):
            self._start(target, cmd_max, memory)               # 5. 새 설정 속도 (첫 점검 포함)
            return self.cmd
        self.cmd = min(self.cmd if self.cmd is not None else self.target, cmd_max)
        t = self.target
        if actual is None or t == 0:
            return self.cmd
        tol = tolerance(t)
        if abs(actual - t) <= tol:                                # 1. 오차 인정 -> 기억
            self.base, self.waiting, self.hold, self.note = None, 0, "", ""
            self.learned = True
            return self.cmd
        if actual > t + tol:                                      # 4. 빠름
            self.base, self.waiting, self.hold, self.note = None, 0, "", ""
            if self.cmd > t:
                self.cmd = max(t, self.cmd - self._step(t, actual - t))
            return self.cmd
        # 2·3. 느림
        if self.hold:
            return self.cmd
        if self.base is not None:                                 # 올린 뒤 반응 확인
            if actual - self.base > tol:
                self.base, self.waiting = None, 0                 # 반응함 -> 아직 느리면 더 올림
            else:
                self.waiting += 1
                if self.waiting >= C.SPEED_STALL_CHECKS:
                    self.hold = "stall"
                    self.note = (f"설정 {t} RPM 보다 느림 (실제 {actual} RPM). 지정값을 {self.cmd} RPM 까지 올렸으나 "
                                 f"{C.SPEED_STALL_CHECKS * C.CHECK_INTERVAL:g}초 동안 속도가 바뀌지 않아 보정 멈춤")
                return self.cmd                                   # 반응을 기다림 (다시 보내지 않음)
        if self.cmd >= cmd_max:
            self.hold = "max"
            self.note = f"설정 {t} RPM 보다 느림 (실제 {actual} RPM). 지정값이 최대 {cmd_max} RPM 이라 더 올리지 못함"
            return self.cmd
        self.base = actual
        self.cmd = min(self.cmd + self._step(t, t - actual), cmd_max)
        return self.cmd

    def _start(self, target: int, cmd_max: int, memory: Optional[Memory]) -> None:
        self.target = target
        self.base, self.waiting, self.hold, self.note = None, 0, "", ""
        cmd = target
        if memory is not None and target > 0:
            m_target, m_cmd = memory
            if m_target > 0 and m_cmd >= m_target:
                if abs(target - m_target) <= tolerance(m_target):
                    cmd = m_cmd                                   # 같은 설정 속도: 기억한 지정값 그대로
                else:
                    cmd = round(target * m_cmd / m_target)        # 다른 설정 속도: 같은 보정 비율
        self.cmd = max(target, min(cmd, cmd_max)) if target > 0 else 0

    @staticmethod
    def _step(target: int, diff: float) -> int:
        """한 번에 바꾸는 양: 차이만큼, 단 설정 속도의 2~3% 범위 (최소 1 RPM)."""
        lo = target * C.SPEED_STEP_MIN_PCT / 100.0
        hi = target * C.SPEED_STEP_MAX_PCT / 100.0
        return max(1, round(min(max(diff, lo), hi)))


class SpeedMemory:
    """오차 안에 든 지정값 기억 -> speed_memory.json. {키: {"target": 설정 속도, "cmd": 지정값, "time": 시각}}
    키 = '구역코드/모터키' (예: 'H1R1/1_5'). 바뀔 때만 표시해 두고 30초에 한 번까지만 파일에 씀(종료 시 반드시 씀)."""

    SAVE_EVERY = 30.0

    def __init__(self, path: Optional[str]) -> None:
        self.path = path
        self._lock = threading.Lock()
        self._data: Dict[str, dict] = {}
        self._dirty = False
        self._saved_at = 0.0
        self._load()

    def _load(self) -> None:
        if not self.path or not os.path.exists(self.path):
            return
        try:
            with open(self.path, encoding="utf-8") as f:
                raw = json.load(f)
            for k, v in raw.items():
                t, c = int(v["target"]), int(v["cmd"])
                if 0 < t <= C.MAX_RPM and t <= c <= C.MAX_RPM:
                    self._data[str(k)] = {"target": t, "cmd": c, "time": str(v.get("time", ""))}
        except (OSError, ValueError, KeyError, TypeError, AttributeError) as e:
            log.warning("속도 기억 파일을 읽지 못함 (새로 시작): %s (%s)", self.path, e)
            self._data = {}

    def get(self, key: str) -> Optional[Memory]:
        with self._lock:
            v = self._data.get(key)
            return (v["target"], v["cmd"]) if v else None

    def remember(self, key: str, target: int, cmd: int) -> None:
        if target <= 0 or cmd < target:
            return
        with self._lock:
            v = self._data.get(key)
            if v and (v["target"], v["cmd"]) == (target, cmd):
                return
            self._data[key] = {"target": int(target), "cmd": int(cmd),
                               "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            self._dirty = True
        self.save(force=False)

    def save(self, force: bool = True) -> None:
        """기억한 값을 파일에 씀 (force=False 이면 마지막 저장 후 SAVE_EVERY 초가 지났을 때만)."""
        if not self.path:
            return
        with self._lock:
            if not self._dirty or (not force and time.monotonic() - self._saved_at < self.SAVE_EVERY):
                return
            data = json.dumps(self._data, ensure_ascii=False, indent=1, sort_keys=True)
            self._dirty = False
            self._saved_at = time.monotonic()
        folder = os.path.dirname(os.path.abspath(self.path))
        tmp = None
        try:
            os.makedirs(folder, exist_ok=True)
            fd, tmp = tempfile.mkstemp(prefix=".speed_", suffix=".tmp", dir=folder)
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(data)
            os.replace(tmp, self.path)
        except OSError as e:
            log.warning("속도 기억 파일 저장 실패: %s (%s)", self.path, e)
            if tmp and os.path.exists(tmp):
                try:
                    os.unlink(tmp)
                except OSError:
                    pass
            with self._lock:
                self._dirty = True
