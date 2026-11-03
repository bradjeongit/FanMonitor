"""구역 단위 모터 관리 및 Modbus TCP 송수신. 기준: 운영 실행파일 C_HOUSE

- 구역 1개 = 기존모터 게이트웨이 1개(SOURCE) + 신규모터 게이트웨이 1개(TARGET)
- 기존모터: RPM 을 읽어 저장
- 신규모터: 1페이지(현재 RPM)와 2페이지(수동/동작/RPM설정)를 읽고, 2페이지 값이 목표와 다르면 써넣음
  목표 RPM = 수동RPM 또는 참조 기존모터 RPM, 속도제한 사용 시 최대속도 이하로 제한
- 상태: 응답 수신 = 정상, 명령 후 60초 무응답 또는 통신장치 접속 실패 = '이상표시'(빨간 !),
  마지막 응답 후 1시간 = 이상. 다시 응답이 오면 바로 정상 표시로 돌아감
- 기존모터제어(OLD_CTRL): 모터마다 지정한 IP/ID 의 ECblue 기존 모터에 직접 접속(PORT 502 고정)해
  h4(제어모드=버스 절대회전수) → h1(Enable) → h2(운전속도) 순으로 맞추고, i14(실제 회전수)를 읽음.
  운전속도 = 수동 지정값, 또는 구역 참조값: 이 구역 신규모터가 따르는 값(수동 RPM 또는 참조 기존모터 회전수)을
  최고속도 MAX_RPM 대비 % 로 바꿔 같은 비율(MAX_RPM × %)로 지정. 값이 없으면 이전 속도를 유지
- 점검 주기(CHECK_INTERVAL, 30초): 모터마다 30초에 한 번 읽고 필요하면 맞춤. 설정을 바꾸거나 제어 모드로
  바꾸면(kick) 기다리지 않고 바로 한 번 점검. 쓰기 명령은 생기는 대로 바로 보냄
- 속도 지정: 신규모터·기존모터제어에는 목표 속도(설정 속도)를 그대로 보냄 (모터가 가진 값과 다를 때만).
  실제 속도로 지정값을 바꾸는 보정은 하지 않음 (2.5.0 의 속도 보정·기억은 2.5.1 에서 삭제)
- 에러 기록(plant.errors): 접속 실패·연결 끊김·응답 없음·쓰기 실패와 그 복구를 상태가 바뀔 때만 기록
- 모니터링 모드(plant.monitoring): 읽기만 하고 쓰기(제어) 명령은 만들지도 보내지도 않음
  (화면에서 실행 중에 바꿀 수 있으며, 매 명령마다 현재 모드를 다시 확인)
"""
from __future__ import annotations

import collections
import logging
import socket
import threading
import time
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Callable, Deque, Dict, List, Optional, Tuple

from . import constants as C
from . import modbus
from .config import MotorEntry
from .constants import CTRL_MODE_MANUAL, CTRL_MODE_REF, MotorStatus, MotorType
from .helpers import get_house_name

if TYPE_CHECKING:
    from .plant import Plant

log = logging.getLogger("motor_monitor.house")

# 디버그 출력 대상 (운영 DEBUG_LINE / DEBUG_IDS). 로그 레벨이 DEBUG 일 때만 출력됩니다.
DEBUG_LINE = "ALL"          # 구역코드 또는 "ALL"
DEBUG_IDS = {5, 50}         # 출력할 모터 ID (0 은 구역 공통 메시지로 항상 출력)


def motor_key(mtype: MotorType, mid: int, ip: str = "", port: int = C.MODBUS_PORT) -> str:
    """모터 고유 키 '종류_ID' (예: '0_5'). 기존모터제어는 접속처가 달라 '2_ID@IP:PORT' (PORT 는 502 고정)."""
    if MotorType(mtype) == MotorType.OLD_CTRL:
        return f"{int(mtype)}_{mid}@{ip}:{port}"
    return f"{int(mtype)}_{mid}"


@dataclass
class Motor:
    id: int
    mtype: MotorType
    use: bool
    use_limit: bool = False           # 속도제한 사용 (신규모터)
    limit_rpm: int = 0                # 속도제한 최대 RPM
    rpm: int = 0                      # 표시 RPM (기존: 현재 RPM, 신규: 레지스터 0x05)
    rdate: Optional[float] = None     # 마지막 응답 시각 (time.monotonic)
    check_since: Optional[float] = None  # 응답 대기 시작 시각 (명령 전송 시 기록, 응답 시 해제)
    had_err: bool = False             # 이상표시: 60초 무응답·접속 실패 (다시 응답이 오면 해제)
    err_note: str = ""                # 이상표시 이유 (툴팁/에러 로그)
    status: MotorStatus = MotorStatus.READY
    must_send_zero: bool = False      # 사용중지 선택 시 RPM 0 을 1회 보내기 위한 표시
    ip: str = ""                      # 기존모터제어: 접속 IP
    port: int = C.MODBUS_PORT         # 기존모터제어: 접속 PORT (502 고정, 화면에 없음)
    set_rpm: int = 0                  # 기존모터제어: 수동 지정 운전 속도
    ctrl_mode: str = CTRL_MODE_MANUAL  # 기존모터제어: manual 수동 지정 / ref 구역 참조
    ref_desc: str = ""                # 기존모터제어(구역 참조): 구역 참조값 출처 (예: '2구역 기존모터 66')
    ref_rpm: Optional[int] = None     # 기존모터제어(구역 참조): 마지막으로 확인한 구역 참조값 (RPM)
    ref_wait_since: Optional[float] = None  # 기존모터제어(구역 참조): 구역 참조값이 없어진 시각 (60초 넘으면 이상표시)
    ctrl_target: Optional[int] = None  # 기존모터제어: 마지막으로 정한 운전속도 (구역 참조값이 없을 때 유지)
    mode_tries: int = 0               # 기존모터제어: 제어모드(h4, NV) 설정 시도 횟수
    ctrl_note: str = ""               # 기존모터제어: 마지막 제어 문제 설명 (있으면 이상표시)
    # 속도 점검 (신규모터·기존모터제어, 표시용)
    target_rpm: Optional[int] = None  # 설정 속도 (이 모터가 돌아야 할 속도)
    cmd_rpm: Optional[int] = None     # 지정값 (모터에 보낸/모터가 가진 값: 신규모터 0x16, 기존모터제어 h2)
    speed_note: str = ""              # 속도 문제 (기존모터제어: 모터 최고속도 h8 이 설정 속도보다 낮음). 있으면 이상 행
    max_rpm: int = 0                  # 기존모터제어: 모터 최고속도 h8 (1/min, 0 = 설정 없음/모름)
    sent_at: Optional[float] = None   # 속도 변경을 보낸 시각 (화면에서 아이콘을 붉게 깜빡임)

    @property
    def has_problem(self) -> bool:
        """화면에서 행에 음영을 칠할 문제 상태: 이상표시(통신·제어), 이상(오래 응답 없음), 속도 문제."""
        return self.use and (self.had_err or self.status == MotorStatus.ERR or bool(self.speed_note))

    @property
    def key(self) -> str:
        return motor_key(self.mtype, self.id, self.ip, self.port)

    def entry(self) -> MotorEntry:
        return MotorEntry(self.id, self.mtype, self.use, self.use_limit, self.limit_rpm,
                          self.ip, self.port, self.set_rpm, self.ctrl_mode)

    @property
    def label(self) -> str:
        """에러 로그용 이름 (예: '신규모터 ID 5', '기존모터제어 ID 180 (172.27.0.95)')"""
        name = C.MOTOR_TYPE_NAME[self.mtype]
        if self.mtype == MotorType.OLD_CTRL:
            return f"{name} ID {self.id} ({self.ip})"
        return f"{name} ID {self.id}"


@dataclass
class _QueueItem:
    dkey: str
    id: int
    msg: bytes
    speed: bool = False               # 속도 변경 쓰기 (보내면 화면 아이콘 깜빡임)


class _Channel:
    """TCP 소켓 1개(기존모터용 또는 신규모터용)."""

    def __init__(self, name: str) -> None:
        self.name = name
        self.sock: Optional[socket.socket] = None
        self.connected = False
        self.net_status = False      # 통신상태 True 정상 / False 이상
        self.down = False            # 접속 실패/끊김을 에러 로그에 기록한 상태 (복구 때 한 번 기록)
        self._lock = threading.Lock()

    def connect(self, ip: str, port: int, timeout: float) -> bool:
        self.close()
        try:
            s = socket.create_connection((ip, port), timeout=timeout)
        except (OSError, ValueError):
            return False
        s.settimeout(2.0)
        with self._lock:
            self.sock = s
            self.connected = True
        return True

    def send(self, data: bytes) -> None:
        s = self.sock
        if s is None:
            raise OSError("not connected")
        s.sendall(data)

    def close(self) -> None:
        with self._lock:
            s, self.sock = self.sock, None
            self.connected = False
        if s is not None:
            try:
                s.shutdown(socket.SHUT_RDWR)
            except OSError:
                pass
            try:
                s.close()
            except OSError:
                pass


class _CtrlLinks:
    """기존모터제어 접속처(IP:PORT)별 연결. 접속 실패 시 일정 시간 뒤 다시 시도, 끊김/복구를 에러 로그에 기록."""

    def __init__(self, house: "HouseLine") -> None:
        self.house = house
        self.links: Dict[tuple, _Channel] = {}
        self.retry_at: Dict[tuple, float] = {}
        self.down: set = set()
        self._tid = 0

    def next_tid(self) -> int:
        self._tid = (self._tid + 1) % (C.CMD_IDX_MAX + 1)
        return self._tid

    def get(self, ep: tuple) -> Optional[_Channel]:
        """연결된 채널. 접속할 수 없으면 None (다음 시도 시각 전에는 시도하지 않음)."""
        link = self.links.setdefault(ep, _Channel(f"C{ep[0]}:{ep[1]}"))
        if link.connected:
            return link
        if time.monotonic() < self.retry_at.get(ep, 0):
            return None
        if not link.connect(ep[0], ep[1], C.CONNECT_CHECK_TIMEOUT):
            log.info("[%s][기존모터제어] 접속 실패 %s:%s (%d초 후 재시도)",
                     self.house.code, ep[0], ep[1], C.RECONNECT_DELAY)
            self.retry_at[ep] = time.monotonic() + C.RECONNECT_DELAY
            self._down(ep, f"{ep[0]} 접속 실패 ({C.RECONNECT_DELAY}초마다 재시도)")
            return None
        log.info("[%s][기존모터제어] 접속 성공 %s:%s", self.house.code, ep[0], ep[1])
        if ep in self.down:
            self.down.discard(ep)
            self.house._log_recover("기존모터제어 통신장치", f"{ep[0]} 다시 접속됨")
        return link

    def _down(self, ep: tuple, msg: str) -> None:
        if ep not in self.down:
            self.down.add(ep)
            self.house._log_error("기존모터제어 통신장치", msg)

    def drop(self, ep: tuple, delay: float, why: str) -> None:
        link = self.links.get(ep)
        if link is not None:
            link.close()
        self.retry_at[ep] = time.monotonic() + delay
        self._down(ep, f"{ep[0]} {why}")

    def close_all(self) -> None:
        for link in self.links.values():
            link.close()


class HouseLine:
    """구역 1개. 운영 C_HOUSE 클래스."""

    def __init__(self, code: str, plant: "Plant", clock: Callable[[], float] = time.monotonic) -> None:
        self.code = code
        self.plant = plant
        self._clock = clock
        self._lock = threading.RLock()

        self._motors: Dict[str, Motor] = {}   # 등록 순서 유지 (설정 저장 순서)
        self._order: List[str] = []           # 화면 표시 순서
        self._sort_key = ""
        self._sort_desc = False
        self.version = 0                      # 데이터가 바뀔 때마다 증가 (화면 갱신 판단용)

        self.source = _Channel("S")
        self.target = _Channel("T")
        self._src_cmd_idx = 0
        self._tgt_cmd_idx = 0
        self._src_read_q: Deque[_QueueItem] = collections.deque()
        self._tgt_read_q: Deque[_QueueItem] = collections.deque()
        self._tgt_write_q: Deque[_QueueItem] = collections.deque()

        self._stop = threading.Event()
        self._stop.set()
        self._threads: Dict[str, threading.Thread] = {}
        self._kick_gen = 0                    # 늘어나면 각 통신 쓰레드가 30초를 기다리지 않고 바로 점검

    # ------------------------------------------------------------------ 상태
    @property
    def source_net_status(self) -> bool:
        return self.source.net_status

    @property
    def target_net_status(self) -> bool:
        return self.target.net_status

    @property
    def running(self) -> bool:
        """구역 제어 중 여부 (운영 RUN_OX)"""
        return not self._stop.is_set()

    def _touch(self) -> None:
        self.version += 1

    def kick(self) -> None:
        """설정이 바뀜: 다음 점검(30초)을 기다리지 않고 바로 한 번 점검."""
        self._kick_gen += 1

    def _debug(self, mid: int, msg: str) -> None:
        if DEBUG_LINE in ("ALL", self.code) and (mid == 0 or mid in DEBUG_IDS):
            log.debug(msg)

    # ------------------------------------------------------------ 모터 목록
    def get_list(self) -> List[Motor]:
        """화면 표시 순서대로 모터 정보 사본 목록."""
        with self._lock:
            return [replace(self._motors[k]) for k in self._order if k in self._motors]

    def get_motor(self, dkey: str) -> Optional[Motor]:
        with self._lock:
            m = self._motors.get(dkey)
            return replace(m) if m else None

    def get_old_motor_ids(self) -> List[str]:
        """등록된 기존모터 ID 목록(등록 순서). 운영 Get_OLDMOTOR_List"""
        with self._lock:
            return [str(m.id) for m in self._motors.values() if m.mtype == MotorType.OLD_MOTOR]

    def export_motors(self) -> List[MotorEntry]:
        """설정 저장용 목록 - 등록 순서 (운영 Get_INI_MOTORINFO 와 같은 순서)."""
        with self._lock:
            return [m.entry() for m in self._motors.values()]

    def add_motor(self, mid: int, mtype: MotorType, use: bool, use_limit: bool = False, limit_rpm: int = 0,
                  ip: str = "", port: int = C.MODBUS_PORT, set_rpm: int = 0,
                  ctrl_mode: str = CTRL_MODE_MANUAL) -> bool:
        key = motor_key(mtype, mid, ip, port)
        with self._lock:
            if key in self._motors:
                return False
            status = MotorStatus.READY if use else MotorStatus.NOTUSE
            self._motors[key] = Motor(id=mid, mtype=MotorType(mtype), use=use, use_limit=use_limit,
                                      limit_rpm=limit_rpm, status=status, ip=ip, port=port, set_rpm=set_rpm,
                                      ctrl_mode=ctrl_mode)
            self._order.append(key)
            self._resort()
            self._touch()
        self.kick()
        return True

    def remove_motor(self, dkey: str) -> bool:
        with self._lock:
            if dkey not in self._motors:
                return False
            del self._motors[dkey]
            if dkey in self._order:
                self._order.remove(dkey)
            self._resort()
            self._touch()
        return True

    def add_entry(self, e: MotorEntry) -> bool:
        return self.add_motor(e.id, e.mtype, e.use, e.use_limit, e.limit_rpm, e.ip, e.port or C.MODBUS_PORT,
                              e.set_rpm, e.ctrl_mode)

    def clear_motors(self) -> None:
        with self._lock:
            self._motors.clear()
            self._order.clear()
            self._touch()

    def mod_motor(self, dkey: str, use: bool, use_limit: bool = False, limit_rpm: int = 0,
                  set_rpm: Optional[int] = None, ctrl_mode: Optional[str] = None) -> None:
        """사용여부/속도제한(/기존모터제어 운전속도·속도 지정 방법) 수정. 운영 Mod_Motor"""
        with self._lock:
            m = self._motors.get(dkey)
            if m is None:
                return
            if use != m.use:
                m.check_since = None
                if use:
                    m.status = MotorStatus.READY
                    m.must_send_zero = False
                else:
                    m.status = MotorStatus.NOTUSE
                    m.rdate = None
                    m.must_send_zero = True   # 사용 중지로 바꾸면 RPM 0 을 한 번 보냄
                m.use = use
            m.use_limit = use_limit
            m.limit_rpm = limit_rpm
            if set_rpm is not None:
                m.set_rpm = set_rpm
            if ctrl_mode is not None:
                if ctrl_mode != m.ctrl_mode:
                    m.ctrl_target = None
                    m.ref_rpm = None
                    m.ref_wait_since = None
                m.ctrl_mode = ctrl_mode
            self._touch()
        self.kick()                          # 바뀐 설정을 30초 기다리지 않고 바로 반영

    def sort(self, skey: str) -> None:
        """목록 정렬. 같은 항목을 다시 누르면 오름/내림 전환."""
        with self._lock:
            if self._sort_key != skey:
                self._sort_key = skey
                self._sort_desc = False
            else:
                self._sort_desc = not self._sort_desc
            self._resort()
            self._touch()

    @property
    def sort_state(self) -> tuple:
        return self._sort_key, self._sort_desc

    def _resort(self) -> None:
        """운영 Update_BIND_LIST(): 정렬 기준에 따라 표시 순서 재구성 (안정 정렬).
        기본은 ID 오름차순 (같은 ID 는 기존모터 → 신규모터 → 기존모터제어). 운영 프로그램 기본은 종류순이었음."""
        if self._sort_key == "":
            self._sort_key = "ID"
            self._sort_desc = False
        keys = list(self._motors.keys())
        ms = self._motors
        desc = self._sort_desc
        if self._sort_key == "MTYPE":
            if desc:
                val = lambda k: int(ms[k].mtype) * 1000 + 1000 - ms[k].id
            else:
                val = lambda k: int(ms[k].mtype) * 1000 + ms[k].id
            keys = sorted(keys, key=val, reverse=desc)
        elif self._sort_key == "ID":
            keys = sorted(keys, key=lambda k: (ms[k].id, int(ms[k].mtype)), reverse=desc)
        elif self._sort_key == "STATUS":
            keys = sorted(keys, key=lambda k: int(ms[k].status), reverse=desc)
        self._order = keys

    def update_status(self) -> bool:
        """모터 상태 갱신(5초 주기, 제어 중일 때만 호출). 운영 TMR_STATUS_UPDATER_TICK."""
        now = self._clock()
        changed = False
        with self._lock:
            for m in self._motors.values():
                old = (m.status, m.had_err)
                if m.use:
                    if m.check_since is not None and int(now - m.check_since) >= C.NO_RESPONSE_WARN_SEC:
                        m.check_since = None
                        self._set_err(m, f"명령 후 {C.NO_RESPONSE_WARN_SEC}초 동안 응답 없음")
                    if m.rdate is None:
                        m.status = MotorStatus.READY
                    else:
                        status = (MotorStatus.ERR if int(now - m.rdate) >= C.MOTOR_TIMEOUT_SEC
                                  else MotorStatus.OK)
                        if status == MotorStatus.ERR and m.status != MotorStatus.ERR:
                            self._log_error(m.label, f"마지막 응답 후 {C.MOTOR_TIMEOUT_SEC // 60}분 넘게 응답 없음 (이상)")
                        m.status = status
                else:
                    m.status = MotorStatus.NOTUSE
                if (m.status, m.had_err) != old:
                    changed = True
            if changed:
                self._touch()
        return changed

    # --- 이상표시 / 에러 기록 ---
    @property
    def zone_name(self) -> str:
        return get_house_name(self.code)

    def _log_error(self, target: str, msg: str) -> None:
        log.warning("[%s] %s: %s", self.code, target, msg)
        el = getattr(self.plant, "errors", None)
        if el is not None:
            el.error(self.zone_name, target, msg)

    def _log_recover(self, target: str, msg: str) -> None:
        log.info("[%s] %s: %s", self.code, target, msg)
        el = getattr(self.plant, "errors", None)
        if el is not None:
            el.recover(self.zone_name, target, msg)

    def _set_err(self, m: Motor, note: str) -> None:
        """이상표시 켜기 (lock 안에서 호출). 이유가 바뀔 때만 에러 기록."""
        if not m.had_err or m.err_note != note:
            self._log_error(m.label, note)
        m.had_err = True
        m.err_note = note
        self._touch()

    def _clear_err(self, m: Motor) -> None:
        """다시 응답이 오면 이상표시 해제 (기존모터제어의 제어 문제가 남아 있으면 유지)."""
        if m.had_err and not m.ctrl_note:
            m.had_err = False
            m.err_note = ""
            self._log_recover(m.label, "응답 정상 (이상표시 해제)")
            self._touch()

    def _set_speed_note(self, m: Motor, note: str) -> None:
        """속도 문제 표시 켜기/끄기 (lock 안에서 호출). 모터 최고속도가 설정 속도보다 낮을 때. 바뀔 때만 기록."""
        if note == m.speed_note:
            return
        if note:
            self._log_error(m.label, note)
        else:
            self._log_recover(m.label, "속도 문제 해제")
        m.speed_note = note
        self._touch()

    def _mark_sent(self, key: str) -> None:
        """속도 변경을 보냄 (화면 아이콘 깜빡임)."""
        with self._lock:
            cur = self._motors.get(key)
            if cur is not None:
                cur.sent_at = time.monotonic()
                self._touch()

    def _mark_channel_down(self, ch: _Channel, mtype: MotorType, msg: str) -> None:
        """통신장치에 접속할 수 없거나 연결이 끊김: 그 장치의 사용 중인 모터를 이상표시 (한 번만 기록)."""
        if not ch.down:
            ch.down = True
            self._log_error(f"{C.MOTOR_TYPE_NAME[mtype]} 통신장치", msg)
        with self._lock:
            for m in self._motors.values():
                if m.mtype == mtype and m.use and not m.had_err:
                    m.had_err = True
                    m.err_note = msg
                    self._touch()

    def _mark_channel_up(self, ch: _Channel, mtype: MotorType, msg: str) -> None:
        if ch.down:
            ch.down = False
            self._log_recover(f"{C.MOTOR_TYPE_NAME[mtype]} 통신장치", msg)

    def _mark_received(self, m: Motor, rpm: Optional[int] = None) -> None:
        if rpm is not None:
            m.rpm = rpm
        m.check_since = None
        m.rdate = self._clock()
        m.status = MotorStatus.OK
        self._clear_err(m)
        self._touch()

    # ------------------------------------------------------------- 서비스
    def service_start(self) -> None:
        """구역 제어 시작. 운영 Service_Start (모든 모터의 이상기록/상태 초기화)"""
        self.service_stop(join=True)
        self._debug(0, f"[{self.code}] Start Service!")
        log.info("[%s] %s 시작", self.code, "모니터링(읽기 전용)" if self.monitoring else "제어")
        self._src_cmd_idx = 0
        self._tgt_cmd_idx = 0
        self._src_read_q.clear()
        self._tgt_read_q.clear()
        self._tgt_write_q.clear()
        with self._lock:
            for m in self._motors.values():
                m.mode_tries = 0
                m.had_err = False
                m.err_note = ""
                m.ctrl_note = ""
                m.ctrl_target = None
                m.ref_rpm = None
                m.ref_wait_since = None
                m.check_since = None
                m.rdate = None
                m.target_rpm = None
                m.cmd_rpm = None
                m.speed_note = ""
                m.status = MotorStatus.READY if m.use else MotorStatus.NOTUSE
            self._touch()
        for ch in (self.source, self.target):
            ch.net_status = False
            ch.down = False
        stop = threading.Event()
        self._stop = stop
        self._threads = {}
        self._spawn("src_send", self._src_send_job, stop)
        self._spawn("tgt_send", self._tgt_send_job, stop)
        self._spawn("ctrl", self._ctrl_job, stop)

    def service_stop(self, join: bool = False) -> None:
        """구역 제어 종료. 운영 Service_Stop"""
        was_running = not self._stop.is_set()
        self._stop.set()
        self._src_read_q.clear()
        self._tgt_read_q.clear()
        self._tgt_write_q.clear()
        self.source.close()
        self.target.close()
        if join:
            for t in list(self._threads.values()):
                if t is not threading.current_thread():
                    t.join(timeout=3)
        if was_running:
            log.info("[%s] 제어 종료", self.code)

    def service_restart(self) -> None:
        self.service_stop(join=True)
        time.sleep(C.RESTART_DELAY)
        self.service_start()

    def _spawn(self, name: str, target, stop: threading.Event) -> None:
        t = threading.Thread(target=target, args=(stop,), name=f"{self.code}-{name}", daemon=True)
        self._threads[name] = t
        t.start()

    def _ensure_recv(self, name: str, target, stop: threading.Event) -> None:
        t = self._threads.get(name)
        if t is None or not t.is_alive():
            self._spawn(name, target, stop)

    def _connect(self, ch: _Channel, ip: str, port: int, stop: threading.Event, mtype: MotorType) -> bool:
        if ch.connected:
            return True
        self._debug(0, f"[{self.code}][{ch.name}] 접속시도")
        if ch.connect(ip, port, C.CONNECT_CHECK_TIMEOUT):
            log.info("[%s][%s] 접속 성공 %s:%s", self.code, ch.name, ip, port)
            self._mark_channel_up(ch, mtype, f"{ip} 다시 접속됨")
            return True
        log.info("[%s][%s] 접속 실패 %s:%s (%d초 후 재시도)", self.code, ch.name, ip, port, C.RECONNECT_DELAY)
        ch.net_status = False
        self._mark_channel_down(ch, mtype, f"{ip} 접속 실패 ({C.RECONNECT_DELAY}초마다 재시도)")
        stop.wait(C.RECONNECT_DELAY)
        return False

    def _channel_lost(self, ch: _Channel, mtype: MotorType, why: str) -> None:
        """연결이 끊김: 소켓 정리 후 다음 주기에 다시 접속."""
        cfg = self.plant.configs[self.code]
        ip = cfg.source_ip if mtype == MotorType.OLD_MOTOR else cfg.target_ip
        ch.close()
        ch.net_status = False
        self._mark_channel_down(ch, mtype, f"{ip} 연결 끊김 ({why})")

    def _next_src_idx(self) -> int:
        i = self._src_cmd_idx
        self._src_cmd_idx = 0 if i + 1 > C.CMD_IDX_MAX else i + 1
        return i

    def _next_tgt_idx(self) -> int:
        i = self._tgt_cmd_idx
        self._tgt_cmd_idx = 0 if i + 1 > C.CMD_IDX_MAX else i + 1
        return i

    def _send_item(self, ch: _Channel, item: _QueueItem, tag: str, stop: threading.Event) -> None:
        with self._lock:
            m = self._motors.get(item.dkey)
            if m is None:
                return
        ch.send(item.msg)
        with self._lock:
            m = self._motors.get(item.dkey)
            if m is not None and m.check_since is None:
                m.check_since = self._clock()     # 응답 대기 시작
            if m is not None and item.speed:
                m.sent_at = time.monotonic()      # 속도 변경을 보냄 -> 화면 아이콘 깜빡임
                self._touch()
        self._debug(item.id, f"[{self.code}][{tag}][{item.id}] {modbus.to_hex(item.msg)}")
        stop.wait(C.SEND_INTERVAL)

    # --- 점검 주기 (CHECK_INTERVAL) ---
    def _new_round_state(self) -> dict:
        return {"next": 0.0, "kick": self._kick_gen}    # 시작하면 바로 첫 점검

    def _round_due(self, rnd: dict) -> bool:
        """점검 차례인지 (30초마다, kick() 하면 바로). 차례면 다음 차례 시각을 정함. rnd: 쓰레드별 상태."""
        now = time.monotonic()
        if now >= rnd["next"] or self._kick_gen != rnd["kick"]:
            rnd["next"] = now + C.CHECK_INTERVAL
            rnd["kick"] = self._kick_gen
            return True
        return False

    @staticmethod
    def _idle(stop: threading.Event, rnd: dict) -> None:
        """다음 점검까지 짧게 나눠 기다림 (그 사이 생긴 쓰기 명령·kick 을 바로 처리하도록)."""
        stop.wait(max(0.01, min(C.CYCLE_INTERVAL, rnd["next"] - time.monotonic())))

    # --- 기존모터 (SOURCE) ---
    def _src_send_job(self, stop: threading.Event) -> None:
        self._debug(0, f"[{self.code}][S] 전송 시작")
        ch = self.source
        rnd = self._new_round_state()
        while not stop.is_set():
            cfg = self.plant.configs[self.code]
            if not self._connect(ch, cfg.source_ip, cfg.source_port, stop, MotorType.OLD_MOTOR):
                continue
            ch.net_status = True
            self._ensure_recv("src_recv", self._src_recv_job, stop)
            try:
                if self._src_read_q:
                    for _ in range(len(self._src_read_q)):
                        if stop.is_set():
                            return
                        self._send_item(ch, self._src_read_q.popleft(), "S", stop)
                elif self._round_due(rnd):
                    # 점검 차례: 사용 중인 기존모터마다 RPM 읽기 1번 (30초마다)
                    with self._lock:
                        targets = [(m.key, m.id) for m in self._motors.values()
                                   if m.mtype == MotorType.OLD_MOTOR and m.use]
                    for dkey, mid in targets:
                        msg = modbus.build_old_rpm_read(self._next_src_idx(), mid)
                        self._src_read_q.append(_QueueItem(dkey, mid, msg))
                else:
                    self._idle(stop, rnd)
            except (OSError, IndexError) as e:
                if not stop.is_set():
                    self._channel_lost(ch, MotorType.OLD_MOTOR, f"전송 오류: {e}")
                stop.wait(C.ERROR_RETRY_DELAY)
        ch.close()

    def _src_recv_job(self, stop: threading.Event) -> None:
        self._recv_loop(self.source, stop, self._handle_src_frame, MotorType.OLD_MOTOR)

    def _handle_src_frame(self, frame: bytes) -> None:
        r = modbus.parse_old_response(frame)
        if r is None:
            return
        self._debug(r.unit_id, f"[{self.code}][S] {r.unit_id}={r.rpm}")
        with self._lock:
            m = self._motors.get(motor_key(MotorType.OLD_MOTOR, r.unit_id))
            if m is not None:
                self._mark_received(m, r.rpm)

    # --- 신규모터 (TARGET) ---
    @property
    def monitoring(self) -> bool:
        return bool(getattr(self.plant, "monitoring", False))

    def _enqueue_write(self, item: _QueueItem) -> None:
        """쓰기(제어) 명령 적재. 모니터링 모드에서는 적재하지 않음 (쓰기 차단의 단일 관문)."""
        if self.monitoring:
            self._debug(item.id, f"[{self.code}][모니터링] 쓰기 생략 {modbus.to_hex(item.msg)}")
            return
        self._tgt_write_q.append(item)

    def _flush_writes(self, ch: _Channel, stop: threading.Event) -> bool:
        if self.monitoring:
            self._tgt_write_q.clear()      # 이중 안전장치: 모니터링 모드에서는 어떤 쓰기도 전송하지 않음
            return True
        for _ in range(len(self._tgt_write_q)):
            if stop.is_set():
                return False
            self._send_item(ch, self._tgt_write_q.popleft(), "W", stop)
        return True

    def _tgt_send_job(self, stop: threading.Event) -> None:
        self._debug(0, f"[{self.code}][T] 전송 시작")
        ch = self.target
        rnd = self._new_round_state()
        while not stop.is_set():
            cfg = self.plant.configs[self.code]
            if not self._connect(ch, cfg.target_ip, cfg.target_port, stop, MotorType.NEW_MOTOR):
                continue
            ch.net_status = True
            self._ensure_recv("tgt_recv", self._tgt_recv_job, stop)
            try:
                if self._tgt_read_q:
                    for _ in range(len(self._tgt_read_q)):
                        if stop.is_set():
                            return
                        # 쓰기(제어) 명령이 있으면 먼저 모두 전송
                        if not self._flush_writes(ch, stop):
                            return
                        self._send_item(ch, self._tgt_read_q.popleft(), "T", stop)
                elif self._tgt_write_q:
                    # 마지막 응답 뒤에 생긴 쓰기 명령은 다음 점검을 기다리지 않고 바로 전송
                    # (사용중인 신규모터가 없어 읽기 명령이 없을 때도 전송 - 운영 프로그램은 쌓이기만 했음)
                    self._flush_writes(ch, stop)
                elif self._round_due(rnd):
                    # 점검 차례 (30초마다): 사용 중인 신규모터마다 1페이지(실제 RPM)·2페이지(설정) 읽기 1번
                    with self._lock:
                        for m in self._motors.values():
                            if m.mtype != MotorType.NEW_MOTOR:
                                continue
                            if m.use:
                                for build in (modbus.build_new_read_page1, modbus.build_new_read):
                                    msg = build(self._next_tgt_idx(), m.id)
                                    self._tgt_read_q.append(_QueueItem(m.key, m.id, msg))
                            elif m.must_send_zero and not self.monitoring:
                                # 사용중지된 신규모터는 RPM 0 을 1회 전송
                                # (모니터링 모드에서는 보류했다가 제어 모드로 바뀌면 전송)
                                m.must_send_zero = False
                                msg = modbus.build_new_rpm_write(self._next_tgt_idx(), m.id, 0)
                                self._enqueue_write(_QueueItem(m.key, m.id, msg, speed=True))
                else:
                    self._idle(stop, rnd)
            except (OSError, IndexError) as e:
                if not stop.is_set():
                    self._channel_lost(ch, MotorType.NEW_MOTOR, f"전송 오류: {e}")
                stop.wait(C.ERROR_RETRY_DELAY)
        ch.close()

    def _tgt_recv_job(self, stop: threading.Event) -> None:
        self._recv_loop(self.target, stop, self._handle_tgt_frame, MotorType.NEW_MOTOR)

    def resolve_setting_rpm(self) -> Optional[int]:
        """신규모터 목표 RPM(속도제한 적용 전). 참조 모터가 정상이 아니면 None(제어 보류)."""
        cfg = self.plant.configs[self.code]
        if cfg.source_manual:
            self._debug(0, f"[{self.code}][T]-참조 수동:{cfg.manual_rpm}")
            return cfg.manual_rpm
        src_house = self.plant.houses.get(cfg.source_line)
        if src_house is None:
            return None
        sm = src_house.get_motor(motor_key(MotorType.OLD_MOTOR, cfg.source_id))
        if sm is None:
            # 운영 프로그램은 참조 모터가 없으면 빈 정보(상태 정상, RPM 0)를 받아 0 으로 제어함.
            # 의도하지 않은 정지를 막기 위해 '상태이상'과 같이 제어를 보류합니다.
            self._debug(0, "참조모터 없음")
            return None
        self._debug(0, f"[{self.code}][T]-참조 :{cfg.source_line}/{cfg.source_id}/{sm.rpm}")
        if sm.status == MotorStatus.OK:
            return sm.rpm
        self._debug(0, "참조모터 상태이상")
        return None

    def zone_reference(self) -> Tuple[str, Optional[int]]:
        """이 구역 신규모터가 따르는 값 (출처 설명, 현재 RPM 또는 None). 기존모터제어 '구역 참조'가 그대로 사용."""
        cfg = self.plant.configs[self.code]
        if cfg.source_manual:
            desc = f"구역 수동 {cfg.manual_rpm} RPM"
        else:
            desc = f"{get_house_name(cfg.source_line) or cfg.source_line or '구역 없음'} 기존모터 {cfg.source_id}"
        return desc, self.resolve_setting_rpm()

    def _handle_tgt_frame(self, frame: bytes) -> None:
        p1 = modbus.parse_new_page1(frame)
        if p1 is not None:
            with self._lock:
                m = self._motors.get(motor_key(MotorType.NEW_MOTOR, p1.unit_id))
                if m is not None:
                    self._mark_received(m, p1.actual_rpm)
            return
        r = modbus.parse_new_response(frame)
        if r is None:
            return
        dkey = motor_key(MotorType.NEW_MOTOR, r.unit_id)
        with self._lock:
            m = self._motors.get(dkey)
            if m is None:
                return
            m.cmd_rpm = r.rpm                     # 모터가 가진 지정값(0x16) - 표시용
            snap = replace(m)
        # 설정 속도 = 구역 목표 RPM, 속도제한 사용 시 최대속도 이하
        setting = self.resolve_setting_rpm()
        if setting is not None and snap.use_limit and setting > snap.limit_rpm:
            setting = snap.limit_rpm              # 속도제한: 최대속도 이하로
        with self._lock:
            cur = self._motors.get(dkey)
            if cur is not None:
                cur.target_rpm = setting          # 표시용
        if setting is not None and not self.monitoring:
            mid = r.unit_id
            if r.automanual != 1:
                self._enqueue_write(_QueueItem(dkey, mid, modbus.build_new_manual_write(self._next_tgt_idx(), mid)))
            if r.startstop != 1:
                self._enqueue_write(_QueueItem(dkey, mid, modbus.build_new_run_write(self._next_tgt_idx(), mid)))
            self._debug(mid, f"[{self.code}][{mid}]-제어RPM:{setting}")
            if r.rpm != setting:                  # 목표를 그대로 (모터가 가진 값과 다를 때만)
                self._enqueue_write(
                    _QueueItem(dkey, mid, modbus.build_new_rpm_write(self._next_tgt_idx(), mid, setting), speed=True))
                with self._lock:
                    cur = self._motors.get(dkey)
                    if cur is not None:
                        cur.cmd_rpm = setting
        with self._lock:
            m = self._motors.get(dkey)
            if m is not None:
                self._mark_received(m)

    # --- 기존모터제어 (OLD_CTRL, ZIEHL-ABEGG ECblue) ---
    def _ctrl_job(self, stop: threading.Event) -> None:
        """기존모터제어 모터들을 접속처(IP:PORT)별 연결 하나로 차례대로 읽고 맞춤."""
        links = _CtrlLinks(self)
        rnd = self._new_round_state()
        try:
            while not stop.is_set():
                if not self._round_due(rnd):            # 30초마다 한 번 (kick 이면 바로)
                    self._idle(stop, rnd)
                    continue
                with self._lock:
                    motors = [replace(m) for m in self._motors.values() if m.mtype == MotorType.OLD_CTRL]
                if not motors:
                    continue
                groups: Dict[tuple, List[Motor]] = {}
                for m in motors:
                    groups.setdefault((m.ip, m.port), []).append(m)
                for ep, group in groups.items():
                    if stop.is_set():
                        break
                    if not any(m.use or m.must_send_zero for m in group):
                        continue
                    link = links.get(ep)
                    if link is None:
                        with self._lock:
                            for m in group:
                                cur = self._motors.get(m.key)
                                if cur is not None and cur.use and not cur.had_err:
                                    cur.had_err, cur.err_note = True, f"{ep[0]} 접속 실패"
                                    self._touch()
                        continue
                    for m in group:
                        if stop.is_set():
                            break
                        try:
                            self._ctrl_cycle(links, link, m, stop)
                        except OSError as e:
                            links.drop(ep, C.ERROR_RETRY_DELAY, f"연결 끊김 ({e})")
                            break
                        except Exception:
                            # 예상 밖 오류로 이 구역의 기존모터제어가 멈추지 않도록 기록 후 연결을 새로 맺음
                            # (응답 순서가 어긋났을 수 있어 연결을 정리)
                            log.exception("[%s][기존모터제어 ID%s %s] 처리 오류", self.code, m.id, m.ip)
                            links.drop(ep, C.ERROR_RETRY_DELAY, "처리 오류 - 다시 접속")
                            break
        finally:
            links.close_all()

    def _ctrl_request(self, link: _Channel, frame: bytes, stop: threading.Event) -> Optional[modbus.Response]:
        """요청 1건 전송 후 같은 트랜잭션 번호의 응답을 기다림 (최대 2초). 응답 없으면 None."""
        link.send(frame)
        tid = frame[0] * 256 + frame[1]
        splitter = modbus.FrameSplitter()
        deadline = time.monotonic() + 2.0
        resp = None
        sock = link.sock
        while sock is not None and time.monotonic() < deadline and not stop.is_set():
            try:
                data = sock.recv(513)
            except socket.timeout:
                break
            if not data:
                raise OSError("connection closed")
            for f in splitter.feed(data):
                r = modbus.parse_response(f)
                if r is not None and r.tid == tid:
                    resp = r
            if resp is not None:
                break
        stop.wait(C.SEND_INTERVAL)
        return resp

    def _ctrl_write(self, link: _Channel, m: Motor, reg: int, value: int, stop: threading.Event,
                    links: "_CtrlLinks") -> bool:
        """쓰기 1건. 모니터링 모드에서는 절대 보내지 않음. 정상 응답(echo)이면 True."""
        if self.monitoring:
            self._debug(m.id, f"[{self.code}][모니터링] 기존모터제어 쓰기 생략 h{reg}={value}")
            return False
        r = self._ctrl_request(link, modbus.build_new_write_single(links.next_tid(), m.id, reg, value), stop)
        if r is None or r.exception is not None:
            log.warning("[%s][기존모터제어 ID%s %s] h%s=%s 쓰기 실패 (%s)", self.code, m.id, m.ip, reg, value,
                        "응답 없음" if r is None else f"예외코드 {r.exception}")
            return False
        return True

    def _ctrl_target(self, m: Motor) -> tuple:
        """운전속도 결정 -> (RPM 또는 None, 문제 설명).
        수동 지정: 운전속도.
        구역 참조: 이 구역 신규모터가 따르는 값(구역설정의 수동 RPM 또는 참조 기존모터 회전수)을 그대로 써서
        최고속도 MAX_RPM 대비 % 를 구하고, 같은 비율(MAX_RPM × %)로 지정. 미니 제어기와 같은 계산(100% = MAX_RPM).
        값이 없으면(참조 기존모터를 아직 읽지 못함·상태 이상 등) 이전에 정한 속도를 유지(없으면 None = 속도를 쓰지 않음).
        값이 없는 상태가 NO_RESPONSE_WARN_SEC(60초) 넘게 이어질 때만 이상표시 (구역을 막 시작해 참조 기존모터를
        아직 읽기 전에는 기다림 - 신규모터의 '60초 응답 없음' 과 같은 기준).
        참조 기존모터가 이 모터 자신이면 따라가지 않음(자기 회전수를 목표로 하면 속도가 그대로 굳음)."""
        if m.ctrl_mode != CTRL_MODE_REF:
            return max(0, min(int(m.set_rpm), C.MAX_RPM)), ""
        desc, value = self.zone_reference()      # 다른 구역의 잠금을 쓰므로 이 구역 잠금 밖에서 읽음
        own = self.plant.reference_endpoint(self.code) == (m.ip, m.port, m.id)
        now = self._clock()
        waited = 0.0
        with self._lock:
            cur = self._motors.get(m.key)
            prev = cur.ctrl_target if cur is not None else None
            if cur is not None:
                shown = (cur.ref_desc, cur.ref_rpm)
                cur.ref_desc = desc
                cur.ref_rpm = None if own else value
                if (cur.ref_desc, cur.ref_rpm) != shown:
                    self._touch()                # 화면 갱신은 바뀔 때만
                if value is not None and not own:
                    cur.ref_wait_since = None
                else:
                    if cur.ref_wait_since is None:
                        cur.ref_wait_since = now
                    waited = now - cur.ref_wait_since
            if value is not None and not own:
                target = C.pct_to_rpm(C.rpm_to_pct(value))
                if cur is not None:
                    cur.ctrl_target = target
                return target, ""
        if own:
            why = "이 모터 자신의 회전수라 따라갈 수 없음 (수동 지정으로 바꾸거나 구역설정의 참조기존모터 변경)"
        elif waited < C.NO_RESPONSE_WARN_SEC:
            return prev, ""                      # 기다림: 이전 속도 유지(없으면 쓰지 않음), 아직 이상표시 안 함
        else:
            why = (f"{C.NO_RESPONSE_WARN_SEC}초 넘게 값 없음 "
                   "(참조 기존모터 응답 없음·상태 이상 또는 그 구역 실행 안 함)")
        if prev is None:
            return None, f"구역 참조값({desc}) {why} - 속도를 지정하지 않음"
        return prev, f"구역 참조값({desc}) {why} - 이전 지정 {C.pct_text(C.rpm_to_pct(prev))} ({prev} RPM) 유지"

    def _ctrl_cycle(self, links: "_CtrlLinks", link: _Channel, m: Motor, stop: threading.Event) -> None:
        key = m.key
        if not m.use:
            if m.must_send_zero and not self.monitoring:
                # 사용 중지로 바꾸면 운전속도 0 을 한 번 보냄 (모니터링 모드에서는 보류했다가 제어 모드로 바뀌면 보냄)
                with self._lock:
                    cur = self._motors.get(key)
                    if cur is not None:
                        cur.must_send_zero = False
                self._mark_sent(key)
                self._ctrl_write(link, m, modbus.ECB_H_SPEED, 0, stop, links)
            return
        with self._lock:
            cur = self._motors.get(key)
            if cur is None:
                return
            if cur.check_since is None:
                cur.check_since = self._clock()
        # 0) 운전속도 결정 (구역 참조면 구역 참조값 확인 - 모니터링 모드에서도 확인해 표시)
        target, note = self._ctrl_target(m)
        # 1) 현재 설정 읽기 h1~h8 (h8 = 모터 최고속도), 2) 실제 회전수 i14 읽기
        hold = self._ctrl_request(link, modbus.build_read_holding(links.next_tid(), m.id, modbus.ECB_H_CONTROL, 8),
                                  stop)
        speed = self._ctrl_request(link, modbus.build_read_input(links.next_tid(), m.id, modbus.ECB_I_SPEED, 1), stop)
        actual = None
        if speed is not None and speed.exception is None and speed.values:
            actual = speed.values[0]
            with self._lock:
                cur = self._motors.get(key)
                if cur is not None:
                    self._mark_received(cur, actual)
        hold_ok = hold is not None and hold.exception is None and len(hold.values) >= 4
        max_rpm = hold.values[7] if hold_ok and len(hold.values) >= 8 else 0      # h8 최고속도 (0 = 설정 없음)
        if hold_ok:
            with self._lock:
                cur = self._motors.get(key)
                if cur is not None:
                    cur.cmd_rpm = hold.values[1]      # 모터가 가진 지정값(h2) - 표시용
                    cur.max_rpm = max_rpm
                    cur.target_rpm = target           # 표시용
        if hold_ok and not self.monitoring and target is not None:
            h1, h2, _h3, h4 = hold.values[:4]
            # 3) 제어모드: h2 절대 회전수 + 버스 Enable (영구저장 NV 레지스터 -> 다를 때만, 시도 횟수 제한)
            if not modbus.ecb_mode_ok(h4):
                with self._lock:
                    cur = self._motors.get(key)
                    tries = cur.mode_tries if cur else C.CTRL_MODE_WRITE_MAX_TRIES
                    if cur is not None and tries < C.CTRL_MODE_WRITE_MAX_TRIES:
                        cur.mode_tries += 1
                if tries >= C.CTRL_MODE_WRITE_MAX_TRIES:
                    note = "제어모드 설정 실패 (PIN 보호 또는 통신 확인 필요)"
                else:
                    log.info("[%s][기존모터제어 ID%s %s] 제어모드 변경 h4: %s -> %s", self.code, m.id, m.ip,
                             h4, modbus.ecb_mode_value(h4))
                    self._ctrl_write(link, m, modbus.ECB_H_PIN, modbus.ECB_PIN_LEVEL2, stop, links)
                    if not self._ctrl_write(link, m, modbus.ECB_H_CONTROLMODE, modbus.ecb_mode_value(h4), stop,
                                            links):
                        note = "제어모드 설정 실패"
            # 4) Enable
            if not h1 & modbus.ECB_ENABLE_BIT:
                if not self._ctrl_write(link, m, modbus.ECB_H_CONTROL, h1 | modbus.ECB_ENABLE_BIT, stop, links):
                    note = note or "운전 허가(h1) 쓰기 실패"
            # 5) 운전속도: 설정 속도를 그대로 (모터가 가진 값과 다를 때만).
            #    모터 최고속도(h8)보다 높은 설정 속도는 h8 로 제한됨(문서 3.4.6 주2) -> 알림만 (보내는 값은 그대로).
            #    속도를 보내는 제어 모드에서만 판단 (모니터링 모드에서는 이 프로그램이 보낸 속도가 아니므로 그대로 둠)
            with self._lock:
                cur = self._motors.get(key)
                if cur is not None:
                    if 0 < max_rpm < target:
                        self._set_speed_note(cur, f"모터 최고속도(h8) {max_rpm} RPM 이 설정 {target} RPM 보다 낮아 "
                                                  f"최고속도로 운전 (모터 스캐너의 ECblue 설정에서 최고속도 확인)")
                    else:
                        self._set_speed_note(cur, "")
            if h2 != target:
                self._mark_sent(key)
                if self._ctrl_write(link, m, modbus.ECB_H_SPEED, target, stop, links):
                    with self._lock:
                        cur = self._motors.get(key)
                        if cur is not None:
                            cur.cmd_rpm = target
                else:
                    note = note or f"운전속도(h2) {target} RPM 쓰기 실패"
        with self._lock:
            cur = self._motors.get(key)
            if cur is None:
                return
            if note:
                if cur.ctrl_note != note:
                    cur.ctrl_note = note
                    self._set_err(cur, note)
            elif cur.ctrl_note:
                cur.ctrl_note = ""
                if cur.rdate is not None and cur.check_since is None:
                    self._clear_err(cur)       # 응답도 정상이면 이상표시 해제
                self._touch()

    # --- 공통 수신 ---
    def _recv_loop(self, ch: _Channel, stop: threading.Event, handler, mtype: MotorType) -> None:
        splitter = modbus.FrameSplitter()
        sock = ch.sock
        if sock is None:
            return
        while not stop.is_set():
            try:
                data = sock.recv(513)
            except socket.timeout:
                continue
            except OSError:
                break
            if not data:
                # 상대가 연결을 끊음 -> 송신 쓰레드가 재접속하도록 소켓 정리
                if ch.sock is sock and not stop.is_set():
                    self._channel_lost(ch, mtype, "통신장치가 연결을 끊음")
                break
            for frame in splitter.feed(data):
                try:
                    handler(frame)
                except Exception:  # 수신 처리 오류로 쓰레드가 죽지 않도록
                    log.exception("[%s][%s] 수신 처리 오류", self.code, ch.name)
