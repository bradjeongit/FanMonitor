"""구역 단위 모터 관리 및 Modbus TCP 송수신. 기준: 운영 실행파일 C_HOUSE

- 구역 1개 = 기존모터 게이트웨이 1개(SOURCE) + 신규모터 게이트웨이 1개(TARGET)
- 기존모터: RPM 을 읽어 저장
- 신규모터: 1페이지(현재 RPM)와 2페이지(수동/동작/RPM설정)를 읽고, 2페이지 값이 목표와 다르면 써넣음
  목표 RPM = 수동RPM 또는 참조 기존모터 RPM, 속도제한 사용 시 최대속도 이하로 제한
- 상태: 응답 수신 = 정상, 명령 후 60초 무응답 = '이상기록', 마지막 응답 후 1시간 = 이상
- 기존모터제어(OLD_CTRL): 모터마다 지정한 IP/PORT/ID 의 ECblue 기존 모터에 직접 접속해
  h4(제어모드=버스 절대회전수) → h1(Enable) → h2(지정 운전속도) 순으로 맞추고, i14(실제 회전수)를 읽음
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
from typing import TYPE_CHECKING, Callable, Deque, Dict, List, Optional

from . import constants as C
from . import modbus
from .config import MotorEntry
from .constants import MotorStatus, MotorType

if TYPE_CHECKING:
    from .plant import Plant

log = logging.getLogger("motor_monitor.house")

# 디버그 출력 대상 (운영 DEBUG_LINE / DEBUG_IDS). 로그 레벨이 DEBUG 일 때만 출력됩니다.
DEBUG_LINE = "ALL"          # 구역코드 또는 "ALL"
DEBUG_IDS = {5, 50}         # 출력할 모터 ID (0 은 구역 공통 메시지로 항상 출력)


def motor_key(mtype: MotorType, mid: int, ip: str = "", port: int = 0) -> str:
    """모터 고유 키 '종류_ID' (예: '0_5'). 기존모터제어는 접속처가 달라 '2_ID@IP:PORT'."""
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
    had_err: bool = False             # 60초 무응답 이력 (구역 제어를 다시 시작할 때까지 유지)
    status: MotorStatus = MotorStatus.READY
    must_send_zero: bool = False      # 사용중지 선택 시 RPM 0 을 1회 보내기 위한 표시
    ip: str = ""                      # 기존모터제어: 접속 IP
    port: int = 0                     # 기존모터제어: 접속 PORT
    set_rpm: int = 0                  # 기존모터제어: 지정 운전 속도
    mode_tries: int = 0               # 기존모터제어: 제어모드(h4, NV) 설정 시도 횟수
    ctrl_note: str = ""               # 기존모터제어: 마지막 제어 상태 설명 (툴팁/로그용)

    @property
    def key(self) -> str:
        return motor_key(self.mtype, self.id, self.ip, self.port)

    def entry(self) -> MotorEntry:
        return MotorEntry(self.id, self.mtype, self.use, self.use_limit, self.limit_rpm,
                          self.ip, self.port, self.set_rpm)


@dataclass
class _QueueItem:
    dkey: str
    id: int
    msg: bytes


class _Channel:
    """TCP 소켓 1개(기존모터용 또는 신규모터용)."""

    def __init__(self, name: str) -> None:
        self.name = name
        self.sock: Optional[socket.socket] = None
        self.connected = False
        self.net_status = False      # 통신상태 True 정상 / False 이상
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
                  ip: str = "", port: int = 0, set_rpm: int = 0) -> bool:
        key = motor_key(mtype, mid, ip, port)
        with self._lock:
            if key in self._motors:
                return False
            status = MotorStatus.READY if use else MotorStatus.NOTUSE
            self._motors[key] = Motor(id=mid, mtype=MotorType(mtype), use=use, use_limit=use_limit,
                                      limit_rpm=limit_rpm, status=status, ip=ip, port=port, set_rpm=set_rpm)
            self._order.append(key)
            self._resort()
            self._touch()
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

    def mod_motor(self, dkey: str, use: bool, use_limit: bool = False, limit_rpm: int = 0,
                  set_rpm: Optional[int] = None) -> None:
        """사용여부/속도제한(/기존모터제어 운전속도) 수정. 운영 Mod_Motor"""
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
            self._touch()

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
        """운영 Update_BIND_LIST(): 정렬 기준에 따라 표시 순서 재구성 (안정 정렬)."""
        if self._sort_key == "":
            self._sort_key = "MTYPE"
            self._sort_desc = True
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
            keys = sorted(keys, key=lambda k: ms[k].id, reverse=desc)
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
                    if m.rdate is None:
                        m.status = MotorStatus.READY
                    else:
                        if m.check_since is not None and int(now - m.check_since) >= C.NO_RESPONSE_WARN_SEC:
                            m.had_err = True
                            m.check_since = None
                        m.status = (MotorStatus.ERR if int(now - m.rdate) >= C.MOTOR_TIMEOUT_SEC
                                    else MotorStatus.OK)
                else:
                    m.status = MotorStatus.NOTUSE
                if (m.status, m.had_err) != old:
                    changed = True
            if changed:
                self._touch()
        return changed

    def _mark_received(self, m: Motor, rpm: Optional[int] = None) -> None:
        if rpm is not None:
            m.rpm = rpm
        m.check_since = None
        m.rdate = self._clock()
        m.status = MotorStatus.OK
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
                m.check_since = None
                m.rdate = None
                m.status = MotorStatus.READY if m.use else MotorStatus.NOTUSE
            self._touch()
        self.source.net_status = False
        self.target.net_status = False
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

    def _connect(self, ch: _Channel, ip: str, port: int, stop: threading.Event) -> bool:
        if ch.connected:
            return True
        self._debug(0, f"[{self.code}][{ch.name}] 접속시도")
        if ch.connect(ip, port, C.CONNECT_CHECK_TIMEOUT):
            log.info("[%s][%s] 접속 성공 %s:%s", self.code, ch.name, ip, port)
            return True
        log.info("[%s][%s] 접속 실패 %s:%s (%d초 후 재시도)", self.code, ch.name, ip, port, C.RECONNECT_DELAY)
        ch.net_status = False
        stop.wait(C.RECONNECT_DELAY)
        return False

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
        self._debug(item.id, f"[{self.code}][{tag}][{item.id}] {modbus.to_hex(item.msg)}")
        stop.wait(C.SEND_INTERVAL)

    # --- 기존모터 (SOURCE) ---
    def _src_send_job(self, stop: threading.Event) -> None:
        self._debug(0, f"[{self.code}][S] 전송 시작")
        ch = self.source
        while not stop.is_set():
            cfg = self.plant.configs[self.code]
            if not self._connect(ch, cfg.source_ip, cfg.source_port, stop):
                continue
            ch.net_status = True
            self._ensure_recv("src_recv", self._src_recv_job, stop)
            try:
                if self._src_read_q:
                    for _ in range(len(self._src_read_q)):
                        if stop.is_set():
                            return
                        self._send_item(ch, self._src_read_q.popleft(), "S", stop)
                else:
                    with self._lock:
                        targets = [(m.key, m.id) for m in self._motors.values()
                                   if m.mtype == MotorType.OLD_MOTOR and m.use]
                    for dkey, mid in targets:
                        msg = modbus.build_old_rpm_read(self._next_src_idx(), mid)
                        self._src_read_q.append(_QueueItem(dkey, mid, msg))
                    stop.wait(C.CYCLE_INTERVAL)
            except (OSError, IndexError):
                ch.close()
                stop.wait(C.ERROR_RETRY_DELAY)
        ch.close()

    def _src_recv_job(self, stop: threading.Event) -> None:
        self._recv_loop(self.source, stop, self._handle_src_frame)

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
        while not stop.is_set():
            cfg = self.plant.configs[self.code]
            if not self._connect(ch, cfg.target_ip, cfg.target_port, stop):
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
                else:
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
                                self._enqueue_write(_QueueItem(m.key, m.id, msg))
                    if not self._tgt_read_q and self._tgt_write_q:
                        # 사용중인 신규모터가 없어 읽기 명령이 없을 때도 쓰기 명령은 전송
                        # (운영 프로그램은 이 경우 RPM 0 명령이 쌓이기만 하고 전송되지 않음)
                        self._flush_writes(ch, stop)
                    stop.wait(C.CYCLE_INTERVAL)
            except (OSError, IndexError):
                ch.close()
                stop.wait(C.ERROR_RETRY_DELAY)
        ch.close()

    def _tgt_recv_job(self, stop: threading.Event) -> None:
        self._recv_loop(self.target, stop, self._handle_tgt_frame)

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
            use_limit, limit_rpm = m.use_limit, m.limit_rpm
        setting = None if self.monitoring else self.resolve_setting_rpm()
        if setting is not None:
            if use_limit and setting > limit_rpm:
                setting = limit_rpm          # 속도제한: 최대속도 이하로
            mid = r.unit_id
            if r.automanual != 1:
                self._enqueue_write(_QueueItem(dkey, mid, modbus.build_new_manual_write(self._next_tgt_idx(), mid)))
            if r.startstop != 1:
                self._enqueue_write(_QueueItem(dkey, mid, modbus.build_new_run_write(self._next_tgt_idx(), mid)))
            self._debug(mid, f"[{self.code}][{mid}]-제어RPM:{setting}")
            if r.rpm != setting:
                self._enqueue_write(
                    _QueueItem(dkey, mid, modbus.build_new_rpm_write(self._next_tgt_idx(), mid, setting)))
        with self._lock:
            m = self._motors.get(dkey)
            if m is not None:
                self._mark_received(m)

    # --- 기존모터제어 (OLD_CTRL, ZIEHL-ABEGG ECblue) ---
    def _ctrl_job(self, stop: threading.Event) -> None:
        """기존모터제어 모터들을 접속처(IP:PORT)별 연결 하나로 차례대로 읽고 맞춤."""
        links: Dict[tuple, _Channel] = {}
        retry_at: Dict[tuple, float] = {}
        tids = {"n": 0}
        try:
            while not stop.is_set():
                with self._lock:
                    motors = [replace(m) for m in self._motors.values() if m.mtype == MotorType.OLD_CTRL]
                if not motors:
                    stop.wait(C.CYCLE_INTERVAL)
                    continue
                groups: Dict[tuple, List[Motor]] = {}
                for m in motors:
                    groups.setdefault((m.ip, m.port), []).append(m)
                for ep, group in groups.items():
                    if stop.is_set():
                        break
                    if not any(m.use or m.must_send_zero for m in group):
                        continue
                    link = links.setdefault(ep, _Channel(f"C{ep[0]}:{ep[1]}"))
                    if not link.connected:
                        if time.monotonic() < retry_at.get(ep, 0):
                            continue
                        if not link.connect(ep[0], ep[1], C.CONNECT_CHECK_TIMEOUT):
                            log.info("[%s][기존모터제어] 접속 실패 %s:%s (%d초 후 재시도)",
                                     self.code, ep[0], ep[1], C.RECONNECT_DELAY)
                            retry_at[ep] = time.monotonic() + C.RECONNECT_DELAY
                            continue
                        log.info("[%s][기존모터제어] 접속 성공 %s:%s", self.code, ep[0], ep[1])
                    for m in group:
                        if stop.is_set():
                            break
                        try:
                            self._ctrl_cycle(link, m, stop, tids)
                        except OSError:
                            link.close()
                            retry_at[ep] = time.monotonic() + C.ERROR_RETRY_DELAY
                            break
                stop.wait(C.CYCLE_INTERVAL)
        finally:
            for link in links.values():
                link.close()

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
                    tids: dict) -> bool:
        """쓰기 1건. 모니터링 모드에서는 절대 보내지 않음. 정상 응답(echo)이면 True."""
        if self.monitoring:
            self._debug(m.id, f"[{self.code}][모니터링] 기존모터제어 쓰기 생략 h{reg}={value}")
            return False
        tids["n"] = (tids["n"] + 1) % (C.CMD_IDX_MAX + 1)
        r = self._ctrl_request(link, modbus.build_new_write_single(tids["n"], m.id, reg, value), stop)
        if r is None or r.exception is not None:
            log.warning("[%s][기존모터제어 ID%s %s] h%s=%s 쓰기 실패 (%s)", self.code, m.id, m.ip, reg, value,
                        "응답 없음" if r is None else f"예외코드 {r.exception}")
            return False
        return True

    def _ctrl_cycle(self, link: _Channel, m: Motor, stop: threading.Event, tids: dict) -> None:
        key = m.key
        if not m.use:
            if m.must_send_zero and not self.monitoring:
                # 사용 중지로 바꾸면 운전속도 0 을 한 번 보냄 (모니터링 모드에서는 보류했다가 제어 모드로 바뀌면 보냄)
                with self._lock:
                    cur = self._motors.get(key)
                    if cur is not None:
                        cur.must_send_zero = False
                self._ctrl_write(link, m, modbus.ECB_H_SPEED, 0, stop, tids)
            return
        with self._lock:
            cur = self._motors.get(key)
            if cur is None:
                return
            if cur.check_since is None:
                cur.check_since = self._clock()
        # 1) 현재 설정 읽기 h1~h4, 2) 실제 회전수 i14 읽기
        tids["n"] = (tids["n"] + 1) % (C.CMD_IDX_MAX + 1)
        hold = self._ctrl_request(link, modbus.build_read_holding(tids["n"], m.id, modbus.ECB_H_CONTROL, 4), stop)
        tids["n"] = (tids["n"] + 1) % (C.CMD_IDX_MAX + 1)
        speed = self._ctrl_request(link, modbus.build_read_input(tids["n"], m.id, modbus.ECB_I_SPEED, 1), stop)
        if speed is not None and speed.exception is None and speed.values:
            with self._lock:
                cur = self._motors.get(key)
                if cur is not None:
                    self._mark_received(cur, speed.values[0])
        if hold is None or hold.exception is not None or len(hold.values) < 4:
            return
        h1, h2, _h3, h4 = hold.values[:4]
        if self.monitoring:
            return
        target = max(0, min(int(m.set_rpm), C.MAX_RPM))
        note = ""
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
                self._ctrl_write(link, m, modbus.ECB_H_PIN, modbus.ECB_PIN_LEVEL2, stop, tids)
                if not self._ctrl_write(link, m, modbus.ECB_H_CONTROLMODE, modbus.ecb_mode_value(h4), stop, tids):
                    note = "제어모드 설정 실패"
        # 4) Enable, 5) 운전속도
        if not h1 & modbus.ECB_ENABLE_BIT:
            self._ctrl_write(link, m, modbus.ECB_H_CONTROL, h1 | modbus.ECB_ENABLE_BIT, stop, tids)
        if h2 != target:
            self._ctrl_write(link, m, modbus.ECB_H_SPEED, target, stop, tids)
        with self._lock:
            cur = self._motors.get(key)
            if cur is not None and cur.ctrl_note != note:
                cur.ctrl_note = note
                if note:
                    cur.had_err = True
                    log.warning("[%s][기존모터제어 ID%s %s] %s", self.code, m.id, m.ip, note)
                self._touch()

    # --- 공통 수신 ---
    def _recv_loop(self, ch: _Channel, stop: threading.Event, handler) -> None:
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
                if ch.sock is sock:
                    ch.close()
                break
            for frame in splitter.feed(data):
                try:
                    handler(frame)
                except Exception:  # 수신 처리 오류로 쓰레드가 죽지 않도록
                    log.exception("[%s][%s] 수신 처리 오류", self.code, ch.name)
