"""전체 구역/설정을 묶는 모델. 기준: 운영 M_Variables(HOUSES, MCONFIG) + M_Config

화면(UI)과 통신 로직 사이의 단일 진입점입니다.
"""
from __future__ import annotations

import datetime
import logging
import os
import shutil
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from . import config as cfgmod
from .config import LineConfig, MiniItem, MotorEntry
from .constants import CTRL_MODE_REF, HOUSE_LINES, MotorStatus, MotorType
from .errorlog import ErrorLog
from .helpers import get_house_name
from .house import HouseLine, Motor, motor_key

log = logging.getLogger(__name__)


@dataclass
class ConfigImport:
    """가져올 운영 프로그램 설정 (미리보기 후 apply_config_import 로 적용)."""
    path: str
    configs: Dict[str, LineConfig]
    motors: Dict[str, List[MotorEntry]]

    def summary_lines(self) -> List[str]:
        out = []
        for s in HOUSE_LINES:
            c = self.configs[s]
            ms = self.motors.get(s, [])
            cnt = {t: sum(1 for m in ms if m.mtype == t) for t in MotorType}
            old = c.source_ip + (f", {c.source_ip2}" if c.source_ip2 else "")
            out.append(f"{get_house_name(s)}: 기존 {old} / 신규 {c.target_ip}"
                       f" / 기존모터 {cnt[MotorType.OLD_MOTOR]}, 신규모터 {cnt[MotorType.NEW_MOTOR]}, "
                       f"기존모터제어 {cnt[MotorType.OLD_CTRL]}")
        return out

    @property
    def has_ctrl(self) -> bool:
        return any(m.mtype == MotorType.OLD_CTRL for ms in self.motors.values() for m in ms)


class Plant:
    def __init__(self, config_path: Optional[str] = None, monitoring: bool = False,
                 error_log_dir: Optional[str] = None) -> None:
        self.config_path = config_path or cfgmod.default_config_path()
        self.errors = ErrorLog(error_log_dir)
        # 모니터링 모드: 모터 상태를 읽기만 하고 제어(쓰기) 신호는 보내지 않음.
        # 시작 모드는 --monitoring 옵션으로 정하고, 실행 중에는 화면의 모드 전환 버튼으로 바꿈(set_monitoring).
        # 모니터링 모드에서는 시작/종료 시 Config.dll 자동 저장을 하지 않음 (사용자가 직접 바꾼 설정만 저장).
        self._monitoring = monitoring
        self.houses: Dict[str, HouseLine] = {}
        self.configs: Dict[str, LineConfig] = {}

    @property
    def monitoring(self) -> bool:
        return self._monitoring

    def set_monitoring(self, monitoring: bool) -> None:
        """실행 중 모드 전환. 통신 스레드는 명령마다 이 값을 다시 읽으므로 즉시 반영됨
        (모니터링 -> 쓰기 즉시 중단, 대기 중인 쓰기 명령도 버림 / 제어 -> 다음 주기부터 쓰기 시작)."""
        monitoring = bool(monitoring)
        if monitoring == self._monitoring:
            return
        self._monitoring = monitoring
        log.warning("실행 모드 전환: %s", "모니터링 모드 (제어 신호 전송 안 함)" if monitoring else "제어 모드")
        if not monitoring:
            self.kick_all()                   # 30초를 기다리지 않고 바로 점검해 제어 신호를 보냄

    def kick_all(self) -> None:
        """모든 구역을 다음 점검(30초)까지 기다리지 않고 바로 한 번 점검."""
        for h in self.houses.values():
            h.kick()

    # ------------------------------------------------------------- 설정
    def load(self) -> None:
        """운영 Get_Local_Config: 구역 생성 -> 설정 읽기 -> 모터 등록 -> 설정 다시 저장.
        설정 파일을 해독할 수 없으면 ConfigError (기존 파일을 덮어쓰지 않음)."""
        configs, motors = cfgmod.load_config(self.config_path)
        self.configs = configs
        self.houses = {}
        for s in HOUSE_LINES:
            h = HouseLine(s, self)
            for e in motors.get(s, []):
                h.add_entry(e)
            self.houses[s] = h
        if not self.monitoring:
            self.save()

    def save(self) -> None:
        """운영 Put_Local_Config"""
        motors = {s: self.houses[s].export_motors() for s in HOUSE_LINES}
        try:
            cfgmod.save_config(self.config_path, self.configs, motors)
        except OSError as e:
            log.error("설정 저장 실패: %s", e)
            self.errors.error("-", "설정 파일", f"저장 실패: {e}")
            raise

    def set_line_config(self, code: str, cfg: LineConfig) -> None:
        self.configs[code] = cfg
        self.save()
        self.kick_all()                       # IP 변경을 바로 반영

    # ------------------------------------------------------------- 조회
    def find_reference_users(self, line_code: str, motor_id: int, gw: int = 1) -> List[str]:
        """해당 기존모터를 참조 속도로 쓰는 모터 목록 (예: '3구역 신규모터 ID 5') - 삭제 전 확인용."""
        out = []
        for s in HOUSE_LINES:
            for m in self.houses[s].get_list():
                if (m.mtype != MotorType.OLD_MOTOR and m.ctrl_mode == CTRL_MODE_REF and m.ref_line == line_code
                        and m.ref_gw == gw and m.ref_id == motor_id):
                    out.append(f"{get_house_name(s)} {m.label}")
        return out

    def reference_ok(self, code: str, m: Motor) -> bool:
        """모터의 속도 출처 상태: 수동이면 정상, 참조이면 참조 모터가 '정상'일 때만 정상."""
        if m.ctrl_mode != CTRL_MODE_REF:
            return True
        sm = self.houses[code].reference_motor(m)
        return sm is not None and sm.status == MotorStatus.OK

    def motor_endpoint(self, line: str, gw: int, mid: int) -> Optional[Tuple[str, int, int]]:
        """기존모터(구역, 통신장치 번호, ID)의 (통신장치 IP, PORT, ID). 없으면 None.
        기존모터제어의 참조가 자기 자신을 따라가지 않도록 확인할 때 씀."""
        cfg = self.configs.get(line)
        if cfg is None or not cfg.gw_ip(gw):
            return None
        return cfg.gw_ip(gw), cfg.source_port, mid

    def find_old_gateway(self, ip: str) -> Optional[Tuple[str, int]]:
        """기존모터 통신장치 IP -> (구역, 통신장치 번호). 없으면 None."""
        for s in HOUSE_LINES:
            c = self.configs[s]
            if c.source_ip == ip:
                return s, 1
            if c.source_ip2 and c.source_ip2 == ip:
                return s, 2
        return None

    # ------------------------------------------------------------- 제어
    def start_line(self, code: str) -> None:
        self.houses[code].service_start()

    def stop_line(self, code: str) -> None:
        self.houses[code].service_stop()

    def update_all_status(self) -> None:
        """제어 중인 구역만 상태 갱신 (운영: 상태 타이머는 구역 제어 시작 시 켜지고 종료 시 꺼짐)."""
        for h in self.houses.values():
            if h.running:
                h.update_status()

    def shutdown(self) -> None:
        """프로그램 종료: 설정 저장(모니터링 모드 제외) 후 모든 구역 제어 종료."""
        try:
            if not self.monitoring:
                self.save()
        finally:
            for h in self.houses.values():
                h.service_stop()

    # ------------------------------------------------------------- 다른 프로그램 설정 가져오기
    @property
    def any_running(self) -> bool:
        return any(h.running for h in self.houses.values())

    def read_config_import(self, path: str) -> ConfigImport:
        """운영 프로그램 설정 파일(Config.dll / uconfig.dll / Config.ini) 읽기. 적용은 하지 않음."""
        configs, motors = cfgmod.config_from_text(cfgmod.read_any_config_text(path))
        return ConfigImport(path, configs, motors)

    def backup_config(self) -> Optional[str]:
        """현재 설정 파일을 Config.dll.bak_YYYYMMDD_HHMMSS 로 복사 (없으면 None)."""
        if not os.path.exists(self.config_path):
            return None
        dst = f"{self.config_path}.bak_{datetime.datetime.now():%Y%m%d_%H%M%S}"
        shutil.copy2(self.config_path, dst)
        return dst

    def apply_config_import(self, imp: ConfigImport, keep_ctrl: bool = True) -> Optional[str]:
        """가져온 설정으로 구역 설정과 모터 목록을 바꾸고 저장. 반환: 백업 파일 경로.
        keep_ctrl: 가져온 파일에 기존모터제어가 없으면 지금의 기존모터제어 목록을 유지."""
        if self.any_running:
            raise RuntimeError("실행 중인 구역이 있어 설정을 바꿀 수 없습니다. 모든 구역을 종료한 뒤 가져오세요.")
        backup = self.backup_config()
        keep = keep_ctrl and not imp.has_ctrl
        for s in HOUSE_LINES:
            h = self.houses[s]
            ctrl = [m.entry() for m in h.get_list() if m.mtype == MotorType.OLD_CTRL] if keep else []
            h.clear_motors()
            for e in imp.motors.get(s, []) + ctrl:
                h.add_entry(e)
            self.configs[s] = imp.configs[s]
        self.save()
        log.warning("설정 가져오기: %s (백업 %s)", imp.path, backup)
        return backup

    def import_mini_items(self, items: List[MiniItem], code: str) -> Tuple[List[str], List[str]]:
        """미니 제어기 항목을 구역 code 의 기존모터제어로 추가하고 저장. 반환: (추가된 항목, 건너뛴 항목)."""
        h = self.houses[code]
        added, skipped = [], []
        for it in items:
            e = cfgmod.mini_item_to_entry(it)
            if e.ctrl_mode == CTRL_MODE_REF:
                # 미니 제어기의 참조팬(IP·ID)이 어느 구역의 기존모터인지 찾아 참조로 지정 (못 찾으면 예전 구역 참조값)
                found = self.find_old_gateway(it.ref_ip)
                if found is not None:
                    e = e._replace(ref_line=found[0], ref_gw=found[1], ref_id=it.ref_uid)
                else:
                    e = cfgmod.migrate_entry(e, self.configs[code])
            name = f"{it.ip} ID {it.uid}"
            if h.get_motor(motor_key(e.mtype, e.id, e.ip, e.port)) is not None:
                skipped.append(f"{name} (이미 있음)")
                continue
            h.add_entry(e)
            added.append(name)
        if added:
            self.save()
        return added, skipped

    def zone_for_ip(self, ip: str) -> Optional[str]:
        """기존모터 통신장치 IP(1 또는 2)가 같은 구역 (없으면 None)."""
        found = self.find_old_gateway(ip)
        return found[0] if found else None
