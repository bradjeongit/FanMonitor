"""전체 구역/설정을 묶는 모델. 기준: 운영 M_Variables(HOUSES, MCONFIG) + M_Config

화면(UI)과 통신 로직 사이의 단일 진입점입니다.
"""
from __future__ import annotations

import logging
from typing import Dict, List, Optional

from . import config as cfgmod
from .config import LineConfig
from .constants import HOUSE_LINES, MotorStatus, MotorType
from .house import HouseLine, motor_key

log = logging.getLogger(__name__)


class Plant:
    def __init__(self, config_path: Optional[str] = None, monitoring: bool = False) -> None:
        self.config_path = config_path or cfgmod.default_config_path()
        # 모니터링 모드: 모터 상태를 읽기만 하고 제어(쓰기) 신호는 보내지 않음. 실행 중에는 바뀌지 않음.
        # 또한 시작/종료 시 Config.dll 자동 저장을 하지 않음 (사용자가 직접 바꾼 설정만 저장).
        self._monitoring = monitoring
        self.houses: Dict[str, HouseLine] = {}
        self.configs: Dict[str, LineConfig] = {}

    @property
    def monitoring(self) -> bool:
        return self._monitoring

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
                h.add_motor(e.id, e.mtype, e.use, e.use_limit, e.limit_rpm)
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
            raise

    def set_line_config(self, code: str, cfg: LineConfig) -> None:
        self.configs[code] = cfg
        self.save()

    # ------------------------------------------------------------- 조회
    def find_reference_users(self, line_code: str, motor_id: int) -> List[str]:
        """해당 기존모터를 참조 RPM 소스로 쓰는 구역 목록 (삭제 전 확인용)."""
        return [
            s for s in HOUSE_LINES
            if not self.configs[s].source_manual
            and self.configs[s].source_line == line_code
            and self.configs[s].source_id == motor_id
        ]

    def reference_ok(self, code: str) -> bool:
        """RPM 소스 상태 점: 수동이면 정상, 참조이면 참조 모터가 '정상'일 때만 정상."""
        cfg = self.configs[code]
        if cfg.source_manual:
            return True
        src = self.houses.get(cfg.source_line)
        if src is None:
            return False
        m = src.get_motor(motor_key(MotorType.OLD_MOTOR, cfg.source_id))
        return m is not None and m.status == MotorStatus.OK

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
