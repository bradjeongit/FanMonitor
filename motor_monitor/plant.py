"""전체 라인/설정을 묶는 모델. 원본: M_Variables.vb(HOUSES, MCONFIG) + M_Config.vb

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
    def __init__(self, config_path: Optional[str] = None) -> None:
        self.config_path = config_path or cfgmod.default_config_path()
        self.houses: Dict[str, HouseLine] = {}
        self.configs: Dict[str, LineConfig] = {}
        # 온실별 제어 실행 여부 (원본 H1_NETWORK_RUN / H2_NETWORK_RUN)
        self.network_run: Dict[str, bool] = {"H1": False, "H2": False}

    # ------------------------------------------------------------- 설정
    def load(self) -> None:
        """원본 Get_Local_Config: 라인 생성 -> 설정 읽기 -> 모터 등록 -> 설정 다시 쓰기."""
        configs, motors = cfgmod.load_config(self.config_path)
        self.configs = configs
        self.houses = {}
        for s in HOUSE_LINES:
            h = HouseLine(s, self)
            for mid, mtype, use in motors.get(s, []):
                h.add_motor(mid, mtype, use)
            self.houses[s] = h
        self.save()

    def save(self) -> None:
        """원본 Put_Local_Config"""
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
        """해당 기존모터를 참조 RPM 소스로 쓰는 라인 목록 (삭제 전 확인용)."""
        return [
            s for s in HOUSE_LINES
            if not self.configs[s].source_manual
            and self.configs[s].source_line == line_code
            and self.configs[s].source_id == motor_id
        ]

    def reference_ok(self, code: str) -> bool:
        """라인 헤더의 참조 상태 표시용: 수동이면 항상 정상, 참조이면 참조 모터가 정상인지."""
        cfg = self.configs[code]
        if cfg.source_manual:
            return True
        src = self.houses.get(cfg.source_line)
        if src is None:
            return False
        m = src.get_motor(motor_key(MotorType.OLD_MOTOR, cfg.source_id))
        return m is not None and m.status == MotorStatus.OK

    # ------------------------------------------------------------- 제어
    def lines_of(self, house_prefix: str) -> List[str]:
        return [s for s in HOUSE_LINES if s[:2] == house_prefix]

    def start_house(self, house_prefix: str) -> None:
        for s in self.lines_of(house_prefix):
            self.houses[s].service_start()
        self.network_run[house_prefix] = True

    def stop_house(self, house_prefix: str) -> None:
        for s in self.lines_of(house_prefix):
            self.houses[s].service_stop()
        self.network_run[house_prefix] = False

    def update_all_status(self) -> None:
        for h in self.houses.values():
            h.update_status()

    def shutdown(self) -> None:
        """프로그램 종료: 설정 저장 후 모든 서비스 종료 (원본 MainWindow_Closing)."""
        try:
            self.save()
        finally:
            for h in self.houses.values():
                h.service_stop()
