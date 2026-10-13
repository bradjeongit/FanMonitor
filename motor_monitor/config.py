"""설정 파일(Config.ini) 읽기/쓰기. 원본: M_Config.vb, M_Variables.vb(Struct_Config)

원본과 같은 섹션/키 이름을 그대로 사용하므로 기존 프로그램의 Config.ini 를 그대로 읽을 수 있습니다.
(원본 키 이름의 오타 '_TARGEIP', '_TARGEPORT' 도 호환을 위해 그대로 유지합니다.)
"""
from __future__ import annotations

import configparser
import io
import logging
import os
import sys
import tempfile
from dataclasses import dataclass, replace
from typing import Dict, List, Tuple

from .constants import (
    CONFIG_FILE_NAME,
    DEFAULT_IP_PREFIX,
    DEFAULT_PORT,
    HOUSE_LINES,
    MotorType,
)

log = logging.getLogger(__name__)

SEC_LINE = "LINE_CONFIG"
SEC_MOTOR = "MOTOR_CONFIG"


@dataclass
class LineConfig:
    """라인(구역/짝홀) 1개의 통신/제어 설정."""
    source_ip: str          # 기존모터 IP
    source_port: int        # 기존모터 PORT
    target_ip: str          # 신규모터 IP
    target_port: int        # 신규모터 PORT
    source_manual: bool     # True: 수동 RPM 사용 / False: 기존모터 RPM 참조
    manual_rpm: int         # 수동 RPM 값
    source_line: str        # 참조 기존모터가 있는 라인 코드 (예: H1R1E)
    source_id: int          # 참조 기존모터 ID

    def copy(self, **changes) -> "LineConfig":
        return replace(self, **changes)


# (id, type, use)
MotorEntry = Tuple[int, MotorType, bool]


def app_base_dir() -> str:
    """실행 파일(또는 run.py)이 있는 폴더. 원본의 AppDomain.CurrentDomain.BaseDirectory 에 해당."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def default_config_path() -> str:
    return os.path.join(app_base_dir(), CONFIG_FILE_NAME)


def _new_parser() -> configparser.ConfigParser:
    p = configparser.ConfigParser(interpolation=None, strict=False, delimiters=("=",))
    p.optionxform = str  # 키 대소문자 유지
    return p


def _read_text(path: str) -> str:
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp949", errors="replace")


def _ini_read(p: configparser.ConfigParser, section: str, key: str, default: str) -> str:
    """원본 INIRead 와 동일: 값이 없으면 기본값, '&' '|' 문자는 제거."""
    if p.has_option(section, key):
        val = p.get(section, key)
    else:
        val = default
    return val.replace("&", "").replace("|", "")


def _to_int(text: str, default: int = 0) -> int:
    """VB Integer.TryParse 와 유사: 실패하면 기본값(0)."""
    try:
        return int(text.strip())
    except (ValueError, AttributeError):
        return default


def parse_motor_list(data: str) -> List[MotorEntry]:
    """'5=0=1,6=1=0' -> [(5, OLD, True), (6, NEW, False)]"""
    result: List[MotorEntry] = []
    if not data:
        return result
    for item in data.split(","):
        parts = item.split("=")
        if len(parts) < 3:
            log.warning("모터 설정 형식 오류(무시): %r", item)
            continue
        mid = _to_int(parts[0])
        mtype_raw = _to_int(parts[1])
        use = _to_int(parts[2]) != 0
        try:
            mtype = MotorType(mtype_raw)
        except ValueError:
            log.warning("알 수 없는 모터 종류(무시): %r", item)
            continue
        result.append((mid, mtype, use))
    return result


def format_motor_list(motors: List[MotorEntry]) -> str:
    return ",".join(f"{mid}={int(mtype)}={1 if use else 0}" for mid, mtype, use in motors)


def load_config(path: str) -> Tuple[Dict[str, LineConfig], Dict[str, List[MotorEntry]]]:
    """설정 읽기. 파일/키가 없으면 원본과 같은 기본값을 사용."""
    p = _new_parser()
    if os.path.exists(path):
        try:
            p.read_string(_read_text(path))
        except (configparser.Error, OSError) as e:
            log.error("설정 파일을 읽지 못했습니다(기본값 사용): %s", e)
            p = _new_parser()

    configs: Dict[str, LineConfig] = {}
    motors: Dict[str, List[MotorEntry]] = {}
    idx = 1
    for s in HOUSE_LINES:
        sip = _ini_read(p, SEC_LINE, f"{s}_SOURCE_IP", f"{DEFAULT_IP_PREFIX}{idx}")
        sport = _to_int(_ini_read(p, SEC_LINE, f"{s}_SOURCE_PORT", str(DEFAULT_PORT)))
        tip = _ini_read(p, SEC_LINE, f"{s}_TARGEIP", f"{DEFAULT_IP_PREFIX}{idx + 1}")
        tport = _to_int(_ini_read(p, SEC_LINE, f"{s}_TARGEPORT", str(DEFAULT_PORT)))
        manual = _ini_read(p, SEC_LINE, f"{s}_SOURCE_MANUAL", "1") != "0"
        mrpm = _to_int(_ini_read(p, SEC_LINE, f"{s}_MANUAL_RPM", "0"))
        sline = _ini_read(p, SEC_LINE, f"{s}_SOURCE_LINE", "")
        sid = _to_int(_ini_read(p, SEC_LINE, f"{s}_SOURCE_ID", "0"))
        configs[s] = LineConfig(sip, sport, tip, tport, manual, mrpm, sline, sid)
        motors[s] = parse_motor_list(_ini_read(p, SEC_MOTOR, f"{s}_MOTORS", ""))
        idx += 2
    return configs, motors


def save_config(path: str, configs: Dict[str, LineConfig], motors: Dict[str, List[MotorEntry]]) -> None:
    """설정 쓰기. 기존 파일의 다른 섹션/키는 보존합니다."""
    p = _new_parser()
    if os.path.exists(path):
        try:
            p.read_string(_read_text(path))
        except (configparser.Error, OSError):
            p = _new_parser()
    for sec in (SEC_LINE, SEC_MOTOR):
        if not p.has_section(sec):
            p.add_section(sec)

    for s in HOUSE_LINES:
        c = configs[s]
        p.set(SEC_LINE, f"{s}_SOURCE_IP", c.source_ip)
        p.set(SEC_LINE, f"{s}_SOURCE_PORT", str(c.source_port))
        p.set(SEC_LINE, f"{s}_TARGEIP", c.target_ip)
        p.set(SEC_LINE, f"{s}_TARGEPORT", str(c.target_port))
        p.set(SEC_LINE, f"{s}_SOURCE_MANUAL", "1" if c.source_manual else "0")
        p.set(SEC_LINE, f"{s}_MANUAL_RPM", str(c.manual_rpm))
        p.set(SEC_LINE, f"{s}_SOURCE_LINE", c.source_line)
        p.set(SEC_LINE, f"{s}_SOURCE_ID", str(c.source_id))
        p.set(SEC_MOTOR, f"{s}_MOTORS", format_motor_list(motors.get(s, [])))

    buf = io.StringIO()
    p.write(buf, space_around_delimiters=False)
    text = buf.getvalue().replace("\r\n", "\n").replace("\n", "\r\n")

    folder = os.path.dirname(os.path.abspath(path))
    os.makedirs(folder, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".config_", suffix=".tmp", dir=folder)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as f:
            f.write(text)
        os.replace(tmp, path)
    except OSError:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise
