"""설정 파일(Config.dll) 읽기/쓰기. 기준: 운영 실행파일 M_Config, M_Variables(Struct_Config)

Config.dll = 3DES 로 암호화한 INI 텍스트 (crypto3des.py).
운영 프로그램과 같은 파일·같은 섹션/키 이름을 사용하므로 현장 Config.dll 을 그대로 읽고 쓸 수 있습니다.
(키 이름의 오타 '_TARGEIP', '_TARGEPORT' 도 호환을 위해 그대로 유지)
"""
from __future__ import annotations

import logging
import os
import sys
import tempfile
from dataclasses import dataclass, replace
from typing import Dict, List, NamedTuple, Tuple

from .constants import CONFIG_FILE_NAME, DEFAULT_IP_PREFIX, DEFAULT_PORT, HOUSE_LINES, MotorType
from .crypto3des import decrypt_config, encrypt_config

log = logging.getLogger(__name__)

SEC_LINE = "LINE_CONFIG"
SEC_MOTOR = "MOTOR_CONFIG"
SEC_CTRL = "CTRL_MOTOR_CONFIG"   # 기존모터제어 목록 (이 프로그램 전용 섹션, 운영 프로그램은 무시)


class ConfigError(Exception):
    """설정 파일을 해독할 수 없음 (운영 프로그램: '프로그램 설정 정보가 처리중 오류가 발생했습니다.')"""


@dataclass
class LineConfig:
    """구역 1개의 통신/제어 설정."""
    source_ip: str          # 기존모터 IP
    source_port: int        # 기존모터 PORT
    target_ip: str          # 신규모터 IP
    target_port: int        # 신규모터 PORT
    source_manual: bool     # True: 수동 RPM / False: 기존모터 RPM 참조
    manual_rpm: int         # 수동 RPM 값
    source_line: str        # 참조 기존모터가 있는 구역 코드 (예: H1R1)
    source_id: int          # 참조 기존모터 ID

    def copy(self, **changes) -> "LineConfig":
        return replace(self, **changes)


class MotorEntry(NamedTuple):
    id: int
    mtype: MotorType
    use: bool
    use_limit: bool = False     # 속도제한 사용 (신규모터)
    limit_rpm: int = 0          # 속도제한 최대 RPM
    ip: str = ""                # 기존모터제어: 접속 IP
    port: int = 0               # 기존모터제어: 접속 PORT
    set_rpm: int = 0            # 기존모터제어: 지정 운전 속도 (RPM)


def app_base_dir() -> str:
    """실행 파일(또는 run.py)이 있는 폴더 (운영 프로그램의 AppDomain.BaseDirectory)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def default_config_path() -> str:
    return os.path.join(app_base_dir(), CONFIG_FILE_NAME)


# ------------------------------------------------------------------ INI 처리
def parse_ini(text: str) -> Dict[str, Dict[str, str]]:
    """Windows GetPrivateProfileString 과 같은 규칙의 단순 INI 해석 (키 대소문자 무시, 첫 값 사용)."""
    data: Dict[str, Dict[str, str]] = {}
    section = None
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith(";"):
            continue
        if line.startswith("[") and "]" in line:
            section = line[1:line.index("]")].strip().upper()
            data.setdefault(section, {})
        elif section is not None and "=" in line:
            k, v = line.split("=", 1)
            data[section].setdefault(k.strip().upper(), v.strip())
    return data


def _ini_read(ini, section: str, key: str, default: str) -> str:
    """운영 INIRead 와 동일: 값이 없으면 기본값, '&' '|' 문자 제거."""
    val = ini.get(section.upper(), {}).get(key.upper(), default)
    return val.replace("&", "").replace("|", "")


def _to_int(text: str, default: int = 0) -> int:
    try:
        return int(text.strip())
    except (ValueError, AttributeError):
        return default


def parse_motor_list(data: str) -> List[MotorEntry]:
    """'2=0=1=0=0,13=1=1=1=500' -> [MotorEntry...]. 예전 3항목 형식(ID=종류=사용)도 읽음."""
    result: List[MotorEntry] = []
    for item in filter(None, data.split(",")):
        parts = item.split("=")
        if len(parts) < 3:
            log.warning("모터 설정 형식 오류(무시): %r", item)
            continue
        try:
            mtype = MotorType(_to_int(parts[1]))
            if mtype == MotorType.OLD_CTRL:      # 기존모터제어는 별도 섹션에만 기록됨
                raise ValueError
        except ValueError:
            log.warning("알 수 없는 모터 종류(무시): %r", item)
            continue
        use_limit, limit = False, 0
        if len(parts) == 5:
            use_limit = _to_int(parts[3]) != 0
            limit = _to_int(parts[4])
        result.append(MotorEntry(_to_int(parts[0]), mtype, _to_int(parts[2]) != 0, use_limit, limit))
    return result


def format_motor_list(motors: List[MotorEntry]) -> str:
    """[MOTOR_CONFIG] 형식 (운영 프로그램과 같음). 기존모터제어는 제외."""
    return ",".join(
        f"{m.id}={int(m.mtype)}={1 if m.use else 0}={1 if m.use_limit else 0}={m.limit_rpm}"
        for m in motors if m.mtype != MotorType.OLD_CTRL)


def parse_ctrl_list(data: str) -> List[MotorEntry]:
    """'2@172.27.0.82:502=1=700,...' -> 기존모터제어 목록 (ID@IP:PORT=사용=운전속도)."""
    result: List[MotorEntry] = []
    for item in filter(None, data.split(",")):
        try:
            head, use, rpm = item.split("=")
            mid, addr = head.split("@")
            ip, port = addr.rsplit(":", 1)
            result.append(MotorEntry(int(mid), MotorType.OLD_CTRL, use.strip() != "0",
                                     ip=ip.strip(), port=int(port), set_rpm=int(rpm)))
        except ValueError:
            log.warning("기존모터제어 설정 형식 오류(무시): %r", item)
    return result


def format_ctrl_list(motors: List[MotorEntry]) -> str:
    return ",".join(f"{m.id}@{m.ip}:{m.port}={1 if m.use else 0}={m.set_rpm}"
                    for m in motors if m.mtype == MotorType.OLD_CTRL)


def read_config_text(path: str) -> str:
    """Config.dll 을 읽어 INI 텍스트 반환. 파일이 없거나 비어 있으면 빈 문자열."""
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        raw = f.read()
    try:
        return decrypt_config(raw)
    except (ValueError, UnicodeDecodeError) as e:
        raise ConfigError(f"설정 파일을 해독할 수 없습니다: {path} ({e})") from e


def load_config(path: str) -> Tuple[Dict[str, LineConfig], Dict[str, List[MotorEntry]]]:
    """설정 읽기. 값이 없으면 운영 프로그램과 같은 기본값."""
    ini = parse_ini(read_config_text(path))
    configs: Dict[str, LineConfig] = {}
    motors: Dict[str, List[MotorEntry]] = {}
    idx = 1
    for s in HOUSE_LINES:
        configs[s] = LineConfig(
            source_ip=_ini_read(ini, SEC_LINE, f"{s}_SOURCE_IP", f"{DEFAULT_IP_PREFIX}{idx}"),
            source_port=_to_int(_ini_read(ini, SEC_LINE, f"{s}_SOURCE_PORT", str(DEFAULT_PORT))),
            target_ip=_ini_read(ini, SEC_LINE, f"{s}_TARGEIP", f"{DEFAULT_IP_PREFIX}{idx + 1}"),
            target_port=_to_int(_ini_read(ini, SEC_LINE, f"{s}_TARGEPORT", str(DEFAULT_PORT))),
            source_manual=_ini_read(ini, SEC_LINE, f"{s}_SOURCE_MANUAL", "1") != "0",
            manual_rpm=_to_int(_ini_read(ini, SEC_LINE, f"{s}_MANUAL_RPM", "0")),
            source_line=_ini_read(ini, SEC_LINE, f"{s}_SOURCE_LINE", ""),
            source_id=_to_int(_ini_read(ini, SEC_LINE, f"{s}_SOURCE_ID", "0")),
        )
        motors[s] = (parse_motor_list(_ini_read(ini, SEC_MOTOR, f"{s}_MOTORS", ""))
                     + parse_ctrl_list(_ini_read(ini, SEC_CTRL, f"{s}_MOTORS", "")))
        idx += 2
    return configs, motors


def build_config_text(configs: Dict[str, LineConfig], motors: Dict[str, List[MotorEntry]]) -> str:
    """운영 프로그램이 만드는 것과 같은 순서/형식의 INI 텍스트 (CRLF)."""
    lines = [f"[{SEC_LINE}]"]
    for s in HOUSE_LINES:
        c = configs[s]
        lines += [
            f"{s}_SOURCE_IP={c.source_ip}",
            f"{s}_SOURCE_PORT={c.source_port}",
            f"{s}_TARGEIP={c.target_ip}",
            f"{s}_TARGEPORT={c.target_port}",
            f"{s}_SOURCE_MANUAL={1 if c.source_manual else 0}",
            f"{s}_MANUAL_RPM={c.manual_rpm}",
            f"{s}_SOURCE_LINE={c.source_line}",
            f"{s}_SOURCE_ID={c.source_id}",
        ]
    lines.append(f"[{SEC_MOTOR}]")
    for s in HOUSE_LINES:
        lines.append(f"{s}_MOTORS={format_motor_list(motors.get(s, []))}")
    # 기존모터제어가 하나라도 있을 때만 섹션 기록 (없으면 운영 프로그램과 바이트 단위로 같은 파일 유지)
    if any(m.mtype == MotorType.OLD_CTRL for s in HOUSE_LINES for m in motors.get(s, [])):
        lines.append(f"[{SEC_CTRL}]")
        for s in HOUSE_LINES:
            lines.append(f"{s}_MOTORS={format_ctrl_list(motors.get(s, []))}")
    return "\r\n".join(lines) + "\r\n"


def save_config(path: str, configs: Dict[str, LineConfig], motors: Dict[str, List[MotorEntry]]) -> None:
    """암호화해 저장. 임시파일에 쓴 뒤 교체(저장 중 전원 차단 시 기존 파일 보호)."""
    data = encrypt_config(build_config_text(configs, motors))
    folder = os.path.dirname(os.path.abspath(path))
    os.makedirs(folder, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".config_", suffix=".tmp", dir=folder)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        os.replace(tmp, path)
    except OSError:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def export_plain_text(path: str) -> str:
    """점검용: Config.dll 내용을 평문 INI 로 반환 (python -m motor_monitor.config 경로)."""
    return read_config_text(path)


if __name__ == "__main__":
    print(export_plain_text(sys.argv[1] if len(sys.argv) > 1 else default_config_path()))
