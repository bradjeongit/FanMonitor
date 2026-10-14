"""프로그램 전역 상수.

기준: 운영 실행파일 BLDC_MOTOR_MONITOR.exe (M_Variables, C_HOUSE)
스펙시트(docs/SPEC_SHEET.md) '5. 동작 파라미터' 항목과 1:1 대응합니다.
"""
from enum import IntEnum

APP_TITLE = "BLDC MOTOR MONITOR"
CONFIG_FILE_NAME = "Config.dll"        # 3DES 로 암호화된 INI (운영 프로그램과 같은 파일)

# 구역 코드 목록: H1R1 = 1구역 ... H1R8 = 8구역 (구역마다 라인 1개)
HOUSE_LINES = ["H1R1", "H1R2", "H1R3", "H1R4", "H1R5", "H1R6", "H1R7", "H1R8"]

# 수동 RPM / 속도제한 RPM 최대값
MAX_RPM = 1100

# 모터 ID 허용 범위 (모터 추가 창)
MOTOR_ID_MIN = 1
MOTOR_ID_MAX = 250

# ---- 타이밍(초) ----
STATUS_UPDATE_INTERVAL = 5      # 모터 상태 갱신 주기 (구역 제어 중일 때만)
NETWORK_UPDATE_INTERVAL = 5     # 통신상태 점 갱신 주기
NO_RESPONSE_WARN_SEC = 60       # 명령 전송 후 이 시간 동안 응답이 없으면 '이상기록'(빨간 !) 표시
MOTOR_TIMEOUT_SEC = 3600        # 마지막 응답 후 이 시간 이상 지나면 '이상' (운영 프로그램 값: 1시간)
CONNECT_CHECK_TIMEOUT = 1       # 접속 제한시간
RECONNECT_DELAY = 10            # 접속 실패 시 재시도 대기
SEND_INTERVAL = 0.3             # 명령 1건 전송 후 대기 (운영 프로그램 SEND_DELAY=300ms)
CYCLE_INTERVAL = 0.3            # 읽기 명령 적재 후 대기
ERROR_RETRY_DELAY = 1.0         # 송신 중 예외 발생 후 대기
RESTART_DELAY = 1.0             # 재시작 시 중지~시작 사이 대기
CMD_IDX_MAX = 9999              # 트랜잭션 번호 최대값(초과 시 0)

# ---- 네트워크 기본값 ----
DEFAULT_PORT = 502
DEFAULT_IP_PREFIX = "192.168.0."


class MotorType(IntEnum):
    OLD_MOTOR = 0   # 기존 모터 (RPM 을 읽어오는 모터)
    NEW_MOTOR = 1   # 신규 모터 (RPM 을 써넣는 모터)


class MotorStatus(IntEnum):
    OK = 0          # 정상
    ERR = 1         # 이상
    NOTUSE = 2      # 중지(사용안함)
    READY = 9       # 대기(제어 시작 후 아직 응답 없음)


MOTOR_TYPE_NAME = {MotorType.OLD_MOTOR: "기존모터", MotorType.NEW_MOTOR: "신규모터"}
