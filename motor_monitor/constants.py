"""프로그램 전역 상수.

원본: M_Variables.vb, C_HOUSE.vb(Enum), MainWindow.xaml.vb(타이머/색상)
스펙시트(docs/SPEC_SHEET.md) '4. 동작 파라미터' 항목과 1:1 대응합니다.
"""
from enum import IntEnum

APP_TITLE = "BLDC MOTOR MONITOR"
CONFIG_FILE_NAME = "Config.ini"

# 라인 코드 목록 (H=온실번호, R=구역번호, E=짝/O=홀)
HOUSE_LINES = [
    "H1R1E", "H1R1O", "H1R2E", "H1R2O", "H1R3E", "H1R3O", "H1R4E", "H1R4O",
    "H2R1E", "H2R1O", "H2R2E", "H2R2O", "H2R3E", "H2R3O", "H2R4E", "H2R4O",
]

# 수동 RPM 최대값
MAX_RPM = 1100

# 모터 ID 허용 범위 (모터 추가 창)
MOTOR_ID_MIN = 1
MOTOR_ID_MAX = 99

# 5~8구역(온실2) 제어 시작/종료 버튼 활성 여부.
# 원본은 H2_StartBtn_Click / H2_StopBtn_Click 첫 줄이 'Exit Sub' 이라 동작하지 않았음.
# v1.1.0 부터 1~8구역을 한 화면에서 운영하므로 활성화.
H2_CONTROL_ENABLED = True

# 신규모터 상태 표시 여부.
# 원본은 신규모터 읽기 응답을 받아도 수신일시(RDate)/RPM을 갱신하지 않아 신규모터는 항상 '대기'로 표시됨.
# True: 응답 수신 시 RPM(레지스터 0x16 값)/수신일시를 갱신해 정상/이상 판정 (v1.1.0 부터 적용, 제어 동작 영향 없음).
NEW_MOTOR_STATUS_TRACKING = True

# ---- 타이밍(초) ----
STATUS_UPDATE_INTERVAL = 5      # 모터 상태 갱신 주기 (C_HOUSE.TMR_STATUS_UPDATER)
NETWORK_UPDATE_INTERVAL = 5     # 통신상태 표시 갱신 주기 (MainWindow.TMR_UPDATE)
MOTOR_TIMEOUT_SEC = 60          # 마지막 수신 후 이 시간 이상 지나면 '이상'
CONNECT_CHECK_TIMEOUT = 1       # 접속 가능 여부 확인 타임아웃
RECONNECT_DELAY = 10            # 접속 실패 시 재시도 대기
SEND_INTERVAL = 0.5             # 명령 1건 전송 후 대기
CYCLE_INTERVAL = 1.0            # 큐 적재 후 대기
ERROR_RETRY_DELAY = 1.0         # 송신 중 예외 발생 후 대기
RESTART_DELAY = 1.0             # Service_Restart 시 중지~시작 사이 대기
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
    READY = 9       # 대기(처음 추가 시 / 아직 수신 없음)


MOTOR_TYPE_NAME = {MotorType.OLD_MOTOR: "기존모터", MotorType.NEW_MOTOR: "신규모터"}
