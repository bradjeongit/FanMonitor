# FanMonitor — BLDC MOTOR MONITOR (Python)

온실 기존 팬 모터의 RPM 을 읽어 신규 팬 모터를 자동 제어하고, 전체 모터 상태를 감시하는 프로그램입니다.
원본 VB.NET(WPF) 프로그램 `Motor_Control` 을 같은 기능·같은 화면 구성으로 파이썬에 옮겼습니다.

## 문서
- **사용자 매뉴얼**: [docs/USER_MANUAL.md](docs/USER_MANUAL.md) — 화면 기능별 사용법
- **스펙시트**: [docs/SPEC_SHEET.md](docs/SPEC_SHEET.md) — 구성요소 ID, 동작 규칙, 통신·설정 사양, 원본과의 차이, 변경 이력
  수정 요청 시 스펙시트의 ID(예: `LINE-*-BTN-SET`, `MSG-LINE-02`)를 사용하면 정확히 전달됩니다.

## 실행
```
python run.py              # Python 3.9+ (표준 라이브러리만 사용)
python run.py --debug      # 통신 패킷 로그 포함
build_exe.bat              # Windows 실행파일 만들기 (PyInstaller)
```
설정은 실행 폴더의 `Config.ini`(원본 프로그램과 형식 호환), 로그는 `logs/` 에 저장됩니다.

## 시험
```
python -m pytest                                  # 로직 시험
xvfb-run -a python -m pytest                      # (리눅스) 화면 시험 포함
python -m motor_monitor.simulator --port 5020 --old 5:950 --new 10   # 가상 모터 통신장치
```

## 폴더 구성
```
run.py                     실행 시작점
motor_monitor/             프로그램 본체 (구성은 스펙시트 10장)
  ui/                      화면 (메인, 대화상자, 위젯, 색상)
  assets/                  그림 파일
tests/                     자동 시험
tools/                     문서용 화면 캡처 스크립트
docs/                      사용자 매뉴얼, 스펙시트, 화면 그림
reference/original_vb/     원본 VB.NET 소스 (비교·참고용, 빌드 산출물/인증서 제외)
```
