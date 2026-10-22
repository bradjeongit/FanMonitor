# FanMonitor — BLDC MOTOR MONITOR (Python)

온실 기존 팬 모터의 RPM 을 읽어 신규 팬 모터를 자동 제어하고, 전체 모터 상태를 감시하는 프로그램입니다.
현재 운영 중인 실행파일 `BLDC_MOTOR_MONITOR.exe`(VB.NET WPF)를 분석해 같은 기능·같은 화면 구성으로 파이썬에 옮겼습니다.
운영 프로그램의 설정 파일 `Config.dll` 을 그대로 읽고 씁니다.

## 문서
- **사용자 매뉴얼**: [docs/USER_MANUAL.md](docs/USER_MANUAL.md) — 화면 기능별 사용법
- **스펙시트**: [docs/SPEC_SHEET.md](docs/SPEC_SHEET.md) — 구성요소 ID, 동작 규칙, 통신·설정 사양, 원본과의 차이, 변경 이력
  수정 요청 시 스펙시트의 ID(예: `ZONE-*-BTN-SET`, `MSG-LINE-02`)를 사용하면 정확히 전달됩니다.

## 실행
```
python run.py              # Python 3.9+ (표준 라이브러리만 사용)
python run.py --debug      # 통신 패킷 로그 포함
python run.py --monitoring # 모니터링 모드로 시작: 읽기만, 제어 신호 전송 안 함 (run_monitoring.bat)
                           # 실행 중에는 화면 맨 위 띠의 [모니터링 모드로 전환]/[제어 모드로 전환] 버튼으로 변경
ecblue_tester.bat          # 기존모터 시험기 (172.27.0.95, ID 180·176 고정) - 매뉴얼 12장
mini_control.bat           # 미니 제어기: IP·ID 로 팬 추가, 속도 % 또는 참조팬 지정 - 매뉴얼 13장
motor_tester.bat           # 모터 시험기(SUNINUS MOTOR TEST, MOTOR_TESTER.exe 새 구현) - 매뉴얼 14장
build_motor_tester_exe.bat # 모터 시험기 실행파일 dist\MOTOR_TESTER.exe 만들기 (GitHub Actions 에서도 자동 빌드)
motor_scanner.bat          # 모터 스캐너: IP 의 모터 ID 찾기, ECblue 설정·통신 끊김 시 자동 운전 - 매뉴얼 15장
build_motor_scanner_exe.bat # 모터 스캐너 실행파일 dist\MOTOR_SCANNER.exe 만들기 (GitHub Actions 에서도 자동 빌드)
build_mini_exe.bat         # 미니 제어기 실행파일 dist\MINI_CONTROL.exe 만들기 (GitHub Actions 에서도 자동 빌드)
build_tester_exe.bat       # 시험기 실행파일 dist\ECBLUE_TESTER.exe 만들기 (GitHub Actions 에서도 자동 빌드)
build_exe.bat              # Windows 실행파일 만들기 (PyInstaller)
```
설정은 실행 폴더의 `Config.dll`(운영 프로그램과 같은 암호화 파일), 로그는 `logs/` 에 저장됩니다.
`Config.dll` 은 현장 정보가 들어 있어 저장소에 올리지 않습니다(.gitignore).

## 최신 버전으로 업데이트
git 으로 받은 폴더에서 `update.bat` 을 더블클릭하면 GitHub `main` 의 최신 내용으로 바뀝니다.
`Config.dll` 과 `logs/` 는 git 관리 대상이 아니라 업데이트해도 그대로 유지됩니다.

## 시험
```
python -m pytest                                  # 로직 시험
xvfb-run -a python -m pytest                      # (리눅스) 화면 시험 포함
python -m motor_monitor.simulator --port 5020 --old 2:770 --new 1,3   # 가상 모터 통신장치
python -m motor_monitor.config Config.dll                           # 설정 파일 평문 확인
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
reference/original_vb/     처음 받은 구버전 VB.NET 소스 (참고용, 현재 운영 버전과 다름)
```
