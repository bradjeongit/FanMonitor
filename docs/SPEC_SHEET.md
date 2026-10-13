# BLDC MOTOR MONITOR — 스펙시트 (Spec Sheet)

> 이 문서는 **수정 요청 시 의사소통 기준 문서**입니다.
> 화면·창·버튼 등 모든 구성요소에 **고유 ID** 를 붙였습니다. 수정을 요청할 때 이 ID 를 사용하면 정확하게 전달됩니다.
> 코드가 바뀔 때마다 이 문서도 함께 갱신합니다(맨 아래 **변경 이력** 참조).

| 항목 | 내용 |
|---|---|
| 문서 버전 | 1.0.0 (프로그램 버전과 동일) |
| 대상 프로그램 | `motor_monitor` 패키지 (실행: `run.py`) |
| 원본 | VB.NET WPF `Motor_Control` (BLDC_MOTOR_MONITOR) — `reference/original_vb/` 에 원본 소스 보관 |
| 사용자 매뉴얼 | [USER_MANUAL.md](USER_MANUAL.md) |

---

## 1. ID 명명 규칙과 수정 요청 방법

### 1.1 ID 체계

| 접두어 | 의미 | 예 |
|---|---|---|
| `SCR-` | 화면(메인 창) | `SCR-MAIN` |
| `DLG-` | 대화상자(팝업 창) | `DLG-LINE`, `DLG-MOTOR` |
| `MAIN-` | 메인 화면 상단 구성요소 | `MAIN-BTN-H1-START` |
| `ZONE-` | 구역 패널 | `ZONE-H1R1` (온실1 1구역) |
| `LINE-` | 라인 패널(구역의 짝/홀 한 쪽) | `LINE-H1R1E` (온실1 1구역 짝) |
| `MSG-` | 메시지 창(확인/오류 알림) | `MSG-LINE-02` |
| `P-` | 동작 파라미터(설정 상수) | `P-MOTOR-TIMEOUT` |
| `C-` | 색상 | `C-NET-OK` |

- 구성요소 종류 약어: `TAB` 탭, `BTN` 버튼, `TXT` 입력칸, `LBL` 표시 글자, `RB` 라디오버튼(선택 동그라미), `CB` 콤보박스(펼침 목록), `DOT` 상태 점, `LIST` 목록, `HDR` 머리글, `COL` 열, `ROW` 줄, `TITLE` 창 제목줄.
- 라인 코드: `H{온실}R{구역}{E=짝 | O=홀}` — 예) `H1R3O` = 온실1 / 3구역 / 홀.
- 16개 라인에 공통인 요소는 `LINE-*-BTN-SET` 처럼 `*` 로 표기합니다. 특정 라인만 지칭할 때는 `LINE-H1R2E-BTN-SET` 처럼 씁니다.

### 1.2 수정 요청 예시

- "`LINE-*-LBL-REF` 글자를 굵게 해줘" → 16개 라인 모두의 RPM 소스 문구 변경
- "`LINE-H1R1E-LIST` 의 `COL-STATUS` 에 마지막 수신시각도 보여줘"
- "`DLG-MOTOR-TXT-ID` 허용 범위를 1~247 로 늘려줘" (관련 파라미터 `P-MOTOR-ID-RANGE`)
- "`MSG-MAIN-04` 문구를 바꿔줘"
- "`C-NET-ERR` 색을 주황색으로"

---

## 2. 화면/창 목록

| ID | 이름 | 여는 방법 | 크기 | 코드 (파일 / 클래스) | 원본 |
|---|---|---|---|---|---|
| `SCR-MAIN` | 메인 화면 | 프로그램 실행 | 기본 1280×820, 최소 1000×650, 크기조절 가능 | `motor_monitor/ui/main_window.py` / `MainWindow` | `MainWindow.xaml` |
| `DLG-LINE` | 라인설정 창 | `LINE-*-BTN-SET` 클릭 | 400×300 고정, 모달 | `motor_monitor/ui/dialogs.py` / `LineSettingsDialog` | `wHOUSE.xaml` |
| `DLG-MOTOR` (ADD) | 모터 추가 창 | `LINE-*-BTN-ADD` 클릭 | 400×260 고정, 모달 | `dialogs.py` / `MotorDialog(mode="ADD")` | `wMotor.xaml` |
| `DLG-MOTOR` (MOD) | 모터 수정/삭제 창 | `LINE-*-LIST` 항목 더블클릭 | 400×260 고정, 모달 | `dialogs.py` / `MotorDialog(mode="MOD")` | `wMotor.xaml` |

---

## 3. 구성요소 상세

### 3.1 `SCR-MAIN` 메인 화면

![메인 화면 구성요소 번호](images/main_running_annotated.png)

| 번호 | ID | 이름 | 종류 | 코드 객체 | 표시 / 동작 |
|---|---|---|---|---|---|
| 1 | `MAIN-TAB-H1` | 온실1 탭 | 탭 | `MainWindow.house_tab` (0번) | 클릭 시 온실1 화면 표시. 선택 시 녹색(`C-TAB-SEL`)·흰 굵은 글씨 |
| 2 | `MAIN-TAB-H2` | 온실2 탭 | 탭 | `MainWindow.house_tab` (1번) | 클릭 시 빈 화면 (원본과 동일하게 내용 없음) |
| 3 | `MAIN-BTN-H1-START` | 온실1 제어 시작 버튼 | 버튼 | `MainWindow.H1_StartBtn` | 확인(`MSG-MAIN-02`) 후 온실1 8개 라인 통신·제어 시작 → 버튼이 `MAIN-BTN-H1-STOP` 으로 바뀜. 위치 x=250 |
| 3 | `MAIN-BTN-H1-STOP` | 온실1 제어 종료 버튼 | 버튼 | `MainWindow.H1_StopBtn` | 확인(`MSG-MAIN-03`) 후 통신·제어 중지 → 시작 버튼으로 바뀜. 같은 위치 |
| 4 | `MAIN-BTN-H2-START` | 온실2 제어 시작 버튼 | 버튼 | `MainWindow.H2_StartBtn` | **동작 안 함** (원본과 동일, `P-H2-CONTROL`=False). 위치 x=400 |
| 4 | `MAIN-BTN-H2-STOP` | 온실2 제어 종료 버튼 | 버튼 | `MainWindow.H2_StopBtn` | 동작 안 함 (위와 같음) |
| 5 | `ZONE-H1R{n}` | 구역 패널 (n=1~4) | 묶음 | `ZonePanel` | 2×2 배치: 1구역 좌상, 2구역 우상, 3구역 좌하, 4구역 우하 |
| 5 | `ZONE-*-TITLE` | 구역 제목줄 | 글자 | `ZonePanel` 내부 Label | "n구역", 배경 `C-ZONE-TITLE`, 흰 굵은 글씨 |
| 6 | `LINE-{코드}` | 라인 패널 | 묶음 | `LinePanel` (`MainWindow.line_panels[코드]`) | 구역 안 왼쪽=짝(E), 오른쪽=홀(O) |
| 6 | `LINE-*-SIDE` | 짝/홀 표시 띠 | 글자 | `LinePanel` 내부 | 짝=주황(`C-SIDE-EVEN`), 홀=파랑(`C-SIDE-ODD`) |
| 7 | `LINE-*-DOT-OLD` | 기존모터 통신상태 점 | 상태 점 | `LinePanel.rec_SOURCE_NETSTATUS` | 4.1절 색상 규칙 |
| 7 | `LINE-*-LBL-OLD` | 기존모터 IP/PORT | 글자 | `LinePanel.lbl_old_ip` | "IP/PORT" 예) `192.168.0.1/502` |
| 8 | `LINE-*-DOT-NEW` | 신규모터 통신상태 점 | 상태 점 | `LinePanel.rec_TARGET_NETSTATUS` | 4.1절 색상 규칙 |
| 8 | `LINE-*-LBL-NEW` | 신규모터 IP/PORT | 글자 | `LinePanel.lbl_new_ip` | "IP/PORT" |
| 9 | `LINE-*-DOT-REF` | RPM 소스 상태 점 | 상태 점 | `LinePanel.rec_SID_NETSTATUS` | 수동이면 항상 정상, 참조이면 참조 기존모터가 '정상'일 때만 정상 |
| 9 | `LINE-*-LBL-REF` | RPM 소스 문구 | 글자 | `LinePanel.lbl_source` | 수동: `수동 > 800 RPM` / 참조: `1구역 > 짝 > 5` (구역 > 짝홀 > 모터ID) |
| 10 | `LINE-*-BTN-SET` | 설정 버튼(톱니) | 그림 버튼 | `LinePanel.Btn_Settings` | 클릭 → `DLG-LINE`. 툴팁 "설정" |
| 11 | `LINE-*-BTN-ADD` | 추가 버튼(+) | 그림 버튼 | `LinePanel.Btn_Add` | 클릭 → `DLG-MOTOR`(ADD). 툴팁 "추가" |
| 12 | `LINE-*-LIST-HDR` | 모터 목록 머리글 | 머리글 | `MotorListView.header` | 열 머리글 클릭 시 정렬 (3.2절) |
| 13 | `LINE-*-LIST` | 모터 목록 | 목록 | `LinePanel.LV` (`MotorListView`) | 3.2절 |

### 3.2 `LINE-*-LIST` 모터 목록

| ID | 이름 | 폭 | 내용 | 머리글 클릭(정렬 기준) |
|---|---|---|---|---|
| `COL-TYPE` | 종류 열 (머리글 빈칸) | 35px | 기존모터 그림(회색 팬) / 신규모터 그림(녹색 팬), 툴팁 "기존모터"/"신규모터" | 종류 (`MTYPE`) |
| `COL-ID` | ID 열 | 60px | 모터 ID | ID |
| `COL-STATUS` | 상태 열 | 나머지 | 상태 그림 + 오른쪽 정렬 문구 (아래 표) | 상태 코드 |

**상태 표시 규칙 (`COL-STATUS`)**

| 상태 코드 | 상태 | 그림 | 문구 | 글자 모양 |
|---|---|---|---|---|
| 0 | 정상 | 녹색 체크(good) — 툴팁 "정상" | 현재 RPM 숫자 | 검정, 보통 |
| 1 | 이상 | 빨간 X(bad) — 툴팁 "이상" | `이상` | 빨강, 굵게 |
| 2 | 중지(사용안함) | 회색 느낌표(nouse) — 툴팁 "중지(사용안함)" | `중지` | 연회색, 굵게 |
| 9 | 대기 | **없음** (원본 동일) | `대기` | 연회색, 굵게 |

**정렬 규칙** (원본 `C_HOUSE.Sort` 와 동일)
- 처음 상태: 종류 내림차순 = **신규모터 먼저**, 같은 종류 안에서는 ID 오름차순.
- 다른 열을 클릭하면 그 열 기준 **오름차순**, 같은 열을 다시 클릭하면 오름/내림 전환.
- 종류 오름차순 = 기존모터 먼저(ID 오름차순).
- 정렬은 클릭·추가·삭제 시점에만 다시 계산합니다. 상태가 바뀌어도 자동 재정렬하지 않습니다(원본 동일).

**기타 동작**
- 한 번 클릭: 항목 선택(하늘색 `C-LIST-SEL`). 마우스 올림: 연하늘색(`C-LIST-HOVER`).
- 더블클릭: `DLG-MOTOR`(MOD) 열기.
- 목록 위 툴팁: "항목 더블클릭시 수정/삭제".
- 항목이 많으면 세로 스크롤바 자동 표시, 마우스 휠 스크롤.
- 화면 갱신: 데이터가 바뀌면 0.3초 이내 반영(`P-LIST-REFRESH`).

### 3.3 `DLG-LINE` 라인설정 창

![라인설정 창 (기존모터 참조 모드)](images/dlg_line_ref_annotated.png) ![라인설정 창 (수동 모드)](images/dlg_line_manual_annotated.png)

| 번호 | ID | 이름 | 종류 | 코드 객체 | 초기값 / 동작 |
|---|---|---|---|---|---|
| 1 | `DLG-LINE-TITLE` | 창 제목줄 "라인설정" | 제목줄 | `header`, `title_label` | 배경 `C-DLG-LINE`. 끌어서 창 이동 |
| 2 | `DLG-LINE-LBL-HOUSE` | 위치 | 글자 | `lbl_house` | 예) `온실1 / 1구역 / 짝` |
| 3 | `DLG-LINE-TXT-SIP` | 기존모터 IP | 입력칸 | `txtSIP` | 현재 설정값 |
| 4 | `DLG-LINE-TXT-SPORT` | 기존모터 PORT | 입력칸 | `txtSPORT` | 현재 설정값 |
| 5 | `DLG-LINE-TXT-TIP` | 신규모터 IP | 입력칸 | `txtTIP` | 현재 설정값 |
| 6 | `DLG-LINE-TXT-TPORT` | 신규모터 PORT | 입력칸 | `txtTPORT` | 현재 설정값 |
| 7 | `DLG-LINE-RB-MANUAL` | RPM 제어 소스: 수동 | 라디오 | `rb_source_manual` | 선택 시 `DLG-LINE-ROW-MANUAL` 표시 |
| 8 | `DLG-LINE-RB-REF` | RPM 제어 소스: 기존모터 | 라디오 | `rb_source_ref` | 선택 시 `DLG-LINE-ROW-REF` 표시 |
| 9 | `DLG-LINE-CB-LINE` | 참조 라인 선택 | 콤보 | `cb_lines` (`DLG-LINE-ROW-REF` 안) | 항목: 라인 / 1구역/짝 … 4구역/홀 (같은 온실 안). 바꾸면 `CB-ID` 목록 다시 채움 |
| 10 | `DLG-LINE-CB-ID` | 참조 모터 ID 선택 | 콤보 | `cb_ids` | 첫 항목 "ID" + 선택 라인의 **기존모터** ID (등록 순서) |
| 11 | `DLG-LINE-BTN-SAVE` | 저장 | 버튼(진청) | `BtnSave` | 검증 → 저장 → 메인 화면 머리 영역 갱신 → 닫기. Enter 키 |
| 12 | `DLG-LINE-BTN-CANCEL` | 취소 | 버튼(회색) | `BtnCancel` | 저장 없이 닫기. Esc 키 |
| 13 | `DLG-LINE-TXT-MRPM` | 수동RPM값 | 입력칸 | `txtMRPM` (`DLG-LINE-ROW-MANUAL` 안) | 현재 설정값 |
| 14 | `DLG-LINE-LBL-RANGE` | 허용 범위 안내 | 글자 | `lbl_rpm_range` | `정수(0~1100)` (`P-MAX-RPM`) |
| – | `DLG-LINE-ROW-MANUAL` | 수동 RPM 줄 | 줄 | `manual_line` | 수동 선택 시만 보임 |
| – | `DLG-LINE-ROW-REF` | 참조기존모터 줄 | 줄 | `ref_line` | 기존모터 선택 시만 보임 |

**초기값 규칙**: 참조 라인이 비어 있으면 자기 라인을 기본 선택. 참조 ID 가 목록에 있으면 그 ID 를 선택, 없으면 "ID".

**저장 시 검증 순서** (첫 오류에서 멈추고 메시지 표시, 제목 "에러")

| 순서 | 조건 | 메시지 ID |
|---|---|---|
| 1 | 기존모터 IP 비어 있음 | `MSG-LINE-01` |
| 2 | 기존모터 IP 형식 오류 (숫자·점만, 7~15자, 점 3개, 각 자리 0~255) | `MSG-LINE-02` |
| 3 | 기존모터 PORT 비어 있음 | `MSG-LINE-03` |
| 4 | 기존모터 PORT 숫자 아님 | `MSG-LINE-04` |
| 5 | 기존모터 PORT 1~65535 밖 | `MSG-LINE-05` |
| 6~10 | 신규모터 IP/PORT 에 대해 1~5와 동일 | `MSG-LINE-06`~`10` |
| 11 | (수동) RPM 이 정수 아님 | `MSG-LINE-11` |
| 12 | (수동) RPM 이 0~1100 밖 | `MSG-LINE-12` |
| 13 | (기존모터) 참조 라인 미선택 | `MSG-LINE-13` |
| 14 | (기존모터) 참조 ID 미선택 | `MSG-LINE-14` |
| 15 | 소스 미선택 | `MSG-LINE-15` |

**저장 값**: 수동 모드에서도 선택된 참조 라인/ID 를 함께 저장(ID 미선택 시 -1). 기존모터 모드에서는 수동 RPM 칸 값을 정수로 읽을 수 있으면 저장, 아니면 0. 참조 라인은 `온실코드 + 선택값` (예: `H1` + `R2O` = `H1R2O`).
저장된 IP/PORT 는 **다음 접속 시도부터** 적용됩니다(제어 중이어도 재시작 불필요, 단 현재 연결이 끊겨야 새 주소로 접속).

### 3.4 `DLG-MOTOR` 모터 추가 / 수정·삭제 창

![모터 추가 창](images/dlg_motor_add_annotated.png) ![모터 수정/삭제 창](images/dlg_motor_mod_annotated.png)

| 번호 | ID | 이름 | 종류 | 코드 객체 | 모드 | 동작 |
|---|---|---|---|---|---|---|
| 1 | `DLG-MOTOR-TITLE` | 창 제목줄 | 제목줄 | `header`, `title_label` | 공통 | ADD: "모터 추가" / MOD: "모터 수정/삭제". 배경 `C-DLG-MOTOR`. 끌어서 이동 |
| 2 | `DLG-MOTOR-LBL-HOUSE` | 위치 | 글자 | `lbl_house` | 공통 | 예) `온실1 / 1구역 / 짝` |
| 3 | `DLG-MOTOR-RB-OLD` | 모터종류: 기존모터 | 라디오 | `rb_oldmotor` | ADD | 처음엔 선택 없음 |
| 4 | `DLG-MOTOR-RB-NEW` | 모터종류: 신규모터 | 라디오 | `rb_newmotor` | ADD | 처음엔 선택 없음 |
| 5 | `DLG-MOTOR-TXT-ID` | 모터ID 입력 | 입력칸 | `txtID` | ADD | 옆에 `정수(1~99)` 안내 |
| 6 | `DLG-MOTOR-RB-USE` | 사용여부: 사용 | 라디오 | `rb_use` | 공통 | ADD: 선택 없음 / MOD: 현재값 |
| 7 | `DLG-MOTOR-RB-NOUSE` | 사용여부: 중지 | 라디오 | `rb_nouse` | 공통 | 〃 |
| 8 | `DLG-MOTOR-BTN-SAVE` | 추가 / 수정 버튼 | 버튼(진청) | `BtnSave` | 공통 | ADD: "추가", MOD: "수정". Enter 키 |
| 9 | `DLG-MOTOR-BTN-CANCEL` | 취소 | 버튼(회색) | `BtnCancel` | 공통 | 닫기. Esc 키 |
| 10 | `DLG-MOTOR-LBL-TYPE` | 모터종류 표시 | 그림+글자 | `mod_type_line` (`img_type`, `lbl_type`) | MOD | 수정 불가 |
| 11 | `DLG-MOTOR-LBL-ID` | 모터ID 표시 | 글자 | `lblID` | MOD | 수정 불가 |
| 12 | `DLG-MOTOR-BTN-DELETE` | 삭제 | 버튼(빨강) | `BtnDelete` | MOD | 3.4.2 |

**3.4.1 추가/수정 처리 순서**
1. (ADD) 종류 미선택 → `MSG-MOTOR-01` / ID 비었거나 숫자 아님 → `MSG-MOTOR-02` / 1~99 밖 → `MSG-MOTOR-03`
2. 사용여부 미선택 → `MSG-MOTOR-04`
3. 확인 `MSG-MOTOR-05` (아니오 → 창 유지)
4. (ADD) 같은 종류+ID 가 이미 있으면 `MSG-MOTOR-06` (창 유지) / 없으면 추가 (사용=대기, 중지=중지 상태로 시작)
5. (MOD) 사용여부 반영. 사용→중지로 바꾼 **신규모터**는 다음 주기에 RPM 0 을 한 번 전송.
6. Config.ini 저장 → 창 닫기 → 목록 갱신

**3.4.2 삭제 처리 순서**
1. 확인 `MSG-MOTOR-07` (아니오 → 취소)
2. 기존모터이고, 어떤 라인이 이 모터를 RPM 참조로 쓰는 중(기존모터 모드)이면 `MSG-MOTOR-08` 후 중단
3. 삭제 → Config.ini 저장 → 창 닫기

### 3.5 메시지 창 목록 (`MSG-`)

| ID | 제목 | 문구 | 버튼 | 발생 위치 |
|---|---|---|---|---|
| `MSG-MAIN-01` | BLDC MOTOR MONITOR | 모니터가 이미 실행중입니다. | 확인 | 프로그램 중복 실행 (후 종료) |
| `MSG-MAIN-02` | 제어 | 온실1 제어를 시작하시겠습니까? | 예/아니오 | `MAIN-BTN-H1-START` |
| `MSG-MAIN-03` | 제어 | 온실1 제어를 종료하시겠습니까? | 예/아니오 | `MAIN-BTN-H1-STOP` |
| `MSG-MAIN-04` | 종료 | 종료하시겠습니까? 종료시 제어기능이 동작하지 않습니다. | 예/아니오 | 메인 창 닫기(X) |
| `MSG-MAIN-05` | 오류 | 설정 저장에 실패했습니다. (+오류내용) | 확인 | 종료 시 Config.ini 저장 실패 (신규) |
| `MSG-LINE-01` | 에러 | 기존모터 접속 IP를 입력해 주세요. | 확인 | `DLG-LINE` 저장 |
| `MSG-LINE-02` | 에러 | 기존모터 접속 IP가 올바른 형식이 아닙니다.. | 확인 | 〃 |
| `MSG-LINE-03` | 에러 | 기존모터 접속 PORT를 입력해 주세요. | 확인 | 〃 |
| `MSG-LINE-04` | 에러 | 기존모터 접속 PORT를 숫자로 입력해 주세요. | 확인 | 〃 |
| `MSG-LINE-05` | 에러 | 기존모터 접속 PORT를 숫자로 1~65535범위로 입력해 주세요. | 확인 | 〃 |
| `MSG-LINE-06` | 에러 | 신규 모터 접속 IP를 입력해 주세요. | 확인 | 〃 |
| `MSG-LINE-07` | 에러 | 신규모터 접속 IP가 올바른 형식이 아닙니다.. | 확인 | 〃 |
| `MSG-LINE-08` | 에러 | 신규모터 접속 PORT를 입력해 주세요. | 확인 | 〃 |
| `MSG-LINE-09` | 에러 | 신규모터 접속 PORT를 숫자로 입력해 주세요. | 확인 | 〃 |
| `MSG-LINE-10` | 에러 | 신규모터 접속 PORT를 숫자로 1~65535범위로 입력해 주세요. | 확인 | 〃 |
| `MSG-LINE-11` | 에러 | 수동 RPM속도를 숫자로 입력해주세요. | 확인 | 〃 |
| `MSG-LINE-12` | 에러 | 수동 RPM속도를 0~1100범위내로 입력해주세요. | 확인 | 〃 |
| `MSG-LINE-13` | 에러 | 참조할 모터가 있는 라인을 선택해주세요. | 확인 | 〃 |
| `MSG-LINE-14` | 에러 | 참조할 모터 ID를 선택해주세요. | 확인 | 〃 |
| `MSG-LINE-15` | 에러 | RPM제어 소스를 선택해주세요. | 확인 | 〃 |
| `MSG-MOTOR-01` | 오류 | 모터 종류을 선택해 주세요. | 확인 | `DLG-MOTOR` 추가 |
| `MSG-MOTOR-02` | 오류 | 모터 ID값을 숫자로 입력해주세요. | 확인 | 〃 |
| `MSG-MOTOR-03` | 오류 | 모터ID값은 숫자로 1-99범위로 입력하셔야 합니다. | 확인 | 〃 |
| `MSG-MOTOR-04` | 오류 | 사용 여부을 선택해 주세요. | 확인 | `DLG-MOTOR` 추가/수정 |
| `MSG-MOTOR-05` | 추가 / 수정 | 해당 모터정보를 추가하시겠습니까? / 해당 모터정보를 수정하시겠습니까? | 예/아니오 | 〃 |
| `MSG-MOTOR-06` | 오류 | 이미 해당 모터 종류와 ID를 갖는 정보가 있습니다. 확인바랍니다. | 확인 | `DLG-MOTOR` 추가 |
| `MSG-MOTOR-07` | 삭제 | 해당 모터정보를 삭제하시겠습니까? | 예/아니오 | `DLG-MOTOR-BTN-DELETE` |
| `MSG-MOTOR-08` | 오류 | 해당 모터를 참조로 사용하는 라인이 있습니다. 변경 후, 삭제할 수 있습니다. | 확인 | 〃 |

> 문구의 맞춤법(예: "종류을", "아닙니다..")은 원본 그대로입니다. 수정을 원하시면 해당 `MSG-` ID 로 요청해 주세요.

---

## 4. 동작 사양

### 4.1 통신상태 점 색상 규칙 (`LINE-*-DOT-OLD/NEW/REF`, 5초마다 갱신)

| 조건 | `DOT-OLD` | `DOT-NEW` | `DOT-REF` |
|---|---|---|---|
| 해당 온실 제어 중지 상태 | 회색 `C-NET-STOP` | 회색 | 회색 |
| 제어 중 + 정상 | 파랑 `C-NET-OK` (기존모터 게이트웨이 접속됨) | 파랑 (신규모터 게이트웨이 접속됨) | 파랑 (수동 모드이거나, 참조 기존모터 상태가 '정상') |
| 제어 중 + 이상 | 빨강 `C-NET-ERR` (접속 실패/끊김) | 빨강 | 빨강 (참조 모터가 이상/대기/중지/없음) |

### 4.2 모터 상태 판정 (모든 라인, 5초마다, 제어 시작 여부와 무관)

| 조건 | 결과 상태 |
|---|---|
| 사용 = 중지 | 2 중지 |
| 사용 중 + 수신 이력 없음 | 9 대기 |
| 사용 중 + 마지막 수신 후 60초 이상 | 1 이상 |
| 사용 중 + 마지막 수신 후 60초 미만 | 0 정상 |

- 기존모터: RPM 읽기 응답을 받을 때마다 RPM·수신시각 갱신.
- 신규모터: 원본과 같이 수신시각을 갱신하지 않으므로 **항상 '대기'** 로 보입니다 (`P-NEW-STATUS` 로 변경 가능, 6장 Q2).

### 4.3 제어 흐름 (라인 1개 기준, `MAIN-BTN-H1-START` 후)

```
[기존모터 송신 쓰레드]                       [신규모터 송신 쓰레드]
 ① 미접속이면 접속(1초 제한)                  ① 미접속이면 접속(1초 제한)
    실패 → 통신이상, 10초 후 재시도              실패 → 통신이상, 10초 후 재시도
 ② 읽기 대기열이 있으면                       ② 읽기 대기열이 있으면, 항목마다
    한 건씩 전송(0.5초 간격)                     a. 쓰기(제어) 대기열 전부 전송(0.5초 간격)
                                                 b. 읽기 1건 전송(0.5초)
 ③ 없으면 '사용 중 기존모터' 마다              ③ 없으면 '사용 중 신규모터' 마다 읽기 명령 적재,
    RPM 읽기 명령 적재 후 1초 대기                 '중지로 바뀐 신규모터' 는 RPM 0 쓰기 1회 적재, 1초 대기

[기존모터 수신]  응답(11바이트, FC04) → 해당 모터 RPM, 수신시각 갱신
[신규모터 수신]  응답(23바이트, FC03) → 목표 RPM 결정
                  · 수동 모드: 수동RPM값
                  · 참조 모드: 참조 기존모터 상태가 '정상'이면 그 RPM, 아니면 이번 응답은 제어하지 않음
                 자동/수동 레지스터 ≠ 1 → '수동(1)' 쓰기 적재
                 동작/정지 레지스터 ≠ 1 → '동작(1)' 쓰기 적재
                 현재 RPM ≠ 목표 RPM   → 'RPM' 쓰기 적재
```

- 제어 종료: 대기열 비움, 소켓 닫음, 쓰레드 종료.
- 프로그램 종료: 설정 저장 후 모든 라인 제어 종료.

---

## 5. 동작 파라미터 (`motor_monitor/constants.py`)

| ID | 이름 | 상수명 | 값 | 의미 |
|---|---|---|---|---|
| `P-MAX-RPM` | 수동 RPM 최대값 | `MAX_RPM` | 1100 | `DLG-LINE-TXT-MRPM` 허용 상한 |
| `P-MOTOR-ID-RANGE` | 모터 ID 범위 | `MOTOR_ID_MIN`/`MAX` | 1~99 | `DLG-MOTOR-TXT-ID` 허용 범위 |
| `P-MOTOR-TIMEOUT` | 이상 판정 시간 | `MOTOR_TIMEOUT_SEC` | 60초 | 마지막 수신 후 경과 시 '이상' |
| `P-STATUS-INTERVAL` | 상태 갱신 주기 | `STATUS_UPDATE_INTERVAL` | 5초 | 모터 상태 판정 |
| `P-NET-INTERVAL` | 통신 점 갱신 주기 | `NETWORK_UPDATE_INTERVAL` | 5초 | `DOT` 색 갱신 |
| `P-LIST-REFRESH` | 목록 반영 주기 | `LIST_REFRESH_MS` (main_window.py) | 0.3초 | 데이터 변경 → 화면 반영 |
| `P-CONNECT-TIMEOUT` | 접속 제한시간 | `CONNECT_CHECK_TIMEOUT` | 1초 | |
| `P-RECONNECT` | 재접속 대기 | `RECONNECT_DELAY` | 10초 | 접속 실패 후 |
| `P-SEND-INTERVAL` | 명령 간격 | `SEND_INTERVAL` | 0.5초 | 명령 1건 전송 후 대기 |
| `P-CYCLE` | 적재 후 대기 | `CYCLE_INTERVAL` | 1초 | |
| `P-ERR-RETRY` | 전송오류 후 대기 | `ERROR_RETRY_DELAY` | 1초 | 소켓 닫고 재접속 |
| `P-CMD-IDX-MAX` | 트랜잭션 번호 최대 | `CMD_IDX_MAX` | 9999 | 초과 시 0 |
| `P-DEFAULT-PORT` | 기본 포트 | `DEFAULT_PORT` | 502 | |
| `P-DEFAULT-IP` | 기본 IP | `DEFAULT_IP_PREFIX` | 192.168.0.x | 라인 순서 n(0부터)에 대해 기존=2n+1, 신규=2n+2 |
| `P-H2-CONTROL` | 온실2 제어 버튼 동작 | `H2_CONTROL_ENABLED` | False | 원본: 동작 안 함 |
| `P-NEW-STATUS` | 신규모터 상태 표시 | `NEW_MOTOR_STATUS_TRACKING` | False | 원본: 신규모터 항상 '대기' |
| `P-DEBUG` | 상세 로그 대상 | `DEBUG_LINE`, `DEBUG_IDS` (house.py) | "ALL", {5, 50} | `--debug` 실행 시 패킷 로그 범위 |

---

## 6. 색상 / 스타일 (`motor_monitor/ui/theme.py`)

| ID | 상수 | 값 | 사용처 |
|---|---|---|---|
| `C-TAB-SEL` | `TAB_SELECTED` | `#16a085` | 선택된 탭 |
| `C-TAB` | `TAB_UNSELECTED` | `#d3d3d3` | 선택 안 된 탭 |
| `C-ZONE-TITLE` | `ZONE_TITLE_BG` | `#34495e` | `ZONE-*-TITLE` |
| `C-SIDE-EVEN` | `SIDE_EVEN_BG` | `#f39c12` | 짝 띠 |
| `C-SIDE-ODD` | `SIDE_ODD_BG` | `#2980b9` | 홀 띠 |
| `C-NET-OK` | `NET_OK` | `#2980b9` | 상태 점 정상 |
| `C-NET-ERR` | `NET_ERR` | `#c0392b` | 상태 점 이상 |
| `C-NET-STOP` | `NET_STOP` | `#808080` | 상태 점 중지 |
| `C-BTN-START` | `START_ICON` | `#2980b9` | 시작 버튼 삼각형 / 마우스오버 배경 |
| `C-BTN-STOP` | `STOP_ICON` | `#c0392b` | 종료 버튼 사각형 / 마우스오버 배경 |
| `C-DLG-LINE` | `LINE_DIALOG_HEADER` | `#2c3e50` | `DLG-LINE-TITLE` |
| `C-DLG-MOTOR` | `MOTOR_DIALOG_HEADER` | `#16a085` | `DLG-MOTOR-TITLE` |
| `C-BTN-NORMAL` | `BUTTON_STYLES["normal"]` | `#ecf0f1` / 오버 `#d4d4d4` | 취소 버튼 |
| `C-BTN-SUBMIT` | `BUTTON_STYLES["submit"]` | `#547791` / 오버 `#3a5e79` | 저장·추가·수정 버튼 |
| `C-BTN-ALERT` | `BUTTON_STYLES["alert"]` | `#e74c3c` / 오버 `#c0392b` | 삭제 버튼 |
| `C-LIST-SEL` | `LIST_SELECTED` | `#cce8ff` | 목록 선택 줄 |
| `C-LIST-HOVER` | `LIST_HOVER` | `#e5f3ff` | 목록 마우스오버 줄 |
| `C-STATUS-ERR` | `STATUS_ERR_FG` | `#ff0000` | "이상" 글자 |
| `C-STATUS-DIM` | `STATUS_DIM_FG` | `#d3d3d3` | "중지"/"대기" 글자 |

글꼴: 맑은 고딕(Windows) 9pt 기본. 그림 파일: `motor_monitor/assets/` (원본 images 폴더와 동일).

---

## 7. 데이터 사양

### 7.1 라인 코드 / 모터 키
- 라인 16개: `H1R1E, H1R1O, H1R2E, … H1R4O, H2R1E, … H2R4O` (온실2 8개는 화면이 없어 설정·제어 불가, 원본 동일)
- 모터 고유 키: `{종류}_{ID}` (종류 0=기존, 1=신규) 예) `0_5`, `1_10`. 같은 라인에 같은 ID 라도 종류가 다르면 별개 모터.

### 7.2 설정 파일 `Config.ini` (실행 폴더, 원본과 형식 호환)

```ini
[LINE_CONFIG]
H1R1E_SOURCE_IP=192.168.0.1      ; 기존모터 IP
H1R1E_SOURCE_PORT=502            ; 기존모터 PORT
H1R1E_TARGEIP=192.168.0.2        ; 신규모터 IP   (키 이름 오타는 원본 호환용으로 유지)
H1R1E_TARGEPORT=502              ; 신규모터 PORT
H1R1E_SOURCE_MANUAL=1            ; 1=수동, 0=기존모터 참조
H1R1E_MANUAL_RPM=0               ; 수동 RPM
H1R1E_SOURCE_LINE=H1R1E          ; 참조 라인
H1R1E_SOURCE_ID=5                ; 참조 기존모터 ID
... (16개 라인 반복)
[MOTOR_CONFIG]
H1R1E_MOTORS=10=1=1,5=0=1        ; ID=종류=사용(1/0), 쉼표로 구분, 화면 표시 순서
```
- 저장 시점: 프로그램 시작 직후(기본값 채움), 라인설정 저장, 모터 추가/수정/삭제, 프로그램 종료.
- 값 안의 `&`, `|` 문자는 읽을 때 제거(원본 동일).

### 7.3 로그
- `logs/motor_monitor.log` (실행 폴더, 2MB × 5개 순환). 접속 성공/실패, 제어 시작/종료 기록.
- `--debug` 로 실행하면 송수신 패킷(16진수)까지 기록 (`P-DEBUG` 범위).

---

## 8. 통신 사양 (Modbus TCP)

| 대상 | 명령 | 기능코드 | 주소 | 개수/값 | 응답 판정 |
|---|---|---|---|---|---|
| 기존모터 | RPM 읽기 | 0x04 Read Input Register | 0x000E (14) | 1개 | 길이 11, FC=04 → RPM = byte[9]·256 + byte[10] |
| 신규모터 | 상태 읽기 | 0x03 Read Holding Register | 0x0013 | 7개 | 길이 23, FC=03, 바이트수 14 → 자동/수동=[9,10], 동작/정지=[11,12], RPM=[15,16] |
| 신규모터 | 수동 설정 | 0x06 Write Single Register | 0x0013 | 1 (수동) | 응답 무시 |
| 신규모터 | 동작 설정 | 0x06 | 0x0014 | 1 (동작) | 응답 무시 |
| 신규모터 | RPM 설정 | 0x06 | 0x0016 | 목표 RPM | 응답 무시 |

- MBAP 헤더: 트랜잭션 번호(0~9999 순환, 기존/신규 별도) · 프로토콜 0 · 길이 · Unit ID = 모터 ID.
- 라인마다 기존모터 게이트웨이 1개, 신규모터 게이트웨이 1개에 각각 TCP 연결 1개.

---

## 9. 원본과의 차이

### 9.1 의도적으로 바꾼 점 (안전성·안정성)

| ID | 내용 | 원본 동작 | 변경 동작 | 위치 |
|---|---|---|---|---|
| D1 | 참조 기존모터가 목록에 없을 때 | 빈 정보(상태 정상, RPM 0)를 받아 **신규모터를 RPM 0 으로 제어** | 상태 이상과 같이 **제어 보류** | `house.py` `resolve_setting_rpm` |
| D2 | 사용 중인 신규모터가 없을 때 '중지' 전환 | RPM 0 명령이 대기열에만 쌓이고 전송 안 됨 | 즉시 전송 | `house.py` `_tgt_send_job` |
| D3 | TCP 수신 처리 | 수신 1회 = 응답 1개로 가정 | MBAP 길이로 응답 분리(붙어 오거나 나뉘어 와도 처리) | `modbus.py` `FrameSplitter` |
| D4 | 접속 확인 | 확인용 연결을 따로 열고 닫지 않음(연결 누수) | 1초 제한 접속 1회로 처리 | `house.py` `_Channel.connect` |
| D5 | 상대가 연결을 끊었을 때 | 다음 전송 오류까지 대기 | 즉시 재접속 절차 | `house.py` `_recv_loop` |
| D6 | 모터 설정 형식 오류 | 프로그램 오류 | 해당 항목만 무시하고 로그 기록 | `config.py` `parse_motor_list` |
| D7 | 중복 실행 확인 | 같은 이름 프로세스 검사 | 잠금 파일 검사 | `runtime.py` |
| D8 | 로그 | 디버그 창 출력만 | 로그 파일 추가, `--debug` 옵션 | `runtime.py`, `app.py` |
| D9 | 설정 저장 | Windows INI API | 임시파일 후 교체(저장 중 전원 차단 시 손상 방지), UTF-8 | `config.py` |

### 9.2 화면 기술 차이 (WPF → tkinter)
- 대화상자 둥근 모서리/투명 배경은 사각형으로 표시. Windows 에서는 원본처럼 제목표시줄 없는 창(머리글 끌어 이동), 그 외 OS 는 일반 창.
- 목록은 데이터 바인딩 대신 0.3초 주기 변경 확인으로 갱신.
- 원본의 숨겨진 시험용 버튼 3개(높이 0 행, "test")는 화면에 보이지 않으므로 옮기지 않음.

### 9.3 원본 그대로 유지한 점 (확인 필요 — 11장)

| ID | 내용 |
|---|---|
| Q1 | 온실2 탭은 비어 있고 온실2 제어 버튼은 동작하지 않음 (`P-H2-CONTROL`) |
| Q2 | 신규모터는 응답을 받아도 상태가 항상 '대기' (`P-NEW-STATUS`) |
| Q3 | '대기' 상태에는 상태 그림이 없음 |
| Q4 | 상태 정렬은 클릭 시점 기준(상태 변화 시 자동 재정렬 없음) |
| Q5 | 수동 모드 저장 시 참조 라인 미선택이면 참조 라인이 `H1` 로 저장됨(동작 영향 없음) |
| Q6 | 제어를 종료해도 상태 판정은 계속되어, 종료 60초 후 기존모터가 '이상' 으로 표시 |
| Q7 | 참조 모터 상태는 5초마다 판정하므로 제어 시작 직후 최대 5초간 참조 모드 라인은 제어 보류 |

---

## 10. 소스 구성

| 파일 | 역할 | 원본 대응 |
|---|---|---|
| `run.py` | 실행 시작점 | – |
| `motor_monitor/app.py` | 시작 절차(로그, 중복실행 확인, 설정 읽기, 메인 창) | `Application.xaml`, `MainWindow_Loaded` |
| `motor_monitor/constants.py` | 상수(5장) | `M_Variables.vb` |
| `motor_monitor/config.py` | Config.ini 읽기/쓰기 | `M_Config.vb` |
| `motor_monitor/helpers.py` | 라인명 변환, IP/정수 검사 | `M_Functions.vb` |
| `motor_monitor/modbus.py` | Modbus 프레임 생성/해석 | `C_HOUSE.vb` modbus_function |
| `motor_monitor/house.py` | 라인별 모터 관리·정렬·상태·통신 쓰레드 | `C_HOUSE.vb`, `C_LIST_BIND_Data.vb` |
| `motor_monitor/plant.py` | 전체 라인/설정 묶음, 제어 시작/종료 | `M_Variables.vb` HOUSES/MCONFIG |
| `motor_monitor/runtime.py` | 중복 실행 방지, 로그 설정 | – |
| `motor_monitor/simulator.py` | 장비 없이 시험하는 가상 게이트웨이 | – |
| `motor_monitor/ui/main_window.py` | `SCR-MAIN`, `ZONE-`, `LINE-` | `MainWindow.xaml(.vb)` |
| `motor_monitor/ui/dialogs.py` | `DLG-LINE`, `DLG-MOTOR` | `wHOUSE.xaml(.vb)`, `wMotor.xaml(.vb)` |
| `motor_monitor/ui/widgets.py` | 버튼, 상태 점, 탭, 툴팁, 모터 목록 | `Application.xaml`/`MainWindow.xaml` 템플릿 |
| `motor_monitor/ui/theme.py` | 색상·글꼴·그림(6장) | XAML 색상값 |
| `tests/` | 자동 시험 (로직 20개 + 화면 4개) | – |
| `tools/make_screenshots.py`, `tools/annotate_screenshots.py` | 문서용 화면 캡처/번호 표시 | – |

---

## 11. 미결 사항 (확인 요청)

| 번호 | 질문 | 현재 처리 |
|---|---|---|
| 1 | 온실2 화면(4개 구역)과 제어 버튼을 활성화할까요? (Q1) | 원본대로 비활성 |
| 2 | 신규모터도 응답 수신 시 RPM/정상·이상을 표시할까요? (Q2) | 원본대로 항상 '대기' |
| 3 | D1~D9 변경 중 원본 동작으로 되돌릴 항목이 있나요? | 변경 적용 |
| 4 | 메시지 문구의 맞춤법을 고칠까요? ("종류을", "여부을", "아닙니다..") | 원본 유지 |

---

## 12. 변경 이력

| 버전 | 날짜 | 변경 내용 | 관련 ID |
|---|---|---|---|
| 1.0.0 | 2026-10-13 | VB.NET 원본을 파이썬으로 최초 이식. 화면 3종, 통신/제어, 설정 파일 호환, 사용자 매뉴얼·스펙시트 작성 | 전체 |
