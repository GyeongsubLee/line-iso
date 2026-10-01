# line-iso

Line List ↔ Source Data QC Tool (`line_iso_desktop_tool`) 유지보수용 저장소.

## 수정 내역

### V1.40
- 사이드바 개발자 정보 변경: `Developed by 김영철, 이경섭` / `02-369-5533`

### V1.39
- 3단계에 **ISO DWG ↔ Line No. 매칭표** 추가: ISO DWG에 Line No.가 없을 때, 사용자가 만든 Excel
  (ISO DWG No. / Line No.)로 Source의 Line No.를 채운 뒤 비교함 (Report 기준 3가지 모두 적용)
  - **매칭표 양식 만들기**: 추출된 ISO DWG No. 목록으로 빈 Excel을 만들어 줌 → Line No. 열만 채우면 됨
  - 머리글에 ISO/DWG가 들어간 열을 ISO DWG No., LINE이 들어간 열을 Line No.로 인식
    (못 찾으면 A열 = ISO DWG No., B열 = Line No.). 대소문자·공백 차이는 무시
  - 매칭표에 없는 ISO DWG는 추출된 Line No.를 그대로 사용
  - 행별 적용 결과는 Extracted 폴더의 `*_with_line_no_table.xlsx`에 있는 `line_no_from` 열에서 확인
- V1.38의 "ISO DWG No. 포함 매칭"은 복잡해서 제거함 (V1.38 파일에 패치하면 자동으로 빠짐)

### V1.37
- 2단계 AutoCAD 추출의 **모든 매핑 항목**에서 추출방식을 고를 수 있도록 함:
  `값 그대로` / `구분자로 나누기` / `정규식 추출`. 결과는 오른쪽 `→`에 바로 미리 보여 줌
- ISO DWG에 Line No. 칸이 따로 없고 ISO DWG No. 안에만 Line No.가 들어 있는 경우,
  Line No.만 잘라서 비교 Key로 쓸 수 있음
  - 예: `AGCC.1917-3010-30100109HE-TK10.ISO-0002` → 구분자 `-`, 3번째 → `30100109HE`
  - 또는 정규식 `\d{8}[A-Z]+` → `30100109HE`
- AutoCAD에 연결하지 않은 상태에서도 저장된 Mapping의 추출방식만 바꿔서 저장하거나 추출할 수 있음

### V1.36
- Step 3 비교 규칙 표를 만들 때 콘솔에 `QTableView::setSpan: single cell span won't be added`가
  반복 출력되던 문제 수정 (기능에는 영향 없던 경고)
- ISO DWG Mapping의 **Tracing**이 NDE Ratio와 같은 텍스트(좌표 104.289, 13.5 / 값 "2")로 잘못
  지정되어 있던 것을 Title Block 배치상 Tracing 칸(157.289, 13.5)으로 수정

### V1.35
V1.34는 `config/project_settings.json`에 이전 담당자 PC의 경로(`D:\00. BCC Local Folder\...`)가
남아 있으면, 해당 드라이브가 없는 PC에서 Extracted 폴더를 만들다가 `FileNotFoundError`로 실행 직후 종료됐다.

- `get_extracted_dir()` / `get_reports_dir()`: 저장된 폴더를 만들 수 없으면 기본 폴더
  (`.py` 옆 `output/extracted`, `output/reports`)로 자동 전환
- `config/*.json`: 현재 PC에 없는 경로를 빈 값으로 정리 (Mapping 정보는 그대로 유지)

## 적용 방법

1. `apply_fix.py`를 지금 쓰는 `line_iso_desktop_tool_V1.3x.py`(V1.34~V1.39)와 같은 폴더에 복사
2. 그 폴더에서 실행:
   ```
   py apply_fix.py
   ```
   - `line_iso_desktop_tool_V1.40.py`가 새로 생성됨 (기존 파일은 그대로 유지, 이미 적용된 수정은 건너뜀)
   - `config/*.json` 정리 (원본은 `config/_backup_날짜시간/`에 백업)
3. `line_iso_desktop_tool_V1.40.py` 실행

`config/project_settings.json`은 위 정리를 적용한 버전이다.

실행 후 Line List Excel, ISO DWG 폴더, MDB 파일 등은 각 화면에서 현재 PC 경로로 다시 선택한다.

## 배포용 exe 만들기

1. 한 폴더에 `line_iso_desktop_tool_V1.40.py`, `build_exe.bat`, `make_icon.py`, (공유할) `config` 폴더를 둔다
2. `build_exe.bat` 더블클릭
   - 빌드 전용 가상환경(`.build_venv`)에 필요한 패키지만 설치해서 빌드 → 용량이 작고 깔끔함
   - 폴더 안에서 가장 높은 버전의 `line_iso_desktop_tool_V*.py`를 자동으로 사용
   - 프로그램에 들어 있는 아이콘을 exe 아이콘으로 사용
   - `config` 폴더의 공용 설정(매핑, 비교 항목, 비교 규칙)은 함께 넣고,
     개인 경로가 들어 있는 `project_settings.json`, `project_mapping.json`은 제외
3. 결과: `release\LineIsoQC_V1.40.zip` ← 이 파일을 배포
   - 받는 사람은 압축을 풀고 `LineIsoQC_V1.40.exe` 실행 (Python 설치 불필요)
   - `input`, `output`, `config` 폴더는 exe 옆에 자동으로 생성됨
