# line-iso

Line List ↔ Source Data QC Tool (`line_iso_desktop_tool`) 유지보수용 저장소.

## 수정 내역

### V1.38
- 3단계에 **"Source(ISO)에 Line No. 칸이 없음"** 체크 추가. 체크하면 Line List의 Line No.가
  ISO DWG No. 안에 들어 있을 때 같은 Line으로 짝을 지음 (Report 기준 3가지 모두 적용)
- 짝을 찾는 순서
  1. Source 값이 Line List Line No.와 정확히 같으면 그대로 사용 (`EXACT`)
  2. `-` `.` `/` `_` 경계로 딱 떨어지는 부분이 Line List Line No.와 같으면 사용 (`ISO_NO_CONTAINS`)
  3. 2가 없으면 경계와 상관없이 들어 있는 6자 이상 Line No.를 사용 (`ISO_NO_CONTAINS_LOOSE`)
  - 2, 3에서 여러 개가 걸리면 가장 긴 값을 쓰고, 같은 길이로 2개 이상이면 짝을 짓지 않음 (`AMBIGUOUS`)
- 행별로 어떤 방식으로 짝지어졌는지는 Extracted 폴더의 `*_line_no_resolved.xlsx`에 있는 `line_no_match` 열에서 확인

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

1. `apply_fix.py`를 지금 쓰는 `line_iso_desktop_tool_V1.3x.py`(V1.34~V1.37)와 같은 폴더에 복사
2. 그 폴더에서 실행:
   ```
   py apply_fix.py
   ```
   - `line_iso_desktop_tool_V1.38.py`가 새로 생성됨 (기존 파일은 그대로 유지, 이미 적용된 수정은 건너뜀)
   - `config/*.json` 정리 (원본은 `config/_backup_날짜시간/`에 백업)
3. `line_iso_desktop_tool_V1.38.py` 실행

`config/project_settings.json`은 위 정리를 적용한 버전이다.

실행 후 Line List Excel, ISO DWG 폴더, MDB 파일 등은 각 화면에서 현재 PC 경로로 다시 선택한다.
