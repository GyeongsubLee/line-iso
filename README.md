# line-iso

Line List ↔ Source Data QC Tool (`line_iso_desktop_tool`) 유지보수용 저장소.

## 수정 내역

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

1. `apply_fix.py`를 `line_iso_desktop_tool_V1.35.py`(또는 V1.34)와 같은 폴더에 복사
2. 그 폴더에서 실행:
   ```
   py apply_fix.py
   ```
   - `line_iso_desktop_tool_V1.36.py`가 새로 생성됨 (기존 파일은 그대로 유지, 이미 적용된 수정은 건너뜀)
   - `config/*.json` 정리 (원본은 `config/_backup_날짜시간/`에 백업)
3. `line_iso_desktop_tool_V1.36.py` 실행

`config/project_settings.json`은 위 정리를 적용한 버전이다.

실행 후 Line List Excel, ISO DWG 폴더, MDB 파일 등은 각 화면에서 현재 PC 경로로 다시 선택한다.
