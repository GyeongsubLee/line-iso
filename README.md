# line-iso

Line List ↔ Source Data QC Tool (`line_iso_desktop_tool`) 유지보수용 저장소.

## V1.35 수정 사항 (다른 PC에서 실행 시 시작 오류 수정)

V1.34는 `config/project_settings.json`에 이전 담당자 PC의 경로(`D:\00. BCC Local Folder\...`)가
남아 있으면, 해당 드라이브가 없는 PC에서 Extracted 폴더를 만들다가 `FileNotFoundError`로 실행 직후 종료됐다.

- `get_extracted_dir()` / `get_reports_dir()`: 저장된 폴더를 만들 수 없으면 기본 폴더
  (`.py` 옆 `output/extracted`, `output/reports`)로 자동 전환
- `config/project_settings.json`: 현재 PC에 없는 `D:` 경로를 모두 빈 값으로 정리
  (Line List 열 매핑, ISO DWG 좌표 매핑, Core Console 설정 등은 그대로 유지)

## 적용 방법

1. `apply_v1_35_fix.py`를 `line_iso_desktop_tool_V1.34.py`와 같은 폴더에 복사
2. 그 폴더에서 실행:
   ```
   py apply_v1_35_fix.py
   ```
   - `line_iso_desktop_tool_V1.35.py`가 새로 생성됨 (V1.34 원본은 그대로 유지)
   - `config/*.json` 안의 현재 PC에 존재하지 않는 경로를 빈 값으로 정리
     (원본은 `config/_backup_날짜시간/`에 백업)
3. `line_iso_desktop_tool_V1.35.py` 실행

`config/project_settings.json`은 경로를 정리한 버전이며, 위 스크립트 대신 이 파일로 직접 교체해도 된다.

실행 후 Line List Excel, ISO DWG 폴더, MDB 파일 등은 각 화면에서 현재 PC 경로로 다시 선택한다.
