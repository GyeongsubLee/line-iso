# 3D BIM BM → DMCS Display Format Tool

3D BIM BM 파일을 병합·정제하고, PCWBS 비교 → Display Format 매핑 →
Temperature & Painting → 파일 분할까지 처리하는 Windows GUI 프로그램입니다.

## 실행 (Python)

```
py -m pip install -r requirements.txt
py BM_DMCS_Tool_v8_13_Beta.py
```

## EXE 만들기 (배포용)

### 방법 1. 내 PC에서 빌드
1. Windows에 Python 3.9 이상 설치 (`py` 명령 사용 가능해야 함)
2. 이 폴더에서 `build_exe.bat` 더블클릭
3. `dist\BM_DMCS_Tool_v8_13_Beta.exe` 생성

### 방법 2. GitHub에서 자동 빌드
1. GitHub 저장소 → **Actions** → **Build Windows EXE**
2. **Run workflow** 실행 (또는 프로그램 파일을 push하면 자동 실행)
3. 완료된 실행 화면 하단 **Artifacts**에서 `BM_DMCS_Tool_v8_13_Beta-exe` 다운로드

### 배포 시 참고
- EXE 파일 하나만 전달하면 됩니다. (Python 설치 불필요)
- 처음 실행하면 EXE와 같은 폴더에 `02_test_input`, `03_reference`, `04_output`, `05_logs` 폴더가 만들어집니다.
- 서명되지 않은 EXE라 Windows SmartScreen 경고가 뜰 수 있습니다. **추가 정보 → 실행**을 누르면 됩니다.
  사내 보안 프로그램이 막는 경우 IT 부서에 예외 등록을 요청하세요.
