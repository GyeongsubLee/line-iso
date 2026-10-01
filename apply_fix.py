# -*- coding: utf-8 -*-
"""
line_iso_desktop_tool V1.34 / V1.35 -> V1.36 패치 스크립트

사용법 (line_iso_desktop_tool_V1.35.py 또는 V1.34.py 와 같은 폴더에 이 파일을 두고):
    py apply_fix.py
또는 경로 직접 지정:
    py apply_fix.py "C:\\...\\line_iso_desktop_tool_V1.35.py"

하는 일
1) line_iso_desktop_tool_V1.36.py 생성 (원본 파일은 그대로 둔다)
   - V1.35: get_extracted_dir() / get_reports_dir() 가 저장된 폴더를 만들 수 없으면
     (다른 PC의 D: 드라이브 등) 기본 폴더(output/extracted, output/reports)로 자동 전환
   - V1.36: Step 3 비교 규칙 표에서 "QTableView::setSpan: single cell span won't be added"
     경고가 콘솔에 반복 출력되던 문제 수정
   - 이미 적용된 수정은 건너뛴다
2) config/*.json 정리 (원본은 config/_backup_YYYYMMDD_HHMMSS/ 폴더에 백업)
   - 현재 PC에 존재하지 않는 절대경로를 빈 값으로 정리
   - ISO DWG Mapping의 Tracing이 NDE Ratio와 같은 텍스트(좌표)로 잘못 지정된 경우,
     Title Block 배치상 Tracing 칸 좌표(x=157.289, y=13.5)로 수정
"""

import json
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path


SRC_NAMES = ["line_iso_desktop_tool_V1.35.py", "line_iso_desktop_tool_V1.34.py"]
DST_NAME = "line_iso_desktop_tool_V1.36.py"
NEW_VERSION = "V1.36"

HEADER_NOTES = [
    "# V1.36: Stops the Step 3 rule table from printing 'single cell span won't be added' Qt warnings.\n",
    "# V1.35: Falls back to default output/extracted and output/reports folders when a saved path "
    "(e.g. another PC's drive) cannot be created, so the program starts on a new PC.\n",
]

REPLACEMENTS = [
    (
        "get_extracted_dir 폴더 생성 실패 시 기본 폴더 사용",
        '        config_text = normalize_text(self.config.get("paths", {}).get("extracted_dir", ""))\n'
        '        path = Path(edit_text or config_text or str(EXTRACTED_DIR))\n'
        '        path.mkdir(parents=True, exist_ok=True)\n'
        '        return path\n',
        '        config_text = normalize_text(self.config.get("paths", {}).get("extracted_dir", ""))\n'
        '        path = Path(edit_text or config_text or str(EXTRACTED_DIR))\n'
        '\n'
        '        try:\n'
        '            path.mkdir(parents=True, exist_ok=True)\n'
        '        except OSError:\n'
        '            # V1.35: 다른 PC에서 저장한 경로(없는 드라이브 등)는 기본 폴더로 되돌린다.\n'
        '            path = EXTRACTED_DIR\n'
        '            path.mkdir(parents=True, exist_ok=True)\n'
        '            self.config.setdefault("paths", {})["extracted_dir"] = str(path)\n'
        '            self.sync_extracted_dir_widgets(path)\n'
        '\n'
        '        return path\n',
    ),
    (
        "get_reports_dir 폴더 생성 실패 시 기본 폴더 사용",
        '    p = Path(path_text) if path_text else REPORTS_DIR\n'
        '    p.mkdir(parents=True, exist_ok=True)\n'
        '    return p\n',
        '    p = Path(path_text) if path_text else REPORTS_DIR\n'
        '    try:\n'
        '        p.mkdir(parents=True, exist_ok=True)\n'
        '    except OSError:\n'
        '        # V1.35: 다른 PC에서 저장한 경로(없는 드라이브 등)는 기본 폴더로 되돌린다.\n'
        '        p = REPORTS_DIR\n'
        '        p.mkdir(parents=True, exist_ok=True)\n'
        '        self.config.setdefault("paths", {})["reports_dir"] = str(p)\n'
        '    return p\n',
    ),
    (
        "비교 규칙 표 setSpan 경고 제거",
        '        # 일단 span 초기화\n'
        '        self.compare_rules_table.setSpan(row, 2, 1, 1)\n'
        '        self.compare_rules_table.setSpan(row, 3, 1, 1)\n',
        '        # 일단 span 초기화. 1x1 setSpan은 기존 span이 있을 때만 해제되고,\n'
        '        # 없으면 Qt가 "single cell span won\'t be added" 경고를 출력하므로 span이 있을 때만 호출한다.\n'
        '        if self.compare_rules_table.columnSpan(row, 2) > 1 or self.compare_rules_table.rowSpan(row, 2) > 1:\n'
        '            self.compare_rules_table.setSpan(row, 2, 1, 1)\n',
    ),
]

WINDOWS_ABS_PATH = re.compile(r"^[A-Za-z]:[\\/]")


def patch_source(src: Path) -> Path:
    raw = src.read_bytes()
    has_bom = raw.startswith(b"\xef\xbb\xbf")
    crlf = b"\r\n" in raw
    text = raw.decode("utf-8-sig").replace("\r\n", "\n")

    for name, old, new in REPLACEMENTS:
        if new in text:
            print(f"[건너뜀] 이미 적용됨: {name}")
            continue
        count = text.count(old)
        if count != 1:
            raise SystemExit(
                f"[중단] '{name}' 수정 위치를 찾지 못했습니다 (발견 {count}회).\n"
                f"원본이 V1.34/V1.35 그대로인지 확인하세요: {src}"
            )
        text = text.replace(old, new)
        print(f"[적용] {name}")

    text, n = re.subn(r'^APP_VERSION = "V1\.3\d"', f'APP_VERSION = "{NEW_VERSION}"', text, count=1, flags=re.M)
    if n != 1:
        raise SystemExit("[중단] APP_VERSION 위치를 찾지 못했습니다.")

    lines = text.split("\n", 1)
    if not lines[0].startswith("# LineIsoCompare / line_iso_desktop_tool V1.3"):
        raise SystemExit("[중단] 첫 줄 버전 표기를 찾지 못했습니다.")
    notes = "".join(note for note in HEADER_NOTES if note not in text)
    text = f"# LineIsoCompare / line_iso_desktop_tool {NEW_VERSION}\n" + notes + lines[1]

    if crlf:
        text = text.replace("\n", "\r\n")
    data = text.encode("utf-8")
    if has_bom:
        data = b"\xef\xbb\xbf" + data

    dst = src.with_name(DST_NAME)
    dst.write_bytes(data)
    return dst


def _clean_paths(value, changes: list, key_path: str):
    """현재 PC에 존재하지 않는 Windows 절대경로 문자열을 빈 값으로 바꾼다."""
    if isinstance(value, dict):
        return {k: _clean_paths(v, changes, f"{key_path}.{k}" if key_path else k) for k, v in value.items()}
    if isinstance(value, list):
        return [_clean_paths(v, changes, f"{key_path}[{i}]") for i, v in enumerate(value)]
    if isinstance(value, str) and WINDOWS_ABS_PATH.match(value.strip()):
        try:
            exists = Path(value.strip()).exists()
        except OSError:
            exists = False
        if not exists:
            changes.append((key_path, value))
            return ""
    return value


_NDE_TRACING_POINT = (104.289, 13.5)
_TRACING_POINT = (157.289, 13.5)


def _fix_iso_tracing(iso_mapping, changes: list, key_path: str):
    """Tracing이 NDE Ratio와 같은 텍스트로 지정된 기존 Mapping을 Tracing 칸 좌표로 바로잡는다."""
    if not isinstance(iso_mapping, dict):
        return
    fields = iso_mapping.get("fields")
    if not isinstance(fields, dict):
        return
    tracing, nde = fields.get("tracing"), fields.get("nde_ratio")
    if not isinstance(tracing, dict) or not isinstance(nde, dict):
        return

    def point(info):
        p = info.get("insertion_point") or {}
        try:
            return (round(float(p.get("x")), 3), round(float(p.get("y")), 3))
        except (TypeError, ValueError):
            return None

    if point(tracing) == point(nde) == _NDE_TRACING_POINT:
        tracing["insertion_point"] = {"x": _TRACING_POINT[0], "y": _TRACING_POINT[1], "z": 0.0}
        tracing["handle"] = ""
        tracing["sample_text"] = ""
        changes.append((f"{key_path}fields.tracing", "NDE Ratio 좌표(104.289, 13.5) -> Tracing 좌표(157.289, 13.5)"))


def clean_config(config_dir: Path):
    if not config_dir.is_dir():
        print(f"[config] 폴더가 없어 건너뜀: {config_dir}")
        return

    backup_dir = config_dir / f"_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    for json_path in sorted(config_dir.glob("*.json")):
        try:
            with open(json_path, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
        except Exception as e:
            print(f"[config] 읽기 실패, 건너뜀: {json_path.name} ({e})")
            continue

        changes = []
        cleaned = _clean_paths(data, changes, "")
        if json_path.name == "iso_dwg_mapping.json":
            _fix_iso_tracing(cleaned, changes, "")
        elif json_path.name in ("project_settings.json", "project_mapping.json") and isinstance(cleaned, dict):
            _fix_iso_tracing(cleaned.get("iso_dwg"), changes, "iso_dwg.")
        if not changes:
            continue

        backup_dir.mkdir(exist_ok=True)
        shutil.copy2(json_path, backup_dir / json_path.name)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(cleaned, f, ensure_ascii=False, indent=2)

        print(f"[config] {json_path.name}: {len(changes)}개 항목 정리")
        for key, old in changes:
            print(f"    - {key}: {old}")

    if backup_dir.exists():
        print(f"[config] 원본 백업: {backup_dir}")


def main():
    if len(sys.argv) > 1:
        src = Path(sys.argv[1]).resolve()
    else:
        here = Path(__file__).resolve().parent
        candidates = [d / name for name in SRC_NAMES for d in (here, Path.cwd())]
        src = next((c for c in candidates if c.exists()), candidates[0])

    if not src.exists():
        raise SystemExit(f"[중단] {' / '.join(SRC_NAMES)} 파일을 찾지 못했습니다. 경로를 인자로 지정하세요.")

    dst = patch_source(src)
    print(f"[완료] {dst.name} 생성: {dst}")
    clean_config(src.parent / "config")
    print(f"\n이제 {dst.name} 를 실행하세요.")


if __name__ == "__main__":
    main()
