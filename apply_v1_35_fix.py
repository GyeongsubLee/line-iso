# -*- coding: utf-8 -*-
"""
line_iso_desktop_tool V1.34 -> V1.35 패치 스크립트

사용법 (line_iso_desktop_tool_V1.34.py 와 같은 폴더에 이 파일을 두고):
    py apply_v1_35_fix.py
또는 경로 직접 지정:
    py apply_v1_35_fix.py "C:\\...\\line_iso_desktop_tool_V1.34.py"

하는 일
1) line_iso_desktop_tool_V1.35.py 생성 (원본 V1.34 파일은 그대로 둔다)
   - get_extracted_dir(): 저장된 Extracted 폴더를 만들 수 없으면(다른 PC의 D: 드라이브 등)
     기본 폴더(output/extracted)로 자동 전환
   - get_reports_dir(): 같은 방식으로 기본 폴더(output/reports)로 자동 전환
   - APP_VERSION을 V1.35로 변경
2) config/*.json 에 남아 있는 "현재 PC에 존재하지 않는 절대경로"를 빈 값으로 정리
   - 원본은 config/_backup_YYYYMMDD_HHMMSS/ 폴더에 백업
   - Mapping(열/좌표/Code 등) 정보는 그대로 유지
"""

import json
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path


SRC_NAME = "line_iso_desktop_tool_V1.34.py"
DST_NAME = "line_iso_desktop_tool_V1.35.py"

REPLACEMENTS = [
    (
        "버전 표기",
        '# LineIsoCompare / line_iso_desktop_tool V1.34\n',
        '# LineIsoCompare / line_iso_desktop_tool V1.35\n'
        '# V1.35: Falls back to default output/extracted and output/reports folders when a saved path '
        '(e.g. another PC\'s drive) cannot be created, so the program starts on a new PC.\n',
    ),
    (
        "APP_VERSION",
        'APP_VERSION = "V1.34"',
        'APP_VERSION = "V1.35"',
    ),
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
]

WINDOWS_ABS_PATH = re.compile(r"^[A-Za-z]:[\\/]")


def patch_source(src: Path) -> Path:
    raw = src.read_bytes()
    has_bom = raw.startswith(b"\xef\xbb\xbf")
    crlf = b"\r\n" in raw
    text = raw.decode("utf-8-sig").replace("\r\n", "\n")

    for name, old, new in REPLACEMENTS:
        count = text.count(old)
        if count != 1:
            raise SystemExit(
                f"[중단] '{name}' 수정 위치를 찾지 못했습니다 (발견 {count}회).\n"
                f"원본이 V1.34 그대로인지 확인하세요: {src}"
            )
        text = text.replace(old, new)

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
        if not changes:
            continue

        backup_dir.mkdir(exist_ok=True)
        shutil.copy2(json_path, backup_dir / json_path.name)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(cleaned, f, ensure_ascii=False, indent=2)

        print(f"[config] {json_path.name}: 존재하지 않는 경로 {len(changes)}개 정리")
        for key, old in changes:
            print(f"    - {key}: {old}")

    if backup_dir.exists():
        print(f"[config] 원본 백업: {backup_dir}")


def main():
    if len(sys.argv) > 1:
        src = Path(sys.argv[1]).resolve()
    else:
        here = Path(__file__).resolve().parent
        src = here / SRC_NAME
        if not src.exists():
            src = Path.cwd() / SRC_NAME

    if not src.exists():
        raise SystemExit(f"[중단] {SRC_NAME} 파일을 찾지 못했습니다. 경로를 인자로 지정하세요.")

    dst = patch_source(src)
    print(f"[완료] {dst.name} 생성: {dst}")
    clean_config(src.parent / "config")
    print(f"\n이제 {dst.name} 를 실행하세요.")


if __name__ == "__main__":
    main()
