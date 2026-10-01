# -*- coding: utf-8 -*-
"""
line_iso_desktop_tool V1.34 ~ V1.37 -> V1.38 패치 스크립트

사용법 (지금 쓰는 line_iso_desktop_tool_V1.3x.py 와 같은 폴더에 이 파일을 두고):
    py apply_fix.py
또는 경로 직접 지정:
    py apply_fix.py "C:\\...\\line_iso_desktop_tool_V1.37.py"

하는 일
1) line_iso_desktop_tool_V1.38.py 생성 (원본 파일은 그대로 둔다)
   - V1.35: get_extracted_dir() / get_reports_dir() 가 저장된 폴더를 만들 수 없으면
     (다른 PC의 D: 드라이브 등) 기본 폴더(output/extracted, output/reports)로 자동 전환
   - V1.36: Step 3 비교 규칙 표에서 "QTableView::setSpan: single cell span won't be added"
     경고가 콘솔에 반복 출력되던 문제 수정
   - V1.37: AutoCAD Mapping의 모든 항목에 추출방식(값 그대로 / 구분자로 나누기 / 정규식 추출)과
     결과 미리보기 추가. ISO DWG No. 안에 들어 있는 Line No.만 잘라 비교 Key로 사용할 수 있다.
   - V1.38: Step 3에 "Source(ISO)에 Line No. 칸이 없음" 체크 추가. 체크하면 Line List의 Line No.가
     ISO DWG No. 안에 포함된 경우 같은 Line으로 매칭한다.
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


SRC_NAMES = [
    "line_iso_desktop_tool_V1.37.py",
    "line_iso_desktop_tool_V1.36.py",
    "line_iso_desktop_tool_V1.35.py",
    "line_iso_desktop_tool_V1.34.py",
]
DST_NAME = "line_iso_desktop_tool_V1.38.py"
NEW_VERSION = "V1.38"

HEADER_NOTES = [
    "# V1.38: Adds a Step 3 option to match Line List Line No. contained in the ISO DWG No. "
    "when the source has no Line No. field.\n",
    "# V1.37: Adds split/regex extraction with live preview to every AutoCAD mapping field, "
    "so Line No. can be cut out of the ISO DWG No. text.\n",
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

# 실행부(# 9. 실행) 바로 앞에 끼워 넣는 코드 블록. 기존 프로그램과 같은 monkey-patch 방식이다.
RUN_SECTION_ANCHOR = "# =========================================================\n# 9. 실행\n"
V137_MARKER = "# 8Q. V1.37 AutoCAD 추출값 Customizing"
V138_MARKER = "# 8R. V1.38 Source에 Line No. 칸이 없는 경우"
V137_BLOCK = r'''# =========================================================
# 8Q. V1.37 AutoCAD 추출값 Customizing: 구분자로 나누기 / 정규식 추출
# =========================================================
# ISO DWG에 Line No. 칸이 따로 없고 ISO DWG No.(예: AGCC.1917-3010-30100109HE-TK10.ISO-0002)
# 안에만 Line No.(30100109HE)가 들어 있는 프로젝트를 위해, AutoCAD Mapping의 모든 항목에
# 추출방식(값 그대로 / 구분자로 나누기 / 정규식 추출)과 결과 미리보기를 제공한다.
# Line No.는 Step 3 비교 Key이므로 비교 규칙(동등/포함)으로는 보정되지 않고,
# 반드시 추출 단계에서 Line List와 같은 형태로 만들어야 한다.

_V37_GROUP_FIELDS = ["trcu_group", "trcu_category", "gost_group", "gost_category"]
_V37_PREV_APPLY_CAD_FIELD_TRANSFORM = apply_cad_field_transform
_V37_PREV_MAKE_CAD_FIELD_MAPPING = make_cad_field_mapping
_V37_PREV_EXTRACT_ACTIVE_DWG = LineIsoDesktopTool.extract_active_dwg
_V37_PREV_EXTRACT_DWG_FOLDER = LineIsoDesktopTool.extract_dwg_folder_core_console


def _v37_regex_extract(text: str, pattern: str) -> str:
    # 괄호 그룹이 있으면 첫 번째 그룹, 없으면 일치한 전체 문자열을 사용한다.
    if not pattern:
        return text
    try:
        m = re.search(pattern, text)
    except re.error:
        return ""
    if not m:
        return ""
    return normalize_text(m.group(1) if m.re.groups else m.group(0))


def _v37_apply_cad_field_transform(raw_value, info: dict) -> str:
    text = normalize_text(raw_value)
    if not text:
        return ""
    info = ensure_dict(info)
    if normalize_text(info.get("transform", "direct")).lower() == "regex":
        return _v37_regex_extract(text, normalize_text(info.get("pattern", "")))
    return _V37_PREV_APPLY_CAD_FIELD_TRANSFORM(text, info)


def _v37_make_cad_field_mapping(field_key: str, row: dict, search_radius: float, transform: str = "direct", delimiter: str = ",", part_index: int = 1, pattern: str = "") -> dict:
    is_regex = normalize_text(transform).lower() == "regex"
    mapping = _V37_PREV_MAKE_CAD_FIELD_MAPPING(
        field_key,
        row,
        search_radius,
        transform="direct" if is_regex else transform,
        delimiter=delimiter,
        part_index=part_index,
    )
    if is_regex:
        mapping["transform"] = "regex"
    mapping["pattern"] = normalize_text(pattern)
    return mapping


def _v37_build_cad_mapping_combos(self):
    self.clear_cad_mapping_layout()
    previous = self.config.get("iso_dwg", {}).get("fields", {})
    previous = previous if isinstance(previous, dict) else {}
    candidate_rows = [row for row in self.cad_candidates if normalize_text(row.get("display_text", ""))]
    has_candidates = bool(candidate_rows)

    help_label = QLabel(
        "추출방식으로 선택한 텍스트의 일부만 사용할 수 있습니다. "
        "예: ISO DWG No. 'AGCC.1917-3010-30100109HE-TK10.ISO-0002'에서 Line No. '30100109HE'만 쓰려면 "
        "Line No.에 같은 ISO DWG No. 텍스트를 선택하고 [구분자로 나누기] 구분자 '-' / 3번째 항목, "
        "또는 [정규식 추출] \\d{8}[A-Z]+ 를 지정하세요. 오른쪽 → 에 결과가 미리 표시됩니다."
    )
    help_label.setWordWrap(True)
    help_label.setStyleSheet("color: #555;")
    self.cad_mapping_layout.addRow(help_label)

    for field_key in self.get_active_source_fields(include_iso_no=True):
        prev_info = previous.get(field_key) if isinstance(previous.get(field_key), dict) else {}
        previous_sample = normalize_text(prev_info.get("sample_text", ""))

        combo = QComboBox()
        if has_candidates:
            combo.addItem("[미지정]", -1)
            for idx, row in enumerate(candidate_rows):
                combo.addItem(make_candidate_label(row), idx)
            if previous_sample:
                for i in range(1, combo.count()):
                    if candidate_rows[combo.itemData(i)].get("display_text") == previous_sample:
                        combo.setCurrentIndex(i)
                        break
        else:
            if previous_sample:
                combo.addItem(f"[저장값: {previous_sample}] DWG 연결/텍스트 선택 후 재매핑 가능", -1)
            else:
                combo.addItem("[DWG 연결 후 텍스트 선택하면 항목 선택 가능]", -1)
            combo.setEnabled(False)
        combo.setMinimumWidth(260)

        is_group_field = field_key in _V37_GROUP_FIELDS
        default_delimiter = "," if is_group_field else "-"
        default_part = 2 if is_group_field and field_key.endswith("category") else 1
        prev_transform = normalize_text(prev_info.get("transform", "direct")).lower()

        transform_combo = QComboBox()
        transform_combo.addItem("값 그대로", "direct")
        transform_combo.addItem("구분자로 나누기", "split")
        transform_combo.addItem("정규식 추출", "regex")
        t_idx = transform_combo.findData(prev_transform)
        transform_combo.setCurrentIndex(t_idx if t_idx >= 0 else 0)

        # 예전 Mapping은 '값 그대로'여도 구분자 ','가 저장되어 있으므로, 그대로 방식이면 항목별 기본값을 보여준다.
        delimiter_text = normalize_text(prev_info.get("delimiter", "")) if prev_transform == "split" else ""
        delimiter_edit = QLineEdit(delimiter_text or default_delimiter)
        delimiter_edit.setMaximumWidth(50)

        part_spin = QSpinBox()
        part_spin.setRange(1, 50)
        try:
            part_spin.setValue(int(prev_info.get("part_index", default_part)) if prev_transform == "split" else default_part)
        except Exception:
            part_spin.setValue(default_part)
        part_spin.setMaximumWidth(60)

        pattern_edit = QLineEdit(normalize_text(prev_info.get("pattern", "")))
        pattern_edit.setPlaceholderText(r"예: \d{8}[A-Z]+")
        pattern_edit.setMaximumWidth(160)

        preview_label = QLabel("")
        preview_label.setMinimumWidth(160)
        preview_label.setStyleSheet("color: #2D74B8; font-weight: bold;")

        def update_controls(*args, cb=combo, rows=candidate_rows, saved=previous_sample,
                            tc=transform_combo, de=delimiter_edit, ps=part_spin, pe=pattern_edit, pl=preview_label,
                            dd=default_delimiter):
            transform = tc.currentData() or "direct"
            de.setEnabled(transform == "split")
            ps.setEnabled(transform == "split")
            pe.setEnabled(transform == "regex")

            data_idx = cb.currentData()
            if rows and data_idx is not None and data_idx != -1:
                sample = normalize_text(rows[int(data_idx)].get("display_text", ""))
            else:
                sample = saved
            if not sample:
                pl.setText("")
                return

            if transform == "regex":
                try:
                    re.compile(normalize_text(pe.text()))
                except re.error:
                    pl.setText("→ 정규식 오류")
                    return
            result = _v37_apply_cad_field_transform(sample, {
                "transform": transform,
                "delimiter": normalize_text(de.text()) or dd,
                "part_index": ps.value(),
                "pattern": normalize_text(pe.text()),
            })
            pl.setText(f"→ {result}" if result else "→ (결과 없음)")

        combo.currentIndexChanged.connect(update_controls)
        transform_combo.currentIndexChanged.connect(update_controls)
        delimiter_edit.textChanged.connect(update_controls)
        part_spin.valueChanged.connect(update_controls)
        pattern_edit.textChanged.connect(update_controls)
        update_controls()

        option_row = QHBoxLayout()
        option_row.addWidget(combo, 1)
        option_row.addWidget(QLabel("추출방식:"))
        option_row.addWidget(transform_combo)
        option_row.addWidget(QLabel("구분자:"))
        option_row.addWidget(delimiter_edit)
        option_row.addWidget(part_spin)
        option_row.addWidget(QLabel("번째"))
        option_row.addWidget(QLabel("정규식:"))
        option_row.addWidget(pattern_edit)
        option_row.addWidget(preview_label)

        self.cad_field_combos[field_key] = (combo, candidate_rows, transform_combo, delimiter_edit, part_spin, pattern_edit)
        self.cad_mapping_layout.addRow(STANDARD_FIELDS.get(field_key, field_key), option_row)


def _v37_read_cad_transform(pack) -> tuple:
    transform_combo = pack[2] if len(pack) > 2 else None
    delimiter_edit = pack[3] if len(pack) > 3 else None
    part_spin = pack[4] if len(pack) > 4 else None
    pattern_edit = pack[5] if len(pack) > 5 else None
    transform = (transform_combo.currentData() if transform_combo is not None else "direct") or "direct"
    delimiter = (normalize_text(delimiter_edit.text()) if delimiter_edit is not None else "") or ","
    part_index = part_spin.value() if part_spin is not None else 1
    pattern = normalize_text(pattern_edit.text()) if pattern_edit is not None else ""
    return transform, delimiter, part_index, pattern


def _v37_refresh_cad_mapping_from_ui(self, persist: bool = True) -> bool:
    if not self.cad_field_combos:
        return False

    search_radius = float(self.search_radius_spin.value())
    fields = {}

    if self.cad_candidates:
        for field_key, pack in self.cad_field_combos.items():
            combo, candidate_rows = pack[0], pack[1]
            idx = combo.currentData()
            if idx is None or idx == -1:
                continue
            transform, delimiter, part_index, pattern = _v37_read_cad_transform(pack)
            fields[field_key] = make_cad_field_mapping(
                field_key=field_key,
                row=candidate_rows[int(idx)],
                search_radius=search_radius,
                transform=transform,
                delimiter=delimiter,
                part_index=part_index,
                pattern=pattern,
            )
    else:
        # AutoCAD에 연결하지 않은 상태: 저장된 텍스트 위치는 그대로 두고 추출방식만 갱신한다.
        saved = self.config.get("iso_dwg", {}).get("fields", {})
        if not isinstance(saved, dict) or not saved:
            return False
        for field_key, info in saved.items():
            info = dict(info) if isinstance(info, dict) else {}
            pack = self.cad_field_combos.get(field_key)
            if pack is not None:
                transform, delimiter, part_index, pattern = _v37_read_cad_transform(pack)
                info.update({
                    "transform": transform,
                    "delimiter": delimiter,
                    "part_index": part_index,
                    "pattern": pattern,
                    "search_radius": search_radius,
                })
            fields[field_key] = info

    self.config["iso_dwg"] = {
        "mapping_method": "cad_user_selection",
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "search_radius": search_radius,
        "fields": fields,
    }

    if persist:
        self.save_path_config()
        save_iso_dwg_mapping_file(self.config["iso_dwg"])
        save_config(self.config)

    return True


def _v37_sync_cad_transform_before_extract(self):
    # 기존 추출 함수는 AutoCAD 텍스트를 선택한 상태에서만 화면 값을 반영하므로,
    # 미연결 상태에서 바꾼 추출방식도 추출 직전에 반영한다.
    if self.cad_field_combos and not self.cad_candidates:
        try:
            self.refresh_cad_mapping_from_ui(persist=True)
        except Exception:
            pass


def _v37_extract_active_dwg(self):
    _v37_sync_cad_transform_before_extract(self)
    return _V37_PREV_EXTRACT_ACTIVE_DWG(self)


def _v37_extract_dwg_folder_core_console(self):
    _v37_sync_cad_transform_before_extract(self)
    return _V37_PREV_EXTRACT_DWG_FOLDER(self)


apply_cad_field_transform = _v37_apply_cad_field_transform
make_cad_field_mapping = _v37_make_cad_field_mapping

LineIsoDesktopTool.build_cad_mapping_combos = _v37_build_cad_mapping_combos
LineIsoDesktopTool.refresh_cad_mapping_from_ui = _v37_refresh_cad_mapping_from_ui
LineIsoDesktopTool.extract_active_dwg = _v37_extract_active_dwg
LineIsoDesktopTool.extract_dwg_folder_core_console = _v37_extract_dwg_folder_core_console


'''
V138_BLOCK = r'''# =========================================================
# 8R. V1.38 Source에 Line No. 칸이 없는 경우: ISO DWG No. 포함 매칭
# =========================================================
# 프로젝트마다 ISO DWG No. 형식이 달라 구분자/정규식으로 Line No.를 잘라낼 수 없는 경우를 위해,
# Step 3에서 체크하면 "Line List의 Line No.가 Source의 ISO DWG No.(또는 Line No./파일명) 안에 들어 있으면"
# 같은 Line으로 짝을 짓는다. 비교 직전에 Source 각 행의 line_no_normalized를 찾은 Line List Key로 바꾸므로
# Report 기준 3가지(Data Source / Line List / 전체)가 모두 같은 방식으로 동작한다.
#
# 오매칭 방지 규칙
# 1) Source 값이 Line List Key와 정확히 같으면 그대로 사용 (EXACT)
# 2) 구분자(- . / _) 경계로 떨어지는 부분 문자열 중 Line List Key와 같은 것 (ISO_NO_CONTAINS)
# 3) 2)가 없으면 경계와 상관없이 포함되는 6자 이상 Key (ISO_NO_CONTAINS_LOOSE)
# 2)/3)에서 여러 개가 걸리면 가장 긴 Key를 쓰고, 같은 길이로 2개 이상이면 매칭하지 않는다 (AMBIGUOUS).

_V38_SEPARATORS = set("-/._")
_V38_LOOSE_MIN_LENGTH = 6
_V38_PREV_CREATE_COMPARE_REPORT = LineIsoDesktopTool.create_compare_report
_V38_PREV_BUILD_COMPARE_DATA_TAB = LineIsoDesktopTool._build_compare_data_tab


def _v38_pick_longest(found) -> tuple:
    found = sorted(set(found), key=len, reverse=True)
    if not found:
        return "", ""
    if len(found) > 1 and len(found[0]) == len(found[1]):
        return "", "AMBIGUOUS: " + ", ".join(k for k in found if len(k) == len(found[0]))
    return found[0], ""


def _v38_find_contained_key(text: str, line_keys: set, loose_keys: list) -> tuple:
    if not text:
        return "", ""

    starts = [0] + [i + 1 for i, ch in enumerate(text) if ch in _V38_SEPARATORS]
    ends = [i for i, ch in enumerate(text) if ch in _V38_SEPARATORS] + [len(text)]
    bounded = [text[i:j] for i in starts for j in ends if j > i and text[i:j] in line_keys]
    key, note = _v38_pick_longest(bounded)
    if key or note:
        return key, note or "ISO_NO_CONTAINS"

    key, note = _v38_pick_longest([k for k in loose_keys if k in text])
    if key or note:
        return key, note or "ISO_NO_CONTAINS_LOOSE"
    return "", ""


def resolve_source_line_keys_by_iso_contains(source_df: pd.DataFrame, line_df: pd.DataFrame) -> tuple:
    # Source 행마다 Line List Key를 찾아 line_no_normalized에 넣고, 매칭 방식은 line_no_match 열에 남긴다.
    source_df = source_df.copy()
    if "line_no_normalized" not in line_df.columns:
        line_df = line_df.copy()
        line_df["line_no_normalized"] = line_df.get("line_no", pd.Series([""] * len(line_df))).apply(normalize_line_no)
    line_keys = {normalize_text(k) for k in line_df["line_no_normalized"].tolist() if normalize_text(k)}
    loose_keys = [k for k in line_keys if len(k) >= _V38_LOOSE_MIN_LENGTH]

    resolved_keys = []
    methods = []
    for _, row in source_df.iterrows():
        texts = [
            normalize_line_no(row.get("line_no", "")),
            normalize_line_no(row.get("iso_dwg_no", "")),
            normalize_line_no(Path(normalize_text(row.get("source_file", ""))).stem),
        ]
        current = normalize_text(row.get("line_no_normalized", "")) or texts[0]

        if current in line_keys:
            resolved_keys.append(current)
            methods.append("EXACT")
            continue

        key, method = "", "NOT_FOUND"
        for text in texts:
            found_key, found_method = _v38_find_contained_key(text, line_keys, loose_keys)
            if found_key:
                key, method = found_key, found_method
                break
            if found_method:
                method = found_method
                break

        resolved_keys.append(key or current)
        methods.append(method)

    source_df["line_no_normalized"] = resolved_keys
    source_df["line_no_match"] = methods
    return source_df, pd.Series(methods).value_counts().to_dict() if methods else {}


def _v38_line_no_contains_enabled(self) -> bool:
    checkbox = getattr(self, "line_no_contains_checkbox", None)
    if checkbox is not None:
        return checkbox.isChecked()
    return bool(self.config.get("compare", {}).get("line_no_contained_in_iso", False))


def _v38_create_compare_report(self):
    if not _v38_line_no_contains_enabled(self):
        return _V38_PREV_CREATE_COMPARE_REPORT(self)

    try:
        extracted_dir = self.save_path_config()
        line_path = extracted_dir / "line_list_extracted.xlsx"
        source_key, source_label, source_path, fallback_paths = self.get_compare_source_info()
        if not source_path.exists():
            source_path = next((p for p in fallback_paths if p.exists()), source_path)
        if not line_path.exists() or not source_path.exists():
            # 파일이 없으면 기존 함수가 안내 메시지를 띄운다.
            return _V38_PREV_CREATE_COMPARE_REPORT(self)

        self.setCursor(Qt.WaitCursor)
        self.append_compare_log("ISO DWG No. 포함 매칭으로 Source Line No. 찾는 중...", 3)
        line_df = pd.read_excel(line_path, dtype=str).fillna("")
        source_df = pd.read_excel(source_path, dtype=str).fillna("")
        resolved_df, stats = resolve_source_line_keys_by_iso_contains(source_df, line_df)
        resolved_path = extracted_dir / f"{source_path.stem}_line_no_resolved.xlsx"
        resolved_df.to_excel(resolved_path, index=False, engine="openpyxl")
        self.setCursor(Qt.ArrowCursor)
    except Exception as e:
        self.setCursor(Qt.ArrowCursor)
        show_error(self, "ISO DWG No. 포함 매칭 오류", str(e))
        return None

    # 기존 비교 로직이 매칭 결과 파일을 Source로 읽도록 잠시 바꿔 끼운다.
    self.get_compare_source_info = lambda: (source_key, source_label, resolved_path, [])
    try:
        result = _V38_PREV_CREATE_COMPARE_REPORT(self)
    finally:
        del self.get_compare_source_info

    stat_text = ", ".join(f"{k} {v}건" for k, v in stats.items())
    self.append_compare_log(f"ISO DWG No. 포함 매칭 결과: {stat_text}", None)
    self.append_compare_log(f"행별 매칭 방식(line_no_match 열): {resolved_path}", None)
    return result


def _v38_build_compare_data_tab(self):
    _V38_PREV_BUILD_COMPARE_DATA_TAB(self)
    tab = self.tabs.widget(self.tabs.count() - 1)
    root = tab.layout()

    group = QGroupBox("Line No. 매칭 방식")
    layout = QVBoxLayout(group)
    self.line_no_contains_checkbox = QCheckBox(
        "Source(ISO)에 Line No. 칸이 없음 - Line List의 Line No.가 ISO DWG No. 안에 포함되어 있으면 같은 Line으로 매칭"
    )
    self.line_no_contains_checkbox.setChecked(bool(self.config.get("compare", {}).get("line_no_contained_in_iso", False)))

    def save_option(*args):
        self.config.setdefault("compare", {})["line_no_contained_in_iso"] = self.line_no_contains_checkbox.isChecked()
        save_config(self.config)

    self.line_no_contains_checkbox.toggled.connect(save_option)
    layout.addWidget(self.line_no_contains_checkbox)

    note = QLabel(
        "예: Line List '30100109HE' ↔ ISO DWG No. 'AGCC.1917-3010-30100109HE-TK10.ISO-0002' → 같은 Line. "
        "구분자(- . / _) 경계로 딱 떨어지는 값을 우선하고, 여러 개가 걸리면 가장 긴 Line No.를 사용합니다. "
        "같은 길이로 2개 이상 걸리면 매칭하지 않습니다. "
        "구분자/정규식으로 Line No.를 정확히 잘라낼 수 있는 프로젝트는 Step 2 추출방식을 쓰는 것이 더 정확합니다. "
        "행별 매칭 결과는 Extracted 폴더의 *_line_no_resolved.xlsx (line_no_match 열)에서 확인할 수 있습니다."
    )
    note.setWordWrap(True)
    note.setStyleSheet("color: #555;")
    layout.addWidget(note)

    insert_at = root.count()
    for i in range(root.count()):
        widget = root.itemAt(i).widget()
        if isinstance(widget, QGroupBox) and "Report 기준" in normalize_text(widget.title()):
            insert_at = i + 1
            break
    root.insertWidget(insert_at, group)


LineIsoDesktopTool.create_compare_report = _v38_create_compare_report
LineIsoDesktopTool._build_compare_data_tab = _v38_build_compare_data_tab


'''

CODE_BLOCKS = [
    ("AutoCAD 추출값 Customizing (구분자로 나누기 / 정규식 추출)", V137_MARKER, V137_BLOCK),
    ("ISO DWG No. 포함 Line No. 매칭 옵션", V138_MARKER, V138_BLOCK),
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
                f"원본이 V1.34~V1.37 그대로인지 확인하세요: {src}"
            )
        text = text.replace(old, new)
        print(f"[적용] {name}")

    for name, marker, block in CODE_BLOCKS:
        if marker in text:
            print(f"[건너뜀] 이미 적용됨: {name}")
        elif text.count(RUN_SECTION_ANCHOR) == 1:
            text = text.replace(RUN_SECTION_ANCHOR, block + RUN_SECTION_ANCHOR)
            print(f"[적용] {name}")
        else:
            raise SystemExit("[중단] '# 9. 실행' 위치를 찾지 못했습니다.")

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
