# -*- coding: utf-8 -*-
"""
line_iso_desktop_tool V1.34 ~ V1.38 -> V1.39 패치 스크립트

사용법 (지금 쓰는 line_iso_desktop_tool_V1.3x.py 와 같은 폴더에 이 파일을 두고):
    py apply_fix.py
또는 경로 직접 지정:
    py apply_fix.py "C:\\...\\line_iso_desktop_tool_V1.38.py"

하는 일
1) line_iso_desktop_tool_V1.39.py 생성 (원본 파일은 그대로 둔다)
   - V1.35: get_extracted_dir() / get_reports_dir() 가 저장된 폴더를 만들 수 없으면
     (다른 PC의 D: 드라이브 등) 기본 폴더(output/extracted, output/reports)로 자동 전환
   - V1.36: Step 3 비교 규칙 표에서 "QTableView::setSpan: single cell span won't be added"
     경고가 콘솔에 반복 출력되던 문제 수정
   - V1.37: AutoCAD Mapping의 모든 항목에 추출방식(값 그대로 / 구분자로 나누기 / 정규식 추출)과
     결과 미리보기 추가. ISO DWG No. 안에 들어 있는 Line No.만 잘라 비교 Key로 사용할 수 있다.
   - V1.39: Step 3에 "ISO DWG ↔ Line No. 매칭표" 추가. ISO DWG에 Line No.가 없을 때
     사용자가 만든 Excel 매칭표로 Line No.를 채워 비교한다. 추출된 ISO 목록으로 빈 양식도 만들어 준다.
     (V1.38의 ISO DWG No. 포함 매칭은 복잡해서 이 방식으로 대체했고, 남아 있으면 제거한다)
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
    "line_iso_desktop_tool_V1.38.py",
    "line_iso_desktop_tool_V1.37.py",
    "line_iso_desktop_tool_V1.36.py",
    "line_iso_desktop_tool_V1.35.py",
    "line_iso_desktop_tool_V1.34.py",
]
DST_NAME = "line_iso_desktop_tool_V1.39.py"
NEW_VERSION = "V1.39"

HEADER_NOTES = [
    "# V1.39: Adds an ISO DWG No. <-> Line No. mapping table (Excel) in Step 3 for drawings without a Line No., "
    "replacing the V1.38 contained-in-ISO-No. matching.\n",
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
V139_MARKER = "# 8S. V1.39 ISO DWG ↔ Line No. 매칭표"
# V1.38에서 넣었다가 V1.39 매칭표로 대체한 블록. 남아 있으면 제거한다.
REMOVED_V138_START = "# =========================================================\n# 8R. V1.38 Source에 Line No. 칸이 없는 경우"
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
V139_BLOCK = r'''# =========================================================
# 8S. V1.39 ISO DWG ↔ Line No. 매칭표
# =========================================================
# ISO DWG에 Line No.가 없는 프로젝트는 사용자가 만든 매칭표(Excel: ISO DWG No. / Line No.)로
# Source 각 행의 Line No.를 채운 뒤 비교한다. 매칭표에 없는 ISO는 추출된 값을 그대로 쓴다.
# Report 기준 3가지(Data Source / Line List / 전체)가 모두 같은 방식으로 동작한다.

_V39_PREV_CREATE_COMPARE_REPORT = LineIsoDesktopTool.create_compare_report
_V39_PREV_BUILD_COMPARE_DATA_TAB = LineIsoDesktopTool._build_compare_data_tab


def _v39_iso_key(value) -> str:
    return re.sub(r"\s+", "", normalize_text(value)).upper()


def _v39_source_iso_value(row) -> str:
    iso = normalize_text(row.get("iso_dwg_no", ""))
    if iso:
        return iso
    return Path(normalize_text(row.get("source_file", ""))).stem


def load_iso_line_no_table(path: Path) -> tuple:
    # 첫 Sheet에서 ISO/DWG가 들어간 머리글을 ISO DWG No. 열, LINE이 들어간 머리글을 Line No. 열로 쓴다.
    # 머리글을 못 찾으면 A열=ISO DWG No., B열=Line No.로 본다.
    df = pd.read_excel(path, dtype=str).fillna("")
    if df.shape[1] < 2:
        raise ValueError("매칭표에는 ISO DWG No.와 Line No. 2개 열이 필요합니다.")
    headers = [normalize_text(c).upper() for c in df.columns]
    iso_idx = next((i for i, h in enumerate(headers) if "ISO" in h or "DWG" in h), 0)
    line_idx = next((i for i, h in enumerate(headers) if "LINE" in h and i != iso_idx), 1 if iso_idx == 0 else 0)

    table = {}
    duplicates = 0
    for iso, line_no in zip(df.iloc[:, iso_idx], df.iloc[:, line_idx]):
        key, line_no = _v39_iso_key(iso), normalize_text(line_no)
        if not key or not line_no:
            continue
        if key in table:
            duplicates += 1
            continue
        table[key] = line_no
    return table, duplicates


def apply_iso_line_no_table(source_df: pd.DataFrame, table: dict) -> tuple:
    source_df = source_df.copy()
    if "line_no" not in source_df.columns:
        source_df["line_no"] = ""
    line_nos, keys, flags = [], [], []
    for _, row in source_df.iterrows():
        line_no = table.get(_v39_iso_key(_v39_source_iso_value(row)), "")
        if line_no:
            flags.append("매칭표")
        else:
            line_no = normalize_text(row.get("line_no", ""))
            flags.append("매칭표에 없음")
        line_nos.append(line_no)
        keys.append(normalize_line_no(line_no))
    source_df["line_no"] = line_nos
    source_df["line_no_normalized"] = keys
    source_df["line_no_from"] = flags
    return source_df, flags.count("매칭표"), flags.count("매칭표에 없음")


def _v39_table_enabled(self) -> bool:
    checkbox = getattr(self, "iso_line_table_checkbox", None)
    if checkbox is not None:
        return checkbox.isChecked()
    return bool(self.config.get("compare", {}).get("iso_line_table_enabled", False))


def _v39_save_table_option(self):
    compare = self.config.setdefault("compare", {})
    compare["iso_line_table_enabled"] = self.iso_line_table_checkbox.isChecked()
    compare["iso_line_table_path"] = normalize_text(self.iso_line_table_edit.text())
    save_config(self.config)


def _v39_choose_iso_line_table(self):
    current = normalize_text(self.iso_line_table_edit.text())
    start_dir = _v22_existing_dir(current, self.get_extracted_dir())
    file_path, _ = QFileDialog.getOpenFileName(self, "ISO DWG ↔ Line No. 매칭표 선택", str(start_dir), "Excel Files (*.xlsx *.xlsm *.xls)")
    if file_path:
        self.iso_line_table_edit.setText(file_path)
        self.iso_line_table_checkbox.setChecked(True)
        _v39_save_table_option(self)


def _v39_make_iso_line_table_template(self):
    # 추출된 Source의 ISO DWG No. 목록으로 빈 매칭표를 만든다. Line No. 열만 채우면 된다.
    try:
        source_key, source_label, source_path, fallback_paths = self.get_compare_source_info()
        if not source_path.exists():
            source_path = next((p for p in fallback_paths if p.exists()), source_path)
        if not source_path.exists():
            show_warn(self, "확인", f"{source_label} 추출 파일이 없습니다. Step 2에서 먼저 추출하세요.\n{source_path}")
            return
        source_df = pd.read_excel(source_path, dtype=str).fillna("")
        iso_list = []
        for _, row in source_df.iterrows():
            iso = _v39_source_iso_value(row)
            if iso and iso not in iso_list:
                iso_list.append(iso)

        output_path = self.get_extracted_dir() / "iso_line_no_table.xlsx"
        if output_path.exists():
            # 사용자가 작성 중인 매칭표를 덮어쓰지 않는다.
            output_path = output_path.with_name(f"iso_line_no_table_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        pd.DataFrame({"ISO DWG No.": iso_list, "Line No.": [""] * len(iso_list)}).to_excel(output_path, index=False, engine="openpyxl")

        self.iso_line_table_edit.setText(str(output_path))
        self.iso_line_table_checkbox.setChecked(True)
        _v39_save_table_option(self)
        show_info(
            self,
            "매칭표 양식 생성",
            f"ISO DWG No. {len(iso_list)}개로 매칭표 양식을 만들었습니다.\n{output_path}\n\n"
            "Line No. 열에 Line List와 같은 형식의 Line No.를 채우고 저장한 뒤 Report를 생성하세요.",
        )
        self.open_excel_file(output_path, "매칭표 없음")
    except Exception as e:
        show_error(self, "매칭표 양식 생성 오류", str(e))


def _v39_create_compare_report(self):
    if not _v39_table_enabled(self):
        return _V39_PREV_CREATE_COMPARE_REPORT(self)

    table_path = Path(normalize_text(self.iso_line_table_edit.text())) if getattr(self, "iso_line_table_edit", None) is not None else Path("")
    if not table_path.is_file():
        show_warn(self, "확인", "ISO DWG ↔ Line No. 매칭표 파일을 선택하세요.\n(매칭표를 쓰지 않으려면 체크를 해제하세요.)")
        return None

    try:
        extracted_dir = self.save_path_config()
        source_key, source_label, source_path, fallback_paths = self.get_compare_source_info()
        if not source_path.exists():
            source_path = next((p for p in fallback_paths if p.exists()), source_path)
        if not source_path.exists():
            # 파일이 없으면 기존 함수가 안내 메시지를 띄운다.
            return _V39_PREV_CREATE_COMPARE_REPORT(self)

        table, duplicates = load_iso_line_no_table(table_path)
        source_df = pd.read_excel(source_path, dtype=str).fillna("")
        resolved_df, matched, unmatched = apply_iso_line_no_table(source_df, table)
        resolved_path = extracted_dir / f"{source_path.stem}_with_line_no_table.xlsx"
        resolved_df.to_excel(resolved_path, index=False, engine="openpyxl")
    except Exception as e:
        show_error(self, "매칭표 적용 오류", str(e))
        return None

    # 기존 비교 로직이 매칭표를 적용한 파일을 Source로 읽도록 잠시 바꿔 끼운다.
    self.get_compare_source_info = lambda: (source_key, source_label, resolved_path, [])
    try:
        result = _V39_PREV_CREATE_COMPARE_REPORT(self)
    finally:
        del self.get_compare_source_info

    self.append_compare_log(f"매칭표 적용: {table_path.name} / 매칭 {matched}건, 매칭표에 없음 {unmatched}건", None)
    if duplicates:
        self.append_compare_log(f"매칭표에 같은 ISO DWG No.가 {duplicates}건 중복되어 첫 번째 행만 사용했습니다.", None)
    self.append_compare_log(f"매칭표 적용 결과(line_no_from 열): {resolved_path}", None)
    return result


def _v39_build_compare_data_tab(self):
    _V39_PREV_BUILD_COMPARE_DATA_TAB(self)
    tab = self.tabs.widget(self.tabs.count() - 1)
    root = tab.layout()
    compare_cfg = self.config.get("compare", {})

    group = QGroupBox("ISO DWG ↔ Line No. 매칭표 (ISO DWG에 Line No.가 없을 때)")
    layout = QVBoxLayout(group)

    self.iso_line_table_checkbox = QCheckBox("매칭표로 Source의 Line No. 채우기")
    self.iso_line_table_checkbox.setChecked(bool(compare_cfg.get("iso_line_table_enabled", False)))
    layout.addWidget(self.iso_line_table_checkbox)

    file_row = QHBoxLayout()
    self.iso_line_table_edit = QLineEdit(normalize_text(compare_cfg.get("iso_line_table_path", "")))
    self.iso_line_table_edit.setPlaceholderText("A열 ISO DWG No. / B열 Line No. 로 작성한 Excel")
    choose_btn = QPushButton("매칭표 선택")
    choose_btn.clicked.connect(lambda _checked=False: _v39_choose_iso_line_table(self))
    template_btn = QPushButton("매칭표 양식 만들기")
    template_btn.clicked.connect(lambda _checked=False: _v39_make_iso_line_table_template(self))
    file_row.addWidget(QLabel("매칭표:"))
    file_row.addWidget(self.iso_line_table_edit, 1)
    file_row.addWidget(choose_btn)
    file_row.addWidget(template_btn)
    layout.addLayout(file_row)

    note = QLabel(
        "[매칭표 양식 만들기]를 누르면 추출된 ISO DWG No. 목록으로 Excel이 만들어집니다. "
        "Line No. 열에 Line List와 같은 형식의 Line No.를 채운 뒤 Report를 생성하세요. "
        "매칭표에 없는 ISO DWG는 추출된 Line No.를 그대로 사용합니다."
    )
    note.setWordWrap(True)
    note.setStyleSheet("color: #555;")
    layout.addWidget(note)

    self.iso_line_table_checkbox.toggled.connect(lambda *_: _v39_save_table_option(self))
    self.iso_line_table_edit.editingFinished.connect(lambda: _v39_save_table_option(self))

    insert_at = root.count()
    for i in range(root.count()):
        widget = root.itemAt(i).widget()
        if isinstance(widget, QGroupBox) and "Report 기준" in normalize_text(widget.title()):
            insert_at = i + 1
            break
    root.insertWidget(insert_at, group)


LineIsoDesktopTool.create_compare_report = _v39_create_compare_report
LineIsoDesktopTool._build_compare_data_tab = _v39_build_compare_data_tab


'''

CODE_BLOCKS = [
    ("AutoCAD 추출값 Customizing (구분자로 나누기 / 정규식 추출)", V137_MARKER, V137_BLOCK),
    ("ISO DWG ↔ Line No. 매칭표", V139_MARKER, V139_BLOCK),
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
                f"원본이 V1.34~V1.38 그대로인지 확인하세요: {src}"
            )
        text = text.replace(old, new)
        print(f"[적용] {name}")

    if REMOVED_V138_START in text:
        start = text.index(REMOVED_V138_START)
        # V1.38 블록은 항상 '# 9. 실행' 바로 앞에 들어갔다.
        end = text.index(RUN_SECTION_ANCHOR, start)
        text = text[:start] + text[end:]
        text = "\n".join(line for line in text.split("\n") if not line.startswith("# V1.38: "))
        print("[제거] V1.38 ISO DWG No. 포함 매칭 (매칭표 방식으로 대체)")

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
