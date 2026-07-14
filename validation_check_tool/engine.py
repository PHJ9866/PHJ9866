"""Core logic: reading Line List / Instrument Datasheets, auto-mapping columns,
comparing Process Data, and building the highlighted Validation Excel report.

No file dialogs / console prompts live here - this module is UI-agnostic so the
GUI (or a future CLI) can drive it.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

PROCESS_FIELDS = ["p_op", "p_des", "t_op", "t_op_max", "t_min", "t_max"]

FIELD_LABELS = {
    "tag": "Tag No",
    "line": "Line No",
    "p_op": "Operating Pressure",
    "p_des": "Design Pressure",
    "t_op": "Operating Temperature",
    "t_op_max": "Max Operating Temperature",
    "t_min": "Min Design Temperature",
    "t_max": "Max Design Temperature",
}

FIELD_ORDER = ["tag", "line", "p_des", "p_op", "t_max", "t_op_max", "t_min", "t_op"]


def _has(text: str, *keywords: str) -> bool:
    return all(re.search(k, text) for k in keywords)


def _lacks(text: str, *keywords: str) -> bool:
    return not any(re.search(k, text) for k in keywords)


# Header cells for grouped columns (e.g. a merged "Temperature" label above
# separate "Operating"/"Design Minimum"/"Design Maximum" sub-columns) can put the
# keywords in either order once concatenated top-to-bottom, so matching is done
# by presence/absence of each keyword rather than a fixed left-to-right sequence.
FIELD_MATCHERS = {
    "tag": lambda t: bool(re.search(r"TAG", t)),
    "line": lambda t: bool(re.search(r"LINE", t)) or bool(re.search(r"P\s*&\s*ID\s*LINE", t)),
    "p_des": lambda t: _has(t, r"PRESS", r"DESIGN"),
    "p_op": lambda t: _has(t, r"PRESS") and _lacks(t, r"DESIGN"),
    "t_max": lambda t: _has(t, r"TEMP", r"MAX", r"DESIGN"),
    "t_op_max": lambda t: _has(t, r"TEMP", r"MAX") and _lacks(t, r"DESIGN"),
    "t_min": lambda t: _has(t, r"TEMP", r"MIN"),
    "t_op": lambda t: _has(t, r"TEMP", r"OPER") and _lacks(t, r"MAX", r"MIN"),
}

TAG_TYPE_KEYWORDS = ["FT", "FZT", "PT", "PG", "FV", "XV", "PSV"]

FAMILY_KEYWORDS = {
    "FT": [
        "CORIOLIS", "THERMALMASS", "VORTEX", "ULTRASONIC", "ROTAMETER",
        "TEMPERATURE TX", "TEMP GAUGE WITH WELL", "RO", "VENTURI", "ORIFICE",
        "FLOW GAUGE TOTALIZER", "ULTRASONIC (SPOOL)", "TEMPERATURE ELEMENT",
        "LOCAL INDICATOR", "THERMOWELL",
    ],
    "PT": ["PT PDT TX", "PRESSURE GAUGE", "DP FLOW TX"],
    "VALVE": ["CONTROL VV", "CONTROL VALVE", "SPECIAL CONTROL VV", "ON_OFF", "MOV"],
    "PSV": ["PSV", "PRV", "RUPTURE DISC"],
}

DEFAULT_MAP = {
    "FT": {"tag": "B", "line": "I", "p_op": "X", "p_des": "BM", "t_op": "AC", "t_min": "BP", "t_max": "BR"},
    "PT": {"tag": "B", "line": "I", "p_op": "W", "p_des": "AW", "t_op": "AB", "t_min": "AZ", "t_max": "BB"},
    "VALVE": {"tag": "B", "line": "I", "p_op": "W", "p_des": "BL", "t_op": "AF", "t_min": "BO", "t_max": "BQ"},
    "PSV": {"tag": "B", "line": "J", "p_op": "R", "p_des": "AK", "t_op": "V", "t_min": "AN", "t_max": "AP"},
}

# Fallback used when the Line List's own headers don't match any FIELD_PATTERNS -
# this mirrors the fixed column layout the original script always assumed, so a
# sheet auto-detection can't find columns for doesn't silently degrade to column A.
MASTER_DEFAULT = {"line": "G", "p_op": "N", "p_des": "O", "t_op": "P", "t_min": "R", "t_max": "S"}


# =====================================================
# small helpers
# =====================================================

def clean(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def col_to_index(col: str) -> int:
    col = col.upper()
    result = 0
    for c in col:
        result = result * 26 + (ord(c) - ord("A") + 1)
    return result - 1


def index_to_col(idx: int) -> str:
    return get_column_letter(idx + 1)


def normalize_sheet_name(sheet: str) -> str:
    parts = sheet.rsplit("-", 1)
    if len(parts) == 2:
        return parts[0].strip()
    return sheet.strip()


def detect_type(tag) -> str:
    tag = str(tag).upper()
    for t in TAG_TYPE_KEYWORDS:
        if f"-{t}" in tag:
            return t
    return "UNKNOWN"


def detect_family(sheet_name: str) -> str | None:
    name = sheet_name.upper()
    for family, keywords in FAMILY_KEYWORDS.items():
        for k in keywords:
            if k in name:
                return family
    return None


# =====================================================
# comparison rules
# =====================================================

def extract_numeric(value) -> float | None:
    match = re.search(r"-?\d+(?:\.\d+)?", str(value).strip())
    return float(match.group()) if match else None


def compare_numeric(master, inst, tolerance: float = 0.05) -> str:
    m = extract_numeric(master)
    i = extract_numeric(inst)
    if m is None or i is None:
        return "PASS" if str(master).strip() == str(inst).strip() else "FAIL"
    return "PASS" if abs(m - i) <= tolerance else "FAIL"


def compare_pressure_op(master, inst) -> str:
    master_text = str(master).replace(" ", "")
    inst_text = str(inst).replace(" ", "")
    try:
        if "~" in master_text:
            low, high = (float(x) for x in master_text.split("~"))
            value = extract_numeric(inst_text)
            if value is None:
                return "FAIL"
            return "PASS" if low <= value <= high else "FAIL"
        return compare_numeric(master_text, inst_text, 0.05)
    except Exception:
        return "FAIL"


def compare_temp_op(master, inst) -> str:
    master_s = str(master).strip().upper()
    if master_s == "AMB":
        return "N/A"
    return "PASS" if str(master).strip() == str(inst).strip() else "FAIL"


COMPARATORS = {
    "p_op": compare_pressure_op,
    "p_des": compare_numeric,
    "t_op": compare_temp_op,
    "t_op_max": compare_numeric,
    "t_min": compare_numeric,
    "t_max": compare_numeric,
}

# t_op is intentionally excluded from the overall PASS/FAIL verdict: it is very
# often "AMB" on the line list against a real number on the datasheet, which is
# an expected, not an erroneous, difference.
FIELDS_IN_VERDICT = ["p_op", "p_des", "t_op_max", "t_min", "t_max"]


# =====================================================
# auto column mapping
# =====================================================

def _is_data_row(row) -> bool:
    non_null = int(row.notna().sum())
    if non_null < 3:
        return False
    numeric_like = sum(1 for v in row if pd.notna(v) and extract_numeric(v) is not None)
    return numeric_like / non_null >= 0.6


def detect_header_row_count(df: pd.DataFrame, max_rows: int = 40, confirm_rows: int = 3) -> int:
    """Finds where the header block ends and real data begins, so header text
    scanning doesn't scoop up actual data values (tag numbers, pressures, ...)
    as if they were column headers. A single numeric-looking row isn't enough -
    a lone row of column-index numbers (e.g. "1 2 3 ...") above the real headers
    would look like data too - so `confirm_rows` consecutive rows must all look
    like data before that point is treated as where the header block ends."""
    limit = min(max_rows, len(df))
    for r in range(limit):
        window = df.iloc[r: min(r + confirm_rows, len(df))]
        if len(window) > 0 and all(_is_data_row(window.iloc[i]) for i in range(len(window))):
            return r
    return limit


def build_header_text(df: pd.DataFrame, header_rows: int | None = None) -> list[str]:
    if header_rows is None:
        header_rows = detect_header_row_count(df)
    n = min(header_rows, len(df))
    if n == 0:
        return ["" for _ in range(df.shape[1])]

    # Forward-fill each header row left-to-right so a merged group label (e.g.
    # "Temperature" spanning several columns, stored only in the leftmost cell)
    # propagates onto every column it visually covers, instead of only the first.
    filled = df.iloc[:n].apply(lambda row: row.ffill(), axis=1)

    headers = []
    for c in range(df.shape[1]):
        parts = []
        for r in range(n):
            v = filled.iat[r, c]
            if pd.notna(v):
                s = str(v).strip()
                if s and s not in parts:
                    parts.append(s)
        headers.append(" ".join(parts).upper())
    return headers


def column_preview(df: pd.DataFrame, col_letter: str, max_len: int = 110) -> str:
    """Human-readable preview of a column: its header text plus one sample value,
    used in the GUI so the user can sanity-check a mapping without opening Excel."""
    idx = col_to_index(col_letter)
    if idx < 0 or idx >= df.shape[1]:
        return "(범위 밖)"
    header_rows = detect_header_row_count(df)
    header = " / ".join(
        str(df.iat[r, idx]).strip()
        for r in range(min(header_rows, len(df)))
        if pd.notna(df.iat[r, idx]) and str(df.iat[r, idx]).strip()
    )
    sample = ""
    for r in range(header_rows, len(df)):
        v = df.iat[r, idx]
        if pd.notna(v) and str(v).strip():
            sample = str(v).strip()
            break
    header = header or "(헤더 없음)"
    text = f"{header}   [예: {sample}]" if sample else header
    if len(text) > max_len:
        text = text[:max_len] + "..."
    return text


def auto_detect_mapping(df: pd.DataFrame) -> dict[str, str]:
    headers = build_header_text(df)
    used: set[int] = set()
    mapping: dict[str, str] = {}
    for f in FIELD_ORDER:
        found = None
        for c, h in enumerate(headers):
            if c in used or not h:
                continue
            if FIELD_MATCHERS[f](h):
                found = c
                break
        if found is not None:
            used.add(found)
            mapping[f] = index_to_col(found)
    return mapping


def resolve_mapping(default: dict[str, str], auto_map: dict[str, str]) -> tuple[dict, dict]:
    mapping, source = {}, {}
    for f in ["tag", "line", *PROCESS_FIELDS]:
        if f in auto_map:
            mapping[f], source[f] = auto_map[f], "auto"
        elif f in default:
            mapping[f], source[f] = default[f], "default"
        else:
            mapping[f], source[f] = "A", "missing"
    return mapping, source


# =====================================================
# data containers
# =====================================================

@dataclass
class SheetMapping:
    file: str
    sheet: str
    key: str
    family: str | None
    mapping: dict = field(default_factory=dict)
    source: dict = field(default_factory=dict)
    include: bool = True


@dataclass
class MasterLine:
    line_no: str
    p_op: str
    p_des: str
    t_op: str
    t_op_max: str
    t_min: str
    t_max: str


@dataclass
class InstrumentRow:
    tag: str
    inst_type: str
    line_no: str
    p_op: str
    p_des: str
    t_op: str
    t_op_max: str
    t_min: str
    t_max: str
    source_file: str
    source_sheet: str


# =====================================================
# master (line list) loading
# =====================================================

MASTER_KEY = "__MASTER__"


def scan_master_file(master_file: str, saved_config: dict) -> tuple[SheetMapping, pd.DataFrame]:
    df = pd.read_excel(master_file, sheet_name=0, header=None)
    if MASTER_KEY in saved_config:
        mapping = dict(saved_config[MASTER_KEY])
        source = {f: "saved" for f in mapping}
    else:
        auto_map = auto_detect_mapping(df)
        mapping, source = resolve_mapping(MASTER_DEFAULT, auto_map)
    sm = SheetMapping(file=master_file, sheet="Line List", key=MASTER_KEY,
                       family="MASTER", mapping=mapping, source=source, include=True)
    return sm, df


def load_master_lines(df: pd.DataFrame, mapping: dict) -> list[MasterLine]:
    idx = {f: col_to_index(mapping[f]) for f in ["line", *PROCESS_FIELDS]}
    lines: list[MasterLine] = []
    for r in range(len(df)):
        try:
            line_no = clean(df.iat[r, idx["line"]])
        except IndexError:
            continue
        if len(line_no.split("-")) < 3:
            continue
        lines.append(
            MasterLine(
                line_no=line_no,
                p_op=clean(df.iat[r, idx["p_op"]]),
                p_des=clean(df.iat[r, idx["p_des"]]),
                t_op=clean(df.iat[r, idx["t_op"]]),
                t_op_max=clean(df.iat[r, idx["t_op_max"]]),
                t_min=clean(df.iat[r, idx["t_min"]]),
                t_max=clean(df.iat[r, idx["t_max"]]),
            )
        )
    return lines


# =====================================================
# instrument datasheet scanning / loading
# =====================================================

def scan_instrument_files(files: list[str], saved_config: dict, progress=None) -> tuple[list[SheetMapping], dict]:
    """Opens every sheet of every instrument file, auto-detects (or reuses saved)
    column mapping, and returns the mapping list plus a df cache for later reuse."""
    results: list[SheetMapping] = []
    df_cache: dict[tuple[str, str], pd.DataFrame] = {}

    for file in files:
        try:
            xl = pd.ExcelFile(file)
        except Exception:
            continue
        for sheet in xl.sheet_names:
            if progress:
                progress(f"스캔 중: {Path(file).name} - {sheet}")
            df = xl.parse(sheet, header=None)
            df_cache[(file, sheet)] = df

            key = normalize_sheet_name(sheet)
            family = detect_family(sheet)

            if key in saved_config:
                mapping = dict(saved_config[key])
                source = {f: "saved" for f in mapping}
            else:
                auto_map = auto_detect_mapping(df)
                mapping, source = resolve_mapping(DEFAULT_MAP.get(family, {}), auto_map)

            include = family is not None or any(s == "auto" for s in source.values())
            results.append(SheetMapping(file=file, sheet=sheet, key=key, family=family,
                                         mapping=mapping, source=source, include=include))
    return results, df_cache


def load_instrument_rows(sheet_mappings: list[SheetMapping], df_cache: dict) -> list[InstrumentRow]:
    rows: list[InstrumentRow] = []
    for sm in sheet_mappings:
        if not sm.include:
            continue
        df = df_cache.get((sm.file, sm.sheet))
        if df is None:
            df = pd.read_excel(sm.file, sheet_name=sm.sheet, header=None)

        idx = {f: col_to_index(sm.mapping[f]) for f in ["tag", "line", *PROCESS_FIELDS]}
        ncols = df.shape[1]

        for r in range(len(df)):
            if idx["tag"] >= ncols:
                break
            tag = clean(df.iat[r, idx["tag"]])
            if not tag:
                continue
            inst_type = detect_type(tag)
            if inst_type == "UNKNOWN":
                continue

            def get(field_name):
                i = idx[field_name]
                return clean(df.iat[r, i]) if i < ncols else ""

            rows.append(
                InstrumentRow(
                    tag=tag,
                    inst_type=inst_type,
                    line_no=get("line"),
                    p_op=get("p_op"),
                    p_des=get("p_des"),
                    t_op=get("t_op"),
                    t_op_max=get("t_op_max"),
                    t_min=get("t_min"),
                    t_max=get("t_max"),
                    source_file=Path(sm.file).name,
                    source_sheet=sm.sheet,
                )
            )
    return rows


# =====================================================
# config persistence
# =====================================================

def load_config(path: str) -> dict:
    p = Path(path)
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return {}


def save_config(path: str, sheet_mappings: list[SheetMapping]) -> None:
    config = load_config(path)
    for sm in sheet_mappings:
        if sm.include:
            config[sm.key] = sm.mapping
    Path(path).write_text(json.dumps(config, indent=4, ensure_ascii=False), encoding="utf-8")


# =====================================================
# report generation
# =====================================================

GRAY = PatternFill("solid", fgColor="D9D9D9")
GREEN = PatternFill("solid", fgColor="C6EFCE")
RED = PatternFill("solid", fgColor="FFC7CE")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(color="FFFFFF", bold=True)

PROCESS_FIELD_COLUMNS = [
    ("p_op", "P_Oper"),
    ("p_des", "P_Design"),
    ("t_op", "T_Oper"),
    ("t_op_max", "T_Oper_Max"),
    ("t_min", "T_Min_Design"),
    ("t_max", "T_Max_Design"),
]

REPORT_HEADERS = (
    ["Line No", "Tag No", "Type"]
    + [h for _, h in PROCESS_FIELD_COLUMNS]
    + ["Source File", "Source Sheet", "Result"]
)

PROCESS_START_COL = 4
FIELD_COL = {f: PROCESS_START_COL + i for i, (f, _) in enumerate(PROCESS_FIELD_COLUMNS)}
SOURCE_FILE_COL = PROCESS_START_COL + len(PROCESS_FIELD_COLUMNS)
SOURCE_SHEET_COL = SOURCE_FILE_COL + 1
RESULT_COL = SOURCE_SHEET_COL + 1


@dataclass
class ReportStats:
    total_lines: int = 0
    instrument_count: int = 0
    matched: int = 0
    missing_lines: int = 0
    fail_count: int = 0
    output_rows: int = 0


def build_report(master_lines: list[MasterLine], instrument_rows: list[InstrumentRow],
                  output_path: str, progress=None) -> ReportStats:
    wb = Workbook()
    ws = wb.active
    ws.title = "Validation"

    for col, h in enumerate(REPORT_HEADERS, start=1):
        cell = ws.cell(1, col, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    stats = ReportStats(total_lines=len(master_lines), instrument_count=len(instrument_rows))
    excel_row = 2
    n_cols = len(REPORT_HEADERS)

    for i, line in enumerate(master_lines):
        if progress and i % 25 == 0:
            progress(f"리포트 작성 중: {i}/{len(master_lines)}")

        matches = [inst for inst in instrument_rows if line.line_no in inst.line_no]

        ws.cell(excel_row, 1, line.line_no)
        ws.cell(excel_row, 2, "MASTER")
        ws.cell(excel_row, 3, "LINE")
        for f in PROCESS_FIELDS:
            ws.cell(excel_row, FIELD_COL[f], getattr(line, f))
        for c in range(1, n_cols + 1):
            ws.cell(excel_row, c).fill = GRAY
        excel_row += 1

        if not matches:
            ws.cell(excel_row, 3, "MISSING")
            for c in range(1, n_cols + 1):
                ws.cell(excel_row, c).fill = YELLOW
            excel_row += 1
            stats.missing_lines += 1
            continue

        for inst in matches:
            stats.matched += 1
            results = {f: COMPARATORS[f](getattr(line, f), getattr(inst, f)) for f in PROCESS_FIELDS}
            final = "FAIL" if any(results[f] == "FAIL" for f in FIELDS_IN_VERDICT) else "PASS"
            if final == "FAIL":
                stats.fail_count += 1

            ws.cell(excel_row, 2, inst.tag)
            ws.cell(excel_row, 3, inst.inst_type)
            for f in PROCESS_FIELDS:
                ws.cell(excel_row, FIELD_COL[f], getattr(inst, f))
            ws.cell(excel_row, SOURCE_FILE_COL, inst.source_file)
            ws.cell(excel_row, SOURCE_SHEET_COL, inst.source_sheet)
            ws.cell(excel_row, RESULT_COL, final)

            ws.cell(excel_row, RESULT_COL).fill = GREEN if final == "PASS" else RED
            for field_name in PROCESS_FIELDS:
                r = results[field_name]
                col = FIELD_COL[field_name]
                if r == "PASS":
                    ws.cell(excel_row, col).fill = GREEN
                elif r == "FAIL":
                    ws.cell(excel_row, col).fill = RED

            excel_row += 1

    stats.output_rows = excel_row - 2

    for col_idx, header in enumerate(REPORT_HEADERS, start=1):
        width = max(12, len(header) + 4)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "B2"

    summary = wb.create_sheet("Summary")
    summary_rows = [
        ("Total Lines", stats.total_lines),
        ("Instrument Count", stats.instrument_count),
        ("Matched", stats.matched),
        ("Missing Lines", stats.missing_lines),
        ("Fail Count", stats.fail_count),
        ("Output Rows", stats.output_rows),
    ]
    for r, (label, value) in enumerate(summary_rows, start=1):
        summary.cell(r, 1, label)
        summary.cell(r, 2, value)
    summary.column_dimensions["A"].width = 20

    wb.save(output_path)
    return stats
