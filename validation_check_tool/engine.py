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
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

PROCESS_FIELDS = ["p_op", "p_des_max", "t_op", "t_op_max", "t_min", "t_max"]

FIELD_LABELS = {
    "tag": "Tag No",
    "line": "Line No",
    "p_op": "Operating Pressure",
    "p_des_max": "Max Design Pressure",
    "t_op": "Operating Temperature",
    "t_op_max": "Max Operating Temperature",
    "t_min": "Min Design Temperature",
    "t_max": "Max Design Temperature",
}

FIELD_ORDER = ["tag", "line", "p_des_max", "p_op", "t_max", "t_op_max", "t_min", "t_op"]


def _has(text: str, *keywords: str) -> bool:
    return all(re.search(k, text) for k in keywords)


def _lacks(text: str, *keywords: str) -> bool:
    return not any(re.search(k, text) for k in keywords)


def _match_p_des_max(t: str) -> bool:
    if not _has(t, r"PRESS", r"DESIGN") or _has(t, r"DIFFERENTIAL"):
        return False
    if _has(t, r"MAX"):
        return True
    # Some Line Lists only have a single, unqualified "Design Pressure" column
    # (no Min/Max split) - treat that as the max design rating, but don't steal
    # a column that's explicitly labelled as the design *minimum*.
    return _lacks(t, r"\bMIN\b")


def _match_p_op(t: str) -> bool:
    if not _has(t, r"PRESS") or _has(t, r"DESIGN"):
        return False
    # Hydrotest/strength-test pressure and Differential Pressure (a DP
    # transmitter's own span, not the line's operating pressure) columns also
    # read "... Pressure" but are never the operating pressure.
    if not _lacks(t, r"TEST", r"STRENGTH", r"HYDRO", r"DIFFERENTIAL"):
        return False
    if _has(t, r"NOR(MAL)?"):
        return True
    # Some sheets have an unrelated Min/Nor/Max pressure range (e.g. an
    # "Upstream Pressure" group) sitting next to Design Pressure. Without an
    # explicit Normal/Operating label, only claim a column here if it isn't one
    # of those Min/Max range columns.
    return _lacks(t, r"\bMIN\b", r"\bMAX\b")


def _match_t_op(t: str) -> bool:
    if not _has(t, r"TEMP") or _has(t, r"DESIGN"):
        return False
    if _has(t, r"NOR(MAL)?"):
        return True
    return _lacks(t, r"\bMIN\b", r"\bMAX\b")


# Header cells for grouped columns (e.g. a merged "Temperature" label above
# separate "Operating"/"Design Minimum"/"Design Maximum" sub-columns) can put the
# keywords in either order once concatenated top-to-bottom, so matching is done
# by presence/absence of each keyword rather than a fixed left-to-right sequence.
FIELD_MATCHERS = {
    "tag": lambda t: bool(re.search(r"TAG", t)),
    "line": lambda t: bool(re.search(r"\bLINE\b", t)),
    "p_des_max": _match_p_des_max,
    "p_op": _match_p_op,
    "t_max": lambda t: _has(t, r"TEMP", r"MAX", r"DESIGN"),
    "t_op_max": lambda t: _has(t, r"TEMP", r"MAX") and _lacks(t, r"DESIGN"),
    "t_min": lambda t: _has(t, r"TEMP", r"MIN", r"DESIGN"),
    "t_op": _match_t_op,
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

# These mirror the fixed columns the original script always assumed. "Min Design
# Pressure" and "Max Operating Temperature" didn't exist as tracked fields back
# then, so there's no legacy column for them - they're left for auto-detection /
# manual mapping. The old single "Design Pressure" default is kept as the Max
# Design Pressure default, since that's what it was actually being used for.
DEFAULT_MAP = {
    "FT": {"tag": "B", "line": "I", "p_op": "X", "p_des_max": "BM", "t_op": "AC", "t_min": "BP", "t_max": "BR"},
    "PT": {"tag": "B", "line": "I", "p_op": "W", "p_des_max": "AW", "t_op": "AB", "t_min": "AZ", "t_max": "BB"},
    "VALVE": {"tag": "B", "line": "I", "p_op": "W", "p_des_max": "BL", "t_op": "AF", "t_min": "BO", "t_max": "BQ"},
    "PSV": {"tag": "B", "line": "J", "p_op": "R", "p_des_max": "AK", "t_op": "V", "t_min": "AN", "t_max": "AP"},
}

# Fallback used when the Line List's own headers don't match any FIELD_MATCHERS -
# this mirrors the fixed column layout the original script always assumed, so a
# sheet auto-detection can't find columns for doesn't silently degrade to column A.
MASTER_DEFAULT = {"line": "G", "p_op": "N", "p_des_max": "O", "t_op": "P", "t_min": "R", "t_max": "S"}


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
    "p_des_max": compare_numeric,
    "t_op": compare_temp_op,
    "t_op_max": compare_numeric,
    "t_min": compare_numeric,
    "t_max": compare_numeric,
}

# t_op is intentionally excluded from the overall PASS/FAIL verdict: it is very
# often "AMB" on the line list against a real number on the datasheet, which is
# an expected, not an erroneous, difference.
FIELDS_IN_VERDICT = ["p_op", "p_des_max", "t_op_max", "t_min", "t_max"]


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


def get_merged_ranges_map(file_path: str) -> dict[str, list[tuple[int, int, int, int]]]:
    """0-indexed (min_row, min_col, max_row, max_col) merged ranges for EVERY
    sheet of a workbook, collected in a single load. Used to correctly propagate
    a merged group header (e.g. "Pressure" spanning several columns, whose text
    openpyxl/pandas only report on the top-left cell) onto every column it
    visually covers - without guessing via forward-fill, which can't tell a real
    merge apart from an unrelated blank column sitting next to a label.

    merged_cells is not exposed in read_only mode, so this has to be a normal
    (in-memory) load - which is why it must happen once per FILE: loading the
    whole workbook again for every sheet made scanning large files painfully slow."""
    try:
        wb = load_workbook(file_path, data_only=True)
    except Exception:
        return {}
    result = {}
    for ws in wb.worksheets:
        result[ws.title] = [(r.min_row - 1, r.min_col - 1, r.max_row - 1, r.max_col - 1)
                            for r in ws.merged_cells.ranges]
    wb.close()
    return result


def _effective_header_grid(df: pd.DataFrame, n_rows: int,
                            merges: list[tuple[int, int, int, int]] | None) -> list[list]:
    """Header cell values for rows [0, n_rows), with merged-cell group labels
    copied onto every column their merge spans (instead of just the top-left
    cell where openpyxl/pandas actually store the value)."""
    ncols = df.shape[1]
    grid = [[df.iat[r, c] if pd.notna(df.iat[r, c]) else None for c in range(ncols)] for r in range(n_rows)]
    for min_row, min_col, max_row, max_col in (merges or []):
        if max_row < 0 or min_row >= n_rows or max_col < min_col or min_row >= len(df):
            continue
        top_val = df.iat[min_row, min_col] if pd.notna(df.iat[min_row, min_col]) else None
        if top_val is None:
            continue
        for r in range(max(min_row, 0), min(max_row, n_rows - 1) + 1):
            for c in range(min_col, min(max_col, ncols - 1) + 1):
                grid[r][c] = top_val
    return grid


def build_header_text(df: pd.DataFrame, header_rows: int | None = None,
                       merges: list[tuple[int, int, int, int]] | None = None) -> list[str]:
    if header_rows is None:
        header_rows = detect_header_row_count(df)
    n = min(header_rows, len(df))
    if n == 0:
        return ["" for _ in range(df.shape[1])]

    grid = _effective_header_grid(df, n, merges)

    headers = []
    for c in range(df.shape[1]):
        parts = []
        for r in range(n):
            v = grid[r][c]
            if v is not None:
                s = str(v).strip()
                if s and s not in parts:
                    parts.append(s)
        headers.append(" ".join(parts).upper())
    return headers


def column_preview(df: pd.DataFrame, col_letter: str,
                    merges: list[tuple[int, int, int, int]] | None = None, max_len: int = 110) -> str:
    """Human-readable preview of a column: its header text plus one sample value,
    used in the GUI so the user can sanity-check a mapping without opening Excel."""
    idx = col_to_index(col_letter)
    if idx < 0 or idx >= df.shape[1]:
        return "(범위 밖)"
    header_rows = detect_header_row_count(df)
    n = min(header_rows, len(df))
    grid = _effective_header_grid(df, n, merges)
    parts = []
    for r in range(n):
        v = grid[r][idx]
        if v is not None:
            s = str(v).strip()
            if s and s not in parts:
                parts.append(s)
    header = " / ".join(parts)
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


# Non-numeric placeholders that legitimately appear in process-data columns:
# AMB(ient), ATM(ospheric), F.V (full vacuum), VAC(uum), N/A. Compared after
# stripping dots/spaces so "F.V", "F.V." and "FV" all normalize the same.
NON_NUMERIC_PROCESS_VALUES = {"AMB", "ATM", "FV", "VAC", "FULLVACUUM", "N/A", "NA", "-"}


def _is_process_value(s: str) -> bool:
    if extract_numeric(s) is not None:
        return True
    return s.upper().replace(".", "").replace(" ", "") in NON_NUMERIC_PROCESS_VALUES


def _column_looks_numeric(df: pd.DataFrame, col_idx: int, data_start: int,
                           sample: int = 12, threshold: float = 0.5) -> bool:
    """A header can accidentally say the right keyword (e.g. a "Description"
    column reading "Pressure Transmitter") while holding text, not the actual
    process value. Before trusting a header match for a numeric field, check
    that the column's real data is actually process values (numbers, plus
    placeholders like ATM/AMB/F.V)."""
    count = numeric = 0
    for r in range(data_start, len(df)):
        v = df.iat[r, col_idx]
        if pd.isna(v):
            continue
        s = str(v).strip()
        if not s:
            continue
        count += 1
        if _is_process_value(s):
            numeric += 1
        if count >= sample:
            break
    if count == 0:
        # An empty column tells us nothing and maps to nothing useful - reject
        # the header match so the search continues (or falls back to defaults),
        # instead of letting e.g. an unused "LINE FLUID" column win.
        return False
    return numeric / count >= threshold


def _line_no_stats(df: pd.DataFrame, col_idx: int, data_start: int,
                    sample: int = 12) -> tuple[int, int]:
    """Counts how many of a column's first data values have the distinctive
    line-number shape: dash-separated with 3+ parts and at least one letter
    (e.g. "301-ATM-0007"). Returns (values seen, values matching)."""
    count = matches = 0
    for r in range(data_start, len(df)):
        v = df.iat[r, col_idx]
        if pd.isna(v):
            continue
        s = str(v).strip()
        if not s:
            continue
        count += 1
        if len(s.split("-")) >= 3 and re.search(r"[A-Za-z]", s):
            matches += 1
        if count >= sample:
            break
    return count, matches


def _column_looks_like_line_no(df: pd.DataFrame, col_idx: int, data_start: int,
                                threshold: float = 0.5) -> bool:
    """Same idea as _column_looks_numeric, but for the Line No column: a wide
    Line List can have unrelated columns elsewhere whose header text happens to
    contain the word "Line" (insulation line, LINE FLUID, ...). Require the
    actual data to look like line numbers - and reject empty columns outright."""
    count, matches = _line_no_stats(df, col_idx, data_start)
    if count == 0:
        return False
    return matches / count >= threshold


# A wide Line List / datasheet can have plenty of unrelated columns whose header
# text loosely matches a field's keywords (e.g. a hydrotest pressure column also
# containing the word "Pressure"). On a Line List, Operating and Design Pressure
# are almost always right next to each other, so once the more tightly-matched
# field is found, columns near it are tried first - but a Line List and an
# Instrument datasheet aren't equally compact (a datasheet can have Operating
# and Design Pressure many columns apart, e.g. with UOM/flag sub-columns in
# between), so this only reorders the search rather than ruling out distant
# columns entirely; the numeric-data check is what actually rejects unrelated
# columns like a "Hydrotest Pressure" value.
PROXIMITY_ANCHOR = {"p_op": "p_des_max", "t_op": "t_max", "t_op_max": "t_max"}
PROXIMITY_WINDOW = 10


def _columns_by_proximity(n_cols: int, anchor: int, window: int = PROXIMITY_WINDOW) -> list[int]:
    near = []
    for d in range(window + 1):
        for c in (anchor - d, anchor + d) if d else (anchor,):
            if 0 <= c < n_cols and c not in near:
                near.append(c)
    rest = [c for c in range(n_cols) if c not in near]
    return near + rest


def auto_detect_mapping(df: pd.DataFrame, merges: list[tuple[int, int, int, int]] | None = None) -> dict[str, str]:
    data_start = detect_header_row_count(df)
    headers = build_header_text(df, data_start, merges)
    n = len(headers)
    used: set[int] = set()
    mapping: dict[str, str] = {}
    for f in FIELD_ORDER:
        anchor_field = PROXIMITY_ANCHOR.get(f)
        anchor_col = mapping.get(anchor_field) if anchor_field else None
        anchor_idx = col_to_index(anchor_col) if anchor_col else None
        order = _columns_by_proximity(n, anchor_idx) if anchor_idx is not None else range(n)

        found = None
        for c in order:
            if c in used or not headers[c]:
                continue
            h = headers[c]
            if not FIELD_MATCHERS[f](h):
                continue
            if f in PROCESS_FIELDS and not _column_looks_numeric(df, c, data_start):
                continue
            if f == "line" and not _column_looks_like_line_no(df, c, data_start):
                continue
            found = c
            break
        if found is not None:
            used.add(found)
            mapping[f] = index_to_col(found)

    if "line" not in mapping:
        # A Line List's line-number column often carries an unhelpful header
        # like "No", which no keyword can find. Fall back to recognizing the
        # data itself: dash-separated identifiers containing letters
        # ("301-ATM-0007"), taking the leftmost strongly-matching column.
        for c in range(n):
            if c in used:
                continue
            count, matches = _line_no_stats(df, c, data_start)
            if count >= 2 and matches / count >= 0.7:
                mapping["line"] = index_to_col(c)
                used.add(c)
                break
    return mapping


def resolve_mapping(default: dict[str, str], auto_map: dict[str, str]) -> tuple[dict, dict]:
    mapping, source = {}, {}
    for f in ["tag", "line", *PROCESS_FIELDS]:
        if f in auto_map:
            mapping[f], source[f] = auto_map[f], "auto"
        elif f in default:
            mapping[f], source[f] = default[f], "default"
        else:
            # Unrecognized: leave the mapping empty (the GUI shows "매핑 필요")
            # instead of silently pointing at column A.
            mapping[f], source[f] = "", "missing"
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
    # Snapshot of the mapping/source as originally detected, so a manual edit
    # that puts a value back to what auto-detection found can drop its
    # "manual" status again.
    orig_mapping: dict = field(default_factory=dict)
    orig_source: dict = field(default_factory=dict)

    def __post_init__(self):
        if not self.orig_mapping:
            self.orig_mapping = dict(self.mapping)
        if not self.orig_source:
            self.orig_source = dict(self.source)


@dataclass
class MasterLine:
    line_no: str
    p_op: str
    p_des_max: str
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
    p_des_max: str
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


def scan_master_file(master_file: str, saved_config: dict) -> tuple[SheetMapping, pd.DataFrame, list]:
    xl = pd.ExcelFile(master_file)
    first_sheet = xl.sheet_names[0]
    df = xl.parse(first_sheet, header=None)
    merges = get_merged_ranges_map(master_file).get(first_sheet, [])
    if MASTER_KEY in saved_config:
        mapping = dict(saved_config[MASTER_KEY])
        source = {f: "saved" for f in mapping}
    else:
        auto_map = auto_detect_mapping(df, merges)
        mapping, source = resolve_mapping(MASTER_DEFAULT, auto_map)
    sm = SheetMapping(file=master_file, sheet="Line List", key=MASTER_KEY,
                       family="MASTER", mapping=mapping, source=source, include=True)
    return sm, df, merges


def _mapping_indexes(mapping: dict, fields: list[str], ncols: int) -> dict[str, int | None]:
    """Column indexes for each field; None when the field is unmapped (empty)
    or points outside the sheet, so callers read "" instead of a wrong column."""
    idx: dict[str, int | None] = {}
    for f in fields:
        v = str(mapping.get(f, "") or "").strip()
        i = col_to_index(v) if v else -1
        idx[f] = i if 0 <= i < ncols else None
    return idx


def load_master_lines(df: pd.DataFrame, mapping: dict) -> list[MasterLine]:
    idx = _mapping_indexes(mapping, ["line", *PROCESS_FIELDS], df.shape[1])
    if idx["line"] is None:
        return []
    lines: list[MasterLine] = []
    for r in range(len(df)):
        def get(field_name):
            i = idx[field_name]
            return clean(df.iat[r, i]) if i is not None else ""

        line_no = get("line")
        if len(line_no.split("-")) < 3:
            continue
        lines.append(
            MasterLine(
                line_no=line_no,
                p_op=get("p_op"),
                p_des_max=get("p_des_max"),
                t_op=get("t_op"),
                t_op_max=get("t_op_max"),
                t_min=get("t_min"),
                t_max=get("t_max"),
            )
        )
    return lines


# =====================================================
# instrument datasheet scanning / loading
# =====================================================

def scan_instrument_files(files: list[str], saved_config: dict, progress=None) -> tuple[list[SheetMapping], dict, dict]:
    """Opens every sheet of every instrument file, auto-detects (or reuses saved)
    column mapping, and returns the mapping list plus df/merge caches for later reuse."""
    results: list[SheetMapping] = []
    df_cache: dict[tuple[str, str], pd.DataFrame] = {}
    merge_cache: dict[tuple[str, str], list] = {}

    for file in files:
        try:
            xl = pd.ExcelFile(file)
        except Exception:
            continue
        # One workbook load per FILE for merged-cell info - loading it once per
        # sheet made scanning large multi-sheet files unbearably slow.
        merges_by_sheet = get_merged_ranges_map(file)
        for sheet in xl.sheet_names:
            if progress:
                progress(f"스캔 중: {Path(file).name} - {sheet}")
            df = xl.parse(sheet, header=None)
            df_cache[(file, sheet)] = df
            merges = merges_by_sheet.get(sheet, [])
            merge_cache[(file, sheet)] = merges

            key = normalize_sheet_name(sheet)
            family = detect_family(sheet)

            if key in saved_config:
                mapping = dict(saved_config[key])
                source = {f: "saved" for f in mapping}
            else:
                auto_map = auto_detect_mapping(df, merges)
                mapping, source = resolve_mapping(DEFAULT_MAP.get(family, {}), auto_map)

            include = family is not None or any(s == "auto" for s in source.values())
            results.append(SheetMapping(file=file, sheet=sheet, key=key, family=family,
                                         mapping=mapping, source=source, include=include))
    return results, df_cache, merge_cache


def load_instrument_rows(sheet_mappings: list[SheetMapping], df_cache: dict) -> list[InstrumentRow]:
    rows: list[InstrumentRow] = []
    for sm in sheet_mappings:
        if not sm.include:
            continue
        df = df_cache.get((sm.file, sm.sheet))
        if df is None:
            df = pd.read_excel(sm.file, sheet_name=sm.sheet, header=None)

        idx = _mapping_indexes(sm.mapping, ["tag", "line", *PROCESS_FIELDS], df.shape[1])
        if idx["tag"] is None:
            continue

        for r in range(len(df)):
            tag = clean(df.iat[r, idx["tag"]])
            if not tag:
                continue
            inst_type = detect_type(tag)
            if inst_type == "UNKNOWN":
                continue

            def get(field_name):
                i = idx[field_name]
                return clean(df.iat[r, i]) if i is not None else ""

            rows.append(
                InstrumentRow(
                    tag=tag,
                    inst_type=inst_type,
                    line_no=get("line"),
                    p_op=get("p_op"),
                    p_des_max=get("p_des_max"),
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
    ("p_des_max", "P_Design_Max"),
    ("t_op", "T_Oper"),
    ("t_op_max", "T_Oper_Max"),
    ("t_min", "T_Min_Design"),
    ("t_max", "T_Max_Design"),
]

REPORT_HEADERS = (
    ["Line No", "Tag No", "Type"]
    + [h for _, h in PROCESS_FIELD_COLUMNS]
    + ["Source Sheet", "Result", "Remark"]
)

PROCESS_START_COL = 4
FIELD_COL = {f: PROCESS_START_COL + i for i, (f, _) in enumerate(PROCESS_FIELD_COLUMNS)}
SOURCE_SHEET_COL = PROCESS_START_COL + len(PROCESS_FIELD_COLUMNS)
RESULT_COL = SOURCE_SHEET_COL + 1
REMARK_COL = RESULT_COL + 1


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

        # Evaluate every match's PASS/FAIL before writing the MASTER row, so its
        # Result cell can roll up "FAIL" whenever any related Instrument fails -
        # otherwise the MASTER row's blank Result cell drops out of an Excel
        # filter set to "FAIL", hiding the Line context above the failing rows.
        match_results = []
        for inst in matches:
            results = {f: COMPARATORS[f](getattr(line, f), getattr(inst, f)) for f in PROCESS_FIELDS}
            final = "FAIL" if any(results[f] == "FAIL" for f in FIELDS_IN_VERDICT) else "PASS"
            match_results.append((inst, results, final))
        line_result = "FAIL" if any(final == "FAIL" for _, _, final in match_results) else (
            "PASS" if match_results else "")

        ws.cell(excel_row, 1, line.line_no)
        ws.cell(excel_row, 2, "MASTER")
        ws.cell(excel_row, 3, "LINE")
        for f in PROCESS_FIELDS:
            ws.cell(excel_row, FIELD_COL[f], getattr(line, f))
        if line_result:
            ws.cell(excel_row, RESULT_COL, line_result)
        for c in range(1, n_cols + 1):
            if c == REMARK_COL:
                continue
            ws.cell(excel_row, c).fill = GRAY
        excel_row += 1

        if not matches:
            ws.cell(excel_row, 3, "No Related Item")
            for c in range(1, n_cols + 1):
                if c == REMARK_COL:
                    continue
                ws.cell(excel_row, c).fill = YELLOW
            excel_row += 1
            stats.missing_lines += 1
            continue

        for inst, results, final in match_results:
            stats.matched += 1
            if final == "FAIL":
                stats.fail_count += 1

            ws.cell(excel_row, 2, inst.tag)
            ws.cell(excel_row, 3, inst.inst_type)
            for f in PROCESS_FIELDS:
                ws.cell(excel_row, FIELD_COL[f], getattr(inst, f))
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
