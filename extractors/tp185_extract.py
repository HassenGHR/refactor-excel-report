#!/usr/bin/env python3
"""
tp185_extract.py — extract a TP 185 (SONATRACH DP, Hassi Messaoud) Daily
Workover Report into the standard dict shape.

Source layout
-------------
Legacy Excel `.xls` binary (OLE2 Compound Document, produced by an old
copy of Excel on the rig).  Sheet name is the well number itself
("OMN473" in the sample).  Title at G1: "Rapport journalier work over".
Same French SONATRACH DP report FAMILY as the gw29 template (same
AVANCEMENTS / OUTILS / USURES / PARAMETRES row-12 section banner) but a
distinct layout AND a different file format — .xls not .xlsx — so it
needs its own extractor and its own reader library (xlrd, not openpyxl).

The parse_source.py router currently sniffs the OLE2 magic bytes
(D0 CF 11 E0…) as "doc" and dispatches to _detect_format_word().  For a
DP-DWR .xls file that dispatch is wrong; the router needs a new branch
for xls (see the note at the end of this module).

Known template quirks
---------------------
* Cell types are inconsistent even within a single column.  Start/end
  times in the ops table appear as TEXT ("00:00") in some rows and as
  DATE floats (0.16666… = 4h/24h) in others.  Duration cells are always
  DATE floats.  All time-like cells are handled uniformly by
  _hours_from_cell().
* The chemicals block V/X/Y/Z has product names but the "Utilisé" and
  "Stock T" columns are (in this sample) entirely blank — the report
  keeps the product list as a template but doesn't record daily
  consumption.  We still emit the list so downstream code can see what
  products the rig tracks.
* Daily tariff totals (R26/S26/T26/U26 for T1/T2/T3/NR "Journalier") may
  disagree with the sum of hours in the operations table by a small
  amount (the sample has T1=22 in the Journalier cell but the ops list
  sums to 23h of T1).  We report what the source says; if you need the
  totals to match exactly, recompute them from the activities list.
* The col-M per-op bill code IS filled in on this template, so the
  _assign_bill_codes() back-fill inherited from tp173_extract.py is a
  safe no-op here; kept anyway for robustness against future reports
  where the operator forgets to fill it.
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import xlrd


# ---------------------------------------------------------------------------
# Type-agnostic cell helpers.  xlrd surfaces cells as ctypes rather than
# native Python values (int/float/datetime/etc), so we normalize here.
# ---------------------------------------------------------------------------
def _clean(v) -> str:
    if v is None: return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _cell_value(ws, r: int, c: int, wb, lookup) -> Any:
    """Return the "logical" value of a (row, col) — respecting merges — as
    a native Python object (str, float, datetime, or None).  Handles the
    xlrd XL_CELL_DATE ambiguity: cells with a value in [-1, 1] are
    treated as time-of-day and returned as float days (so 0.25 == 6:00
    == 6 hours), while true dates go back as datetime objects."""
    if r < 0 or r >= ws.nrows or c < 0 or c >= ws.ncols:
        return None
    cell = ws.cell(r, c)
    ct = cell.ctype
    if ct == xlrd.XL_CELL_EMPTY or ct == xlrd.XL_CELL_BLANK:
        # Fall back to merge lookup — the cell may be an interior cell of
        # a merged range whose value lives at the top-left origin.
        return lookup.get((r, c))
    if ct == xlrd.XL_CELL_DATE:
        # Return the raw float — see _hours_from_cell / _date_from_cell
        # for how it's later interpreted.
        return float(cell.value)
    if ct == xlrd.XL_CELL_ERROR:
        # xlrd sometimes exposes cached values (e.g. 23) alongside an
        # error type — pass the value through as-is so downstream numeric
        # coercion still works.
        return cell.value
    if ct == xlrd.XL_CELL_NUMBER:
        return float(cell.value)
    if ct == xlrd.XL_CELL_BOOLEAN:
        return bool(cell.value)
    return cell.value


def _build_merged_lookup(ws) -> Dict[Tuple[int, int], Any]:
    """Map every interior (row, col) of a merged range to the value at
    that range's top-left cell.  Uses xlrd's half-open (rlo, rhi, clo,
    chi) tuple layout.  0-indexed to match ws.cell()."""
    lookup: Dict[Tuple[int, int], Any] = {}
    for rlo, rhi, clo, chi in ws.merged_cells:
        origin_cell = ws.cell(rlo, clo)
        # Materialize the origin's Python value once, so all interior
        # cells share the same normalized representation.
        if origin_cell.ctype == xlrd.XL_CELL_DATE:
            val: Any = float(origin_cell.value)
        elif origin_cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
            val = None
        else:
            val = origin_cell.value
        if val is None:
            continue
        for r in range(rlo, rhi):
            for c in range(clo, chi):
                if (r, c) != (rlo, clo):
                    lookup[(r, c)] = val
    return lookup


def _build_span_end_cols(ws) -> Dict[Tuple[int, int], int]:
    """Map every (row, col) inside a merged range to that range's LAST
    column index (inclusive).  Callers use this to skip past a wide
    label's own merge in one step instead of accidentally re-reading the
    label text column-by-column.  Same shape as the helper of the same
    name in tp215_extract.py / enf30_extract.py."""
    spans: Dict[Tuple[int, int], int] = {}
    for rlo, rhi, clo, chi in ws.merged_cells:
        for r in range(rlo, rhi):
            for c in range(clo, chi):
                # xlrd chi is half-open — subtract 1 to get inclusive end
                spans[(r, c)] = chi - 1
    return spans


def _hours_from_cell(v) -> float:
    """Convert a "time-like" cell value to a duration in hours.

    xlrd stores Excel time-of-day cells as a float in [0, 1] where 1.0 ==
    24 hours.  Some cells are TEXT ("04:00") instead of DATE; some are
    plain numbers (e.g. 19.0 in an ERROR-typed cell).  Cover all three."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        # xlrd float in [0, 1] → hours; larger values are already hours.
        if 0 <= float(v) <= 1.5:
            return float(v) * 24.0
        return float(v)
    if isinstance(v, datetime):
        # Reachable if a caller pre-converts an xlrd date to datetime.
        return v.hour + v.minute / 60.0 + v.second / 3600.0
    if isinstance(v, time):
        return v.hour + v.minute / 60.0
    if isinstance(v, timedelta):
        return v.total_seconds() / 3600.0
    s = _clean(v)
    if not s: return 0.0
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$", s)
    if m:
        h = int(m.group(1)); mi = int(m.group(2))
        se = int(m.group(3) or 0)
        return h + mi / 60.0 + se / 3600.0
    try:
        f = float(s.replace(",", "."))
        if 0 <= f <= 1.5:
            return f * 24.0
        return f
    except ValueError:
        return 0.0


def _time_from_cell(v) -> Optional[time]:
    """Convert a time-like cell to a datetime.time (24:00 → 00:00)."""
    if v is None: return None
    if isinstance(v, time): return v
    if isinstance(v, datetime): return v.time()
    if isinstance(v, timedelta):
        total = int(v.total_seconds())
        if total == 86400: return time(0, 0)
        return time((total // 3600) % 24, (total % 3600) // 60)
    if isinstance(v, (int, float)):
        # xlrd float [0, 1] → time-of-day
        f = float(v)
        if f >= 1.0:  # 24:00 or beyond
            return time(0, 0)
        total_min = int(round(f * 24 * 60))
        return time((total_min // 60) % 24, total_min % 60)
    s = _clean(v)
    if not s or s in ("-", "None"): return None
    if s in ("24:00", "24:00:00"): return time(0, 0)
    for fmt in ("%H:%M:%S", "%H:%M", "%Hh%M"):
        try: return datetime.strptime(s, fmt).time()
        except ValueError: continue
    return None


def _float(v, default: float = 0.0) -> float:
    if v is None: return default
    if isinstance(v, bool): return float(v)
    if isinstance(v, (int, float)): return float(v)
    s = _clean(v).replace(",", ".").replace(" ", "")
    if not s or s in ("-", "/", "None"): return default
    stripped = re.sub(r"[a-zA-Zé°%/³]+$", "", s).strip()
    try: return float(stripped)
    except ValueError: pass
    m = re.match(r"^-?\d+(?:\.\d+)?", s)
    if m:
        try: return float(m.group(0))
        except ValueError: pass
    return default


def _int(v, default: int = 0) -> int:
    return int(_float(v, float(default)))


def _date_from_cell(v, wb) -> Optional[date_type]:
    """Try to read a cell value as a calendar date.  xlrd date floats > 1
    are true dates; convert via xldate_as_datetime.  Strings are parsed
    via strptime.  Returns None for anything that doesn't look like a
    real date (times, empty, placeholders)."""
    if v is None: return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date_type): return v
    if isinstance(v, (int, float)):
        # Fraction of a day (time-of-day) → not a real date
        if 0 <= float(v) < 1.5:
            return None
        try:
            dt = xlrd.xldate.xldate_as_datetime(float(v), wb.datemode)
            return dt.date()
        except Exception:
            return None
    s = _clean(v)
    if not s or s in ("-", "/", "//"): return None
    if "--/--/" in s or "/--/" in s: return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: continue
    # Last-ditch: pull a "dd/mm/yyyy" out of a longer string
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b", s)
    if m:
        d, mo, y = (int(x) for x in m.groups())
        if y < 100: y += 2000
        try: return date_type(y, mo, d)
        except ValueError: return None
    return None


def _strip_prefix(text: str, *prefixes: str) -> str:
    """Remove any leading "LABEL :" prefix from `text`."""
    s = _clean(text)
    for p in prefixes:
        rx = rf"^\s*{re.escape(p)}\s*:?\s*"
        new = re.sub(rx, "", s, flags=re.IGNORECASE)
        if new != s: return new.strip()
    return s


def _find_row(ws, wb, lookup, keywords, row_range, col_range):
    """Locate the first (row, col) in the given ranges whose cleaned
    upper-case text contains any of the given keyword substrings.
    Ranges are half-open [start, end) in 0-indexed row/col space."""
    for r in range(row_range[0], min(row_range[1], ws.nrows)):
        for c in range(col_range[0], min(col_range[1], ws.ncols)):
            v = _cell_value(ws, r, c, wb, lookup)
            if v is None: continue
            txt = _clean(str(v)).upper()
            for kw in keywords:
                if kw in txt:
                    return (r, c)
    return None


def _value_after(ws, wb, lookup, spans, row, after_col, max_scan=6,
                 reject_labels=True):
    """Return the first non-empty cell strictly right of `after_col` on
    `row`, skipping past `after_col`'s own merged span so we don't just
    re-read the label text again from an interior cell of its merge.

    When `reject_labels=True`, cells whose value looks like a label
    ("Puits :", "Champ :", "N°", etc.) are also skipped — those are
    common false positives when a label is followed by ANOTHER label
    before the actual value cell."""
    start = spans.get((row, after_col), after_col)
    for c in range(start + 1, start + 1 + max_scan):
        v = _cell_value(ws, row, c, wb, lookup)
        if v is None or _clean(str(v)) == "":
            continue
        if reject_labels and isinstance(v, str):
            s = _clean(v)
            # Bare "Xxxx :" labels — trailing colon, no other punctuation
            if re.match(r"^[A-Za-zÀ-ÿ\s\.']+\s*:\s*$", s):
                continue
        return v
    return None


def _numeric_value_after(ws, wb, lookup, spans, row, after_col, max_scan=6):
    """Like _value_after, but only accepts a cell that parses as a number
    (skips over stray label text bleeding in from further merges)."""
    start = spans.get((row, after_col), after_col)
    for c in range(start + 1, start + 1 + max_scan):
        v = _cell_value(ws, row, c, wb, lookup)
        if v is None:
            continue
        if isinstance(v, (int, float)):
            return v
        s = _clean(str(v))
        if s and re.match(r"^-?[\d.,]+$", s):
            return v
    return None


def _text_value_after(ws, wb, lookup, spans, row, after_col, max_scan=6):
    """Like _value_after, but only accepts a cell that is a non-numeric,
    non-label string (e.g. a well name, a person's name, a rig name).
    Numbers and bare "Xxxx :" labels are skipped."""
    start = spans.get((row, after_col), after_col)
    for c in range(start + 1, start + 1 + max_scan):
        v = _cell_value(ws, row, c, wb, lookup)
        if v is None: continue
        if isinstance(v, (int, float, datetime, date_type, time, timedelta)):
            continue
        s = _clean(str(v))
        if not s: continue
        if re.match(r"^-?[\d.,]+$", s): continue
        if re.match(r"^[A-Za-zÀ-ÿ\s\.']+\s*:\s*$", s):
            continue  # bare label
        return s
    return None


# ---------------------------------------------------------------------------
# Bill-code back-assignment (inherited from tp173_extract.py — same logic,
# in case a future report on this template leaves col M blank).  See that
# module's _assign_bill_codes() docstring for the full rationale.
# ---------------------------------------------------------------------------
def _assign_bill_codes(activities, tarif_totals):
    if not tarif_totals or not activities:
        return activities
    DAILY_KEYS = ["t1", "t2", "t3", "t4", "t0", "nr"]
    remaining = {code: float(tarif_totals.get(code, 0.0)) for code in DAILY_KEYS}
    for op in activities:
        code = (op.get("bill") or "").lower()
        if code in remaining:
            remaining[code] -= _hours_from_cell(op.get("hours"))
    bucket_order = [c for c in DAILY_KEYS if remaining.get(c, 0.0) > 1e-6]
    if not bucket_order:
        return activities
    bi = 0
    for op in activities:
        if op.get("bill"):
            continue
        while bi < len(bucket_order) and remaining[bucket_order[bi]] <= 1e-6:
            bi += 1
        if bi >= len(bucket_order):
            break
        code = bucket_order[bi]
        op["bill"] = code.upper()
        remaining[code] -= _hours_from_cell(op.get("hours"))
    return activities


# ---------------------------------------------------------------------------
# Main extractor.  All row/col constants below are 0-INDEXED (xlrd
# convention), unlike openpyxl.  Comments show A1-style coordinates for
# readability, e.g. "row 7 col 18" == "S8" in the spreadsheet.
# ---------------------------------------------------------------------------
def parse_tp185(source: Union[Path, str, BytesIO]) -> dict:
    if isinstance(source, (str, Path)):
        wb = xlrd.open_workbook(str(source), formatting_info=True)
    else:
        # xlrd needs bytes when reading from memory
        data = source.read()
        wb = xlrd.open_workbook(file_contents=data, formatting_info=True)

    # The rig writes the well number as the sheet name; pick the first
    # (and typically only) data sheet, skipping empty template sheets.
    ws = None
    for name in wb.sheet_names():
        candidate = wb.sheet_by_name(name)
        if candidate.nrows > 5:
            ws = candidate
            break
    if ws is None:
        ws = wb.sheet_by_index(0)

    L = _build_merged_lookup(ws)
    S = _build_span_end_cols(ws)

    def C(r, c):  # short alias
        return _cell_value(ws, r, c, wb, L)

    header: Dict[str, Any] = {}

    # =====================================================================
    # HEADER  (rows 1-11 in A1 notation → rows 0-10 here)
    #
    # Layout of the top band:
    #   R1:  "Division production"        "Rapport journalier work over"
    #        "DU :"  <date>  "N°"  <day-number>
    #   R8:  "Puits :"  <well>   "Champ :"  ...  "Appareil :"  <rig>
    #        "Sabot dernier Tubage 7\":"  <depth>   " liner 7\""  <depth>
    #   R10: field name value, "Fond initial", "type", <mud-type>
    #
    # We use dynamic label search rather than hard-coded rows so that a
    # ±1-row shift on another report doesn't lose the field.
    # =====================================================================
    date_pos = _find_row(ws, wb, L, ("DU :", "DU:"), (0, 8), (0, 25))
    if date_pos:
        v = _value_after(ws, wb, L, S, date_pos[0], date_pos[1])
        d = _date_from_cell(v, wb)
        if d:
            header["date"] = d

    dayn_pos = _find_row(ws, wb, L, ("N°",), (0, 8), (16, 28))
    if dayn_pos:
        v = _numeric_value_after(ws, wb, L, S, dayn_pos[0], dayn_pos[1])
        if v is not None:
            header["day_number"] = _int(v)

    well_pos = _find_row(ws, wb, L, ("PUITS",), (5, 12), (0, 8))
    if well_pos:
        v = _text_value_after(ws, wb, L, S, well_pos[0], well_pos[1])
        if v: header["well_name"] = v

    rig_pos = _find_row(ws, wb, L, ("APPAREIL",), (5, 12), (0, 20))
    if rig_pos:
        v = _text_value_after(ws, wb, L, S, rig_pos[0], rig_pos[1])
        if v: header["rig_name"] = v

    champ_pos = _find_row(ws, wb, L, ("CHAMP",), (5, 12), (0, 20))
    if champ_pos:
        # "Champ :" value lives in a cell merged across several ROWS
        # BELOW the label (G10[I11] in the sample) — NOT on the same row.
        # On the same row there's a "Appareil :" label and then the rig
        # name ("TP 185"), which we must NOT pick up as the field.  We
        # constrain the search to columns strictly BEFORE the "Appareil"
        # label to keep the two apart, and try 2 rows down first (where
        # the value's merge actually starts on this template).
        app_col = None
        app_search = _find_row(ws, wb, L, ("APPAREIL",),
                                (champ_pos[0], champ_pos[0] + 1),
                                (champ_pos[1] + 1, 25))
        if app_search:
            app_col = app_search[1]
        for dr in (2, 1, 0):    # bottom-up: value most likely 2 rows down
            r_scan = champ_pos[0] + dr
            start_c = S.get((r_scan, champ_pos[1]), champ_pos[1]) + 1
            hi_c = app_col if app_col is not None else start_c + 6
            found = None
            for c in range(start_c, hi_c):
                v = C(r_scan, c)
                if v is None: continue
                if isinstance(v, (int, float, datetime, date_type, time, timedelta)):
                    continue
                s = _clean(str(v))
                if not s: continue
                if re.match(r"^-?[\d.,]+$", s): continue
                if re.match(r"^[A-Za-zÀ-ÿ\s\.']+\s*:\s*$", s): continue
                found = s
                break
            if found:
                header["field_name"] = found
                break

    # Last casing shoe — label "Sabot dernier Tubage 7\"" with the depth
    # to its right.  The label itself embeds the casing size.
    sabot_pos = _find_row(ws, wb, L, ("SABOT",), (5, 12), (0, 20))
    if sabot_pos:
        sabot_label = _clean(str(C(*sabot_pos)))
        size_m = re.search(r"(\d+[\"'](?:\s*\d+/\d+)?)", sabot_label)
        v = _numeric_value_after(ws, wb, L, S, sabot_pos[0], sabot_pos[1])
        if v is not None:
            depth = _float(v)
            if depth > 0:
                sz = size_m.group(1) if size_m else ""
                header["top_shoe"] = f'{sz} @ {int(depth)}m' if sz else f'@ {int(depth)}m'
                header["top_shoe_depth"] = depth

    # Liner 7" — label sits mid-row to the right of the shoe cell.  The
    # label cell IS the value in this template ("liner 7\"" + a depth
    # cell right after), so grab both parts.
    liner_pos = _find_row(ws, wb, L, ("LINER",), (5, 12), (0, 25))
    if liner_pos:
        liner_label = _clean(str(C(*liner_pos)))
        liner_size_m = re.search(r"(\d+[\"'](?:\s*\d+/\d+)?)", liner_label)
        depth_v = _value_after(ws, wb, L, S, liner_pos[0], liner_pos[1])
        if depth_v is not None:
            depth_s = _clean(str(depth_v))
            sz = liner_size_m.group(1) if liner_size_m else "liner"
            header["top_liner"] = f'{sz} @ {depth_s}' if depth_s else sz
        elif liner_size_m:
            header["top_liner"] = liner_label

    # Mud type — the mud-header block has "Fond intial" ... "type" ...
    # <mud-type> in a row of tightly-packed labels+values.  We want the
    # cell after "type", skipping the "type" label's own merge.
    type_pos = _find_row(ws, wb, L, ("TYPE",), (8, 13), (18, 28))
    if type_pos:
        v = _text_value_after(ws, wb, L, S, type_pos[0], type_pos[1])
        if v: header["mud_type"] = v

    # BHA description — under the "BHA" merged label in the AVANCEMENTS
    # block.  The value lives in a merged cell 2-3 rows below the label.
    bha_pos = _find_row(ws, wb, L, ("BHA",), (18, 25), (0, 8))
    if bha_pos:
        for dr in (1, 2, 3):
            v = C(bha_pos[0] + dr, 0)
            if v is not None:
                s = _clean(str(v))
                if s and "BHA" not in s.upper() and not re.match(r"^[A-Za-z\s\.']+\s*:\s*$", s):
                    header["bha_details"] = s
                    break

    # Tige (drill pipe) and DC specs — combined into a single string.
    # The row has 3 semantically meaningful fields (OD, ID, grade) after
    # the label, then the next SECTION header ("MATERIELS DE FOND",
    # "Type", etc.) bleeds in from a wide merged cell further right.  We
    # cap at 3 values AND stop at anything that looks like a header
    # (all-caps or ends-with-colon) to avoid picking that up.
    def _collect_row_values(label_pos, max_values=3):
        r, lc = label_pos
        start_c = S.get(label_pos, lc) + 1
        parts, last = [], None
        for c in range(start_c, start_c + 10):
            v = C(r, c)
            if v is None:
                # First empty cell AFTER we've collected something = the
                # end of this row's content.  Anything further right is
                # the next block's header/data bleeding in.
                if parts:
                    break
                continue
            s = _clean(str(v))
            if not s:
                if parts: break
                continue
            if s == last: continue                 # dedupe consecutive merges
            # Stop at obvious section-header/label bleed
            if re.match(r"^[A-Za-zÀ-ÿ\s\.']+\s*:\s*$", s):
                break
            if s.isupper() and len(s) > 4:         # ALL-CAPS block header
                break
            parts.append(s); last = s
            if len(parts) >= max_values:
                break
        return parts

    tige_pos = _find_row(ws, wb, L, ("TIGE",), (15, 22), (0, 8))
    if tige_pos:
        parts = _collect_row_values(tige_pos, max_values=3)
        if parts:
            header["drill_pipe"] = " / ".join(parts)
    dc_pos = _find_row(ws, wb, L, ("D C ",), (15, 22), (0, 8))
    if dc_pos:
        parts = _collect_row_values(dc_pos, max_values=3)
        if parts:
            header["drill_collar"] = " / ".join(parts)

    # BOP test date — "Dernière date test BOP " label near row 45.
    bop_pos = _find_row(ws, wb, L, ("DERNI", "BOP"), (40, 50), (14, 22))
    if bop_pos:
        v = _value_after(ws, wb, L, S, bop_pos[0], bop_pos[1])
        d = _date_from_cell(v, wb)
        if d: header["bop_test"] = d

    # =====================================================================
    # OPERATIONS  (typically rows 28..38 in A1 → 27..37 here)
    #   A: start time  |  B: end time  |  C..L: description (wide merge)
    #   M: bill code   |  N: duration hours
    #
    # We locate the header row dynamically by scanning for "DE" in col A
    # near "Analyse des temps et des opérations" so we survive the ±1-row
    # drift I saw between other templates in this family.
    # =====================================================================
    activities: List[Dict[str, Any]] = []

    ops_hdr = _find_row(ws, wb, L, ("ANALYSE DES TEMPS", "ANALYSE DES OP"),
                        (20, 40), (0, 20))
    ops_start = (ops_hdr[0] + 2) if ops_hdr else 27  # header + "DE/A" row
    # But the "DE" row itself is at ops_hdr[0]+1 in this template; scan
    # forward from there to skip label rows.
    if ops_hdr is not None:
        for r in range(ops_hdr[0] + 1, min(ops_hdr[0] + 5, ws.nrows)):
            v_a = C(r, 0)
            v_b = C(r, 1)
            if isinstance(v_a, str) and _clean(v_a).upper() == "DE":
                ops_start = r + 1
                break

    ops_max_row = min(ops_start + 20, ws.nrows)

    for r in range(ops_start, ops_max_row):
        # Stop at obvious block terminators
        row_text = " ".join(_clean(C(r, c) or "") for c in range(0, 15)).upper()
        if "TOTAL" in _clean(C(r, 12) or "").upper():
            # Reached the "Total" summary row (col M) — end of ops
            break
        if ("SITUATION" in row_text and "MINUIT" in row_text) or "MINUIT" in row_text:
            break

        a_val = C(r, 0)   # start
        b_val = C(r, 1)   # end
        desc  = _clean(C(r, 2) or "")
        bill  = _clean(C(r, 12) or "")   # M
        hours_cell = C(r, 13)             # N

        # Skip fully blank rows
        if (a_val is None and b_val is None and not desc and not bill
                and hours_cell is None):
            continue
        if not desc and not bill and _hours_from_cell(hours_cell) < 1e-6:
            # Row is present in the template but the operator left it empty
            continue

        start_t = _time_from_cell(a_val)
        end_t   = _time_from_cell(b_val)

        # Description-only continuation (no times, no code, no hours) →
        # fold into the previous op.
        if (start_t is None and end_t is None and not bill
                and _hours_from_cell(hours_cell) < 1e-6 and desc):
            if activities:
                activities[-1]["description"] = (
                    activities[-1]["description"] + "\n" + desc
                ).strip()
            continue

        # Prefer the explicit duration cell (col N) when present, else
        # compute from start/end (handling wrap over midnight).
        hours = _hours_from_cell(hours_cell)
        if hours == 0.0 and start_t and end_t:
            sm = start_t.hour * 60 + start_t.minute
            em = end_t.hour * 60 + end_t.minute
            if em == sm: hours = 24.0
            elif em > sm: hours = (em - sm) / 60.0
            else:         hours = (em + 1440 - sm) / 60.0

        activities.append({
            "start_time": start_t,
            "end_time":   end_t,
            "hours":      hours,
            "phase_name": "",
            "code": "", "sub": "",
            "description": desc,
            "start_md": 0, "end_md": 0,
            "npt": 0, "npt_detail": "",
            "npt_company": "", "op_company": "",
            "bill": bill,
        })

    # =====================================================================
    # TARIFF TOTALS  (row 26 in A1: "Journalier" label + T1/T2/T3/NR cells
    #                 row 27 in A1: "Cumul" label + cumulative cells)
    # =====================================================================
    tarif_totals: Dict[str, float] = {}
    # Header row with T1/T2/T3/NR labels
    tar_hdr = _find_row(ws, wb, L, ("T1",), (20, 30), (14, 25))
    if tar_hdr:
        hr, hc = tar_hdr
        # Read the 4 code labels off the header row
        codes: List[str] = []
        for c in range(hc, hc + 6):
            v = C(hr, c)
            if isinstance(v, str) and re.match(r"^(T\d|NR|T0)$", _clean(v)):
                codes.append(_clean(v).lower())
            else:
                # Stop at first non-code cell so we don't drift into LGS/ClNa
                break
        # Journalier row is 1 below the header; Cumul is 2 below.  Values
        # sit at the same columns as the code labels.
        for dr, prefix in ((1, ""), (2, "cum_")):
            for i, code in enumerate(codes):
                v = C(hr + dr, hc + i)
                if v is None or _clean(str(v)) == "":
                    continue
                hrs = _hours_from_cell(v) if dr == 1 else _float(v)
                tarif_totals[f"{prefix}{code}"] = hrs

    # =====================================================================
    # MUD CHECKS  (rows 19-25 in A1 → col V/W left pair + col X/Y right pair)
    # =====================================================================
    mud_checks: Dict[str, Any] = {}
    if "mud_type" in header:
        mud_checks["mud_type"] = header["mud_type"]

    # Map of (label substring, key) — both left and right label columns
    # feed the same dict, since Excel puts labels on both sides of the
    # narrow property block.
    MUD_LABELS = {
        "DENSIT":       "density",
        "V.MARSH":      "fun_vis",
        "VMARSH":       "fun_vis",
        "Y.P":          "yp",
        "YP":           "yp",
        "FILTRAT":      "filtrat",
        "SABL":         "sand_pct",
        "SOLIDE":       "solid_pct",
        "GEL 10/S":     "gel10sec",
        "GEL 10/M":     "gel10m",
        "P V":          "pv",
        "PV":           "pv",
        "PH":           "ph",
        "LGS":          "lgs",
        "CLNA":         "cl_na",
        "EAU":          "water_pct",
    }
    # Scan the mud property strip.  Labels appear in cols V or X (21, 23);
    # values in the immediate next col (W=22 or Y=24).
    for r in range(18, 26):
        for label_c, value_c in ((21, 22), (23, 24)):
            label = _clean(C(r, label_c) or "")
            if not label:
                continue
            up = label.upper()
            for kw, key in MUD_LABELS.items():
                if kw in up:
                    v = C(r, value_c)
                    if v is None or _clean(str(v)) == "":
                        break
                    if key in ("cl_na",):
                        mud_checks[key] = _clean(str(v))
                    else:
                        mud_checks[key] = _float(v)
                    break

    # =====================================================================
    # MUD VOLUMES  (col V/W/X/Y in rows 12-18 A1 → same-shape label/value
    # pairs as mud checks, but for volumes rather than properties)
    # =====================================================================
    mud_volume: Dict[str, Any] = {}
    VOLUME_LABELS = {
        "FABRICATION":       "fabrication_vol",
        "STOCKAGE D":        "water_storage_vol",
        "PERTE FORMATION":   "formation_loss",
        "PERTE ACCID":       "accident_loss",
        "EJECTE BOUE":       "dumped_volume",
        "VOL PUITS":         "string_volume",
        "VOL SURFACE":       "pits_volume",
        "VOL. PUITS":        "string_volume",
        "VOL. SURFACE":      "pits_volume",
    }
    for r in range(11, 19):
        for label_c, value_c in ((21, 22), (21, 24), (23, 24)):
            label = _clean(C(r, label_c) or "")
            if not label:
                continue
            up = label.upper()
            for kw, key in VOLUME_LABELS.items():
                if kw in up:
                    v = C(r, value_c)
                    if v is None or _clean(str(v)) == "":
                        break
                    mud_volume[key] = _float(v)
                    break

    # =====================================================================
    # CHEMICALS  (rows 28-38 A1 → col V/W name, X used, Y stock)
    # We emit the product list even if usage values are blank, because
    # downstream code may still want to know which products the rig
    # tracks.
    # =====================================================================
    chemicals: List[Dict[str, str]] = []
    prod_hdr = _find_row(ws, wb, L, ("PRODUITS",), (23, 30), (20, 27))
    if prod_hdr:
        for r in range(prod_hdr[0] + 2, min(prod_hdr[0] + 15, ws.nrows)):
            name = _clean(C(r, 21) or "")
            if not name:
                continue
            up = name.upper()
            # Stop the block cleanly when we hit the next section
            if "PERSONNEL" in up or "TOTAL" in up:
                break
            used = C(r, 23)
            stock = C(r, 24)
            if (used is None or _clean(str(used)) == "") and \
               (stock is None or _clean(str(stock)) == ""):
                # Still keep the name so consumers can see the product list
                chemicals.append({
                    "item": name, "units": "",
                    "received": "", "used": "", "on_loc": "",
                })
                continue
            chemicals.append({
                "item":     name,
                "units":    "",
                "received": "",
                "used":     _clean(str(used)) if used is not None else "",
                "on_loc":   _clean(str(stock)) if stock is not None else "",
            })

    # =====================================================================
    # PERSONNEL  (rows 40-46 in A1 → col V/W label, Y count).
    # The list ends with a "TOTAL" row we skip.
    # =====================================================================
    personnel: List[Dict[str, Any]] = []
    pers_hdr = _find_row(ws, wb, L, ("PERSONNEL",), (35, 45), (20, 28))
    if pers_hdr:
        for r in range(pers_hdr[0] + 1, min(pers_hdr[0] + 10, ws.nrows)):
            role = _clean(C(r, 21) or "")
            if not role:
                continue
            up = role.upper()
            if up.startswith("TOTAL"):
                total_count = C(r, 24)
                if total_count is not None:
                    header["personnel_total"] = _int(total_count)
                break
            count = C(r, 24)
            if count is None or _clean(str(count)) == "":
                continue
            personnel.append({
                "company": role,
                "number":  _int(count),
                "hours":   "",
                "names":   "",
            })

    # =====================================================================
    # TEXT SECTIONS
    #   "Situation après minuit :" (label row) + value on the row below
    #   "Situation au rapport"     (col A label) + value in col D
    #   "Programme prévu"          (col A label) + value in col D
    # =====================================================================
    text_sections: Dict[str, str] = {}

    am_pos = _find_row(ws, wb, L, ("MINUIT",), (30, 45), (0, 15))
    if am_pos:
        # Value sits on the NEXT row (or two down), spanning C..L merged.
        # Read col C (index 2) directly since the value's merge origin is
        # at (am_row+1, 2).
        for dr in (1, 2):
            v = C(am_pos[0] + dr, 2)
            if v is not None:
                s = _clean(str(v))
                if s and "MINUIT" not in s.upper():
                    text_sections["after_midnight"] = s
                    break

    # "Situation au rapport" appears TWICE on this sheet: once as a
    # label-only cell at C46 (merged C:L, no value on its row), and once
    # as a label-and-value pair at A47 (with the situation text in D47).
    # We want the second one — so restrict the label search to col A
    # only, which is where the "with value" version sits.
    sit_pos = _find_row(ws, wb, L, ("SITUATION AU RAPPORT",), (40, 50), (0, 1))
    if sit_pos:
        v = _text_value_after(ws, wb, L, S, sit_pos[0], sit_pos[1], max_scan=10)
        if v:
            text_sections["current_operation"] = v
            text_sections["day_summary"] = v

    plan_pos = _find_row(ws, wb, L, ("PROGRAMME",), (40, 52), (0, 5))
    if plan_pos:
        v = _text_value_after(ws, wb, L, S, plan_pos[0], plan_pos[1], max_scan=10)
        # Reject stray "CODE" tokens sitting to the right on this row
        if v and "CODE" not in v.upper():
            text_sections["plan_operations"] = v

    # =====================================================================
    # SUPERVISOR — the "N° SH/DP" label sits at S46 with the rep's name
    # in V47 (merged Y47) on the row below.  We use _text_value_after to
    # guarantee we skip numeric cells (personnel counts) that would
    # otherwise get picked up first.
    # =====================================================================
    sup_pos = _find_row(ws, wb, L, ("N° SH", "SH/DP", "SH / DP"),
                         (40, 50), (14, 25))
    if sup_pos:
        REJECT_KWS = ("SH/DP", "SH / DP", "PERSONNEL", "N° SH", "TOTAL",
                      "TYPE", "EN LOCATION", "VEHICULES")
        for dr in (0, 1, 2):
            v = _text_value_after(ws, wb, L, S, sup_pos[0] + dr,
                                   sup_pos[1], max_scan=10)
            if not v: continue
            up = v.upper()
            if any(kw in up for kw in REJECT_KWS):
                continue
            header["supervisor"] = v
            break

    # =====================================================================
    # BACK-FILL BILL CODES (no-op if the source already tagged every op)
    # =====================================================================
    activities = _assign_bill_codes(activities, tarif_totals)

    return {
        "header":              header,
        "activities":          activities,
        "text_sections":       text_sections,
        "mud_checks":          mud_checks,
        "mud_volume":          mud_volume,
        "mud_chemical_usage":  chemicals,
        "personnel_data":      personnel,
        "pumps":               [],
        "well_location":       {},
        "survey_data":         [],
        "safety":              {},
        "tarif_totals":        tarif_totals,
    }


# Drop-in compat with the rest of the extractor set
parse_daily_excel_report = parse_tp185


# ---------------------------------------------------------------------------
# NOTE ON parse_source.py INTEGRATION
# ---------------------------------------------------------------------------
# The current parse_source.py sniffs OLE2 magic bytes (D0 CF 11 E0…) as
# "doc" and dispatches to _detect_format_word().  That will fail for a
# .xls file — python-docx can't open .xls.  To wire this extractor in,
# add BOTH of these to parse_source.py:
#
#   1. In _sniff_kind(), branch on the file extension for OLE2 headers:
#
#        if head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
#            if isinstance(source, (str, Path)):
#                suffix = Path(source).suffix.lower()
#                if suffix == ".xls":
#                    return "xls"
#            return "doc"
#
#   2. Add a new _detect_format_xls(source) that opens with xlrd, scans
#      sheet name + first 12 rows for markers, and returns "tp185" when
#      it sees "TP 185" / "Rapport journalier work over" / "OMN".  Wire
#      it into _detect_format() alongside pdf/xlsx/word branches.
#
#   3. In parse_source(), add:
#        elif fmt == "tp185":
#            from extractors.tp185_extract import parse_tp185
#            data = parse_tp185(source)
# ---------------------------------------------------------------------------


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: tp185_extract.py SOURCE.xls")
    data = parse_tp185(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
