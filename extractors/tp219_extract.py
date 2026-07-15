#!/usr/bin/env python3
"""
entp219_extract.py — extract an ENTP219 (SONATRACH DP, Hassi Messaoud)
Daily Work-Over Report into the standard dict shape, following the same
conventions as tp185_extract.py / tp173_extract.py.

Source layout
-------------
Modern .xlsx, single sheet named "DWR N°1".  Title at J3: "daily
work-over report".  Same SHDP report family as tp185_extract.py (same
AVANCEMENTS / OUTILS / PARAMETRES / USURES row-8 section banner) but
English-labeled header fields (WELL / RIG / LAST CSG SHOS / TOL / TOP
LINER) and a cleaner, mostly same-row label/value layout — no merged
"label spans 3 rows down to its value" quirk like tp185's Champ/Puits
block.

Known layout notes
-------------------
* The ops table (rows 27+ in the sample) has FIVE relevant columns:
  C=start, D=end, E=description (wide merge to O), P=bill code,
  Q=hours.  Columns A and B, immediately to the LEFT of the real start/
  end columns, hold an entirely unrelated set of decimal-hour numbers
  that do NOT correspond to the same operations (verified against the
  sample: e.g. row 29 has real times C=03:00/D=12:00 but A=5/B=14, which
  don't match). We deliberately ignore columns A/B.
* Start/end cells are inconsistently typed — sometimes a timedelta,
  sometimes a text string like " 14:00" or " 24:00" (with a leading
  space). _time_from_cell() below normalizes all of these.
* "After midnight" is not a separate section — it's a single description
  cell (e.g. row 45) whose text is *prefixed* with "After midnight :".
  We detect that prefix directly rather than treating it as a stop
  marker for a whole new sub-table (there isn't one here — the rows
  between the last real op and the after-midnight line are all blank or
  belong to the unrelated A/B helper column).
* "Current Status" / "24 H Forecast" / "REQUIREMENTS" / "REMARKS" are
  clean same-row label (col C, merged to E) / value (col F, merged to P)
  pairs — no multi-row lookups needed, unlike tp185's "Situation"
  fields.
* Per-operation bill codes ARE filled in on this template (col P), so
  the _assign_bill_codes() back-fill (inherited from tp173_extract.py)
  is normally a no-op here; kept for robustness in case a future report
  leaves it blank.
* Personnel and mud-volume blocks use a "label spans 3 cols, value 1 col
  further" layout (e.g. X46:Z46 label, AA46 value) — handled with the
  same merge-span-skip helpers used in tp185_extract.py /
  tp215_extract.py.
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _clean(v) -> str:
    if v is None: return ""
    return re.sub(r"\s+", " ", str(v)).strip()


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


def _time_from_cell(v) -> Optional[time]:
    """Normalize a start/end cell to datetime.time.  Cells on this
    template appear as timedelta, text like ' 14:00' / ' 24:00' (leading
    space), or occasionally a bare datetime."""
    if v is None: return None
    if isinstance(v, time): return v
    if isinstance(v, datetime): return v.time()
    if isinstance(v, timedelta):
        total = int(v.total_seconds())
        if total >= 86400: return time(0, 0)
        return time((total // 3600) % 24, (total % 3600) // 60)
    s = _clean(v)
    if not s or s in ("-", "None"): return None
    if s in ("24:00", "24:00:00"): return time(0, 0)
    for fmt in ("%H:%M:%S", "%H:%M", "%Hh%M"):
        try: return datetime.strptime(s, fmt).time()
        except ValueError: continue
    return None


def _hours_from_cell(v) -> float:
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    if isinstance(v, timedelta): return v.total_seconds() / 3600.0
    if isinstance(v, time): return v.hour + v.minute / 60.0
    if isinstance(v, datetime): return v.hour + v.minute / 60.0
    return _float(v)


def _date_parse(v):
    if v is None: return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date_type): return v
    s = _clean(v)
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: continue
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b", s)
    if m:
        d, mo, y = (int(x) for x in m.groups())
        if y < 100: y += 2000
        try: return date_type(y, mo, d)
        except ValueError: return None
    return None


def _build_merged_lookup(ws):
    lookup = {}
    for mr in ws.merged_cells.ranges:
        av = ws.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                if (r, c) != (mr.min_row, mr.min_col):
                    lookup[(r, c)] = av
    return lookup


def _build_span_end_cols(ws):
    spans = {}
    for mr in ws.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                spans[(r, c)] = mr.max_col
    return spans


def _cell(ws, r, c, lookup):
    v = ws.cell(r, c).value
    return v if v is not None else lookup.get((r, c))


def _find_row(ws, lookup, keywords, row_range, col_range):
    for r in range(row_range[0], row_range[1] + 1):
        for c in range(col_range[0], col_range[1] + 1):
            v = _cell(ws, r, c, lookup)
            if v is None: continue
            txt = _clean(str(v)).upper()
            for kw in keywords:
                if kw in txt:
                    return (r, c)
    return None


def _value_after(ws, lookup, row, after_col, max_scan=6, spans=None,
                 reject_labels=True):
    start = after_col
    if spans is not None:
        start = spans.get((row, after_col), after_col)
    for c in range(start + 1, start + 1 + max_scan):
        v = _cell(ws, row, c, lookup)
        if v is None or _clean(str(v)) == "":
            continue
        if reject_labels and isinstance(v, str):
            s = _clean(v)
            if re.match(r"^[A-Za-zÀ-ÿ\s\.'/]+\s*:\s*$", s):
                continue
        return v
    return None


def _numeric_value_after(ws, lookup, row, after_col, max_scan=6, spans=None):
    start = after_col
    if spans is not None:
        start = spans.get((row, after_col), after_col)
    for c in range(start + 1, start + 1 + max_scan):
        v = _cell(ws, row, c, lookup)
        if v is None: continue
        if isinstance(v, (int, float)):
            return v
        s = _clean(str(v))
        if s and re.match(r"^-?[\d.,]+$", s):
            return v
    return None


def _assign_bill_codes(activities, tarif_totals):
    """Same chronological best-effort back-fill used in tp173_extract.py /
    tp185_extract.py — see those modules for the full rationale. No-op
    when every op already has a bill code (the common case here)."""
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
# Main extractor
# ---------------------------------------------------------------------------
def parse_entp219(source: Union[Path, str, BytesIO]) -> dict:
    wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)
    S = _build_span_end_cols(ws)

    header: Dict[str, Any] = {}

    # =====================================================================
    # HEADER  (rows 3-7) — date, report N°, well/rig/casing block.
    # =====================================================================
    date_pos = _find_row(ws, L, ("LA DATE",), (1, 6), (1, 30))
    if date_pos:
        v = _value_after(ws, L, date_pos[0], date_pos[1], spans=S)
        d = _date_parse(v)
        if d: header["date"] = d

    dwr_pos = _find_row(ws, L, ("DWR REPORT",), (1, 6), (1, 30))
    if dwr_pos:
        raw = _clean(_cell(ws, dwr_pos[0], dwr_pos[1], L) or "")
        m = re.search(r"(\d+)", raw)
        if m:
            header["day_number"] = int(m.group(1))

    well_pos = _find_row(ws, L, ("WELL",), (4, 9), (1, 10))
    if well_pos:
        v = _value_after(ws, L, well_pos[0], well_pos[1], spans=S)
        if v: header["well_name"] = _clean(str(v))

    rig_pos = _find_row(ws, L, ("RIG",), (4, 9), (1, 15))
    if rig_pos:
        v = _value_after(ws, L, rig_pos[0], rig_pos[1], spans=S)
        if v: header["rig_name"] = _clean(str(v))

    csg_pos = _find_row(ws, L, ("LAST CSG",), (4, 9), (1, 25))
    if csg_pos:
        label = _clean(_cell(ws, csg_pos[0], csg_pos[1], L) or "")
        size_m = re.search(r"(\d+[\"'](?:\s*\d+/\d+)?)", label)
        v = _value_after(ws, L, csg_pos[0], csg_pos[1], spans=S)
        if v is not None:
            depth_s = _clean(str(v))
            sz = size_m.group(1) if size_m else ""
            header["last_csg_size"] = sz
            header["top_shoe"] = f'{sz} @ {depth_s}' if sz else depth_s

    tol_pos = _find_row(ws, L, ("TOL",), (4, 9), (1, 30))
    if tol_pos:
        label = _clean(_cell(ws, tol_pos[0], tol_pos[1], L) or "")
        size_m = re.search(r"(\d+[\"'](?:\s*\d+/\d+)?)", label)
        v = _value_after(ws, L, tol_pos[0], tol_pos[1], spans=S)
        if v is not None:
            depth_s = _clean(str(v))
            sz = size_m.group(1) if size_m else ""
            header["top_of_liner"] = f'{sz} @ {depth_s}' if sz else depth_s

    tl_pos = _find_row(ws, L, ("TOP LINER",), (4, 9), (1, 30))
    if tl_pos:
        label = _clean(_cell(ws, tl_pos[0], tl_pos[1], L) or "")
        size_m = re.search(r"(\d+[\-\s]*\d*/?\d*)\s*:?\s*$", label)
        depth_pos = _find_row(ws, L, ("DEPTH",), (tl_pos[0], tl_pos[0] + 2), (tl_pos[1], tl_pos[1] + 6))
        if depth_pos:
            v = _value_after(ws, L, depth_pos[0], depth_pos[1], spans=S)
            if v is not None:
                depth_s = _clean(str(v))
                header["top_liner_depth"] = depth_s
                header["top_liner"] = f'4-1/2 @ {depth_s}'

    type_pos = _find_row(ws, L, ("TYPE",), (4, 9), (20, 30))
    if type_pos:
        v = _value_after(ws, L, type_pos[0], type_pos[1], spans=S)
        if v: header["mud_type"] = _clean(str(v))

    # =====================================================================
    # OPERATIONS — header row found via "WORK-OVER OPERATIONS"; the real
    # data columns are C(start)/D(end)/E(desc, wide)/P(bill)/Q(hours).
    # Columns A/B (immediately left of start/end) are a stray, unrelated
    # table — see module docstring — and are intentionally never read.
    # "After midnight :" is detected as a text PREFIX on a description
    # cell rather than a separate stop-marked section.
    # =====================================================================
    activities: List[Dict[str, Any]] = []
    text_sections: Dict[str, str] = {}

    ops_hdr = None
    for r in range(24, 30):
        v = _cell(ws, r, 3, L)   # col C
        if isinstance(v, str) and _clean(v).upper() == "DE":
            ops_hdr = (r, 3)
            break
    ops_start = ops_hdr[0] + 1 if ops_hdr else 28

    for r in range(ops_start, ops_start + 25):
        start_v = _cell(ws, r, 3, L)   # C
        end_v   = _cell(ws, r, 4, L)   # D
        desc    = _clean(_cell(ws, r, 5, L) or "")   # E
        bill    = _clean(_cell(ws, r, 16, L) or "")  # P
        hours_v = _cell(ws, r, 17, L)                # Q

        if desc.upper().lstrip().startswith("AFTER MIDNIGHT"):
            v = re.sub(r"^\s*AFTER\s*MIDNIGHT\s*:?\s*", "", desc, flags=re.IGNORECASE).strip()
            if v:
                text_sections["after_midnight"] = v
            continue

        if desc.upper().lstrip().startswith("NB"):
            v = re.sub(r"^\s*NB\s*:?\s*", "", desc, flags=re.IGNORECASE).strip()
            if v:
                text_sections["notes"] = v
            continue

        start_t = _time_from_cell(start_v)
        end_t   = _time_from_cell(end_v)
        hours   = _hours_from_cell(hours_v)

        if start_t is None and end_t is None and not bill and hours < 1e-6:
            continue  # blank row (including the stray A/B-only rows)

        if hours == 0.0 and start_t and end_t:
            sm = start_t.hour * 60 + start_t.minute
            em = end_t.hour * 60 + end_t.minute
            if em == sm:    hours = 24.0
            elif em > sm:   hours = (em - sm) / 60.0
            else:           hours = (em + 1440 - sm) / 60.0

        activities.append({
            "start_time": start_t, "end_time": end_t, "hours": hours,
            "phase_name": "", "code": "", "sub": "",
            "description": desc,
            "start_md": 0, "end_md": 0,
            "npt": 0, "npt_detail": "", "npt_company": "", "op_company": "",
            "bill": bill,
        })

    # =====================================================================
    # TARIFF TOTALS — T1/T2/T3/T4 as COLUMN headers (not row labels like
    # the other templates); the daily hours sit one row below each code.
    # =====================================================================
    tarif_totals: Dict[str, float] = {}
    tar_hdr = _find_row(ws, L, ("T1",), (20, 27), (15, 25))
    if tar_hdr:
        hr, hc = tar_hdr
        codes = []
        for c in range(hc, hc + 6):
            v = _cell(ws, hr, c, L)
            if isinstance(v, str) and re.match(r"^(T\d|NR|T0)$", _clean(v)):
                codes.append((c, _clean(v).lower()))
            elif codes:
                break
        for c, code in codes:
            v = _cell(ws, hr + 1, c, L)
            if v is not None and _clean(str(v)) != "":
                tarif_totals[code] = _float(v)

    # =====================================================================
    # CURRENT STATUS / FORECAST / REQUIREMENTS / REMARKS — clean same-row
    # label/value pairs.
    # =====================================================================
    for keywords, key in ((("CURRENT STATUS",), "current_operation"),
                           (("24 H FORECAST",), "plan_operations"),
                           (("REQUIREMENTS",), "requirements"),
                           (("REMARKS",), "remarks")):
        pos = _find_row(ws, L, keywords, (55, 70), (1, 6))
        if pos:
            v = _value_after(ws, L, pos[0], pos[1], max_scan=12, spans=S)
            if v:
                text_sections[key] = _clean(str(v))
    if "current_operation" in text_sections:
        text_sections["day_summary"] = text_sections["current_operation"]

    # =====================================================================
    # MUD CHECKS  (label/value pairs scattered around rows 15-23, cols
    # X/Y/Z/AA — inconsistent which of Y/AA holds the value, so we scan a
    # small window after each label's merge).
    # =====================================================================
    mud_checks: Dict[str, Any] = {}
    if "mud_type" in header:
        mud_checks["mud_type"] = header["mud_type"]
    MUD_LABELS = {
        "DENSIT":   "density",
        "V.MARSH":  "fun_vis",
        "Y.P":      "yp",
        "FILTRAT":  "filtrat",
        "GEL  0":   "gel0sec",
        "GEL 0":    "gel0sec",
        "GEL  10":  "gel10sec",
        "GEL 10":   "gel10sec",
        "SABL":     "sand_pct",
        "SOLIDE":   "solid_pct",
        "HUILE":    "oil_pct",
        "LGS":      "lgs",
        "PV":       "pv",
        "H/E":      "he_ratio",
    }
    for r in range(14, 24):
        for c in range(23, 27):   # cols X..AA
            label = _clean(_cell(ws, r, c, L) or "")
            if not label:
                continue
            up = label.upper()
            for kw, key in MUD_LABELS.items():
                if kw in up:
                    v = _numeric_value_after(ws, L, r, c, max_scan=3, spans=S)
                    if v is not None:
                        mud_checks[key] = _float(v)
                    break

    # =====================================================================
    # MUD VOLUME  (rows 8-14, cols X/Y label, AA or AD value — both
    # observed in the sample, so we scan a wider window).
    # =====================================================================
    mud_volume: Dict[str, Any] = {}
    VOLUME_LABELS = {
        "FABRICATION":        "fabrication_vol",
        "R" "\u00c9CEPTION":  "reception_vol",
        "RECEPTION":          "reception_vol",
        "DUMPING":            "dumped_volume",
        "TRIPPING":           "tripping_vol",
        "VOLUME PUITS":       "string_volume",
        "VOLUME SURFACE":     "pits_volume",
    }
    for r in range(7, 15):
        for c in range(23, 27):   # cols X..AA
            label = _clean(_cell(ws, r, c, L) or "")
            if not label:
                continue
            up = label.upper()
            for kw, key in VOLUME_LABELS.items():
                if kw in up:
                    v = _numeric_value_after(ws, L, r, c, max_scan=4, spans=S)
                    if v is not None:
                        mud_volume[key] = _float(v)
                    break

    # =====================================================================
    # PERSONNEL  (rows ~45-63, col X label spans to Z, value at AA).
    # The "TOTAL PERSON" row is captured as header["personnel_total"]
    # rather than as a personnel_data entry.
    # =====================================================================
    personnel: List[Dict[str, Any]] = []
    pers_start = _find_row(ws, L, ("SH / DP SUPERVISOR",), (40, 50), (23, 27))
    scan_from = pers_start[0] if pers_start else 46
    for r in range(scan_from, scan_from + 20):
        label = _clean(_cell(ws, r, 24, L) or "")   # col X
        if not label:
            continue
        up = label.upper()
        if up.startswith("TOTAL PERSON"):
            v = _numeric_value_after(ws, L, r, 24, max_scan=4, spans=S)
            if v is not None:
                header["personnel_total"] = _int(v)
            continue
        if up.startswith("TOTAL"):
            continue
        v = _numeric_value_after(ws, L, r, 24, max_scan=4, spans=S)
        if v is None:
            continue
        personnel.append({"company": label, "number": _int(v),
                           "hours": "", "names": ""})

    # =====================================================================
    # SUPERVISOR NAME — "SV/DP" label near the bottom, value merged wide.
    # =====================================================================
    sup_pos = _find_row(ws, L, ("SV/DP",), (60, 70), (20, 27))
    if sup_pos:
        v = _value_after(ws, L, sup_pos[0], sup_pos[1], spans=S)
        if v:
            header["supervisor"] = _clean(str(v)).strip()

    # BOP test dates
    last_bop_pos = _find_row(ws, L, ("LAST  TEST BOP", "LAST TEST BOP"), (50, 62), (16, 24))
    if last_bop_pos:
        v = _value_after(ws, L, last_bop_pos[0], last_bop_pos[1], spans=S)
        d = _date_parse(v)
        if d: header["bop_test"] = d
    next_bop_pos = _find_row(ws, L, ("NEXT TEST BOP",), (50, 62), (16, 24))
    if next_bop_pos:
        v = _value_after(ws, L, next_bop_pos[0], next_bop_pos[1], spans=S)
        d = _date_parse(v)
        if d: header["next_bop_test"] = d

    # =====================================================================
    # BACK-FILL BILL CODES (no-op if every op is already tagged)
    # =====================================================================
    activities = _assign_bill_codes(activities, tarif_totals)

    wb.close()

    return {
        "header":             header,
        "activities":         activities,
        "text_sections":      text_sections,
        "mud_checks":         mud_checks,
        "mud_volume":         mud_volume,
        "mud_chemical_usage": [],
        "personnel_data":     personnel,
        "pumps":              [],
        "well_location":      {},
        "survey_data":        [],
        "safety":             {},
        "tarif_totals":       tarif_totals,
    }


# Drop-in compat
parse_daily_excel_report = parse_entp219


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: entp219_extract.py SOURCE.xlsx")
    data = parse_entp219(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
