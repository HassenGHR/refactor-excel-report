#!/usr/bin/env python3
"""
tp195_extract.py — extract a TP-195 (ENTP rig 195, AIN T'SILA field,
AT-NN wells) Daily Workover Report into the standard dict shape.

Source layout
-------------
SONATRACH DIVISION PRODUCTION "Daily Drilling Report" — the SAME master
spreadsheet template as entp204_extract.py (verified cell-by-cell against
a real report: identical row layout for the header, BIT DATA / LAST BHA /
MUD CHECKS banner, SURVEY DATA, PERSONNEL DATA, operations table, tariff
totals, and text-sections blocks — only the rig/well/field values differ).
See entp204_extract.py's docstring for the general notes on this
template family; this file follows the same conventions.

Header fields are "LABEL : value" combined into a SINGLE cell (e.g.
"WELL : AT 43", "DATE : 14/07/2026") — NOT split across separate
label/value cells as an earlier version of this extractor assumed. The
header block is duplicated on the sheet (page 1 around row 5, page 2
around row 60); we scan a wide enough row range to catch whichever copy
comes first.

Known quirks specific to this template's data (verified against the
sample report):
  * FROM/TO/HRS cells in the operations table are TEXT in "HHhMM" form
    ("00h00", "24h00") rather than time/timedelta objects — already
    handled by the existing _tod_minutes()/_duration_hours() 'h'
    separator support, no special-casing needed.
  * The tariff totals row combines label+value in one cell again
    ("T1 =24h00"), same as some entp204 reports — handled generically.
  * "ACTUEL OPERATIONS" (French spelling) is what this file uses, while
    the entp204 sample used "ACTUAL OPERATIONS" (English) — we check
    both spellings.
  * The right-hand mud-checks column's VALUE sits at col O (15); the
    label spans cols M:N (13:14), so a naive "read col 14 or col 15"
    fallback picks up the label text at col 14 first and never reaches
    the real value at col 15. Fixed by reading col 15 directly.
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Union, Optional

from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _clean(v) -> str:
    if v is None: return ""
    s = str(v).replace("\u202f", " ").replace("\u00a0", " ")
    return re.sub(r"\s+", " ", s).strip()


def _float(v, default=0.0) -> float:
    if v is None: return default
    if isinstance(v, (int, float)): return float(v)
    s = _clean(v).replace(",", ".").replace(" ", "")
    if s in ("", "-", "/", "None", "\\", "\\\\"): return default
    m = re.match(r"^([-+]?\d+(?:\.\d+)?)", s)
    if m:
        try: return float(m.group(1))
        except ValueError: pass
    return default


def _int(v, default=0) -> int:
    return int(_float(v, float(default)))


def _date_parse(v):
    if v is None: return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date_type): return v
    s = _clean(v)
    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        try: return date_type(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError: return None
    return None


def _label_value(cell_text: str, label: str) -> str:
    """From a 'LABEL : value' combined string, return the value part."""
    s = _clean(cell_text)
    rx = re.compile(rf"^{re.escape(label)}\s*[:：]\s*(.*)$", re.IGNORECASE)
    m = rx.match(s)
    return m.group(1).strip() if m else ""


def _scan_label_value(ws, L, label, row_range=(1, 65), col_range=(1, 20)) -> str:
    """Scan a region for a cell whose text starts with 'LABEL :' and
    return the (non-empty) value part."""
    for r in range(row_range[0], row_range[1] + 1):
        for c in range(col_range[0], col_range[1] + 1):
            v = _cell(ws, r, c, L)
            if v is None: continue
            got = _label_value(v, label)
            if got:
                return got
    return ""


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
    """Map every (row, col) inside a merged range to that range's LAST
    column index, so we can jump past a label's own merge in one step
    when looking for an adjacent value cell."""
    spans = {}
    for mr in ws.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                spans[(r, c)] = mr.max_col
    return spans


def _cell(ws, r, c, lookup):
    v = ws.cell(r, c).value
    return v if v is not None else lookup.get((r, c))


def _scan_label_cell(ws, L, S, label, row_range=(1, 65), col_range=(1, 20)):
    """Find a cell whose text starts with LABEL, and return its RAW value
    (unconverted — could be str/datetime/timedelta/int) plus (row, col).

    Handles both layouts seen across this template's report variants:
      - combined:  "LABEL : value" in one cell.
      - split:     the label cell holds ONLY the label; the real value
        lives in the very next cell after the label's own merged span.

    Returns (None, None, None) if the label isn't found."""
    rx = re.compile(rf"^{re.escape(label)}\s*[:：=]?\s*(.*)$", re.IGNORECASE)
    for r in range(row_range[0], row_range[1] + 1):
        for c in range(col_range[0], col_range[1] + 1):
            v = _cell(ws, r, c, L)
            if v is None or not isinstance(v, str):
                continue
            m = rx.match(_clean(v))
            if not m:
                continue
            trailing = m.group(1).strip()
            if trailing:
                return trailing, r, c
            end_c = S.get((r, c), c)
            nv = _cell(ws, r, end_c + 1, L)
            if nv is not None and _clean(str(nv)) != "":
                return nv, r, c
            return None, r, c
    return None, None, None


def _normalize_rig(name: str) -> str:
    s = _clean(name).upper()
    m = re.search(r"ENTP[\s#\-]*(\d+)", s)
    if m: return f"ENTP-{m.group(1)}"
    m = re.search(r"TP[\s#\-]*(\d+)", s)
    if m: return f"TP-{m.group(1)}"
    return _clean(name)


def _parse_op_time(v) -> Optional[time]:
    """Parse operation time cells: datetime.time, '0:00:00', '24h00',
    '1 day, 0:00:00' (= midnight), or timedelta."""
    if v is None: return None
    if isinstance(v, time): return v
    if isinstance(v, timedelta):
        total = int(v.total_seconds())
        if total >= 86400: return time(0, 0)
        return time((total // 3600) % 24, (total % 3600) // 60)
    if isinstance(v, datetime):
        return v.time()
    s = _clean(v)
    if not s: return None
    if "day" in s.lower():
        return time(0, 0)
    m = re.match(r"^(\d{1,2})\s*[:hH]\s*(\d{2})(?:\s*[:hH]\s*\d{2})?$", s)
    if m:
        return time(int(m.group(1)) % 24, int(m.group(2)))
    return None


def _prev_end_minutes(prev_op) -> Optional[int]:
    st = prev_op.get("start_time")
    if st is None:
        return None
    start_min = st.hour * 60 + st.minute
    hours = prev_op.get("hours", 0.0) or 0.0
    return int(round(start_min + hours * 60))


def _tod_minutes(v) -> Optional[int]:
    """Convert an operation time cell to minutes-since-midnight (0-1440),
    where 1440 represents a genuine 24:00 end-of-day (not 0)."""
    if v is None:
        return None
    if isinstance(v, timedelta):
        total = int(v.total_seconds())
        return min(total // 60, 1440)
    if isinstance(v, time):
        return v.hour * 60 + v.minute
    if isinstance(v, datetime):
        return v.hour * 60 + v.minute
    s = _clean(v)
    if not s:
        return None
    if "day" in s.lower():
        return 1440
    if s in ("24:00", "24:00:00", "24h00"):
        return 1440
    m = re.match(r"^(\d{1,2})\s*[:hH]\s*(\d{2})(?:\s*[:hH]\s*\d{2})?$", s)
    if m:
        h, mn = int(m.group(1)), int(m.group(2))
        return h * 60 + mn
    return None


def _hhmm_to_hours(s: str) -> float:
    """'24h00' / '00h00' / 'CUM 00h00' → decimal hours."""
    s = _clean(s)
    m = re.search(r"(\d{1,2})\s*[hH:]\s*(\d{2})", s)
    if not m: return 0.0
    return int(m.group(1)) + int(m.group(2)) / 60.0


def _hours_from_any(v) -> float:
    """Convert a tariff-hours value to decimal hours, whichever form it
    takes: a timedelta (split-cell layout) or a string like '24h00'
    (combined-cell layout)."""
    if v is None:
        return 0.0
    if isinstance(v, timedelta):
        return v.total_seconds() / 3600.0
    if isinstance(v, (int, float)):
        return float(v)
    return _hhmm_to_hours(str(v))


def _duration_hours(v) -> float:
    """Decode an HRS cell — number, timedelta, time, 'HHhMM' text, or
    '1 day, 0:00:00' string."""
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    if isinstance(v, timedelta):
        return v.total_seconds() / 3600.0
    if isinstance(v, time):
        return v.hour + v.minute / 60.0
    if isinstance(v, datetime):
        if v.year == 1900 and v.month == 1 and v.day == 1:
            return v.hour + v.minute / 60.0
        return 0.0
    s = _clean(v)
    m = re.match(r"^(\d+)\s+day", s)
    if m:
        days = int(m.group(1))
        rest = re.sub(r"^\d+\s+day,?\s*", "", s)
        t = _parse_op_time(rest)
        return days * 24 + (t.hour + t.minute / 60.0 if t else 0)
    t = _parse_op_time(v)
    if t: return t.hour + t.minute / 60.0
    return _float(v)


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_tp195(source: Union[Path, str, BytesIO]) -> dict:
    wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)
    S = _build_span_end_cols(ws)

    header = {}

    # =====================================================================
    # HEADER — "LABEL : value" combined cells (some reports split label
    # and value into adjacent cells instead — _scan_label_cell tries both
    # forms). The header block appears twice on the sheet (page 1 near
    # row 5, page 2 near row 60); scanning rows 1-65 catches whichever
    # copy comes first.
    # =====================================================================
    well = _scan_label_value(ws, L, "WELL", (1, 65))
    if well: header["well_name"] = well

    dv, _, _ = _scan_label_cell(ws, L, S, "DATE", (1, 65))
    dd = _date_parse(dv)
    if dd: header["date"] = dd

    for lbl in ("TOTAL MD", "MD"):
        mdv, _, _ = _scan_label_cell(ws, L, S, lbl, (1, 65))
        if mdv is not None:
            v = _float(mdv)
            if v: header["well_md"] = v
            break
    for lbl in ("TOTAL TVD", "TVD"):
        tvdv, _, _ = _scan_label_cell(ws, L, S, lbl, (1, 65))
        if tvdv is not None:
            v = _float(tvdv)
            if v: header["tvd"] = v
            break

    formation = _scan_label_value(ws, L, "FORMATION TOP", (1, 65))
    if formation: header["formation_top"] = formation

    reprise_v, _, _ = _scan_label_cell(ws, L, S, "REPRISE DATE", (1, 65))
    reprise_d = _date_parse(reprise_v)
    if reprise_d: header["reprise_date"] = reprise_d

    repv, _, _ = _scan_label_cell(ws, L, S, "REP N°", (1, 65))
    if repv is None:
        repv, _, _ = _scan_label_cell(ws, L, S, "REP N", (1, 65))
    if repv is not None:
        n = _int(repv, 0)
        if n > 0: header["day_number"] = n

    # Supervisors — names joined with '+', '&', ',' or '/' depending on
    # the report.
    sup_raw = _scan_label_value(ws, L, "Supervisor", (1, 65))
    if sup_raw:
        names = re.split(r"\s*[+&,/]\s*", sup_raw)
        names = [n.strip() for n in names if n.strip()]
        if names:
            header["supervisor"] = names[0]
        if len(names) >= 2:
            header["superintendent"] = names[1]

    rig = _scan_label_value(ws, L, "RIG NAME", (1, 65))
    if rig: header["rig_name"] = _normalize_rig(rig)

    field = _scan_label_value(ws, L, "FIELD", (1, 65))
    if field: header["field_name"] = field

    # "OFFICE REP" is this template's alternate name for a second
    # supervisor on some reports (distinguishing marker vs TP-182's
    # "SUPERINTANDANT"); prefer it over the generic Superintandant field
    # when it holds a real name.
    office_rep = _scan_label_value(ws, L, "OFFICE REP", (1, 65))
    if office_rep and office_rep not in ("/", "-", ""):
        header["superintendent"] = office_rep
    else:
        supt_raw = _scan_label_value(ws, L, "Superintandant", (1, 65))
        if supt_raw and supt_raw not in ("/", "-", ""):
            header["superintendent"] = supt_raw

    accv, _, _ = _scan_label_cell(ws, L, S, "ACC FREE", (1, 65))
    if accv is not None:
        header["accident_free_days"] = _int(accv, 0)

    dnpt_v, _, _ = _scan_label_cell(ws, L, S, "Daily NPT", (1, 65))
    if dnpt_v is not None:
        header["daily_npt"] = _hours_from_any(dnpt_v)

    lft = _scan_label_value(ws, L, "Last Formation Test", (1, 65))
    if lft and lft not in ("/", "-", ""):
        header["last_formation_test"] = lft

    csgv, _, _ = _scan_label_cell(ws, L, S, "Last Csg Shoe", (1, 65))
    if csgv is not None:
        s = _clean(str(csgv))
        if s and s.replace("\\", "").strip() not in ("", "m"):
            try:
                d = int(float(re.sub(r"[^\d.]", "", s)))
                header["top_shoe"] = f"@ {d}m"
            except (ValueError, TypeError):
                header["top_shoe"] = f"@ {s}"

    liner = _scan_label_value(ws, L, "Last Liner", (1, 65))
    if liner and liner.replace("\\", "").strip() not in ("", "m"):
        header["last_tol"] = liner

    # BOP test — "NEXT BOP TEST" or "LAST TEST BOP" depending on report.
    for lbl, key in (("NEXT BOP TEST", "next_bop_test"),
                     ("NEXT  TEST BOP", "next_bop_test"),
                     ("LAST  TEST BOP", "bop_test"),
                     ("LAST TEST BOP", "bop_test"),
                     ("L,BOP", "bop_test"), ("L.BOP", "bop_test")):
        bv, _, _ = _scan_label_cell(ws, L, S, lbl, (1, 65))
        bd = _date_parse(bv)
        if bd:
            header[key] = bd
            break

    safety_v, _, _ = _scan_label_cell(ws, L, S, "Last Safety Meeting", (1, 65))
    safety_d = _date_parse(safety_v)
    if safety_d: header["last_safety_meeting"] = safety_d

    cnpt_v, _, _ = _scan_label_cell(ws, L, S, "Cum NPT", (1, 65))
    if cnpt_v is not None:
        header["cum_npt"] = _hours_from_any(cnpt_v)

    reason = _scan_label_value(ws, L, "WORK-OVER REASON", (1, 65))
    if reason: header["well_objective"] = reason

    # =====================================================================
    # BIT DATA (rows 11-27, col E value) — unchanged from the previous
    # version of this extractor; verified these positions ARE correct on
    # the real report (rows 13+, col E).
    # =====================================================================
    bit_data = {}
    bit_labels = [
        (13, "bit_size"), (14, "bit_manufacture"), (15, "bit_type"),
        (16, "bit_iadc"), (17, "bit_serial"), (18, "bit_jets"),
        (19, "bit_tfa"), (20, "bit_depth_in"),
    ]
    for row, key in bit_labels:
        v = _cell(ws, row, 5, L)        # col E
        if v is not None:
            bit_data[key] = _clean(v)
    if bit_data:
        header["bit_data"] = " | ".join(f"{k}={v}" for k, v in bit_data.items())

    # LAST BHA (rows 13-29, col F desc / H OD / I length)
    bha_parts = []
    for row in range(13, 30):
        name = _clean(_cell(ws, row, 6, L) or "")     # F
        if not name or name.upper() == "BHA TOTAL":
            continue
        od     = _clean(_cell(ws, row, 8, L) or "")   # H
        length = _cell(ws, row, 9, L)                 # I
        if length is not None:
            try: length = f"{float(length):.2f}m"
            except (ValueError, TypeError): length = _clean(length)
        parts = [name]
        if od: parts.append(od)
        if length: parts.append(str(length))
        bha_parts.append(" / ".join(parts))
    if bha_parts:
        header["bha_details"] = " | ".join(bha_parts)

    # =====================================================================
    # MUD CHECKS  (rows 11-29; left: label col J/K, value col L.
    # Right: label col M/N, value col O — NOT col N, which is part of the
    # label's own merge and would just re-read the label text.)
    # =====================================================================
    mud_checks = {}
    mud_type_raw = _cell(ws, 11, 13, L) or _cell(ws, 11, 14, L)  # M11 / N11
    if mud_type_raw:
        s = _clean(mud_type_raw)
        if s.upper() not in ("MUD", "TYPE", "TYPE1/TYPE2"):
            mud_checks["mud_type"] = s

    left_mud = [
        (13, "density"), (14, "depth"), (15, "fl_tmp"), (16, "fun_vis"),
        (17, "pv"), (18, "yp"), (19, "gel10sec"), (20, "gel10m"),
        (21, "hpht_fl"), (22, "lgs"), (24, "hgs"), (27, "solid"),
    ]
    for row, key in left_mud:
        v = _cell(ws, row, 12, L)        # col L
        if v is None: continue
        try: mud_checks[key] = _float(v)
        except (ValueError, TypeError): pass

    right_mud = [
        (13, "pom"), (14, "sand_pct"), (15, "solids_pct"), (16, "oil_pct"),
        (17, "water_pct"), (22, "es"), (23, "brine"), (27, "excess_lime"),
    ]
    for row, key in right_mud:
        label = _clean(_cell(ws, row, 13, L) or "")   # M
        v = _cell(ws, row, 15, L)                      # O — the real value col
        if v is None:
            continue
        if key == "oil_water_ratio_hint":
            continue
        try:
            mud_checks[key] = _float(v)
        except (ValueError, TypeError):
            mud_checks[key] = _clean(str(v))

    # OIL/WATER ratio (text like '86/14') sits at row 18, col O too.
    ow = _cell(ws, 18, 15, L)
    if ow is not None:
        mud_checks["oil_water_ratio"] = _clean(str(ow))

    # =====================================================================
    # MUD VOLUMES  (row 28: OBM Surface Vol @ col L, Well Vol @ col O;
    # row 29: Losses @ col L, crude oil Vol @ col O)
    # =====================================================================
    mud_volume = {}
    well_v = _cell(ws, 28, 15, L)
    if well_v is not None: mud_volume["string_volume"] = _float(well_v)
    obm_surface = _cell(ws, 28, 12, L)
    if obm_surface is not None: mud_volume["pits_volume"] = _float(obm_surface)
    losses = _cell(ws, 29, 12, L)
    if losses is not None:
        try: mud_volume["surface_loss"] = _float(losses)
        except (ValueError, TypeError): pass
    gas_oil = _cell(ws, 29, 15, L)
    if gas_oil is not None:
        try: mud_volume["gas_oil_volume"] = _float(gas_oil)
        except (ValueError, TypeError): pass
    if mud_volume:
        mud_volume["total_volume"] = sum(
            v for v in (mud_volume.get("string_volume"), mud_volume.get("pits_volume"))
            if v is not None
        )

    # =====================================================================
    # MUD CHEMICAL USAGE  (rows 47-55)
    # Left block:  col B item, col E units, col F rec'd, col G used, col H end
    # Right block: col I item, col L units, col M rec'd, col O used, col P end
    # =====================================================================
    chemicals = []
    for row in range(47, 56):
        item_l = _clean(_cell(ws, row, 2, L) or "")
        if item_l:
            chemicals.append({
                "item": item_l,
                "units": _clean(_cell(ws, row, 5, L) or ""),
                "received": _clean(_cell(ws, row, 6, L)) if _cell(ws, row, 6, L) is not None else "",
                "used":     _clean(_cell(ws, row, 7, L)) if _cell(ws, row, 7, L) is not None else "",
                "on_loc":   _clean(_cell(ws, row, 8, L)) if _cell(ws, row, 8, L) is not None else "",
            })
        item_r = _clean(_cell(ws, row, 9, L) or "")
        if item_r:
            chemicals.append({
                "item": item_r,
                "units": _clean(_cell(ws, row, 12, L) or ""),
                "received": _clean(_cell(ws, row, 13, L)) if _cell(ws, row, 13, L) is not None else "",
                "used":     _clean(_cell(ws, row, 15, L)) if _cell(ws, row, 15, L) is not None else "",
                "on_loc":   _clean(_cell(ws, row, 16, L)) if _cell(ws, row, 16, L) is not None else "",
            })

    # =====================================================================
    # SURVEY DATA — header found dynamically via "SURVEY DATA" label
    # (row drifts slightly report-to-report), data 3 rows below it.
    # =====================================================================
    survey_data = []
    surv_hdr = None
    for r in range(55, 70):
        if "SURVEY DATA" in _clean(_cell(ws, r, 2, L) or "").upper():
            surv_hdr = r
            break
    if surv_hdr:
        for row in range(surv_hdr + 3, surv_hdr + 6):
            md = _cell(ws, row, 2, L)
            if md is None or not isinstance(md, (int, float)):
                continue
            survey_data.append({
                "md":  _float(md),
                "tvd": _float(_cell(ws, row, 4, L)),
                "azimuth": _float(_cell(ws, row, 6, L)),
                "inclination": _float(_cell(ws, row, 8, L)),
                "dls": _float(_cell(ws, row, 10, L)),
                "ns":  _float(_cell(ws, row, 12, L)),
                "ew":  _float(_cell(ws, row, 14, L)),
                "vs":  _float(_cell(ws, row, 16, L)),
            })

    # =====================================================================
    # PERSONNEL — header found dynamically via "PERSONNEL DATA" label.
    # Left: col B company, D number, E hours, G names.
    # Right: col J company, L number, M hours, N names.
    # =====================================================================
    personnel = []
    pers_hdr = None
    for r in range(60, 80):
        if "PERSONNEL DATA" in _clean(_cell(ws, r, 2, L) or "").upper():
            pers_hdr = r
            break
    if pers_hdr:
        for row in range(pers_hdr + 2, pers_hdr + 12):
            company_l = _clean(_cell(ws, row, 2, L) or "")
            if company_l.upper() in ("FROM",):
                break
            if company_l and company_l.upper() not in ("COMPANY", "PERSONNEL DATA"):
                count_l = _cell(ws, row, 4, L)
                names_l = _clean(_cell(ws, row, 7, L) or "")
                if "WATER TRUCK" in company_l.upper() and count_l is not None:
                    header["water_truck"] = _int(count_l)
                personnel.append({
                    "company": company_l,
                    "number":  _int(count_l, 0),
                    "hours":   _clean(_cell(ws, row, 5, L) or ""),
                    "names":   names_l,
                })
            company_r = _clean(_cell(ws, row, 10, L) or "")
            if company_r and company_r.upper() not in ("COMPANY",):
                count_r = _cell(ws, row, 12, L)
                names_r = _clean(_cell(ws, row, 14, L) or "")
                personnel.append({
                    "company": company_r,
                    "number":  _int(count_r, 0),
                    "hours":   _clean(_cell(ws, row, 13, L) or ""),
                    "names":   names_r,
                })
            if not company_l and not company_r:
                continue

    # =====================================================================
    # OPERATIONS — header found dynamically via "FROM"/"TO" cells (row
    # drifts depending on how tall the SURVEY DATA/PERSONNEL DATA blocks
    # are that day, so we don't hard-code row 77).
    # Cols: B=FROM C=TO D=HRS E=DESCRIPTION(wide) O=BILL P=COMPANY.
    # "AFTER MIDNIGTH"/"AFTER MIDNIGHT" marks the boundary into the next
    # day's early-hours operations, which we collect as free text rather
    # than as timed ops (their times restart from 00:00 and would
    # otherwise double-count against the 24h day total).
    # =====================================================================
    activities = []
    after_midnight_parts = []
    ops_header_row = None
    for r in range(70, 105):
        b = _clean(_cell(ws, r, 2, L) or "")
        c = _clean(_cell(ws, r, 3, L) or "")
        if b.upper() == "FROM" and c.upper() == "TO":
            ops_header_row = r
            break

    if ops_header_row:
        in_after_midnight = False
        for r in range(ops_header_row + 1, ops_header_row + 30):
            from_v = _cell(ws, r, 2, L)
            to_v   = _cell(ws, r, 3, L)
            desc   = _clean(_cell(ws, r, 5, L) or "")
            bill   = _clean(_cell(ws, r, 15, L) or "")     # col O
            comp   = _clean(_cell(ws, r, 16, L) or "")     # col P

            b_text = _clean(_cell(ws, r, 2, L) or "").upper()
            if any(m in b_text for m in ("T1 =", "ACTUEL", "ACTUAL",
                                          "PLAN OPER", "REMARKS",
                                          "REQUIREMENTS", "SURVEY",
                                          "WELL :", "SONATRACH")):
                break

            if "AFTER MIDN" in desc.upper():
                in_after_midnight = True
                continue

            if desc.upper() in ("DESCRIPTION", "") and from_v is None and to_v is None:
                continue
            if from_v is None and to_v is None and not desc:
                continue

            if in_after_midnight:
                if desc:
                    after_midnight_parts.append(desc)
                continue

            sm = _tod_minutes(from_v)
            em = _tod_minutes(to_v)

            if activities and sm is not None:
                prev_end_min = _prev_end_minutes(activities[-1])
                if prev_end_min is not None and sm < prev_end_min <= (em if em is not None else 1440):
                    sm = prev_end_min

            start_t = _parse_op_time(from_v)
            end_t   = _parse_op_time(to_v)
            if sm is not None:
                start_t = time((sm // 60) % 24, sm % 60)

            if sm is not None and em is not None:
                if em > sm:
                    hours = (em - sm) / 60.0
                elif em == sm:
                    hours = 24.0 if sm == 0 else 0.0
                else:
                    hours = (em + 1440 - sm) / 60.0

                activities.append({
                    "start_time": start_t, "end_time": end_t, "hours": hours,
                    "phase_name": "", "code": "", "sub": "",
                    "description": desc,
                    "start_md": 0, "end_md": 0,
                    "npt": 0, "npt_detail": "", "npt_company": "",
                    "op_company": comp,
                    "bill": bill,
                })
            elif desc and activities and desc.upper() != "DESCRIPTION":
                activities[-1]["description"] = (
                    activities[-1]["description"] + "\n" + desc
                ).strip()

    # =====================================================================
    # TARIFF TOTALS — "T1 =24h00 | T2 = 00h00 | ..." row, found
    # dynamically. Handles BOTH combined-string values ("T1 =24h00") and
    # split label/adjacent-value-cell layouts (label "T1 =" alone, value
    # in the next cell as a timedelta), same as entp204_extract.py.
    # =====================================================================
    tarif_totals = {}
    for r in range(100, 120):
        row_has_code = any(
            re.match(r"^(T[1-4]|FT|NR)\s*=", _clean(_cell(ws, r, c, L) or ""))
            for c in range(2, 16)
        )
        if not row_has_code:
            continue
        for c in range(2, 16):
            cell_text = _clean(_cell(ws, r, c, L) or "")
            m = re.match(r"^(T[1-4]|FT|NR)\s*=\s*(.*)$", cell_text)
            if not m:
                continue
            code = m.group(1).lower()
            trailing = m.group(2).strip()
            is_cum = False
            if trailing:
                if trailing.upper().startswith("CUM"):
                    is_cum = True
                    trailing = re.sub(r"^CUM\s*", "", trailing, flags=re.IGNORECASE)
                hrs = _hhmm_to_hours(trailing)
            else:
                end_c = S.get((r, c), c)
                val = _cell(ws, r, end_c + 1, L)
                # A separate adjacent "CUM" text cell can also precede
                # the actual value cell on the split-cell layout.
                if isinstance(val, str) and val.strip().upper() == "CUM":
                    is_cum = True
                    val = _cell(ws, r, end_c + 2, L)
                hrs = _hours_from_any(val)
            if hrs > 0:
                # A 'CUM' value is cumulative-since-start, not today's
                # daily total — keep it out of the plain t1/t2/.. keys so
                # it doesn't get mistaken for (or summed against) today's
                # actual operations.
                tarif_totals[f"cum_{code}" if is_cum else code] = hrs
        break

    # =====================================================================
    # TEXT SECTIONS — "ACTUEL OPERATIONS"/"ACTUAL OPERATIONS" (either
    # spelling), "PLAN OPERATIONS", "REQUIREMENTS", "REMARKS".
    # =====================================================================
    text_sections = {}
    actuel = (_scan_label_value(ws, L, "ACTUEL OPERATIONS", (100, 122))
              or _scan_label_value(ws, L, "ACTUAL OPERATIONS", (100, 122)))
    if actuel:
        v = actuel
        if len(v) > 300:
            v = v[:300].rsplit(None, 1)[0] + "…"
        text_sections["current_operation"] = v
        text_sections["day_summary"]       = v
    plan = _scan_label_value(ws, L, "PLAN OPERATIONS", (100, 122))
    if plan:
        text_sections["plan_operations"] = plan
    requirements = _scan_label_value(ws, L, "REQUIREMENTS", (100, 122))
    if requirements:
        text_sections["requirements"] = requirements
    remarks = _scan_label_value(ws, L, "REMARKS", (100, 122))
    if remarks and remarks != "/":
        text_sections["remarks"] = remarks
    if after_midnight_parts:
        text_sections["after_midnight"] = " ".join(after_midnight_parts)

    # =====================================================================
    # SAFETY
    # =====================================================================
    safety = {}
    if header.get("accident_free_days") is not None:
        safety["accident_free_days"] = header["accident_free_days"]
    if header.get("water_truck") is not None:
        safety["water_truck"] = header["water_truck"]

    wb.close()

    return {
        "header": header,
        "activities": activities,
        "text_sections": text_sections,
        "mud_checks": mud_checks,
        "mud_volume": mud_volume,
        "mud_chemical_usage": chemicals,
        "personnel_data": personnel,
        "pumps": [],
        "well_location": {},
        "survey_data": survey_data,
        "safety": safety,
        "tarif_totals": tarif_totals,
    }


# Drop-in compat
parse_daily_excel_report = parse_tp195


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: tp195_extract.py SOURCE.xlsx")
    data = parse_tp195(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))