#!/usr/bin/env python3
"""
tp127_extract.py — extract a TP-127 (ENTP DF rig, DAD wells, DRAA DAOUI
field) Daily Work-Over Report into the standard router-dict shape.

Source layout
-------------
Single sheet (typically named after the well, e.g. "DAD #1"), ~84 rows ×
~30 active cols, ENGLISH template. Distinguishing markers:
    - "ENTP   DF" in A1 (operator marker)
    - "TP 127" in B1 (rig)
    - "RAPPORT  N°" / "DATE" / "WELL" / "SUPERINTENDANT" / "TOOL PUSHER'S"
      labels at the right side (cols N-S, rows 1-5)
    - Casing data in cols D-F rows 2-7 (CASING 30"/18⅝"/13⅜"/9⅝", TOP LINER 7", SABOT 7")
    - "LAST  BOP TEST" in I1 with date in I2
    - Operations table at row 25 header (E=FROM | F=TO | G=HRS | I=OPERATIONS),
      data rows 26-42, with AFTER MIDNIGHT label at r43 and after-midnight
      activity at r44+
    - Tarif totals at rows 15-18 cols O-Q (Daily T3 / Cumul T3 / Daily NR /
      Cumul NR)

This is DIFFERENT from the other TP-127 file (which uses the gw29 French
"Rapport journalier de Work - Over" template); detection routes by the
"ENTP   DF" + "RAPPORT  N°" English markers.
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Union

from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Helpers (same shape as the other extractors in this package)
# ---------------------------------------------------------------------------
def _clean(v) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _float(v, default=0.0) -> float:
    if v is None:
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = _clean(v).replace(",", ".").replace(" ", "")
    if s in ("", "-", "/", "None", "#REF!"):
        return default
    s = re.sub(r"[a-zA-Z%]+$", "", s).strip()
    try:
        return float(s)
    except ValueError:
        return default


def _int(v, default=0) -> int:
    f = _float(v, default)
    try:
        return int(f)
    except (TypeError, ValueError):
        return default


def _date_parse(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date_type):
        return v
    s = _clean(v)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except (ValueError, TypeError):
            continue
    return None


def _time_parse(v):
    """Parse a cell to time.

    Handles datetime, time, timedelta, "HH:MM"/"HH:MM:SS" strings,
    and Excel's "1900-01-01 00:00:00" sentinel (= midnight end-of-day).
    """
    if v is None:
        return None
    if isinstance(v, time):
        return v
    if isinstance(v, datetime):
        # Excel epoch 1899/1900 with time 0:00 means end-of-day 24:00 in
        # the workover convention used here.
        if v.year < 1901:
            return time(0, 0)
        return v.time()
    if isinstance(v, timedelta):
        total = int(v.total_seconds() // 60)
        return time((total // 60) % 24, total % 60)
    s = _clean(v)
    if not s or s in ("#REF!", "-"):
        return None
    if s in ("24:00", "24:00:00"):
        return time(0, 0)
    for fmt in ("%H:%M:%S", "%H:%M", "%Hh%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            continue
    return None


def _hours_value(v) -> float:
    """Extract hours from a cell. Accepts time (HH:MM = decimal hours),
    timedelta, "HH:MM" string, plain number, or "#REF!"/None.

    Excel quirk: when a duration cell is formatted as "h:mm" but the
    underlying value is a plain number ≤ 1.0 (e.g. 0.75), that is an
    Excel day-fraction (0.75 day = 18 hours), NOT 0.75 hours.  Values
    > 1.0 are treated as decimal hours directly.
    """
    if v is None:
        return 0.0
    if isinstance(v, time):
        return v.hour + v.minute / 60.0
    if isinstance(v, timedelta):
        return v.total_seconds() / 3600.0
    if isinstance(v, (int, float)):
        f = float(v)
        # Day-fraction heuristic: 0 < f <= 1 with no explicit fraction
        # boundary at 24 → interpret as Excel day-fraction.  Values that
        # are clearly hours (e.g. 13.25 from a fraction-of-day = 0.5521
        # which couldn't have come from typing "13:15") are > 1 and
        # treated as decimal hours.
        if 0 < f <= 1.0:
            return f * 24.0
        return f
    s = _clean(v)
    if s in ("", "#REF!", "-"):
        return 0.0
    if ":" in s:
        try:
            h, m = s.split(":")[:2]
            return int(h) + int(m) / 60.0
        except (ValueError, TypeError):
            return 0.0
    return _float(v)


def _build_merged_lookup(ws):
    lookup = {}
    for mr in ws.merged_cells.ranges:
        anchor_val = ws.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                if (r, c) != (mr.min_row, mr.min_col):
                    lookup[(r, c)] = anchor_val
    return lookup


def _cell(ws, r, c, lookup):
    v = ws.cell(r, c).value
    return v if v is not None else lookup.get((r, c))


def _pick_sheet(wb):
    """Pick the first sheet that has the TP-127 markers; fall back to active."""
    for name in wb.sheetnames:
        ws = wb[name]
        a1 = _clean(ws.cell(1, 1).value).upper()
        if "ENTP" in a1 and "DF" in a1:
            return ws
        b1 = _clean(ws.cell(1, 2).value).upper()
        if "TP 127" in b1 or "TP-127" in b1 or "TP127" in b1:
            return ws
    return wb.active


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_tp127(source: Union[Path, str, BytesIO]) -> dict:
    """Extract a TP-127 English-template daily report into the router dict."""
    wb = load_workbook(source, data_only=True)
    ws = _pick_sheet(wb)
    L = _build_merged_lookup(ws)

    # =====================================================================
    # HEADER
    # =====================================================================
    header = {}

    # Well — S2 (after "WELL" label in R2).  Some files render the label as
    # "WELL " with a trailing space.
    header["well_name"] = _clean(_cell(ws, 2, 19, L))           # S2

    # Rig — B1 ("TP 127") with operator marker in A1 ("ENTP   DF").
    header["rig_name"] = _clean(_cell(ws, 1, 2, L))             # B1

    # Field/Client — S3 (CLIENT label in R3).  No explicit field cell in
    # this template; derive from filename or leave blank.
    header["field_name"] = ""

    # Client / operator
    header["client"] = _clean(_cell(ws, 3, 19, L))              # S3

    # Region (often blank in this file)
    header["region"] = _clean(_cell(ws, 4, 19, L))              # S4

    # Date — S1 (right of "DATE" label in R1)
    header["date"] = _date_parse(_cell(ws, 1, 19, L))           # S1

    # Report number — P1 (right of "RAPPORT  N°" label in N1)
    header["report_no"] = _int(_cell(ws, 1, 16, L))             # P1

    # Superintendant — P2
    header["superintendent"] = _clean(_cell(ws, 2, 16, L))      # P2
    # Tool pusher — P3
    header["supervisor"] = _clean(_cell(ws, 3, 16, L))          # P3
    # Night tool pusher — P4 (extra context, not a separate column)
    header["night_tool_pusher"] = _clean(_cell(ws, 4, 16, L))
    # HSE supervisor — P5
    header["hse_supervisor"] = _clean(_cell(ws, 5, 16, L))
    # Site doctor — P6
    header["medecin"] = _clean(_cell(ws, 6, 16, L))

    # Casing & shoe data (cols D label, F depth, rows 2-7)
    #   D2 CASING 30"      F2 (depth or empty)
    #   D3 CASING 18"5/8   F3
    #   D4 CASING 13"3/8   F4 = 2006
    #   D5 CASING 9"5/8    F5 = 3452
    #   D6 TOP LINER 7"    F6 = 3354
    #   D7 SABOT 7"        F7 = 3663
    casing_rows = []
    for r in range(2, 8):
        label = _clean(_cell(ws, r, 4, L))
        depth = _cell(ws, r, 6, L)
        if label and depth is not None and _clean(depth):
            casing_rows.append((label, _clean(depth)))
    # Most-recent casing = last non-empty
    if casing_rows:
        header["last_casing"] = casing_rows[-1][0] + " @ " + casing_rows[-1][1]
        # Top Shoe = SABOT row if present
        for lbl, dep in casing_rows:
            if "SABOT" in lbl.upper() or "SHOE" in lbl.upper():
                header["top_shoe"] = lbl + " @ " + dep
                break
        # Top Liner
        for lbl, dep in casing_rows:
            if "LINER" in lbl.upper():
                header["top_liner"] = lbl + " @ " + dep
                break
        # Total depth = deepest numeric value seen
        depths = [_float(d) for _, d in casing_rows if _float(d) > 0]
        if depths:
            header["total_depth"] = max(depths)
    else:
        header["last_casing"] = ""
        header["top_shoe"] = ""
        header["top_liner"] = ""
        header["total_depth"] = 0.0

    # Last BOP test — I2 (date)
    header["last_bop_test"] = _date_parse(_cell(ws, 2, 9, L))

    # Operation at 00H00 — R5/R6 text panel (description of current state)
    op_00 = _clean(_cell(ws, 5, 18, L)) or _clean(_cell(ws, 5, 19, L))
    op_00_2 = _clean(_cell(ws, 6, 18, L)) or _clean(_cell(ws, 6, 19, L))
    header["operation_at_0000"] = " ".join(filter(None, [op_00, op_00_2])).strip()

    # =====================================================================
    # OPERATIONS — header at r25, data starts r26.  Cols:
    #   E=FROM  F=TO  G=HRS  H=PHASE  I=OPERATIONS  (description)
    # =====================================================================
    activities = []
    after_midnight_parts = []

    # Find the actual header row in case the layout shifts a row or two
    ops_header_row = None
    for r in range(20, 30):
        e_val = _clean(_cell(ws, r, 5, L)).upper()
        f_val = _clean(_cell(ws, r, 6, L)).upper()
        if e_val == "FROM" and f_val == "TO":
            ops_header_row = r
            break
    if ops_header_row is None:
        ops_header_row = 25  # default

    in_after_midnight = False
    for r in range(ops_header_row + 1, ops_header_row + 25):
        from_v = _cell(ws, r, 5, L)
        to_v = _cell(ws, r, 6, L)
        hrs_v = _cell(ws, r, 7, L)
        desc = _clean(_cell(ws, r, 9, L))

        # Stop if we drift into a different section
        section_marker = _clean(_cell(ws, r, 1, L)).upper()
        if any(m in section_marker for m in (
            "MUD CARAC", "VEHICULES", "MOTOPOMPE", "CITERNE",
            "POMPES DES BACS",
        )):
            # Mud panel etc. lives in col A; don't break here if the ops
            # table runs alongside it. Only break when we're clearly past
            # the ops area.
            pass

        # Detect AFTER MIDNIGHT label (appears in col I around r43)
        if desc.upper() == "AFTER MIDNIGHT":
            in_after_midnight = True
            continue

        # Skip empty rows entirely
        if from_v is None and to_v is None and not desc and hrs_v in (None, 0):
            continue

        if in_after_midnight:
            if desc:
                after_midnight_parts.append(desc)
            continue

        start_t = _time_parse(from_v)
        end_t = _time_parse(to_v)
        hours = _hours_value(hrs_v)

        # Continuation lines (no times, just description) fold into the
        # previous activity rather than creating phantom entries.
        if start_t is None and end_t is None:
            if desc and activities:
                activities[-1]["description"] = (
                    activities[-1]["description"] + " " + desc
                ).strip()
            continue

        # If start/end are set but hours is missing or #REF!, compute it
        if hours == 0 and start_t is not None and end_t is not None:
            sm = start_t.hour * 60 + start_t.minute
            em = end_t.hour * 60 + end_t.minute
            if em == sm == 0:
                # 00:00 → 00:00 with no hours = skip
                pass
            else:
                if em <= sm:
                    em += 1440
                hours = (em - sm) / 60.0

        if not desc and hours == 0:
            continue

        activities.append({
            "start_time": start_t,
            "end_time": end_t,
            "hours": round(hours, 2),
            "phase_name": "",
            "code": "",
            "sub": "",
            "description": desc,
            "start_md": 0,
            "end_md": 0,
            "npt": 0,
            "npt_detail": "",
            "npt_company": "",
            "op_company": "",
            "bill": "",   # back-assigned from tarif totals below
        })

    # =====================================================================
    # TARIF totals — Daily T3 (Q15), Daily NR (Q17).  Everything else of
    # the 24h is T1 (billable) by convention.
    # =====================================================================
    daily_t3 = _hours_value(_cell(ws, 15, 17, L))   # Q15
    daily_nr = _hours_value(_cell(ws, 17, 17, L))   # Q17
    total_hours = sum(a["hours"] for a in activities)
    daily_t1 = max(0.0, total_hours - daily_t3 - daily_nr)

    tarif_totals = {}
    if daily_t1 > 0:
        tarif_totals["T1"] = round(daily_t1, 2)
    if daily_t3 > 0:
        tarif_totals["T3"] = round(daily_t3, 2)
    if daily_nr > 0:
        tarif_totals["NR"] = round(daily_nr, 2)
    # If everything is zero (the file has no NPT today), default to all T1
    if not tarif_totals and total_hours > 0:
        tarif_totals["T1"] = round(total_hours, 2)

    # Back-assign bill codes when none are set per-op.  Simple chronological
    # fill: ops fill T1 first, then T3, then NR — matching most workover days
    # where NPT is a contiguous segment.  The downstream universal
    # assign_bill_codes pass will refine this with subset-sum if richer.
    if all(not a.get("bill") for a in activities) and tarif_totals:
        remaining = dict(tarif_totals)
        for a in activities:
            assigned = False
            for code in ("T1", "T2", "T3", "T4", "NR"):
                if remaining.get(code, 0) >= a["hours"] - 0.01:
                    a["bill"] = code
                    remaining[code] -= a["hours"]
                    assigned = True
                    break
            if not assigned:
                # Pick the code with the most remaining capacity
                code = max(remaining, key=lambda k: remaining[k]) if remaining else "T1"
                a["bill"] = code
                remaining[code] = max(0, remaining.get(code, 0) - a["hours"])

    # =====================================================================
    # AFTER MIDNIGHT — already collected above; format as text section
    # =====================================================================
    text_sections = {}
    if after_midnight_parts:
        text_sections["after_midnight"] = " ".join(after_midnight_parts).strip()[:300]

    # Operation-at-00:00 free text from R5/R6 → "situation"
    if header.get("operation_at_0000"):
        text_sections["situation"] = header["operation_at_0000"][:300]

    # =====================================================================
    # SAFETY counters — Free Days (J62), Last Accident (J64)
    # =====================================================================
    safety = {}
    safety["accident_free_days"] = _int(_cell(ws, 62, 10, L))
    safety["last_accident_date"] = _date_parse(_cell(ws, 64, 10, L))

    # =====================================================================
    # MUD characteristics (rows 26-50, col A label, col B value)
    # =====================================================================
    mud_checks = {}
    mud_label_map = {
        "DENSITY": "density",
        "MARSH VISC": "marsh_viscosity",
        "HTHP-FILTRATE": "filtrate",
        "CAKE": "cake",
        "PH": "ph",
        "YIELD POINT": "yield_point",
        "PLASTIC VISC": "plastic_viscosity",
        "TOTAL VOLUME": "total_volume_m3",
        "GEL.0": "gel_0",
        "GEL.10": "gel_10",
        "SOLID %": "solid_pct",
        "OIL %": "oil_pct",
    }
    for r in range(26, 53):
        label = _clean(_cell(ws, r, 1, L)).upper()
        value = _cell(ws, r, 2, L)
        if not label or value is None:
            continue
        for key, slug in mud_label_map.items():
            if label.startswith(key):
                mud_checks[slug] = _clean(value) or _float(value)
                break

    wb.close()

    return {
        "header": header,
        "activities": activities,
        "text_sections": text_sections,
        "mud_checks": mud_checks,
        "mud_volume": {},
        "mud_chemical_usage": [],
        "personnel_data": [],
        "pumps": [],
        "well_location": {},
        "survey_data": [],
        "safety": safety,
        "tarif_totals": tarif_totals,
    }


# Drop-in compat with the rest of the pipeline
parse_daily_excel_report = parse_tp127
parse_ddr = parse_tp127


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: tp127_extract.py SOURCE.xlsx")
    data = parse_tp127(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)):
            return o.isoformat()
        if isinstance(o, time):
            return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta):
            return o.total_seconds()
        return str(o)

    print(json.dumps(data, indent=2, default=default))
