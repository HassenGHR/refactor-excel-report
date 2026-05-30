#!/usr/bin/env python3
"""
entp204_extract.py — extract an ENTP-204 (ENTP rig 204, AIN T'SILA field,
TXNO wells) Daily Workover Report into the standard dict shape.

Source layout
-------------
SONATRACH AIN T'SILA "DAILY WORKOVER REPORT" — English, single sheet
(~127 rows × 19 cols).  Same template family as TP-182 / TP-195 but the
header cells store "LABEL : value" as a SINGLE combined string (e.g.
"WELL : TXNO-07", "RIG NAME : ENTP 204") rather than split label/value
cells.  The sheet repeats the header block twice (page 1 rows 5-9, page 2
rows 60-64); we read from the first occurrence.

Header (row 5-7, combined "LABEL : value" strings):
    r5  WELL : TXNO-07 | DATE : 05/05/2026 | TOTAL MD | TOTAL TVD | FORMATION TOP | REP N° : 09
    r6  Supervisor: L.BOUHENACHE +N.Ayad | RIG NAME : ENTP 204 | FIELD : AIN T'SILA |
        Superintandant: / | ACC FREE : | Daily NPT : 00 Hrs
    r7  Last Formation Test | Last Csg Shoe | Last Liner | L,BOP 27/04/2026 |
        Last Safety Meeting | Cum NPT
    r8  WORK-OVER REASON: SHORT-RADIUS DRILLING (RE-ENTRY)

Operations table (header at row 78): FROM | TO | HRS | DESCRIPTION | … | BILL | COMPANY
    r79+  0:00:00 | 1 day, 0:00:00 | | <desc> | … | T2 | ENTP
    (TO of "1 day, 0:00:00" = 24:00; an after-midnight row may follow)

Tarif totals at row 107 (combined strings):
    T1 =00h00 | T2 =24h00 | T3 = CUM 00h00 | T4 = 00h00 | FT = 00h00 | NR = 00h00

Text sections:
    ACTUEL OPERATIONS: …   (row 108)
    PLAN OPERATIONS: …     (row 110)
    REMARKS: …             (row 114)

Two supervisor names appear joined with '+' in the Supervisor cell
("L.BOUHENACHE +N.Ayad") → split into supervisor + superintendent per
the two-supervisor convention.

Distinguishing markers (used by helpers.parse_source._detect_format_xlsx):
    - "AIN T'SILA" or "AIN TSILA"
    - "ENTP 204" / "ENTP204"
    - well "TXNO" prefix
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Union, List, Optional

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
    """From a 'LABEL : value' combined string, return the value part.
    Matches label case-insensitively; tolerates ':' or '：' separators."""
    s = _clean(cell_text)
    rx = re.compile(rf"^{re.escape(label)}\s*[:：]\s*(.*)$", re.IGNORECASE)
    m = rx.match(s)
    if m:
        return m.group(1).strip()
    return ""


def _scan_label_value(ws, L, label, row_range=(1, 65), col_range=(1, 20)) -> str:
    """Scan a region for a cell whose text starts with 'LABEL :' and
    return the value part."""
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


def _cell(ws, r, c, lookup):
    v = ws.cell(r, c).value
    return v if v is not None else lookup.get((r, c))


def _normalize_rig(name: str) -> str:
    """'ENTP 204' / 'ENTP-204' → 'ENTP-204'."""
    s = _clean(name).upper()
    m = re.search(r"ENTP[\s#\-]*(\d+)", s)
    if m: return f"ENTP-{m.group(1)}"
    m = re.search(r"TP[\s#\-]*(\d+)", s)
    if m: return f"TP-{m.group(1)}"
    return _clean(name)


def _parse_op_time(v) -> Optional[time]:
    """Parse operation time cells.  Handles datetime.time, '0:00:00',
    '1 day, 0:00:00' (= 24:00 → time(0,0)), and timedelta."""
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
    # "1 day, 0:00:00" → midnight (24:00)
    if "day" in s.lower():
        return time(0, 0)
    m = re.match(r"^(\d{1,2})\s*[:hH]\s*(\d{2})(?:\s*[:hH]\s*\d{2})?$", s)
    if m:
        return time(int(m.group(1)) % 24, int(m.group(2)))
    return None


def _prev_end_minutes(prev_op) -> Optional[int]:
    """Recover the previous op's end time as minutes-since-midnight,
    reconstructing it from start + hours so that a midnight end (stored as
    time(0,0)) is correctly treated as 1440 rather than 0."""
    st = prev_op.get("start_time")
    if st is None:
        return None
    start_min = st.hour * 60 + st.minute
    hours = prev_op.get("hours", 0.0) or 0.0
    return int(round(start_min + hours * 60))


def _tod_minutes(v) -> Optional[int]:
    """Convert an operation time cell to minutes-since-midnight as an
    ABSOLUTE time-of-day.  Crucially, an end-of-day midnight is encoded in
    this template as "1 day, 0:00:00" (timedelta of 24h) or the text
    "24:00", and must map to 1440 (= 24:00), NOT 0 — otherwise an op like
    21:00 → 00:00 would be mis-measured as 24h instead of 3h.

    Returns minutes (0–1440) or None if unparseable.
    """
    if v is None:
        return None
    if isinstance(v, timedelta):
        total = int(v.total_seconds())
        # Cap at 1440 (a full day's worth of minutes = 24:00 / midnight-end)
        return min(total // 60, 1440)
    if isinstance(v, time):
        return v.hour * 60 + v.minute
    if isinstance(v, datetime):
        return v.hour * 60 + v.minute
    s = _clean(v)
    if not s:
        return None
    # "1 day, 0:00:00" → 1440
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


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_entp204(source: Union[Path, str, BytesIO]) -> dict:
    if isinstance(source, (str, Path)):
        wb = load_workbook(source, data_only=True)
    else:
        wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)

    header = {}

    # =====================================================================
    # HEADER — combined "LABEL : value" cells, page-1 block (rows 5-9)
    # =====================================================================
    well = _scan_label_value(ws, L, "WELL", (1, 12))
    if well: header["well_name"] = well

    d = _scan_label_value(ws, L, "DATE", (1, 12))
    if d:
        dd = _date_parse(d)
        if dd: header["date"] = dd

    rep = _scan_label_value(ws, L, "REP N°", (1, 12)) or _scan_label_value(ws, L, "REP N", (1, 12))
    if rep:
        n = _int(rep, 0)
        if n > 0: header["day_number"] = n

    rig = _scan_label_value(ws, L, "RIG NAME", (1, 12))
    if rig: header["rig_name"] = _normalize_rig(rig)

    field = _scan_label_value(ws, L, "FIELD", (1, 12))
    if field: header["field_name"] = field

    # MD / TVD / formation
    md = _scan_label_value(ws, L, "TOTAL MD", (1, 12))
    if md:
        v = _float(md)
        if v: header["well_md"] = v
    tvd = _scan_label_value(ws, L, "TOTAL TVD", (1, 12))
    if tvd:
        v = _float(tvd)
        if v: header["tvd"] = v

    # Supervisors — "Supervisor: L.BOUHENACHE +N.Ayad"
    sup_raw = _scan_label_value(ws, L, "Supervisor", (1, 12))
    if sup_raw:
        # Split on '+' or '&' or ',' into up to two names
        names = re.split(r"\s*[+&,]\s*", sup_raw)
        names = [n.strip() for n in names if n.strip()]
        if names:
            header["supervisor"] = names[0]
        if len(names) >= 2:
            header["superintendent"] = names[1]
    # If the explicit "Superintandant:" field has a real name, prefer it
    supt_raw = _scan_label_value(ws, L, "Superintandant", (1, 12))
    if supt_raw and supt_raw not in ("/", "-", ""):
        header["superintendent"] = supt_raw

    # Casing / liner / formation test / BOP
    csg = _scan_label_value(ws, L, "Last Csg Shoe", (1, 12))
    if csg and csg.replace("\\", "").strip() not in ("", "m"):
        header["top_shoe"] = csg
    liner = _scan_label_value(ws, L, "Last Liner", (1, 12))
    if liner and liner.replace("\\", "").strip() not in ("", "m"):
        header["last_csg_top"] = liner
    bop = _scan_label_value(ws, L, "L,BOP", (1, 12)) or _scan_label_value(ws, L, "L.BOP", (1, 12))
    if not bop:
        # The BOP cell uses spaces, not a colon: "L,BOP   27/04/2026"
        for r in range(1, 12):
            for c in range(1, 16):
                v = _clean(_cell(ws, r, c, L) or "")
                if re.match(r"^L[,.]?\s*BOP\b", v, re.IGNORECASE):
                    bd = _date_parse(v)
                    if bd:
                        header["bop_test"] = bd
                    break
            if header.get("bop_test"): break
    else:
        bd = _date_parse(bop)
        if bd: header["bop_test"] = bd

    # Accident-free / NPT
    acc = _scan_label_value(ws, L, "ACC FREE", (1, 12))
    if acc:
        header["accident_free_days"] = _int(acc, 0)

    # Work-over reason → objective
    reason = _scan_label_value(ws, L, "WORK-OVER REASON", (1, 12))
    if reason:
        header["well_objective"] = reason

    # =====================================================================
    # OPERATIONS — table header "FROM | TO | HRS | DESCRIPTION | ... BILL | COMPANY"
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

            # Stop when we reach the tarif-totals / section labels
            b_text = _clean(_cell(ws, r, 2, L) or "").upper()
            if any(m in b_text for m in ("T1 =", "ACTUEL", "PLAN OPER",
                                          "REMARKS", "REQUIREMENTS", "SURVEY",
                                          "WELL :", "SONATRACH")):
                break

            # "AFTER MIDNIGHT" (often misspelled "AFTER MIDNIGTH") is a
            # section label: everything below it belongs to the next day's
            # first hours.  This is the ONLY boundary between same-day ops
            # and after-midnight ops — we must NOT infer the boundary from a
            # missing bill code, because same-day ops legitimately lack bill
            # codes mid-day (they get back-assigned from the tarif totals).
            if "AFTER MIDN" in desc.upper():
                in_after_midnight = True
                continue

            # Skip the repeated header row and obvious label leaks
            if desc.upper() in ("DESCRIPTION", "") and from_v is None and to_v is None:
                continue
            if from_v is None and to_v is None and not desc:
                continue

            if in_after_midnight:
                # Capture after-midnight description text (don't make it an op)
                if desc:
                    after_midnight_parts.append(desc)
                continue

            sm = _tod_minutes(from_v)
            em = _tod_minutes(to_v)

            # Some source files have a data-entry quirk where an op's FROM is
            # left at 00:00 (or otherwise earlier than the previous op's TO),
            # creating an overlap that inflates the daily total past 24h.
            # Snap the start to the previous op's end so the timeline is
            # contiguous.  Only do this when it produces a sane forward span.
            if activities and sm is not None:
                prev_em = _tod_minutes(activities[-1]["end_time"])
                # prev end stored as time(0,0) for a midnight-end would read 0;
                # recompute from the stored hours instead when needed.
                prev_end_min = _prev_end_minutes(activities[-1])
                if prev_end_min is not None and sm < prev_end_min <= (em if em is not None else 1440):
                    sm = prev_end_min

            start_t = _parse_op_time(from_v)
            end_t   = _parse_op_time(to_v)
            # Reflect any snap in the displayed start time
            if sm is not None:
                start_t = time((sm // 60) % 24, sm % 60)

            if sm is not None and em is not None:
                if em > sm:
                    hours = (em - sm) / 60.0
                elif em == sm:
                    # identical → only a real op if it's the lone 00:00→00:00
                    hours = 24.0 if sm == 0 else 0.0
                else:
                    # end < start: wrapped past midnight (rare in this template)
                    hours = (em + 1440 - sm) / 60.0

                activities.append({
                    "start_time": start_t,
                    "end_time":   end_t,
                    "hours":      hours,
                    "phase_name": "",
                    "code": "", "sub": "",
                    "description": desc,
                    "start_md": 0, "end_md": 0,
                    "npt": 0, "npt_detail": "",
                    "npt_company": "", "op_company": comp,
                    "bill": bill,
                })
            elif desc and activities and desc.upper() != "DESCRIPTION":
                # Genuine continuation line for the current op
                activities[-1]["description"] = (
                    activities[-1]["description"] + "\n" + desc
                ).strip()

    # =====================================================================
    # TARIF TOTALS — row with "T1 =00h00 | T2 =24h00 | ..."
    # =====================================================================
    tarif_totals = {}
    for r in range(100, 115):
        b = _clean(_cell(ws, r, 2, L) or "")
        if re.match(r"^T1\s*=", b):
            # This row holds all the T-codes across columns
            for c in range(2, 20):
                cell_text = _clean(_cell(ws, r, c, L) or "")
                m = re.match(r"^(T[1-4])\s*=\s*(.*)$", cell_text)
                if m:
                    code = m.group(1).lower()
                    hrs = _hhmm_to_hours(m.group(2))
                    if hrs > 0:
                        tarif_totals[code] = hrs
            break

    # Back-assign bill codes for any operations that don't already carry
    # one.  Same-day ops frequently leave the BILL column blank mid-day
    # (only some rows are tariffed in the source); the tarif-totals row
    # gives the daily breakdown.  assign_bill_codes keeps existing per-op
    # codes and partitions the remaining unbilled hours into the leftover
    # T-buckets.
    has_unbilled = any(not _clean(a.get("bill")) for a in activities)
    if has_unbilled and tarif_totals:
        from helpers.bill_code_assign import assign_bill_codes
        activities = assign_bill_codes(activities, tarif_totals)

    # If the ops table has per-op bills but no hours yet, and tarif_totals
    # tells us the breakdown, the universal normalize pass + insert handle
    # the codes.  But if an op got hours=0 with a real bill, fix the single
    # 24h op case: when there's exactly one billed op covering the day.
    if len(activities) == 1 and activities[0]["hours"] == 0.0:
        # Single op spanning the whole day
        total_tarif = sum(tarif_totals.values())
        activities[0]["hours"] = total_tarif if total_tarif else 24.0

    # =====================================================================
    # MUD TYPE
    # =====================================================================
    mud_checks = {}
    for r in range(10, 13):
        for c in range(13, 20):
            v = _clean(_cell(ws, r, c, L) or "")
            if "OBM" in v.upper() or "WATER BASE" in v.upper():
                mud_checks["mud_type"] = v
                break
        if mud_checks: break

    # =====================================================================
    # TEXT SECTIONS — ACTUEL OPERATIONS / PLAN OPERATIONS / REMARKS
    # =====================================================================
    text_sections = {}
    actuel = _scan_label_value(ws, L, "ACTUEL OPERATIONS", (100, 120))
    if actuel:
        v = actuel
        if len(v) > 300:
            v = v[:300].rsplit(None, 1)[0] + "…"
        text_sections["current_operation"] = v
        text_sections["day_summary"]       = v
    plan = _scan_label_value(ws, L, "PLAN OPERATIONS", (100, 120))
    if plan:
        text_sections["plan_operations"] = plan
    remarks = _scan_label_value(ws, L, "REMARKS", (100, 122))
    if remarks:
        text_sections["remarks"] = remarks
    if after_midnight_parts:
        text_sections["after_midnight"] = " ".join(after_midnight_parts)

    # =====================================================================
    # PERSONNEL (header at "PERSONNEL DATA", table below)
    # =====================================================================
    personnel = []
    pers_row = None
    for r in range(65, 80):
        if "PERSONNEL DATA" in _clean(_cell(ws, r, 2, L) or "").upper():
            pers_row = r
            break
    if pers_row:
        for r in range(pers_row + 2, pers_row + 12):
            comp_l = _clean(_cell(ws, r, 2, L) or "")
            # Stop when we hit the next section header or the ops table
            if comp_l.upper() in ("FROM", "DESCRIPTION") or comp_l.upper().startswith(
                    ("WORK-OVER", "SURVEY", "BIT DATA", "MUD CHEMICAL")):
                break
            no_l   = _cell(ws, r, 4, L)
            comp_r = _clean(_cell(ws, r, 10, L) or "")
            no_r   = _cell(ws, r, 12, L)
            names_r = _clean(_cell(ws, r, 14, L) or "")
            # A blank left-and-right row signals the end of the table
            if not comp_l and not comp_r:
                continue
            if comp_l and comp_l.upper() not in ("COMPANY", "PERSONNEL DATA"):
                personnel.append({"company": comp_l, "number": _int(no_l, 0),
                                  "hours": "", "names": ""})
            if comp_r and comp_r.upper() not in ("COMPANY",):
                personnel.append({"company": comp_r, "number": _int(no_r, 0),
                                  "hours": "", "names": names_r})

    # =====================================================================
    # SURVEY DATA
    # =====================================================================
    survey = []
    surv_row = None
    for r in range(63, 72):
        if "SURVEY DATA" in _clean(_cell(ws, r, 2, L) or "").upper():
            surv_row = r
            break
    if surv_row:
        # header at surv_row+1, values at surv_row+2
        vals = [_cell(ws, surv_row + 2, c, L) for c in (2, 4, 6, 8, 10, 12, 14, 16)]
        if any(v is not None for v in vals):
            survey.append({
                "md": _float(vals[0]), "tvd": _float(vals[1]),
                "azimuth": _float(vals[2]), "inclination": _float(vals[3]),
                "dls": _float(vals[4]), "ns": _float(vals[5]),
                "ew": _float(vals[6]), "vs": _float(vals[7]),
            })

    wb.close()

    return {
        "header": header,
        "activities": activities,
        "text_sections": text_sections,
        "mud_checks": mud_checks,
        "mud_volume": {},
        "mud_chemical_usage": [],
        "personnel_data": personnel,
        "pumps": [],
        "well_location": {},
        "survey_data": survey,
        "safety": {},
        "tarif_totals": tarif_totals,
    }


# Drop-in compat
parse_daily_excel_report = parse_entp204


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: entp204_extract.py SOURCE.xlsx")
    data = parse_entp204(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))