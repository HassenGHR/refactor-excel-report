#!/usr/bin/env python3
"""
enf06_extract.py — extract an ENF#06 (ENAFOR rig 06, Hassi Guettar field,
HGAS wells) Daily Workover Report (.xlsx) into the standard dict shape.

Source layout
-------------
Single-sheet French/English template (~70 rows × 254 cols, sparse).
Wide ENAFOR "DAILY WORKOVER REPORT" format.

  Row 2   : A2 = ' '
  Row 3   : I3 = "E.NA.FOR DAILY WORKOVER REPORT"
  Row 5   : E5 = "TP.Senior", F5 = senior name
  Row 6   : A6 = "ENF# 06", E6 = "TP.Junior", F6 = junior name,
            R6 = "Date", S6 = date string
  Row 7   : E7 = "Field", G7 = "Well name", I7 = "Present depth",
            K7 = "Previous depth", etc.
  Row 8   : E8 = field name, G8 = well name, I8 = present depth (numeric)
  Row 9   : HSE block: B9 = "Accident free Days", E9 = count
  Row 11  : A11 = "SURVEY", B11-E11 = survey headers
  Row 12  : O12 = "Last BOP Test", Q12 = date
  Row 14  : A14 = "Last Csg:", C14 = "Size", D14 = "7"",
            E14 = "KOP", G14 = "BHA details", I14 = BHA text (wide)
  Row 15-16: C15/C16 = SHOE/TOL labels with numeric values at D/E/F
  Row 17-20: Personnel rows: row 17 = labels, row 18 = counts,
             row 19 = labels, row 20 = counts
            Layout: A=Supt, C=S/tool, E=J/tool, G=Driller, I=A/Driller,
                    K=Cont.lab, M=Cas.lab, N=Vigiles, P=DTM/PERS,
                    Q=Other, S=Formateur
            Row 19: A=Elect, C=Mech, E=Transport, G=Catering, I=Company man,
                    K=Trainee, M=HSE, N=ADG, P=Medecin, Q=Intendant,
                    R=Welder, S=Forklift
  Row 21  : A21 = "Timing" (merged A:C), D21 = "OPERATIONS" (merged D:L),
            M21 = "Rate" (merged M:Q? actually M21:Q22 "MUD PUMPS DATA"),
            N21 = "MUD PUMPS DATA" merged N:Q
  Row 22  : A22 = "Start", B22 = "End", C22 = "Hrs"
  Rows 23+: Operations: A=start, B=end, C=hours, D=description (merged D:L),
            M=rate/bill code, N-Q = mud pump data / mud properties
  Row 37  : D37 = "AFTER MIDNIGHT OPERATIONS"
  Row 39+: After-midnight ops (same cols)
  Row 47  : D47 = "PLAN OPERATIONS:"
  Rows 44-53: Genset/equipment data (cols N-Q)
  Rows 52-69: KPI, equipment inventory, cost data
            A62 = "Daily cost", C62 = value
  Row 61  : A61 = "Report time" (well status info)
  Rows 62-68: Cost/NPT rows

Distinguishing markers (helpers.parse_source._detect_format_xlsx):
    - "E.NA.FOR" + "DAILY WORKOVER REPORT"
    - "ENF# 06" or "ENF#06"
    - "TP.Senior" + "TP.Junior" headers
    - "Timing" + "MUD PUMPS DATA" ops headers
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Union

from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _clean(v) -> str:
    if v is None: return ""
    s = str(v)
    if s.strip() in ("#VALUE!", "#REF!", "#NAME?", "#N/A", "#DIV/0!", "#NULL!"):
        return ""
    return re.sub(r"\s+", " ", s).strip()


def _float(v, default=0.0) -> float:
    if v is None: return default
    if isinstance(v, (int, float)): return float(v)
    s = _clean(v).replace(",", ".").replace(" ", "")
    if s in ("", "-", "/", "None"): return default
    s = re.sub(r"[a-zA-Zé°%/³]+$", "", s).strip()
    try: return float(s)
    except ValueError: return default


def _int(v, default=0) -> int:
    return int(_float(v, float(default)))


def _date_parse(v):
    if v is None: return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date_type): return v
    s = _clean(v)
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: continue
    return None


def _time_parse(v):
    if v is None: return None
    if isinstance(v, time): return v
    if isinstance(v, datetime):
        if v.year == 1900 and v.month == 1 and v.day == 1:
            return v.time()
        return v.time()
    if isinstance(v, timedelta):
        total = int(v.total_seconds())
        if total == 86400: return time(0, 0)
        return time((total // 3600) % 24, (total % 3600) // 60)
    s = _clean(v)
    if not s or s in ("-", "None"): return None
    if s in ("24:00", "24:00:00", "24h00", "24H00", "24h"): return time(0, 0)
    m = re.match(r"^(\d{1,2})[hH](\d{0,2})$", s)
    if m: return time(int(m.group(1)) % 24, int(m.group(2) or 0))
    for fmt in ("%H:%M:%S", "%H:%M", "%Hh%M"):
        try: return datetime.strptime(s, fmt).time()
        except ValueError: continue
    return None


def _duration_hours(v) -> float:
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    if isinstance(v, timedelta): return v.total_seconds() / 3600.0
    if isinstance(v, time): return v.hour + v.minute / 60.0
    if isinstance(v, datetime):
        if v.year == 1900 and v.month == 1 and v.day == 1:
            return v.hour + v.minute / 60.0
        return 0.0
    t = _time_parse(v)
    if t: return t.hour + t.minute / 60.0
    return _float(v)


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
    """'ENF# 06' / 'ENF#06' -> 'ENF#06'."""
    s = _clean(name).upper()
    m = re.search(r"ENF\s*#?\s*(\d+)", s)
    if m: return f"ENF#{int(m.group(1)):02d}"
    return _clean(name)


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_enf06(source: Union[Path, str, BytesIO]) -> dict:
    if isinstance(source, (str, Path)):
        wb = load_workbook(source, data_only=True)
    else:
        wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)

    header = {}

    # =====================================================================
    # HEADER
    # =====================================================================
    # Row 6: rig, date
    header["rig_name"] = _normalize_rig(_cell(ws, 6, 1, L) or "")  # A6 "ENF# 06"
    header["date"] = _date_parse(_cell(ws, 6, 19, L))              # S6 "27-05-2026"

    # Row 8: field, well, present depth
    header["field_name"] = _clean(_cell(ws, 8, 5, L))              # E8
    header["well_name"] = _clean(_cell(ws, 8, 7, L))               # G8

    present_depth = _cell(ws, 8, 9, L)                             # I8
    if present_depth is not None:
        header["well_md"] = _float(present_depth)

    previous_depth = _cell(ws, 8, 11, L)                            # K8
    if previous_depth is not None:
        header["tmd"] = _float(previous_depth)

    # Overall-days at O8
    overall_days = _cell(ws, 8, 15, L)                              # O8
    if overall_days is not None:
        header["day_number"] = _int(overall_days)

    # HSE — Accident free Days (B9 label, E9 value)
    header["accident_free_days"] = _int(_cell(ws, 9, 5, L))        # E9

    # BOP test
    bop_date = _date_parse(_cell(ws, 12, 17, L))                   # Q12
    if bop_date:
        header["bop_test"] = bop_date

    # Supervisors
    tp_senior = _clean(_cell(ws, 5, 6, L))                         # F5
    tp_junior = _clean(_cell(ws, 6, 6, L))                         # F6
    if tp_senior:
        header["supervisor"] = tp_senior
    if tp_junior:
        header["superintendent"] = tp_junior

    # Last casing / BHA
    csg_size = _clean(_cell(ws, 14, 4, L))                         # D14 "7""
    csg_shoe_val = _cell(ws, 15, 4, L)                             # D15 (shoe depth)
    if csg_size or csg_shoe_val is not None:
        if csg_shoe_val is not None:
            try:
                d = int(float(str(csg_shoe_val)))
                header["top_shoe"] = f'{csg_size} @ {d}m' if csg_size else f'@ {d}m'
            except (ValueError, TypeError):
                header["top_shoe"] = f'{csg_size} @ {_clean(csg_shoe_val)}m'
        else:
            header["top_shoe"] = csg_size

    # KOP
    kop = _cell(ws, 15, 5, L)                                      # E15
    if kop is not None:
        try: header["kop"] = _float(kop)
        except: pass

    # Last CSG TOP / TOL
    tol = _cell(ws, 16, 4, L)                                      # D16
    if tol is not None:
        try: header["last_csg_top"] = f"@ {int(float(str(tol)))}m"
        except: pass

    # BHA details (wide merged cell starting at I14)
    bha_raw = _clean(_cell(ws, 14, 9, L) or "")                    # I14 (merged I14:N16)
    if bha_raw:
        header["bha_details"] = bha_raw

    # =====================================================================
    # PERSONNEL (rows 17-20, two-row label/value layout)
    # Row 16 labels: A=Supt, C=S/tool, E=J/tool, G=Driller, I=A/Driller,
    #                K=Cont.lab, M=Cas.lab, N=Vigiles, P=DTM/PERS, Q=Other, S=Formateur
    # Row 17 values (counts)
    # Row 19 labels: A=Elect, C=Mech, E=Transport, G=Catering, I=Company man,
    #                K=Trainee, M=HSE, N=ADG, P=Medecin, Q=Intendant, R=Welder, S=Forklift
    # Row 20 values (counts)
    # =====================================================================
    personnel = []
    pers_map = [
        # (label_row, label_col, value_row, value_col)
        (17, 1, 18, 1),    # Supt
        (17, 3, 18, 3),    # S/toolpusher
        (17, 5, 18, 5),    # J/toolpusher
        (17, 7, 18, 7),    # Driller
        (17, 9, 18, 9),    # A/Driller
        (17, 11, 18, 11),  # Cont .lab
        (17, 13, 18, 13),  # Cas.lab
        (17, 14, 18, 14),  # Vigiles
        (17, 16, 18, 16),  # DTM / PERS
        (17, 17, 18, 17),  # Other
        (17, 19, 18, 19),  # Formateur
        (19, 1, 20, 1),    # Elect
        (19, 3, 20, 3),    # Mechanic
        (19, 5, 20, 5),    # Transport
        (19, 7, 20, 7),    # Catering
        (19, 9, 20, 9),    # Company man
        (19, 11, 20, 11),  # Trainee
        (19, 13, 20, 13),  # HSE
        (19, 14, 20, 14),  # ADG
        (19, 16, 20, 16),  # Medecin+inf
        (19, 17, 20, 17),  # Intendant
        (19, 18, 20, 18),  # Welder
        (19, 19, 20, 19),  # Forklift
    ]
    for lr, lc, vr, vc in pers_map:
        label = _clean(_cell(ws, lr, lc, L) or "")
        val = _cell(ws, vr, vc, L)
        if not label or val is None:
            continue
        text = _clean(str(val))
        if not text:
            continue
        # Parse counts: pure number, "X+Y" (sum), or text fallback
        if re.match(r"^\d+(\.\d+)?$", text):
            n = _int(text)
        elif re.match(r"^[\d.]+\s*\+\s*[\d.]+(\s*\+\s*[\d.]+)*$", text):
            n = sum(_int(p.strip()) for p in text.split("+"))
        elif re.match(r"^[\d.]+(/[\d.]+)+$", text):
            n = sum(_int(p.strip()) for p in text.split("/"))
        else:
            n = 0
        names_field = "" if re.match(r"^\d+(\.\d+)?$", text) else text
        personnel.append({
            "company": label,
            "number": n,
            "hours": "",
            "names": names_field,
        })

    # =====================================================================
    # OPERATIONS
    # Header at row 21: A21="Timing" (merged A:C), D21="OPERATIONS" (merged D:L)
    # Sub-headers row 22: A=Start, B=End, C=Hrs
    # Data rows 23+: A=start, B=end, C=hours, D=description (merged D:L),
    #               M=rate/bill code
    # Mud data interleaved in cols N-Q on same rows
    # Stop at "AFTER MIDNIGHT OPERATIONS" (D37) for same-day ops
    # =====================================================================
    activities = []
    after_midnight_parts = []
    in_after_midnight = False

    for row in range(23, ws.max_row + 1):
        start = _cell(ws, row, 1, L)                                # A
        end   = _cell(ws, row, 2, L)                                # B
        hrs   = _cell(ws, row, 3, L)                                # C
        desc  = _clean(_cell(ws, row, 4, L) or "")                  # D (merged D:L)
        bill  = _clean(_cell(ws, row, 13, L) or "")                 # M

        # Detect section boundaries
        if "AFTER MIDNIGHT" in desc.upper():
            in_after_midnight = True
            continue
        if "PLAN OPERATIONS" in desc.upper():
            break
        if "BESOIN URGENT" in desc.upper():
            break

        if in_after_midnight:
            if desc:
                after_midnight_parts.append(desc)
            # After-midnight ops may also have timing
            start_t = _time_parse(start)
            end_t = _time_parse(end)
            if start_t is not None and end_t is not None:
                hours = _duration_hours(hrs) if hrs is not None else 0.0
                if hours == 0.0:
                    sm = start_t.hour * 60 + start_t.minute
                    em = end_t.hour * 60 + end_t.minute
                    if em == sm: hours = 24.0
                    elif em > sm: hours = (em - sm) / 60.0
                    else: hours = (em + 1440 - sm) / 60.0
                activities.append({
                    "start_time": start_t,
                    "end_time": end_t,
                    "hours": hours,
                    "phase_name": "",
                    "code": "", "sub": "",
                    "description": desc,
                    "start_md": 0, "end_md": 0,
                    "npt": 0, "npt_detail": "",
                    "npt_company": "", "op_company": "",
                    "bill": bill,
                })
            continue

        # Skip empty rows
        if start is None and end is None and not desc and not bill:
            continue

        start_t = _time_parse(start)
        end_t = _time_parse(end)

        if start_t is not None and end_t is not None:
            hours = 0.0
            if isinstance(hrs, (int, float)):
                hours = float(hrs)
            elif hrs is not None:
                hours = _duration_hours(hrs)
            if hours == 0.0:
                sm = start_t.hour * 60 + start_t.minute
                em = end_t.hour * 60 + end_t.minute
                if em == sm: hours = 24.0
                elif em > sm: hours = (em - sm) / 60.0
                else: hours = (em + 1440 - sm) / 60.0

            activities.append({
                "start_time": start_t,
                "end_time": end_t,
                "hours": hours,
                "phase_name": "",
                "code": "", "sub": "",
                "description": desc,
                "start_md": 0, "end_md": 0,
                "npt": 0, "npt_detail": "",
                "npt_company": "", "op_company": "",
                "bill": bill,
            })
        elif desc and activities:
            # Continuation row — fold into previous op's description
            activities[-1]["description"] = (
                activities[-1]["description"] + "\n" + desc
            ).strip()

    # =====================================================================
    # TARIF TOTALS — sum per-op hours by bill code
    # =====================================================================
    tarif_totals = {}
    for a in activities:
        code = _clean(a.get("bill")).lower()
        if re.match(r"^t\d+$", code):
            tarif_totals[code] = tarif_totals.get(code, 0.0) + a.get("hours", 0.0)

    # =====================================================================
    # MUD CHECKS  (rows 30-35, cols N-Q)
    #   Row 30: N=LOSS label, P=VISC label, Q=45 (VISC value)
    #   Row 31: N=HOLE VOLUME, O=98, P=YIELD POINT, Q=14
    #   Row 32: N=ACTIVE, O=34, P=OIL, Q=0.9
    #   Row 33: N=RESERVE, O=116, P=WATER, Q=0.1
    #   Row 34: N=TOTAL, O=248, P=SOLIDS, Q=0.06
    #   Row 35: N=DENSITY, O=1.45, P=SAND, Q=0.001
    # =====================================================================
    mud_checks = {}
    mud_volume = {}

    # Row 30: VISC at Q30
    v = _cell(ws, 30, 17, L)                                       # Q30
    if v is not None: mud_checks["fun_vis"] = _float(v)

    # Row 31: HOLE VOLUME at O31, YIELD POINT at Q31
    v = _cell(ws, 31, 15, L)                                       # O31
    if v is not None: mud_volume["string_volume"] = _float(v)
    v = _cell(ws, 31, 17, L)                                       # Q31
    if v is not None: mud_checks["yp"] = _float(v)

    # Row 32: ACTIVE at O32, OIL at Q32
    v = _cell(ws, 32, 15, L)                                       # O32
    if v is not None: mud_volume["active_volume"] = _float(v)
    v = _cell(ws, 32, 17, L)                                       # Q32
    if v is not None: mud_checks["oil"] = _float(v)

    # Row 33: RESERVE at O33, WATER at Q33
    v = _cell(ws, 33, 15, L)                                       # O33
    if v is not None: mud_volume["reserve_volume"] = _float(v)
    v = _cell(ws, 33, 17, L)                                       # Q33
    if v is not None: mud_checks["h2o"] = _float(v)

    # Row 34: TOTAL at O34, SOLIDS at Q34
    v = _cell(ws, 34, 15, L)                                       # O34
    if v is not None: mud_volume["total_volume"] = _float(v)
    v = _cell(ws, 34, 17, L)                                       # Q34
    if v is not None: mud_checks["solid"] = _float(v)

    # Row 35: DENSITY at O35, SAND at Q35
    v = _cell(ws, 35, 15, L)                                       # O35
    if v is not None: mud_checks["density"] = _float(v)
    v = _cell(ws, 35, 17, L)                                       # Q35
    if v is not None: mud_checks["sand"] = _float(v)

    # =====================================================================
    # TEXT SECTIONS
    # =====================================================================
    text_sections = {}

    # After midnight
    if after_midnight_parts:
        text_sections["after_midnight"] = " | ".join(after_midnight_parts)

    # Plan operations (D47)
    plan = _clean(_cell(ws, 47, 4, L) or "")                       # D47
    if plan:
        plan_text = re.sub(r"^PLAN\s+OPERATIONS\s*:?\s*", "", plan, flags=re.IGNORECASE)
        if plan_text:
            text_sections["plan_operations"] = plan_text

    # Situation — last op description (capped 300)
    if activities:
        sit = activities[-1]["description"]
        if len(sit) > 300:
            sit = sit[:300].rsplit(None, 1)[0] + "\u2026"
        text_sections["current_operation"] = sit
        text_sections["day_summary"] = sit

    # =====================================================================
    # DAILY COST (A62/C62)
    # =====================================================================
    daily_cost = _cell(ws, 62, 3, L)                               # C62
    if daily_cost is not None:
        try:
            v = _float(daily_cost)
            if v > 0:
                header["daily_cost"] = int(v)
        except:
            pass

    # =====================================================================
    # SAFETY
    # =====================================================================
    safety = {}
    if header.get("accident_free_days") is not None:
        safety["accident_free_days"] = header["accident_free_days"]

    wb.close()

    return {
        "header": header,
        "activities": activities,
        "text_sections": text_sections,
        "mud_checks": mud_checks,
        "mud_volume": mud_volume,
        "mud_chemical_usage": [],
        "personnel_data": personnel,
        "pumps": [],
        "well_location": {},
        "survey_data": [],
        "safety": safety,
        "tarif_totals": tarif_totals,
    }


# Drop-in compat
parse_daily_excel_report = parse_enf06


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: enf06_extract.py SOURCE.xlsx")
    data = parse_enf06(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
