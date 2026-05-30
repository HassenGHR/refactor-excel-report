#!/usr/bin/env python3
"""
enf27_extract.py — extract an ENF#27 (ENAFOR rig 27, Direction Régionale
Ohanet, DIMW wells) Daily Workover Report (.xlsx) into the standard dict
shape.

Source layout
-------------
Compact French workover template (~38 rows × 14 cols).  Title at D2:
"RAPPORT JOURNALIER WORK OVER" with regional direction "DIRECTION
REGIONALE OHANET".

  Row 2   : C2 = regional direction, D2 = title, K2 = "JOURNEE DU",
            M2 = "RAPPORT"
  Row 3   : M3 = day number
  Row 4   : Headers — A4="PUITS", C4="APPAREIL", D4="Dernier Tubage",
            E4="Cote sabot", G4="LINER:", H4="Liner", J4="Fond",
            K4="TYPE BOUE", M4 = mud type (OBM)
  Row 5   : A5 = well name, C5 = rig (ENF 27), D5 = casing description,
            E5 = shoe depth (numeric), J5 = TD/Fond depth
  Row 6   : A5/A7 = well objective (merged across A7:D9),
            K6/P6 = "Fabrication"/"Ejection", E6="TARIFICATION"
  Row 7   : F7="T1", H7="T2", I7="T3", J7="T4", K7="Perte surface",
            M7="Transfert"
  Row 8   : E8="H/Jour", F8=T1 hrs, H8=T2 hrs, I8=T3, J8=T4,
            K8="Perte accident", M8="V. puits", N8=volume puits
  Row 9   : E9="H/Cum" (cumulative hours), K9="Perte formation",
            M9="V. Surf", N9=volume surface
  Row 10  : E10="COUTS (DA)", K10="Densité", L10=density value,
            M10="Yield"
  Row 11  : Single-row operation: A11=start time, B11=end time (e.g. "24h00"),
            C11=description, E11="JOURNALIER", H11="CUMULE",
            K11="Visc. March.", L11=viscosity value, M11="PH"
  Row 12  : E12=daily cost value (numeric), K12="PARAMETRES DE FRAISAGE"
  Rows 13-26: Cost breakdown by vendor (E=vendor name, I=amount DA)
            Row 26: E26="PERSONNEL" header
  Rows 27-35: Personnel — E=role label, J=count, K/N=chemicals header rows
            K28:L28="PRODUITS SH DP OHT" header
            Rows 29-35: role + count in E+J, chemicals in K-N
  Row 36  : E36="TOTAL" sum of personnel, J36=total count,
            K36="Rep maître oeuvre", M36="Resp sce puits"
  Row 37  : A37="SITUATION AU RAPPORT:" + description,
            K37=supervisor name, M37=second supervisor name
            E37="Véhicule", F37="Type", J37="N°  SH"
  Row 38  : A38="PROGRAMME  PRÉVU :" + description,
            rows 37-36 cols E-N: vehicle info

Distinguishing markers (helpers.parse_source._detect_format_xlsx):
    - "ENF 27" or "ENF#27" rig name
    - "DIRECTION REGIONALE OHANET"
    - well "DIMW-" prefix
    - title "RAPPORT JOURNALIER WORK OVER" matches TP-179 but ENF#27
      has OHANET regional direction to distinguish it
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
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: continue
    return None


def _time_parse(v):
    if v is None: return None
    if isinstance(v, time): return v
    if isinstance(v, datetime):
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
    """'ENF 27' / 'ENF#27' -> 'ENF#27'."""
    s = _clean(name).upper()
    m = re.search(r"ENF\s*#?\s*(\d+)", s)
    if m: return f"ENF#{int(m.group(1)):02d}"
    return _clean(name)


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_enf27(source: Union[Path, str, BytesIO]) -> dict:
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
    header["day_number"] = _int(_cell(ws, 3, 13, L))           # M3
    header["rig_name"]   = _normalize_rig(_cell(ws, 5, 3, L) or "")  # C5 "ENF 27"
    header["well_name"]  = _clean(_cell(ws, 5, 1, L))          # A5

    # Mud type at M4 (merged M4:N4)
    header["mud_type"] = _clean(_cell(ws, 4, 13, L))           # M4 (merged K4:N4)

    # Casing description (D5)
    csg_desc = _clean(_cell(ws, 5, 4, L))                       # D5
    csg_shoe = _cell(ws, 5, 5, L)                               # E5 (shoe depth)
    if csg_desc or csg_shoe is not None:
        if csg_shoe is not None:
            try:
                d = int(float(str(csg_shoe)))
                header["top_shoe"] = f"{csg_desc} @ {d}m" if csg_desc else f"@ {d}m"
            except (ValueError, TypeError):
                header["top_shoe"] = f"{csg_desc} @ {_clean(csg_shoe)}".strip()
        else:
            header["top_shoe"] = csg_desc

    # Liner/Top liner — labels at G4/H4, value at H5 (merged H5:I5)
    # In this template H5 is often empty; only set if a real value exists
    liner = _cell(ws, 5, 8, L)                                 # H5 (merged H5:I5)
    if liner is not None:
        s = _clean(str(liner))
        if s and s.upper() not in ("LINER", "FOND"):
            header["last_csg_top"] = s

    # Total Depth (FOND at J4, value at J5)
    td = _cell(ws, 5, 10, L)                                    # J5
    if td is not None:
        header["well_md"] = _float(td)

    # Well objective (A6/A7 merged)
    obj = _clean(_cell(ws, 7, 1, L) or _cell(ws, 6, 1, L))     # A7 (merged A7:D9)
    if obj:
        header["well_objective"] = obj

    # Fabrication / Ejection (K6/M6)
    fabrication = _cell(ws, 5, 11, L)                           # K5 "PARAMETRES"
    ejection = _cell(ws, 5, 13, L)                              # M5

    # =====================================================================
    # TARIF TOTALS  (row 8)
    # =====================================================================
    tarif_totals = {}
    for code, col in [("t1", 6), ("t2", 8), ("t3", 9), ("t4", 10)]:
        v = _cell(ws, 8, col, L)
        if v is not None:
            try:
                hrs = _float(v)
                if hrs > 0:
                    tarif_totals[code] = hrs
            except:
                pass
    # Cumul row 9
    for code, col in [("cum_t1", 6), ("cum_t2", 8), ("cum_t3", 9), ("cum_t4", 10)]:
        v = _cell(ws, 9, col, L)
        if v is not None:
            try:
                hrs = _float(v)
                if hrs > 0:
                    tarif_totals[code] = hrs
            except:
                pass

    # =====================================================================
    # MUD CHECKS  (rows 8-11)
    # Dense row: K10=L10(density), M10=Yield label
    # Row 11: K11=Visc. March., L11=viscosity value, M11=PH
    # =====================================================================
    mud_checks = {}

    density = _cell(ws, 10, 12, L)                             # L10
    if density is not None: mud_checks["density"] = _float(density)

    # Viscosity
    visc = _cell(ws, 11, 12, L)                                # L11
    if visc is not None: mud_checks["fun_vis"] = _float(visc)

    # Yield point — M10 label, value may be M10 or nearby
    # In this template M10 is "Yield" label only, no separate value cell

    # =====================================================================
    # MUD VOLUMES
    # =====================================================================
    mud_volume = {}
    # V. puits (N8), V. Surf (N9)
    v = _cell(ws, 8, 14, L)                                    # N8
    if v is not None: mud_volume["string_volume"] = _float(v)
    v = _cell(ws, 9, 14, L)                                    # N9
    if v is not None: mud_volume["pits_volume"] = _float(v)
    # Perte accident (K8 label), value likely no data cell in this sample
    # Perte formation (K9 label)
    if mud_volume:
        mud_volume["total_volume"] = sum(
            v for v in (mud_volume.get("string_volume"), mud_volume.get("pits_volume"))
            if isinstance(v, (int, float))
        )

    # =====================================================================
    # OPERATIONS  (row 11+)
    # Row 11: A=start, B=end (possibly "24h00"), C=description
    # Only one op on this single-op template. May extend to row 13 etc.
    # =====================================================================
    activities = []
    for row in range(11, 20):
        start = _cell(ws, row, 1, L)                            # A
        end   = _cell(ws, row, 2, L)                            # B
        desc  = _clean(_cell(ws, row, 3, L) or "")              # C

        # Stop at cost/personnel section headers
        e_val = _clean(_cell(ws, row, 5, L) or "")
        if e_val.upper() in ("COUTS (DA)", "COUT DETAILLE (DA)",
                              "PERSONNEL", "TOTAL"):
            break
        if "SITUATION" in desc.upper():
            break

        start_t = _time_parse(start)
        end_t   = _time_parse(end)

        if start_t is not None and end_t is not None:
            hours = 0.0
            if start_t and end_t:
                sm = start_t.hour * 60 + start_t.minute
                em = end_t.hour * 60 + end_t.minute
                if em == sm: hours = 24.0
                elif em > sm: hours = (em - sm) / 60.0
                else: hours = (em + 1440 - sm) / 60.0

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
                "bill": "",
            })
        elif desc and activities:
            activities[-1]["description"] = (
                activities[-1]["description"] + "\n" + desc
            ).strip()

    # Back-assign bill codes from tarif totals
    from helpers.bill_code_assign import assign_bill_codes
    activities = assign_bill_codes(activities, tarif_totals)

    # =====================================================================
    # DAILY COST  (E12)
    # =====================================================================
    daily_cost = _cell(ws, 12, 5, L)                           # E12
    if daily_cost is not None:
        try:
            v = _float(daily_cost)
            if v > 0:
                header["daily_cost"] = int(v)
        except:
            pass

    # =====================================================================
    # PERSONNEL  (rows 27-35)
    # Row 26: "PERSONNEL" header
    # Rows 27-35: E=role, J=count
    # After row 28 (chemicals header at K28:L28), skip chemical items
    # =====================================================================
    personnel = []
    CHEMICAL_HEADERS = {"PRODUITS SH DP OHT", "SOUDE", "BENTONITE", "SEL",
                        "KCL", "AVAHEC", "EAU", "BRUT"}

    for row in range(27, 36):
        role = _clean(_cell(ws, row, 5, L))                    # E
        count = _cell(ws, row, 10, L)                          # J
        chem = _clean(_cell(ws, row, 11, L))                   # K

        if not role:
            continue
        if role.upper() == "TOTAL":
            continue
        if role.upper() == "PERSONNEL":
            continue
        if role in CHEMICAL_HEADERS:
            continue
        # Skip if K is a known chemical item and col E is empty (pure chemical rows)
        if chem in CHEMICAL_HEADERS and not role:
            continue
        if not count or not isinstance(count, (int, float)):
            continue

        personnel.append({
            "company": role,
            "number": _int(count),
            "hours": "",
            "names": "",
        })

    # =====================================================================
    # SUPERVISORS  (row 36)
    # =====================================================================
    # K36="Rep maître oeuvre", M36="Resp sce puits"
    # Values at K37, M37
    sup1_label = _clean(_cell(ws, 36, 11, L) or "")            # K36
    sup2_label = _clean(_cell(ws, 36, 13, L) or "")            # M36
    sup1_name  = _clean(_cell(ws, 37, 11, L) or "")            # K37
    sup2_name  = _clean(_cell(ws, 37, 13, L) or "")            # M37
    if "MAÎTRE" in sup1_label.upper() or "MAITRE" in sup1_label.upper():
        if sup1_name:
            header["supervisor"] = sup1_name
    if "SCE PUITS" in sup2_label.upper() or "RESP" in sup2_label.upper():
        if sup2_name:
            header["superintendent"] = sup2_name

    # =====================================================================
    # TEXT SECTIONS
    # =====================================================================
    text_sections = {}

    # Situation
    sit = _clean(_cell(ws, 37, 1, L) or "")                    # A37
    if sit:
        sit_text = re.sub(r"^SITUATION\s+AU\s+RAPPORT\s*:?\s*", "", sit, flags=re.IGNORECASE)
        if sit_text:
            if len(sit_text) > 300:
                sit_text = sit_text[:300].rsplit(None, 1)[0] + "\u2026"
            text_sections["current_operation"] = sit_text
            text_sections["day_summary"] = sit_text

    # Programme prévu
    plan = _clean(_cell(ws, 38, 1, L) or "")                   # A38
    if plan:
        plan_text = re.sub(r"^PROGRAMME\s+PR[ÉE]VU[ÉE]?\s*:?\s*", "", plan, flags=re.IGNORECASE)
        if plan_text:
            text_sections["plan_operations"] = plan_text

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
        "safety": {},
        "tarif_totals": tarif_totals,
    }


# Drop-in compat
parse_daily_excel_report = parse_enf27


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: enf27_extract.py SOURCE.xlsx")
    data = parse_enf27(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
