#!/usr/bin/env python3
"""
enf08_extract.py — extract an ENF#18 (ENAFOR rig 18, HMD field, MD wells)
Daily Workover Report (.xlsx) into the standard dict shape.

Source layout
-------------
Single-sheet French/English template (~60 rows × 35 cols).

  Row 3   : Title "Rapport journalier work-over" (H3)
  Row 4   : Date (V4), Day number (Y4), Division Production (C4)
  Row 6   : Puits (B6), Champ (E6), Appareil (I6=ENF#18),
            Dernier Tubage (K6), Top liner (Q6), bouue (W6)
  Row 7   : FOND (Q7), mud Type (X7, e.g. OBM)
  Row 8   : Section headers: AVANCEMENTS | OUTILS | USURES | PARAMETRES
  Row 9-10: Sub-headers for the above blocks
  Row 11  : Advancement / tool / wear / parameter data
  Row 15  : TOTAL JOUR (A15) + MATERIELS DE FOND (G15) +
            MESURES DE DEVIATION (O15) + mud panel header col labels (W15:Y15)
  Rows 16-19: Mud checks (W=label X=value Y=label Z=value), survey, BHA
            + deviation measurements
  Row 20-21: Mud continuation + tarif T1/T2/T3/T4 labels
  Row 22  : Operations header: Horaire + Analyse des temps et des opérations,
            Code/Heure/Jour, total-jours (Q22), PRODUITS (W22)
  Row 23  : DE/A sub-headers + COUTS/Utilisé/Stock
  Rows 24+: Operations: A=start B=end C=description M=code N=hours
            Chemicals in cols W-Y starting same rows
  Row 21  : Tarif column headers + H/E ratio label
  Row 29  : ∑ total row (N29 = sum of hours), more chemicals
  Row 30-31: Après minuit label + description
  Row 35  : Note block
  Row 38  : Situation à 06:00 (A38, D38)
  Row 39  : Programme prévue (A39, D39)
  Rows 38-39: Personnel/supervisor (W38 label, W39 name)

Distinguishing markers (helpers.parse_source._detect_format_xlsx):
    - "ENF#18" or "ENF # 18" rig name
    - well "MD " prefix + HMD field + "Rapport journalier work-over"
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
    """'ENF#18' or 'ENF # 18' -> 'ENF#18'."""
    s = _clean(name).upper()
    m = re.search(r"ENF\s*#?\s*(\d+)", s)
    if m: return f"ENF#{int(m.group(1)):02d}"
    return _clean(name)


def _first_numeric(ws, L, row, col_start, col_end):
    """Scan cols left-to-right and return the first parsable numeric value."""
    for c in range(col_start, col_end + 1):
        v = _cell(ws, row, c, L)
        if v is None: continue
        try:
            return _float(v)
        except Exception:
            continue
    return None


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_enf08(source: Union[Path, str, BytesIO]) -> dict:
    if isinstance(source, (str, Path)):
        wb = load_workbook(source, data_only=True)
    else:
        wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)

    # =====================================================================
    # HEADER
    # =====================================================================
    header = {}

    # Row 4: Date at V4, day number at Y4
    header["date"]       = _date_parse(_cell(ws, 4, 22, L))   # V4
    header["day_number"] = _int(_cell(ws, 4, 25, L))           # Y4

    # Row 6: well / field / rig / casing / liner / mud
    header["well_name"]    = _clean(_cell(ws, 6, 2,  L))       # B6  Puits
    header["field_name"]   = _clean(_cell(ws, 6, 5,  L))       # E6  Champ (HMD)
    header["rig_name"]     = _normalize_rig(_cell(ws, 6, 9,  L) or "")  # I6  Appareil

    # Dernier Tubage: K6="Dernier Tubage" label, M6:N6="Sabot 7"" (casing description)
    csg_desc = _clean(_cell(ws, 6, 13, L) or "")               # K6 label
    csg_shoe = _clean(_cell(ws, 6, 14, L) or "")               # L6:M6 "Sabot 7""
    if csg_shoe:
        header["top_shoe"] = csg_shoe
    elif csg_desc and "TUBAGE" not in csg_desc.upper():
        header["top_shoe"] = csg_desc

    # Top liner
    liner = _clean(_cell(ws, 6, 17, L) or "")                   # Q6 (merged Q6:S6)
    if liner:
        size = re.sub(r"^\s*Top\s+liner\s*", "", liner, flags=re.IGNORECASE).strip()
        if size:
            header["last_csg_top"] = size

    # Total Depth (FOND) - label at Q7, may or may not have an adjacent value cell
    for c in range(18, 23):
        v = _cell(ws, 7, c, L)
        if v is not None:
            try:
                header["well_md"] = _float(v)
                break
            except (ValueError, TypeError):
                pass

    # Mud type
    mud_type = _clean(_cell(ws, 7, 24, L) or "")                # X7 (merged X7:Z7)
    if mud_type:
        header["mud_type"] = mud_type

    # =====================================================================
    # AVANCEMENT data (row 11)
    # =====================================================================
    # A11=type, B11=op number, C11=depth, D11=advance, E11=duration, F11=cumulative

    # =====================================================================
    # MUD CHECKS  (rows 15-21, cols W-Z interleaved)
    #   Row 15: W=Densit� label, X=1.45 (density value), Y=VB
    #   Row 16: W=V.Marsh label, X=49 (FV), Y=Y.p label, Z=25 (YP)
    #   Row 17: W=Filtrat label, X=4, Y=Gel 0  (Z17 may have gel0 value)
    #   Row 18: W=Sabl% label, X=0, Y=Gel 10   (Z18 may have gel10 value)
    #   Row 19: W=Solide% label, X=0.1, Y=PB, AC=value (decimal day fraction)
    #   Row 20: W=Huile label, X=0.9, Y=PV
    #   Row 21: W=H/E label, X=90/10 ratio, Y=LGS
    # =====================================================================
    mud_checks = {}
    if mud_type:
        mud_checks["mud_type"] = mud_type

    # Row 15: Densit� at W15, value at X15
    v = _cell(ws, 15, 24, L)                                   # X15
    if v is not None: mud_checks["density"] = _float(v)

    # Row 16: V.Marsh at W16, value at X16; YP at Z16
    v = _cell(ws, 16, 24, L)                                   # X16
    if v is not None: mud_checks["fun_vis"] = _float(v)
    v = _cell(ws, 16, 26, L)                                   # Z16
    if v is not None: mud_checks["yp"] = _float(v)

    # Row 17: Filtrat at W17, value at X17; Gel 0 label at Y17, value at Z17
    v = _cell(ws, 17, 24, L)                                   # X17
    if v is not None: mud_checks["apl_fl"] = _float(v)
    v = _cell(ws, 17, 26, L)                                   # Z17
    if v is not None:
        try: mud_checks["gel10sec"] = _float(v)
        except: pass

    # Row 18: Sabl% at W18, value at X18; Gel 10 label at Y18, value at Z18
    v = _cell(ws, 18, 24, L)                                   # X18
    if v is not None: mud_checks["sand"] = _float(v)
    v = _cell(ws, 18, 26, L)                                   # Z18
    if v is not None:
        try: mud_checks["gel10m"] = _float(v)
        except: pass

    # Row 19: Solide% at W19, value at X19; PB at Y19, value at AC19
    v = _cell(ws, 19, 24, L)                                   # X19
    if v is not None: mud_checks["solid"] = _float(v)
    pb_val = _cell(ws, 19, 29, L)                              # AC19 (decimal day fraction)
    if pb_val is not None:
        try: mud_checks["pf"] = _float(pb_val)
        except: pass

    # Row 20: Huile at W20, value at X20; PV at Y20, value at Z20
    v = _cell(ws, 20, 24, L)                                   # X20
    if v is not None: mud_checks["oil"] = _float(v)
    pv_val = _cell(ws, 20, 26, L)                              # Z20
    if pv_val is not None:
        try: mud_checks["pv"] = _float(pv_val)
        except: pass

    # Row 21: H/E at W21, value at X21; LGS at Y21
    he_val = _cell(ws, 21, 24, L)                              # X21
    if he_val is not None:
        s = _clean(he_val)
        if s and "/" in s:
            mud_checks["oil_water_ratio"] = s
    lgs_val = _cell(ws, 21, 25, L)                             # Y21
    if lgs_val is not None:
        try: mud_checks["lgs"] = _float(lgs_val)
        except: pass

    # =====================================================================
    # MUD VOLUMES (rows 12-14, col W labels, scan nearby cols for values)
    #   W11=W14: labels in W11-Y14:
    #     W11 "Perte formation", W12 "Tripping", W13 "Volume puits (m3)",
    #     W14 "Volume surface (m3)"
    #   Values are around col 23-26 area. Let's read from known positions.
    #   Based on the layout these are near the mud panel (W/Y area).
    #   Row 13 col Z or AA = Volume puits value
    #   Row 14 col Z or AA = Volume surface value
    #   Row 11 col Z or AA = Perte formation value
    # =====================================================================
    mud_volume = {}
    # Volume puits
    for c in range(23, 30):
        v = _cell(ws, 13, c, L)
        if v is not None:
            try:
                mud_volume["string_volume"] = _float(v)
                break
            except: pass
    # Volume surface
    for c in range(23, 30):
        v = _cell(ws, 14, c, L)
        if v is not None:
            try:
                mud_volume["pits_volume"] = _float(v)
                break
            except: pass
    # Perte formation
    for c in range(23, 30):
        v = _cell(ws, 11, c, L)
        if v is not None:
            try:
                mud_volume["formation_loss"] = _float(v)
                break
            except: pass
    # Perte surface
    for c in range(23, 30):
        v = _cell(ws, 12, c, L)
        if v is not None:
            try:
                if "Tripping" in _clean(_cell(ws, 12, 23, L) or ""):
                    continue
                mud_volume["surface_loss"] = _float(v)
                break
            except: pass
    if mud_volume:
        parts = [v for v in (mud_volume.get("string_volume"),
                              mud_volume.get("pits_volume"))
                 if isinstance(v, (int, float)) and v]
        if parts:
            mud_volume["total_volume"] = sum(parts)

    # =====================================================================
    # TARIF TOTALS (row 21: T1/T2/T3/T4 labels; row 22 col Q = total hours;
    # row 29: ∑ + total hours)
    # Per-op bill codes in col M, hours in col N.
    # =====================================================================
    tarif_totals = {}

    # Read T1/T2/T3/T4 total hours from row 29 col N (∑ row)
    total_hrs = _cell(ws, 29, 14, L)                            # N29
    if total_hrs is not None:
        try:
            th = _float(total_hrs)
        except:
            th = 0.0
    else:
        th = 0.0

    # Also check tarif labels at row 21 cols Q/S/T/U for per-code totals
    for code, col in [("t1", 17), ("t2", 19), ("t3", 20), ("t4", 21)]:
        v = _cell(ws, 21, col, L)                               # Q21="T1", etc.
        # These are just labels here; the totals come from summing per-op hours
        pass

    # =====================================================================
    # OPERATIONS  (rows 24+, header at row 22: A=Horaire B=A C=desc M=code N=hours)
    # Stop at blank rows or end-of-ops markers (∑ row 29 marks the end).
    # =====================================================================
    activities = []
    for row in range(24, 30):                                   # rows 24-29
        start = _cell(ws, row,  1, L)                            # A
        end   = _cell(ws, row,  2, L)                            # B
        desc  = _clean(_cell(ws, row,  3, L) or "")              # C
        code  = _clean(_cell(ws, row, 13, L) or "")              # M
        hrs   = _cell(ws, row, 14, L)                            # N

        # Stop at summary/total row
        if "∑" in code.upper() or "SOMMA" in code.upper() or "TOTAL" in code.upper():
            break
        # Stop at section headers below ops
        if "APRÈS MINUIT" in desc.upper() or "APRES MINUIT" in desc.upper():
            break

        start_t = _time_parse(start)
        end_t   = _time_parse(end)

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
                "end_time":   end_t,
                "hours":      hours,
                "phase_name": "",
                "code": "", "sub": "",
                "description": desc,
                "start_md": 0, "end_md": 0,
                "npt": 0, "npt_detail": "",
                "npt_company": "", "op_company": "",
                "bill": code.upper(),
            })
        elif desc and activities:
            activities[-1]["description"] = (
                activities[-1]["description"] + "\n" + desc
            ).strip()

    # Derive tarif totals from per-op bill codes
    if not tarif_totals:
        for a in activities:
            code = _clean(a.get("bill")).lower()
            if re.match(r"^t\d+$", code):
                tarif_totals[code] = tarif_totals.get(code, 0.0) + a.get("hours", 0.0)

    # =====================================================================
    # AFTER MIDNIGHT (rows 31-32, col C)
    # =====================================================================
    text_sections = {}
    am_parts = []
    for r in range(31, 35):
        c_val = _clean(_cell(ws, r, 3, L) or "")
        if not c_val:
            continue
        if c_val.upper().startswith("APR"):
            continue  # skip the label row itself
        if any(t in c_val.upper() for t in ("SITUATION", "PROGRAMME", "NOTE:")):
            break
        am_parts.append(c_val)
    if am_parts:
        text_sections["after_midnight"] = " | ".join(am_parts)

    # =====================================================================
    # SITUATION + PROGRAMME (rows 38-39, col A label, col D value)
    # =====================================================================
    sit_label = _clean(_cell(ws, 38, 1, L) or "")               # A38
    if "SITUATION" in sit_label.upper():
        sit_val = _clean(_cell(ws, 38, 4, L) or "")             # D38
        if sit_val:
            if len(sit_val) > 300:
                sit_val = sit_val[:300].rsplit(None, 1)[0] + "\u2026"
            text_sections["current_operation"] = sit_val
            text_sections["day_summary"]       = sit_val

    plan_label = _clean(_cell(ws, 39, 1, L) or "")              # A39
    if "PROGRAMME" in plan_label.upper() or "PROGRAM" in plan_label.upper():
        plan_val = _clean(_cell(ws, 39, 4, L) or "")            # D39
        if plan_val:
            text_sections["plan_operations"] = plan_val

    # =====================================================================
    # SUPERVISOR / PERSONNEL (rows 38-39, cols W)
    #   W38: "Représentant SH/DP" label, W39: supervisor name
    # =====================================================================
    personnel = []
    sup_label = _clean(_cell(ws, 38, 23, L) or "")              # W38
    sup_name  = _clean(_cell(ws, 39, 23, L) or "")              # W39
    if "REPRÉSENTANT" in sup_label.upper() or "REPRESENTANT" in sup_label.upper() or "SH/DP" in sup_label.upper():
        if sup_name:
            header["supervisor"] = sup_name
    elif sup_name:
        # No explicit label but name present
        header["supervisor"] = sup_name

    # =====================================================================
    # MUD CHEMICAL USAGE
    #   Row 22 col W: "PRODUITS" header
    #   Rows 24+: W=item, Y=used (merged with AA..AC), Z=stock
    #   Row 23 col Y: "Utilisé", col Z: "Stock"
    # =====================================================================
    chemicals = []
    for row in range(24, ws.max_row + 1):
        item = _clean(_cell(ws, row, 23, L) or "")              # W
        if not item or item.upper() in ("PRODUITS",):
            continue
        # Stop at non-chemical rows
        o = _clean(_cell(ws, row, 15, L) or "")
        if o.upper() in ("COUTS", "JOURNALIER", "CUMUL",
                          "MATERIELS EN LOCATION",
                          "DERNIÈRE DATE TEST BOP",
                          "DERNIERE DATE TEST BOP"):
            break
        used = _cell(ws, row, 25, L)                            # Y (merged Y:AA)
        stock = _cell(ws, row, 26, L)                           # Z
        # Filter out rows where the "item" cell is actually a date or non-chemical
        if re.match(r"^\d{4}[-/]", item) or len(item) > 60:
            continue
        chemicals.append({
            "item":     item,
            "units":    "",
            "received": "",
            "used":     _clean(str(used)  if used  is not None else ""),
            "on_loc":   _clean(str(stock) if stock is not None else ""),
        })

    # =====================================================================
    # NOTE / REMARKS (row 35, col C)
    # =====================================================================
    note = _clean(_cell(ws, 35, 3, L) or "")
    if note:
        text_sections["remarks"] = note

    # =====================================================================
    # BOP TEST (row 37, label at O37, date at T37)
    # =====================================================================
    bop_label = _clean(_cell(ws, 37, 15, L) or "")              # O37
    if "BOP" in bop_label.upper():
        bop_val = _cell(ws, 37, 20, L)                           # T37
        if bop_val:
            d = _date_parse(bop_val)
            if d:
                header["bop_test"] = d

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
        "survey_data": [],
        "safety": {},
        "tarif_totals": tarif_totals,
    }


# Drop-in compat
parse_daily_excel_report = parse_enf08


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: enf08_extract.py SOURCE.xlsx")
    data = parse_enf08(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
