#!/usr/bin/env python3
"""
enf33_extract.py — extract an ENF#33 (ENAFOR rig 33, BERKINE field, BKNS
wells) Daily Drilling Report (.xlsx / .xlsm) into the standard dict shape.

This is a DIFFERENT layout from the ENF#17 DDR (enf17_extract.py).  It's a
wide single-sheet English template (~77 rows × 74 cols) with:
    - Title "DAILY DRILLING REPORT" (J2-J4), "RIG.S.I" marker (F2)
    - "ENF # 33" in column A
    - Header block rows 2-17 (field/well/depth/HSE/casing/personnel)
    - Two named TPs: "TP. Sénior" (G3) and "TP. Junior" (G4) → the
      two-supervisor convention maps them to supervisor / superintendent
    - Date at W4 (datetime)
    - Operations table rows 20-35: A=Start B=End C=H D=Tarif E=Description
      (per-op tarif codes present; continuation rows have desc but no time)
    - After-midnight at E36+
    - Daily cost C56, Plan operations G56
    - Mud panel in cols R-U around rows 24-30

Distinguishing markers (helpers.parse_source._detect_format_xlsx):
    - "RIG.S.I" + "DAILY DRILLING REPORT"
    - "ENF # 33" / "ENF#33"
    - well "BKNS" prefix
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
    if s in ("", "-", "/", "None"): return default
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
    m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m:
        try: return date_type(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError: return None
    return None


def _time_parse(v) -> Optional[time]:
    if v is None: return None
    if isinstance(v, time): return v
    if isinstance(v, datetime): return v.time()
    if isinstance(v, timedelta):
        total = int(v.total_seconds())
        if total >= 86400: return time(0, 0)
        return time((total // 3600) % 24, (total % 3600) // 60)
    s = _clean(v)
    if not s: return None
    if s in ("24:00", "24:00:00", "24h00"): return time(0, 0)
    m = re.match(r"^(\d{1,2})\s*[:hH]\s*(\d{2})(?:\s*[:hH]\s*\d{2})?$", s)
    if m: return time(int(m.group(1)) % 24, int(m.group(2)))
    return None


def _hours_between(start_t: time, end_t: time) -> float:
    sm = start_t.hour * 60 + start_t.minute
    em = end_t.hour * 60 + end_t.minute
    if em == sm:    return 24.0
    if em > sm:     return (em - sm) / 60.0
    return (em + 1440 - sm) / 60.0


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
    """'ENF # 33' → 'ENF#33'."""
    s = _clean(name).upper()
    m = re.search(r"ENF\s*#?\s*(\d+)", s)
    if m: return f"ENF#{int(m.group(1)):02d}"
    return _clean(name)


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_enf33(source: Union[Path, str, BytesIO]) -> dict:
    # data_only reads cached values; xlsm macros are ignored (we only read data)
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
    # Rig — "ENF # 33" appears in column A; also derive from sheet name
    rig_raw = _clean(_cell(ws, 3, 1, L) or _cell(ws, 2, 1, L) or "")
    if rig_raw:
        header["rig_name"] = _normalize_rig(rig_raw)

    # Field / Well — labels in col F, values in col G
    for r in range(5, 8):
        lab = _clean(_cell(ws, r, 6, L) or "")
        val = _clean(_cell(ws, r, 7, L) or "")
        if lab.upper() == "FIELD" and val:
            header["field_name"] = val
        elif lab.upper() == "WELL" and val:
            header["well_name"] = val

    # Present depth (J6 like "4779 m") and day number (S6)
    pd = _cell(ws, 6, 10, L)
    if pd is not None:
        v = _float(pd)
        if v: header["well_md"] = v
    day = _cell(ws, 6, 19, L)
    if day is not None:
        d = _int(day, 0)
        if d > 0: header["day_number"] = d

    # Date — W4
    dt = _cell(ws, 4, 23, L)
    d = _date_parse(dt)
    if d: header["date"] = d

    # HSE accident-free days — F7
    acc = _cell(ws, 7, 6, L)
    if acc is not None:
        header["accident_free_days"] = _int(acc, 0)

    # Casing / shoe / liner — rows 11-13, label col C, size col D, depth col E
    for r, key in [(11, "last_csg_shoe"), (12, "top_shoe"), (13, "last_csg_top")]:
        lab = _clean(_cell(ws, r, 3, L) or "")
        size = _clean(_cell(ws, r, 4, L) or "")
        depth = _clean(_cell(ws, r, 5, L) or "")
        if size or depth:
            combined = f"{size} @ {depth}" if (size and depth) else (size or depth)
            header[key] = combined

    # Supervisors — TP Sénior (G3) + TP Junior (G4).
    # Two-supervisor convention: senior → supervisor, junior → superintendent.
    tp_senior = _clean(_cell(ws, 3, 7, L) or "")
    tp_junior = _clean(_cell(ws, 4, 7, L) or "")
    if tp_senior and tp_senior.upper() not in ("TP. SÉNIOR", "TP. SENIOR"):
        header["supervisor"] = tp_senior
    if tp_junior and tp_junior.upper() not in ("TP. JUNIOR",):
        header["superintendent"] = tp_junior

    # Daily cost — C56
    dc = _cell(ws, 56, 3, L)
    if dc is not None:
        v = _float(dc)
        if v: header["daily_cost"] = int(v)

    # Water truck — RIG WATER at R35/T35
    for r in range(34, 37):
        if "RIG WATER" in _clean(_cell(ws, r, 18, L) or "").upper():
            wt = _cell(ws, r, 20, L)
            if wt is not None:
                header["water_truck"] = _int(wt, 0)
            break

    # =====================================================================
    # OPERATIONS (rows 20-35): A=Start B=End C=H D=Tarif E=Description
    # =====================================================================
    activities = []
    for r in range(20, 36):
        start = _cell(ws, r, 1, L)
        end   = _cell(ws, r, 2, L)
        hrs   = _cell(ws, r, 3, L)
        tarif = _clean(_cell(ws, r, 4, L) or "")
        desc  = _clean(_cell(ws, r, 5, L) or "")

        # Stop at the after-midnight / section labels
        if "AFTER MIDNIGHT" in desc.upper():
            break

        start_t = _time_parse(start)
        end_t   = _time_parse(end)

        if start_t is not None and end_t is not None:
            hours = _float(hrs) if hrs is not None else 0.0
            if hours == 0.0:
                hours = _hours_between(start_t, end_t)
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
                "bill": tarif.upper(),
            })
        elif desc and activities:
            # Continuation row — fold into previous op's description
            activities[-1]["description"] = (
                activities[-1]["description"] + "\n" + desc
            ).strip()

    # =====================================================================
    # TARIF TOTALS — sum per-op hours by code
    # =====================================================================
    tarif_totals = {}
    for a in activities:
        code = _clean(a.get("bill")).lower()
        if re.match(r"^t\d+$", code):
            tarif_totals[code] = tarif_totals.get(code, 0.0) + a.get("hours", 0.0)

    # =====================================================================
    # MUD CHECKS (cols R label / S value, T label / U value, rows 26-30)
    # =====================================================================
    mud_checks = {}
    mud_label_map = {
        "density": ("density", 19),       # R26=Density, S26=value
        "visc": ("fun_vis", 21),          # T26=Visc, U26=value
        "solide": ("solid", 19),
        "yield point": ("yp", 21),
        "lgs %": ("lgs", 21),
        "oil / water ratio": ("oil_water_ratio", 19),
        "water": ("h2o", 21),
        "sand": ("sand", 21),
    }
    for r in range(24, 36):
        # Left pair: R label / S value
        rlab = _clean(_cell(ws, r, 18, L) or "").lower().rstrip(": ").strip()
        rval = _cell(ws, r, 19, L)
        # Right pair: T label / U value
        tlab = _clean(_cell(ws, r, 20, L) or "").lower().rstrip(": ").strip()
        tval = _cell(ws, r, 21, L)
        for lab, val in ((rlab, rval), (tlab, tval)):
            for key_sub, (dest, _col) in mud_label_map.items():
                if lab.startswith(key_sub):
                    if dest == "oil_water_ratio":
                        if _clean(val):
                            mud_checks[dest] = _clean(val)
                    else:
                        fv = _float(val)
                        if _clean(val) not in ("", "/") and (fv or fv == 0.0):
                            mud_checks[dest] = fv
                    break

    # =====================================================================
    # MUD VOLUMES — Total Circ Vol (U24), Loss Hole Vol (U25)
    # =====================================================================
    mud_volume = {}
    for r in range(24, 27):
        rlab = _clean(_cell(ws, r, 18, L) or "").lower()
        tlab = _clean(_cell(ws, r, 20, L) or "").lower()
        if "total circ" in tlab:
            v = _float(_cell(ws, r, 21, L))
            if v: mud_volume["total_volume"] = v
        if "loss hole" in tlab:
            v = _float(_cell(ws, r, 21, L))
            if v: mud_volume["surface_loss"] = v

    # =====================================================================
    # TEXT SECTIONS
    # =====================================================================
    text_sections = {}
    # After-midnight (E36 label, text in E37+)
    am_parts = []
    capture_am = False
    for r in range(36, 48):
        e = _clean(_cell(ws, r, 5, L) or "")
        if "AFTER MIDNIGHT" in e.upper():
            capture_am = True
            continue
        if capture_am:
            if not e or e.upper().startswith(("NB :", "PLAN")):
                break
            am_parts.append(e)
    if am_parts:
        text_sections["after_midnight"] = " ".join(am_parts)

    # Plan operations — G56 (label "Plan Opetations :" at E56)
    plan = _clean(_cell(ws, 56, 7, L) or "")
    if plan:
        text_sections["plan_operations"] = plan

    # Situation = last op description (capped 300)
    if activities:
        sit = activities[-1]["description"]
        if len(sit) > 300:
            sit = sit[:300].rsplit(None, 1)[0] + "…"
        text_sections["current_operation"] = sit
        text_sections["day_summary"] = sit

    # =====================================================================
    # PERSONNEL (rows 14-17: label/value pairs across the row)
    # r14 labels: Sup int / S Tool Pusher / J Tool Pusher / Driller / A/Driller / ...
    # r15 values; r16 labels; r17 values
    # =====================================================================
    personnel = []
    for label_row, value_row in ((14, 15), (16, 17)):
        seen_labels = set()
        last_label = None
        for c in range(1, 30):
            lab = _clean(_cell(ws, label_row, c, L) or "")
            val = _cell(ws, value_row, c, L)
            # Skip merged-cell repeats: a label identical to the immediately
            # preceding column is the same merged cell spanning columns.
            if lab and lab == last_label:
                continue
            last_label = lab if lab else last_label
            if not lab or lab in seen_labels:
                continue
            if val is None or not _clean(val):
                continue
            seen_labels.add(lab)
            cnt = _clean(val)
            is_count = bool(re.match(r"^[\d+]+$", cnt.replace(" ", "")))
            personnel.append({
                "company": lab,
                "number":  _int(cnt, 0) if is_count else 0,
                "hours":   "",
                "names":   "" if is_count else cnt,
            })

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
parse_daily_excel_report = parse_enf33


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: enf33_extract.py SOURCE.xlsx|SOURCE.xlsm")
    data = parse_enf33(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
