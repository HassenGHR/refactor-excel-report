#!/usr/bin/env python3
"""
rap_entp_extract.py — extract the ENTP "RAP ENTP" wide, single-sheet daily
drilling report (filename pattern RAP_ENTP_N_<report#>_<rig>_MD_<well>_DU_
<date>.xlsx) into the same overall dict shape used by the other extractors
in this set (ddr_extract / tp179_extract / wo_report_extract / dwr_extract),
plus template-specific extras.

This is the densest of the four templates seen so far — everything lives on
one sheet (~65 rows x 24 cols), packed into many small side-by-side blocks
rather than one tall report:

  Rows 1-7    header banner: rig, casing-shoe depths, last-BOP-test dates,
              BOP stack inventory, report number/date, well/client/region,
              org-chart names (superintendant/tool pushers/HSE/SH-DP sup).
  Rows 8-23   bit data (blank on non-bit-change days), drill-string tally
              (drill pipe / drill collar joint counts), surface BOP-stack
              equipment list, mud-pump rate table.
  Rows 24-47  mud characteristics + mud volumes (left block, cols A-B);
              operations log — E=from, F=to, G=duration, H:P=description
              (cols Q=tarif code, R=cost/reference code) (rows 25-27ish);
              a BHA-component tally sharing the same rows but cols S-T;
              inventaire tubulaire (pipe-on-location tally, cols C-D);
              free-text notes (situation après minuit, concerns) cols H-P;
              cumulative tarif/NPT stats, cols R-T rows 14-17.
  Rows 48-53  fuel consumption; mud-additive treated/remaining tallies;
              TARIF "TIMING" totals (T1-T4, NR, total), cols S-T.
  Rows 54-65  five side-by-side blocks: VEHICLES (A-C), drilling-line
              inspection (G-J), safety-meeting / STOP-cards / gas bottles
              (K-M), PERSONNEL headcount by company (N-P), and the
              CURRENT OPERATION free-text block (Q-T).

Because several of these blocks repeat the same value across many merged
rows (e.g. the BOP-test date, or a safety topic repeated for five rows),
list-building code below de-duplicates consecutive repeats.
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Union

from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def _clean(v) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _float(v, default=0.0):
    if v is None or v == "":
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = _clean(v).replace(",", ".").replace(" ", "")
    if s in ("", "-", "/", "None"):
        return default
    s = re.sub(r"[a-zA-Z%]+$", "", s).strip()
    try:
        return float(s)
    except ValueError:
        return default


def _int(v, default=0) -> int:
    return int(_float(v, float(default)))


def _hours(v):
    """Normalize a timedelta / datetime.time / numeric-fraction cell to hours."""
    if v is None or v == "":
        return None
    if isinstance(v, timedelta):
        return round(v.total_seconds() / 3600, 3)
    if isinstance(v, time):
        return round(v.hour + v.minute / 60 + v.second / 3600, 3)
    if isinstance(v, datetime):
        return round(v.hour + v.minute / 60 + v.second / 3600, 3)
    if isinstance(v, (int, float)):
        return round(v * 24, 3) if 0 <= v <= 1 else float(v)
    return None


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


def _dedup(seq):
    out = []
    for x in seq:
        if x and (not out or out[-1] != x):
            out.append(x)
    return out


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_rap_entp(source: Union[Path, str, BytesIO]) -> dict:
    wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)
    C = lambda r, c: _cell(ws, r, c, L)   # shorthand

    # =====================================================================
    # HEADER
    # =====================================================================
    header = {}
    header["rig_name"] = _clean(C(1, 2))                 # B1
    header["report_number"] = _int(C(1, 16), 0)          # P1
    header["date"] = C(1, 19)                             # S1
    if isinstance(header["date"], datetime):
        header["date"] = header["date"].date()
    header["well_name"] = _clean(C(2, 19))                # S2
    header["client"] = _clean(C(3, 19))                   # S3
    header["region"] = _clean(C(4, 19))                   # S4
    header["op_start_note"] = _clean(C(5, 18))            # R5, free text

    # Casing shoe depths — rows 2-7, label col D, depth col F (blank if N/A)
    casing_shoes = []
    for row in range(2, 8):
        label = _clean(C(row, 4))
        if not label:
            continue
        depth = C(row, 6)
        casing_shoes.append({"item": label, "depth_m": depth if depth is not None else ""})
    header["casing_shoes"] = casing_shoes

    header["last_bop_test"] = C(2, 8)                      # H2 (date, repeated rows 2-7)
    if isinstance(header["last_bop_test"], datetime):
        header["last_bop_test"] = header["last_bop_test"].date()

    header["bop_stack"] = _dedup(_clean(C(row, 11)) for row in range(2, 8))   # K2:K7

    # Org-chart names — rows 2-7, role col N, name col P
    org = {}
    for row in range(2, 8):
        role = _clean(C(row, 14))
        name = _clean(C(row, 16))
        if role and name:
            org[role] = name
    header["org_chart"] = org

    # =====================================================================
    # DRILL STRING TALLY (rows 14-19, cols A=item,B=daily joints,C=total joints)
    # =====================================================================
    drill_string = []
    for row in range(14, 20):
        item = _clean(C(row, 1))
        if not item:
            continue
        drill_string.append({
            "item": item,
            "daily": C(row, 2),
            "total": C(row, 3),
        })

    # =====================================================================
    # MUD CHECKS  (rows 25-47, label col A, value col B)
    # =====================================================================
    MUD_KEYS = {
        "WEIGHT": "density", "APP VISC": "app_vis", "PLASTIC VISC": "pv",
        "FUNNEL VISC": "fun_vis", "LGS": "lgs", "PH": "ph",
        "WATER %": "water_pct", "SOLIDES %": "solids_pct", "OIL       %": "oil_pct",
        "GEL.0": "gel0", "GEL.10": "gel10", "YIELD POINT": "yp", "PF": "pf",
        "KCL": "kcl", "HP HT FILTRATE": "hpht_filtrate", "PM": "pm",
    }
    mud_checks = {}
    for row in range(25, 42):
        label = _clean(C(row, 1)).upper()
        key = MUD_KEYS.get(label)
        if key:
            v = C(row, 2)
            if v is not None:
                mud_checks[key] = _float(v)

    mud_volume = {}
    VOL_KEYS = {
        "EJECTION BOUE": "ejected", "PERTE FORMATION": "perte_formation",
        "PERTE TRIPPING": "perte_tripping", "VOLUME PUITS": "well_volume",
        "TOTAL SURF.": "total_surface", "EVAPORATION": "evaporation",
    }
    for row in range(42, 48):
        label = _clean(C(row, 1)).upper()
        key = VOL_KEYS.get(label)
        if key:
            v = C(row, 2)
            if v is not None:
                mud_volume[key] = _float(v)

    # =====================================================================
    # FUEL  (rows 48-53, cols A-D + F-G)
    # =====================================================================
    fuel = {}
    rig_cons = C(50, 3)     # C50 (Rig m3/j)
    cam_cons = C(51, 3)     # C51 (Cam m3/j)
    if rig_cons is not None:
        fuel["rig_consumption_m3_per_day"] = _float(rig_cons)
    if cam_cons is not None:
        fuel["camp_consumption_m3_per_day"] = _float(cam_cons)
    order = C(53, 2)        # B53
    if order is not None:
        fuel["order_m3"] = _float(order)

    # =====================================================================
    # OPERATIONS LOG — cols E(from)/F(to)/G(duration)/H:P(description,
    # merged wide) / Q(tarif code)/R(reference code). Rows 25 onward, until
    # a row has none of these populated.
    # =====================================================================
    activities = []
    for row in range(25, 34):
        frm = C(row, 5)     # E
        to = C(row, 6)      # F
        dur = C(row, 7)     # G
        desc = _clean(C(row, 8))     # H
        code = _clean(C(row, 17))    # Q
        ref = _clean(C(row, 18))     # R

        if frm is None and to is None and not desc:
            continue

        def _clock(v):
            h = _hours(v)
            if h is None:
                return None
            h = h % 24
            hh = int(h)
            mm = int(round((h - hh) * 60))
            if mm == 60:
                mm, hh = 0, (hh + 1) % 24
            return time(hh, mm)

        activities.append({
            "start_time": _clock(frm),
            "end_time": _clock(to),
            "hours": _hours(dur) or 0.0,
            "phase_name": "",
            "code": ref,
            "sub": "",
            "description": desc,
            "start_md": 0, "end_md": 0,
            "npt": 0, "npt_detail": "",
            "npt_company": "", "op_company": "",
            "bill": code,
        })

    # =====================================================================
    # BHA COMPONENT TALLY — cols S(item)/T(length), rows 25-32 (independent
    # of the operations log above, despite sharing the same row range).
    # =====================================================================
    bha_components = []
    for row in range(25, 33):
        item = _clean(C(row, 19))    # S
        if not item or item.upper() == "BHA":
            length = C(row, 20)
            if item.upper() == "BHA" and length is not None:
                bha_components.append({"item": "TOTAL", "length_m": _float(length)})
            continue
        length = C(row, 20)          # T
        bha_components.append({"item": item, "length_m": _float(length) if length is not None else ""})

    # =====================================================================
    # TARIF "TIMING" totals — cols S(code)/T(hours), rows 47-53
    # =====================================================================
    tarif_totals = {}
    for row in range(47, 53):
        code = _clean(C(row, 19)).upper()   # S
        if not code or not re.match(r"^T[1-4]$|^NR$", code):
            continue
        h = _hours(C(row, 20))              # T
        if h:
            tarif_totals[code] = h
    total_h = _hours(C(53, 20))             # T53
    if total_h:
        tarif_totals["_total"] = total_h

    # =====================================================================
    # TEXT SECTIONS
    # =====================================================================
    text_sections = {}
    situation = _clean(C(39, 8))            # H39, under "Situation après minuit :" (H38)
    if situation:
        text_sections["situation_after_midnight"] = situation

    concerns = []
    for row in (40, 43, 45, 49):
        v = _clean(C(row, 8))
        if v:
            concerns.append(v)
    if concerns:
        text_sections["concerns"] = concerns

    current_op = _dedup(_clean(C(row, 17)) for row in range(55, 60))   # Q55:Q59
    if current_op:
        text_sections["current_operation"] = "\n".join(current_op)
        text_sections["day_summary"] = current_op[0]

    # =====================================================================
    # VEHICLES  (rows 55-62, cols A=name,B=type,C=plate)
    # =====================================================================
    vehicles = []
    for row in range(55, 63):
        name = _clean(C(row, 1))
        if not name:
            continue
        vehicles.append({
            "vehicle": name,
            "type": _clean(C(row, 2)),
            "plate": _clean(C(row, 3)),
        })

    # =====================================================================
    # DRILLING LINE  (rows 55-65, cols G=label,H/J=value)
    # =====================================================================
    drilling_line = {}
    DL_KEYS = {
        "Diamètre": "diameter", "Manufacturer": "manufacturer",
        "Slipping": "slipping", "Last Slipping": "last_slipping",
        "Last Cutting": "last_cutting", "Reste Câble": "cable_remaining_m",
        "Cumm Cutting": "cumulative_cutting", "FREE DAYS": "free_days",
        "Total Accident": "total_accident", "Last Accident": "last_accident",
    }
    for row in range(55, 66):
        label = _clean(C(row, 7))    # G
        key = DL_KEYS.get(label)
        if not key:
            continue
        col = 10 if row >= 63 else 8   # rows 63-65 (FREE DAYS/Total Accident/Last Accident) use col J; others col H
        v = ws.cell(row, col).value
        if isinstance(v, datetime):
            v = v.date()
        if v is not None:
            drilling_line[key] = v

    # =====================================================================
    # SAFETY MEETING / STOP CARDS / GAS BOTTLES  (rows 55-65, cols K-M)
    # =====================================================================
    safety = {}
    topics = _dedup(_clean(C(row, 11)) for row in range(55, 60))   # K55:K59
    if topics:
        safety["topics"] = topics
    stop_daily = C(61, 12)     # L61
    stop_total = C(61, 13)     # M61
    if stop_daily is not None or stop_total is not None:
        safety["stop_cards"] = {"daily": stop_daily, "total": stop_total}
    gas_bottles = []
    for row in range(63, 66):
        item = _clean(C(row, 11))     # K
        if not item:
            continue
        gas_bottles.append({
            "item": item,
            "full": C(row, 12),        # L
            "empty": C(row, 13),       # M
        })
    if gas_bottles:
        safety["gas_bottles"] = gas_bottles

    # =====================================================================
    # PERSONNEL  (rows 55-64, cols N=company,P=count)
    # =====================================================================
    personnel = []
    for row in range(55, 65):
        company = _clean(C(row, 14))    # N
        if not company or company.upper() == "TOTAL":
            continue
        count = C(row, 16)               # P
        personnel.append({
            "company": company,
            "number": _int(count, 0),
            "hours": "",
            "names": "",
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
        "safety": safety,
        "tarif_totals": tarif_totals,
        # Extra info beyond the common dict shape:
        "drill_string": drill_string,
        "bha_components": bha_components,
        "fuel": fuel,
        "vehicles": vehicles,
        "drilling_line": drilling_line,
    }


# Drop-in compat for the converter's import
parse_daily_excel_report = parse_rap_entp
parse_ddr = parse_rap_entp


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: rap_entp_extract.py SOURCE.xlsx")
    data = parse_rap_entp(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)):
            return o.isoformat()
        if isinstance(o, time):
            return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta):
            return o.total_seconds()
        return str(o)

    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
