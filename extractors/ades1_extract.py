#!/usr/bin/env python3
"""
ades1_extract.py - extractor for the ADES-1 layout family.

English-language sibling of the GW-29 / TP-127 template: same overall
section layout (AVANCEMENTS / OUTILS / USURES / PARAMETRES, TIME /
Operations table, MUD PRODUCTS block, TARIFICATION), but titled
"DAILY WORK OVER REPORT" instead of "RAPPORT JOURNALIER DE WORK-OVER",
and with the header block shifted down two rows (well/rig/field/casing
info sits on rows 4-5 instead of rows 1-2).

Operations window: 00:00 to 00:00 (24 h); after-midnight rows excluded.
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Union
from openpyxl import load_workbook


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
    if s in ("", "-", "/", "None"):
        return default
    s = re.sub(r"[a-zA-Z\xc3\xa9\xb0%/\xb3m]+$", "", s).strip()
    try:
        return float(s)
    except ValueError:
        return default


def _int(v, default=0) -> int:
    return int(_float(v, float(default)))


def _date_parse(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date_type):
        return v
    s = _clean(v)
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d,%m,%y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _time_parse(v):
    if v is None:
        return None
    if isinstance(v, time):
        return v
    if isinstance(v, datetime):
        return v.time()
    if isinstance(v, timedelta):
        total = int(v.total_seconds())
        return time((total // 3600) % 24, (total % 3600) // 60)
    s = _clean(v)
    if not s or s in ("-", "None"):
        return None
    if s in ("24:00", "24:00:00"):
        return time(0, 0)
    for fmt in ("%H:%M:%S", "%H:%M", "%Hh%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            continue
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


def _strip_prefix(text, *prefixes):
    s = _clean(text)
    for p in prefixes:
        rx = r"^\s*" + re.escape(p) + r"\s*:?\s*"
        new = re.sub(rx, "", s, flags=re.IGNORECASE)
        if new != s:
            return new.strip()
    return s


def _format_csg(label, depth_val):
    label_clean = _clean(label)
    size = re.sub(r"^\s*(?:Last\s+Csg|Top\s+Liner|Last\s+Lnr)\s*",
                  "", label_clean, flags=re.IGNORECASE).strip()
    if depth_val is None:
        return size
    try:
        depth = int(float(str(depth_val).replace("m", "").replace(",", ".")))
        return "{} @ {}m".format(size, depth)
    except (ValueError, TypeError):
        return "{} @ {}".format(size, _clean(depth_val))


def parse_ades1(source):
    wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)

    header = {}
    header["well_name"]  = _clean(_cell(ws, 4, 2, L))
    header["rig_name"]   = _clean(_cell(ws, 4, 9, L))
    header["field_name"] = _clean(_cell(ws, 4, 5, L))
    header["well_class"] = ""
    header["date"]       = _date_parse(_cell(ws, 1, 22, L))
    day_raw = _clean(_cell(ws, 1, 25, L) or "")
    m = re.search(r"(\d+)", day_raw)
    header["day_number"] = int(m.group(1)) if m else None
    header["bop_test"]      = _date_parse(_cell(ws, 41, 20, L))
    header["next_bop_test"] = None
    header["last_accident"] = None
    header["mud_type"]    = _clean(_cell(ws, 5, 24, L))
    header["supervisor"]  = _clean(_cell(ws, 44, 23, L) or "")
    header["water_truck"] = 0
    header["well_md"] = 0.0

    # LAST CASING/LINER -> label M4 (description), depth O4
    csg_desc  = _cell(ws, 4, 13, L)
    csg_depth = _cell(ws, 4, 15, L)
    if csg_desc or csg_depth:
        header["top_shoe"] = _format_csg(str(csg_desc or ""), csg_depth)

    # Top liner -> label Q4, "depth" (often a text like "TO SURFACE") T4
    liner_label = _cell(ws, 4, 17, L)
    liner_depth = _cell(ws, 4, 20, L)
    if liner_label or liner_depth:
        size = re.sub(r"^\s*Top\s+Liner\s*", "",
                      _clean(liner_label or ""), flags=re.IGNORECASE).strip()
        depth_s = _clean(liner_depth or "").replace("m", "").strip() if liner_depth else ""
        try:
            d = int(float(depth_s.replace(",", ".") or "0"))
            header["last_csg_top"] = "{} @ {}m".format(size, d) if size else "@ {}m".format(d)
        except (ValueError, TypeError):
            header["last_csg_top"] = "{} @ {}".format(size, _clean(liner_depth)).strip()

    bha_raw = _clean(_cell(ws, 12, 2, L) or "")
    if bha_raw:
        header["bha_details"] = bha_raw
    header["well_objective"] = ""

    mud_checks = {}
    if header["mud_type"]:
        mud_checks["mud_type"] = header["mud_type"]

    def _mc(r, c):
        v = _cell(ws, r, c, L)
        return _float(v) if v is not None else None

    # left column (W=label, X=value) and right column (Y=label, Z=value),
    # rows 13-19
    for key, r, c in [("density", 13, 24), ("fun_vis", 14, 24),
                      ("filtrate", 15, 24), ("water_pct", 16, 24),
                      ("solid", 17, 24), ("oil_pct", 18, 24)]:
        val = _mc(r, c)
        if val is not None:
            mud_checks[key] = val
    for key, r, c in [("pom", 13, 26), ("yp", 14, 26), ("gel0", 15, 26),
                      ("gel10sec", 16, 26), ("es", 17, 26), ("pv", 18, 26),
                      ("lsyp", 19, 26)]:
        val = _mc(r, c)
        if val is not None:
            mud_checks[key] = val
    # H/E (Huile/Eau, i.e. oil/water) ratio, e.g. "95/05"
    oil_water_raw = _clean(_cell(ws, 19, 24, L) or "")
    if oil_water_raw:
        parts = re.split(r"[/\\]", oil_water_raw)
        if len(parts) == 2:
            try:
                mud_checks["oil"] = float(parts[0].strip())
                mud_checks["h2o"] = float(parts[1].strip())
            except (ValueError, TypeError):
                pass

    mud_volume = {}
    for row, key in [(7, "reception"), (8, "centrifuge"), (9, "left_in_hole"),
                     (10, "surface_cleaning_settling"), (11, "well_volume"),
                     (12, "surface_volume")]:
        v = _cell(ws, row, 26, L)
        if v is not None:
            mud_volume[key] = _float(v)

    chemicals = []
    for row in range(22, 33):
        name = _clean(_cell(ws, row, 23, L) or "")
        if not name:
            continue
        used_v  = _cell(ws, row, 25, L)
        stock_v = _cell(ws, row, 26, L)
        used  = _clean(used_v) if used_v is not None else ""
        stock = _clean(stock_v) if stock_v is not None else ""
        chemicals.append({"item": name, "units": "", "received": "",
                          "used": used, "on_loc": stock})

    personnel = []
    rep_name = _clean(_cell(ws, 44, 23, L) or "")
    if rep_name:
        personnel.append({"company": "SH/DP Representative", "number": 1,
                          "hours": "", "names": rep_name})

    activities = []
    midnight_seen = False
    for row in range(20, ws.max_row + 1):
        a   = _cell(ws, row, 1, L)
        b   = _cell(ws, row, 2, L)
        c_v = _cell(ws, row, 3, L)
        h_v = _cell(ws, row, 14, L)
        desc   = _clean(c_v)
        a_text = _clean(a).upper() if a else ""
        if any(m in a_text for m in ("AFTER MIDNIGHT", "PLANNED PROGRAM",
                                     "MATERIELS EN LOCATION")):
            break
        if midnight_seen:
            continue
        start_t = _time_parse(a)
        end_t   = _time_parse(b)
        if start_t is None and end_t is None:
            if desc and activities:
                activities[-1]["description"] = (
                    activities[-1]["description"] + "\n" + desc).strip()
            continue
        tarif = _clean(_cell(ws, row, 13, L)).upper()
        if not re.match(r"^T[1-4]$|^NR$|^FT$", tarif):
            tarif = ""
        hours = 0.0
        if h_v is not None:
            ht = _time_parse(h_v)
            if ht is not None:
                hours = ht.hour + ht.minute / 60.0
            elif isinstance(h_v, timedelta):
                hours = h_v.total_seconds() / 3600.0
            elif isinstance(h_v, (int, float)):
                hours = float(h_v)
        if hours == 0.0 and start_t and end_t:
            sm = start_t.hour * 60 + start_t.minute
            em = end_t.hour  * 60 + end_t.minute
            if em == sm:   hours = 24.0 if sm == 0 else 0.0
            elif em > sm:  hours = (em - sm) / 60.0
            else:          hours = (em + 1440 - sm) / 60.0
        activities.append({"start_time": start_t, "end_time": end_t,
            "hours": hours, "phase_name": "", "code": "", "sub": "",
            "description": desc, "start_md": 0, "end_md": 0,
            "npt": 0, "npt_detail": "", "npt_company": "", "op_company": "",
            "bill": tarif})
        if end_t == time(0, 0) and hours > 0:
            midnight_seen = True

    # TARIFICATION cumulative totals: T1/T2/T3/T4 on the "Cumul" row (21)
    tarif_totals = {}
    for code, col in [("T1", 17), ("T2", 19), ("T3", 20), ("T4", 22)]:
        v = _cell(ws, 21, col, L)
        if v is not None:
            f = _float(v)
            if f > 0:
                tarif_totals[code] = f

    text_sections = {}
    status = _clean(_cell(ws, 41, 1, L) or "")
    if status:
        text_sections["current_status"] = status
    remarks = _clean(_cell(ws, 41, 10, L) or "")
    if remarks:
        text_sections["remarks"] = remarks
    plan = _clean(_cell(ws, 44, 1, L) or "")
    if plan:
        text_sections["plan_operations"] = plan
    midnight_parts = []
    for row in range(35, 40):
        v = _clean(_cell(ws, row, 1, L) or "")
        if v:
            midnight_parts.append(v)
    if midnight_parts:
        text_sections["after_midnight_operations"] = " ".join(midnight_parts)

    safety = {}
    if header.get("bop_test"):
        safety["bop_test"] = header["bop_test"]

    wb.close()
    return {"header": header, "activities": activities,
            "text_sections": text_sections, "mud_checks": mud_checks,
            "mud_volume": mud_volume, "mud_chemical_usage": chemicals,
            "personnel_data": personnel, "pumps": [], "well_location": {},
            "survey_data": [], "safety": safety, "tarif_totals": tarif_totals}


parse_daily_excel_report = parse_ades1


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: ades1_extract.py SOURCE.xlsx")
    data = parse_ades1(Path(sys.argv[1]))
    def _default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time):     return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return str(o)
        return str(o)
    print(json.dumps(data, indent=2, default=_default, ensure_ascii=False))
