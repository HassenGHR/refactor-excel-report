#!/usr/bin/env python3
"""
enf18_extract.py — extract the ENF#18 workover-rig "Rapport journalier
work-over" daily report (single sheet named e.g. "MD 300 (135)", well
MD-300) into the same overall dict shape used by the other extractors in
this set (ddr_extract / tp179_extract / wo_report_extract / dwr_extract /
rap_entp_extract).

This template is closely related to the legacy DWR_SH_DP...xls layout
(same "Division Production / Rapport journalier work-over" banner, same
bit/BHA/deviation scaffolding, same parallel "code + hours" tally next to
the operations log) but as a native .xlsx workbook with everything shifted
a few rows and with some column positions changed:

  Row 6      header: well (B6), field (E6), rig (I6), last-casing note
             (M6, free text e.g. 'Sabot 7"'), top-liner label/value (Q6/Q7).
  Row 4      report date (V4) and report number (Y4).
  Row 7      mud type (X7, under "Type"/"boue" at W6/W7).
  Rows 9-11  bit-run summary (type/no/depth/advance/duration/ROP/OD/make) —
             usually only one row filled, on days with a bit change.
  Rows 15-21 mud checks, cols W(label)/X(value) and Y(label)/Z(value) —
             real filled-in copies of this template are inconsistent about
             which of X/Y/Z actually holds a given label's value, so this
             block is parsed with a generic "nearest value to the right"
             scan (see _labelled_row_values).
  Rows 22-47 mud/chemical "PRODUITS" tally, col W=item, X='Utilisé' (used),
             Z='Stock' (on location) — often left blank day-to-day.
  Rows 24-4x operations log: A=start time, B=end time, C:L=description
             (merged wide), M=tarif code, N=duration **in plain hours**
             (unlike the legacy .xls sibling template, not a date/fraction).
  Row 21-22  TARIFICATION today's-hours-by-code row (Q/S/T/V = T1/T2/T3/T4)
             — used as a cross-check/fallback for the M/N tally above.
  Row 44     "Après minuit :" label + value (situation after midnight).
  Row 52-54  free-text "Note:" block (well status / rig moves / targets)
             and "Dernière date test BOP" (last BOP test date).
  Row 55-56  "Situation à 06:00" / "Programme prévue" free text, plus a
             vehicle / "Représentant SH/DP" block.

No PERSONNEL breakdown table is present on this template (unlike the DWR
.xls sibling), so `personnel_data` is always empty here.
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


def _strip_stray_ref(text: str) -> str:
    """
    Some copies of this template have leftover cell-reference text pasted
    into a label (e.g. "Baryte+W24:X27") from a broken copy/paste in the
    source workbook. Strip that kind of trailing artifact.
    """
    return re.sub(r"\+[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$", "", text).strip()


def _hours_from_time(t):
    if isinstance(t, time):
        return t.hour + t.minute / 60 + t.second / 3600
    if isinstance(t, datetime):
        return t.hour + t.minute / 60 + t.second / 3600
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


_TARIF_CODE_RE = re.compile(r"^T[1-4]$|^NR$")


def _labelled_row_values(ws, row, col_start, col_end):
    """
    Generic 'label picks up nearest value to its right' scan for the mud
    checks / products side block, where real filled-in copies of this
    template don't consistently put a label's value in the very next cell.
    Returns list of (label, value).
    """
    cells = []
    for c in range(col_start, col_end + 1):
        v = ws.cell(row, c).value
        if v is not None and v != "":
            cells.append((c, v))

    out = []
    i = 0
    while i < len(cells):
        c, v = cells[i]
        if isinstance(v, str):
            label = _clean(v)
            value = None
            j = i + 1
            if j < len(cells) and not isinstance(cells[j][1], str):
                value = cells[j][1]
            out.append((label, value))
        i += 1
    return out


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_enf18_report(source: Union[Path, str, BytesIO]) -> dict:
    wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)
    C = lambda r, c: _cell(ws, r, c, L)

    # =====================================================================
    # HEADER
    # =====================================================================
    header = {}
    header["well_name"] = _clean(C(6, 2))               # B6
    header["field_name"] = _clean(C(6, 5))               # E6
    header["rig_name"] = _clean(C(6, 9))                  # I6
    header["last_casing_note"] = _clean(C(6, 13))         # M6, e.g. 'Sabot 7"'

    liner_label = _clean(C(6, 17))                        # Q6
    liner_value = _clean(C(7, 17))                        # Q7
    if liner_label or liner_value:
        header["top_liner"] = f"{liner_label}: {liner_value}".strip(": ")

    rdate = C(4, 22)                                      # V4
    header["date"] = rdate.date() if isinstance(rdate, datetime) else rdate
    header["report_number"] = _clean(C(4, 25))            # Y4

    mud_type = _clean(C(7, 24))                           # X7
    if mud_type:
        header["mud_type"] = mud_type

    bop_date = C(54, 20)                                  # T54, "Dernière date test BOP"
    if isinstance(bop_date, datetime):
        bop_date = bop_date.date()
    if bop_date:
        header["last_bop_test"] = bop_date

    representative = _clean(C(56, 23))                    # W56, under "Représentant SH/DP"
    if representative:
        header["representative"] = representative

    veh_type = _clean(C(56, 18))                          # R56
    veh_no = _clean(C(56, 20))                            # T56
    if veh_type or veh_no:
        header["vehicle_type"] = veh_type
        header["vehicle_number"] = veh_no

    # =====================================================================
    # BIT RUN (rows 9-14): A=type, B=no, C=depth, D=advance, E=duration,
    # F=ROP, G=OD, H=make/type. Usually only one row filled.
    # =====================================================================
    bit_runs = []
    for row in range(11, 15):
        op_type = _clean(C(row, 1))
        if not op_type:
            continue
        bit_runs.append({
            "type": op_type,
            "number": _clean(C(row, 2)),
            "depth_m": C(row, 3),
            "advance_m": C(row, 4),
            "duration_h": C(row, 5),
            "rop_m_h": C(row, 6),
            "od": _clean(C(row, 7)),
            "make": _clean(C(row, 8)),
        })

    # =====================================================================
    # OPERATIONS LOG — A=start, B=end, C=description (merged wide),
    # M=tarif code, N=duration in plain hours. Rows 24 onward, stopping at
    # the "Après minuit :" label row.
    # =====================================================================
    activities = []
    for row in range(24, 44):
        a = C(row, 1)     # A
        b = C(row, 2)     # B
        desc = _clean(C(row, 3))     # C
        code = _clean(C(row, 13)).upper()   # M
        dur = C(row, 14)                    # N

        if desc.rstrip().endswith(":") and re.search(r"APRES MINUIT|APRÈS MINUIT", desc.upper()):
            break

        start_t = a if isinstance(a, time) else None
        end_t = b if isinstance(b, time) else None

        if start_t is None and end_t is None:
            if desc and activities:
                activities[-1]["description"] = (
                    activities[-1]["description"] + "\n" + desc
                ).strip()
            continue
        if not desc:
            continue

        sh, eh = _hours_from_time(start_t), _hours_from_time(end_t)
        if sh is not None and eh is not None:
            diff = eh - sh
            hrs = diff if diff > 0 else (24.0 if diff == 0 else diff + 24)
        else:
            hrs = _float(dur, 0.0)

        activities.append({
            "start_time": start_t,
            "end_time": end_t,
            "hours": round(hrs, 2),
            "phase_name": "",
            "code": "",
            "sub": "",
            "description": desc,
            "start_md": 0, "end_md": 0,
            "npt": 0, "npt_detail": "",
            "npt_company": "", "op_company": "",
            "bill": code if _TARIF_CODE_RE.match(code) else "",
        })

    # =====================================================================
    # TARIF hours-by-code — primary source: M(code)/N(hours) tally next to
    # the operations log, rows 24-42 (stop at the '∑' total row). Falls
    # back to nothing if that tally is empty; the "TARIFICATION / Jour" row
    # (Q22=T1, S22=T2, T22=T3, V22=T4) is kept separately as a cross-check.
    # =====================================================================
    tarif_totals = {}
    for row in range(24, 43):
        code = _clean(C(row, 13)).upper()   # M
        if code in ("", "∑", "TOTAL"):
            continue
        if not _TARIF_CODE_RE.match(code):
            continue
        dur = _float(C(row, 14), 0.0)        # N
        if dur:
            tarif_totals[code] = round(tarif_totals.get(code, 0) + dur, 2)

    total_h = C(42, 14)                     # N42 ('∑' row)
    if total_h:
        tarif_totals["_total"] = _float(total_h)

    tarif_daily_reported = {}
    for code, col in (("T1", 17), ("T2", 19), ("T3", 20), ("T4", 22)):   # Q,S,T,V
        v = C(22, col)
        if v is not None:
            tarif_daily_reported[code] = _float(v)

    # =====================================================================
    # TEXT SECTIONS
    # =====================================================================
    text_sections = {}
    after_midnight = _clean(C(45, 3))       # C45, under "Après minuit:" (C44)
    if after_midnight:
        text_sections["situation_after_midnight"] = after_midnight

    situation = _clean(C(55, 4))            # D55, under "Situation à 06:00"
    if situation:
        text_sections["current_operation"] = situation
        text_sections["day_summary"] = situation

    programme = _clean(C(56, 4))            # D56, under "Programme prévue"
    if programme:
        text_sections["plan_operations"] = programme

    notes = _clean(C(52, 3))                # C52, big free-text "Note:" block
    if notes:
        text_sections["notes"] = notes

    # =====================================================================
    # MUD CHECKS  (rows 15-21, cols W/X label-value and Y/Z label-value)
    # =====================================================================
    MUD_KEYS = {
        "densité": "density", "v.marsh": "fun_vis", "filtrat": "apl_fl",
        "sabl%": "sand_pct", "solide%": "solids_pct", "huile": "oil_pct",
        "h/e": "oil_water_ratio", "y.p": "yp", "gel 0": "gel0", "gel  0": "gel0",
        "gel 10": "gel10", "gel  10": "gel10", "pb": "pb", "pv": "pv", "lgs": "lgs",
        "vb": "vb",
    }
    mud_checks = {}
    if mud_type:
        mud_checks["mud_type"] = mud_type
    for row in range(15, 22):
        for label, value in _labelled_row_values(ws, row, 23, 26):    # W:Z
            key = MUD_KEYS.get(label.lower())
            if key and value is not None:
                mud_checks[key] = _float(value)

    mud_volume = {}
    VOL_KEYS = {
        "perte formation": "perte_formation", "perte surface": "perte_surface",
        "tripping": "perte_tripping", "volume puits (m3)": "string_volume",
        "volume surface (m3)": "pits_volume",
    }
    for row in range(10, 15):
        for label, value in _labelled_row_values(ws, row, 23, 26):
            key = VOL_KEYS.get(label.lower())
            if key and value is not None:
                mud_volume[key] = _float(value)

    # =====================================================================
    # MUD CHEMICAL USAGE ("PRODUITS") — col W=item, X='Utilisé', Z='Stock'
    # =====================================================================
    chemicals = []
    for row in range(24, 47):
        item = ws.cell(row, 23).value     # W (raw, not merge-filled)
        if not item:
            continue
        item = _strip_stray_ref(_clean(item))
        used = ws.cell(row, 24).value      # X
        stock = ws.cell(row, 26).value     # Z
        chemicals.append({
            "item": item,
            "units": "",
            "received": "",
            "used": "" if used is None else used,
            "on_loc": "" if stock is None else stock,
        })

    wb.close()

    return {
        "header": header,
        "activities": activities,
        "text_sections": text_sections,
        "mud_checks": mud_checks,
        "mud_volume": mud_volume,
        "mud_chemical_usage": chemicals,
        "personnel_data": [],   # this template has no personnel table
        "pumps": [],
        "well_location": {},
        "survey_data": [],
        "safety": {},
        "tarif_totals": tarif_totals,
        # Extra info beyond the common dict shape:
        "bit_runs": bit_runs,
        "tarif_daily_reported": tarif_daily_reported,
    }


# Drop-in compat for the converter's import
parse_daily_excel_report = parse_enf18_report
parse_ddr = parse_enf18_report


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: enf18_extract.py SOURCE.xlsx")
    data = parse_enf18_report(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)):
            return o.isoformat()
        if isinstance(o, time):
            return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta):
            return o.total_seconds()
        return str(o)

    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
