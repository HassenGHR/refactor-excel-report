#!/usr/bin/env python3
"""
tp183_extract.py — extract a TP-183 (rig) / MD-19 (well) / HMD (field)
Daily Work-Over Report from a native .xlsx into the standard dict shape.

Source layout
-------------
Modern .xlsx, single sheet named like "DWR 07-TP-183 MD-19". Title at
J3: "Rapport journalier work-over". A dense single-sheet French
drilling/workover report, structurally similar in spirit to
tp127_extract.py's GW29-family layout (granular AVANCEMENT/OUTILS/
PARAMETRES/USURES engineering sections, BHA composition, deviation
surveys) but with its own distinct column layout.

Known layout notes
-------------------
* HEADER: "La date" (Y3, value at Y4 — row BELOW, same column) and
  "WO N" (AB3, value at AB4 — also row below) follow the tp219-style
  "label above, value below" convention; everything else ("Puits",
  "Champ", "Appareil", "Dernier Tubage") is "label, value a cell or
  two to the right, same row" like tp127.
* "Dernier Tubage" (M6) gives only a pipe SIZE (" 7''"), no depth in
  the same cell — there's no matching depth field on this report
  instance, so last_csg_shoe is only set when the source actually has
  something numeric alongside it (it doesn't here — no fabricated
  depth is attached to a bare size).
* TARIFICATION (T24 section header) hours are the DAILY row (T26
  "Jour", values at V26:Y26 for T1-T4) — there's no separate
  cumulative row on this template.
* ACTIVITIES (rows 28+): the description (E, wide merge), bill code
  (R), and hours (S) columns are reliable; the start/end time columns
  (A/B and C/D) use THREE different, inconsistent encodings across
  rows on the same report (plain integers, text "00:00", and
  Excel timedeltas used as "seconds since midnight") — since the S
  column already gives pre-computed hours directly, hours come from
  S and start/end times are read best-effort from C/D for display
  only, never relied on for the hours figure itself.
  Consistent with the established fix (tp182/tp187/entp204): a row
  only counts as a real operation if it has an explicit bill code
  (column R). This report has an "Apres minuit:" (after midnight)
  marker (E40) followed by a continuation-description row with NO
  bill code of its own — excluded from activities by the same filter,
  and instead captured separately as `text_sections['after_midnight']`
  (same convention as entp204_extract.py). A further row (E45) is a
  "Remarque:" (remark) note, not an activity or after-midnight
  continuation — captured separately as `header['remarks']`.
* MUD CHECKS (Z15:AB24 area) are simple unmerged label/value pairs —
  read generically (immediate next cell as value, if value-like).
* PRODUITS (chemicals, Z27 header — item merged 2 cols, used at col
  AB, stock at col AC) and PERSONNEL (Z43+, role labels with counts —
  all blank on the sample report) are both read generically the same
  way as the mud-checks scan.
* "Situation @ 06:00" (C60/F60) and "Programme prévu" (C61/F61) are
  both "label, value a few cells right, same row".

Distinguishing markers (for helpers.parse_source._detect_format):
    - file extension .xlsx
    - contains "Rapport journalier work-over" + "Appareil" + a rig
      value matching "TP-183" / "TP183" (or generically "AVANCEMENTS"
      + "USURES" + "MATERIELS DE FOND" section headers together, or an
      "MD-" well prefix)
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _clean(v) -> str:
    if v is None: return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _float(v, default: float = 0.0) -> float:
    if v is None: return default
    if isinstance(v, bool): return float(v)
    if isinstance(v, (int, float)): return float(v)
    s = _clean(v).replace(",", ".").replace(" ", "")
    if not s or s in ("-", "/", "None"): return default
    stripped = re.sub(r"[a-zA-Zé°%/³]+$", "", s).strip()
    try: return float(stripped)
    except ValueError: pass
    m = re.match(r"^-?\d+(?:\.\d+)?", s)
    if m:
        try: return float(m.group(0))
        except ValueError: pass
    return default


def _int(v, default: int = 0) -> int:
    return int(_float(v, float(default)))


def _time_from_cell(v) -> Optional[time]:
    if v is None: return None
    if isinstance(v, time): return v
    if isinstance(v, datetime): return v.time()
    if isinstance(v, timedelta):
        total = int(v.total_seconds())
        if total >= 86400: return time(0, 0)
        return time((total // 3600) % 24, (total % 3600) // 60)
    s = _clean(v)
    if not s or s in ("-", "None"): return None
    if s in ("24:00", "24:00:00"): return time(0, 0)
    for fmt in ("%H:%M:%S", "%H:%M", "%Hh%M"):
        try: return datetime.strptime(s, fmt).time()
        except ValueError: continue
    return None


def _hours_from_cell(v) -> float:
    """Read an Excel duration cell — these use the TIME data type to mean
    a duration (e.g. time(11, 0) == "11 hours", NOT 11 AM), the same
    convention used for H/Jour-style hour cells in the sibling extractors."""
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    if isinstance(v, timedelta): return v.total_seconds() / 3600.0
    if isinstance(v, time): return v.hour + v.minute / 60.0
    if isinstance(v, datetime): return v.hour + v.minute / 60.0
    return _float(v)


def _date_parse(v):
    if v is None: return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date_type): return v
    s = _clean(v)
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: continue
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b", s)
    if m:
        d, mo, y = (int(x) for x in m.groups())
        if y < 100: y += 2000
        try: return date_type(y, mo, d)
        except ValueError: return None
    return None


def _build_merged_lookup(ws):
    """Value lookup for non-top-left cells of a merged range."""
    lookup = {}
    for mr in ws.merged_cells.ranges:
        av = ws.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                if (r, c) != (mr.min_row, mr.min_col):
                    lookup[(r, c)] = av
    return lookup


def _build_span_end_cols(ws):
    """For every cell inside a merged range, the range's rightmost column
    — used so a label/value scan can skip past the LABEL's own merge span
    (e.g. 'Perte TRIPING' merged W6:Y6, value at Z6) instead of re-reading
    the label's own text back via the merged-cell lookup and mistaking it
    for a value."""
    spans = {}
    for mr in ws.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                spans[(r, c)] = mr.max_col
    return spans


def _cell(ws, r, c, lookup):
    v = ws.cell(r, c).value
    return v if v is not None else lookup.get((r, c))


def _find_cell(ws, lookup, target, row_range, col_range):
    """(row, col) of the first cell whose text EQUALS target
    (case-insensitive, whitespace-normalized, colon-stripped)."""
    t = _clean(target).upper().rstrip(":").strip()
    for r in range(row_range[0], row_range[1] + 1):
        for c in range(col_range[0], col_range[1] + 1):
            v = _cell(ws, r, c, lookup)
            if v is None: continue
            if _clean(str(v)).upper().rstrip(":").strip() == t:
                return (r, c)
    return None


def _value_right(ws, lookup, row, after_col, max_scan=6, spans=None):
    """Nearest non-empty cell to the right of (row, after_col) — skipping
    past the label's OWN merge span first if `spans` is given."""
    start = after_col
    if spans is not None:
        start = spans.get((row, after_col), after_col)
    for c in range(start + 1, start + 1 + max_scan):
        v = _cell(ws, row, c, lookup)
        if v is None or _clean(str(v)) == "":
            continue
        return v
    return None


def _is_value_like(v) -> bool:
    """True if a cell looks like a real DATA value (number, date/time, or
    a short numeric-ish string like '90/10' or '10/17') rather than
    another label — used when scanning a row for label/value pairs with
    irregular spacing, so a label immediately followed by ANOTHER label
    (no value in between) isn't mistaken for having one."""
    if v is None: return False
    if isinstance(v, (int, float, datetime, date_type, time, timedelta)):
        return True
    s = _clean(v)
    return bool(re.match(r"^[\d.,/\-\s]+$", s)) and s != ""


def _slug(label: str) -> str:
    s = _clean(label).lower()
    s = re.sub(r"[^\w]+", "_", s)
    return re.sub(r"_+", "_", s).strip("_")


_MUD_LABEL_ALIASES = {
    "densité": "density", "densite": "density",
    "v. marsh": "fun_vis", "v marsh": "fun_vis",
    "filtrat": "filtrat",
    "e stability": "electrical_stability",
    "solide (%)": "solids_pct",
    "oil/w/ratio": "oil_water_ratio",
    "yield- p": "yield_pt", "yield p": "yield_pt",
    "gel 0/10": "gel_0_10",
    "plast vis": "pv",
    "ph": "ph",
    "perte triping": "tripping_loss",
    "transfert": "transfert_vol",
    "volume puits (m3) d": "string_volume",
    "totale surface (m3)": "pits_volume",
}


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_tp183(source: Union[Path, str, BytesIO]) -> dict:
    wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)
    S = _build_span_end_cols(ws)

    header: Dict[str, Any] = {}

    # =====================================================================
    # HEADER
    # =====================================================================
    # "La date" / "WO N" — label above, value directly BELOW (same col).
    pos = _find_cell(ws, L, "La date", (1, 5), (20, 30))
    if pos:
        v = _cell(ws, pos[0] + 1, pos[1], L)
        d = _date_parse(v)
        if d: header["date"] = d
    pos = _find_cell(ws, L, "WO N", (1, 5), (25, 32))
    if pos:
        v = _cell(ws, pos[0] + 1, pos[1], L)
        if v is not None:
            m = re.search(r"\d+", _clean(str(v)))
            if m: header["day_number"] = int(m.group(0))

    # "Puits" / "Champ" / "Appareil" — label, value a cell or two right,
    # same row.
    pos = _find_cell(ws, L, "Puits", (5, 8), (1, 6))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v: header["well_name"] = _clean(str(v))
    pos = _find_cell(ws, L, "Champ", (5, 8), (1, 10))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v: header["field_name"] = _clean(str(v))
    pos = _find_cell(ws, L, "Appareil", (5, 8), (1, 14))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v:
            rig = _clean(str(v)).upper()
            m = re.match(r"^([A-Z]+)\s*#?\s*-?\s*(\d+)$", rig)
            header["rig_name"] = f"{m.group(1)}-{m.group(2)}" if m else _clean(str(v))

    # "Dernier Tubage" — this report instance gives only a pipe SIZE
    # (" 7''"), no depth in the same cell or the next one over. No
    # last_csg_shoe is fabricated from a bare size with no matching depth.
    pos = _find_cell(ws, L, "Dernier Tubage", (5, 8), (10, 16))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v:
            size = _clean(str(v))
            header["last_csg_size"] = size
            dm = re.search(r"([\d.,]+)", size)
            # Only a bare size like " 7''" here — no digits at all besides
            # the pipe fraction, so this deliberately does NOT become a
            # depth (re.search below would wrongly grab "7" as a depth).
            # Kept only as the raw size field.

    # Mud type ("Type" label, value in the next merged cell)
    pos = _find_cell(ws, L, "Type", (5, 8), (24, 28))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v: header["mud_type"] = _clean(str(v))

    # Current bit/drilling parameters (AVANCEMENT/OUTILS/PARAMETRES block)
    # — the day's real operational figures, captured as a flat dict.
    current_params = {}
    for r in range(9, 13):
        row_vals = {}
        for c, key in ((4, "outil_no"), (5, "profondeur"), (6, "avance"),
                       (12, "debit"), (13, "pression"), (14, "rpm"),
                       (15, "poids"), (16, "rob")):
            v = _cell(ws, r, c, L)
            if isinstance(v, (int, float)):
                row_vals[key] = float(v)
        if row_vals:
            current_params.update(row_vals)
    if current_params:
        header["current_params"] = current_params

    # Situation @ HH:MM / Programme prévu — label, value a few cells
    # right, same row.
    pos = None
    for r in range(58, 65):
        for c in range(1, 6):
            v = _cell(ws, r, c, L)
            if v and _clean(str(v)).upper().startswith("SITUATION"):
                pos = (r, c); break
        if pos: break
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v: header["situation"] = _clean(str(v))
    pos = None
    for r in range(58, 65):
        for c in range(1, 6):
            v = _cell(ws, r, c, L)
            if v and "PROGRAMME" in _clean(str(v)).upper():
                pos = (r, c); break
        if pos: break
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v: header["plan_operations"] = _clean(str(v))

    # =====================================================================
    # TARIFICATION — T1-T4 header row, then the "Jour" (daily) hours row
    # right below it (no separate cumulative row on this template).
    # =====================================================================
    tarif_totals: Dict[str, float] = {}
    t1_pos = _find_cell(ws, L, "T1", (20, 28), (18, 26))
    if t1_pos:
        hdr_row, t1_col = t1_pos
        codes = []
        for c in range(t1_col, t1_col + 4):
            v = _cell(ws, hdr_row, c, L)
            if isinstance(v, str) and re.match(r"^T\d$", _clean(v)):
                codes.append((c, _clean(v).lower()))
        jour_row = None
        for r in range(hdr_row + 1, hdr_row + 3):
            label = _clean(str(_cell(ws, r, 20, L) or ""))
            if "JOUR" in label.upper(): jour_row = r; break
        if jour_row is not None:
            for c, code in codes:
                v = _cell(ws, jour_row, c, L)
                if v is not None: tarif_totals[code] = _float(v)

    # =====================================================================
    # ACTIVITIES: E=description (wide merge), R=bill code, S=hours
    # (pre-computed by the source spreadsheet). A/B and C/D start/end-time
    # columns use inconsistent encodings across rows on this report (plain
    # integers, text "00:00", Excel timedeltas as seconds-since-midnight),
    # so hours always come from S — never recomputed from a start/end
    # difference — and C/D are read best-effort for display only. Only
    # rows with an explicit bill code count as real operations (see
    # module docstring), which excludes the "Apres minuit" continuation.
    # =====================================================================
    activities: List[Dict[str, Any]] = []
    for r in range(28, 45):
        bill = _clean(_cell(ws, r, 18, L) or "")   # col R
        if not bill:
            continue
        desc = _clean(ws.cell(r, 5).value or "")   # col E
        hrs_v = _cell(ws, r, 19, L)                 # col S
        hours = _float(hrs_v)

        start_t = _time_from_cell(ws.cell(r, 3).value)  # col C
        end_t   = _time_from_cell(ws.cell(r, 4).value)  # col D

        activities.append({
            "start_time": start_t, "end_time": end_t, "hours": hours,
            "phase_name": "", "code": "", "sub": "",
            "description": desc,
            "start_md": 0, "end_md": 0,
            "npt": 0, "npt_detail": "", "npt_company": "", "op_company": "",
            "bill": bill,
        })

    # =====================================================================
    # AFTER MIDNIGHT / REMARKS — supplementary description text, not
    # activities (no bill code of their own). Captured separately, same
    # convention as entp204_extract.py's after_midnight_summary.
    # =====================================================================
    text_sections: Dict[str, str] = {}
    am_pos = None
    for r in range(28, 45):
        v = ws.cell(r, 5).value
        if v and "APRES MINUIT" in _clean(str(v)).upper():
            am_pos = r; break
    if am_pos:
        parts = []
        for r in range(am_pos, am_pos + 6):
            v = ws.cell(r, 5).value
            if not v: continue
            vs = _clean(str(v))
            if not vs: continue
            if vs.upper().startswith("APRES MINUIT"):
                vs = re.sub(r"^APRES\s*MINUIT\s*:?\s*", "", vs, flags=re.IGNORECASE)
            if vs.upper().startswith("REMARQUE"):
                # A "Remarque:" line is a distinct note, not part of the
                # after-midnight continuation — captured separately below.
                continue
            if vs and vs not in parts:
                parts.append(vs)
        if parts:
            text_sections["after_midnight"] = " ".join(parts)

    for r in range(28, 50):
        v = ws.cell(r, 5).value
        if v and _clean(str(v)).upper().startswith("REMARQUE"):
            remark = re.sub(r"^REMARQUE\s*:?\s*", "", _clean(str(v)), flags=re.IGNORECASE)
            if remark: header["remarks"] = remark
            break

    if header.get("situation"):
        text_sections["current_operation"] = header["situation"]
    if header.get("plan_operations"):
        text_sections["plan_operations"] = header["plan_operations"]

    # =====================================================================
    # MUD CHECKS — simple unmerged label/value pairs (immediate next cell
    # as value, if it looks like real data rather than another label).
    # =====================================================================
    mud_checks: Dict[str, Any] = {}
    if "mud_type" in header:
        mud_checks["mud_type"] = header["mud_type"]
    for r in range(15, 25):
        for lab_c in (26, 28):   # col Z, col AB
            v = _cell(ws, r, lab_c, L)
            if v is None or not isinstance(v, str) or not _clean(v):
                continue
            label = _clean(v)
            norm = label.lower().rstrip(":").strip()
            val = _value_right(ws, L, r, lab_c, max_scan=2, spans=S)
            if val is None or not _is_value_like(val):
                continue
            key = _MUD_LABEL_ALIASES.get(norm, _slug(label))
            if not key or key in mud_checks: continue
            vs = _clean(str(val))
            mud_checks[key] = _float(val) if re.match(r"^-?[\d.,]+$", vs) else vs

    # =====================================================================
    # PRODUITS (chemicals) — item (col Z, sometimes merged 2 cols), used
    # (col AB), stock (col AC). Bounded from the "Produits" section header
    # down to the "TOTAL" row.
    # =====================================================================
    chemicals: List[Dict[str, Any]] = []
    prod_pos = None
    for r in range(24, 30):
        for c in range(26, 32):
            v = _cell(ws, r, c, L)
            if v and _clean(str(v)).upper().startswith("UTILISE"):
                prod_pos = (r, 26); break
        if prod_pos: break
    if prod_pos:
        for r in range(prod_pos[0] + 1, prod_pos[0] + 20):
            item = _cell(ws, r, 26, L)  # col Z
            if not item or not isinstance(item, str): continue
            item_s = _clean(item)
            if not item_s: continue
            if item_s.upper().startswith("TOTAL"): break
            used  = _cell(ws, r, 28, L)   # col AB
            stock = _cell(ws, r, 29, L)   # col AC
            chemicals.append({
                "item": item_s, "units": "T", "received": "",
                "used":  _clean(str(used)) if used is not None else "",
                "on_loc": _clean(str(stock)) if stock is not None else "",
            })

    # =====================================================================
    # PERSONNEL — role labels (col Z) from the row after the chemicals
    # TOTAL down to "TOTAL PERSON". Counts (if filled in) sit a few
    # columns to the right; all blank on the sample report.
    # =====================================================================
    personnel: List[Dict[str, Any]] = []
    total_pos = None
    for r in range(40, 46):
        v = _cell(ws, r, 26, L)
        if v and _clean(str(v)).upper().startswith("TOTAL"):
            total_pos = r; break
    if total_pos:
        for r in range(total_pos + 1, total_pos + 22):
            label = _cell(ws, r, 26, L)  # col Z
            if not label or not isinstance(label, str): continue
            label_s = _clean(label)
            if not label_s: continue
            if label_s.upper().startswith("TOTAL PERSON"):
                v = _value_right(ws, L, r, 26, max_scan=3, spans=S)
                if v is not None and _is_value_like(v):
                    header["personnel_total"] = _int(v)
                break
            v = _value_right(ws, L, r, 26, max_scan=3, spans=S)
            n = _int(v) if (v is not None and _is_value_like(v)) else 0
            personnel.append({"company": label_s, "number": n, "hours": "", "names": ""})

    wb.close()

    return {
        "header": header,
        "activities": activities,
        "text_sections": text_sections,
        "mud_checks": mud_checks,
        "mud_volume": {},
        "mud_chemical_usage": chemicals,
        "personnel_data": personnel,
        "pumps": [],
        "well_location": {},
        "survey_data": [],
        "safety": {},
        "tarif_totals": tarif_totals,
        "tarif_cumul": {},
    }


# Drop-in compat
parse_daily_excel_report = parse_tp183
parse_gw29 = parse_tp183


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: tp183_extract.py SOURCE.xlsx")
    data = parse_tp183(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
