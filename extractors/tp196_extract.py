#!/usr/bin/env python3
"""
tp196_extract.py — extract a TP-196 (rig) / OMJ-701 (well) / HMD (field)
Daily Work-Over Report from a native .xlsx into the standard dict shape.

Source layout
-------------
Modern .xlsx, single sheet named like "Nr 06 Le 03-07-2017" (the sheet
NAME's date is a stale leftover from an older report cycle — the real
report date lives in the "DU" header cell and is read from there, not
the sheet name). Title at H1: "Rapport journalier de Work - Over".
Same GW29-family layout as tp127/tp183 (AVANCEMENT/OUTILS/USURE/
PARAMETRES engineering sections, BHA composition, deviation surveys),
with its own header/column offsets.

Known layout notes
-------------------
* HEADER: unlike tp183 (date/day-number below their labels), THIS
  report uses "label, value a cell or two right, SAME row" for
  everything, including "DU" (S1, date at U1) and "N°" (Y1, day
  number at Z1) — the opposite convention from tp183 for those two
  fields specifically, so don't assume "row below" carries over.
* "Last Tubage 7\"" (K2) and "Top Liner 4\"1/2" (N2) are both bare
  labels with NO adjacent depth value on this report instance — no
  last_csg_shoe/last_lnr_top is fabricated from a size with nothing
  next to it; only the raw size text is kept.
* STALE LEFTOVER DATA: columns AE/AF (rows 12-16) contain a
  "Tarification appareil TP 197" reference block — cost/tariff figures
  for a COMPLETELY DIFFERENT RIG, apparently leftover from a shared
  template — and columns AC/AD/AE/AF (rows 19-40) contain a dated
  table running from 2017-06-28 to 2017-07-19, clearly a stale
  artifact from a much older report cycle. Both blocks happen to
  contain "T1"/"T2"/"T3"/"T4"-looking text, so the real TARIFICATION
  search is bounded tightly to columns S-V (where the report's OWN
  T1-T4 table actually lives, row 15) to avoid matching them.
* TARIFICATION (O14 section header, T1-T4 at row 15, "Jour" row 16,
  "Cumul" row 17).
* ACTIVITIES (rows 18+): A=start, B=end, C=description (wide merge),
  M=bill code, N=hours (Excel timedelta). Note: on the sample report
  ALL SIX main-table activities are tagged bill code "T1" in column M,
  while the "Jour" tarification row (S16:V16) shows T1=12.5h/T2=11.5h
  — a genuine inconsistency in the SOURCE data (the daily summary
  wasn't reconciled against the individual line-item tags). Both are
  extracted faithfully as-is rather than one being silently "corrected"
  to match the other. Consistent with the established fix (tp182/
  tp183/tp187/entp204): a row only counts as a real operation if it
  has an explicit bill code, which excludes the row-38 "total" check
  row (M38='total', not a T-code) and the after-midnight continuation.
* "Après minuit :" (C33) + continuation (C34) — captured separately as
  `text_sections['after_midnight']`, not as an activity (no bill code).
* A "NB:" remark line (C31) — captured as `header['remarks']`.
* "Situation au rapport" (A39/D39) and "Programme prévu" (A40/D40) —
  label, value a few cells right, same row.
* "Représentant SH/DP" (W39, value BELOW at W40) — mapped onto
  `supervisor`, the same way other reports map their site-rep field.
* MUD CHECKS / mud volumes / PRODUITS (chemicals, W17 title) /
  PERSONNEL (W31 title) all live in columns W-Z with the same
  irregular label/value spacing as tp183 (some labels span W:X, some
  W:Y) — read the same generic, span-aware way.

Distinguishing markers (for helpers.parse_source._detect_format):
    - file extension .xlsx
    - contains "Rapport journalier de Work - Over" + "Appareil" + a rig
      value matching "TP 196" / "TP-196" (or generically "AVANCEMENT" +
      "USURE" + "MATERIEL DE FOND" section headers together, or an
      "OMJ" well prefix)
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
    "h/e": "oil_water_ratio",
    "salinité": "salinity", "salinite": "salinity",
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

# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_tp196(source: Union[Path, str, BytesIO]) -> dict:
    wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)
    S = _build_span_end_cols(ws)

    header: Dict[str, Any] = {}

    # =====================================================================
    # HEADER — "label, value a cell or two right, same row" throughout,
    # including DU/N° (unlike tp183, which has those two below their
    # labels — don't assume that convention carries over between reports).
    # =====================================================================
    pos = _find_cell(ws, L, "DU", (1, 1), (17, 24))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=4, spans=S)
        d = _date_parse(v)
        if d: header["date"] = d

    pos = _find_cell(ws, L, "N°", (1, 1), (23, 28))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v is not None:
            header["day_number"] = _int(v)

    pos = _find_cell(ws, L, "Puits", (1, 3), (1, 6))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v: header["well_name"] = _clean(str(v))

    pos = _find_cell(ws, L, "Champ", (1, 3), (1, 10))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v: header["field_name"] = _clean(str(v))

    pos = _find_cell(ws, L, "Appareil", (1, 3), (1, 14))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v:
            rig = _clean(str(v)).upper()
            m = re.match(r"^([A-Z]+)\s*#?\s*-?\s*(\d+)$", rig)
            header["rig_name"] = f"{m.group(1)}-{m.group(2)}" if m else _clean(str(v))

    # "Last Tubage 7\"" / "Top Liner 4\"1/2" — bare labels; on this
    # report instance there's no adjacent depth value at all, so only
    # the raw size text is kept (no fabricated last_csg_shoe/last_lnr_top
    # from a size with nothing next to it).
    for prefix, raw_key in ((["Last Tubage"], "last_csg_size_raw"),
                             (["Top Liner"], "last_lnr_top_size_raw")):
        pos = None
        for r in range(1, 4):
            for c in range(1, 22):
                v = _cell(ws, r, c, L)
                if v and _clean(str(v)).upper().startswith(prefix[0].upper()):
                    pos = (r, c); break
            if pos: break
        if pos:
            label_text = _clean(str(_cell(ws, *pos, L)))
            header[raw_key] = label_text
            v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
            if v is not None and _is_value_like(v):
                # A real depth WAS found next to the label — build the
                # standardized field after all.
                m = re.search(r'(\d+"?[\d/\s]*)$', label_text)
                size = m.group(1).strip() if m else label_text
                depth = _float(v)
                if depth:
                    dest = "last_csg" if "csg" in raw_key else "last_lnr_top"
                    header[f"{dest}_size"] = size
                    header[f"{dest}_depth"] = depth
                    header[f"{dest}_shoe" if dest == "last_csg" else dest] = f"{size} @ {depth:g}m"

    # Mud type
    pos = _find_cell(ws, L, "Type", (2, 4), (22, 26))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v: header["mud_type"] = _clean(str(v))

    # Situation au rapport / Programme prévu — label, value a few cells
    # right, same row.
    pos = _find_cell(ws, L, "Situation au rapport", (36, 42), (1, 4))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v: header["situation"] = _clean(str(v))
    pos = _find_cell(ws, L, "Programme prévu", (36, 42), (1, 4))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=3, spans=S)
        if v: header["plan_operations"] = _clean(str(v))

    # "Représentant SH/DP" — value BELOW the label this time.
    pos = _find_cell(ws, L, "Représentant SH/DP", (36, 42), (20, 26))
    if pos:
        v = _cell(ws, pos[0] + 1, pos[1], L)
        if v:
            header["site_representative"] = _clean(str(v))
            header["supervisor"] = header["site_representative"]

    # A "NB:"-prefixed remark line.
    for r in range(28, 33):
        v = ws.cell(r, 3).value
        if v and _clean(str(v)).upper().startswith("NB"):
            remark = re.sub(r"^NB\s*:?\s*", "", _clean(str(v)), flags=re.IGNORECASE)
            if remark: header["remarks"] = remark
            break

    # =====================================================================
    # TARIFICATION — bounded tightly to columns S-V (18-22), where this
    # report's OWN T1-T4 table lives, to avoid the stale "Tarification
    # appareil TP 197" leftover block sitting in columns AE/AF (see
    # module docstring) which also contains "T1"/"T2"/"T3"/"T4" text.
    # =====================================================================
    tarif_totals: Dict[str, float] = {}
    tarif_cumul: Dict[str, float] = {}
    t1_pos = _find_cell(ws, L, "T1", (13, 18), (18, 22))
    if t1_pos:
        hdr_row, t1_col = t1_pos
        codes = []
        for c in range(t1_col, t1_col + 4):
            v = _cell(ws, hdr_row, c, L)
            if isinstance(v, str) and re.match(r"^T\d$", _clean(v)):
                codes.append((c, _clean(v).lower()))
        jour_row = None
        cumul_row = None
        for r in range(hdr_row + 1, hdr_row + 4):
            label = _clean(str(ws.cell(r, 15).value or ""))  # col O
            if "JOUR" in label.upper(): jour_row = r
            if "CUMUL" in label.upper(): cumul_row = r
        for c, code in codes:
            if jour_row is not None:
                v = _cell(ws, jour_row, c, L)
                if v is not None: tarif_totals[code] = _float(v)
            if cumul_row is not None:
                v = _cell(ws, cumul_row, c, L)
                if v is not None: tarif_cumul[code] = _float(v)

    # =====================================================================
    # ACTIVITIES: A=start, B=end, C=description (wide merge), M=bill
    # code, N=hours (Excel timedelta). Only rows with an explicit bill
    # code count as real operations — see module docstring re: the
    # T1-everywhere-in-column-M quirk on this report, and the row-38
    # "total" check row (M38='total', not a real T-code) that this
    # filter also correctly excludes.
    # =====================================================================
    activities: List[Dict[str, Any]] = []
    for r in range(18, 38):
        bill = _clean(ws.cell(r, 13).value or "")   # col M
        if not bill or not re.match(r"^T\d$", bill):
            continue
        desc = _clean(ws.cell(r, 3).value or "")     # col C
        hrs_v = _cell(ws, r, 14, L)                   # col N
        hours = _hours_from_cell(hrs_v)

        start_t = _time_from_cell(ws.cell(r, 1).value)  # col A
        end_t   = _time_from_cell(ws.cell(r, 2).value)  # col B

        activities.append({
            "start_time": start_t, "end_time": end_t, "hours": hours,
            "phase_name": "", "code": "", "sub": "",
            "description": desc,
            "start_md": 0, "end_md": 0,
            "npt": 0, "npt_detail": "", "npt_company": "", "op_company": "",
            "bill": bill,
        })

    # =====================================================================
    # AFTER MIDNIGHT — supplementary description text, not an activity
    # (no bill code of its own). Same convention as entp204/tp183.
    # =====================================================================
    text_sections: Dict[str, str] = {}
    am_pos = None
    for r in range(28, 38):
        v = ws.cell(r, 3).value
        if v and "MINUIT" in _clean(str(v)).upper():
            am_pos = r; break
    if am_pos:
        parts = []
        for r in range(am_pos, am_pos + 6):
            v = ws.cell(r, 3).value
            if not v: continue
            vs = _clean(str(v))
            if not vs: continue
            if vs.upper().startswith("APR") and "MINUIT" in vs.upper():
                vs = re.sub(r"^Apr[eè]s\s*minuit\s*:?\s*", "", vs, flags=re.IGNORECASE)
            if vs and vs not in parts:
                parts.append(vs)
        if parts:
            text_sections["after_midnight"] = " ".join(parts)

    if header.get("situation"):
        text_sections["current_operation"] = header["situation"]
    if header.get("plan_operations"):
        text_sections["plan_operations"] = header["plan_operations"]

    # =====================================================================
    # MUD CHECKS — generic span-aware label/value scan, columns W-Z.
    # =====================================================================
    mud_checks: Dict[str, Any] = {}
    if "mud_type" in header:
        mud_checks["mud_type"] = header["mud_type"]
    for r in range(6, 17):
        c = 23  # column W
        while c <= 26:  # through column Z
            v = _cell(ws, r, c, L)
            if v is None or not isinstance(v, str) or not _clean(v):
                c += 1; continue
            label = _clean(v)
            norm = label.lower().rstrip(":").strip()
            val = _value_right(ws, L, r, c, max_scan=3, spans=S)
            span_end = S.get((r, c), c)
            c = span_end + 1
            if val is None or not _is_value_like(val):
                continue
            key = _MUD_LABEL_ALIASES.get(norm, _slug(label))
            if not key or key in mud_checks: continue
            vs = _clean(str(val))
            mud_checks[key] = _float(val) if re.match(r"^-?[\d.,]+$", vs) else vs

    # =====================================================================
    # PRODUITS (chemicals) — item (col W, sometimes merged W:X), used
    # (col Y), stock (col Z). Bounded from the "PRODUITS" title down to
    # the "PERSONNEL" title.
    # =====================================================================
    chemicals: List[Dict[str, Any]] = []
    prod_pos = _find_cell(ws, L, "PRODUITS", (16, 18), (22, 27))
    pers_pos = _find_cell(ws, L, "PERSONNEL", (29, 33), (22, 27))
    ceiling = pers_pos[0] if pers_pos else 31
    if prod_pos:
        for r in range(prod_pos[0] + 1, ceiling):
            item = _cell(ws, r, 23, L)  # col W
            if not item or not isinstance(item, str): continue
            item_s = _clean(item)
            if not item_s or item_s.upper() in ("UTILISÉS", "STOCK"): continue
            used  = _cell(ws, r, 25, L)   # col Y
            stock = _cell(ws, r, 26, L)   # col Z
            chemicals.append({
                "item": item_s, "units": "", "received": "",
                "used":  _clean(str(used)) if used is not None else "",
                "on_loc": _clean(str(stock)) if stock is not None else "",
            })

    # =====================================================================
    # PERSONNEL — label (col W, merged W:Y) + count (col Z).
    # =====================================================================
    personnel: List[Dict[str, Any]] = []
    if pers_pos:
        for r in range(pers_pos[0] + 1, pers_pos[0] + 8):
            label = _cell(ws, r, 23, L)  # col W
            if not label or not isinstance(label, str): continue
            label_s = _clean(label)
            if not label_s: continue
            v = _cell(ws, r, 26, L)  # col Z
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
        "tarif_cumul": tarif_cumul,
    }

# Drop-in compat
parse_daily_excel_report = parse_tp196
parse_gw29 = parse_tp196


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: tp196_extract.py SOURCE.xlsx")
    data = parse_tp196(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
