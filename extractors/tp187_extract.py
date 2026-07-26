#!/usr/bin/env python3
"""
tp187_extract.py — extract a TP-187 (rig) / TG-61 (well) / IN AMENAS
(region) Daily Work-Over Report from a native .xlsx into the standard
dict shape.

Source layout
-------------
Modern .xlsx, single sheet named like "RAP N°(01)". Title at E1:
"RAPPORT JOURNALIER \\nWORK OVER". This is a different report family
from the SHDP/tp185/tp219 English-labeled ones and from the
PDF-exported TP-212/ENF-03/TP-217/TP-237 family — French-labeled,
native Excel, region "IN AMENAS".

Known layout notes
-------------------
* HEADER FIELDS mostly follow a "label on one row, value directly
  below it in the SAME column" convention (e.g. "PUITS" at A3, value
  "TG-61" at A4; "JOURNEE DU" at K1, date at K2; "APPAREIL" at D3, rig
  at D4). This is the same pattern whose ABSENCE caused a real bug in
  tp219_extract.py (date/day-number silently defaulting because the
  code only scanned rightward on the same row) — here we check the
  cell directly below the label FIRST, for every header field.
* "TYPE BOUE" is the one header field that does NOT follow that
  pattern — its value ("WBM") sits on the SAME row, two columns to the
  right (K3 label, M3 value).
* "Dernier Tubage" (E3/E4) is a single combined string listing BOTH
  the 9-5/8" and 7" casing strings run so far, e.g. '9"5/8 (P110 47
  lbs/f) - 7" (P110 29 lbs/ft)' — not a simple "size : depth". The
  matching depth for the LAST (deepest/most recent) size mentioned is
  a separate field, "Côte sabot" (F3/F4). last_csg_shoe is built by
  pairing the last size token found in the Dernier Tubage string with
  the Côte sabot depth. "Top Liner" (H3/H4) is a separate liner-top
  field that's often "-" (no liner run yet) — last_lnr_top is only
  set when it's a real value. "Fond" (J3/J4) is the well's current/
  total depth reference, with no equivalent in the other report
  families, so it gets its own field (`fond_depth`) rather than being
  forced into an existing slot.
* TARIFICATION is a clean small table: T1-T4 column headers on one row
  (G6:J6), daily hours the row after ("H/Jour", G7:J7) and cumulative
  hours the row after that ("H/Cum", G8:J8) — no special-casing needed
  beyond locating the T1 header cell.
* ACTIVITIES (rows 10+): columns A=start, B=end, C=bill code, D=
  description (wide, but openpyxl reads the full string from the
  top-left cell of its merge regardless of how narrow the merge
  itself is — Excel allows unmerged overflow, which doesn't affect
  the cell's own .value). EVERY activity row in the observed sample
  has an explicit bill code, so — consistent with the fix applied to
  tp182_extract.py after review — a row only counts as a real
  operation if it has one; this naturally excludes the "Opération
  après minuit" section-marker row (A24, no bill/time of its own) while
  still picking up the one real post-marker activity row that follows
  it (which DOES have its own start/end/bill).
* MUD CHECKS / MUD VOLUME / PRODUITS (chemicals) are all sparse and
  inconsistently populated in practice (most fields blank on any given
  day), living in columns K-N across rows 5-21 as label/value pairs
  (sometimes same row, sometimes with a units-only cell and no number).
  Read generically by scanning that whole block for label/value pairs
  rather than hardcoding which of the ~20 possible fields will be
  filled in on a given report.
* PERSONNEL: header "PERSONNEL" at F22, role labels in F23+ (merged
  F:I), counts in column J. "TOTAL" (F31/J31) is the day's headcount
  total, stored separately from the per-role personnel_data list.
* "Rep maître œuvre" (K31/K32, the site/client representative name(s))
  and "Resp sce puits" (M31/M32, the well-service department's
  responsible person) are this report's closest equivalents to a
  supervisor/superintendent pair — mapped onto those standard fields,
  the same way Emetteur/Récepteur are mapped in the PDF report family.

Distinguishing markers (for helpers.parse_source._detect_format):
    - file extension .xlsx
    - contains "RAPPORT JOURNALIER" + "WORK OVER" + "APPAREIL" + a rig
      value matching "TP 187" (or generically "PUITS" + "TG-" well
      naming + "IN AMENAS" in the header banner)
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
    """Value lookup for non-top-left cells of a merged range (defensive —
    every field this extractor reads happens to be the top-left cell of
    its merge already, but this guards against a report instance where
    that isn't true)."""
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
    (e.g. 'TYPE BOUE' merged K3:L3) instead of re-reading the label's own
    text back via the merged-cell lookup and mistaking it for a value."""
    spans = {}
    for mr in ws.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                spans[(r, c)] = mr.max_col
    return spans


def _cell(ws, r, c, lookup):
    v = ws.cell(r, c).value
    return v if v is not None else lookup.get((r, c))


def _below(ws, r, c, lookup):
    """The value one row below (r, c), same column — the dominant
    label/value convention on this template."""
    return _cell(ws, r + 1, c, lookup)


def _find_cell(ws, lookup, target, row_range, col_range):
    """(row, col) of the first cell whose text EQUALS target
    (case-insensitive, whitespace-normalized)."""
    t = _clean(target).upper()
    for r in range(row_range[0], row_range[1] + 1):
        for c in range(col_range[0], col_range[1] + 1):
            v = _cell(ws, r, c, lookup)
            if v is None: continue
            if _clean(str(v)).upper().rstrip(":").strip() == t:
                return (r, c)
    return None


def _value_right(ws, lookup, row, after_col, max_scan=6, spans=None):
    start = after_col
    if spans is not None:
        start = spans.get((row, after_col), after_col)
    for c in range(start + 1, start + 1 + max_scan):
        v = _cell(ws, row, c, lookup)
        if v is None or _clean(str(v)) == "":
            continue
        return v
    return None


def _slug(label: str) -> str:
    s = _clean(label).lower()
    s = re.sub(r"[^\w]+", "_", s)
    return re.sub(r"_+", "_", s).strip("_")


_MUD_LABEL_ALIASES = {
    "densité": "density", "densite": "density",
    "visc. march.": "fun_vis", "visc march": "fun_vis",
    "yield": "yield_pt", "ph": "ph",
    "perte surface": "surface_loss",
    "perte accident": "accidental_loss",
    "perte formation": "formation_loss",
    "fabrication": "fabrication_vol",
    "ejection": "ejection_vol",
    "transfert": "transfert_vol",
    "v. puits": "string_volume",
    "v. surf": "pits_volume",
    "poids": "wob",
    "rotation": "rpm",
    "débit": "flow_rate", "debit": "flow_rate",
}


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_tp187(source: Union[Path, str, BytesIO]) -> dict:
    wb = load_workbook(source, data_only=True)
    ws = wb.active
    L = _build_merged_lookup(ws)
    S = _build_span_end_cols(ws)

    header: Dict[str, Any] = {}

    # =====================================================================
    # HEADER — date/day-number/well/rig/casing block (rows 1-4)
    # =====================================================================
    pos = _find_cell(ws, L, "JOURNEE DU", (1, 2), (1, 20))
    if pos:
        d = _date_parse(_below(ws, *pos, L))
        if d: header["date"] = d

    pos = _find_cell(ws, L, "RAPPORT", (1, 2), (1, 20))
    if pos:
        v = _below(ws, *pos, L)
        if isinstance(v, (int, float)):
            header["day_number"] = int(v)
        else:
            m = re.search(r"(\d+)", _clean(v))
            if m: header["day_number"] = int(m.group(1))

    pos = _find_cell(ws, L, "PUITS", (2, 4), (1, 10))
    if pos:
        v = _below(ws, *pos, L)
        if v: header["well_name"] = _clean(str(v))

    pos = _find_cell(ws, L, "APPAREIL", (2, 4), (1, 15))
    if pos:
        v = _below(ws, *pos, L)
        if v:
            rig = _clean(str(v)).upper()
            m = re.match(r"^([A-Z]+)\s*#?\s*-?\s*(\d+)$", rig)
            header["rig_name"] = f"{m.group(1)}-{m.group(2)}" if m else _clean(str(v))

    # "Dernier Tubage" (combined 9-5/8"+7" casing string) + "Côte sabot"
    # (shoe depth for whichever is the deepest/last one mentioned).
    pos = _find_cell(ws, L, "Dernier Tubage", (2, 4), (1, 15))
    if pos:
        v = _below(ws, *pos, L)
        if v: header["last_csg_string"] = _clean(str(v))
    pos = _find_cell(ws, L, "Côte sabot", (2, 4), (1, 15))
    if pos:
        v = _below(ws, *pos, L)
        if v is not None and _clean(str(v)) not in ("", "-"):
            header["last_csg_depth"] = _float(v)
            if header.get("last_csg_string"):
                sizes = re.findall(r'\d+"[\d/]*', header["last_csg_string"])
                if sizes:
                    header["last_csg_size"] = sizes[-1]
                    header["last_csg_shoe"] = f'{sizes[-1]} @ {_float(v):g}m'

    pos = _find_cell(ws, L, "Top Liner", (2, 4), (1, 20))
    if pos:
        v = _below(ws, *pos, L)
        vs = _clean(str(v)) if v is not None else ""
        if vs and vs != "-":
            header["last_lnr_top"] = vs

    pos = _find_cell(ws, L, "Fond", (2, 4), (1, 20))
    if pos:
        v = _below(ws, *pos, L)
        if v is not None and _clean(str(v)) not in ("", "-"):
            header["fond_depth"] = _float(v)

    pos = _find_cell(ws, L, "TYPE BOUE", (2, 4), (1, 20))
    if pos:
        v = _value_right(ws, L, pos[0], pos[1], max_scan=4, spans=S)
        if v: header["mud_type"] = _clean(str(v))

    # BUT DU WORK OVER — purpose text (if any) sits directly below the
    # label, same column; often blank on this template.
    pos = _find_cell(ws, L, "BUT DU WORK OVER", (5, 9), (1, 6))
    if pos:
        v = _below(ws, *pos, L)
        if v and _clean(str(v)):
            header["workover_purpose"] = _clean(str(v))

    # PROGRAMME PREVU — label + value combined in one cell.
    pos = None
    for r in range(28, 40):
        for c in range(1, 6):
            v = _cell(ws, r, c, L)
            if v and "PROGRAMME" in _clean(str(v)).upper():
                pos = (r, c); break
        if pos: break
    if pos:
        raw = _clean(str(_cell(ws, *pos, L)))
        plan = re.sub(r"^PROGRAMME\s*PR[ÉE]VU\s*:?\s*", "", raw, flags=re.IGNORECASE)
        if plan: header["plan_operations"] = plan

    # =====================================================================
    # TARIFICATION — T1-T4 column headers, then H/Jour and H/Cum rows.
    # Bounded to rows 1-9 to avoid colliding with activity bill codes
    # (also "T1"/"T2"/...) in column C from row 10 onward.
    # =====================================================================
    tarif_totals: Dict[str, float] = {}
    tarif_cumul: Dict[str, float] = {}
    t1_pos = _find_cell(ws, L, "T1", (1, 9), (5, 12))
    if t1_pos:
        hdr_row, t1_col = t1_pos
        codes = []
        for c in range(t1_col, t1_col + 4):
            v = _cell(ws, hdr_row, c, L)
            if isinstance(v, str) and re.match(r"^T\d$", _clean(v)):
                codes.append((c, _clean(v).lower()))
        hjour_row = None
        hcum_row = None
        for r in range(hdr_row + 1, hdr_row + 4):
            label = _clean(str(_cell(ws, r, 6, L) or _cell(ws, r, 1, L) or ""))
            if "H/JOUR" in label.upper(): hjour_row = r
            if "H/CUM" in label.upper(): hcum_row = r
        for c, code in codes:
            if hjour_row is not None:
                v = _cell(ws, hjour_row, c, L)
                if v is not None: tarif_totals[code] = _float(v)
            if hcum_row is not None:
                v = _cell(ws, hcum_row, c, L)
                if v is not None: tarif_cumul[code] = _float(v)

    # =====================================================================
    # ACTIVITIES (rows 10+): A=start, B=end, C=bill, D=description.
    # Only rows with an explicit bill code count as real operations (see
    # module docstring) — this naturally skips the "Opération après
    # minuit" marker row while still keeping the real activity after it.
    # =====================================================================
    activities: List[Dict[str, Any]] = []
    for r in range(10, 31):
        # Raw cell values only (NOT the merged-cell lookup) — the "Opération
        # après minuit" marker row is a WIDE merge (A24:E24), and via the
        # merge lookup that single label's text would otherwise bleed into
        # what look like the bill(C)/description(D) columns for that row.
        # None of A-D are ever legitimately merged that way on a real
        # activity row, so reading raw values is always correct here and
        # naturally makes the marker row's bill/desc genuinely blank.
        start_v = ws.cell(r, 1).value
        end_v   = ws.cell(r, 2).value
        bill    = _clean(ws.cell(r, 3).value or "")
        desc    = _clean(ws.cell(r, 4).value or "")

        if not bill:
            continue

        start_t = _time_from_cell(start_v)
        end_t   = _time_from_cell(end_v)
        hours = 0.0
        if start_t and end_t:
            sm = start_t.hour * 60 + start_t.minute
            em = end_t.hour * 60 + end_t.minute
            if em == sm: hours = 24.0
            elif em > sm: hours = (em - sm) / 60.0
            else: hours = (em + 1440 - sm) / 60.0

        activities.append({
            "start_time": start_t, "end_time": end_t, "hours": hours,
            "phase_name": "", "code": "", "sub": "",
            "description": desc,
            "start_md": 0, "end_md": 0,
            "npt": 0, "npt_detail": "", "npt_company": "", "op_company": "",
            "bill": bill,
        })

    # =====================================================================
    # MUD CHECKS / MUD VOLUME — sparse label/value pairs scattered across
    # cols K-N, rows 5-14. Read generically: known labels map to clear
    # field names; anything else falls back to a slug rather than being
    # silently dropped.
    # =====================================================================
    mud_checks: Dict[str, Any] = {}
    if "mud_type" in header:
        mud_checks["mud_type"] = header["mud_type"]
    for r in range(5, 15):
        for lab_c, val_c in ((11, 12), (13, 14)):   # K/L and M/N
            label = _clean(_cell(ws, r, lab_c, L) or "")
            if not label: continue
            val = _cell(ws, r, val_c, L)
            if val is None: continue
            vs = _clean(str(val))
            if vs == "" or vs.upper() == label.upper():
                # A merged label (e.g. "PARAMETRES DE FRAISAGE" spanning
                # K11:N11) reports its own text back for every cell in its
                # span via the merged-cell lookup — not a real value.
                continue
            norm = label.lower().rstrip(":").strip()
            key = _MUD_LABEL_ALIASES.get(norm, _slug(label))
            if not key or key in mud_checks: continue
            if re.match(r"^-?[\d.,]+$", vs):
                mud_checks[key] = _float(val)
            elif vs not in ("m3", "t", "m"):   # skip bare unit-only cells
                mud_checks[key] = vs

    # =====================================================================
    # PRODUITS (chemical/fluid usage) — item(K)/unit(L)/used(M)/stock(N),
    # scanned generically over the whole block (the list runs past the
    # "Opération après minuit" row without a second header, continuing
    # into what are really water/fuel items like Gas-oil/Eau/brut).
    # =====================================================================
    chemicals: List[Dict[str, Any]] = []
    for r in range(16, 28):
        item = _clean(_cell(ws, r, 11, L) or "")
        if not item: continue
        unit  = _clean(_cell(ws, r, 12, L) or "")
        used  = _cell(ws, r, 13, L)
        stock = _cell(ws, r, 14, L)
        chemicals.append({
            "item": item, "units": unit, "received": "",
            "used":  _clean(str(used)) if used is not None else "",
            "on_loc": _clean(str(stock)) if stock is not None else "",
        })

    # =====================================================================
    # PERSONNEL — header "PERSONNEL" (col F), role labels F+1.. with
    # counts in column J. "TOTAL" row stored separately.
    # =====================================================================
    personnel: List[Dict[str, Any]] = []
    pers_pos = _find_cell(ws, L, "PERSONNEL", (18, 24), (5, 10))
    scan_from = pers_pos[0] + 1 if pers_pos else 23
    for r in range(scan_from, scan_from + 12):
        label = _clean(_cell(ws, r, 6, L) or "")
        if not label: continue
        if label.upper() == "TOTAL":
            v = _cell(ws, r, 10, L)
            if v is not None: header["personnel_total"] = _int(v)
            break
        v = _cell(ws, r, 10, L)
        if v is None: continue
        personnel.append({"company": label, "number": _int(v), "hours": "", "names": ""})

    # =====================================================================
    # "Rep maître œuvre" / "Resp sce puits" — this report's closest
    # equivalents to a supervisor/superintendent pair (site client rep
    # vs. well-service dept. responsible), mapped the same way
    # Emetteur/Récepteur are mapped in the PDF report family.
    # =====================================================================
    pos = _find_cell(ws, L, "Rep maître œuvre", (28, 34), (10, 14))
    if pos:
        v = _below(ws, *pos, L)
        if v:
            header["site_representative"] = _clean(str(v))
            header["supervisor"] = header["site_representative"]
    pos = _find_cell(ws, L, "Resp sce puits", (28, 34), (10, 16))
    if pos:
        v = _below(ws, *pos, L)
        if v:
            header["well_service_responsible"] = _clean(str(v))
            header["superintendent"] = header["well_service_responsible"]

    wb.close()

    return {
        "header": header,
        "activities": activities,
        "text_sections": {
            "plan_operations": header.get("plan_operations", ""),
        },
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
parse_daily_excel_report = parse_tp187


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: tp187_extract.py SOURCE.xlsx")
    data = parse_tp187(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))
