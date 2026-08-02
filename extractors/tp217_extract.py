#!/usr/bin/env python3
"""
tp217_extract.py — extract a TP-217 (rig) / ONRS-01 (well) / HRM (field)
Daily Workover Report from a PDF into the standard dict shape.

Source layout
-------------
Same Excel-exported-to-PDF template family as tp212_extract.py /
enf03_extract.py (one page, A4 landscape, one big pdfplumber grid
table). This report instance has two real quirks not seen in the
earlier two reports:

  - The mud-checks panel's FIRST label/value row sits on the very same
    grid row as the "TARIFICATION" section header (one row earlier
    than in the TP-212/ENF-03 reports, where the panel starts the row
    AFTER "TARIFICATION"). So the mud-checks row range here starts
    at the "TARIFICATION" row itself (inclusive) rather than the next
    row — any row with no mud data in those columns is simply skipped,
    so this is safe for either layout.

  - Two of that panel's label/value pairs are corrupted by the Excel
    export splitting a decimal point or a closing parenthesis across
    the cell boundary, e.g. "Réception OBM d=0." / "9 0" instead of
    "Réception OBM d=0.9" / "0", and "V. S Bacs SH (gas-oil" / ") 27"
    instead of "V. S Bacs SH (gas-oil)" / "27". `_repair_mud_pair()`
    detects an unterminated decimal point or open parenthesis in the
    label and reassembles the true label/value split from the
    following cell — this matters because without it a real value
    (the 27 m3 gas-oil stock figure) would come out as the garbled
    string ") 27" instead of the number 27.

  - Well/rig names use yet another prefix ("ONRS 01", not "HR"/"HRZ"),
    so naming is normalized generically for ANY alpha-prefix + number
    pattern rather than hardcoding "HR"/"ENF"/"TP".

As with the other extractors in this family, activities in "ANALYSE
DES TEMPS ET DES OPERATIONS" are rebuilt from raw word coordinates
(not the grid) because the time-band and description cells aren't
line-index-aligned, and the Tarif column / bill-code positions are
detected dynamically from word coordinates since Excel column widths
shift between report instances.

Distinguishing markers (for helpers.parse_source._detect_format):
    - file extension .pdf, Creator "Microsoft Excel" / "Microsoft® Excel®"
    - contains "RAPPORT JOURNALIER" + "WORK OVER" + "APPAREIL:TP 217"
      (or generally "PUITS: ONRS" + "CHAMP: HRM")
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from pathlib import Path
from typing import Union, List, Optional

import pdfplumber


# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------
def _clean(v) -> str:
    if v is None: return ""
    s = str(v).replace("\u202f", " ").replace("\u00a0", " ")
    return re.sub(r"[ \t]+", " ", s).strip()


def _float(v, default=0.0) -> float:
    if v is None: return default
    if isinstance(v, (int, float)): return float(v)
    s = _clean(v).replace(" ", "").replace(",", ".")
    if s in ("", "-", "/", "None"): return default
    m = re.match(r"^([-+]?\d+(?:\.\d+)?)", s)
    if m:
        try: return float(m.group(1))
        except ValueError: pass
    return default


def _int(v, default=0) -> int:
    return int(_float(v, float(default)))


def _money_to_int(s) -> int:
    """'2 802 818' / '243 230 632' / '6 427 056,00' -> int DA."""
    if s is None: return 0
    if isinstance(s, (int, float)): return int(s)
    cleaned = _clean(s)
    cleaned = re.sub(r"\s*DA\s*$", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.split(r"[.,]\d{2,3}\s*$", cleaned)[0]
    digits = re.sub(r"\D", "", cleaned)
    return int(digits) if digits else 0


def _date_from_text(s: str) -> Optional[date_type]:
    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if not m: return None
    try:
        return date_type(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def _hours_between(start_t: time, end_t: time) -> float:
    sm = start_t.hour * 60 + start_t.minute
    em = end_t.hour * 60 + end_t.minute
    if em == sm: return 24.0
    if em > sm:  return (em - sm) / 60.0
    return (em + 1440 - sm) / 60.0


def _normalize_alpha_num(name: str) -> str:
    """'ONRS 01' / 'TP 217' / 'HR#152' / 'ENF-03' -> 'ONRS-01' / 'TP-217' /
    'HR-152' / 'ENF-03'. Generic across whatever alpha prefix the rig/well
    naming convention uses — this report family alone has used TP, ENF,
    HR, HRZ and ONRS, so the prefix itself is not hardcoded."""
    s = _clean(name).upper()
    m = re.match(r"^([A-Z]+)\s*#?\s*-?\s*(\d+)$", s)
    if m: return f"{m.group(1)}-{m.group(2)}"
    return _clean(name)


def _normalize_rig(name: str) -> str:
    return _normalize_alpha_num(name)


def _normalize_well(name: str) -> str:
    return _normalize_alpha_num(name)


def _parse_size_depth(text: str):
    """Generic 'LABEL WORDS <size> : <depth> m' parser that works
    regardless of the exact label wording (e.g. "Sabot 9''5/8 : 2187 m",
    "Top of Liner 7'' : 2147 m", "Bridge Plug 7'' : 2219 m") by taking
    the substring from the first digit onward as the size, and pulling
    the depth from whatever number precedes a trailing 'm'.
    Returns (size_str, depth_float_or_None); size_str is '' if unparseable.
    """
    text = _clean(text)
    m = re.search(r":\s*([\d,\.]+)\s*m\b", text, re.IGNORECASE)
    depth = _float(m.group(1)) if m else None
    before = text.split(":", 1)[0]
    m2 = re.search(r"\d.*$", before)
    size = m2.group(0).strip() if m2 else ""
    return size, depth


def _slug(label: str) -> str:
    s = _clean(label).lower()
    s = re.sub(r"[^\w]+", "_", s)
    return re.sub(r"_+", "_", s).strip("_")


def _repair_mud_pair(label: str, value: str):
    """Fix a specific Excel-export artifact seen in this report family:
    a long label ending mid-token at the cell boundary bleeds its
    remainder into the START of the value cell, e.g.
        label='Réception OBM d=0.'  value='9 0'
        -> label='Réception OBM d=0.9', value='0'
        label='V. S Bacs SH (gas-oil'  value=') 27'
        -> label='V. S Bacs SH (gas-oil)', value='27'
    Handles: an unterminated decimal point ("...d=0.") pulling a leading
    digit run off the value cell, and an unterminated open-parenthesis
    pulling a leading ")" off the value cell. Falls back to discarding
    a stray leftover token in an otherwise-plain value (e.g. "0 0" -> "0")
    since these panels only ever hold single short tokens.
    """
    if not label or value is None:
        return label, value
    value = value.strip()
    open_parens = label.count("(") - label.count(")")
    ends_incomplete_decimal = bool(re.search(r"\d\.$", label.rstrip()))
    if open_parens > 0 and value.startswith(")"):
        label = label + ")"
        value = value[1:].strip()
    elif ends_incomplete_decimal:
        m = re.match(r"^(\d+)\s*(.*)$", value)
        if m:
            label = label.rstrip() + m.group(1)
            value = m.group(2).strip()
    else:
        m2 = re.match(r"^(\S+)\s+\S+$", value)
        if m2:
            value = m2.group(1)
    return label, value


# Known mud-check label -> clear field name aliases. Anything not listed
# here falls back to a slugified version of the label (see _slug), so no
# field is silently dropped if a report variant introduces a new one.
_MUD_LABEL_ALIASES = {
    "transfert": "transfert",
    "réception": "reception",
    "reception": "reception",
    "dumping": "dumping",
    "coating": "coating",
    "tripping": "tripping",
    "perte surface": "surface_loss",
    "v.surface": "v_surface",
    "v surface": "v_surface",
    "v.s obm": "v_s_obm",
    "densité": "density",
    "densite": "density",
    "v. puits": "v_puits",
    "v puits": "v_puits",
    "visc march": "visc_march",
    "pv": "pv",
    "yp": "yp",
    "vp": "vp",
    "cake thickns": "cake_thickness",
    "%e": "pct_e",
    "%s": "pct_s",
    "%h": "pct_h",
    "hp/ht @ 200°f": "hp_ht_200f",
    "hp ht 200 f": "hp_ht_200f",
    "elect stab": "electrical_stability",
    "e stability": "electrical_stability",
    "rapport h/e": "oil_water_ratio",   # Huile/Eau = Oil/Water ratio
    "oil / water ratio": "oil_water_ratio",
    "salinity gr/l": "salinity",
    "lgs": "lgs",
    "solides": "solides",
    "filtrat hpht": "filtrat_hpht",
}


# ---------------------------------------------------------------------------
# Grid helpers (label -> value-in-same-row lookups)
# ---------------------------------------------------------------------------
def _norm_label(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def _grid_find(grid: List[List[Optional[str]]], label: str,
               occurrence: int = 0) -> Optional[tuple]:
    """(row, col) of the cell whose first line (or whole text) equals
    `label`, case-insensitively. `occurrence` selects the Nth match."""
    target = _norm_label(label)
    hits = []
    for r, row in enumerate(grid):
        for c, cell in enumerate(row):
            if cell is None: continue
            first_line = _norm_label(cell.split("\n")[0])
            if first_line == target or _norm_label(cell) == target:
                hits.append((r, c))
    if occurrence < len(hits):
        return hits[occurrence]
    return None


def _grid_find_prefix(grid: List[List[Optional[str]]], prefixes: List[str]
                       ) -> Optional[tuple]:
    """(row, col) of the first cell whose text starts with any of the
    given prefixes (case-insensitive) — used for header labels whose
    exact wording varies between report instances."""
    for r, row in enumerate(grid):
        for c, cell in enumerate(row):
            if not cell: continue
            norm = _norm_label(cell)
            for p in prefixes:
                if norm.startswith(_norm_label(p)):
                    return (r, c)
    return None


def _row_next(grid: List[List[Optional[str]]], r: int, c: int,
              skip: int = 1) -> Optional[str]:
    """The `skip`-th cell to the right of (r, c) in the same row, counting
    only real cells (None = merged-cell continuation, skipped; ""
    = a genuinely blank value, and DOES count, so a label with no data
    returns "" rather than bleeding into the next field's text)."""
    row = grid[r]
    count = 0
    for cc in range(c + 1, len(row)):
        if row[cc] is not None:
            count += 1
            if count >= skip:
                return row[cc]
    return None


def _label_value(grid, label: str, occurrence: int = 0) -> str:
    pos = _grid_find(grid, label, occurrence)
    if not pos: return ""
    v = _row_next(grid, *pos)
    return _clean(v) if v else ""


# ---------------------------------------------------------------------------
# Activities: word-position based parser
# ---------------------------------------------------------------------------
_TIME_RE = re.compile(r"^(\d{1,2}):(\d{2})$")


def _norm_time_token(s: str) -> str:
    """Strip a stray leading/trailing colon typo sometimes present in the
    source report (e.g. ':07:00' instead of '07:00') so the time band
    still parses instead of the whole activity row being silently
    dropped for lack of a matching end-time token."""
    return s.strip().strip(":")


def _parse_activities(page) -> List[dict]:
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)

    # 1) Time-band anchors: pairs of HH:MM tokens in the leftmost column
    time_words = [w for w in words if w["x0"] < 95 and _TIME_RE.match(_norm_time_token(w["text"]))]
    by_top = {}
    for w in time_words:
        by_top.setdefault(round(w["top"], 1), []).append(w)
    anchors = []
    for top, ws in sorted(by_top.items()):
        ws = sorted(ws, key=lambda w: w["x0"])
        if len(ws) < 2: continue
        m1 = _TIME_RE.match(_norm_time_token(ws[0]["text"]))
        m2 = _TIME_RE.match(_norm_time_token(ws[1]["text"]))
        if not (m1 and m2): continue
        st = time(int(m1.group(1)) % 24, int(m1.group(2)))
        et = time(int(m2.group(1)) % 24, int(m2.group(2)))
        anchors.append({"top": top, "start_time": st, "end_time": et, "desc_lines": []})
    if not anchors:
        return []
    anchors.sort(key=lambda a: a["top"])

    # 2) Description column: right edge detected dynamically from the
    # "Tarif" column header word (Excel column widths shift between
    # report instances, so this can't be a fixed pixel cutoff).
    tarif_header_words = [w for w in words if w["text"] == "Tarif"]
    desc_x1 = (min(w["x0"] for w in tarif_header_words) - 10) if tarif_header_words else 435
    DESC_X0 = 95
    desc_words = [w for w in words if DESC_X0 <= w["x0"] < desc_x1]
    lines_by_top = {}
    for w in desc_words:
        lines_by_top.setdefault(round(w["top"], 1), []).append(w)

    first_top = anchors[0]["top"]
    last_top = anchors[-1]["top"]
    # The stray "Le Directeur Eng.& Prod." signature sits somewhere below
    # the last real activity line, within the same tall merged cell. A
    # fixed "+60pt" margin isn't always enough — reports with fewer/shorter
    # activities push it visually closer to the last real line — so look
    # for the actual signature text first and cut exactly above it;
    # fall back to the margin heuristic only if it's not found.
    directeur_tops = [w["top"] for w in words
                      if "directeur" in w["text"].lower() and DESC_X0 <= w["x0"] < desc_x1]
    cutoff = (min(directeur_tops) - 2) if directeur_tops else (last_top + 60)
    floor = first_top - 2   # drop page-header/title text above the block
    for top in sorted(lines_by_top):
        if top < floor or top > cutoff:
            continue
        candidate = None
        for a in anchors:
            if a["top"] <= top + 2:
                candidate = a
            else:
                break
        if candidate is None:
            continue
        ws = sorted(lines_by_top[top], key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in ws)
        candidate["desc_lines"].append(text)

    # 3) Per-activity tarif/bill codes: matched by TOP proximity to each
    # anchor rather than a fixed x-range, since the bill-code column's
    # x-position also shifts between report instances.
    code_words = [w for w in words if re.match(r"^T[1-4]$", w["text"])]
    for a in anchors:
        code = ""
        best_dist = None
        for w in code_words:
            d = abs(round(w["top"], 1) - a["top"])
            if d <= 3 and (best_dist is None or d < best_dist):
                code, best_dist = w["text"], d
        a["bill"] = code

    activities = []
    for a in anchors:
        hours = _hours_between(a["start_time"], a["end_time"])
        activities.append({
            "start_time": a["start_time"],
            "end_time": a["end_time"],
            "hours": hours,
            "phase_name": "",
            "code": a.get("bill", ""), "sub": "",
            "description": _clean(" ".join(a["desc_lines"])),
            "start_md": 0, "end_md": 0,
            "npt": 0, "npt_detail": "",
            "npt_company": "", "op_company": "",
            "bill": a.get("bill", ""),
        })
    return activities


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_tp217(source: Union[Path, str]) -> dict:
    """Parse a TP-217 / ONRS-01 (PDF) daily workover report."""
    source_path = Path(source)
    with pdfplumber.open(str(source_path)) as pdf:
        page = pdf.pages[0]
        tables = page.find_tables()
        if not tables:
            raise ValueError("No table grid found on page 1 — unexpected layout")
        grid = tables[0].extract()

        # ================================================================
        # HEADER
        # ================================================================
        header = {}

        pos = None
        for r, row in enumerate(grid):
            for c, cell in enumerate(row):
                if cell and "journée du" in cell.lower():
                    pos = (r, c)
        if pos:
            r, c = pos
            d = _date_from_text(grid[r][c])
            if d: header["date"] = d
            day_val = _row_next(grid, r, c)
            if day_val and re.match(r"^\d+$", _clean(day_val)):
                header["day_number"] = int(_clean(day_val))

        for r, row in enumerate(grid):
            for c, cell in enumerate(row):
                if not cell: continue
                m = re.search(r"PUITS\s*:\s*([A-Z0-9\-\s]+?)\s+CHAMP\s*:\s*([A-Z0-9\-\s]+)", cell, re.IGNORECASE)
                if m:
                    header["well_name"] = _normalize_well(m.group(1))
                    header["field_name"] = _clean(m.group(2))
                m2 = re.search(r"APPAREIL\s*:\s*([A-Z0-9\-\s]+)", cell, re.IGNORECASE)
                if m2:
                    header["rig_name"] = _normalize_rig(m2.group(1))

        # Well construction reference points — matched by label PREFIX
        # (wording varies: "Sabot 9''5/8", "Top of Liner 7''",
        # "Tête Liner 7''", "Sabot 7''", "Bridge Plug 7''", "KOP...").
        for prefixes, raw_key in [
            (["Sabot 9"], "sabot_9_5_8"),
            (["Sabot 7"], "sabot_7"),
            (["Top of Liner", "Tête Liner", "Tete Liner"], "top_of_liner"),
            (["Bridge Plug"], "bridge_plug"),
        ]:
            pos = _grid_find_prefix(grid, prefixes)
            if pos:
                header[raw_key] = _clean(grid[pos[0]][pos[1]])
        for r, row in enumerate(grid):
            for c, cell in enumerate(row):
                if cell and cell.upper().startswith("KOP"):
                    header["kop_toc"] = _clean(cell)

        # Standardized last_csg_shoe / last_lnr_top / last_lnr_shoe
        # (size@depth, matching the enf17 DDR extractor's naming/format).
        if header.get("sabot_9_5_8"):
            size, depth = _parse_size_depth(header["sabot_9_5_8"])
            if size:
                header["last_csg_size"] = size
                if depth is not None:
                    header["last_csg_depth"] = depth
                    header["last_csg_shoe"] = f"{size} @ {depth:g}m"
        if header.get("top_of_liner"):
            size, depth = _parse_size_depth(header["top_of_liner"])
            if size:
                header["last_lnr_top_size"] = size
                if depth is not None:
                    header["last_lnr_top_depth"] = depth
                    header["last_lnr_top"] = f"{size} @ {depth:g}m"
        if header.get("sabot_7"):
            size, depth = _parse_size_depth(header["sabot_7"])
            if size:
                header["last_lnr_shoe_size"] = size
                if depth is not None:
                    header["last_lnr_shoe_depth"] = depth
                    header["last_lnr_shoe"] = f"{size} @ {depth:g}m"
        if header.get("bridge_plug"):
            size, depth = _parse_size_depth(header["bridge_plug"])
            if size:
                header["bridge_plug_size"] = size
                if depth is not None:
                    header["bridge_plug_depth"] = depth

        # BOUE type
        for r, row in enumerate(grid):
            for c, cell in enumerate(row):
                if cell and cell.upper().startswith("TYPE:"):
                    header["mud_type"] = _clean(cell.split(":", 1)[1])

        # BUT DU WORK OVER (purpose) — whole heading + body live in ONE
        # multi-line grid cell; strip the heading line itself.
        pos = _grid_find(grid, "BUT DU WORK OVER")
        if pos:
            r0, c0 = pos
            lines = (grid[r0][c0] or "").split("\n")
            body = " ".join(_clean(l) for l in lines[1:] if _clean(l))
            if body:
                header["workover_purpose"] = body

        # Progress / situation / plan
        avanc = _label_value(grid, "AVANCEMENT PHYSIQUE :")
        if avanc: header["progress_pct"] = _float(avanc.replace("%", ""))
        sit = _label_value(grid, "SITUATION AU RAPPORT:")
        if sit: header["situation"] = sit
        plan = _label_value(grid, "PROGRAMME PREVU:")
        if plan: header["plan_operations"] = plan

        # Vehicule
        pos = _grid_find(grid, "VEHICULE")
        if pos:
            r, c = pos
            vals = [x for x in grid[r][c + 1:] if x]
            if vals: header["vehicule"] = _clean(vals[0])

        # Emetteur (Chantier) / Récepteur (Service) — this report family
        # has no dedicated shift-supervisor field, so the sign-off names
        # are mapped onto the standard supervisor/superintendent slots.
        emet = _grid_find(grid, "Emetteur (Chantier)")
        if emet:
            r, c = emet
            if r + 1 < len(grid):
                header["emetteur"] = _clean(grid[r + 1][c]) if grid[r + 1][c] else ""
                for cc in range(c + 1, len(grid[r + 1])):
                    if grid[r + 1][cc]:
                        header["recepteur"] = _clean(grid[r + 1][cc])
                        break
        if header.get("emetteur"):
            header["supervisor"] = header["emetteur"]
        if header.get("recepteur"):
            header["superintendent"] = header["recepteur"]

        # ================================================================
        # TARIFICATION (T1-T4 hours/jour and cumulative)
        # ================================================================
        tarif_totals = {}
        tarif_cumul = {}
        pos = _grid_find(grid, "H/jour")
        if pos:
            r, c = pos
            vals = [v for v in grid[r][c + 1:] if v is not None]
            for i, key in enumerate(["t1", "t2", "t3", "t4"]):
                if i < len(vals):
                    tarif_totals[key] = _float(vals[i])
        pos = _grid_find(grid, "H. Cumul")
        if pos:
            r, c = pos
            vals = [v for v in grid[r][c + 1:] if v is not None]
            for i, key in enumerate(["t1", "t2", "t3", "t4"]):
                if i < len(vals):
                    tarif_cumul[key] = _float(vals[i])

        # ================================================================
        # COSTS: daily / cumulative + breakdown by company
        # ================================================================
        pos = _grid_find(grid, "Journalier")
        if pos:
            r, c = pos
            if r + 1 < len(grid):
                daily_val = grid[r + 1][c]
                if daily_val: header["daily_cost"] = _money_to_int(daily_val)
                cumul_pos = _grid_find(grid, "Cumulé") or _grid_find(grid, "Cumule")
                if cumul_pos:
                    rc, cc = cumul_pos
                    if r + 1 < len(grid) and cc < len(grid[r + 1]):
                        cumul_val = grid[r + 1][cc]
                        if cumul_val: header["cum_cost"] = _money_to_int(cumul_val)

        cost_breakdown = {}
        pos = _grid_find(grid, "DETAILS")
        if pos:
            r0, c0 = pos
            for r in range(r0 + 1, len(grid)):
                label = grid[r][c0]
                if not label: continue
                lab_clean = _clean(label)
                if lab_clean.upper() in ("PERSONNEL", "PRODUITS", "GDR"):
                    break
                val = _row_next(grid, r, c0)
                if val:
                    amt = _money_to_int(val)
                    if amt or _clean(val) in ("0,00", "0"):
                        cost_breakdown[lab_clean] = amt
        if cost_breakdown:
            header["cost_breakdown"] = cost_breakdown

        # ================================================================
        # MUD CHECKS — read generically from the row range between the
        # "TARIFICATION" header and the "PRODUITS" section title, rather
        # than a hardcoded per-report label list (see module docstring).
        # ================================================================
        mud_checks = {}
        tarification_pos = _grid_find(grid, "TARIFICATION")
        produits_title_pos = None
        for r, row in enumerate(grid):
            for c, cell in enumerate(row):
                if cell and _norm_label(cell) == "produits" and c >= 10:
                    produits_title_pos = (r, c)
                    break
            if produits_title_pos: break

        if tarification_pos and produits_title_pos:
            # This report's mud panel starts on the SAME row as the
            # "TARIFICATION" header (unlike TP-212/ENF-03, where it starts
            # the row after) — start inclusive; blank rows are just skipped.
            r0 = tarification_pos[0]
            r1 = produits_title_pos[0]
            mud_col = tarification_pos[1] + 5  # right-panel column (label1)
            for r in range(r0, r1):
                if mud_col + 3 >= len(grid[r]): continue
                for lab_c, val_c in ((mud_col, mud_col + 1), (mud_col + 2, mud_col + 3)):
                    label = grid[r][lab_c]
                    if not label: continue
                    val = grid[r][val_c]
                    if val is None: continue
                    label, val = _repair_mud_pair(label, val)
                    norm = _norm_label(label)
                    key = _MUD_LABEL_ALIASES.get(norm, _slug(label))
                    if not key or key in mud_checks: continue
                    if _clean(val) == "": continue
                    fv = _float(val)
                    # Keep genuinely non-numeric values (e.g. "95/05") as text
                    mud_checks[key] = fv if re.match(r"^[\d.,\-]+$", _clean(val)) else _clean(val)

        # ================================================================
        # PRODUITS (mud chemical usage) — header row identified by having
        # BOTH "Produits" and "Unité" cells together (distinguishes it from
        # the plain "PRODUITS" section-title cell a few rows above).
        # ================================================================
        chemicals = []
        prod_header = None
        for r, row in enumerate(grid):
            norm_cells = [_norm_label(c) for c in row if c]
            if "produits" in norm_cells and "unité" in norm_cells:
                c0 = next(c for c, cell in enumerate(row)
                          if cell and _norm_label(cell) == "produits")
                prod_header = (r, c0)
                break
        ceiling = len(grid)
        water_pos = _grid_find(grid, "Eau-Rig")
        if water_pos:
            ceiling = water_pos[0]
        if prod_header:
            r0, c0 = prod_header
            for r in range(r0 + 1, ceiling):
                item = grid[r][c0]
                if not item: continue
                unit = _row_next(grid, r, c0, skip=1) or ""
                used = _row_next(grid, r, c0, skip=2) or ""
                stock = _row_next(grid, r, c0, skip=3) or ""
                chemicals.append({
                    "item": _clean(item),
                    "units": _clean(unit),
                    "received": "",
                    "used": _clean(used),
                    "on_loc": _clean(stock),
                })

        # ================================================================
        # PERSONNEL
        # ================================================================
        personnel_data = []
        pos = _grid_find(grid, "PERSONNEL")
        if pos:
            r0, c0 = pos
            for r in range(r0 + 1, len(grid)):
                label = grid[r][c0]
                if not label: continue
                lc = _clean(label)
                if lc.upper() == "TOTAL":
                    val = _row_next(grid, r, c0)
                    if val: header["personnel_total"] = _int(val)
                    break
                val = _row_next(grid, r, c0)
                if val is not None:
                    personnel_data.append({"company": lc, "number": _int(val), "hours": "", "names": ""})

        # ================================================================
        # WATER / GDR VOLUMES — read generically (label/unit/used/stock)
        # from "Eau-Rig" down to the "Emetteur (Chantier)" row, since the
        # number of rows here varies between report instances (this one
        # adds "Gasoil").
        # ================================================================
        water = {}
        if water_pos:
            wr0, wc0 = water_pos
            end_pos = _grid_find(grid, "Emetteur (Chantier)")
            wceiling = end_pos[0] if end_pos else len(grid)
            for r in range(wr0, wceiling):
                if wc0 >= len(grid[r]): continue
                label = grid[r][wc0]
                if not label: continue
                unit = _row_next(grid, r, wc0, skip=1) or ""
                used = _row_next(grid, r, wc0, skip=2) or ""
                stock = _row_next(grid, r, wc0, skip=3) or ""
                water[_slug(label)] = {"unit": _clean(unit), "used": _clean(used), "stock": _clean(stock)}
        if water:
            header["water"] = water

        # ================================================================
        # ACTIVITIES (word-position parser)
        # ================================================================
        activities = _parse_activities(page)

    return {
        "header": header,
        "activities": activities,
        "text_sections": {
            "current_operation": header.get("situation", ""),
            "plan_operations": header.get("plan_operations", ""),
        },
        "mud_checks": mud_checks,
        "mud_volume": {},
        "mud_chemical_usage": chemicals,
        "personnel_data": personnel_data,
        "pumps": [],
        "well_location": {},
        "survey_data": [],
        "safety": {},
        "tarif_totals": tarif_totals,
        "tarif_cumul": tarif_cumul,
        "service_companies": [],
    }


# Drop-in compat
parse_daily_pdf_report = parse_tp217


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: tp217_extract.py SOURCE.pdf")
    data = parse_tp217(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))