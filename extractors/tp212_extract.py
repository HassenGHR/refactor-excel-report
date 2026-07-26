#!/usr/bin/env python3
"""
tp212_extract.py — extract a TP-212 (rig) / HRZ-007 (well) / HRM (field)
Daily Workover Report from a PDF into the standard dict shape.

Source layout
-------------
This report family is exported straight from Excel to PDF (one page,
A4 landscape, "Creator: Microsoft Excel"), so the whole page is really
one big spreadsheet grid rendered as a single pdfplumber table
(observed: 39 rows x 15 columns). Two extraction strategies are used:

1. GRID EXTRACTION (pdfplumber `page.find_tables()`) for every labeled
   field — header info, tarification (T1-T4), cost breakdown by
   company, mud checks, chemical usage (PRODUITS), personnel, water
   volumes (GDR), progress/situation/plan text, and the vehicle /
   emetteur / recepteur line.  Cells are matched by label text and the
   value is read from the next non-empty cell in the same row, so the
   extractor tolerates the row/column shifting that's common between
   different report instances of the same template.

2. WORD-POSITION EXTRACTION (pdfplumber `page.extract_words()`) for the
   "ANALYSE DES TEMPS ET DES OPERATIONS" activities block specifically.
   This block is NOT safe to read from the grid: the time-band cell
   and the description cell are two separately-wrapped multi-line
   Excel cells that are *not* line-index-aligned (a description can
   wrap across 2+ lines for one time band, unlike the TP-186/ENF-08
   report families). It's also common for a stray label — e.g. "Le
   Directeur Eng.& Prod." — to sit physically inside the same tall
   merged cell purely because of vertical space, well below the last
   real activity. So activities are rebuilt from raw word coordinates:
     - time-band anchors are words with x0 < ~95 matching "HH:MM"
     - descriptions are words with 95 <= x0 < ~435 (the "Tarif" column
       starts around x0=440), grouped into lines by y-position and
       assigned to the nearest preceding time-band anchor
     - any description line more than ~60pt below the last time-band
       anchor is dropped (this is what filters out stray labels like
       the "Le Directeur..." signature placeholder)
     - per-activity tarif/bill codes ("T1".."T4") are read from words
       at x0 ~= 440-450, matched by y-position to each time-band anchor

Distinguishing markers (for helpers.parse_source._detect_format):
    - file extension .pdf, Creator "Microsoft Excel"
    - contains "RAPPORT JOURNALIER" + "WORK OVER" + "APPAREIL:TP 212"
      (or generally "PUITS: HRZ" / "CHAMP: HRM")
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
    """'2 849 149' / '78 434 535' / '1 227 984,00' -> int DA."""
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


def _normalize_rig(name: str) -> str:
    s = _clean(name).upper()
    m = re.search(r"TP\s*#?\s*(\d+)", s)
    if m: return f"TP-{m.group(1)}"
    return _clean(name)


def _normalize_well(name: str) -> str:
    s = _clean(name).upper()
    m = re.search(r"HRZ\s*#?\s*(\d+)", s)
    if m: return f"HRZ-{m.group(1)}"
    return _clean(name)


def _parse_size_depth(text: str):
    """'Sabot 9"5/8 : 2246 m' -> ('9"5/8', 2246.0)
    'Tête Liner 7'' : 1837 m' -> ("7''", 1837.0)
    'Sabot 7":3001 m' -> ('7"', 3001.0)
    Returns (size_str, depth_float_or_None); size_str is '' if unparseable."""
    text = _clean(text)
    m = re.search(r":\s*([\d,\.]+)\s*m\b", text, re.IGNORECASE)
    depth = _float(m.group(1)) if m else None
    before = text.split(":", 1)[0]
    size = re.sub(r"^(Sabot|T[êe]te\s*Liner)\s*", "", before, flags=re.IGNORECASE).strip()
    return size, depth


# ---------------------------------------------------------------------------
# Grid helpers (label -> value-in-same-row lookups)
# ---------------------------------------------------------------------------
def _norm_label(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def _grid_find(grid: List[List[Optional[str]]], label: str,
               occurrence: int = 0) -> Optional[tuple]:
    """Find the (row, col) of a cell whose FIRST LINE equals `label`
    (case-insensitive). `occurrence` selects the Nth match (0-based) —
    useful for repeated labels like 'GSF' appearing in two sections."""
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


def _row_next(grid: List[List[Optional[str]]], r: int, c: int,
              skip: int = 1) -> Optional[str]:
    """The `skip`-th cell to the right of (r, c) in the same row, counting
    only real cells (None entries are merged-cell continuations and are
    skipped; empty-string cells are genuine — usually blank — values and
    DO count, so a label whose value is legitimately blank returns "" not
    a bled-in value from the next field over)."""
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


def _col_block(grid: List[List[Optional[str]]], header_label: str,
               n_fields: int, value_col_offset: int = 1) -> List[str]:
    """Read `n_fields` consecutive rows below `header_label`'s row,
    returning the value found value_col_offset cells to the right of the
    header's column in each of those rows (used for PERSONNEL-style
    label/value column stacks)."""
    pos = _grid_find(grid, header_label)
    out = []
    if not pos: return out
    r0, c0 = pos
    for r in range(r0 + 1, r0 + 1 + n_fields):
        if r >= len(grid): break
        out.append(grid[r])
    return out


# ---------------------------------------------------------------------------
# Activities: word-position based parser (see module docstring)
# ---------------------------------------------------------------------------
_TIME_RE = re.compile(r"^(\d{1,2}):(\d{2})$")


def _parse_activities(page) -> List[dict]:
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)

    # 1) Time-band anchors: pairs of HH:MM tokens in the leftmost column
    time_words = [w for w in words if w["x0"] < 95 and _TIME_RE.match(w["text"])]
    by_top = {}
    for w in time_words:
        by_top.setdefault(round(w["top"], 1), []).append(w)
    anchors = []  # (top, start_time, end_time)
    for top, ws in sorted(by_top.items()):
        ws = sorted(ws, key=lambda w: w["x0"])
        if len(ws) < 2: continue
        m1, m2 = _TIME_RE.match(ws[0]["text"]), _TIME_RE.match(ws[1]["text"])
        if not (m1 and m2): continue
        st = time(int(m1.group(1)) % 24, int(m1.group(2)))
        et = time(int(m2.group(1)) % 24, int(m2.group(2)))
        anchors.append({"top": top, "start_time": st, "end_time": et, "desc_lines": []})

    if not anchors:
        return []
    anchors.sort(key=lambda a: a["top"])

    # 2) Description column words -> lines -> assign to nearest anchor above
    DESC_X0, DESC_X1 = 95, 435
    desc_words = [w for w in words if DESC_X0 <= w["x0"] < DESC_X1]
    lines_by_top = {}
    for w in desc_words:
        lines_by_top.setdefault(round(w["top"], 1), []).append(w)

    first_top = anchors[0]["top"]
    last_top = anchors[-1]["top"]
    cutoff = last_top + 60  # drop stray labels (e.g. signature placeholders)
                            # sitting well below the last real activity line
    floor = first_top - 2   # drop page-header / title text sitting above
                            # the first real activity line
    for top in sorted(lines_by_top):
        if top < floor or top > cutoff:
            continue
        # find the last anchor whose top <= this line's top (+tolerance)
        candidate = None
        for a in anchors:
            if a["top"] <= top + 2:
                candidate = a
            else:
                break
        if candidate is None:
            continue  # shouldn't happen given the floor check above
        ws = sorted(lines_by_top[top], key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in ws)
        candidate["desc_lines"].append(text)

    # 3) Tarif/bill codes: words at x0 ~ 435-460, matched by anchor top
    tarif_words = [w for w in words if 435 <= w["x0"] < 465]
    tarif_by_top = {round(w["top"], 1): w["text"] for w in tarif_words}
    for a in anchors:
        code = ""
        for top, txt in tarif_by_top.items():
            if abs(top - a["top"]) <= 3:
                code = txt
                break
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
def parse_tp212(source: Union[Path, str]) -> dict:
    """Parse a TP-212 / HRZ-007 (PDF) daily workover report."""
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

        # Date + day number — top-right cell "Journée du\n24/07/2026" + "14"
        pos = _grid_find(grid, "Journée du") or _grid_find(grid, "Journee du")
        if not pos:
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

        # PUITS / CHAMP / APPAREIL
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

        # Well construction reference points — raw text kept as-is (matches
        # what's printed on the report) ...
        for label, key in [
            ("Sabot 9\"5/8", "sabot_9_5_8"), ("Sabot 9\"5/8 :", "sabot_9_5_8"),
            ("Tête Liner 7''", "tete_liner_7"),
            ("Sabot 7\"", "sabot_7"),
        ]:
            for r, row in enumerate(grid):
                for c, cell in enumerate(row):
                    if cell and _norm_label(cell.split(":")[0]) == _norm_label(label.split(":")[0]):
                        header[key] = _clean(cell)
                        break

        # ... plus standardized "last_csg_shoe" / "last_lnr_top" /
        # "last_lnr_shoe" fields (size@depth, matching the naming/format
        # used by the enf17 DDR extractor) derived from those same cells:
        #   Sabot 9"5/8 : 2246 m   -> last casing shoe
        #   Tête Liner 7'' : 1837 m -> liner top
        #   Sabot 7":3001 m         -> liner shoe
        if header.get("sabot_9_5_8"):
            size, depth = _parse_size_depth(header["sabot_9_5_8"])
            if size:
                header["last_csg_size"] = size
                if depth is not None:
                    header["last_csg_depth"] = depth
                    header["last_csg_shoe"] = f"{size} @ {depth:g}m"
        if header.get("tete_liner_7"):
            size, depth = _parse_size_depth(header["tete_liner_7"])
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

        for r, row in enumerate(grid):
            for c, cell in enumerate(row):
                if cell and cell.upper().startswith("KOP"):
                    header["kop_toc"] = _clean(cell)

        # BOUE type
        for r, row in enumerate(grid):
            for c, cell in enumerate(row):
                if cell and cell.upper().startswith("TYPE:"):
                    header["mud_type"] = _clean(cell.split(":", 1)[1])

        # BUT DU WORK OVER (purpose) — the whole heading + body live in ONE
        # multi-line grid cell (not spread across subsequent grid rows), so
        # just strip the "BUT DU WORK OVER" heading line from that cell.
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

        # Vehicule / emetteur / recepteur
        pos = _grid_find(grid, "VEHICULE")
        if pos:
            r, c = pos
            vals = [x for x in grid[r][c + 1:] if x]
            if vals: header["vehicule"] = _clean(vals[0])
        emet = _grid_find(grid, "Emetteur (Chantier)")
        if emet:
            r, c = emet
            if r + 1 < len(grid):
                header["emetteur"] = _clean(grid[r + 1][c]) if grid[r + 1][c] else ""
                # recepteur sits a few cols right in the same row
                for cc in range(c + 1, len(grid[r + 1])):
                    if grid[r + 1][cc]:
                        header["recepteur"] = _clean(grid[r + 1][cc])
                        break
        # This report family has no dedicated shift-supervisor field (unlike
        # e.g. enf17's TP Sénior/TP Junior), so the Emetteur (Chantier) /
        # Récepteur (Service) sign-off names are mapped onto the standard
        # supervisor/superintendent slots.
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
                    # Store even zero hours — a rig genuinely idle on T3/T4
                    # for the day is real data, not a missing field.
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
                if label is None:
                    label = grid[r][c0] if c0 < len(grid[r]) else None
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
        # MUD CHECKS
        # ================================================================
        mud_checks = {}
        for label, key, is_float in [
            ("Transfert OBM", "transfert_obm", False),
            ("Tripping", "tripping", True),
            ("Dumping", "dumping", True),
            ("Shaker+surface", "shaker_surface", True),
            ("Densité sg", "density", True),
            ("Visc March", "visc_march", True),
            ("YP", "yp", True),
            ("VP", "vp", True),
            ("E Stability", "e_stability", True),
            ("Solides", "solides", True),
            ("LGS", "lgs", True),
            ("Filtrat HPHT", "filtrat_hpht", True),
            ("Oil / Water ratio", "oil_water_ratio", False),
            ("SALINITY gr/l", "salinity", True),
            ("V Surf OBM d = 1.05", "v_surf_obm_1_05", True),
            ("V.Surf OBM d= 1.54", "v_surf_obm_1_54", True),
            ("V.Surf OBM", "v_surf_obm", True),
            ("V,Puit OBM", "v_puit_obm", True),
        ]:
            v = _label_value(grid, label)
            if v:
                mud_checks[key] = _float(v) if is_float else v

        # ================================================================
        # PRODUITS (mud chemical usage)
        # ================================================================
        chemicals = []
        # The plain section title "PRODUITS" (a few rows above) and the
        # actual table header cell "Produits" normalize to the same string
        # case-insensitively, so match on the ROW that has both "Produits"
        # AND "Unité" cells together — that's the real header row.
        prod_header = None
        for r, row in enumerate(grid):
            norm_cells = [_norm_label(c) for c in row if c]
            if "produits" in norm_cells and "unité" in norm_cells:
                c0 = next(c for c, cell in enumerate(row)
                          if cell and _norm_label(cell) == "produits")
                prod_header = (r, c0)
                break
        # Bound the item list so it can't wander into the next section
        # (GDR / water volumes) if a blank row is encountered.
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
        personnel_data = []  # enf17-style list: [{"company","number","hours","names"}]
        pos = _grid_find(grid, "PERSONNEL")
        if pos:
            r0, c0 = pos
            for r in range(r0 + 1, len(grid)):
                label = grid[r][c0]
                if not label: continue
                lc = _clean(label)
                if lc.upper() in ("GDR",):
                    # GDR appears as the last personnel-style row before the
                    # water section; still record it (count may be blank)
                    val = _row_next(grid, r, c0)
                    n = _int(val) if val else 0
                    personnel_data.append({"company": lc, "number": n, "hours": "", "names": ""})
                    continue
                if lc.upper() == "TOTAL":
                    val = _row_next(grid, r, c0)
                    if val: header["personnel_total"] = _int(val)
                    break
                val = _row_next(grid, r, c0)
                if val is not None:
                    n = _int(val)
                    personnel_data.append({"company": lc, "number": n, "hours": "", "names": ""})

        # ================================================================
        # WATER / GDR VOLUMES
        # ================================================================
        water = {}
        for label, key in [("Eau-Rig", "eau_rig"),
                            ("Eau-Dinc-SH+ENTP", "eau_dinc_sh_entp"),
                            ("Fluide de compt", "fluide_de_compt")]:
            pos = _grid_find(grid, label)
            if pos:
                r, c = pos
                unit = _row_next(grid, r, c, skip=1) or ""
                used = _row_next(grid, r, c, skip=2) or ""
                stock = _row_next(grid, r, c, skip=3) or ""
                water[key] = {"unit": _clean(unit), "used": _clean(used), "stock": _clean(stock)}
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
parse_daily_pdf_report = parse_tp212


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: tp212_extract.py SOURCE.pdf")
    data = parse_tp212(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))