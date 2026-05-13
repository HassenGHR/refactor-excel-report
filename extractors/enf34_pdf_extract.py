#!/usr/bin/env python3
"""
enf34_pdf_extract.py — extract an ENF#34 (ENAFOR rig 34, Direction
Régionale Gassi-Touil, GT-XX wells) Daily Workover Report from a PDF
source into the standard dict shape used by all the Excel extractors.

Source layout
-------------
4-page PDF, French.  Distinguishing markers:
    - Page 2 title: "RAPPORT JOURNALIER DE WORK OVER" (with spaces)
    - Header "Direction Régionale Gassi-Touil"
    - Page 2: header block (Activité du / Puits / Appareil / Date début / BUT)
    - Page 2: "OPERATIONS REALISEES" with time-banded rows
    - Page 2: "After M/night", "A / E :" current op, "Programme Prévu"
    - Page 3: PUITS / Appareil / CHAMP single line, "DERNIER TUBAGE",
      "BILAN BOUE" block with ~17 mud properties, "PRODUIT & ETATS DES
      STOCKS" chemical table
    - Page 4: "ANALYSE DES COUTS" cost table with T1/T2/T3/T4, service
      breakdown, "SUPERVISEUR SH DP/GTL : NAME" at the bottom

This module is selected by parse_source.py when the uploaded file is a
PDF whose text contains "GASSI-TOUIL" / "GASSI TOUIL" + "RAPPORT
JOURNALIER DE WORK OVER".
"""
from __future__ import annotations
import re
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Union, List

import pdfplumber


# ---------------------------------------------------------------------------
# Helpers (same shape as the Excel extractors)
# ---------------------------------------------------------------------------
def _clean(v) -> str:
    if v is None: return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _float(v, default=0.0) -> float:
    if v is None: return default
    if isinstance(v, (int, float)): return float(v)
    s = _clean(v).replace(",", ".").replace(" ", "")
    if s in ("", "-", "/", "None", "---"): return default
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
    if not s or s in ("-", "/", "//", "---"): return None
    if "--/" in s: return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: continue
    return None


def _time_parse(v):
    if v is None: return None
    if isinstance(v, time): return v
    s = _clean(v)
    if not s or s in ("-", "None"): return None
    if s in ("24:00", "24:00:00"): return time(0, 0)
    for fmt in ("%H:%M:%S", "%H:%M", "%Hh%M"):
        try: return datetime.strptime(s, fmt).time()
        except ValueError: continue
    return None


def _normalize_rig(name: str) -> str:
    """'ENF-34' or 'ENAFOR # 34' -> 'ENF#34'."""
    s = _clean(name).upper()
    m = re.search(r"(?:ENAFOR|ENF)\s*[#\-]?\s*(\d+)", s)
    if m:
        return f"ENF#{int(m.group(1)):02d}"
    return _clean(name)


def _money_to_int(s) -> int:
    """'2 503 636' or '2503636' or '0' → 2503636."""
    if s is None: return 0
    if isinstance(s, (int, float)): return int(s)
    cleaned = re.sub(r"[^\d]", "", str(s))
    return int(cleaned) if cleaned else 0


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_enf34_pdf(source: Union[Path, str, BytesIO]) -> dict:
    """Parse an ENF#34 Gassi-Touil workover PDF into the standard dict shape."""
    if isinstance(source, (str, Path)):
        pdf = pdfplumber.open(source)
    else:
        pdf = pdfplumber.open(source)

    try:
        pages_text = [p.extract_text() or "" for p in pdf.pages]
        # Pre-extract tables once (slow operation)
        pages_tables = [p.extract_tables() or [] for p in pdf.pages]
    finally:
        pdf.close()

    full_text = "\n".join(pages_text)
    # Page 2 holds the structured header block (Puits, Appareil, dates,
    # objective). The cover letter on page 1 contains project codes like
    # "PUITS PED" that the regexes below would otherwise match.
    page2_text = pages_text[1] if len(pages_text) > 1 else ""

    # =====================================================================
    # HEADER (page 2 — block of "Label : value" lines)
    # =====================================================================
    header = {}

    # Day number from "Rapport N°09"
    m = re.search(r"Rapport\s+N[°o]\s*0*(\d+)", full_text, re.IGNORECASE)
    if m: header["day_number"] = int(m.group(1))

    # Activity date — scoped to page 2 (page 1 cover letter has another date)
    m = re.search(r"Activit[ée]\s+du\s*:?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
                  page2_text, re.IGNORECASE)
    if m: header["date"] = _date_parse(m.group(1))

    # Well — page 2 only.  Page 1 has "PUITS PED" (project code) which we
    # must not pick up.
    m = re.search(r"Puits\s*:?\s*([A-Z0-9][\w\-]+)", page2_text)
    if m: header["well_name"] = _clean(m.group(1))

    # Rig (Appareil : ENF-34) — page 2
    m = re.search(r"Appareil\s*:?\s*([A-Z0-9][\w\-]+)", page2_text)
    if m: header["rig_name"] = _normalize_rig(m.group(1))

    # Field (CHAMP : Gassi Touil) — on page 3
    m = re.search(r"CHAMP\s*:?\s*([A-Za-zé\-\s]+?)(?:\n|$|DERNIER|BILAN)",
                  full_text, re.IGNORECASE)
    if m: header["field_name"] = _clean(m.group(1))

    # Operation start date (Date début opération : 04/05/2026)
    m = re.search(r"Date\s+d[ée]but\s+op[ée]ration\s*:?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
                  full_text, re.IGNORECASE)
    if m:
        d = _date_parse(m.group(1))
        if d: header["spud_date"] = d

    # Workover objective (BUT DU WORK OVER : ... [until OPERATIONS])
    m = re.search(r"BUT\s+DU\s+WORK\s+OVER\s*:?\s*\n?(.+?)\n\s*OPERATIONS",
                  full_text, re.IGNORECASE | re.DOTALL)
    if m:
        header["well_objective"] = _clean(m.group(1))

    # Last casing shoe: "DERNIER TUBAGE : Sabot Liner 5'' @ 2068.7 m"
    m = re.search(r"DERNIER\s+TUBAGE\s*:?\s*(.+?)(?:\n|$)",
                  full_text, re.IGNORECASE)
    if m:
        # Frontend "Last CSG SHOE" reads from header.top_shoe via insert
        shoe = _clean(m.group(1))
        # Normalize quotes
        shoe = shoe.replace("''", '"').replace("'’", '"').replace("’’", '"')
        header["top_shoe"] = shoe

    # Supervisor (page 4: "SUPERVISEUR SH DP/GTL : S.K. KADRI")
    # Per the two-supervisor convention, the only name on this report goes
    # to header.supervisor — header.superintendent stays empty.
    m = re.search(r"SUPERVISEUR(?:\s+SH\s+DP/GTL)?\s*:?\s*([A-Z][A-Z\.\s\-]+?)(?:\n|$)",
                  full_text)
    if m:
        name = _clean(m.group(1))
        # Strip trailing footer junk if any
        name = re.sub(r"\s+Direction.*$", "", name).strip()
        if name and len(name) <= 60:
            header["supervisor"] = name

    # =====================================================================
    # OPERATIONS (page 2 — "OPERATIONS REALISEES:")
    # Time-banded rows of form: "HH:MM HH:MM description (multi-line)"
    # =====================================================================
    activities = _extract_operations_from_text(pages_text[1] if len(pages_text) > 1 else "")

    # =====================================================================
    # TEXT SECTIONS (page 2)
    # =====================================================================
    text_sections = {}

    # "A / E : Suite fraisage du packer 5'' de 1946,52 m à 1946,72 m en cours"
    m = re.search(r"A\s*/\s*E\s*:?\s*(.+?)(?:\n\s*Programme|\n\s*Direction|$)",
                  full_text, re.DOTALL | re.IGNORECASE)
    if m:
        v = _clean(m.group(1))
        text_sections["current_operation"] = v
        text_sections["day_summary"]       = v

    # "Programme Prévu : Fraisage Packer Electrique 5"
    m = re.search(r"Programme\s+Pr[ée]vu\s*:?\s*(.+?)(?:\n\s*Direction|\n\s*Tél|$)",
                  full_text, re.DOTALL | re.IGNORECASE)
    if m:
        text_sections["plan_operations"] = _clean(m.group(1))

    # "After M/night ... [text until A / E :]"
    m = re.search(r"After\s+M\s*/\s*night\s*\n?(.+?)(?:\n\s*A\s*/\s*E\s*:|\n\s*Direction)",
                  full_text, re.DOTALL | re.IGNORECASE)
    if m:
        v = _clean(m.group(1))
        if v: text_sections["after_midnight"] = v

    # =====================================================================
    # MUD CHECKS (page 3 — "BILAN BOUE" block)
    # =====================================================================
    mud_checks, mud_volume = _extract_mud(pages_text[2] if len(pages_text) > 2 else "")

    # =====================================================================
    # MUD CHEMICAL USAGE (page 3 table)
    # =====================================================================
    chemicals = []
    if len(pages_tables) > 2:
        for tbl in pages_tables[2]:
            # Look for the chemicals table — header row has "Product" or "Réception"
            if not tbl or not tbl[0]: continue
            head = " ".join(str(c) for c in tbl[0] if c)
            if "Product" not in head and "Réception" not in head:
                continue
            for row in tbl[1:]:
                if not row: continue
                # Filter Nones, collapse to columns: item, received, used, stock
                cells = [_clean(c) if c else "" for c in row]
                # item is col 0
                item = cells[0] if cells else ""
                if not item: continue
                # received: first non-empty among cols 1-3
                received = next((c for c in cells[1:4] if c), "")
                # used: prefer the middle column ("Utilisé En T ou M3")
                used  = cells[4] if len(cells) > 4 else ""
                # stock: last column
                stock = cells[-1] if len(cells) >= 2 else ""
                chemicals.append({
                    "item":     item,
                    "units":    "",
                    "received": received,
                    "used":     used,
                    "on_loc":   stock,
                })
            break  # only the first matching table

    # =====================================================================
    # COSTS  (page 4 — "ANALYSE DES COUTS")
    # =====================================================================
    tarif_totals = {}
    daily_cost_total = 0
    cum_cost_total   = 0
    if len(pages_tables) > 3:
        for tbl in pages_tables[3]:
            if not tbl: continue
            head = " ".join(str(c) for c in tbl[0] if c).upper()
            if "TEMPS" in head and "MONTANT" in head:
                # Cost-by-bill table
                for row in tbl[1:]:
                    if not row: continue
                    cells = [_clean(c) if c else "" for c in row]
                    if not cells: continue
                    label = cells[0].upper() if cells[0] else ""
                    if label in ("T1", "T2", "T3", "T4"):
                        # Hours in col 1, hourly cost col 2, total DA in last col
                        try:
                            hrs = _float(cells[1]) if len(cells) > 1 else 0.0
                            tarif_totals[label.lower()] = hrs
                        except Exception:
                            pass
                        try:
                            tarif_totals[f"{label.lower()}_amount"] = _money_to_int(cells[-1])
                        except Exception:
                            pass
                    elif "TOTAL JOURNALIER" in label:
                        daily_cost_total = _money_to_int(cells[-1])
                    elif "CUMUL" in label and "DTM" in label:
                        cum_cost_total = _money_to_int(cells[-1])
            elif "NATURE DES CHARGES" in head:
                # Service-cost table — pick up cum services + grand totals
                for row in tbl[1:]:
                    if not row: continue
                    cells = [_clean(c) if c else "" for c in row]
                    label = cells[0].upper() if cells and cells[0] else ""
                    if "CUMUL APPAREIL & SERVICES AU" in label:
                        # This is the grand total cumulative including services
                        cum_cost_total = max(cum_cost_total, _money_to_int(cells[-1]))

    if daily_cost_total: header["daily_cost"] = daily_cost_total
    if cum_cost_total:   header["cum_cost"]   = cum_cost_total

    # =====================================================================
    # Back-assign the bill code(s) to each operation row.
    #
    # This PDF format doesn't tag individual operations with T1/T2/T3/T4 —
    # only the daily totals in "ANALYSE DES COUTS" know the breakdown.  We
    # partition the ops into groups whose hours sum exactly to each
    # T-bucket total.  See bill_code_assign.assign_bill_codes for the
    # subset-sum partitioning algorithm.
    #
    # In the unambiguous case (only one bucket has hours, e.g. T1=24h and
    # T2=T3=T4=0) every operation simply gets that single code.
    # =====================================================================
    from helpers.bill_code_assign import assign_bill_codes
    activities = assign_bill_codes(activities, tarif_totals)

    # =====================================================================
    # PERSONNEL — not present in this PDF format
    # =====================================================================
    personnel = []

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


# ---------------------------------------------------------------------------
# Operation row extraction
# ---------------------------------------------------------------------------
def _extract_operations_from_text(page_text: str) -> List[dict]:
    """
    Pull out the time-banded operation rows from page 2's
    "OPERATIONS REALISEES" block.

    Lines look like:
        00:00 12:00 Assemblage Junk Mill ...
        12:00 14:00 Descente avec rotation ...
        14:00 00:00 Suite descente avec rotation ...
    With continuation lines indented under each row.
    """
    if not page_text: return []

    # Isolate the OPERATIONS REALISEES block
    block_m = re.search(
        r"OPERATIONS\s+REALISEES\s*:?\s*\n(.+?)(?:\n\s*After\s+M\s*/\s*night|\n\s*A\s*/\s*E\s*:|\n\s*Direction)",
        page_text, re.IGNORECASE | re.DOTALL)
    block = block_m.group(1) if block_m else ""

    if not block: return []

    # Split into rows where each row STARTS with HH:MM HH:MM, and continuation
    # text lines are folded into the preceding row.
    activities = []
    current = None
    row_start_re = re.compile(r"^(\d{2}:\d{2})\s+(\d{2}:\d{2})\s+(.*)$")

    for raw in block.split("\n"):
        line = raw.strip()
        if not line: continue
        m = row_start_re.match(line)
        if m:
            if current is not None:
                activities.append(_finalize_activity(current))
            current = {
                "start": m.group(1),
                "end":   m.group(2),
                "desc":  [m.group(3)],
            }
        elif current is not None:
            current["desc"].append(line)
    if current is not None:
        activities.append(_finalize_activity(current))

    return activities


def _finalize_activity(d: dict) -> dict:
    start_t = _time_parse(d["start"])
    end_t   = _time_parse(d["end"])
    hours = 0.0
    if start_t and end_t:
        sm = start_t.hour * 60 + start_t.minute
        em = end_t.hour * 60 + end_t.minute
        if em == sm:    hours = 24.0
        elif em > sm:   hours = (em - sm) / 60.0
        else:           hours = (em + 1440 - sm) / 60.0
    return {
        "start_time": start_t,
        "end_time":   end_t,
        "hours":      hours,
        "phase_name": "",
        "code": "", "sub": "",
        "description": _clean(" ".join(d["desc"])),
        "start_md": 0, "end_md": 0,
        "npt": 0, "npt_detail": "",
        "npt_company": "", "op_company": "",
        "bill": "",
    }


# ---------------------------------------------------------------------------
# Mud panel extraction
# ---------------------------------------------------------------------------
def _extract_mud(page_text: str):
    """
    Page 3 has a free-text mud block:

        Boue : OBM  Type : OBM  Densité : 1.26  V : 52  Filtrat : 6%
        Solide : 15 %  Salinité : --- g/l  Gel : 6  Gel : 8  YP : 14
                                            0           10
        Vol puits : 38 m³  Perte surf : - m3  Recep :  H/E : 95/5  ES : 100
        Vol surf : 1 m³    V a : 17           V Plast : 10  Eau : 4%  H : 81%
        Transf :  Rempl : m³  Eject : m³  HPHT filt 7%

    pdfplumber sometimes places the subscript digits ("0", "10") on a
    separate physical line. We scan multiple lines with regexes per field.
    """
    mud_checks = {}
    mud_volume = {}
    if not page_text: return mud_checks, mud_volume

    # Restrict scan to the BILAN BOUE block to avoid bleed-in from the title
    # or chemicals table.
    m = re.search(r"BILAN\s+BOUE\s*:?\s*(.+?)(?:PRODUIT\s+&\s+ETATS|\Z)",
                  page_text, re.IGNORECASE | re.DOTALL)
    blob = m.group(1) if m else page_text

    # mud_type: "Boue : OBM" or "Type : OBM"
    m = re.search(r"\bType\s*:?\s*([A-Za-z][A-Za-z0-9\s\-()]*?)(?=\s+(?:Densit|D\s*e\s*n\s*s|V\s*:|Filtrat|\n)|$)",
                  blob, re.IGNORECASE)
    if m:
        mud_checks["mud_type"] = _clean(m.group(1))
    else:
        m2 = re.search(r"\bBoue\s*:?\s*([A-Za-z][A-Za-z0-9\s\-()]*?)(?=\s+(?:Type|Densit|\n)|$)",
                       blob, re.IGNORECASE)
        if m2: mud_checks["mud_type"] = _clean(m2.group(1))

    # Numeric / labeled fields — each looks like "Label : value"
    # Note %-typed fields strip the trailing "%"
    patterns = [
        # (key,                   regex)
        ("density",  r"\bDensit[ée]\s*:?\s*([\d.,]+)"),
        ("fun_vis",  r"\bV\s*:\s*([\d.,]+)"),                       # plain V (Viscosité)
        ("apl_fl",   r"\bFiltrat\s*:?\s*([\d.,]+)\s*%?"),
        ("solid",    r"\bSolide\s*:?\s*([\d.,]+)\s*%?"),
        ("yp",       r"\bYP\s*:?\s*([\d.,]+)"),
        ("es",       r"\bES\s*:?\s*([\d.,]+)"),
        ("pv",       r"\bV\s+Plast\s*:?\s*([\d.,]+)"),
        ("hpht_fl",  r"\bHPHT\s+filt(?:\s+\d*)?\s*[:.]?\s*([\d.,]+)\s*%?"),
        ("h2o",      r"\bEau\s*:?\s*([\d.,]+)\s*%?"),
        ("oil",      r"\bH\s*:?\s*([\d.,]+)\s*%"),   # H : 81% (oil percentage in OBM)
    ]
    for key, rx in patterns:
        m = re.search(rx, blob, re.IGNORECASE)
        if m:
            val = _float(m.group(1))
            if val or val == 0.0:
                mud_checks[key] = val

    # Gel0 / Gel10 — labels and values sometimes on separate lines
    # Try inline pattern first: "Gel0 : 6  Gel10 : 8"
    m = re.search(r"Gel\s*0?\s*:?\s*([\d.,]+).{0,40}Gel\s*1?0?\s*:?\s*([\d.,]+)",
                  blob, re.IGNORECASE | re.DOTALL)
    if m:
        mud_checks["gel10sec"] = _float(m.group(1))
        mud_checks["gel10m"]   = _float(m.group(2))

    # Oil/Water ratio "H/E : 95/5" -> store as ratio text
    m = re.search(r"H\s*/\s*E\s*:?\s*(\d+\s*/\s*\d+)", blob, re.IGNORECASE)
    if m:
        mud_checks["oil_water_ratio"] = _clean(m.group(1))

    # Volumes
    m = re.search(r"Vol\s+puits\s*:?\s*([\d.,]+)", blob, re.IGNORECASE)
    if m: mud_volume["string_volume"] = _float(m.group(1))
    m = re.search(r"Vol\s+surf\s*:?\s*([\d.,]+)", blob, re.IGNORECASE)
    if m: mud_volume["pits_volume"]   = _float(m.group(1))
    m = re.search(r"Perte\s+surf\s*:?\s*([\d.,\-]+)", blob, re.IGNORECASE)
    if m:
        try: mud_volume["surface_loss"] = _float(m.group(1))
        except: pass
    if mud_volume:
        mud_volume["total_volume"] = sum(
            v for v in (mud_volume.get("string_volume"),
                        mud_volume.get("pits_volume"))
            if isinstance(v, (int, float))
        )

    return mud_checks, mud_volume


# Drop-in compat
parse_daily_excel_report = parse_enf34_pdf


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: enf34_pdf_extract.py SOURCE.pdf")
    data = parse_enf34_pdf(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))