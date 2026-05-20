#!/usr/bin/env python3
"""
tp186_extract.py — extract a TP-186 (ENTP rig 186, ZR wells, ZARZAITINE
field) Daily Workover Report from a Word file (.doc or .docx) into the
standard dict shape.

Source layout
-------------
Microsoft Word document (legacy .doc or modern .docx).  Title at
paragraph 2: "RAPPORT JOURNALIER WORK-OVER".

The document uses 4 tables:
    Table 0 — header (well, field, rig, mud, BHA, params)
    Table 1 — operations + chemicals + situation + plan
              Operations are in a SINGLE cell with newline-separated
              lines, each prefixed by a time band "HHhMM-HHhMM" or
              continuation lines (no leading time band).
    Table 2 — tarif breakdown: T1/T2/T3/T4/NR hours and amounts (DA)
    Table 3 — service companies and their costs

Two supervisor names sometimes appear (Responsable Section + supervisor
signature).  When only one is present, it goes to `header.supervisor`
and `superintendent` stays empty (per the two-supervisor convention).

.doc files are auto-converted to .docx using LibreOffice headless (must
be installed on the deployment host).  .docx files are read directly.

Distinguishing markers (used by helpers.parse_source._detect_format):
    - file extension .doc or .docx
    - contains "RAPPORT JOURNALIER WORK-OVER" or "TP # 186" / "TP 186"
"""
from __future__ import annotations
import re
import subprocess
import tempfile
from datetime import datetime, time, date as date_type, timedelta
from io import BytesIO
from pathlib import Path
from typing import Union, List, Tuple, Optional

from docx import Document


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _clean(v) -> str:
    if v is None: return ""
    s = str(v)
    # Strip Word's narrow no-break space (\u202f) and regular no-break (\u00a0)
    s = s.replace("\u202f", " ").replace("\u00a0", " ")
    return re.sub(r"\s+", " ", s).strip()


def _float(v, default=0.0) -> float:
    if v is None: return default
    if isinstance(v, (int, float)): return float(v)
    s = _clean(v).replace(",", ".").replace(" ", "")
    if s in ("", "-", "/", "None"): return default
    # Strip trailing unit (handles 'm3', 'm³', '%', 'kg', 'psi', etc.):
    # take everything up to the last digit + optional decimal portion.
    m = re.match(r"^([-+]?\d+(?:\.\d+)?)", s)
    if m:
        try: return float(m.group(1))
        except ValueError: pass
    return default


def _int(v, default=0) -> int:
    return int(_float(v, float(default)))


def _money_to_int(s) -> int:
    """'2 350 649,53 DA' or '2\\u202f350\\u202f649' → 2350649."""
    if s is None: return 0
    if isinstance(s, (int, float)): return int(s)
    cleaned = _clean(s)
    # Drop currency suffix
    cleaned = re.sub(r"\s*DA\s*$", "", cleaned, flags=re.IGNORECASE)
    # Drop decimals (we want integer DA)
    cleaned = re.split(r"[.,]\d{2,3}\s*$", cleaned)[0]
    # Strip non-digits
    digits = re.sub(r"\D", "", cleaned)
    return int(digits) if digits else 0


def _parse_hhmm_duration(s: str) -> float:
    """'22H30' / '22h30' / '00H00' → decimal hours."""
    s = _clean(s)
    m = re.match(r"^(\d{1,2})[hH:](\d{2})$", s)
    if not m: return 0.0
    return int(m.group(1)) + int(m.group(2)) / 60.0


def _parse_time_band(text: str) -> Optional[Tuple[time, time, str]]:
    """Parse 'HHhMM-HHhMM <description>' → (start, end, desc) or None."""
    s = text.strip()
    m = re.match(r"^(\d{1,2})[hH:](\d{2})\s*[-–]\s*(\d{1,2})[hH:](\d{2})\s*(.*)$",
                 s, re.DOTALL)
    if not m: return None
    try:
        sh, sm, eh, em = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
        # End-of-day "00h00" means midnight tomorrow → keep as time(0,0)
        start_t = time(sh % 24, sm)
        end_t   = time(eh % 24, em)
        return start_t, end_t, m.group(5).strip()
    except ValueError:
        return None


def _hours_between(start_t: time, end_t: time) -> float:
    sm = start_t.hour * 60 + start_t.minute
    em = end_t.hour * 60 + end_t.minute
    if em == sm:    return 24.0
    if em > sm:     return (em - sm) / 60.0
    return (em + 1440 - sm) / 60.0


def _find_libreoffice() -> str:
    """Locate the LibreOffice executable across platforms.

    Search order:
      1. LIBREOFFICE_PATH environment variable (lets the user override)
      2. `soffice` / `libreoffice` on PATH (Linux / macOS / Windows-with-PATH)
      3. Common Windows install locations (Program Files / Program Files (x86))
      4. macOS .app bundle
    Returns the executable path as a string, or raises RuntimeError with
    a platform-specific install hint.
    """
    import os, shutil, sys

    # 1. Explicit override
    env = os.environ.get("LIBREOFFICE_PATH")
    if env and Path(env).exists():
        return env

    # 2. On PATH — both common command names
    for cmd in ("soffice", "libreoffice"):
        found = shutil.which(cmd)
        if found:
            return found

    # 3. Windows default install locations
    if sys.platform.startswith("win"):
        candidates = [
            Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
            Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
            # LibreOffice 7.x / 25.x specific paths
            Path(r"C:\Program Files\LibreOffice 7\program\soffice.exe"),
            Path(r"C:\Program Files\LibreOffice 25\program\soffice.exe"),
        ]
        for p in candidates:
            if p.exists():
                return str(p)

    # 4. macOS bundle
    if sys.platform == "darwin":
        for p in (
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
            Path.home() / "Applications/LibreOffice.app/Contents/MacOS/soffice",
        ):
            if p.exists():
                return str(p)

    # Not found — raise with platform-appropriate hint
    if sys.platform.startswith("win"):
        hint = (
            "LibreOffice not found. Install it from https://www.libreoffice.org/\n"
            "and either:\n"
            "  (a) add C:\\Program Files\\LibreOffice\\program to your PATH, or\n"
            "  (b) set the LIBREOFFICE_PATH environment variable to the full\n"
            "      path of soffice.exe, e.g.\n"
            "      setx LIBREOFFICE_PATH \"C:\\Program Files\\LibreOffice\\program\\soffice.exe\""
        )
    elif sys.platform == "darwin":
        hint = ("LibreOffice not found. Install via "
                "`brew install --cask libreoffice` or download from "
                "https://www.libreoffice.org/")
    else:
        hint = ("LibreOffice not found on PATH. Install with "
                "`apt-get install libreoffice` (Debian/Ubuntu) or set "
                "LIBREOFFICE_PATH to the soffice executable.")
    raise RuntimeError(hint)


def _ensure_docx(source_path: Path) -> Path:
    """If source is .doc (legacy binary), convert to .docx using LibreOffice
    headless.  Returns the path to the .docx file (may be the same path if
    the source was already .docx)."""
    suffix = source_path.suffix.lower()
    if suffix == ".docx":
        return source_path
    if suffix != ".doc":
        raise ValueError(f"Unsupported Word file extension: {suffix}")

    soffice = _find_libreoffice()

    # Convert via LibreOffice — must be installed on the host.  Output goes
    # to a temp dir to avoid polluting the input directory.
    tmp_dir = Path(tempfile.mkdtemp(prefix="docconv_"))
    try:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "docx",
             "--outdir", str(tmp_dir), str(source_path)],
            capture_output=True, text=True, timeout=120,
        )
    except FileNotFoundError:
        # Can still happen if soffice was found but lacks exec permission, or
        # on Windows where shutil.which returned a stub that doesn't exist.
        raise RuntimeError(
            f"LibreOffice was located at {soffice!r} but couldn't be executed. "
            "Check that the file exists and is runnable, or set "
            "LIBREOFFICE_PATH to a working soffice executable."
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("LibreOffice conversion timed out (>120s)")
    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed: {result.stderr[:500]}"
        )
    docx_path = tmp_dir / (source_path.stem + ".docx")
    if not docx_path.exists():
        raise RuntimeError(
            f"Expected output not found after conversion: {docx_path}"
        )
    return docx_path


def _date_from_text(s: str) -> Optional[date_type]:
    """Extract first DD/MM/YYYY or DD-MM-YYYY from a string."""
    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if not m: return None
    try:
        return date_type(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def _normalize_rig(name: str) -> str:
    """'TP # 186' / 'TP-186' → 'TP-186'."""
    s = _clean(name).upper()
    m = re.search(r"TP[\s#\-]*(\d+)", s)
    if m: return f"TP-{m.group(1)}"
    return _clean(name)


# ---------------------------------------------------------------------------
# Operations parser
# ---------------------------------------------------------------------------
def _parse_parallel_ops(time_lines: List[str], desc_lines: List[str]) -> List[dict]:
    """
    Time bands and descriptions live in separate parallel cells, with
    lines aligned by index:

        cell 0 ('HORAIRE'):              cell 1 ('OPERATIONS'):
        L0: 00h00-07h30                  L0: Suite descente complétion …
        L1: 07h30-09h00                  L1: Contrôle Wire Line: …
        L2: (blank)                      L2:     Descente GC = 91,50mm …  ← continuation
        L3: 09h00-00h00                  L3: Suite descente complétion …

    Rule: a line at index i in cell 0 with a parseable HHhMM-HHhMM band
    starts a new op; the same index in cell 1 is its description.  A
    blank time-band line at index i means cell 1's line at index i is a
    continuation of the previous op's description.
    """
    activities = []
    current = None
    # Pad whichever list is shorter so the indices line up
    n = max(len(time_lines), len(desc_lines))
    time_lines = list(time_lines) + [""] * (n - len(time_lines))
    desc_lines = list(desc_lines) + [""] * (n - len(desc_lines))

    for i in range(n):
        tl = time_lines[i].strip()
        dl = desc_lines[i].rstrip()
        # Try to parse a time band on this line of cell 0
        m = re.match(r"^(\d{1,2})[hH:](\d{2})\s*[-–]\s*(\d{1,2})[hH:](\d{2})\s*$", tl)
        if m:
            # New op starts here
            if current is not None:
                activities.append(_finalize_op(current))
            start_t = time(int(m.group(1)) % 24, int(m.group(2)))
            end_t   = time(int(m.group(3)) % 24, int(m.group(4)))
            current = {
                "start_time": start_t,
                "end_time":   end_t,
                "desc_parts": [dl.strip()] if dl.strip() else [],
            }
        else:
            # Time band absent / unparseable → cell 1 line is a continuation
            if current is not None and dl.strip():
                current["desc_parts"].append(dl.strip())
    if current is not None:
        activities.append(_finalize_op(current))
    return activities


def _parse_operations_cell(cell_text: str) -> List[dict]:
    """
    The operations live in a single cell as newline-separated lines.
    Some lines start with a time band 'HHhMM-HHhMM' and have a description
    on the same line.  Continuation lines (indented or without a time
    band) belong to the previous operation.

    Returns a list of activity dicts.
    """
    activities = []
    current = None
    for raw_line in cell_text.split("\n"):
        line = raw_line.rstrip()
        if not line.strip():
            continue
        band = _parse_time_band(line)
        if band:
            # New operation row
            if current is not None:
                activities.append(_finalize_op(current))
            start_t, end_t, desc = band
            current = {
                "start_time": start_t,
                "end_time":   end_t,
                "desc_parts": [desc] if desc else [],
            }
        elif current is not None:
            # Continuation line — fold into current op's description
            current["desc_parts"].append(line.strip())
        # else: line before any time band — ignore (likely a stray header)
    if current is not None:
        activities.append(_finalize_op(current))
    return activities


def _finalize_op(d: dict) -> dict:
    start_t = d["start_time"]
    end_t   = d["end_time"]
    hours   = _hours_between(start_t, end_t)
    return {
        "start_time": start_t,
        "end_time":   end_t,
        "hours":      hours,
        "phase_name": "",
        "code": "", "sub": "",
        "description": _clean(" ".join(d["desc_parts"])),
        "start_md": 0, "end_md": 0,
        "npt": 0, "npt_detail": "",
        "npt_company": "", "op_company": "",
        "bill": "",        # back-filled from tarif totals
    }


# ---------------------------------------------------------------------------
# Cell-by-cell table extraction with merged-cell handling
# ---------------------------------------------------------------------------
def _row_cells(row) -> List[str]:
    """Return one entry per VISUAL cell.  Word's `cells` collection
    repeats the same cell object for spans across columns, so adjacent
    cells with identical text are deduplicated — but only when their
    underlying tc element is the same (true merged cell), not when two
    different cells just happen to hold the same string."""
    out = []
    prev_tc = None
    for c in row.cells:
        tc = c._tc        # underlying XML element
        if tc is prev_tc:
            continue      # same merged cell — skip
        out.append(c.text.strip())
        prev_tc = tc
    return out


def _find_value_after(text_row: List[str], label_substr: str) -> str:
    """In a list of row cells, find the cell whose normalized text EQUALS
    label_substr (case-insensitive, whitespace-normalized) and return the
    next cell's value.

    Exact match is important here because some rows contain multiple
    "Volume *" or "Type *" labels — substring matching would pick the
    first one every time.  When no exact match is found, falls back to
    "starts with" matching."""
    norm = lambda s: re.sub(r"\s+", " ", s.strip().lower())
    target = norm(label_substr)
    # First pass: exact (normalized) match
    for i, cell in enumerate(text_row):
        if norm(cell) == target and i + 1 < len(text_row):
            return text_row[i + 1].strip()
    # Second pass: starts-with match
    for i, cell in enumerate(text_row):
        if norm(cell).startswith(target) and i + 1 < len(text_row):
            return text_row[i + 1].strip()
    return ""


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------
def parse_tp186(source: Union[Path, str, BytesIO]) -> dict:
    """Parse a TP-186 (Word .doc / .docx) workover report."""
    if isinstance(source, BytesIO):
        # Stream a docx directly; .doc binary streams need conversion which
        # requires a file path — buffered to disk first
        try:
            doc = Document(source)
            tmp_path = None
        except Exception:
            with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as tf:
                tf.write(source.getvalue())
                tmp_path = Path(tf.name)
            docx_path = _ensure_docx(tmp_path)
            doc = Document(str(docx_path))
    else:
        source_path = Path(source)
        docx_path = _ensure_docx(source_path)
        doc = Document(str(docx_path))

    # Extract paragraphs and tables
    paragraphs = [_clean(p.text) for p in doc.paragraphs]
    tables = doc.tables

    # =====================================================================
    # HEADER (table 0 + paragraphs)
    # =====================================================================
    header = {}

    # Date — from paragraph "RAPPORT N° 40 du 16/05/2026"
    for p in paragraphs:
        m = re.search(r"du\s+(\d{1,2}[/-]\d{1,2}[/-]\d{4})", p, re.IGNORECASE)
        if m:
            d = _date_from_text(m.group(1))
            if d:
                header["date"] = d
                break
        m = re.search(r"RAPPORT\s*N\s*[°o]\s*(\d+)", p, re.IGNORECASE)
        if m:
            header["day_number"] = int(m.group(1))

    # Pull rig/well/field from table 0
    if tables:
        t0 = tables[0]
        for ri, row in enumerate(t0.rows):
            cells = _row_cells(row)
            joined = " | ".join(cells)
            # PUITS : ZR#708 | Champ | ZARZAITINE | Appareil | TP # 186
            puits = _find_value_after(cells, "puits")
            if puits and "well_name" not in header:
                header["well_name"] = puits.replace(" ", "")
            champ = _find_value_after(cells, "champ")
            if champ and "field_name" not in header:
                header["field_name"] = champ
            app = _find_value_after(cells, "appareil")
            if app and "rig_name" not in header:
                header["rig_name"] = _normalize_rig(app)
            # Mud volumes
            vt = _find_value_after(cells, "volume total")
            if vt: header["_mud_total"] = vt
            vs = _find_value_after(cells, "volume surface")
            if vs: header["_mud_surface"] = vs
            vp = _find_value_after(cells, "volume puits")
            if vp: header["_mud_well"] = vp
            # Mud density / viscosity / filtrat
            for label, key in [("densité", "_density"), ("v. marsh", "_fun_vis"),
                                ("filtrat", "_filtrat"), ("type", "_mud_type")]:
                v = _find_value_after(cells, label)
                if v:
                    header[key] = v
            # Bit
            diam = _find_value_after(cells, "diamètre")
            if diam: header["_bit_diam"] = diam

    # Supervisor — paragraph after "Responsable de la Section"
    for i, p in enumerate(paragraphs):
        if "RESPONSABLE" in p.upper() and "SECTION" in p.upper():
            # Name is usually 1-3 paragraphs down
            for j in range(i + 1, min(i + 6, len(paragraphs))):
                cand = paragraphs[j]
                if cand and len(cand) < 60 and not cand.startswith(("S", "RAPPORT", "OPERATION")):
                    # Looks like a name: short, no leading section keyword
                    if re.match(r"^[A-Z][\.\s]", cand) or "." in cand[:5]:
                        header["supervisor"] = cand
                        break
            break

    # =====================================================================
    # OPERATIONS (table 1)
    # Layout: cell 0 holds time bands ('HHhMM-HHhMM' newline-separated),
    # cell 1 holds the matching descriptions (also newline-separated).
    # Lines line up by position; description lines that lack a parallel
    # time band are continuations of the previous op.
    # =====================================================================
    activities = []
    if len(tables) > 1:
        t1 = tables[1]
        # Find the "HORAIRE" header row, then read the cells of the next row
        for ri, row in enumerate(t1.rows):
            cells = _row_cells(row)
            if any("HORAIRE" in c.upper() for c in cells):
                if ri + 1 < len(t1.rows):
                    data_cells = _row_cells(t1.rows[ri + 1])
                    if len(data_cells) >= 2:
                        time_lines = [l for l in data_cells[0].split("\n")]
                        desc_lines = [l for l in data_cells[1].split("\n")]
                        activities = _parse_parallel_ops(time_lines, desc_lines)
                break

    # =====================================================================
    # TARIF TOTALS (table 2)
    # =====================================================================
    tarif_totals = {}
    if len(tables) > 2:
        t2 = tables[2]
        for row in t2.rows:
            cells = _row_cells(row)
            if len(cells) < 2: continue
            label = _clean(cells[0]).upper().replace(" ", "").replace(".", "")
            if label in ("T1", "T2", "T3", "T4"):
                hrs = _parse_hhmm_duration(cells[1])
                if hrs > 0:
                    tarif_totals[label.lower()] = hrs

    # Back-fill bill codes via the shared helper
    from helpers.bill_code_assign import assign_bill_codes
    activities = assign_bill_codes(activities, tarif_totals)

    # =====================================================================
    # COSTS (table 2 TOTAL row + paragraph "CUMUL DEPUIS ORIGINE")
    # =====================================================================
    if len(tables) > 2:
        t2 = tables[2]
        for row in t2.rows:
            cells = _row_cells(row)
            if cells and "TOTAL" in cells[0].upper():
                # Find the rightmost monetary value
                for c in reversed(cells):
                    if "DA" in c.upper():
                        v = _money_to_int(c)
                        if v > 0:
                            header["daily_cost"] = v
                            break
                break

    # Cumul depuis origine — paragraph
    for p in paragraphs:
        if "CUMUL DEPUIS ORIGINE" in p.upper() and "APPAREIL" in p.upper():
            v = _money_to_int(p)
            if v > 0:
                header["cum_cost"] = v
                break

    # =====================================================================
    # MUD CHECKS + VOLUMES (parsed from header values stashed above)
    # =====================================================================
    mud_checks = {}
    mud_volume = {}
    if header.pop("_mud_type", None):
        mt = header.get("_mud_type") or ""
        # was already popped — read directly from local var instead
    # Properly extract via header keys (mutate-copy)
    for src_key, dst_key in [("_density", "density"), ("_fun_vis", "fun_vis"),
                              ("_filtrat", "apl_fl")]:
        raw = header.pop(src_key, None)
        if raw is not None:
            v = _float(raw)
            if v: mud_checks[dst_key] = v
    mt_raw = header.pop("_mud_type", None) or header.get("_mud_type")
    # Use the type already extracted (may be 'Huile')
    # Mud type was popped above (if it ever was); refind it from t0
    if tables:
        for row in tables[0].rows:
            cells = _row_cells(row)
            tv = _find_value_after(cells, "type")
            if tv and tv.lower() in ("huile", "obm", "oil", "wbm", "water"):
                mud_checks["mud_type"] = tv
                break
    # Volumes
    for src_key, dst_key in [("_mud_total", "total_volume"),
                              ("_mud_surface", "pits_volume"),
                              ("_mud_well", "string_volume")]:
        raw = header.pop(src_key, None)
        if raw is not None:
            v = _float(raw)
            if v: mud_volume[dst_key] = v

    # Strip the temporary _bit_diam etc.
    for k in list(header.keys()):
        if k.startswith("_"):
            header.pop(k, None)

    # =====================================================================
    # MUD CHEMICAL USAGE (table 1: a sub-cell after "DESIGNATIONS")
    # =====================================================================
    chemicals = []
    if len(tables) > 1:
        t1 = tables[1]
        for ri, row in enumerate(t1.rows):
            cells = _row_cells(row)
            if any("DESIGNATION" in c.upper() for c in cells):
                if ri + 1 < len(t1.rows):
                    next_cells = _row_cells(t1.rows[ri + 1])
                    # Format: [items_block, used_block, stock_block]
                    if len(next_cells) >= 3:
                        items = next_cells[0].split("\n")
                        useds  = next_cells[1].split("\n") if next_cells[1] else []
                        stocks = next_cells[2].split("\n") if next_cells[2] else []
                        for i, item in enumerate(items):
                            item = item.strip()
                            if not item: continue
                            chemicals.append({
                                "item": item,
                                "units": "",
                                "received": "",
                                "used":   (useds[i].strip() if i < len(useds) else ""),
                                "on_loc": (stocks[i].strip() if i < len(stocks) else ""),
                            })
                break

    # =====================================================================
    # TEXT SECTIONS (table 1: "Situation à 06h00" and "Programme prévu")
    # =====================================================================
    text_sections = {}
    if len(tables) > 1:
        for row in tables[1].rows:
            cells = _row_cells(row)
            if not cells: continue
            label = cells[0].upper() if cells[0] else ""
            value = cells[1] if len(cells) > 1 else ""
            value_clean = _clean(value)
            if not value_clean: continue
            if "SITUATION" in label:
                # Cap at 300 chars (frontend display limit)
                if len(value_clean) > 300:
                    value_clean = value_clean[:300].rsplit(None, 1)[0] + "…"
                text_sections["current_operation"] = value_clean
                text_sections["day_summary"]       = value_clean
            elif "PROGRAMME" in label:
                text_sections["plan_operations"] = value_clean

    # =====================================================================
    # SERVICE COMPANIES (table 3 — not used downstream but capture for future)
    # =====================================================================
    service_companies = []
    if len(tables) > 3:
        for row in tables[3].rows:
            cells = _row_cells(row)
            if len(cells) < 3: continue
            name = cells[0].strip()
            if not name or name.upper() in ("SOCIETES", "TOTAL"): continue
            if name in ("/",): continue
            service_companies.append({
                "company": name,
                "nature":  cells[1].strip(),
                "amount":  _money_to_int(cells[2]),
            })

    return {
        "header": header,
        "activities": activities,
        "text_sections": text_sections,
        "mud_checks": mud_checks,
        "mud_volume": mud_volume,
        "mud_chemical_usage": chemicals,
        "personnel_data": [],
        "pumps": [],
        "well_location": {},
        "survey_data": [],
        "safety": {},
        "tarif_totals": tarif_totals,
        "service_companies": service_companies,
    }


# Drop-in compat
parse_daily_excel_report = parse_tp186


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        sys.exit("Usage: tp186_extract.py SOURCE.doc|SOURCE.docx")
    data = parse_tp186(Path(sys.argv[1]))

    def default(o):
        if isinstance(o, (date_type, datetime)): return o.isoformat()
        if isinstance(o, time): return o.strftime("%H:%M:%S")
        if isinstance(o, timedelta): return o.total_seconds()
        return str(o)
    print(json.dumps(data, indent=2, default=default, ensure_ascii=False))