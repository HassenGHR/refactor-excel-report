#!/usr/bin/env python3
"""
parse_source.py — auto-detect rig-report format and dispatch to the
appropriate extractor.

Accepts both Excel sources (.xlsx — most rigs) and PDF sources (some rigs
deliver their daily reports as PDF instead of Excel). The file type is
sniffed from the first few bytes (PDF starts with %PDF-, xlsx with PK\x03)
so callers don't need to declare the format up front.

Usage (programmatic):
    from helpers.parse_source import parse_source
    data = parse_source(Path("report.xlsx"))   # Excel
    data = parse_source(Path("report.pdf"))    # PDF
    data = parse_source(BytesIO(file_bytes))   # any in-memory source

Usage (CLI):
    python parse_source.py report.xlsx           # prints detected format
    python parse_source.py report.pdf --json     # prints the extracted dict
"""
from __future__ import annotations
import re
import sys
from pathlib import Path
from typing import Union
from io import BytesIO

from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# File-type sniffing — by magic bytes, not by extension
# ---------------------------------------------------------------------------
def _read_head(source) -> bytes:
    """Read the first 8 bytes of a source without consuming it."""
    if isinstance(source, (str, Path)):
        with open(source, "rb") as f:
            return f.read(8)
    # BytesIO or similar: remember position, peek, restore
    pos = source.tell()
    head = source.read(8)
    source.seek(pos)
    return head


def _sniff_kind(source) -> str:
    """Return 'pdf', 'xlsx', 'docx', 'doc', or 'unknown' based on the
    file's magic bytes (and, for .docx, the file extension since it shares
    the zip magic with .xlsx)."""
    head = _read_head(source)
    if head.startswith(b"%PDF-"):
        return "pdf"
    # Legacy Word/Excel format (OLE2 Compound File Binary)
    if head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
        return "doc"
    # .xlsx and .docx are both zip archives — disambiguate by extension
    if head.startswith(b"PK\x03\x04"):
        if isinstance(source, (str, Path)):
            suffix = Path(source).suffix.lower()
            if suffix == ".docx":
                return "docx"
        return "xlsx"
    return "unknown"


# ---------------------------------------------------------------------------
# Excel format detection
# ---------------------------------------------------------------------------
def _detect_format_xlsx(source) -> str:
    """Peek at an Excel source and return the rig template key."""
    if isinstance(source, (str, Path)):
        wb = load_workbook(source, data_only=True, read_only=True)
    else:
        wb = load_workbook(source, data_only=True, read_only=True)
    ws = wb.active

    # Scan the first 12 rows × 28 cols for marker strings
    markers = []
    for row in ws.iter_rows(min_row=1, max_row=12, max_col=28, values_only=True):
        for v in row:
            if v is not None:
                markers.append(str(v).upper())
    blob = " || ".join(markers)
    wb.close()

    # ENTP-204 — AIN T'SILA DWR, TXNO wells, rig ENTP 204.  Same SONATRACH
    # template family as TP-182 / TP-195 but stores header as combined
    # "LABEL : value" strings and has a FROM/TO/HRS/DESCRIPTION/BILL ops
    # table.  Must be checked BEFORE TP-195 and TP-182 since it shares
    # their "OFFICE REP"-less SONATRACH markers; distinguished by the
    # ENTP-204 rig name or TXNO well prefix.
    if ("ENTP 204" in blob or "ENTP204" in blob or "ENTP-204" in blob
            or re.search(r"\bTXNO[-\s]?\d", blob)):
        return "entp204"

    # TP-195 — SONATRACH AIN T'SILA format.  Same template family as TP-182
    # (English, "DAILY DRILLING REPORT" title) but uses split label/value
    # cells, "OFFICE REP" instead of "SUPERINTANDANT", and only "NEXT BOP
    # TEST" (no LAST BOP).  Must be checked BEFORE TP-182 since both share
    # the SONATRACH PRODUCTION DIVISION title.
    if "OFFICE REP" in blob:
        return "tp195"

    # TP-182 — SONATRACH PRODUCTION DIVISION Daily Drilling Report format
    # (English, "SUPERINTANDANT" misspelling, has "WORKOVER REASON")
    if "SUPERINTANDANT" in blob:        # specific to TP-182 template
        return "tp182"
    if "SONATRACH PRODUCTION DIVISION" in blob and "DAILY DRILLING REPORT" in blob:
        return "tp182"

    # ENAFOR ENF#04 — French workover format (Haoud Berkaoui, DDNH wells).
    # Two layouts in circulation: 2026-05-10 title "RAPPORT JOURNALIER DE
    # WORKOVER" (one word), 2026-05-16 title "RAPPORT JOURNALIER DE
    # WORK-OVER" (hyphenated, lowercase).  Both have "Haoud Berkaoui" in
    # the regional-direction line OR DDNH-NN wells.  Must be checked
    # BEFORE TP-173 and TP-179 since they all share "RAPPORT JOURNALIER".
    if ("HAOUD BERKAOUI" in blob
        or "RAPPORT JOURNALIER DE WORKOVER" in blob
        or "RAPPORT JOURNALIER DE WORK-OVER" in blob
        or re.search(r"\bDDNH[-\s]?\d", blob)):
        return "enf04"

    # TP-183 — ENTP rig 183, TMLS wells.  Different title from TP-173/TP-179
    # ("RAPPORT JOURNALIER WORK OVER" — no "DE", no "DU") and uses a
    # compact ~62-row single-sheet template with cost analysis.  Must be
    # checked BEFORE both TP-173 (shares "DERNIER TUBAGE" label) AND GW29
    # (shares "PARAMETRES" marker).
    if "TP 183" in blob or "TP-183" in blob or re.search(r"\bTMLS\b", blob):
        return "tp183"

    # ENTP-127 — English-template daily work-over report (ENTP DF rig,
    # DAD wells, DRAA DAOUI field).  Distinguishing markers vs the French
    # GW29/TP-127 template: "ENTP   DF" in A1 (rendered as "ENTP  DF" or
    # with extra spaces) and English-only labels "TOOL PUSHER" / "LAST BOP
    # TEST" in the header area.  Must be checked BEFORE gw29 since the gw29
    # block matches "TP 127" which also appears here.
    if (re.search(r"ENTP\s+DF", blob)
            and ("TOOL PUSHER" in blob or "LAST BOP TEST" in blob)):
        return "entp127"

    # GW29 layout family — French "Rapport journalier de Work - Over" with
    # the AVANCEMENT / OUTILS / USURE / PARAMETRES section headers at row 3
    # and the wide header in rows 1-2 (label/value pairs).  Originally for
    # GW-series rigs (GWDC operator, RBL wells, REB field); ALSO matches
    # TP-127 (DAD#1ST well, DRAA DAOUI field) which uses the same template
    # — confirmed identical layout.
    # Must be checked BEFORE the generic TP-179 catch-all below since these
    # files all have "RAPPORT JOURNALIER" + "WORK".
    if "RAPPORT JOURNALIER" in blob and ("AVANCEMENT" in blob
                                         or "PARAMETRES" in blob
                                         or re.search(r"\bGW\s?\d{2}\b", blob)
                                         or "TP#127" in blob
                                         or "TP-127" in blob
                                         or "TP 127" in blob
                                         or re.search(r"\bDAD[#\-\s]?\d", blob)):
        return "gw29"

    # TP-173 — same template family as TP-179 (RAPPORT JOURNALIER DU WORK
    # OVER, ADRAR region) but with different cell positions: "APPAREIL:"
    # instead of "RIG :", "DERNIER TUBAGE:" combined string for casing.
    # Must be checked BEFORE TP-179 since both share the title.
    if "TP-173" in blob or "DERNIER TUBAGE" in blob:
        return "tp173"

    # TP-179 — French workover format (ENTP rigs)
    if "RAPPORT JOURNALIER" in blob and "WORK" in blob:
        return "tp179"

    # ENAFOR ENF#33 — BERKINE field, BKNS wells.  Wide single-sheet English
    # DDR with a DIFFERENT cell layout from the ENF#17 DDR, so it needs its
    # own extractor.  Both share the "DAILY DRILLING REPORT" title and the
    # generic "RIG.S.I" field label, so we distinguish ENF#33 by its
    # specific rig number ("ENF # 33") or BKNS well prefix — NOT by RIG.S.I,
    # which appears in every ENAFOR DDR.  Must be checked BEFORE the generic
    # ENF branch below.
    if "ENF # 33" in blob or "ENF#33" in blob or re.search(r"\bBKNS[-\s]?\d", blob):
        return "enf33"

    # ENAFOR DDR (ENF#NN rigs) — distinguishing markers
    if "DAILY DRILLING REPORT" in blob and ("ENF#" in blob or "ENF #" in blob):
        return "enf"

    # Generic drilling — try enf as a fallback
    if "DAILY DRILLING REPORT" in blob:
        return "enf"

    return "unknown"


# ---------------------------------------------------------------------------
# PDF format detection
# ---------------------------------------------------------------------------
def _detect_format_pdf(source) -> str:
    """Peek at a PDF source and return the rig template key.

    Reads the first 2 pages and scans for distinguishing marker phrases.
    pdfplumber is imported lazily so the module still loads on systems
    without it when only Excel/PDF sources are used.
    """
    import pdfplumber

    if isinstance(source, BytesIO):
        # pdfplumber consumes BytesIO; rewind it for the caller after detection
        pos = source.tell()
        try:
            with pdfplumber.open(source) as pdf:
                pages_to_scan = pdf.pages[:2]
                blob = " || ".join(
                    (p.extract_text() or "").upper() for p in pages_to_scan
                )
        finally:
            source.seek(pos)
    else:
        with pdfplumber.open(source) as pdf:
            pages_to_scan = pdf.pages[:2]
            blob = " || ".join(
                (p.extract_text() or "").upper() for p in pages_to_scan
            )

    # ENF#34 — Gassi-Touil workover PDF (this is the first PDF format we
    # support).  Distinguishing markers: "GASSI" + "RAPPORT JOURNALIER DE
    # WORK OVER" (note the space in "WORK OVER" — TP-179 uses "WORKOVER").
    if "GASSI" in blob and "RAPPORT JOURNALIER" in blob and "WORK" in blob:
        return "enf34_pdf"
    if "DIRECTION RÉGIONALE GASSI" in blob or "GASSI-TOUIL" in blob:
        return "enf34_pdf"

    return "unknown"


# ---------------------------------------------------------------------------
# Word format detection
# ---------------------------------------------------------------------------
def _detect_format_word(source) -> str:
    """Peek at a Word source (.doc or .docx) and return the rig template
    key.  Legacy .doc files are converted to .docx via LibreOffice (the
    extractor handles this); for detection we just need to read the text.

    Imports python-docx and (for .doc) the extractor's _ensure_docx helper
    lazily so the dispatcher doesn't depend on Word support being
    installed when callers only use Excel/PDF sources.
    """
    from docx import Document

    # If it's a legacy .doc, convert first (the extractor would do this
    # anyway).  We import the helper from the extractor module to avoid
    # duplicating the LibreOffice subprocess logic.
    if isinstance(source, (str, Path)):
        suffix = Path(source).suffix.lower()
        if suffix == ".doc":
            from extractors.tp186_extract import _ensure_docx
            docx_path = _ensure_docx(Path(source))
            doc = Document(str(docx_path))
        else:
            doc = Document(str(source))
    else:
        # BytesIO — try opening as docx first; if it fails it's probably .doc
        pos = source.tell()
        try:
            doc = Document(source)
        except Exception:
            source.seek(pos)
            # Buffer to a temp .doc file and convert
            import tempfile
            from extractors.tp186_extract import _ensure_docx
            with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as tf:
                tf.write(source.read())
                tmp_path = Path(tf.name)
            source.seek(pos)
            docx_path = _ensure_docx(tmp_path)
            doc = Document(str(docx_path))

    # Concatenate paragraph text and table text
    parts = [p.text for p in doc.paragraphs]
    for tbl in doc.tables[:2]:                # first 2 tables = enough
        for row in tbl.rows:
            for cell in row.cells:
                parts.append(cell.text)
    blob = " || ".join(parts).upper()

    # RNSE-08 — ENTP rig 188, RNSE wells.  French workover report with a
    # "Déroulement des opérations" operations table and per-op tarif codes.
    # Distinguishing markers: the operations-table title or RNSE well /
    # TP 188 rig.
    if ("DÉROULEMENT DES OPÉRATIONS" in blob
            or "DEROULEMENT DES OPERATIONS" in blob
            or re.search(r"\bRNSE[-\s]?\d", blob)
            or "TP 188" in blob or "TP-188" in blob):
        return "rnse08"

    # TP-186 — ENTP rig 186, ZR wells, ZARZAITINE field, telex-style
    # Word .doc/.docx.  Distinguishing markers:
    #   - "RAPPORT JOURNALIER WORK-OVER" (hyphenated)
    #   - "TP # 186" or "TP-186" or "TP 186"
    #   - ZR# well prefix
    if ("TP # 186" in blob or "TP-186" in blob or "TP 186" in blob
            or re.search(r"\bZR\s*#\s*\d", blob)
            or ("RAPPORT JOURNALIER WORK-OVER" in blob and "ZARZAITINE" in blob)):
        return "tp186"

    return "unknown"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def _detect_format(source) -> str:
    """Single entry point for format detection.  Returns a rig key
    like 'enf', 'tp182', 'enf34_pdf', etc.  Returns 'unknown' if the file
    type or layout couldn't be identified."""
    kind = _sniff_kind(source)
    if kind == "pdf":
        return _detect_format_pdf(source)
    if kind == "xlsx":
        return _detect_format_xlsx(source)
    if kind in ("doc", "docx"):
        return _detect_format_word(source)
    return "unknown"


def parse_source(source: Union[Path, str, BytesIO]) -> dict:
    """Detect the source format and call the right extractor.

    Accepts both Excel (.xlsx) and PDF inputs.  Returns the standard
    dict shape produced by all extractors (header / activities /
    text_sections / mud_checks / mud_volume / mud_chemical_usage /
    personnel_data / pumps / well_location / survey_data / safety /
    tarif_totals) plus a "_meta" key with the source format and kind.
    """
    fmt = _detect_format(source)

    # Excel-backed extractors
    if fmt == "enf":
        from extractors.enf17_extract import parse_ddr
        data = parse_ddr(source)
    elif fmt == "enf33":
        from extractors.enf33_extract import parse_enf33
        data = parse_enf33(source)
    elif fmt == "tp179":
        from extractors.tp179_extract import parse_tp179
        data = parse_tp179(source)
    elif fmt == "tp173":
        from extractors.tp173_extract import parse_tp173
        data = parse_tp173(source)
    elif fmt == "tp183":
        from extractors.tp183_extract import parse_tp183
        data = parse_tp183(source)
    elif fmt == "tp182":
        from extractors.tp182_extract import parse_tp182
        data = parse_tp182(source)
    elif fmt == "tp195":
        from extractors.tp195_extract import parse_tp195
        data = parse_tp195(source)
    elif fmt == "entp204":
        from extractors.entp204_extract import parse_entp204
        data = parse_entp204(source)
    elif fmt == "entp127":
        from extractors.entp127_extract import parse_tp127
        data = parse_tp127(source)
    elif fmt == "gw29":
        from extractors.gw29_extract import parse_gw29
        data = parse_gw29(source)
    elif fmt == "enf04":
        from extractors.enf04_extract import parse_enf04
        data = parse_enf04(source)

    # PDF-backed extractors
    elif fmt == "enf34_pdf":
        from extractors.enf34_pdf_extract import parse_enf34_pdf
        data = parse_enf34_pdf(source)

    # Word-backed extractors (.doc auto-converted via LibreOffice → .docx)
    elif fmt == "tp186":
        from extractors.tp186_extract import parse_tp186
        data = parse_tp186(source)
    elif fmt == "rnse08":
        from extractors.rnse08_extract import parse_rnse08
        data = parse_rnse08(source)

    else:
        raise ValueError(
            f"Unrecognised report format. Markers in the file did not match "
            f"any known rig layout. Add a new extractor module and register "
            f"it in parse_source._detect_format_xlsx() or "
            f"_detect_format_pdf()."
        )

    data.setdefault("_meta", {})["source_format"] = fmt
    data["_meta"]["source_kind"] = _sniff_kind(source)

    # Universal bill-code normalization.  Different rig templates use
    # different bill formats — some have clean codes ("T1"), others use
    # multiplier prefixes ("1,05xT1", "0.95XT2").  The downstream insert
    # function's strict regex ^T(\d+)$ only matches the clean form, so
    # we normalize every activity's bill code here after the extractor
    # runs.  Empty / unrecognized values are left as "" rather than
    # silently mis-tagged.
    from helpers.bill_code_assign import normalize_bill_code
    for a in data.get("activities", []) or []:
        raw = a.get("bill", "")
        if raw:
            a["bill"] = normalize_bill_code(raw)

    return data


def main(argv=None) -> int:
    import argparse, json
    from datetime import date as date_type, datetime, time, timedelta

    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("source", type=Path)
    p.add_argument("--json", action="store_true", help="dump the extracted dict")
    args = p.parse_args(argv)

    if not args.source.exists():
        sys.exit(f"ERROR: source not found: {args.source}")

    fmt = _detect_format(args.source)
    print(f"Detected format: {fmt}", file=sys.stderr)

    if args.json:
        data = parse_source(args.source)
        def default(o):
            if isinstance(o, (date_type, datetime)): return o.isoformat()
            if isinstance(o, time): return o.strftime("%H:%M:%S")
            if isinstance(o, timedelta): return o.total_seconds()
            return str(o)
        print(json.dumps(data, indent=2, default=default))
    return 0


if __name__ == "__main__":
    sys.exit(main())
