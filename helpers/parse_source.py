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
    # Legacy OLE2 Compound File Binary — shared by .xls (Excel) and legacy
    # .doc (Word).  Disambiguate by extension; BytesIO (no extension) falls
    # back to "doc" since in-memory Word is the more common case here.
    if head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
        if isinstance(source, (str, Path)) and Path(source).suffix.lower() == ".xls":
            return "xls"
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

    # Scan the first 18 rows × 28 cols for marker strings.  18 (not 12) so
    # the rig-number cells of the Haoud Berkaoui workover templates
    # (ENF#10 / TP#215 / ENF#30) — whose "Appareil" / rig cell sits on row
    # 14 — are captured.  The remaining branches only match specific,
    # low-collision keywords, so the extra rows don't misroute other
    # formats.
    markers = []
    for row in ws.iter_rows(min_row=1, max_row=18, max_col=28, values_only=True):
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
    # (English, "DAILY DRILLING REPORT" title).  Some TP-195 reports use
    # "OFFICE REP" instead of "SUPERINTANDANT"; others (verified against a
    # real report) use "Superintandant" just like TP-182, which makes that
    # label alone useless for telling them apart.  The reliable signal is
    # the rig number itself ("ENTP 195" / "ENTP195" / "ENTP-195") or the
    # AIN T'SILA "AT-NN" well prefix — check these FIRST, before the
    # generic TP-182 fallback, and keep "OFFICE REP" as a secondary catch
    # for report variants that do use it instead of Superintandant.
    if (re.search(r"\bENTP[\s#\-]*195\b", blob)
            or (re.search(r"\bAT[\s\-]?\d{1,3}\b", blob) and "AIN T" in blob)
            or "OFFICE REP" in blob):
        return "tp195"

    # TP-182 — SONATRACH PRODUCTION DIVISION Daily Drilling Report format
    # (English, "SUPERINTANDANT" misspelling, has "WORKOVER REASON")
    if "SUPERINTANDANT" in blob:        # specific to TP-182 template
        return "tp182"
    if "SONATRACH PRODUCTION DIVISION" in blob and "DAILY DRILLING REPORT" in blob:
        return "tp182"

    # ENF#30 — Haoud Berkaoui workover format (same report family as
    # ENF#10 / TP#215, title "RAPPORT JOURNALIER DE WORKOVER").  Distinct
    # rig number "ENF#30" — must be checked BEFORE the generic ENF#04
    # catch-all below (which also fires on "HAOUD BERKAOUI" / "WORKOVER").
    if (re.search(r"\bENF\s*#?\s*30\b", blob)
            or "ENAFOR # 30" in blob or "ENAFOR#30" in blob
            or "ENF#30" in blob or "ENF #30" in blob or "ENF 30" in blob):
        return "enf30"

    # ENF#10 — Haoud Berkaoui workover format (rig "ENF#10", same template
    # as TP#215).  Distinct rig number — must be checked BEFORE the generic
    # ENF#04 catch-all below (which also fires on "HAOUD BERKAOUI").
    if (re.search(r"\bENF\s*#?\s*10\b", blob)
            or re.search(r"\bENAFOR\s*#?\s*10\b", blob)
            or "ENF#10" in blob or "ENF #10" in blob or "ENF 10" in blob):
        return "enf10"

    # TP#215 — Haoud Berkaoui workover format (rig "TP#215", same template
    # as ENF#10).  Distinct rig number — must be checked BEFORE the generic
    # ENF#04 catch-all below (which also fires on "HAOUD BERKAOUI").
    if (re.search(r"\bTP\s*#?\s*215\b", blob)
            or "TP#215" in blob or "TP #215" in blob or "TP 215" in blob):
        return "tp215"

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

    # ENTP-219 — ENTP rig 219, SONATRACH DP Hassi Messaoud daily work-over
    # report (modern .xlsx, single sheet "DWR N°1").  English-labeled header;
    # title "DAILY WORK-OVER REPORT" at J3 and a "WORK-OVER OPERATIONS"
    # table header.  Distinct from the French RAPPORT JOURNALIER workover
    # templates (TP-173/TP-179/TP-236) and from TP-185 (which is legacy .xls
    # and shares the SHDP family but uses "RAPPORT JOURNALIER WORK OVER").
    if ("ENTP 219" in blob or "ENTP-219" in blob or "ENTP219" in blob
            or "WORK-OVER OPERATIONS" in blob
            or "DAILY WORK-OVER REPORT" in blob):
        return "tp219"

    # ENF#18 — ENF rig 18 native .xlsx workover report. Distinctive banner
    # "DIVISION PRODUCTION / RAPPORT JOURNALIER WORK-OVER" and either a
    # specific rig identifier or ETP/ENF 18 reference in the header.
    if ("DIVISION PRODUCTION" in blob and "RAPPORT JOURNALIER WORK-OVER" in blob
            or re.search(r"\bENF\s*#?\s*18\b", blob)
            or re.search(r"\bENTP\s*#?\s*18\b", blob)):
        return "enf18"

    # ENTP-27 — some TP-189 files include an ENTP N°27 rig identifier in the
    # header (filename examples: "RAP ENTP N°27 TP189 ...").  These use the
    # same RAP ENTP layout as TP-189; detect and route them explicitly so
    # they don't fall through to unknown when the generic TP-189 markers
    # are slightly different.
    if re.search(r"\bENTP\s*(?:N°|N|#)?\s*27\b", blob) or "ENTP N\u00B027" in blob:
        return "entp27"

    # TP-189 — ENTP "RAP ENTP" wide single-sheet daily drilling report
    # (modern .xlsx, ~65 rows x 24 cols).  Filename pattern
    # RAP_ENTP_N_<report#>_<rig>_MD_<well>_DU_*.  Distinctive markers:
    # "RAP ENTP" title/banner, "DRILL STRING TALLY" section, "BHA COMPONENT
    # TALLY", and "inventaire tubulaire" (tubular inventory).  Must be
    # checked BEFORE any generic DDR catch-all since it shares some English
    # labels with other ENTP templates.
    if ("RAP ENTP" in blob
            or "DRILL STRING" in blob
            or "BHA COMPONENT" in blob
            or "inventaire tubulaire" in blob
            or "TP189" in blob
            or re.search(r"\bTP\s*-?\s*189\b", blob)):
        return "tp189"

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

    # TP-236 — ENTP rig 236, SONATRACH "RAPPORT JOURNALIER DU WORK OVER"
    # French/English workover format.  Shares the same title as TP-179 /
    # TP-173 but is distinguished by the rig number ("TP 236" / "TP-236" /
    # "ENTP 236") in the header.  Must be checked BEFORE the generic TP-179
    # catch-all below (which also matches "RAPPORT JOURNALIER" + "WORK").
    if ("TP 236" in blob or "TP-236" in blob
            or "ENTP 236" in blob or "ENTP-236" in blob
            or re.search(r"\bENTP\s?236\b", blob)):
        return "tp236"

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

    # ENF#24 — newer ENAFOR DDR template (rig ENF#24).  Same ENF DDR family as
    # ENF#17 / ENF#33 but with a few header cells shifted and a restructured
    # "Last Csg" block.  Distinguish by the rig number ("ENF#24" / "ENF #24"
    # / "ENF 24"); must be checked BEFORE the generic ENF catch-all below.
    if (re.search(r"ENF\s*#\s*24\b", blob)
            or "ENF#24" in blob or "ENF #24" in blob or "ENF 24" in blob):
        return "enf24"

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

    # ENF-03 — Excel-exported PDF report from ENF-03 / HR-152 / HRM.
    # Distinctive markers: "RAPPORT JOURNALIER" + "WORK OVER" +
    # "APPAREIL: ENF 03" or both "PUITS: HR" and "CHAMP: HRM".
    if ("RAPPORT JOURNALIER" in blob and "WORK OVER" in blob
            and (re.search(r"\bAPPAREIL:\s*ENF\s*03\b", blob)
                 or (re.search(r"\bPUITS[:\s]+HR\b", blob)
                     and "CHAMP: HRM" in blob))):
        return "enf03"

    # TP-212 — Excel-exported PDF report from TP-212 / HRZ-007.
    # Distinctive markers: "RAPPORT JOURNALIER" + "WORK OVER" +
    # "APPAREIL:TP 212" or "PUITS: HRZ".
    if ("RAPPORT JOURNALIER" in blob and "WORK OVER" in blob
            and ("APPAREIL:TP 212" in blob
                 or re.search(r"\bPUITS[:\s]+HRZ\b", blob))):
        return "tp212"

    # TP-217 — Excel-exported PDF report from TP-217 / ONRS-01 / HRM.
    # Distinctive markers: "RAPPORT JOURNALIER" + "WORK OVER" +
    # "APPAREIL:TP 217" or "PUITS: ONRS".
    if ("RAPPORT JOURNALIER" in blob and "WORK OVER" in blob
            and ("APPAREIL:TP 217" in blob
                 or re.search(r"\bPUITS[:\s]+ONRS\b", blob))):
        return "tp217"

    # TP-237 — Excel-exported PDF report from TP-237 / HR-128 / HRM.
    # Distinctive markers: "RAPPORT JOURNALIER" + "DTM" +
    # "APPAREIL: TP 237" or well/field markers like HR + HRM.
    if ("RAPPORT JOURNALIER" in blob and "DTM" in blob
            and (re.search(r"\bAPPAREIL:\s*TP\s*[- ]?\s*237\b", blob)
                 or (re.search(r"\bPUITS[:\s]+HR\b", blob)
                     and "CHAMP: HRM" in blob))):
        return "tp237"

    # ENF#34 — Gassi-Touil workover PDF (this is the first PDF format we
    # support).  Distinguishing markers: "GASSI" + "RAPPORT JOURNALIER DE
    # WORK OVER" (note the space in "WORK OVER" — TP-179 uses "WORKOVER").
    if "GASSI" in blob and "RAPPORT JOURNALIER" in blob and "WORK" in blob:
        return "enf34_pdf"
    if "DIRECTION RÉGIONALE GASSI" in blob or "GASSI-TOUIL" in blob:
        return "enf34_pdf"

    return "unknown"


# ---------------------------------------------------------------------------
# Legacy Excel (.xls) format detection
# ---------------------------------------------------------------------------
def _detect_format_xls(source) -> str:
    """Peek at a legacy .xls source (OLE2) and return the rig template key.

    .xls is the only spreadsheet flavour openpyxl can't read, so we open it
    with xlrd (imported lazily) and scan the first sheet's name plus its
    first 12 rows × 28 cols for distinguishing marker strings.
    """
    import xlrd

    if isinstance(source, (str, Path)):
        book = xlrd.open_workbook(str(source), formatting_info=False)
    else:
        pos = source.tell()
        raw = source.read()
        source.seek(pos)
        book = xlrd.open_workbook(file_contents=raw, formatting_info=False)

    sheet = book.sheet_by_index(0)
    markers = [sheet.name.upper()]
    for r in range(min(sheet.nrows, 12)):
        for c in range(min(sheet.ncols, 28)):
            v = sheet.cell_value(r, c)
            if v is not None and v != "":
                markers.append(str(v).upper())
    blob = " || ".join(markers)
    book.release_resources()

    # TP-185 — SONATRACH DP Hassi Messaoud daily workover report (.xls).
    # Title "RAPPORT JOURNALIER WORK OVER" + rig "TP 185" / "TP-185" /
    # "OMN"-prefixed well.  Must be checked before any generic "RAPPORT
    # JOURNALIER" workover catch-all, which would otherwise misroute it.
    if ("TP 185" in blob or "TP-185" in blob
            or "RAPPORT JOURNALIER WORK OVER" in blob
            or re.search(r"\bOMN[-\s]?\d", blob)):
        return "tp185"

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

    # ENF-08 — ENAFOR rig 08 Word report.  Same overall title as TP-186
    # but distinguished by the ENF rig number plus its ISNO/TINRHERT markers.
    if ("RAPPORT JOURNALIER WORK-OVER" in blob
            and (re.search(r"\bENF\s*#?\s*08\b", blob)
                 or "TINRHERT" in blob
                 or re.search(r"\bISNO[-\s]?\d+\b", blob))):
        return "enf08"

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
    if kind == "xls":
        return _detect_format_xls(source)
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
    if fmt == "enf30":
        from extractors.enf30_extract import parse_enf30
        data = parse_enf30(source)
    elif fmt == "enf10":
        from extractors.enf10_extract import parse_enf10
        data = parse_enf10(source)
    elif fmt == "tp215":
        from extractors.tp215_extract import parse_tp215
        data = parse_tp215(source)
    elif fmt == "enf24":
        from extractors.enf24_extract import parse_ddr
        data = parse_ddr(source)
    elif fmt == "enf":
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
    elif fmt == "tp236":
        from extractors.tp236_extract import parse_wo_report
        data = parse_wo_report(source)
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
    elif fmt == "tp219":
        from extractors.tp219_extract import parse_entp219
        data = parse_entp219(source)
    elif fmt == "enf18":
        from extractors.enf18_extract import parse_enf18_report
        data = parse_enf18_report(source)
    elif fmt == "entp27":
        from extractors.entp27_extract import parse_entp27
        data = parse_entp27(source)
    elif fmt == "tp189":
        from extractors.tp189_extract import parse_rap_entp
        data = parse_rap_entp(source)
    elif fmt == "gw29":
        from extractors.gw29_extract import parse_gw29
        data = parse_gw29(source)
    elif fmt == "enf04":
        from extractors.enf04_extract import parse_enf04
        data = parse_enf04(source)

    # Legacy .xls extractor (xlrd-backed)
    elif fmt == "tp185":
        from extractors.tp185_extract import parse_tp185
        data = parse_tp185(source)

    # PDF-backed extractors
    elif fmt == "enf34_pdf":
        from extractors.enf34_pdf_extract import parse_enf34_pdf
        data = parse_enf34_pdf(source)
    elif fmt == "tp212":
        from extractors.tp212_extract import parse_tp212
        data = parse_tp212(source)
    elif fmt == "tp217":
        from extractors.tp217_extract import parse_tp217
        data = parse_tp217(source)
    elif fmt == "tp237":
        from extractors.tp237_extract import parse_tp237
        data = parse_tp237(source)
    elif fmt == "enf03":
        from extractors.enf03_extract import parse_enf03
        data = parse_enf03(source)

    # Word-backed extractors (.doc auto-converted via LibreOffice → .docx)
    elif fmt == "tp186":
        from extractors.tp186_extract import parse_tp186
        data = parse_tp186(source)
    elif fmt == "enf08":
        from extractors.enf08_extract import parse_enf08
        data = parse_enf08(source)
    elif fmt == "rnse08":
        from extractors.rnse08_extract import parse_rnse08
        data = parse_rnse08(source)

    else:
        raise ValueError(
            f"Unrecognised report format. Markers in the file did not match "
            f"any known rig layout. Add a new extractor module and register "
            f"it in parse_source._detect_format_xlsx(), "
            f"_detect_format_xls(), or _detect_format_pdf()."
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