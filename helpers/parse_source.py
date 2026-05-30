#!/usr/bin/env python3
"""
helpers/parse_source.py — detect Daily Workover Report format and dispatch to appropriate extractor.
"""
from __future__ import annotations

import re
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Union

from docx import Document
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet

# Import extractors - handle import errors gracefully for development
try:
    from ..extractors.enf08_extract import parse_enf08
except ImportError:
    try:
        from extractors.enf08_extract import parse_enf08
    except ImportError:
        try:
            from .enf08_extract import parse_enf08
        except ImportError:
            parse_enf08 = None  # type: ignore

try:
    from ..extractors.enf06_extract import parse_enf06
except ImportError:
    try:
        from extractors.enf06_extract import parse_enf06
    except ImportError:
        try:
            from .enf06_extract import parse_enf06
        except ImportError:
            parse_enf06 = None  # type: ignore

try:
    from ..extractors.enf27_extract import parse_enf27
except ImportError:
    try:
        from extractors.enf27_extract import parse_enf27
    except ImportError:
        try:
            from .enf27_extract import parse_enf27
        except ImportError:
            parse_enf27 = None  # type: ignore

try:
    from ..extractors.enf33_extract import parse_enf33
except ImportError:
    try:
        from extractors.enf33_extract import parse_enf33
    except ImportError:
        try:
            from .enf33_extract import parse_enf33
        except ImportError:
            parse_enf33 = None  # type: ignore

try:
    from ..extractors.tp182_extract import parse_tp182
except ImportError:
    try:
        from extractors.tp182_extract import parse_tp182
    except ImportError:
        try:
            from .tp182_extract import parse_tp182
        except ImportError:
            parse_tp182 = None  # type: ignore

try:
    from ..extractors.entp204_extract import parse_entp204
except ImportError:
    try:
        from extractors.entp204_extract import parse_entp204
    except ImportError:
        try:
            from .entp204_extract import parse_entp204
        except ImportError:
            parse_entp204 = None  # type: ignore

try:
    from ..extractors.rnse08_extract import parse_rnse08
except ImportError:
    try:
        from extractors.rnse08_extract import parse_rnse08
    except ImportError:
        try:
            from .rnse08_extract import parse_rnse08
        except ImportError:
            parse_rnse08 = None  # type: ignore

try:
    from ..extractors.enf34_pdf_extract import parse_enf34_pdf
except ImportError:
    try:
        from extractors.enf34_pdf_extract import parse_enf34_pdf
    except ImportError:
        try:
            from .enf34_pdf_extract import parse_enf34_pdf
        except ImportError:
            parse_enf34_pdf = None  # type: ignore

try:
    from ..extractors.tp186_extract import parse_tp186, _ensure_docx
except ImportError:
    try:
        from extractors.tp186_extract import parse_tp186, _ensure_docx
    except ImportError:
        try:
            from .tp186_extract import parse_tp186, _ensure_docx
        except ImportError:
            parse_tp186 = None  # type: ignore
            _ensure_docx = None  # type: ignore

try:
    from ..extractors.enf17_extract import parse_ddr as parse_enf17
except ImportError:
    try:
        from extractors.enf17_extract import parse_ddr as parse_enf17
    except ImportError:
        try:
            from .enf17_extract import parse_ddr as parse_enf17
        except ImportError:
            parse_enf17 = None  # type: ignore


def _clean(v) -> str:
    """Clean cell value: strip whitespace, replace multiple spaces."""
    if v is None:
        return ""
    s = str(v)
    if s.strip() in ("#VALUE!", "#REF!", "#NAME?", "#N/A", "#DIV/0!", "#NULL!"):
        return ""
    return re.sub(r"\s+", " ", s).strip()


def _cell(ws: Worksheet, row: int, col: int, lookup: dict) -> object:
    """Get cell value, accounting for merged cells."""
    v = ws.cell(row, col).value
    return v if v is not None else lookup.get((row, col))


def _build_merged_lookup(ws: Worksheet) -> dict:
    """Build lookup dictionary for merged cell values."""
    lookup = {}
    for mr in ws.merged_cells.ranges:
        av = ws.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                if (r, c) != (mr.min_row, mr.min_col):
                    lookup[(r, c)] = av
    return lookup


def _extract_doc_text(doc: Document) -> str:
    """Extract all text from a Word document for format detection."""
    parts = []
    parts.extend(p.text for p in doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return _clean(" ".join(parts)).upper()


def _extract_pdf_text(source: Union[Path, str, bytes]) -> str:
    """Extract all text from a PDF source for format detection."""
    try:
        from io import BytesIO
        import pdfplumber
    except ImportError as exc:
        raise ImportError(
            "pdfplumber is required to detect and parse PDF sources."
        ) from exc

    if isinstance(source, bytes):
        pdf_source = BytesIO(source)
    else:
        pdf_source = source

    pdf = pdfplumber.open(pdf_source)
    try:
        pages_text = [p.extract_text() or "" for p in pdf.pages]
    finally:
        pdf.close()

    return _clean(" ".join(pages_text)).upper()


def _detect_format_pdf(source: Union[Path, str, bytes]) -> str | None:
    """Detect the DWR format for a PDF document."""
    text = _extract_pdf_text(source)
    if ("GASSI-TOUIL" in text or "GASSI TOUIL" in text) and \
            "RAPPORT JOURNALIER DE WORK OVER" in text:
        return "enf34"
    return None


def _detect_format_word(doc: Document) -> str | None:
    """Detect the DWR format for a Word document."""
    text = _extract_doc_text(doc)
    if ("DÉROULEMENT DES OPÉRATIONS" in text or
            "DEROULEMENT DES OPERATIONS" in text):
        if "RNSE-" in text or "TP 188" in text or "TP-188" in text:
            return "rnse08"
    if "RAPPORT JOURNALIER WORK-OVER" in text or re.search(r"TP\s*#?\s*186", text):
        return "tp186"
    return None


def _detect_format_xlsx(ws: Worksheet) -> str | None:
    """
    Detect the DWR format based on worksheet content.
    Returns format identifier: 'enf06', 'enf18', 'enf17', 'enf27', 'enf33', 'tp182', 'entp204', or None.
    """
    L = _build_merged_lookup(ws)
    
    # Helper to get cleaned cell value
    def get_cell(r, c):
        return _clean(_cell(ws, r, c, L))

    def _make_search_text(max_rows=40, max_cols=20):
        parts = []
        for r in range(1, min(ws.max_row, max_rows) + 1):
            for c in range(1, min(ws.max_column, max_cols) + 1):
                parts.append(get_cell(r, c).upper())
        return " ".join(parts)

    sheet_text = _make_search_text()
    if "TXNO" in sheet_text or "AIN TSILA" in sheet_text or "AIN T'SILA" in sheet_text:
        return "entp204"
    if "WELL :" in sheet_text and "WIH" in sheet_text and "RIG NAME" in sheet_text and "TOTAL MD" in sheet_text:
        return "tp182"
    if ("RIG.S.I" in sheet_text or "ENF # 33" in sheet_text or "ENF#33" in sheet_text) and "DAILY DRILLING REPORT" in sheet_text:
        return "enf33"
    if re.search(r"ENF\s*#?\s*17\b", sheet_text, re.IGNORECASE) and "DAILY DRILLING REPORT" in sheet_text:
        return "enf17"

    # Scan key areas for distinguishing markers
    
    # Check for ENF#06 markers
    # - "E.NA.FOR" + "DAILY WORKOVER REPORT"
    # - "ENF# 06" or "ENF#06"
    # - "TP.Senior" + "TP.Junior" headers
    # - "Timing" + "MUD PUMPS DATA" ops headers
    
    # Look for E.NA.FOR pattern in title area (rows 2-5, cols A-I)
    title_found = False
    for r in range(1, 6):  # rows 1-5 (0-indexed as 1-5 in 1-based)
        for c in range(1, 10):  # cols A-I (1-9)
            val = get_cell(r, c)
            if "E.NA.FOR" in val.upper():
                # Check for DAILY WORKOVER REPORT nearby
                for r2 in range(max(1, r-2), min(ws.max_row+1, r+3)):
                    for c2 in range(max(1, c-2), min(ws.max_column+1, c+3)):
                        if "DAILY WORKOVER REPORT" in get_cell(r2, c2).upper():
                            title_found = True
                            break
                    if title_found:
                        break
                if title_found:
                    break
        if title_found:
            break
    
    if title_found:
        # Additional ENF#06 checks
        enf06_rig = False
        tp_senior_junior = False
        timing_mud_pumps = False
        
        # Check for ENF#06 rig designation
        for r in range(1, 10):
            for c in range(1, 10):
                val = get_cell(r, c)
                if re.search(r"ENF\s*#?\s*06", val, re.IGNORECASE):
                    enf06_rig = True
                    break
            if enf06_rig:
                break
        
        # Check for TP.Senior/TP.Junior in the top area; they may appear on
        # different rows in some ENF#06 variants.
        tp_senior_present = False
        tp_junior_present = False
        for r in range(1, 10):
            row_vals = [get_cell(r, c) for c in range(1, 10)]
            row_str = " ".join(row_vals).upper()
            if "TP.SENIOR" in row_str:
                tp_senior_present = True
            if "TP.JUNIOR" in row_str:
                tp_junior_present = True
            if tp_senior_present and tp_junior_present:
                tp_senior_junior = True
                break
        
        # Check for Timing + MUD PUMPS DATA
        for r in range(1, 15):
            row_vals = [get_cell(r, c) for c in range(1, 20)]
            row_str = " ".join(row_vals)
            if "TIMING" in row_str.upper() and "MUD PUMPS DATA" in row_str.upper():
                timing_mud_pumps = True
                break
        
        if enf06_rig and (tp_senior_junior or timing_mud_pumps):
            return "enf06"
    
    # Check for ENF#18 markers
    # - "ENF#18" or "ENF # 18" rig name
    # - well "MD " prefix + HMD field + "Rapport journalier work-over"
    
    enf18_rig = False
    md_hmd_workover = False
    
    # Check for ENF#18 rig designation
    for r in range(1, 15):
        for c in range(1, 20):
            val = get_cell(r, c)
            if re.search(r"ENF\s*#?\s*18", val, re.IGNORECASE):
                enf18_rig = True
                break
        if enf18_rig:
            break
    
    # Check for MD well + HMD field + work-over
    well_name = ""
    field_name = ""
    work_over_title = ""
    
    # Well name typically around B6-E6 area
    for r in [5, 6]:  # rows 5-6 (1-based 6-7)
        for c in [2, 3, 4, 5]:  # cols B-E (2-5)
            val = get_cell(r, c)
            if val and len(val) > 3 and not val.isdigit():
                if not well_name:
                    well_name = val
    
    # Field name typically around E6 area
    for r in [5, 6]:
        for c in [4, 5, 6]:  # cols D-F (4-6)
            val = get_cell(r, c)
            if val and ("HMD" in val.upper() or "HASSI" in val.upper()):
                field_name = val
    
    # Work-over title
    for r in [2, 3, 4]:
        for c in [1, 2, 3, 4, 5, 6, 7]:  # cols A-G
            val = get_cell(r, c)
            if "RAPPORT" in val.upper() and ("WORK" in val.upper() or "OVER" in val.upper()):
                work_over_title = val
    
    if well_name and field_name and work_over_title:
        if "MD" in well_name.upper() and "HMD" in field_name.upper():
            md_hmd_workover = True
    
    if enf18_rig or md_hmd_workover:
        return "enf18"
    
    # Check for ENF#27 markers
    # - "ENF 27" or "ENF#27" rig name
    # - "DIRECTION REGIONALE OHANET"
    # - well "DIMW-" prefix
    
    enf27_rig = False
    ohANET_direction = False
    dimw_well = False
    
    # Check for ENF#27 rig designation
    for r in range(1, 15):
        for c in range(1, 20):
            val = get_cell(r, c)
            if re.search(r"ENF\s*#?\s*27", val, re.IGNORECASE):
                enf27_rig = True
                break
        if enf27_rig:
            break
    
    # Check for DIRECTION REGIONALE OHANET
    for r in range(1, 10):
        for c in range(1, 15):
            val = get_cell(r, c)
            if "DIRECTION" in val.upper() and "REGIONALE" in val.upper() and "OHANET" in val.upper():
                ohANET_direction = True
                break
        if ohANET_direction:
            break
    
    # Check for DIMW- well prefix
    well_name_27 = ""
    for r in [4, 5]:  # rows 4-5 (1-based 5-6)
        for c in [1, 2, 3]:  # cols A-C (1-3)
            val = get_cell(r, c)
            if val and ("DIMW" in val.upper() or val.startswith("DIMW-")):
                dimw_well = True
                break
        if dimw_well:
            break
    
    if enf27_rig and (ohANET_direction or dimw_well):
        return "enf27"
    
    return None


def parse_source(source: Union[Path, str, bytes]) -> dict:
    """
    Main entry point: detect format and dispatch to appropriate extractor.
    
    Args:
        source: File path, string, or bytes of the Excel or Word file
        
    Returns:
        dict: Extracted data in standard format
        
    Raises:
        ValueError: If format cannot be detected or extractor not available
    """
    source_path = None
    format_id = None
    doc = None
    wb = None

    if isinstance(source, bytes):
        from io import BytesIO
        stream = BytesIO(source)
        try:
            wb = load_workbook(stream, data_only=True)
            ws = wb.active
            format_id = _detect_format_xlsx(ws)
        except InvalidFileException:
            stream.seek(0)
            try:
                format_id = _detect_format_pdf(stream)
            except ImportError:
                format_id = None
            if format_id is None:
                stream.seek(0)
                try:
                    doc = Document(stream)
                except Exception:
                    stream.seek(0)
                    with tempfile.NamedTemporaryFile(suffix='.doc', delete=False) as tf:
                        tf.write(stream.read())
                        tmp_path = Path(tf.name)
                    if _ensure_docx is None:
                        raise ValueError(
                            "Cannot parse Word document bytes: LibreOffice conversion is unavailable."
                        )
                    doc = Document(str(_ensure_docx(tmp_path)))
                format_id = _detect_format_word(doc)
    else:
        source_path = Path(source)
        suffix = source_path.suffix.lower()
        if suffix in ('.doc', '.docx'):
            if suffix == '.doc':
                if _ensure_docx is None:
                    raise ValueError(
                        "Cannot parse .doc file: LibreOffice conversion is unavailable."
                    )
                doc = Document(str(_ensure_docx(source_path)))
            else:
                doc = Document(str(source_path))
            format_id = _detect_format_word(doc)
        elif suffix == '.pdf':
            try:
                format_id = _detect_format_pdf(source_path)
            except ImportError as exc:
                raise ValueError(
                    "Cannot parse PDF file: pdfplumber is required."
                ) from exc
        else:
            if suffix not in ('.xlsx', '.xls', '.xlsm', '.xltx', '.xltm'):
                raise ValueError(
                    f"Unsupported source file type: {suffix!r}. "
                    "Supported formats: .xlsx, .xlsm, .xltx, .xltm, .doc, .docx, .pdf"
                )
            try:
                wb = load_workbook(source_path, data_only=True)
            except InvalidFileException as exc:
                raise ValueError(
                    f"Unsupported source file type: {suffix!r}. "
                    "Supported formats: .xlsx, .xlsm, .xltx, .xltm, .doc, .docx, .pdf"
                ) from exc
            ws = wb.active
            format_id = _detect_format_xlsx(ws)
    
    # Dispatch to appropriate extractor
    if format_id == "enf06" and parse_enf06 is not None:
        result = parse_enf06(source)
    elif format_id == "enf18" and parse_enf08 is not None:  # enf08 handles ENF#18
        result = parse_enf08(source)
    elif format_id == "enf27" and parse_enf27 is not None:
        result = parse_enf27(source)
    elif format_id == "enf33" and parse_enf33 is not None:
        result = parse_enf33(source)
    elif format_id == "enf17" and parse_enf17 is not None:
        result = parse_enf17(source)
    elif format_id == "tp182" and parse_tp182 is not None:
        result = parse_tp182(source)
    elif format_id == "entp204" and parse_entp204 is not None:
        result = parse_entp204(source)
    elif format_id == "enf34" and parse_enf34_pdf is not None:
        result = parse_enf34_pdf(source)
    elif format_id == "rnse08" and parse_rnse08 is not None:
        result = parse_rnse08(source)
    elif format_id == "tp186" and parse_tp186 is not None:
        result = parse_tp186(source)
    else:
        if isinstance(source, bytes):
            wb = None
        else:
            wb = load_workbook(source_path, data_only=True) if source_path.suffix.lower() in ('.xlsx', '.xls', '.xlsm', '.xltx', '.xltm') else None
        available = []
        if parse_enf06 is not None:
            available.append("enf06")
        if parse_enf08 is not None:
            available.append("enf18")
        if parse_enf27 is not None:
            available.append("enf27")
        if parse_enf33 is not None:
            available.append("enf33")
        if parse_enf17 is not None:
            available.append("enf17")
        if parse_tp182 is not None:
            available.append("tp182")
        if parse_entp204 is not None:
            available.append("entp204")
        if parse_enf34_pdf is not None:
            available.append("enf34")
        if parse_rnse08 is not None:
            available.append("rnse08")
        if parse_tp186 is not None:
            available.append("tp186")
        
        if format_id is None:
            raise ValueError(
                f"Could not detect DWR format. Available extractors: {available}"
            )
        else:
            raise ValueError(
                f"Format '{format_id}' detected but no extractor available. "
                f"Available extractors: {available}"
            )
    if 'wb' in locals() and wb is not None:
        wb.close()
    return result


# Drop-in compat for direct Excel parsing
parse_daily_excel_report = parse_source


if __name__ == "__main__":
    import sys
    import json
    
    if len(sys.argv) < 2:
        print("Usage: python parse_source.py SOURCE.xlsx")
        sys.exit(1)
    
    try:
        data = parse_source(Path(sys.argv[1]))
        print(json.dumps(data, indent=2, default=str, ensure_ascii=False))
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)