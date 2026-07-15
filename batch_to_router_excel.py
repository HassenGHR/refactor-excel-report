#!/usr/bin/env python3
"""
batch_to_router_excel.py — process a list of report files or directories
through the same pipeline as to_router_excel.py.

Usage:
    python batch_to_router_excel.py REPORT1.xlsx REPORT2.pdf report.docx
    python batch_to_router_excel.py INPUT_DIR -o OUTPUT_DIR
    python batch_to_router_excel.py INPUT_DIR --recursive
    python batch_to_router_excel.py "*.xlsx" "*.pdf" "*.xls"

Supports .xlsx/.xlsm/.xls (Excel, including legacy .xls via the TP-185
extractor), .pdf, and Word .doc/.docx sources. Each input file is routed
through helpers.parse_source.parse_source which auto-detects its rig
template. Files whose format isn't recognized are skipped with a warning
rather than aborting the batch.

For each successful source, a separate ZIP archive is created containing
its single router-facing .xlsx output.
"""
from __future__ import annotations
import argparse
import sys
import zipfile
from pathlib import Path
from typing import List, Tuple

from to_router_excel import _ensure_date, parse_ddr, make_router_excel_bytes
from datetime import datetime, date as date_type


SUPPORTED_SUFFIXES = (".xlsx", ".xlsm", ".xls", ".pdf", ".doc", ".docx")


def _iter_sources(inputs: List[Path], recursive: bool, pattern: str) -> List[Path]:
    """Find all candidate source files from files and directories."""
    sources: List[Path] = []
    for entry in inputs:
        if entry.is_dir():
            if recursive:
                matches = sorted(entry.rglob(pattern))
            else:
                matches = sorted(entry.glob(pattern))
            sources.extend(
                p for p in matches
                if p.is_file()
                and p.suffix.lower() in SUPPORTED_SUFFIXES
                and "_router" not in p.stem.lower()
            )
        elif entry.is_file():
            if (entry.suffix.lower() in SUPPORTED_SUFFIXES
                    and "_router" not in entry.stem.lower()):
                sources.append(entry)
        else:
            print(f"Warning: input path not found or unsupported: {entry}")
    return sources


def _process_one(source: Path) -> dict:
    """Convert a single source file. Returns a result record (with bytes on success)."""
    try:
        data = parse_ddr(source)
        _ensure_date(data, source)

        rig = (data["header"].get("rig_name") or "rig").replace("/", "_")
        rd = data["header"].get("date")
        date_str = (rd.strftime("%Y-%m-%d")
                    if isinstance(rd, (datetime, date_type)) else "out")
        inner_name = f"{rig}_{date_str}_router.xlsx"
        xlsx_bytes = make_router_excel_bytes(data)

        return {
            "source": str(source),
            "status": "ok",
            "format": data.get("_meta", {}).get("source_format"),
            "inner": inner_name,
            "bytes": xlsx_bytes,
            "activities": len(data["activities"]),
            "date": str(rd) if rd else None,
        }
    except ValueError as e:
        # Unrecognised format — skip but don't fail the batch
        return {"source": str(source), "status": "skipped", "error": str(e)}
    except Exception as e:
        return {"source": str(source), "status": "error", "error": str(e)}


def main(argv=None) -> int:
    p = argparse.ArgumentParser(
        description="Batch-convert report files to router-ready Excel files, "
                    "with each output packaged into its own ZIP archive."
    )
    p.add_argument("inputs", type=Path, nargs="+",
                    help="Input files or directories containing report sources")
    p.add_argument("-o", "--output-dir", type=Path, default=None,
                    help="Directory where individual output ZIP archives will be written. "
                         "If omitted, each ZIP is written next to its source file.")
    p.add_argument("-r", "--recursive", action="store_true",
                   help="Recurse into subdirectories")
    p.add_argument("--pattern", default="*",
                   help="Glob pattern to match (default: *)")
    args = p.parse_args(argv)

    missing = [p for p in args.inputs if not p.exists()]
    if missing:
        sys.exit(f"ERROR: input path(s) not found: {', '.join(str(p) for p in missing)}")

    if args.output_dir:
        output_dir = args.output_dir
        output_dir.mkdir(parents=True, exist_ok=True)

    sources = _iter_sources(args.inputs, args.recursive, args.pattern)
    if not sources:
        print(f"No source files found. Checked inputs: {', '.join(str(p) for p in args.inputs)}")
        return 0

    if args.output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)
        print(f"Found {len(sources)} file(s); will write ZIP archives to {output_dir}")
    else:
        print(f"Found {len(sources)} file(s); will write ZIP archives next to each source file")
    print()

    ok = skipped = errored = 0
    for src in sources:
        result = _process_one(src)
        if result["status"] == "ok":
            ok += 1
            output_target_dir = args.output_dir or src.parent
            output_target_dir.mkdir(parents=True, exist_ok=True)
            date_str = result.get("date") or "out"
            zip_name = f"{src.stem}_{date_str}_router.zip"
            zip_path = output_target_dir / zip_name
            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(result["inner"], result["bytes"])

            print(f"  ✓  {src.name}  →  {zip_path.name}  "
                  f"[{result['format']}, {result['activities']} ops, {result['date']}]")
        elif result["status"] == "skipped":
            skipped += 1
            print(f"  -  {src.name}: skipped — {result['error']}")
        else:
            errored += 1
            print(f"  ✗  {src.name}: ERROR — {result['error']}")

    print()
    print(f"Summary: {ok} succeeded, {skipped} skipped, {errored} errored "
          f"(of {len(sources)} total)")
    return 0 if errored == 0 else 1


if __name__ == "__main__":
    sys.exit(main())