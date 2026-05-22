#!/usr/bin/env python3
"""
batch_to_router_excel.py — process every .xlsx / .pdf in a directory
through the same pipeline as to_router_excel.py.

Usage:
    python batch_to_router_excel.py INPUT_DIR
    python batch_to_router_excel.py INPUT_DIR -o OUTPUT_DIR
    python batch_to_router_excel.py INPUT_DIR --recursive
    python batch_to_router_excel.py INPUT_DIR --pattern "*.pdf"   # PDFs only

Each input file is routed through helpers.parse_source.parse_source which
auto-detects its rig template. Files whose format isn't recognized are
skipped with a warning rather than aborting the batch.

All successfully generated router Excel files are collected into a single
ZIP archive (router_batch_YYYY-MM-DD_HHMM.zip) in the output directory.
The archive contains the individual .xlsx files (not further zipped).
"""
from __future__ import annotations
import argparse
import sys
import zipfile
from pathlib import Path
from typing import List, Tuple

from to_router_excel import build_router_excel, _ensure_date, parse_ddr, make_router_excel_bytes
from datetime import datetime, date as date_type


SUPPORTED_SUFFIXES = (".xlsx", ".xls", ".pdf")


def _iter_sources(input_dir: Path, recursive: bool, pattern: str) -> List[Path]:
    """Find all candidate source files in input_dir."""
    if recursive:
        matches = sorted(input_dir.rglob(pattern))
    else:
        matches = sorted(input_dir.glob(pattern))
    # Filter to supported file types (router_ready outputs are skipped)
    return [
        p for p in matches
        if p.is_file()
        and p.suffix.lower() in SUPPORTED_SUFFIXES
        and "_router" not in p.stem.lower()        # skip our own outputs
    ]


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
        description="Batch-convert a directory of rig reports to "
                    "router-ready Excel files, delivered as one ZIP archive "
                    "containing all the .xlsx files."
    )
    p.add_argument("input_dir", type=Path,
                    help="Directory containing source .xlsx/.pdf files")
    p.add_argument("-o", "--output-dir", type=Path, default=None,
                    help="Directory where the final ZIP archive will be written "
                         "(default: same as input_dir)")
    p.add_argument("-r", "--recursive", action="store_true",
                   help="Recurse into subdirectories")
    p.add_argument("--pattern", default="*",
                   help="Glob pattern to match (default: *)")
    args = p.parse_args(argv)

    if not args.input_dir.exists() or not args.input_dir.is_dir():
        sys.exit(f"ERROR: input dir not found: {args.input_dir}")

    output_dir = args.output_dir or args.input_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    sources = _iter_sources(args.input_dir, args.recursive, args.pattern)
    if not sources:
        print(f"No source files found in {args.input_dir} (pattern {args.pattern!r}).")
        return 0

    print(f"Found {len(sources)} file(s); will create ZIP archive in {output_dir}")
    print()

    ok = skipped = errored = 0
    results = []
    for src in sources:
        result = _process_one(src)
        results.append(result)
        if result["status"] == "ok":
            ok += 1
            print(f"  ✓  {src.name}  →  {result['inner']}  "
                  f"[{result['format']}, {result['activities']} ops, {result['date']}]")
        elif result["status"] == "skipped":
            skipped += 1
            print(f"  -  {src.name}: skipped — {result['error']}")
        else:
            errored += 1
            print(f"  ✗  {src.name}: ERROR — {result['error']}")

    # Package all successful .xlsx into one ZIP archive
    goods: List[Tuple[str, bytes]] = [
        (r["inner"], r["bytes"]) for r in results
        if r["status"] == "ok" and "bytes" in r
    ]
    if goods:
        ts = datetime.now().strftime("%Y-%m-%d_%H%M")
        zip_name = f"router_batch_{ts}.zip"
        zip_path = output_dir / zip_name
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for inner, content in goods:
                zf.writestr(inner, content)
        print()
        print(f"Created ZIP archive: {zip_path}  (contains {len(goods)} Excel file(s))")

    print()
    print(f"Summary: {ok} succeeded, {skipped} skipped, {errored} errored "
          f"(of {len(sources)} total)")
    return 0 if errored == 0 else 1


if __name__ == "__main__":
    sys.exit(main())