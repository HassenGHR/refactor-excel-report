#!/usr/bin/env python3
"""
batch_to_router_excel.py — Batch process a folder of Excel reports, converting each
to a router-friendly Excel file compressed in a zip, outputting to an output folder.

For each .xlsx file in the input folder, runs the to_router_excel.py logic and saves
the resulting zip to the output folder with the same basename.

CLI:
    python batch_to_router_excel.py INPUT_FOLDER [-o OUTPUT_FOLDER]
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

# Import the processing functions
from parse_source import parse_source as parse_ddr
from to_router_excel import build_router_excel


def main(argv=None) -> int:
    p = argparse.ArgumentParser(
        description="Batch convert Excel reports in a folder to router-friendly zips "
                    "containing the optimized Excel files."
    )
    p.add_argument("input_folder", type=Path, help="Path to the folder containing .xlsx files")
    p.add_argument("-o", "--output", type=Path, default=Path.cwd() / "output",
                   help="Output folder for the .zip files (default: ./output)")
    args = p.parse_args(argv)

    if not args.input_folder.exists() or not args.input_folder.is_dir():
        sys.exit(f"ERROR: input folder not found or not a directory: {args.input_folder}")

    # Create output folder if it doesn't exist
    args.output.mkdir(parents=True, exist_ok=True)

    # Find all .xlsx files in the input folder
    xlsx_files = list(args.input_folder.glob("*.xlsx"))
    if not xlsx_files:
        sys.exit(f"ERROR: no .xlsx files found in {args.input_folder}")

    processed = 0
    for source in xlsx_files:
        # Output path: output_folder / source.stem.zip
        output_path = args.output / f"{source.stem}.zip"

        try:
            # 1. Extract data from source
            data = parse_ddr(source)

            # 2. Build the zip containing the Excel
            build_router_excel(data, output_path)

            print(f"Processed {source.name} -> {output_path}")
            processed += 1
        except Exception as e:
            print(f"ERROR processing {source.name}: {e}")

    print(f"\nBatch processing complete. Processed {processed} files.")
    return 0


if __name__ == "__main__":
    sys.exit(main())