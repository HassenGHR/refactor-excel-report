# Refactor Excel Report

This project contains scripts to process and refactor Excel reports for daily drilling operations. The scripts extract data from source Excel files and generate a new Excel file in a standardized layout optimized for maximum data extraction by the existing excel_import_service.py (TP parser path), without needing any changes to the FastAPI route.

The output is a ZIP file containing the refactored Excel report, which can then be uploaded to the API endpoint for full data parsing.

## Usage

To convert a single Excel file, run:

```bash
python to_router_excel.py report.xlsx
```

To process multiple Excel files, place them in the `reports` folder and run:

```bash
python batch_to_router_excel.py reports -o outputs
```

Each script outputs a ZIP file containing the refactored Excel report.

## Supported Report Formats

The scripts support various Excel report formats from different rigs:

- TP-179 (HTJW well, ENTP rig)
- ENF04
- GW29
- TP-173
- TP-182
- TP-195
- ENF-17

## Requirements

See requirements.txt for dependencies.