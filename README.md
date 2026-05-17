# Refactor Excel Report

This project contains scripts to process and refactor Excel and PDF reports for daily drilling operations. The scripts extract data from source Excel (.xlsx) or PDF files and generate a new Excel file in a standardized layout optimized for maximum data extraction by the existing excel_import_service.py (TP parser path), without needing any changes to the FastAPI route.

The output is a ZIP file containing the refactored Excel report, which can then be uploaded to the API endpoint for full data parsing.

## Usage

To convert a single Excel or PDF file, run:

```bash
python to_router_excel.py report.xlsx
# or
python to_router_excel.py report.pdf
```

To process multiple Excel files, place them in the `reports` folder and run:

```bash
python batch_to_router_excel.py reports -o outputs
```

Each script outputs a ZIP file containing the refactored Excel report.

## Supported Report Formats

The scripts support various Excel and PDF report formats from different rigs:

- TP-179 (HTJW well, ENTP rig)
- ENF04
- GW29
- TP-173
- TP-182
- TP-183
- TP-195
- ENF-17
- ENF34 PDF

## Project Structure

```
.
├── extractors/                 # Extraction scripts for different report formats
│   ├── __init__.py
│   ├── enf17_extract.py          # ENF17 format extraction
│   ├── enf04_extract.py        # ENF04 format extraction
│   ├── enf34_pdf_extract.py    # ENF34 PDF format extraction
│   ├── gw29_extract.py         # GW29 format extraction
│   ├── tp173_extract.py        # TP-173 format extraction
│   ├── tp179_extract.py        # TP-179 format extraction
│   ├── tp182_extract.py        # TP-182 format extraction
│   ├── tp183_extract.py        # TP-183 format extraction
│   └── tp195_extract.py        # TP-195 format extraction
├── helpers/                    # Helper modules
│   ├── bill_code_assign.py     # Bill code assignment logic
│   └── parse_source.py         # Format detection and dispatcher
├── batch_to_router_excel.py    # Batch processing multiple files
├── to_router_excel.py          # Single file processing
├── .gitignore
├── README.md
└── requirements.txt
```

## Requirements

See requirements.txt for dependencies.