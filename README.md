# Drilling / Workover Report Converter

Converts daily drilling and workover reports from multiple rig templates
(Excel + PDF) into a single router-ready Excel format that the existing
FastAPI `/v1/excel_import/upload` endpoint can ingest.

## Project structure

```
.
├── extractors/                 # Per-rig extraction modules
│   ├── __init__.py
│   ├── enf17_extract.py        # ENAFOR rig 17, English DDR
│   ├── enf04_extract.py        # ENAFOR rig 04, Haoud Berkaoui (2 layouts)
│   ├── enf34_pdf_extract.py    # ENAFOR rig 34, Gassi-Touil PDF
│   ├── gw29_extract.py         # GWDC rig 29, REB
│   ├── tp173_extract.py        # ENTP rig 173, ADRAR / ODZ wells
│   ├── tp179_extract.py        # ENTP rig 179, ADRAR / HTJW wells
│   ├── tp182_extract.py        # ENTP rig 182, SONATRACH / WIH wells
│   ├── tp183_extract.py        # ENTP rig 183, TMLS wells
│   └── tp195_extract.py        # ENTP rig 195, AIN T'SILA / AT wells
├── helpers/
│   ├── __init__.py
│   ├── bill_code_assign.py     # Bill-code partitioning + normalization
│   └── parse_source.py         # Format detection + dispatcher
├── to_router_excel.py          # Single-file CLI
├── batch_to_router_excel.py    # Batch processor (planned)
├── requirements.txt
└── README.md
```

## Install

```bash
python -m venv .venv
source .venv/bin/activate           # or `.venv\Scripts\activate` on Windows
pip install -r requirements.txt
```

## Usage

### Single file

```bash
python to_router_excel.py path/to/source.xlsx
python to_router_excel.py path/to/source.pdf
python to_router_excel.py source.xlsx --date 2026-05-16   # override date
python to_router_excel.py source.xlsx -o custom_output.xlsx
```

The dispatcher auto-detects which rig template the file is from based on
its content (magic bytes + scanned markers), so the same command works
for every supported rig.

### Programmatic

```python
from helpers.parse_source import parse_source

data = parse_source("source.xlsx")   # or source.pdf
print(data["_meta"]["source_format"])
print(data["header"]["well_name"])
for op in data["activities"]:
    print(op["start_time"], op["end_time"], op["bill"], op["description"])
```

## Supported formats

| Module               | Rig    | Template                        |
|----------------------|--------|----------------------------------|
| enf17_extract        | ENF#17 | DDR English (OpenWells)          |
| enf04_extract        | ENF#04 | Haoud Berkaoui — Layouts A and B |
| tp173_extract        | TP-173 | ADRAR, ODZ wells                 |
| tp179_extract        | TP-179 | ADRAR, HTJW wells                |
| tp182_extract        | TP-182 | SONATRACH PRODUCTION / WIH       |
| tp183_extract        | TP-183 | TMLS, single-sheet               |
| tp195_extract        | TP-195 | AIN T'SILA, OFFICE REP           |
| gw29_extract         | GW29   | GWDC REB                         |
| enf34_pdf_extract    | ENF#34 | Gassi-Touil PDF                  |

## Conventions

- **Two-supervisor rule.** Every rig has two 12-hour-shift supervisors;
  the DB columns `supervisor` and `superintendent` are slot 1 and slot 2.
  The "superintendent" column name is a misnomer — both slots hold
  supervisors.
- **Bill code normalization.** Templates with multiplier prefixes
  (`1,05xT1`, `0.95XT2`) are normalized to clean `T<n>` by
  `helpers.bill_code_assign.normalize_bill_code`, applied universally in
  `parse_source` after every extractor runs.
- **Bill code back-assignment.** For templates that don't tag individual
  operations (`ENF#04`, `ENF#34`, `TP-183`),
  `helpers.bill_code_assign.assign_bill_codes` partitions ops into groups
  whose hours sum exactly to each T-bucket total. Falls back to
  chronological best-effort when no exact partition exists.
- **Frontend column-name swap (do NOT "fix").** The DB column
  `lastCasingSHOE` renders as "Last LNR SHOE" on the frontend, and
  `lastCSNlnrSHOE` renders as "Last CSG SHOE". Extractors set
  `header.top_shoe` for the casing shoe (which flows to `lastCSNlnrSHOE`).
- **Situation cap.** `current_operation` / `day_summary` are capped at
  300 characters with a word-aware cut and ellipsis to fit the frontend
  display.
- **Date resolution order** in `to_router_excel.py`:
  `--date` CLI override → extractor's `header.date` → date in filename →
  today (last-resort fallback so uploads never fail with "no date").

## Adding a new rig template

1. Inspect the source's cell layout. Look for the distinguishing marker
   strings in the top 12 rows.
2. Copy the closest existing extractor in `extractors/` to
   `extractors/<rig>_extract.py` and adjust cell positions.
3. Add **one** branch to `helpers.parse_source._detect_format_xlsx` (or
   `_detect_format_pdf`) returning the rig key. Mind the ordering — more
   specific markers must come before more general ones (e.g. TP-195's
   `OFFICE REP` must check before TP-182's `SONATRACH PRODUCTION
   DIVISION`).
4. Add **one** dispatch elif to `helpers.parse_source.parse_source`.

No other code changes are needed. The router converter, parser-side
patch, and DB insert function don't care which rig the data came from.