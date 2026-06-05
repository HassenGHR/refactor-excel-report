# Drilling / Workover Report Converter

Converts daily drilling and workover reports from multiple rig templates
(Excel + PDF) into a single router-ready Excel format that the existing
FastAPI `/v1/excel_import/upload` endpoint can ingest.

## Project structure

```
.
├── extractors/                 # Per-rig extraction modules
│   ├── __init__.py
│   ├── enf06_extract.py        # ENAFOR rig 06, Daily Workover Report
│   ├── enf08_extract.py        # ENAFOR rig 08, Daily Workover Report
│   ├── enf17_extract.py        # ENAFOR rig 17, English DDR
│   ├── enf27_extract.py        # ENAFOR rig 27, Direction Régionale OHANET
│   ├── enf33_extract.py        # ENAFOR rig 33, BERKINE/BKNS DDR (.xlsm)
│   ├── enf04_extract.py        # ENAFOR rig 04, Haoud Berkaoui (2 layouts)
│   ├── enf34_pdf_extract.py    # ENAFOR rig 34, Gassi-Touil PDF
│   ├── entp127_extract.py      # ENTP DF rig 127, DAD wells, English template
│   ├── gw29_extract.py         # GW-29 / TP-127 layout family — French template (GWDC REB + DRAA DAOUI)
│   ├── tp173_extract.py        # ENTP rig 173, ADRAR / ODZ wells
│   ├── tp179_extract.py        # ENTP rig 179, ADRAR / HTJW wells
│   ├── tp182_extract.py        # ENTP rig 182, SONATRACH / WIH wells
│   ├── tp183_extract.py        # ENTP rig 183, TMLS wells
│   ├── tp186_extract.py        # ENTP rig 186, ZARZAITINE telex (Word .doc/.docx)
│   ├── rnse08_extract.py       # ENTP rig 188, RNSE wells (Word .doc/.docx)
│   ├── tp195_extract.py        # ENTP rig 195, AIN T'SILA / AT wells
│   └── entp204_extract.py      # ENTP rig 204, AIN T'SILA / TXNO wells
├── helpers/
│   ├── __init__.py
│   ├── bill_code_assign.py     # Bill-code partitioning + normalization
│   └── parse_source.py         # Magic-byte sniffing, format detection, dispatcher
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

### Batch processing

```bash
# Process a mix of sources and write a ZIP per source next to the input file
python batch_to_router_excel.py report.xlsx report.pdf report.docx

# Process all supported files in a directory recursively
python batch_to_router_excel.py reports --recursive

# Write every output ZIP into a dedicated directory
python batch_to_router_excel.py reports -o outputs
```

Each successful source produces its own ZIP archive containing a single
`*_router.xlsx` file. Unsupported or unrecognized inputs are skipped with a
warning, while errors are reported per file.

### Notes

- Supported input formats: `.xlsx`, `.xls`, `.pdf`, `.doc`, `.docx`
- If `--output-dir` is omitted, each ZIP is written next to its source file.

## Supported formats

| Module               | Rig    | Source type | Template                        |
|----------------------|--------|-------------|----------------------------------|
| enf04_extract        | ENF#04 | .xlsx       | Haoud Berkaoui — Layouts A and B |
| enf06_extract        | ENF#06 | .xlsx       | Daily Workover Report            |
| enf08_extract        | ENF#08 | .xlsx       | Daily Workover Report            |
| enf17_extract        | ENF#17 | .xlsx       | DDR English (OpenWells)          |
| enf27_extract        | ENF#27 | .xlsx       | DIRECTION REGIONALE OHANET       |
| enf33_extract        | ENF#33 | .xlsx/.xlsm | DDR English, BERKINE/BKNS wells  |
| tp173_extract        | TP-173 | .xlsx       | ADRAR, ODZ wells                 |
| tp179_extract        | TP-179 | .xlsx       | ADRAR, HTJW wells                |
| tp182_extract        | TP-182 | .xlsx       | SONATRACH PRODUCTION / WIH       |
| tp183_extract        | TP-183 | .xlsx       | TMLS, single-sheet               |
| tp186_extract        | TP-186 | .doc, .docx | ZR wells, ZARZAITINE telex       |
| rnse08_extract       | TP-188 | .doc, .docx | RNSE wells, ops-table format     |
| tp195_extract        | TP-195 | .xlsx       | AIN T'SILA, OFFICE REP           |
| entp204_extract      | ENTP-204 | .xlsx      | AIN T'SILA, TXNO, LABEL:value    |
| entp127_extract      | ENTP TP-127 | .xlsx   | DAD wells, DRAA DAOUI, English   |
| gw29_extract         | GW29 / TP-127 | .xlsx  | GWDC REB · TP-127 DRAA DAOUI (French) |
| enf34_pdf_extract    | ENF#34 | .pdf        | Gassi-Touil PDF                  |

### Word .doc support

Legacy Word `.doc` files (Word 97-2003 binary format) are auto-converted
to `.docx` using **LibreOffice headless** before parsing.  This requires
LibreOffice on the deployment host:

```bash
sudo apt-get install libreoffice          # Debian/Ubuntu
brew install --cask libreoffice           # macOS
```

`.docx` files are read directly via python-docx — no conversion needed.

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