# BDAP Data Atlas

A browsable, searchable data dictionary for the BDAP repair-order data warehouse, built from the internal "BDAP Staging <> Warehouse Mapping" spreadsheet.

**[Open `index.html`](./index.html)** in any browser — no server or build step required. Everything (217 tables, 7,875 fields) is embedded in the page.

## What it covers

The source workbook has one summary index (staging source, source system, table type, and description for every warehouse table) plus one sheet per table with its field-level mapping (source field → target field, data type, mandatory flag, transformation rule, data-quality rules, notes). This project parses all of that into a single dataset and renders it as a navigable app:

- **Search** across table names, field names, and definitions
- **Filter** by table type — Fact, Dimension, Aggregate, Workday journals, Audit, Other/staging
- Per-table view with every field's data type, whether it's required, its business definition, and (expandable) its source lineage, transformation rule, and data-quality rules
- Shareable/bookmarkable links (`index.html#table/dim_customer`)

## Project layout

```
index.html                    the app (generated — do not hand-edit; regenerate instead)
data/data_dictionary.json     the same dataset as structured JSON, for reuse elsewhere
scripts/build_dictionary.py   parses the source .xlsx and regenerates the JSON + HTML
scripts/app_template.html     the app shell (HTML/CSS/JS) with a __DATA_JSON__ placeholder
```

## Regenerating from an updated mapping spreadsheet

The source `.xlsx` isn't checked into this repo (it's a large binary and changes independently of this project). To rebuild after the mapping spreadsheet is updated:

```bash
pip install openpyxl
python3 scripts/build_dictionary.py /path/to/BDAP_Staging_Warehouse_Mapping.xlsx
```

This overwrites `data/data_dictionary.json` and `index.html`. Commit both.

Useful flags:
- `--out-json` / `--out-html` — write elsewhere instead of the defaults
- `--skip-html` — only refresh the JSON

## Known gaps in the source data

The build script logs these to stderr on every run; they reflect gaps in the source spreadsheet, not the parser:

- 3 sheets have no field-mapping table it can recognize and are skipped: `it-analyst`, `CSI Insurance Mapping (US)`, `Insurance_names_mapping`.
- 12 tables listed in the summary index have no corresponding sheet in the workbook (e.g. `audit_job`, `fct_vehicle_scan_review`, `subaru_oe_extr`) and so show up without a source-system/description in the atlas.
- Many fields have no "Field definition" in the source sheet — the app shows these as "no definition provided" rather than guessing.
