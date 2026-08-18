# BDAP Data Atlas

A browsable, searchable data dictionary for BDAP's data platform, built from two internal mapping spreadsheets: the "Staging <> Warehouse Mapping" spec and the "Source <> Data Lake Mapping" spec.

**[Open `index.html`](./index.html)** in any browser — no server or build step required. Everything (376 tables, 12,624 fields across both datasets) is embedded in the page.

## What it covers

The app has two tabs, one per dataset, each built from its own workbook:

- **Data Warehouse** — 217 fact/dimension/aggregate tables, from the staging-to-warehouse mapping spec.
- **Data Lake** — 159 raw ingestion tables/feeds, from the source-to-data-lake mapping spec. Many raw feeds are catalogued (source system, feed type, format, frequency, S3 path) without a field-by-field mapping yet; those still show up, clearly marked as "no field mapping yet," rather than being dropped.

Each source workbook has one summary/index sheet describing every table, plus one sheet per table with its field-level mapping (source field → target field, data type, mandatory flag, transformation rule, notes, etc). The app renders all of it as:

- **Search** across table names, field names, and definitions, scoped to whichever tab is open
- **Filter** by category — table type for the warehouse (Fact, Dimension, Aggregate, …), source system for the data lake (CCC Cloud Connect, Workday, Hertz, …)
- Per-table view with every field's data type, whether it's required, its business definition, and (expandable) its source lineage, transformation rule, and data-quality rules
- Shareable/bookmarkable links, e.g. `index.html#warehouse/table/dim_customer` or `index.html#datalake/table/repair_order`

## Project layout

```
index.html                          the app (generated — do not hand-edit; regenerate instead)
data/data_dictionary.json           Data Warehouse dataset as structured JSON
data/data_lake_dictionary.json      Data Lake dataset as structured JSON
scripts/mapping_common.py           shared header-matching / field-parsing logic
scripts/build_dictionary.py         parses the warehouse mapping .xlsx -> data/data_dictionary.json
scripts/build_datalake_dictionary.py parses the data lake mapping .xlsx -> data/data_lake_dictionary.json
scripts/build_app.py                combines the dataset JSON files + app_template.html -> index.html
scripts/app_template.html           the app shell (HTML/CSS/JS) with a __DATA_JSON__ placeholder
```

## Regenerating from updated mapping spreadsheets

The source `.xlsx` files aren't checked into this repo (they're large binaries that change independently of this project). To rebuild after either mapping spreadsheet is updated:

```bash
pip install openpyxl
python3 scripts/build_dictionary.py /path/to/BDAP_Staging_Warehouse_Mapping.xlsx
python3 scripts/build_datalake_dictionary.py /path/to/BDAP_Source_DataLake_Mapping.xlsx
python3 scripts/build_app.py
```

Run only the first (or second) script if just one dataset changed, then always re-run `build_app.py` to re-embed both datasets into `index.html`. Commit whichever JSON file(s) changed plus `index.html`.

`build_app.py` takes dataset JSON paths as optional positional args if you want to build from a subset or from files at non-default paths; run any script with `--help` for its full option list.

## Adding a third dataset later

1. Write a `scripts/build_<name>_dictionary.py` that parses the new workbook and writes `data/<name>_dictionary.json` with the shape `{id, label, tables: {...}}` (use `mapping_common.py`'s `parse_detail_sheets` / `build_fields_and_rules` for the per-sheet parsing; each table needs at least `id`, `display_name`, `category`, `field_count`, `fields`).
2. Add that JSON path to `DEFAULT_DATASETS` in `scripts/build_app.py`.
3. Optionally add a `DATASET_UI` entry for it in `scripts/app_template.html` (hero title/lead/how-to copy and stat cards) — without one it falls back to generic copy, so this step is cosmetic only.
4. Re-run `build_app.py`.

## Known gaps in the source data

The build scripts log these to stderr on every run; they reflect gaps in the source spreadsheets, not the parser:

**Data Warehouse** — 3 sheets have no field-mapping table it can recognize and are skipped: `it-analyst`, `CSI Insurance Mapping (US)`, `Insurance_names_mapping`. 12 tables listed in the summary index have no corresponding sheet in the workbook (e.g. `audit_job`, `fct_vehicle_scan_review`, `subaru_oe_extr`) and so show up without a source-system/description in the atlas.

**Data Lake** — 4 sheets are skipped as not real field-mapping tables: `SFTP - AWS Sync` (a different kind of tracking sheet), `Navex Location` and `Dynamo DB Audit` (no header row found), `Repair Order Canada` (header only, no data rows). 33 tables listed in the catalog have no field-level mapping sheet of their own — these show up with full catalog metadata (source system, feed type, format, frequency, S3 paths) but zero documented fields, flagged as "no field mapping yet."

**Both** — many fields have no "Field definition" in the source sheet — the app shows these as "no definition provided" rather than guessing. Source-system category labels come directly from the spreadsheets and aren't normalized, so near-duplicates like "asTech" / "asTech / OPUS" or "QCOMM" / "QComm" can appear as separate filter chips.
