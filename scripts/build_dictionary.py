#!/usr/bin/env python3
"""
Build the Data Warehouse dataset from the staging-to-warehouse mapping
workbook ("BDAP Staging <> Warehouse Mapping.xlsx").

The source workbook has one "Summary  Index" sheet listing every warehouse
table (staging source, source system, table type, description) plus one
sheet per table containing its field-level mapping. See mapping_common.py
for how headers are matched.

Usage:
    python3 scripts/build_dictionary.py <path-to-xlsx> [--out-json data/data_dictionary.json]

Requires: openpyxl (pip install openpyxl)
Then run scripts/build_app.py to assemble index.html from this dataset
(and any others produced by the sibling build_*.py scripts).
"""
import argparse
import collections
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from mapping_common import parse_detail_sheets, build_fields_and_rules, clean, norm_key

SUMMARY_SHEET_NAME = 'Summary  Index'


def category_for(name):
    n = name.lower()
    if n.startswith('dim_'):
        return 'Dimension'
    if n.startswith('fct_') or n.startswith('fact_'):
        return 'Fact'
    if n.startswith('agg_'):
        return 'Aggregate'
    if n.startswith('wd_validation') or n.startswith('wd_journal_entry'):
        return 'Workday Journal / Validation'
    if 'audit' in n:
        return 'Audit'
    return 'Other'


def parse_catalog(wb):
    if SUMMARY_SHEET_NAME not in wb.sheetnames:
        raise SystemExit(f"Expected a '{SUMMARY_SHEET_NAME}' sheet, not found in workbook")
    summary_ws = wb[SUMMARY_SHEET_NAME]
    catalog = {}
    sidx_headers = [c.value for c in summary_ws[1]]
    for row in summary_ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(sidx_headers, row))
        tname = clean(d.get('BDAP Table Name (DW)'))
        if not tname:
            continue
        catalog[tname] = {
            'staging_source': clean(d.get('Staging Source')),
            'source_system': clean(d.get('Source System')),
            'table_type': clean(d.get('Type')),
            'description': clean(d.get('BDAP Table Description (DW)')),
            'comments': clean(d.get('Comments')),
        }
    return catalog


def build_tables(catalog, sheets):
    catalog_lookup = {norm_key(k): (k, v) for k, v in catalog.items()}
    tables_out = {}

    for sheet_name, sheet in sheets.items():
        rows = sheet['columns']
        target_names = [r.get('target_table') for r in rows if r.get('target_table')]
        canonical = collections.Counter(target_names).most_common(1)[0][0] if target_names else sheet_name

        field_list, table_rules = build_fields_and_rules(rows)

        match = catalog_lookup.get(norm_key(canonical))
        if not match:
            stem = norm_key(canonical).rstrip('_')
            candidates = [(k, v) for nk, (k, v) in catalog_lookup.items() if nk.rstrip('_').startswith(stem) or stem.startswith(nk.rstrip('_'))]
            if len(candidates) == 1:
                match = candidates[0]
        cat_key, cat_info = match if match else (None, None)
        display_name = cat_key if cat_key else canonical

        tables_out[sheet_name] = {
            'id': sheet_name,
            'display_name': display_name,
            'sheet_name': sheet_name,
            'category': category_for(display_name),
            'warehouse_type': (cat_info or {}).get('table_type'),
            'description': (cat_info or {}).get('description'),
            'source_system': (cat_info or {}).get('source_system'),
            'staging_source': (cat_info or {}).get('staging_source'),
            'catalog_comments': (cat_info or {}).get('comments'),
            'field_count': len(field_list),
            'fields': field_list,
            'table_rules': table_rules,
        }

    matched_keys = {norm_key(t['display_name']) for t in tables_out.values()}
    unmatched_catalog = [k for k in catalog if norm_key(k) not in matched_keys]
    return tables_out, unmatched_catalog


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('xlsx_path', help='Path to the BDAP Staging <> Warehouse Mapping .xlsx file')
    ap.add_argument('--out-json', default='data/data_dictionary.json')
    args = ap.parse_args()

    import openpyxl
    wb = openpyxl.load_workbook(args.xlsx_path, data_only=True)

    catalog = parse_catalog(wb)
    sheets, skipped = parse_detail_sheets(wb, exclude_sheet_names={SUMMARY_SHEET_NAME})
    tables, unmatched_catalog = build_tables(catalog, sheets)

    total_fields = sum(t['field_count'] for t in tables.values())
    print(f"[warehouse] Parsed {len(tables)} tables, {total_fields} fields.", file=sys.stderr)
    if skipped:
        print(f"[warehouse] Skipped {len(skipped)} sheet(s) with no recognizable header: {skipped}", file=sys.stderr)
    if unmatched_catalog:
        print(f"[warehouse] {len(unmatched_catalog)} catalog entries had no matching sheet: {unmatched_catalog}", file=sys.stderr)

    data = {
        'id': 'warehouse',
        'label': 'Data Warehouse',
        'tables': tables,
        'unmatched_catalog': unmatched_catalog,
    }

    out_json = Path(args.out_json)
    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_json.write_text(json.dumps(data, indent=1), encoding='utf-8')
    print(f"[warehouse] Wrote {out_json}", file=sys.stderr)


if __name__ == '__main__':
    main()
