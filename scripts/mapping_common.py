"""
Shared parsing helpers for BDAP mapping workbooks (staging<>warehouse,
source<>data lake, and any future one-sheet-per-table mapping spec).

Each such workbook has one summary/index sheet describing every table, plus
one sheet per table with its field-level mapping (source field -> target
field, data type, mandatory flag, transformation rule, notes, etc). Column
headers are not fully consistent sheet to sheet or workbook to workbook, so
headers are matched by pattern rather than by fixed position.
"""
import collections
import datetime
import re

HEADER_MAP_RULES = [
    (re.compile(r'^target table \(oubound table\)$', re.I), 'target_table'),
    (re.compile(r'^target table$', re.I), 'target_table'),
    (re.compile(r'^target file$', re.I), 'target_table'),
    (re.compile(r'^target file name$', re.I), 'target_file_name'),
    (re.compile(r'^target field \(oubound table\)$', re.I), 'target_field'),
    (re.compile(r'^target field \(outbound file\)$', re.I), 'target_field_outbound'),
    (re.compile(r'^target field$', re.I), 'target_field'),
    (re.compile(r'^source table\b', re.I), 'source_table'),
    (re.compile(r'^source file$', re.I), 'source_table'),
    (re.compile(r'^source$', re.I), 'source_table'),
    (re.compile(r'^bi-weekly report$', re.I), 'source_table'),
    (re.compile(r'^source field.*$', re.I), 'source_field'),
    (re.compile(r'^source file column$', re.I), 'source_field'),
    (re.compile(r'^athena table$', re.I), 'athena_table'),
    (re.compile(r'^athena field$', re.I), 'athena_field'),
    (re.compile(r'^csv fields?$', re.I), 'csv_field'),
    (re.compile(r'^processed path$', re.I), 'processed_path'),
    (re.compile(r'^mandatory\??$', re.I), 'mandatory'),
    (re.compile(r'^data ?type( and length)?$', re.I), 'data_type'),
    (re.compile(r'^field type$', re.I), 'data_type'),
    (re.compile(r'^transformation rule$', re.I), 'transformation_rule'),
    (re.compile(r'^source format\s*/\s*examples$', re.I), 'source_format'),
    (re.compile(r'^ge ?rules?$', re.I), 'ge_rules'),
    (re.compile(r'^notes$', re.I), 'notes'),
    (re.compile(r'^what it represents$', re.I), 'notes'),
    (re.compile(r'^comments$', re.I), 'comments'),
    (re.compile(r'^field def[ei]ni?tion$', re.I), 'field_definition'),
    (re.compile(r'^join condition$', re.I), 'join_condition'),
    (re.compile(r'^legacy col(umn)?s?$', re.I), 'legacy_column'),
    (re.compile(r'^legacy table/sp$', re.I), 'legacy_column'),
    (re.compile(r'^legacy table$', re.I), 'legacy_column'),
    (re.compile(r'^legacy field name$', re.I), 'legacy_column'),
    (re.compile(r'^legacy logic/mapping$', re.I), 'legacy_logic'),
    (re.compile(r'^sp logic/mapping$', re.I), 'legacy_logic'),
    (re.compile(r'^mapping in yaml$', re.I), 'yaml_mapping'),
    (re.compile(r'^status$', re.I), 'status'),
]


def norm_header(h):
    if h is None:
        return None
    h = str(h).strip()
    for rx, key in HEADER_MAP_RULES:
        if rx.match(h):
            return key
    return None


def clean(v):
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        return v.isoformat()
    return v


def strip_suffix(name):
    return re.sub(r'\s*\([^)]*\)\s*$', '', name).strip()


def norm_key(name):
    return strip_suffix(name).strip().lower()


def parse_detail_sheets(wb, exclude_sheet_names):
    """Find the header row of every sheet (except the excluded ones) and
    extract its data rows as normalized dicts. Returns (sheets, skipped)."""
    sheets = {}
    skipped = []
    for name in wb.sheetnames:
        if name in exclude_sheet_names:
            continue
        ws = wb[name]
        header_row_idx = None
        header_map = {}
        for r in range(1, 8):
            row_vals = [c.value for c in ws[r]]
            if not row_vals:
                continue
            mapped = {i: norm_header(v) for i, v in enumerate(row_vals) if norm_header(v)}
            if 'target_table' in mapped.values() or 'target_field' in mapped.values():
                header_row_idx = r
                header_map = mapped
                break
        if header_row_idx is None:
            skipped.append(name)
            continue

        rows_out = []
        seen_target_tables = set()
        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            rec = {}
            for i, key in header_map.items():
                if i < len(row):
                    val = clean(row[i])
                    if val is not None:
                        rec[key] = val
            if not rec:
                continue
            if rec.get('target_table'):
                seen_target_tables.add(rec['target_table'])
            if any(k in rec for k in ('target_field', 'target_field_outbound', 'source_field', 'source_table')):
                rows_out.append(rec)

        if not rows_out and not seen_target_tables:
            skipped.append(name)
            continue

        sheets[name] = {'sheet_name': name, 'columns': rows_out}

    return sheets, skipped


FIELD_LIST_BUCKETS = (
    ('transformation_rule', 'transformation_rules'),
    ('ge_rules', 'ge_rules'),
    ('notes', 'notes'),
    ('comments', 'comments'),
    ('join_condition', 'join_conditions'),
    ('legacy_column', 'legacy'),
    ('legacy_logic', 'legacy'),
    ('yaml_mapping', 'yaml_mapping'),
    ('csv_field', 'csv_fields'),
    ('processed_path', 'processed_paths'),
)


def build_fields_and_rules(rows):
    """Group a sheet's raw rows into one entry per target field, plus any
    table-level rule/note rows that carry no target field of their own."""
    fields = collections.OrderedDict()
    table_rules = []

    for r in rows:
        fname = r.get('target_field') or r.get('target_field_outbound')
        if not fname:
            content = r.get('source_field') or r.get('transformation_rule') or r.get('notes') or r.get('comments')
            if content and (r.get('source_table') or r.get('transformation_rule')):
                table_rules.append({'label': r.get('source_table') or 'Note', 'text': content})
            continue
        f = fields.setdefault(fname, {
            'field_name': fname, 'data_type': None, 'mandatory': None, 'definition': None,
            'transformation_rules': [], 'ge_rules': [], 'notes': [], 'comments': [],
            'sources': [], 'join_conditions': [], 'legacy': [], 'yaml_mapping': [],
            'csv_fields': [], 'processed_paths': [], 'status': None,
        })
        if r.get('data_type') and not f['data_type']:
            f['data_type'] = r['data_type']
        if r.get('mandatory'):
            if not f['mandatory'] or r['mandatory'].strip().upper().startswith('Y'):
                f['mandatory'] = r['mandatory']
        if r.get('field_definition') and not f['definition']:
            f['definition'] = r['field_definition']
        if r.get('status') and not f['status']:
            f['status'] = r['status']
        for key, bucket in FIELD_LIST_BUCKETS:
            v = r.get(key)
            if v and v not in f[bucket]:
                f[bucket].append(v)
        src = {k: r[k] for k in ('source_table', 'source_field', 'athena_table', 'athena_field', 'source_format') if r.get(k)}
        if src and src not in f['sources']:
            f['sources'].append(src)

    return list(fields.values()), table_rules
