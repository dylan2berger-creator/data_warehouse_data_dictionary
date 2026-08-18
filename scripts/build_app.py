#!/usr/bin/env python3
"""
Assemble index.html from the built datasets (data/data_dictionary.json,
data/data_lake_dictionary.json, ...) and the app shell template.

Run this after build_dictionary.py / build_datalake_dictionary.py.

Usage:
    python3 scripts/build_app.py [--out index.html] [--template scripts/app_template.html] [dataset.json ...]
"""
import argparse
import json
import sys
from pathlib import Path

DEFAULT_DATASETS = ['data/data_dictionary.json', 'data/data_lake_dictionary.json']


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('datasets', nargs='*', default=DEFAULT_DATASETS, help='Dataset JSON files to embed (default: both built-in datasets)')
    ap.add_argument('--out', default='index.html')
    ap.add_argument('--template', default='scripts/app_template.html')
    args = ap.parse_args()

    datasets = {}
    order = []
    for path in args.datasets:
        p = Path(path)
        if not p.exists():
            print(f"Skipping missing dataset file: {path}", file=sys.stderr)
            continue
        d = json.loads(p.read_text(encoding='utf-8'))
        datasets[d['id']] = d
        order.append(d['id'])
        print(f"Embedding dataset '{d['id']}' ({len(d['tables'])} tables) from {path}", file=sys.stderr)

    if not datasets:
        raise SystemExit("No dataset files found -- run the build_*.py scripts first")

    combined = {'datasetOrder': order, 'datasets': datasets}
    data_json = json.dumps(combined, separators=(',', ':')).replace('</', '<\\/')

    template = Path(args.template).read_text(encoding='utf-8')
    if '__DATA_JSON__' not in template:
        raise SystemExit(f"Template {args.template} is missing the __DATA_JSON__ placeholder")
    out = template.replace('__DATA_JSON__', data_json)

    Path(args.out).write_text(out, encoding='utf-8')
    print(f"Wrote {args.out}", file=sys.stderr)


if __name__ == '__main__':
    main()
