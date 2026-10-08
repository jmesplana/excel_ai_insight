"""Compare an exported regression run with proposed expectations; no API calls."""
import argparse
import json
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]


def compare(rows, expected):
    by_id = {}
    for row in rows:
        key = row.get('case_id')
        if key in by_id:
            raise ValueError(f'Duplicate case_id: {key}')
        by_id[key] = row
    failures, checks = [], 0
    for case_id, fields in expected.items():
        if case_id not in by_id:
            failures.append({'case_id': case_id, 'issue': 'Missing row'})
            continue
        for field, allowed in fields.items():
            checks += 1
            actual = by_id[case_id].get(field)
            if actual not in allowed:
                failures.append({'case_id': case_id, 'field': field, 'actual': actual, 'expectedOneOf': allowed})
    return {'cases': len(expected), 'fieldChecks': checks, 'mismatches': failures,
            'note': 'Agreement with proposed expectations is not a validated population accuracy score.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('export', type=Path)
    parser.add_argument('--expectations', type=Path, default=ROOT / 'sample_jev/jev_regression_expectations.json')
    args = parser.parse_args()
    expected = json.loads(args.expectations.read_text(encoding='utf-8'))['cases']
    book = load_workbook(args.export, read_only=True, data_only=True)
    try:
        values = iter(book.active.values)
        headers = next(values)
        rows = [dict(zip(headers, row)) for row in values]
    finally:
        book.close()
    report = compare(rows, expected)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    raise SystemExit(1 if report['mismatches'] else 0)


if __name__ == '__main__':
    main()
