#!/usr/bin/env python3
"""Select every initial comparable latency/RSS regression above five percent."""
import csv
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent


def main():
    with (HERE / 'baseline/comparable/raw.csv').open() as stream:
        baseline = {(r['phase'], r['case']): r for r in csv.DictReader(stream)}
    flags = []
    with (HERE / 'candidate/comparable/raw.csv').open() as stream:
        for row in csv.DictReader(stream):
            old = baseline[row['phase'], row['case']]
            assert old['status'] == row['status'] == '0'
            delta = {key: (float(row[key]) / float(old[key]) - 1) * 100
                     if float(old[key]) else 0
                     for key in ['p50_ns', 'p95_ns', 'p99_ns', 'max_rss_kib']}
            if any(value > 5 for value in delta.values()):
                flags.append({'group': 'comparable', 'phase': row['phase'],
                              'case': row['case'], 'repeat': int(row['repeat']),
                              'delta_pct': delta})
    (HERE / 'initial-flags.json').write_text(json.dumps(flags, indent=2) + '\n')
    print(f'{len(flags)} initial flags')


if __name__ == '__main__':
    main()
