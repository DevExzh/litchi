#!/usr/bin/env python3
"""Verify and summarize an extracted run_pairs.py capture."""
import hashlib
import json
from pathlib import Path
import re
import statistics
import sys

capture = Path(sys.argv[1])
receipt = json.loads((capture / 'receipt.json').read_text())
assert receipt['binary_unchanged']
for name, expected in receipt['artifacts'].items():
    assert hashlib.sha256((capture / name).read_bytes()).hexdigest() == expected, name
records = receipt['records']
assert len(records) == 6 * len(receipt['cases'])
assert all(row['status'] == 0 for row in records)
# Compare every non-timing result field, including allocation and live memory.
ignored = {'elapsed_ns', 'elapsed_ns_mean', 'elapsed_ns_p95', 'elapsed_ns_p99', 'elapsed_ns_per_repeat'}
summary = []
for case in receipt['cases']:
    groups = {label: [r for r in records if r['case'] == case and r['label'] == label]
              for label in ('before', 'after')}
    results = [r['result'] for rows in groups.values() for r in rows]
    baseline = {k: v for k, v in results[0].items() if k not in ignored}
    assert all({k: v for k, v in r.items() if k not in ignored} == baseline for r in results), case
    row = {'case': case, 'deterministic_parity': True}
    for label, rows in groups.items():
        assert [r['round'] for r in rows] == [1, 2, 3]
        rss = []
        for r in rows:
            prefix = f"{r['round']:02}-{case}-{label}"
            raw = json.loads((capture / (prefix + '.jsonl')).read_text())
            assert raw == r['result']
            rss.append(int(re.search(r'Maximum resident set size \(kbytes\): (\d+)',
                                     (capture / (prefix + '.time')).read_text())[1]))
        row[label] = {'rss_kib': rss}
        for metric in ('elapsed_ns', 'elapsed_ns_p95'):
            row[label][metric] = [r['result'][metric] for r in rows]
    row['delta_percent'] = {metric: 100 * (statistics.median(row['after'][metric]) /
                                          statistics.median(row['before'][metric]) - 1)
                            for metric in ('rss_kib', 'elapsed_ns', 'elapsed_ns_p95')}
    row['review_flag'] = any(delta > 5 for delta in row['delta_percent'].values())
    summary.append(row)
for number in (1, 2, 3):
    expected = [(case, label) for case in receipt['cases']
                for label in (('before', 'after') if number % 2 else ('after', 'before'))]
    assert [(r['case'], r['label']) for r in records if r['round'] == number] == expected
print(json.dumps({'runs': len(records), 'failed_runs': 0, 'deterministic_mismatches': [],
                  'cases': summary}, indent=2))
