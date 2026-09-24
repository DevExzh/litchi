#!/usr/bin/env python3
"""Summarize the existing-value-workload paired capture from its archive."""
import json
import statistics
import sys
import tarfile

with tarfile.open(sys.argv[1]) as archive:
    receipt = json.load(archive.extractfile('run.json'))
    assert receipt['status'] == 'complete' and not receipt['failures']
    assert receipt['child_count'] == 162 and receipt['pair_count'] == 81
    assert receipt['deterministic_mismatch_pairs'] == 0
    assert all(receipt[k] for k in ('binaries_unchanged', 'harness_unchanged', 'runner_unchanged', 'source_receipts_unchanged'))
    children = receipt['children']
    assert all(c['status'] == 0 and not c['validation_errors'] for c in children)
    cases = []
    for case in sorted({c['case'] for c in children}):
        groups = {s: [c['row'] for c in children if c['case'] == case and c['slot'] == s] for s in ('A', 'B')}
        assert all(len(g) == 3 for g in groups.values())
        row = {'case': case, 'repeat': groups['A'][0]['repeat']}
        metrics = ('p50_ns', 'p95_ns', 'max_rss_kib')
        for slot in groups:
            row[slot] = {k: [int(r[k]) for r in groups[slot]] for k in metrics}
        row['delta_percent'] = {k: 100 * (statistics.median(row['B'][k]) / statistics.median(row['A'][k]) - 1) for k in metrics}
        counters = [k for k in groups['A'][0] if k.startswith(('alloc_', 'dealloc_', 'requested_', 'released_', 'live_', 'peak_live_', 'memory_retained_', 'adapter_index_'))]
        row['memory_changes'] = {}
        for k in counters:
            values = {s: sorted({r[k] for r in groups[s]}) for s in groups}
            if values['A'] != values['B']:
                row['memory_changes'][k] = values
        row['review_flag'] = any(v > 5 for v in row['delta_percent'].values())
        cases.append(row)
print(json.dumps({'children': 162, 'pairs': 81, 'deterministic_mismatches': 0, 'cases': cases}, indent=2))
