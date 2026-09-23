#!/usr/bin/env python3
"""Per-case allocation medians for both legs of change 0752."""
import glob, json, os, statistics, sys
directory = sys.argv[1]
table = {}
for path in sorted(glob.glob(os.path.join(directory, '*-before.json')) + glob.glob(os.path.join(directory, '*-after.json'))):
    leg = 'before' if path.endswith('-before.json') else 'after'
    for r in json.load(open(path))['results']:
        a = r.get('operation_metrics', {}).get('allocation')
        if not a or a.get('status') != 'measured':
            continue  # this case reports no allocator counters
        key = (r['case'], r['corpus']['name'])
        table.setdefault(key, {})[leg] = {m: statistics.median(a[m]['values']) for m in
                                          ('allocation_calls', 'reallocation_calls', 'allocated_bytes', 'region_peak_live_bytes')}
rows = []
print(f"{'case':30} {'corpus':42} {'alloc calls':>18} {'allocated bytes':>26} {'region peak bytes':>24}")
for key, legs in sorted(table.items()):
    b, a = legs['before'], legs['after']
    rows.append({'case': key[0], 'corpus': key[1], 'before': b, 'after': a})
    print(f"{key[0]:30} {key[1]:42} {b['allocation_calls']:>8.0f} -> {a['allocation_calls']:<8.0f} {b['allocated_bytes']:>12.0f} -> {a['allocated_bytes']:<12.0f} {b['region_peak_live_bytes']:>11.0f} -> {a['region_peak_live_bytes']:<11.0f}")
json.dump(rows, open(os.path.join(directory, 'summary.json'), 'w'), indent=1)
