#!/usr/bin/env python3
"""Ratio = median(after legs) / median(before legs) per scenario."""
import json, sys, glob, statistics, os
out = {}
for d in sys.argv[1:]:
    rows = {}
    for f in sorted(glob.glob(os.path.join(d, 'probe-*.json'))):
        leg = 'after' if 'after' in f else 'before'
        for r in json.load(open(f)):
            rows.setdefault(r['scenario'], {'before': [], 'after': []})[leg].append(r['median_us'])
    out[d] = {k: {'before': statistics.median(v['before']), 'after': statistics.median(v['after']),
                  'ratio': statistics.median(v['after']) / statistics.median(v['before'])} for k, v in rows.items()}
keys = list(next(iter(out.values())).keys())
for k in keys:
    cells = [f"{out[d][k]['before']:.1f} {out[d][k]['after']:.1f} {out[d][k]['ratio']:.3f}" for d in out]
    print(f"| {k} | " + " | ".join(cells) + " |")
json.dump(out, open(os.path.join(os.path.dirname(sys.argv[1]), 'summary.json'), 'w'), indent=1)
