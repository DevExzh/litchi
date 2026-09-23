#!/usr/bin/env python3
"""Median of per-process p50s per arm for each probe measure (A = base-code
probe, B = candidate probe), with B/A."""
import collections, csv, glob, os, re, statistics, sys
root = sys.argv[1] if len(sys.argv) > 1 else '.'
rows = collections.defaultdict(list)
for path in sorted(glob.glob(os.path.join(root, '*.csv'))):
    m = re.match(r'(.+)-r(\d)-s(\d)-([AB])\.csv', os.path.basename(path))
    arm = m.group(4)
    for r in csv.DictReader(open(path)):
        rows[(r['label'], r['side'], r['measure'], arm)].append(float(r['p50_ns']))
out = []
for key in sorted({k[:3] for k in rows}):
    a = rows.get(key + ('A',)); b = rows.get(key + ('B',))
    am = statistics.median(a) if a else None
    bm = statistics.median(b) if b else None
    ratio = (bm / am) if (am and bm) else None
    out.append((key, am, bm, ratio, len(a or []), len(b or [])))
    fmt = lambda v: '-' if v is None else f'{v:,.0f}'
    print(f"{key[0]:<24} {key[1]:<12} {key[2]:<26} A {fmt(am):>12} B {fmt(bm):>12} {'' if ratio is None else f'{ratio:.4f}'}")
