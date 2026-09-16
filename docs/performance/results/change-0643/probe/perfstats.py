#!/usr/bin/env python3
"""Difference-of-medians over the reps=0 / reps=1 perf pairs, per op and leg."""
import csv, sys, statistics as st
from collections import defaultdict
rows = list(csv.DictReader(open(sys.argv[1])))
counters = ["cycles", "instructions", "branch_misses", "dtlb_load_misses", "page_faults"]
g = defaultdict(lambda: defaultdict(list))
for r in rows:
    for c in counters:
        g[(r["op"], r["leg"], r["reps"])][c].append(int(r[c]))
ops = sorted({r["op"] for r in rows})
legs = sorted({r["leg"] for r in rows})
print(f"{'op':<38}{'counter':<20}" + "".join(f"{leg:>16}" for leg in legs) + f"{'delta %':>10}")
for op in ops:
    for c in counters:
        per = {}
        for leg in legs:
            lo = g[(op, leg, "0")][c]
            hi = g[(op, leg, "1")][c]
            if not lo or not hi:
                continue
            per[leg] = st.median(hi) - st.median(lo)
        if len(per) < 2:
            continue
        base = per[legs[0]]
        last = per[legs[-1]]
        d = (last - base) / base * 100 if base else float("nan")
        print(f"{op:<38}{c:<20}" + "".join(f"{per.get(leg, float('nan')):>16.0f}" for leg in legs) + f"{d:>10.2f}")
