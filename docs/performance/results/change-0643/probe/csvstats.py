#!/usr/bin/env python3
"""Per-(op, order, leg) medians of every counter column in a first-call CSV,
then the pooled before/after deltas and the A/A floor."""
import csv, sys, statistics as st
from collections import defaultdict

rows = list(csv.DictReader(open(sys.argv[1])))
counters = [c for c in rows[0] if c not in ("op", "order", "leg", "rep")]
groups = defaultdict(lambda: defaultdict(list))
for row in rows:
    for c in counters:
        groups[(row["op"], row["order"], row["leg"])][c].append(int(row[c]))

ops = sorted({r["op"] for r in rows})
print(f"{'op':<38}{'order':<7}{'leg':<9}" + "".join(f"{c:>20}" for c in counters))
for op in ops:
    for order in ("1", "2", "3", "4"):
        for (o, ord_, leg), data in groups.items():
            if o == op and ord_ == order:
                print(f"{op:<38}{order:<7}{leg:<9}" + "".join(f"{st.median(data[c]):>20.0f}" for c in counters))
print()
print(f"{'op':<38}{'counter':<22}{'A1':>14}{'B1':>14}{'B2':>14}{'A2':>14}{'A/A %':>9}{'pooled %':>10}")
for op in ops:
    for c in counters:
        try:
            a1 = st.median(groups[(op, "1", "rev0592")][c]); b1 = st.median(groups[(op, "2", "base")][c])
            b2 = st.median(groups[(op, "3", "base")][c]);    a2 = st.median(groups[(op, "4", "rev0592")][c])
        except KeyError:
            continue
        pa, pb = (a1 + a2) / 2, (b1 + b2) / 2
        aa = (a2 - a1) / a1 * 100 if a1 else float("nan")
        pooled = (pb - pa) / pa * 100 if pa else float("nan")
        print(f"{op:<38}{c:<22}{a1:>14.0f}{b1:>14.0f}{b2:>14.0f}{a2:>14.0f}{aa:>9.2f}{pooled:>10.2f}")
