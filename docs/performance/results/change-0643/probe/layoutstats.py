#!/usr/bin/env python3
"""p50 per leg for the A1 C1 B1 B2 C2 A2 layout-control runs."""
import sys, os, statistics as st
def load(path):
    v = sorted(int(x) for x in open(path) if x.strip())
    return {"n": len(v), "p50": v[len(v)//2], "mean": st.fmean(v), "p95": v[min(len(v)-1,int(len(v)*0.95))],
            "p99": v[min(len(v)-1,int(len(v)*0.99))], "min": v[0]}
d = sys.argv[1]
cases = sorted({f.rsplit("-",1)[0] for f in os.listdir(d) if f.endswith(".txt")})
runs = ("A1","C1","B1","B2","C2","A2")
print(f"{'case':<44}" + "".join(f"{r+' p50':>11}" for r in runs) + f"{'A/A %':>8}{'C/C %':>8}{'B/B %':>8}{'C-A %':>8}{'B-C %':>8}{'B-A %':>8}")
for case in cases:
    try:
        legs = {r: load(os.path.join(d, f"{case}-{r}.txt")) for r in runs}
    except FileNotFoundError:
        continue
    pa = (legs["A1"]["p50"]+legs["A2"]["p50"])/2
    pc = (legs["C1"]["p50"]+legs["C2"]["p50"])/2
    pb = (legs["B1"]["p50"]+legs["B2"]["p50"])/2
    f = lambda x, y: (y-x)/x*100
    print(f"{case:<44}" + "".join(f"{legs[r]['p50']:>11}" for r in runs)
          + f"{f(legs['A1']['p50'],legs['A2']['p50']):>8.2f}{f(legs['C1']['p50'],legs['C2']['p50']):>8.2f}"
          + f"{f(legs['B1']['p50'],legs['B2']['p50']):>8.2f}{f(pa,pc):>8.2f}{f(pc,pb):>8.2f}{f(pa,pb):>8.2f}")
print()
print(f"{'case':<44}{'leg':<5}{'n':>5}{'p50':>12}{'mean':>13}{'p95':>11}{'p99':>11}{'min':>11}")
for case in cases:
    for r in runs:
        p = os.path.join(d, f"{case}-{r}.txt")
        if not os.path.exists(p): continue
        row = load(p)
        print(f"{case:<44}{r:<5}{row['n']:>5}{row['p50']:>12}{row['mean']:>13.1f}{row['p95']:>11}{row['p99']:>11}{row['min']:>11}")
