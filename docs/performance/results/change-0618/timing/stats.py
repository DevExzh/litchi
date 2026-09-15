"""Percentiles, paired deltas in both directions, and the A/A floor.

Adapted from change 0607's `timing/stats.py`. A is the before leg, B the after
leg; A1/A2 bracket B1/B2 so the A/A pair measures the floor in the same window.
"""
import statistics as st, sys, pathlib
def load(p): return sorted(float(x) for x in pathlib.Path(p).read_text().split())
def q(v, f):
    i = min(len(v)-1, int(round(f*(len(v)-1))))
    return v[i]
def line(name, v):
    return (f"{name:24s} n={len(v):<5d} p50={q(v,0.5)/1000:10.2f} mean={st.mean(v)/1000:10.2f} "
            f"p95={q(v,0.95)/1000:10.2f} p99={q(v,0.99)/1000:10.2f} us")
d = sys.argv[1]
legs = {n: load(f"{d}/{n}.txt") for n in ("A1","B1","B2","A2")}
print(f"# {d}")
for n in ("A1","A2","B1","B2"): print(line(n, legs[n]))
A = sorted(legs["A1"]+legs["A2"]); B = sorted(legs["B1"]+legs["B2"])
print(line("A (before, pooled)", A)); print(line("B (after, pooled)", B))
for stat, f in (("p50", lambda v: q(v,0.5)), ("mean", st.mean), ("p95", lambda v: q(v,0.95)), ("p99", lambda v: q(v,0.99))):
    a, b = f(A), f(B)
    print(f"delta {stat:5s} B-A = {(b-a)/1000:+10.2f} us  {100*(b-a)/a:+7.2f}% of A   "
          f"A-B = {(a-b)/1000:+10.2f} us  {100*(a-b)/b:+7.2f}% of B")
print(f"A/A floor (A2 vs A1)     p50 {100*(q(legs['A2'],0.5)-q(legs['A1'],0.5))/q(legs['A1'],0.5):+.2f}%"
      f"  p95 {100*(q(legs['A2'],0.95)-q(legs['A1'],0.95))/q(legs['A1'],0.95):+.2f}%"
      f"  p99 {100*(q(legs['A2'],0.99)-q(legs['A1'],0.99))/q(legs['A1'],0.99):+.2f}%")
print(f"B/B floor (B2 vs B1)     p50 {100*(q(legs['B2'],0.5)-q(legs['B1'],0.5))/q(legs['B1'],0.5):+.2f}%"
      f"  p95 {100*(q(legs['B2'],0.95)-q(legs['B1'],0.95))/q(legs['B1'],0.95):+.2f}%"
      f"  p99 {100*(q(legs['B2'],0.99)-q(legs['B1'],0.99))/q(legs['B1'],0.99):+.2f}%")
