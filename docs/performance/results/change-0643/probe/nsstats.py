#!/usr/bin/env python3
"""p50/mean/p95/p99 of one-nanosecond-per-line probe timing files, and the
A1 B1 B2 A2 paired deltas."""
import sys, os, statistics as st

def load(path):
    values = sorted(int(line) for line in open(path) if line.strip())
    n = len(values)
    return {
        "n": n,
        "p50": values[n // 2],
        "mean": st.fmean(values),
        "p95": values[min(n - 1, int(n * 0.95))],
        "p99": values[min(n - 1, int(n * 0.99))],
        "min": values[0],
    }

directory = sys.argv[1]
cases = sorted({f.rsplit("-", 1)[0] for f in os.listdir(directory) if f.endswith(".txt")})
print(f"{'case':<40}{'A1 p50':>11}{'B1 p50':>11}{'B2 p50':>11}{'A2 p50':>11}{'A/A %':>9}{'pooled %':>10}")
for case in cases:
    try:
        legs = {run: load(os.path.join(directory, f"{case}-{run}.txt")) for run in ("A1", "B1", "B2", "A2")}
    except FileNotFoundError:
        continue
    pa = (legs["A1"]["p50"] + legs["A2"]["p50"]) / 2
    pb = (legs["B1"]["p50"] + legs["B2"]["p50"]) / 2
    print(f"{case:<40}{legs['A1']['p50']:>11}{legs['B1']['p50']:>11}{legs['B2']['p50']:>11}{legs['A2']['p50']:>11}"
          f"{(legs['A2']['p50'] - legs['A1']['p50']) / legs['A1']['p50'] * 100:>9.2f}{(pb - pa) / pa * 100:>10.2f}")
print()
print(f"{'case':<40}{'leg':<5}{'n':>6}{'p50':>12}{'mean':>14}{'p95':>12}{'p99':>12}{'min':>12}")
for case in cases:
    for run in ("A1", "B1", "B2", "A2"):
        path = os.path.join(directory, f"{case}-{run}.txt")
        if not os.path.exists(path):
            continue
        row = load(path)
        print(f"{case:<40}{run:<5}{row['n']:>6}{row['p50']:>12}{row['mean']:>14.1f}{row['p95']:>12}{row['p99']:>12}{row['min']:>12}")
