#!/usr/bin/env python3
"""Paired-timing analysis for change 0660.

Pools the two runs of each leg, reports p50/mean/p95/p99 per case and shape, the
delta in both directions, and the A/A floor observed in the same window.
"""
import json, sys, statistics

def rows(path):
    out = {}
    for r in json.load(open(path))["results"]:
        out[(r["case"], r["corpus"]["shape"])] = r["elapsed_ns"]["samples"]
    return out

def stats(samples):
    s = sorted(samples)
    n = len(s)
    q = lambda p: s[min(n - 1, max(0, int(round(p * (n - 1)))))]
    return {"n": n, "p50": q(0.5), "mean": statistics.fmean(s),
            "p95": q(0.95), "p99": q(0.99), "min": s[0], "max": s[-1]}

def pct(new, old):
    return (new - old) / old * 100.0

def main():
    out = sys.argv[1]
    a1, b1, b2, a2 = (rows(f"{out}/{leg}.json") for leg in ("A1", "B1", "B2", "A2"))
    keys = sorted(a1)
    print(f"{'case':38s} {'shape':7s} {'before p50':>11s} {'after p50':>11s} "
          f"{'delta%':>8s} {'inv%':>8s} {'floor%':>8s} {'b p95':>10s} {'a p95':>10s} "
          f"{'b p99':>10s} {'a p99':>10s} {'b mean':>10s} {'a mean':>10s}")
    report = {}
    for key in keys:
        before = stats(a1[key] + a2[key])
        after = stats(b1[key] + b2[key])
        floor = pct(stats(a2[key])["p50"], stats(a1[key])["p50"])
        bfloor = pct(stats(b2[key])["p50"], stats(b1[key])["p50"])
        delta = pct(after["p50"], before["p50"])
        inverse = pct(before["p50"], after["p50"])
        report["/".join(key)] = {"before": before, "after": after,
                                 "p50_delta_pct": delta, "p50_inverse_pct": inverse,
                                 "aa_floor_p50_pct": floor, "bb_floor_p50_pct": bfloor}
        print(f"{key[0]:38s} {key[1]:7s} {before['p50']:11d} {after['p50']:11d} "
              f"{delta:8.2f} {inverse:8.2f} {floor:8.2f} {before['p95']:10d} {after['p95']:10d} "
              f"{before['p99']:10d} {after['p99']:10d} {before['mean']:10.0f} {after['mean']:10.0f}")
    json.dump(report, open(f"{out}/paired-summary.json", "w"), indent=1)

main()
