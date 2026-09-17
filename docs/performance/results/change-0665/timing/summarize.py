#!/usr/bin/env python3
"""Pool each leg's two runs and print p50/mean/p95/p99, both deltas and the
A/A and B/B floors observed in the same window."""
import json, statistics, sys

def rows(path):
    out = {}
    for r in json.load(open(path))["results"]:
        out[(r["case"], r["corpus"].get("shape") or r["corpus"].get("name", "-"))] = r["elapsed_ns"]["samples"]
    return out

def stats(samples):
    s = sorted(samples); n = len(s)
    q = lambda p: s[min(n - 1, max(0, int(round(p * (n - 1)))))]
    return {"n": n, "p50": q(0.5), "mean": statistics.fmean(s), "p95": q(0.95), "p99": q(0.99),
            "min": s[0], "max": s[-1]}

def pct(new, old):
    return (new - old) / old * 100.0

def main():
    root = sys.argv[1]
    legs = {leg: rows(f"{root}/{leg}.json") for leg in ("A1", "B1", "B2", "A2", "A3", "A4")}
    report = {}
    print(f"{'case':40s} {'shape':13s} {'before p50':>12s} {'after p50':>12s} {'a-b%':>7s} {'b-a%':>7s} {'A/A%':>7s} {'B/B%':>7s}")
    for key in sorted(legs["A1"]):
        before = stats(legs["A1"][key] + legs["A2"][key])
        after = stats(legs["B1"][key] + legs["B2"][key])
        aa = pct(stats(legs["A4"][key])["p50"], stats(legs["A3"][key])["p50"])
        bb = pct(stats(legs["B2"][key])["p50"], stats(legs["B1"][key])["p50"])
        report["/".join(key)] = {"before": before, "after": after,
                                 "p50_delta_pct": pct(after["p50"], before["p50"]),
                                 "p50_inverse_pct": pct(before["p50"], after["p50"]),
                                 "aa_floor_p50_pct": aa, "bb_floor_p50_pct": bb}
        print(f"{key[0]:40s} {key[1]:13s} {before['p50']:12d} {after['p50']:12d} "
              f"{pct(after['p50'], before['p50']):7.2f} {pct(before['p50'], after['p50']):7.2f} {aa:7.2f} {bb:7.2f}")
    json.dump(report, open(f"{root}/summary.json", "w"), indent=1)

main()
