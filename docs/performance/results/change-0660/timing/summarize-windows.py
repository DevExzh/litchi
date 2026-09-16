#!/usr/bin/env python3
"""Pool each leg's two runs per window and print p50/mean/p95/p99, both deltas
and the A/A and B/B floors observed in that window."""
import json, statistics, sys, os

def rows(path):
    out = {}
    for r in json.load(open(path))["results"]:
        out[(r["case"], r["corpus"].get("shape") or r["corpus"].get("name", "-"))] = r["elapsed_ns"]["samples"]
    return out

def stats(samples):
    s = sorted(samples); n = len(s)
    q = lambda p: s[min(n - 1, max(0, int(round(p * (n - 1)))))]
    return {"n": n, "p50": q(0.5), "mean": statistics.fmean(s), "p95": q(0.95), "p99": q(0.99)}

def pct(new, old):
    return (new - old) / old * 100.0

def main():
    root = sys.argv[1]
    report = {}
    for window in sorted(d for d in os.listdir(root) if d.startswith("w")):
        report[window] = {}
        for prefix, family in (("", "semantic"), ("ord-", "ordinary_save")):
            try:
                a1, b1, b2, a2 = (rows(f"{root}/{window}/{prefix}{leg}.json") for leg in ("A1", "B1", "B2", "A2"))
            except FileNotFoundError:
                continue
            print(f"\n== {window} / {family}")
            print(f"{'case':38s} {'shape':7s} {'before p50':>11s} {'after p50':>11s} {'d%':>7s} {'inv%':>7s} {'A/A%':>7s} {'B/B%':>7s}")
            for key in sorted(a1):
                before, after = stats(a1[key] + a2[key]), stats(b1[key] + b2[key])
                aa = pct(stats(a2[key])["p50"], stats(a1[key])["p50"])
                bb = pct(stats(b2[key])["p50"], stats(b1[key])["p50"])
                report[window]["/".join(key)] = {
                    "before": before, "after": after,
                    "p50_delta_pct": pct(after["p50"], before["p50"]),
                    "p50_inverse_pct": pct(before["p50"], after["p50"]),
                    "aa_floor_p50_pct": aa, "bb_floor_p50_pct": bb}
                print(f"{key[0]:38s} {key[1]:7s} {before['p50']:11d} {after['p50']:11d} "
                      f"{pct(after['p50'], before['p50']):7.2f} {pct(before['p50'], after['p50']):7.2f} {aa:7.2f} {bb:7.2f}")
    json.dump(report, open(f"{root}/paired-summary.json", "w"), indent=1)

main()
