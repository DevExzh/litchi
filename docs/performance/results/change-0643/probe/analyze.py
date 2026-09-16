#!/usr/bin/env python3
"""Summarise A1 B1 B2 A2 harness timing runs: per-leg p50/mean/p95/p99 and paired deltas."""
import json, sys, os

def load(directory, prefix, run):
    with open(os.path.join(directory, f"{prefix}-{run}.json")) as handle:
        data = json.load(handle)
    rows = {}
    for row in data["results"]:
        key = (row["case"], row.get("cache_state") or "-", row["corpus"]["shape"])
        rows[key] = row["elapsed_ns"]
    return rows

def pct(new, old):
    return (new - old) / old * 100.0

def main(directory, prefix):
    runs = {run: load(directory, prefix, run) for run in ("A1", "B1", "B2", "A2")}
    keys = sorted(set(runs["A1"]) & set(runs["B1"]) & set(runs["B2"]) & set(runs["A2"]))
    print(f"{'case':<42}{'A1 p50':>11}{'B1 p50':>11}{'B2 p50':>11}{'A2 p50':>11}{'A/A %':>9}{'B1-A1 %':>10}{'B2-A2 %':>10}{'pooled %':>10}")
    for key in keys:
        a1, a2 = runs["A1"][key], runs["A2"][key]
        b1, b2 = runs["B1"][key], runs["B2"][key]
        pa = (a1["p50"] + a2["p50"]) / 2
        pb = (b1["p50"] + b2["p50"]) / 2
        print(f"{key[0]:<42}{a1['p50']:>11}{b1['p50']:>11}{b2['p50']:>11}{a2['p50']:>11}"
              f"{pct(a2['p50'], a1['p50']):>9.2f}{pct(b1['p50'], a1['p50']):>10.2f}"
              f"{pct(b2['p50'], a2['p50']):>10.2f}{pct(pb, pa):>10.2f}")
    print()
    print(f"{'case':<42}{'leg':<5}{'p50':>12}{'mean':>14}{'p95':>12}{'p99':>12}{'min':>12}")
    for key in keys:
        for run in ("A1", "B1", "B2", "A2"):
            row = runs[run][key]
            print(f"{key[0]:<42}{run:<5}{row['p50']:>12}{row['mean']:>14.1f}{row['p95']:>12}{row['p99']:>12}{row['min']:>12}")

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "fs")
