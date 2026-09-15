#!/usr/bin/env python3
"""Summarise the A1 B1 B2 A2 timing runs: per-leg p50/mean/p95/p99 and paired deltas."""
import json, sys, os
SP = os.path.dirname(os.path.abspath(__file__))
T = os.path.join(SP, "out", "timing")

def load(prefix, run):
    path = os.path.join(T, f"{prefix}-{run}.json")
    with open(path) as handle:
        data = json.load(handle)
    rows = {}
    for row in data["results"]:
        key = (row["case"], row.get("cache_state") or "-", row["corpus"]["shape"])
        rows[key] = row["elapsed_ns"]
    return rows

def pct(new, old):
    return (new - old) / old * 100.0

for prefix in ("sem", "fs"):
    runs = {}
    ok = True
    for run in ("A1", "B1", "B2", "A2"):
        try:
            runs[run] = load(prefix, run)
        except FileNotFoundError:
            ok = False
    if not ok:
        print(f"[{prefix}] incomplete"); continue
    keys = sorted(set(runs["A1"]) & set(runs["B1"]) & set(runs["B2"]) & set(runs["A2"]))
    print(f"\n=== {prefix} ===")
    hdr = f"{'case':<44}{'shape':<12}{'A1 p50':>12}{'A2 p50':>12}{'B1 p50':>12}{'B2 p50':>12}{'A/A %':>9}{'B1-A1 %':>10}{'B2-A2 %':>10}{'pooled %':>10}"
    print(hdr)
    for key in keys:
        a1, a2 = runs["A1"][key], runs["A2"][key]
        b1, b2 = runs["B1"][key], runs["B2"][key]
        pooled_a = (a1["p50"] + a2["p50"]) / 2
        pooled_b = (b1["p50"] + b2["p50"]) / 2
        print(f"{key[0]:<44}{key[2]:<12}{a1['p50']:>12}{a2['p50']:>12}{b1['p50']:>12}{b2['p50']:>12}"
              f"{pct(a2['p50'], a1['p50']):>9.2f}{pct(b1['p50'], a1['p50']):>10.2f}"
              f"{pct(b2['p50'], a2['p50']):>10.2f}{pct(pooled_b, pooled_a):>10.2f}")
    print()
    print(f"{'case':<44}{'leg':<5}{'p50':>12}{'mean':>14}{'p95':>12}{'p99':>12}{'min':>12}")
    for key in keys:
        for run in ("A1", "B1", "B2", "A2"):
            row = runs[run][key]
            print(f"{key[0]:<44}{run:<5}{row['p50']:>12}{row['mean']:>14.1f}{row['p95']:>12}{row['p99']:>12}{row['min']:>12}")
