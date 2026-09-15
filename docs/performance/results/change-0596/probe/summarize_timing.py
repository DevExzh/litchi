#!/usr/bin/env python3
"""Pool the paired timing legs and report per-scenario percentiles and deltas.

usage: summarize_timing.py <timing-dir>
  reads a1,a2,a3,a4 (before binary) and b1,b2 (after binary)
  A = a1 + a2 pooled, B = b1 + b2 pooled, A/A floor = a3 vs a4
"""
import json, sys, os, statistics

d = sys.argv[1]

def load(tag):
    with open(os.path.join(d, f"{tag}.json")) as f:
        doc = json.load(f)
    out = {}
    for record in doc["results"]:
        key = (record["case"], record["corpus"]["shape"], record["corpus"]["name"])
        out.setdefault(key, []).extend(record["elapsed_ns"]["samples"])
    return doc, out

def pct(values, q):
    values = sorted(values)
    if not values:
        return float("nan")
    index = min(len(values) - 1, max(0, int(round((q / 100.0) * (len(values) - 1)))))
    return values[index]

def stats(values):
    return {
        "n": len(values),
        "p50": pct(values, 50),
        "mean": statistics.fmean(values),
        "p95": pct(values, 95),
        "p99": pct(values, 99),
    }

legs = {}
docs = {}
for tag in ("a1", "a2", "a3", "a4", "b1", "b2"):
    docs[tag], legs[tag] = load(tag)

print(f"host={docs['a1']['environment']['cpu_model']} kernel={docs['a1']['environment']['kernel']} "
      f"cpu_affinity={docs['a1']['environment']['cpu_affinity']} rustc={docs['a1']['environment']['rustc_version']}")
print(f"before sha256={docs['a1']['binary_identity']['binary_sha256']}")
print(f"after  sha256={docs['b1']['binary_identity']['binary_sha256']}")
print(f"samples per leg per scenario={docs['a1']['configuration']['samples_per_case']} "
      f"warmup={docs['a1']['configuration']['warmup_iterations_per_case']}")
print()
header = f"{'scenario':52} {'metric':6} {'before ns':>12} {'after ns':>12} {'B vs A':>9} {'A vs B':>9} {'A/A':>8}"
print(header)
print("-" * len(header))
keys = sorted(set(legs["a1"]) | set(legs["b1"]))
for key in keys:
    a = stats(legs["a1"].get(key, []) + legs["a2"].get(key, []))
    b = stats(legs["b1"].get(key, []) + legs["b2"].get(key, []))
    f3 = stats(legs["a3"].get(key, []))
    f4 = stats(legs["a4"].get(key, []))
    label = f"{key[0]}/{key[1]}"
    for metric in ("p50", "mean", "p95", "p99"):
        ba = 100.0 * (b[metric] - a[metric]) / a[metric]
        ab = 100.0 * (a[metric] - b[metric]) / b[metric]
        floor = 100.0 * (f4[metric] - f3[metric]) / f3[metric]
        print(f"{label:52} {metric:6} {a[metric]:12,.0f} {b[metric]:12,.0f} {ba:+8.2f}% {ab:+8.2f}% {floor:+7.2f}%")
    print(f"{'':52} {'n':6} {a['n']:12,} {b['n']:12,}")
