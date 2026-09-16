#!/usr/bin/env python3
"""Paired-timing summary for change 0634: per-leg p50/mean/p95/p99 and the
paired deltas in both directions, beside the A/A floor measured in the same
window."""
import json, os, statistics, sys

S = os.environ.get("SCRATCH", os.path.dirname(os.path.abspath(__file__)))
T = os.path.join(S, "timing")
LEGS = ["A1", "B1", "B2", "A2", "AA1", "AA2"]


def pct(a, b):
    return (b - a) / a * 100.0


def stats(samples):
    s = sorted(samples)
    n = len(s)
    return {
        "p50": s[n // 2],
        "mean": statistics.fmean(s),
        "p95": s[min(n - 1, int(round(0.95 * (n - 1))))],
        "p99": s[min(n - 1, int(round(0.99 * (n - 1))))],
    }


def emit(title, per_leg):
    print(f"\n### {title}")
    print(f"{'leg':<5}{'p50':>12}{'mean':>12}{'p95':>12}{'p99':>12}")
    for leg in LEGS:
        if leg not in per_leg:
            continue
        v = per_leg[leg]
        print(f"{leg:<5}{v['p50']:>12,.0f}{v['mean']:>12,.1f}{v['p95']:>12,.0f}{v['p99']:>12,.0f}")
    if not all(leg in per_leg for leg in ("A1", "A2", "B1", "B2")):
        return
    a = statistics.median([per_leg["A1"]["p50"], per_leg["A2"]["p50"]])
    b = statistics.median([per_leg["B1"]["p50"], per_leg["B2"]["p50"]])
    floor = ""
    if "AA1" in per_leg and "AA2" in per_leg:
        aa = abs(pct(per_leg["AA1"]["p50"], per_leg["AA2"]["p50"]))
        floor = f"   A/A floor p50 {aa:.2f}% (AA1 {per_leg['AA1']['p50']:,.0f} vs AA2 {per_leg['AA2']['p50']:,.0f})"
    print(f"A p50 {a:,.0f}  B p50 {b:,.0f}  before->after {pct(a, b):+.2f}%  after->before {pct(b, a):+.2f}%{floor}")


# Registered harness selectors.
per_case = {}
for leg in LEGS:
    path = os.path.join(T, f"{leg}.json")
    if not os.path.exists(path):
        continue
    doc = json.load(open(path))
    for entry in doc["results"]:
        key = (entry["case"], entry["corpus"]["name"])
        per_case.setdefault(key, {})[leg] = stats(entry["elapsed_ns"]["samples"])
for key in sorted(per_case):
    emit(f"selector {key[0]} [corpus {key[1]}]", per_case[key])

# Whole-operation timings on the real fixtures.
per_fx = {}
for name in sorted(os.listdir(T)):
    if not name.startswith("fx-") or not name.endswith(".txt"):
        continue
    stem = name[3:-4]
    leg = stem.rsplit("-", 1)[1]
    label = stem.rsplit("-", 1)[0]
    samples = [int(line) for line in open(os.path.join(T, name)) if line.strip()]
    if samples:
        per_fx.setdefault(label, {})[leg] = stats(samples)
for key in sorted(per_fx):
    emit(f"fixture {key}", per_fx[key])
