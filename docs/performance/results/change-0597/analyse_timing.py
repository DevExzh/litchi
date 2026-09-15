#!/usr/bin/env python3
"""Change 0597: summarize the ABBA paired timing and the A/A floor."""
import json, os, sys, statistics

SC = "/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0597"
D = os.path.join(SC, "out", "timing")


def load(tag):
    path = os.path.join(D, f"{tag}.json")
    if not os.path.exists(path):
        return None
    out = {}
    for row in json.load(open(path))["results"]:
        key = (row["case"], row["corpus"]["shape"])
        e = row["elapsed_ns"]
        s = sorted(e["samples"])
        out[key] = {
            "p50": statistics.median(s),
            "mean": statistics.fmean(s),
            "p95": s[min(len(s) - 1, int(round(0.95 * (len(s) - 1))))],
            "p99": s[min(len(s) - 1, int(round(0.99 * (len(s) - 1))))],
            "n": len(s),
        }
    return out


def pct(a, b):
    return (a - b) / a * 100.0


def report(title, left, right, lname, rname):
    print(f"\n## {title}  ({lname} -> {rname}; negative = {rname} slower)")
    print(f"{'case / shape':<46} {'stat':<5} {lname:>12} {rname:>12} {'delta %':>9}")
    for key in sorted(set(left) & set(right)):
        for stat in ("p50", "mean", "p95", "p99"):
            l, r = left[key][stat], right[key][stat]
            print(f"{key[0]+' / '+key[1]:<46} {stat:<5} {l:>12.0f} {r:>12.0f} {pct(l, r):>9.3f}")


A1, B1, B2, A2 = (load(t) for t in ("A1", "B1", "B2", "A2"))
F = [load(t) for t in ("F1", "F2", "F3", "F4")]
if A1 and B1:
    report("leg 1: A1 (base) vs B1 (change)", A1, B1, "A1-base", "B1-chg")
if A2 and B2:
    report("leg 2: A2 (base) vs B2 (change)", A2, B2, "A2-base", "B2-chg")
if all(F):
    print("\n## A/A floor (four base runs in the same window)")
    print(f"{'case / shape':<46} {'stat':<5} {'spread %':>9}  (max-min)/min over F1..F4")
    for key in sorted(F[0]):
        for stat in ("p50", "mean", "p95", "p99"):
            vals = [f[key][stat] for f in F]
            print(f"{key[0]+' / '+key[1]:<46} {stat:<5} {(max(vals)-min(vals))/min(vals)*100:>9.3f}")
