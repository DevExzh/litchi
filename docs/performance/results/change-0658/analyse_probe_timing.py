#!/usr/bin/env python3
"""Change 0658: summarize the probe's ineligible-read ABBA timing and A/A floor."""
import os, statistics

S = "/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0658"
D = os.path.join(S, "out", "timing")
ITERS = 20


def load(fixture, tag):
    path = os.path.join(D, f"probe-{fixture}-{tag}.txt")
    values = sorted(int(line) / ITERS for line in open(path) if line.strip())
    return {
        "p50": statistics.median(values),
        "mean": statistics.fmean(values),
        "p95": values[min(len(values) - 1, int(round(0.95 * (len(values) - 1))))],
        "p99": values[min(len(values) - 1, int(round(0.99 * (len(values) - 1))))],
        "n": len(values),
    }


def pct(a, b):
    return (a - b) / a * 100.0


print("Ineligible source-backed one-cell read, ns per read (20 reads per sample,")
print("30 samples per leg, 5 warmup samples, CPU 13).\n")
for fixture in ("control", "real"):
    a1, b1, b2, a2 = (load(fixture, t) for t in ("A1", "B1", "B2", "A2"))
    floor = [load(fixture, t) for t in ("F1", "F2", "F3", "F4")]
    print(f"## {fixture}")
    print(f"{'stat':<6} {'A1 base':>12} {'B1 chg':>12} {'d1 %':>8} "
          f"{'A2 base':>12} {'B2 chg':>12} {'d2 %':>8} {'A/A floor %':>12}")
    for stat in ("p50", "mean", "p95", "p99"):
        vals = [f[stat] for f in floor]
        spread = (max(vals) - min(vals)) / min(vals) * 100.0
        print(f"{stat:<6} {a1[stat]:>12.0f} {b1[stat]:>12.0f} {pct(a1[stat], b1[stat]):>8.2f} "
              f"{a2[stat]:>12.0f} {b2[stat]:>12.0f} {pct(a2[stat], b2[stat]):>8.2f} {spread:>12.2f}")
    print()
