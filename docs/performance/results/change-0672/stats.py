#!/usr/bin/env python3
"""Summarize one change 0672 ABBA run.

usage: stats.py RUNS_DIR TAG [TAG...]

Prints, per tag, the p50/mean/p95/p99 of each leg in nanoseconds and the paired
deltas in both directions (A1 against B1, and B2 against A2), as a percentage
reduction of the before leg.
"""

import statistics
import sys
from pathlib import Path


def load(path):
    return sorted(int(line) for line in path.read_text().split() if line.strip())


def quantile(values, fraction):
    if not values:
        return float("nan")
    index = min(len(values) - 1, max(0, round(fraction * (len(values) - 1))))
    return values[index]


def summary(values):
    return {
        "n": len(values),
        "p50": statistics.median(values),
        "mean": statistics.fmean(values),
        "p95": quantile(values, 0.95),
        "p99": quantile(values, 0.99),
    }


def main():
    runs = Path(sys.argv[1])
    for tag in sys.argv[2:]:
        legs = {}
        for name in ("a1", "b1", "b2", "a2"):
            path = runs / f"{tag}-{name}.txt"
            if not path.exists():
                print(f"{tag}: missing {path.name}")
                break
            legs[name] = summary(load(path))
        else:
            print(f"== {tag}")
            print(f"{'leg':<4}{'n':>5}{'p50':>14}{'mean':>14}{'p95':>14}{'p99':>14}")
            for name in ("a1", "b1", "b2", "a2"):
                s = legs[name]
                print(
                    f"{name:<4}{s['n']:>5}{s['p50']:>14,.0f}{s['mean']:>14,.1f}"
                    f"{s['p95']:>14,.0f}{s['p99']:>14,.0f}"
                )
            for before, after, label in (("a1", "b1", "A1->B1"), ("a2", "b2", "B2<-A2")):
                parts = []
                for metric in ("p50", "mean", "p95", "p99"):
                    old = legs[before][metric]
                    new = legs[after][metric]
                    parts.append(f"{metric} {100.0 * (old - new) / old:+.3f}%")
                parts.append(f"p50 speedup {legs[before]['p50'] / legs[after]['p50']:.2f}x")
                print(f"{label}: " + "  ".join(parts))
            print()


if __name__ == "__main__":
    main()
