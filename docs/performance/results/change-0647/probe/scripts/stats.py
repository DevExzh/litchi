#!/usr/bin/env python3
"""Change 0647: summarize one paired-timing directory.

Reads `before.txt`, `after.txt`, `floor-a.txt` and `floor-b.txt` (one sample in
nanoseconds per line, as `reuse_iters --timing` prints them) and reports p50,
mean, p95 and p99 for each leg, the paired delta in both directions, and the
A/A floor measured in the same window.

Usage: stats.py <timing-dir> [label]
"""

import pathlib
import statistics
import sys


def load(path):
    return sorted(
        int(line) for line in path.read_text().split() if line.strip().isdigit()
    )


def pct(values, q):
    if not values:
        return 0
    index = min(len(values) - 1, int(round(q * (len(values) - 1))))
    return values[index]


def describe(name, values):
    return (
        f"{name:<10} n={len(values):<5} "
        f"p50={pct(values, 0.50) / 1000.0:10.2f}us "
        f"mean={statistics.fmean(values) / 1000.0:10.2f}us "
        f"p95={pct(values, 0.95) / 1000.0:10.2f}us "
        f"p99={pct(values, 0.99) / 1000.0:10.2f}us"
    )


def delta(label, base, other):
    """`other` relative to `base`, at each quantile, in both directions."""
    rows = []
    for tag, q in (("p50", 0.50), ("mean", None), ("p95", 0.95), ("p99", 0.99)):
        if q is None:
            left, right = statistics.fmean(base), statistics.fmean(other)
        else:
            left, right = pct(base, q), pct(other, q)
        forward = (right - left) / left * 100.0 if left else 0.0
        backward = (left - right) / right * 100.0 if right else 0.0
        rows.append(f"  {tag:<5} {forward:+7.2f}% (inverse {backward:+7.2f}%)")
    return f"{label}\n" + "\n".join(rows)


def main():
    root = pathlib.Path(sys.argv[1])
    label = sys.argv[2] if len(sys.argv) > 2 else root.name
    before = load(root / "before.txt")
    after = load(root / "after.txt")
    floor_a = load(root / "floor-a.txt")
    floor_b = load(root / "floor-b.txt")

    print(f"### {label}")
    print(describe("before", before))
    print(describe("after", after))
    print(describe("floor-a", floor_a))
    print(describe("floor-b", floor_b))
    print(delta("after relative to before:", before, after))
    print(delta("A/A floor (floor-b relative to floor-a):", floor_a, floor_b))


if __name__ == "__main__":
    main()
