#!/usr/bin/env python3
"""Summarise the ABBA legs of the delayed-transport open probe.

Usage: summarise.py <work-dir>

Prints, per fixture, the request count and median of each leg, the paired
deltas in both directions (A1 vs B1 and A2 vs B2) and the A/A floor measured
in the same window (A3 vs A4, and A1 vs A2).
"""
import json
import sys
from pathlib import Path

LEGS = ("A1", "B1", "B2", "A2", "A3", "A4")


def load(work: Path, leg: str):
    rows = {}
    for line in (work / f"{leg}.json").read_text().splitlines():
        if not line.strip():
            continue
        row = json.loads(line)
        rows[row["fixture"]] = row
    return rows


def pct(before: float, after: float) -> float:
    return (after - before) / before * 100.0


def main() -> None:
    work = Path(sys.argv[1])
    legs = {leg: load(work, leg) for leg in LEGS}
    for fixture in legs["A1"]:
        print(f"== {fixture}")
        for leg in LEGS:
            row = legs[leg][fixture]
            print(
                f"   {leg}: requests={row['requests_min']}..{row['requests_max']} "
                f"bytes={row['bytes_min']} p50={row['p50_us']:.1f}us "
                f"mean={row['mean_us']:.1f} p95={row['p95_us']:.1f} p99={row['p99_us']:.1f}"
            )
        a1, b1 = legs["A1"][fixture]["p50_us"], legs["B1"][fixture]["p50_us"]
        a2, b2 = legs["A2"][fixture]["p50_us"], legs["B2"][fixture]["p50_us"]
        a3, a4 = legs["A3"][fixture]["p50_us"], legs["A4"][fixture]["p50_us"]
        print(f"   A1->B1 p50 {pct(a1, b1):+.2f}%   A2->B2 p50 {pct(a2, b2):+.2f}%")
        print(
            f"   floor A3->A4 p50 {pct(a3, a4):+.2f}%   floor A1->A2 p50 {pct(a1, a2):+.2f}%"
            f"   floor B1->B2 p50 {pct(b1, b2):+.2f}%"
        )
        print(f"   saving A1-B1 {(a1 - b1) / 1000:.2f} ms   A2-B2 {(a2 - b2) / 1000:.2f} ms")
        print()


if __name__ == "__main__":
    main()
