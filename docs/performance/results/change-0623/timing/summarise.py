#!/usr/bin/env python3
"""Summarise the ABBA timing legs of change 0611.

Usage: summarise.py <work-dir> <kind>

Prints, per case, each leg's p50/mean/p95/p99, the paired deltas in both
directions (A1 vs B1 and A2 vs B2), and the A/A floor measured in the same
window (A3 vs A4, and A1 vs A2).
"""

import json
import sys
from pathlib import Path

STATS = ("p50", "mean", "p95", "p99")


def load(work: Path, kind: str, leg: str):
    path = work / f"{kind}-{leg}.json"
    data = json.loads(path.read_text())
    out = {}
    for result in data["results"]:
        key = (result["case"], result.get("corpus", {}).get("name"))
        out[key] = result["elapsed_ns"]
    return out


def pct(before: float, after: float) -> float:
    """Percentage change from `before` to `after`; negative is faster."""
    return (after - before) / before * 100.0


def main() -> None:
    work = Path(sys.argv[1])
    kind = sys.argv[2]
    legs = {leg: load(work, kind, f"{kind}-{leg}") for leg in ("A1", "B1", "B2", "A2", "A3", "A4")}
    cases = sorted(legs["A1"])
    for case in cases:
        name = f"{case[0]} [{case[1]}]"
        print(f"== {name}")
        for leg in ("A1", "B1", "B2", "A2", "A3", "A4"):
            elapsed = legs[leg][case]
            values = " ".join(f"{stat}={elapsed[stat] / 1000.0:.1f}us" for stat in STATS)
            print(f"   {leg:<3} {values}")
        for label, before, after in (
            ("paired A1->B1", "A1", "B1"),
            ("paired A2->B2", "A2", "B2"),
            ("floor  A3->A4", "A3", "A4"),
            ("floor  A1->A2", "A1", "A2"),
        ):
            deltas = " ".join(
                f"{stat}={pct(legs[before][case][stat], legs[after][case][stat]):+.2f}%"
                for stat in STATS
            )
            print(f"   {label}: {deltas}")


if __name__ == "__main__":
    main()
