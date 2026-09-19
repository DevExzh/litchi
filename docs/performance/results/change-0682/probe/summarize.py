#!/usr/bin/env python3
"""Summarize change 0682 JSONL samples without dropping raw evidence."""

from __future__ import annotations

import json
import math
import pathlib
import statistics
import sys
from collections import defaultdict


def percentile(values: list[float], fraction: float) -> float:
    values = sorted(values)
    if not values:
        return float("nan")
    rank = fraction * (len(values) - 1)
    low = math.floor(rank)
    high = math.ceil(rank)
    if low == high:
        return values[low]
    return values[low] + (values[high] - values[low]) * (rank - low)


def stats(rows: list[dict]) -> dict:
    elapsed = [float(row["elapsed_ns"]) for row in rows]
    allocations = [float(row["allocations"]) for row in rows]
    allocated_bytes = [float(row["allocated_bytes"]) for row in rows]
    return {
        "n": len(rows),
        "elapsed_ns": {
            "p50": percentile(elapsed, 0.50),
            "p95": percentile(elapsed, 0.95),
            "p99": percentile(elapsed, 0.99),
            "mean": statistics.fmean(elapsed),
        },
        "allocations": {
            "p50": percentile(allocations, 0.50),
            "p95": percentile(allocations, 0.95),
            "mean": statistics.fmean(allocations),
        },
        "allocated_bytes": {
            "p50": percentile(allocated_bytes, 0.50),
            "p95": percentile(allocated_bytes, 0.95),
            "mean": statistics.fmean(allocated_bytes),
        },
        # Grouping uses the content hash below, so an absolute checkout path
        # in a real-fixture row cannot split before/after or ABBA samples.
        "corpus_labels": sorted({row["corpus"] for row in rows}),
        "input_sha256": rows[0]["input_sha256"],
        "result_digests": sorted({row["result_digest"] for row in rows}),
        "source_versions": sorted(
            {(row.get("source_version_id"), row.get("source_revision")) for row in rows}
        ),
    }


def load(path: pathlib.Path) -> list[dict]:
    return [json.loads(line) for line in path.read_text().splitlines() if line.strip()]


def key(row: dict) -> tuple:
    return row["route"], row["mode"], row["input_sha256"], row["repetitions"]


def main() -> int:
    if len(sys.argv) != 3:
        raise SystemExit("usage: summarize.py <aa-or-abba.jsonl> <summary.json>")
    source = pathlib.Path(sys.argv[1])
    destination = pathlib.Path(sys.argv[2])
    rows = load(source)
    grouped: dict[tuple, list[dict]] = defaultdict(list)
    for row in rows:
        grouped[key(row)].append(row)

    output = {
        "schema": "change-0682-summary-v1",
        "source": str(source),
        "rows": len(rows),
        "groups": {"/".join(map(str, group)): stats(values) for group, values in sorted(grouped.items())},
    }
    if any("leg" in row for row in rows):
        aa = {}
        for group, values in sorted(grouped.items()):
            by_leg: dict[str, list[dict]] = defaultdict(list)
            for row in values:
                by_leg[row["leg"]].append(row)
            if {"A1", "A2"}.issubset(by_leg):
                a1 = stats(by_leg["A1"])
                a2 = stats(by_leg["A2"])
                aa["/".join(map(str, group))] = {
                    "A1": a1,
                    "A2": a2,
                    "p50_delta_percent": 100.0
                    * (a2["elapsed_ns"]["p50"] / a1["elapsed_ns"]["p50"] - 1.0),
                    "p95_delta_percent": 100.0
                    * (a2["elapsed_ns"]["p95"] / a1["elapsed_ns"]["p95"] - 1.0),
                }
            if {"A1", "A2", "B1", "B2"}.issubset(by_leg):
                a1 = stats(by_leg["A1"])
                a2 = stats(by_leg["A2"])
                b1 = stats(by_leg["B1"])
                b2 = stats(by_leg["B2"])
                aa["/".join(map(str, group))].update(
                    {
                        "B1": b1,
                        "B2": b2,
                        "before_after_p50_percent": 100.0
                        * (b1["elapsed_ns"]["p50"] / a1["elapsed_ns"]["p50"] - 1.0),
                        "after_before_p50_percent": 100.0
                        * (b2["elapsed_ns"]["p50"] / a2["elapsed_ns"]["p50"] - 1.0),
                    }
                )
        output["paired"] = aa
    destination.write_text(json.dumps(output, indent=2, sort_keys=True) + "\n")
    print(f"wrote {destination} ({len(grouped)} groups)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
