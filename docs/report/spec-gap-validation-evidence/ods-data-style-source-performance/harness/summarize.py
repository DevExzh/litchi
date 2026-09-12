#!/usr/bin/env python3
"""Summarize the JSONL receipts emitted by the source-style profile."""

from __future__ import annotations

import json
import statistics
import sys
from pathlib import Path


def percentile(values: list[int], fraction: float) -> int:
    values = sorted(values)
    if not values:
        raise ValueError("empty measurement set")
    index = min(len(values) - 1, int((len(values) - 1) * fraction + 0.999999))
    return values[index]


def summarize(path: Path) -> dict[str, object]:
    rows = [
        json.loads(line)
        for line in path.read_text().splitlines()
        if line and json.loads(line).get("kind") == "measurement"
    ]
    groups: dict[tuple[int, str], list[dict[str, object]]] = {}
    for row in rows:
        groups.setdefault((int(row["scale"]), str(row["operation"])), []).append(row)

    def stats(group: list[dict[str, object]]) -> dict[str, object]:
        def values(name: str) -> list[int]:
            return [int(row[name]) for row in group]

        return {
            "iterations": len(group),
            "elapsed_ns": {
                "min": min(values("elapsed_ns")),
                "median": statistics.median(values("elapsed_ns")),
                "p95": percentile(values("elapsed_ns"), 0.95),
                "max": max(values("elapsed_ns")),
            },
            "allocations_median": statistics.median(values("allocations")),
            "deallocations_median": statistics.median(values("deallocations")),
            "requested_bytes_median": statistics.median(values("requested_bytes")),
            "released_bytes_median": statistics.median(values("released_bytes")),
            "net_live_bytes_delta_median": statistics.median(
                values("net_live_bytes_delta")
            ),
            "peak_live_bytes_delta_median": statistics.median(
                values("peak_live_bytes_delta")
            ),
            "copy_bytes_observed_median": statistics.median(
                values("copy_bytes_observed")
            ),
            "result_bytes_median": statistics.median(values("result_bytes")),
            "exact_noop_values": sorted({bool(row["exact_noop"]) for row in group}),
        }

    return {
        "input": path.name,
        "measurement_rows": len(rows),
        "groups": [
            {
                "scale": scale,
                "operation": operation,
                **stats(groups[(scale, operation)]),
            }
            for scale, operation in sorted(groups)
        ],
    }


def main() -> None:
    if len(sys.argv) not in (2, 3):
        raise SystemExit(f"usage: {sys.argv[0]} RAW.jsonl [SUMMARY.json]")
    summary = summarize(Path(sys.argv[1]))
    rendered = json.dumps(summary, indent=2, sort_keys=True) + "\n"
    if len(sys.argv) == 3:
        Path(sys.argv[2]).write_text(rendered)
    else:
        sys.stdout.write(rendered)


if __name__ == "__main__":
    main()
