#!/usr/bin/env python3
"""Summarize the change 0627 OLE2 range-source baseline.

Reads the raw schema-1 reports in ``raw/`` and prints three tables: the
deterministic per-scenario counts, the paired owned-source A/A floor, and the
range-source against owned-source comparison. Standard library only.
"""

from __future__ import annotations

import gzip
import json
import sys
from pathlib import Path

FIXTURES = [
    ("withcustomviews", "WithCustomViews.xls"),
    ("conditionalformatting", "ConditionalFormattingSamples.xls"),
    ("54016", "54016.xls"),
    ("45543", "45543.ppt"),
]
SCENARIOS = [
    "open",
    "open+list-worksheets",
    "open+one-cell",
    "open+all-cells",
    "open+full-text",
    "open+one-shape-text",
]


def load(directory: Path, name: str) -> dict:
    """Read one leg's schema-1 report. Legs above 1 MB are retained gzipped."""
    path = directory / "raw" / f"{name}.json"
    compressed = path.with_suffix(".json.gz")
    if compressed.exists():
        report = json.loads(gzip.decompress(compressed.read_bytes()).decode("utf-8"))
    elif path.exists():
        report = json.loads(path.read_text(encoding="utf-8"))
    else:
        return {}
    out = {}
    for result in report["results"]:
        evidence = result["source"]["ole2_range_source"]
        out[evidence["scenario"]] = (result, evidence)
    return out


def unique(values: list[int]) -> str:
    """One number when every retained sample agreed, else the observed set."""
    distinct = sorted(set(values))
    if len(distinct) == 1:
        return f"{distinct[0]:,}"
    return "/".join(f"{value:,}" for value in distinct)


def percent(after: float, before: float) -> str:
    if before == 0:
        return "n/a"
    return f"{(after - before) / before * 100.0:+.2f}%"


def main() -> int:
    directory = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    print("## Deterministic counts (identical on both transports)\n")
    print("| fixture | scenario | logical reads | logical bytes | physical requests | "
          "physical bytes | sequence digest stable | observation |")
    print("|---|---|---:|---:|---:|---:|---|---|")
    for key, label in FIXTURES:
        owned = load(directory, f"owned-{key}")
        ranged = load(directory, f"range-{key}")
        for scenario in SCENARIOS:
            if scenario not in owned:
                continue
            _, owned_evidence = owned[scenario]
            calls = unique(owned_evidence["logical_read_calls"])
            byte_totals = unique(owned_evidence["logical_read_bytes"])
            requests = bytes_requested = "n/a"
            stable = "n/a"
            if scenario in ranged:
                _, range_evidence = ranged[scenario]
                requests = unique(range_evidence["physical_request_count"])
                bytes_requested = unique(range_evidence["physical_request_bytes"])
                stable = "yes" if range_evidence["request_sequence_identical"] else "NO"
                for field in ("logical_read_calls", "logical_read_bytes", "observation"):
                    if range_evidence[field] != owned_evidence[field]:
                        raise SystemExit(
                            f"{label} {scenario}: {field} differs between transports"
                        )
            observation = owned_evidence["observation"].replace("\n", "\\n")
            if len(observation) > 46:
                observation = observation[:43] + "..."
            print(
                f"| `{label}` | {scenario} | {calls} | {byte_totals} | {requests} "
                f"| {bytes_requested} | {stable} | `{observation}` |"
            )

    print("\n## Owned-source A/A floor (same selector, same window, second run vs first)\n")
    print("| fixture | scenario | A1 p50 (us) | A2 p50 (us) | delta p50 | delta p95 | delta p99 |")
    print("|---|---|---:|---:|---:|---:|---:|")
    worst = 0.0
    for key, label in FIXTURES:
        first = load(directory, f"owned-{key}")
        second = load(directory, f"aa-owned-{key}")
        for scenario in SCENARIOS:
            if scenario not in first or scenario not in second:
                continue
            a1 = first[scenario][0]["elapsed_ns"]
            a2 = second[scenario][0]["elapsed_ns"]
            worst = max(worst, abs((a2["p50"] - a1["p50"]) / a1["p50"] * 100.0))
            print(
                f"| `{label}` | {scenario} | {a1['p50'] / 1000.0:,.2f} | {a2['p50'] / 1000.0:,.2f} "
                f"| {percent(a2['p50'], a1['p50'])} | {percent(a2['p95'], a1['p95'])} "
                f"| {percent(a2['p99'], a1['p99'])} |"
            )
    print(f"\nWorst |A/A| at p50: {worst:.2f}%")

    print("\n## Range source against owned source (elapsed is modelled on the range leg)\n")
    print("| fixture | scenario | owned p50 (ms) | range p50 (ms) | ratio | "
          "modelled service floor (ms) | floor share of range p50 |")
    print("|---|---|---:|---:|---:|---:|---:|")
    for key, label in FIXTURES:
        owned = load(directory, f"owned-{key}")
        ranged = load(directory, f"range-{key}")
        for scenario in SCENARIOS:
            if scenario not in owned or scenario not in ranged:
                continue
            owned_p50 = owned[scenario][0]["elapsed_ns"]["p50"]
            range_p50 = ranged[scenario][0]["elapsed_ns"]["p50"]
            floor = ranged[scenario][1]["simulated_service_floor_ns"]
            print(
                f"| `{label}` | {scenario} | {owned_p50 / 1e6:,.3f} | {range_p50 / 1e6:,.3f} "
                f"| {range_p50 / owned_p50:,.1f}x | {floor / 1e6:,.3f} "
                f"| {floor / range_p50 * 100.0:.1f}% |"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
