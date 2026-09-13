#!/usr/bin/env python3
"""Summarize the change-0562 A/G/G/A latency matrix.

A cell is reported as improving only when both ABBA directions agree. Cells that
are adverse in both directions by more than five percent are listed explicitly
with their absolute values rather than folded into an aggregate.
"""

from __future__ import annotations

import argparse
import json
import pathlib
import statistics

LEGS = ("A1", "G1", "G2", "A2")
STATS = ("p50", "mean", "p95", "p99")


def quantile(values: list[int], fraction: float) -> float:
    ordered = sorted(values)
    if len(ordered) == 1:
        return float(ordered[0])
    position = fraction * (len(ordered) - 1)
    low = int(position)
    high = min(low + 1, len(ordered) - 1)
    weight = position - low
    return ordered[low] * (1.0 - weight) + ordered[high] * weight


def cells(path: pathlib.Path) -> dict:
    report = json.loads(path.read_text())
    found = {}
    for record in report.get("results", []):
        samples = record.get("elapsed_ns", {}).get("samples")
        if not samples:
            continue
        corpus = (record.get("corpus") or {}).get("name")
        found[(record["case"], str(corpus))] = {
            "p50": quantile(samples, 0.50),
            "mean": statistics.fmean(samples),
            "p95": quantile(samples, 0.95),
            "p99": quantile(samples, 0.99),
            "samples": len(samples),
        }
    return found


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--latency", default="docs/performance/results/change-0562/latency")
    parser.add_argument("--output", default="docs/performance/results/change-0562/latency-analysis.json")
    args = parser.parse_args()
    root = pathlib.Path(args.latency)

    names = sorted({path.name.split("-", 1)[1][: -len(".json")] for path in root.glob("A1-*.json")})
    rows, triggers = [], []
    for name in names:
        legs = {}
        for leg in LEGS:
            path = root / f"{leg}-{name}.json"
            if not path.exists():
                legs = {}
                break
            legs[leg] = cells(path)
        if not legs:
            rows.append({"case": name, "status": "incomplete"})
            continue
        for key in sorted(set.intersection(*(set(value) for value in legs.values()))):
            entry = {"case": key[0], "corpus": key[1], "statistics": {}}
            for stat in STATS:
                first = (legs["G1"][key][stat] - legs["A1"][key][stat]) / legs["A1"][key][stat] * 100.0
                second = (legs["G2"][key][stat] - legs["A2"][key][stat]) / legs["A2"][key][stat] * 100.0
                entry["statistics"][stat] = {
                    "first_direction_percent": first,
                    "second_direction_percent": second,
                    "improves_both": first < 0.0 and second < 0.0,
                    "adverse_both": first > 0.0 and second > 0.0,
                    "baseline_ns": legs["A1"][key][stat],
                    "candidate_ns": legs["G1"][key][stat],
                }
                if first > 5.0 and second > 5.0:
                    triggers.append({"case": key[0], "corpus": key[1], "statistic": stat,
                                     **entry["statistics"][stat]})
            rows.append(entry)

    complete = [row for row in rows if "statistics" in row]
    improve = sum(1 for row in complete for value in row["statistics"].values() if value["improves_both"])
    summary = {
        "schema_version": 1,
        "analysis_kind": "litchi-perf-change-0562-latency",
        "rows": rows,
        "summary": {
            "rows": len(complete),
            "comparisons": len(complete) * len(STATS),
            "improve_both_directions": improve,
            "review_triggers_adverse_both_over_5_percent": len(triggers),
        },
        "review_triggers": triggers,
    }
    pathlib.Path(args.output).write_text(json.dumps(summary, indent=2) + "\n")
    for row in complete:
        p50 = row["statistics"]["p50"]
        flag = "improves" if p50["improves_both"] else ("ADVERSE" if p50["adverse_both"] else "mixed")
        print(f"{row['case']:44s} {row['corpus'][:20]:20s} p50 {p50['first_direction_percent']:+7.2f}% / "
              f"{p50['second_direction_percent']:+7.2f}%  {flag}")
    print(f"\n{summary['summary']}")
    for trigger in triggers:
        print(f"  TRIGGER {trigger['case']} {trigger['corpus']} {trigger['statistic']}: "
              f"{trigger['first_direction_percent']:+.2f}%/{trigger['second_direction_percent']:+.2f}% "
              f"({trigger['baseline_ns']:.0f}ns -> {trigger['candidate_ns']:.0f}ns)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
