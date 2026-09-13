#!/usr/bin/env python3
"""Summarize the change-0558 OLE2 guardrail matrix.

Each case ran once per ABBA leg with the same corpus generator. A case is
reported as improved only when both directions agree; anything else is printed
as a review trigger rather than folded into an aggregate.
"""

from __future__ import annotations

import argparse
import json
import pathlib
import statistics

LEGS = ("A1", "C1", "C2", "A2")


def quantile(values: list[int], fraction: float) -> float:
    ordered = sorted(values)
    if len(ordered) == 1:
        return float(ordered[0])
    position = fraction * (len(ordered) - 1)
    low = int(position)
    high = min(low + 1, len(ordered) - 1)
    weight = position - low
    return ordered[low] * (1.0 - weight) + ordered[high] * weight


def cells(report: dict) -> dict:
    """Map (case, corpus) to elapsed statistics computed from retained samples."""
    found = {}
    for record in report.get("results", []):
        samples = record.get("elapsed_ns", {}).get("samples")
        if not samples:
            continue
        corpus = record.get("corpus") or {}
        key = (record["case"], str(corpus.get("name")))
        found[key] = {
            "p50_ns": quantile(samples, 0.50),
            "mean_ns": statistics.fmean(samples),
            "p95_ns": quantile(samples, 0.95),
            "p99_ns": quantile(samples, 0.99),
            "samples": len(samples),
        }
    return found


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--packet", default="docs/performance/results/change-0559/guardrail")
    parser.add_argument("--output", default="docs/performance/results/change-0559/guardrail-analysis.json")
    args = parser.parse_args()
    root = pathlib.Path(args.packet)

    names = sorted({path.name.split("-", 1)[1][: -len(".json")] for path in root.glob("A1-*.json")})
    rows = []
    for name in names:
        legs = {}
        for leg in LEGS:
            path = root / f"{leg}-{name}.json"
            if not path.exists():
                break
            legs[leg] = cells(json.loads(path.read_text()))
        if len(legs) != len(LEGS):
            rows.append({"case": name, "status": "incomplete"})
            continue
        keys = set(legs["A1"]) & set(legs["C1"]) & set(legs["C2"]) & set(legs["A2"])
        for key in sorted(keys):
            entry = {"case": name, "cell": list(key), "statistics": {}}
            for stat in ("p50_ns", "mean_ns", "p95_ns", "p99_ns"):
                if stat not in legs["A1"][key]:
                    continue
                first = (legs["C1"][key][stat] - legs["A1"][key][stat]) / legs["A1"][key][stat] * 100.0
                second = (legs["C2"][key][stat] - legs["A2"][key][stat]) / legs["A2"][key][stat] * 100.0
                entry["statistics"][stat] = {
                    "first_direction_percent": first,
                    "second_direction_percent": second,
                    "improves_both": first < 0.0 and second < 0.0,
                    "adverse_both": first > 0.0 and second > 0.0,
                    "adverse_both_over_5_percent": first > 5.0 and second > 5.0,
                }
            rows.append(entry)

    triggers = [
        {"case": row["case"], "cell": row["cell"], "statistic": stat, **value}
        for row in rows
        if "statistics" in row
        for stat, value in row["statistics"].items()
        if value["adverse_both_over_5_percent"]
    ]
    summary = {
        "schema_version": 1,
        "analysis_kind": "litchi-perf-change-0559-guardrail",
        "rows": rows,
        "review_triggers_adverse_both_over_5_percent": triggers,
    }
    pathlib.Path(args.output).write_text(json.dumps(summary, indent=2) + "\n")
    for row in rows:
        if "statistics" not in row:
            print(f"{row['case']:44s} {row['status']}")
            continue
        p50 = row["statistics"].get("p50_ns")
        if p50 is None:
            continue
        flag = "improves" if p50["improves_both"] else ("ADVERSE" if p50["adverse_both"] else "mixed")
        print(
            f"{row['case']:44s} {row['cell'][1][:14]:14s} p50 {p50['first_direction_percent']:+8.2f}% / "
            f"{p50['second_direction_percent']:+8.2f}%  {flag}"
        )
    print(f"\nreview triggers (adverse in both directions by more than 5%): {len(triggers)}")
    for trigger in triggers:
        print(f"  {trigger['case']} {trigger['cell']} {trigger['statistic']}: "
              f"{trigger['first_direction_percent']:+.2f}% / {trigger['second_direction_percent']:+.2f}%")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
