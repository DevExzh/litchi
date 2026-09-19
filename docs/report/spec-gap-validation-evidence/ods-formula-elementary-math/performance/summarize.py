#!/usr/bin/env python3
"""Write deterministic paired statistics for a completed profile."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
RESULTS = HERE / "results"
BASELINE = "baseline-8ef0057e5"
CANDIDATE = "candidate-final"
CONTROL_CASES = {
    "scalar-control-arithmetic",
    "scalar-control-round",
    "array-control-4x4-arithmetic",
    "array-control-4x4-round",
    "array-control-16x16-arithmetic",
    "array-control-16x16-round",
    "reference-scalar-arithmetic",
    "reference-scalar-round",
    "reference-array-arithmetic",
    "reference-array-round",
    "scalar-control-sin",
    "array-control-4x4-sin",
    "array-control-16x16-sin",
    "reference-scalar-sin",
    "reference-array-sin",
}

METRICS = {
    "allocator_calls_p50": "allocator_calls",
    "requested_bytes_p50": "requested_bytes",
    "released_bytes_p50": "released_bytes",
    "peak_live_delta_p50": "peak_live_bytes",
    "memory_retained_p50": "result_live_memory",
    "work_p50": "work",
    "rss_kib": "rss_kib",
}


def read_rows(directory: Path) -> list[dict[str, Any]]:
    path = directory / "measurements.jsonl"
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def p50(values: list[int]) -> int:
    ordered = sorted(values)
    return ordered[(len(ordered) - 1) * 50 // 100]


def grouped(rows: list[dict[str, Any]]) -> dict[tuple[str, str], list[dict[str, Any]]]:
    groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for row in rows:
        groups.setdefault((row["case"], row["phase"]), []).append(row)
    return groups


def group_stats(rows: list[dict[str, Any]]) -> dict[str, Any]:
    stats: dict[str, Any] = {
        "case": rows[0]["case"],
        "operation": rows[0]["operation"],
        "phase": rows[0]["phase"],
        "supported": all(bool(row["supported"]) for row in rows),
        "samples": len(rows),
    }
    stats["time_ns_per_repeat"] = p50([int(row["elapsed_ns_per_repeat"]) for row in rows])
    stats["time_ns_batch_p50"] = p50([int(row["elapsed_ns_p50"]) for row in rows])
    for source, name in METRICS.items():
        stats[name] = p50([int(row[source]) for row in rows])
    repeat = int(rows[0]["repeat"])
    stats["repeat"] = repeat
    stats["work_per_repeat"] = p50([int(row["work_per_repeat"]) for row in rows])
    return stats


def delta_percent(before: int, after: int) -> float | None:
    if before == 0:
        return None
    return (after - before) * 100.0 / before


def fmt_delta(value: float | None) -> str:
    return "n/a" if value is None else f"{value:+.1f}%"


def fmt_int(value: int | None) -> str:
    return "n/a" if value is None else f"{value:,}"


def markdown_controls(
    baseline: dict[tuple[str, str], dict[str, Any]],
    candidate: dict[tuple[str, str], dict[str, Any]],
) -> list[str]:
    lines = [
        "| case | phase | baseline time ns/repeat | candidate time ns/repeat | time delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |",
        "| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for key in sorted(baseline):
        before = baseline[key]
        after = candidate[key]
        lines.append(
            f"| {key[0]} | {key[1]} | {fmt_int(before['time_ns_per_repeat'])} | "
            f"{fmt_int(after['time_ns_per_repeat'])} | "
            f"{fmt_delta(delta_percent(before['time_ns_per_repeat'], after['time_ns_per_repeat']))} | "
            f"{fmt_int(before['allocator_calls'])} | {fmt_int(after['allocator_calls'])} | "
            f"{fmt_int(before['rss_kib'])} | {fmt_int(after['rss_kib'])} |"
        )
    return lines


def markdown_candidate_only(candidate: dict[tuple[str, str], dict[str, Any]]) -> list[str]:
    lines = [
        "| case | phase | supported | time ns/repeat | alloc calls | requested bytes | result-live memory | RSS KiB |",
        "| --- | --- | :---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for key in sorted(candidate):
        if key[0] in CONTROL_CASES:
            continue
        row = candidate[key]
        lines.append(
            f"| {key[0]} | {key[1]} | {'yes' if row['supported'] else 'no'} | "
            f"{fmt_int(row['time_ns_per_repeat'])} | {fmt_int(row['allocator_calls'])} | "
            f"{fmt_int(row['requested_bytes'])} | {fmt_int(row['result_live_memory'])} | "
            f"{fmt_int(row['rss_kib'])} |"
        )
    return lines


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, default=RESULTS)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    results = args.results.resolve()
    output = args.output.resolve() if args.output else results / "performance-report.md"
    baseline_groups = grouped(read_rows(results / BASELINE))
    candidate_groups = grouped(read_rows(results / CANDIDATE))
    baseline = {key: group_stats(rows) for key, rows in baseline_groups.items()}
    candidate = {key: group_stats(rows) for key, rows in candidate_groups.items()}
    matched = sorted(set(baseline) & set(candidate))
    controls = {key: baseline[key] for key in matched if key[0] in CONTROL_CASES}
    candidate_controls = {key: candidate[key] for key in matched if key[0] in CONTROL_CASES}
    candidate_only = {key: value for key, value in candidate.items() if key not in baseline}
    summary = {
        "baseline_groups": len(baseline),
        "candidate_groups": len(candidate),
        "matched_control_groups": len(controls),
        "candidate_only_groups": len(candidate_only),
        "controls": [
            {"baseline": controls[key], "candidate": candidate_controls[key]}
            for key in sorted(controls)
        ],
        "candidate_only": [candidate[key] for key in sorted(candidate_only)],
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        "# ODS elementary-math performance profile\n\n"
        "The baseline timing lane contains matched arithmetic/ROUND/trigonometric "
        "controls. The candidate lane contains all named scalar, array, and "
        "synthetic local-reference cases. New elementary-math cases have no "
        "baseline timing comparison because the baseline does not implement "
        "them. Time deltas are candidate minus baseline; positive values are "
        "slower. Samples are p50 across fresh child processes, and each row's "
        "time is normalized by the harness repeat count.\n\n"
        "## Matched controls\n\n"
        + "\n".join(markdown_controls(controls, candidate_controls))
        + "\n\n## Candidate elementary-math and representative workloads\n\n"
        + "\n".join(markdown_candidate_only(candidate))
        + "\n\nThe synthetic local-reference rows use the profile's immutable resolver. "
        "They do not measure the production worksheet adapter or full-workbook "
        "recalculation. The direct f64 oracle checks evaluator projection and "
        "is not an independent libm-accuracy implementation.\n",
        encoding="utf-8",
    )
    output.with_suffix(".json").write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
