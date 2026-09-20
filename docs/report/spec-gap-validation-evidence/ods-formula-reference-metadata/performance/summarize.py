#!/usr/bin/env python3
"""Summarize matched controls and reference-metadata receipts."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from run_profile import BASELINE_COMMIT, MATCHED_CONTROL_CASES

HERE = Path(__file__).resolve().parent
RESULTS = HERE / "results"
BASELINE = f"baseline-{BASELINE_COMMIT}"
CANDIDATE = "candidate-final"


def rows(directory: Path) -> list[dict[str, Any]]:
    path = directory / "measurements.jsonl"
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def p50(values: list[int]) -> int:
    ordered = sorted(values)
    return ordered[(len(ordered) - 1) * 50 // 100]


def grouped(records: list[dict[str, Any]]) -> dict[tuple[str, str], list[dict[str, Any]]]:
    result: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for record in records:
        result.setdefault((record["case"], record["phase"]), []).append(record)
    return result


def stats(group: list[dict[str, Any]]) -> dict[str, Any]:
    first = group[0]
    result: dict[str, Any] = {
        "case": first["case"],
        "operation": first.get("operation"),
        "phase": first["phase"],
        "shape": first.get("shape"),
        "rows": first.get("rows"),
        "columns": first.get("columns"),
        "elements": first.get("elements"),
        "repeat": first.get("repeat"),
        "samples": len(group),
    }
    for source, target in (
        ("elapsed_ns_per_repeat", "time_ns_per_repeat"),
        ("allocator_calls_p50", "allocator_calls"),
        ("requested_bytes_p50", "requested_bytes"),
        ("released_bytes_p50", "released_bytes"),
        ("peak_live_delta_p50", "peak_live_bytes"),
        ("memory_retained_p50", "result_live_budget"),
        ("work_per_repeat", "work_per_repeat"),
        ("reference_reads_per_repeat", "reference_reads"),
        ("rss_kib", "rss_kib"),
    ):
        values = [int(record[source]) for record in group if record.get(source) is not None]
        result[target] = p50(values) if values else None
    result["checksum"] = p50([int(record["checksum_p50"]) for record in group])
    result["input_bytes"] = first.get("input_bytes")
    result["output_bytes_p50"] = first.get("output_bytes_p50")
    result["bytes_per_repeat_p50"] = first.get("bytes_per_repeat_p50")
    result["expected"] = first.get("expected")
    return result


def delta(before: int | None, after: int | None) -> str:
    if before in (None, 0) or after is None:
        return "n/a"
    return f"{(after - before) * 100.0 / before:+.1f}%"


def fmt(value: int | None) -> str:
    return "n/a" if value is None else f"{value:,}"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, default=RESULTS)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    results = args.results.resolve()
    output = args.output.resolve() if args.output else results / "performance-report.md"
    baseline = {key: stats(value) for key, value in grouped(rows(results / BASELINE)).items()}
    candidate = {key: stats(value) for key, value in grouped(rows(results / CANDIDATE)).items()}
    matched = sorted(set(baseline) & set(candidate))
    controls = [key for key in matched if key[0] in MATCHED_CONTROL_CASES]
    candidate_only = [key for key in sorted(candidate) if key[0] not in MATCHED_CONTROL_CASES]
    lines = [
        "# ODS reference-metadata evaluator performance profile",
        "",
        f"The baseline is committed `{BASELINE_COMMIT}`. The profile uses three warmups and fifteen fresh child processes in both evaluator phases; every row below is the p50 across those fresh children with time, work, and resolver reads normalized by the fixed repeat count.",
        "",
        "## Matched controls",
        "",
        "| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline bytes/repeat | candidate bytes/repeat | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |",
        "| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for key in controls:
        before = baseline[key]
        after = candidate[key]
        lines.append(
            f"| {key[0]} | {key[1]} | {fmt(before['time_ns_per_repeat'])} | {fmt(after['time_ns_per_repeat'])} | {delta(before['time_ns_per_repeat'], after['time_ns_per_repeat'])} | {fmt(before['bytes_per_repeat_p50'])} | {fmt(after['bytes_per_repeat_p50'])} | {fmt(before['allocator_calls'])} | {fmt(after['allocator_calls'])} | {fmt(before['rss_kib'])} | {fmt(after['rss_kib'])} |"
        )
    lines.extend(
        [
            "",
            "## Byte-position and projected-reducer workloads",
            "",
            "The candidate matrix covers all eight reference-metadata functions over contiguous and 3-D reference descriptors, ordered reference lists, current coordinates, inline arrays, projected matrix coordinates, scalar refusal, large geometry, typed resource limits, and cancellation. Input/output bytes and the reviewed domain labels remain with the raw case receipts.",
            "",
            "| case | phase | input bytes | output bytes p50 | bytes/repeat | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |",
            "| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
        ]
    )
    for key in candidate_only:
        value = candidate[key]
        lines.append(
            f"| {key[0]} | {key[1]} | {fmt(value['input_bytes'])} | {fmt(value['output_bytes_p50'])} | {fmt(value['bytes_per_repeat_p50'])} | {fmt(value['time_ns_per_repeat'])} | {fmt(value['work_per_repeat'])} | {fmt(value['reference_reads'])} | {fmt(value['allocator_calls'])} | {fmt(value['requested_bytes'])} | {fmt(value['peak_live_bytes'])} | {fmt(value['result_live_budget'])} | {fmt(value['rss_kib'])} |"
        )
    lines.extend(
        [
            "",
            "The resolver is an immutable borrowing fixture. Each child validates one typed text, number, or logical result (or typed failure) before timing the evaluator and drop path. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity.",
            "",
        ]
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text("\n".join(lines), encoding="utf-8")
    summary = {
        "baseline_commit": BASELINE_COMMIT,
        "baseline_groups": len(baseline),
        "candidate_groups": len(candidate),
        "matched_control_groups": len(controls),
        "candidate_only_groups": len(candidate_only),
        "controls": [{"baseline": baseline[key], "candidate": candidate[key]} for key in controls],
        "candidate_only": [candidate[key] for key in candidate_only],
    }
    output.with_suffix(".json").write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
