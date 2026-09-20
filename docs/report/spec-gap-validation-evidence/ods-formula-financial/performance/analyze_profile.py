#!/usr/bin/env python3
"""Validate and summarize a retained ODS financial capture.

Financial rows unsupported by the old ODS baseline remain capability records.
They are reported as candidate-only cost/scaling observations and never enter
before/after deltas.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
import json
import math
from pathlib import Path
import random
import statistics
from typing import Any

HERE = Path(__file__).resolve().parent
MATRIX_PATH = HERE / "case-matrix.json"
THRESHOLD = 0.05
BOOTSTRAP_SEED = 20260920
BOOTSTRAP_REPLICATES = 100_000
PHASES = ("evaluate", "parse-evaluate")


class AnalysisError(RuntimeError):
    pass


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise AnalysisError(f"cannot load {path}: {error}") from error


def finite(value: Any, label: str) -> float:
    try:
        result = float(value)
    except (TypeError, ValueError) as error:
        raise AnalysisError(f"{label}: expected number") from error
    if not math.isfinite(result):
        raise AnalysisError(f"{label}: expected finite number")
    return result


def matrix(path: Path) -> tuple[dict[str, dict[str, Any]], set[str], set[str]]:
    value = load(path)
    controls = value.get("matched_controls")
    candidate = value.get("candidate_cases")
    if not isinstance(controls, list) or not isinstance(candidate, list):
        raise AnalysisError("matrix case lists are malformed")
    rows = {row["name"]: row for row in controls + candidate}
    control_names = {row["name"] for row in controls}
    candidate_names = {row["name"] for row in candidate}
    if len(rows) != len(controls) + len(candidate):
        raise AnalysisError("matrix contains duplicate case names")
    return rows, control_names, candidate_names


def records(path: Path) -> list[dict[str, Any]]:
    file = path / "measurements.jsonl"
    if not file.is_file():
        raise AnalysisError(f"missing measurements: {file}")
    result = []
    for line in file.read_text(encoding="utf-8").splitlines():
        if line.strip():
            row = json.loads(line)
            if not isinstance(row, dict):
                raise AnalysisError(f"non-object measurement in {file}")
            result.append(row)
    return result


def flatten_samples(record: dict[str, Any], matrix_row: dict[str, Any]) -> list[dict[str, Any]]:
    samples = record.get("samples")
    if not isinstance(samples, list) or not samples:
        raise AnalysisError(f"{record.get('case')}: missing sample list")
    repeat = record.get("repeat")
    if not isinstance(repeat, int) or repeat <= 0:
        raise AnalysisError(f"{record.get('case')}: invalid repeat")
    expected_repeat = matrix_row["repeat"]
    if repeat != expected_repeat:
        raise AnalysisError(
            f"{record.get('case')}: capture repeat {repeat} differs from matrix {expected_repeat}"
        )
    flattened = []
    for index, sample in enumerate(samples, 1):
        if not isinstance(sample, dict):
            raise AnalysisError(f"{record.get('case')}: sample {index} is not an object")
        elapsed = finite(sample.get("elapsed_ns"), f"{record.get('case')} sample elapsed")
        checksum = sample.get("checksum")
        if not isinstance(checksum, int):
            raise AnalysisError(f"{record.get('case')}: sample checksum is not integer")
        for field in (
            "alloc_calls",
            "dealloc_calls",
            "requested_bytes",
            "released_bytes",
            "live_before",
            "live_after",
            "peak_live_delta",
            "work",
            "memory_retained",
            "reference_reads",
            "output_bytes",
        ):
            if not isinstance(sample.get(field), int) or sample[field] < 0:
                raise AnalysisError(f"{record.get('case')} sample {field} is invalid")
        expected_reads = matrix_row["reference_reads"] * repeat
        if matrix_row["class"] == "financial-cancellation":
            expected_reads = matrix_row["reference_reads"]
        if sample["reference_reads"] != expected_reads:
            raise AnalysisError(
                f"{record.get('case')}: expected read accounting {expected_reads}, observed {sample['reference_reads']}"
            )
        flattened.append(
            {
                "elapsed_ns_per_repeat": elapsed / repeat,
                "rss_kib": sample.get("rss_kib"),
                "checksum": checksum,
                "reference_reads": sample["reference_reads"],
                "output_bytes": sample["output_bytes"],
                "work": sample["work"],
                "memory_retained": sample["memory_retained"],
                "alloc_calls": sample["alloc_calls"],
                "requested_bytes": sample["requested_bytes"],
                "released_bytes": sample["released_bytes"],
            }
        )
    checksums = {sample["checksum"] for sample in flattened}
    if len(checksums) != 1:
        raise AnalysisError(f"{record.get('case')}: checksum changes between samples")
    return flattened


def percentile(values: list[float], fraction: float) -> float:
    ordered = sorted(values)
    if not ordered:
        raise AnalysisError("empty sample set")
    index = min(len(ordered) - 1, int((len(ordered) - 1) * fraction))
    return ordered[index]


def summarize(samples: list[dict[str, Any]]) -> dict[str, Any]:
    latency = [sample["elapsed_ns_per_repeat"] for sample in samples]
    return {
        "samples": len(samples),
        "elapsed_ns_p50": percentile(latency, 0.50),
        "elapsed_ns_mean": statistics.fmean(latency),
        "elapsed_ns_p95": percentile(latency, 0.95),
        "elapsed_ns_p99": percentile(latency, 0.99),
        "reference_reads": samples[0]["reference_reads"],
        "work_p50": percentile([float(sample["work"]) for sample in samples], 0.50),
        "memory_retained_p50": percentile([float(sample["memory_retained"]) for sample in samples], 0.50),
        "allocator_calls_p50": percentile([float(sample["alloc_calls"]) for sample in samples], 0.50),
        "requested_bytes_p50": percentile([float(sample["requested_bytes"]) for sample in samples], 0.50),
        "released_bytes_p50": percentile([float(sample["released_bytes"]) for sample in samples], 0.50),
        "output_bytes_p50": percentile([float(sample["output_bytes"]) for sample in samples], 0.50),
        "checksum": samples[0]["checksum"],
    }


def bootstrap_delta(baseline: list[float], candidate: list[float]) -> dict[str, Any]:
    if len(baseline) != len(candidate) or not baseline:
        raise AnalysisError("paired bootstrap requires equal nonempty sample sets")
    observed = statistics.median(candidate) / statistics.median(baseline) - 1.0
    rng = random.Random(BOOTSTRAP_SEED)
    values = []
    size = len(baseline)
    for _ in range(BOOTSTRAP_REPLICATES):
        indices = [rng.randrange(size) for _ in range(size)]
        before = statistics.median(baseline[index] for index in indices)
        after = statistics.median(candidate[index] for index in indices)
        values.append(after / before - 1.0 if before else 0.0)
    values.sort()
    lower = values[int(0.025 * len(values))]
    upper = values[int(0.975 * len(values))]
    return {
        "delta": observed,
        "bootstrap_seed": BOOTSTRAP_SEED,
        "bootstrap_replicates": BOOTSTRAP_REPLICATES,
        "confidence": 0.95,
        "ci95": [lower, upper],
    }


def analyze(baseline_dir: Path, candidate_dir: Path, matrix_path: Path) -> dict[str, Any]:
    rows, controls, candidate_names = matrix(matrix_path)
    all_records = {"baseline": records(baseline_dir), "candidate": records(candidate_dir)}
    grouped: dict[str, dict[tuple[str, str], list[dict[str, Any]]]] = {"baseline": {}, "candidate": {}}
    for label, values in all_records.items():
        for record in values:
            name = record.get("case")
            phase = record.get("phase")
            if name not in rows or phase not in PHASES:
                raise AnalysisError(f"{label}: unknown case/phase {name!r}/{phase!r}")
            if (name, phase) in grouped[label]:
                raise AnalysisError(f"duplicate {label} row {name}/{phase}")
            grouped[label][(name, phase)] = flatten_samples(record, rows[name])
    expected_baseline = {(name, phase) for name in controls for phase in PHASES}
    expected_candidate = {(name, phase) for name in rows for phase in PHASES}
    baseline_unsupported = set()
    baseline_summary = load(baseline_dir / "preflight.json")
    for name in baseline_summary.get("unsupported", []):
        baseline_unsupported.add(name)
    expected_baseline -= {(name, phase) for name in baseline_unsupported for phase in PHASES}
    if set(grouped["baseline"]) != expected_baseline:
        raise AnalysisError("baseline sample geometry differs from matched-control matrix")
    if set(grouped["candidate"]) != expected_candidate:
        raise AnalysisError("candidate sample geometry differs from financial matrix")

    matched = []
    flags = []
    for name in sorted(controls):
        for phase in PHASES:
            before = grouped["baseline"][(name, phase)]
            after = grouped["candidate"][(name, phase)]
            if len(before) != len(after):
                raise AnalysisError(f"{name}/{phase}: sample counts differ")
            if before[0]["reference_reads"] != after[0]["reference_reads"]:
                raise AnalysisError(f"{name}/{phase}: matched read accounting differs")
            comparison = bootstrap_delta(
                [sample["elapsed_ns_per_repeat"] for sample in before],
                [sample["elapsed_ns_per_repeat"] for sample in after],
            )
            before_summary = summarize(before)
            after_summary = summarize(after)
            row = {
                "case": name,
                "phase": phase,
                "baseline": before_summary,
                "candidate": after_summary,
                **comparison,
            }
            matched.append(row)
            rss_before = [sample.get("rss_kib") for sample in before if sample.get("rss_kib") is not None]
            rss_after = [sample.get("rss_kib") for sample in after if sample.get("rss_kib") is not None]
            if rss_before and rss_after:
                rss_delta = statistics.median(rss_after) / statistics.median(rss_before) - 1.0
                row["rss_delta"] = rss_delta
            if abs(comparison["delta"]) >= THRESHOLD or abs(row.get("rss_delta", 0.0)) >= THRESHOLD:
                flags.append({"case": name, "phase": phase, "latency_delta": comparison["delta"], "rss_delta": row.get("rss_delta")})

    candidate_only = []
    for name in sorted(candidate_names):
        if name in baseline_unsupported:
            reason = "baseline-unsupported"
        else:
            reason = "candidate-only-financial-lane"
        for phase in PHASES:
            candidate_only.append({
                "case": name,
                "phase": phase,
                "reason": reason,
                "candidate": summarize(grouped["candidate"][(name, phase)]),
            })
    return {
        "schema": 1,
        "contract_sha256": load(matrix_path).get("contract_sha256"),
        "baseline_unsupported": sorted(baseline_unsupported),
        "matched": matched,
        "candidate_only": candidate_only,
        "flags": flags,
        "threshold": THRESHOLD,
        "bootstrap_seed": BOOTSTRAP_SEED,
        "bootstrap_replicates": BOOTSTRAP_REPLICATES,
        "disposition": "descriptive; unsupported baseline lanes have no speedup delta",
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--baseline", type=Path, required=True)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--matrix", type=Path, default=MATRIX_PATH)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = analyze(args.baseline, args.candidate, args.matrix)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": "ok", "matched": len(result["matched"]), "candidate_only": len(result["candidate_only"]), "flags": len(result["flags"])}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
