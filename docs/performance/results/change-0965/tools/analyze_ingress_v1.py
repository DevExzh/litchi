#!/usr/bin/env python3
"""Replay descriptive ingress-tail evidence from the valid 0964 native packet.

This helper is intentionally read-only.  It consumes only the epoch-2 native
child reports, validates the packet's bound plan and report hashes, and emits
bounded JSON containing nearest-rank quantiles, adjacent paired legs, sample
windows, launch-order trends, and within-sample phase correlations.  It does
not read allocator reports, rerun a workload, or make an adoption decision.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import statistics
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


PHASES = (
    "ingress_ns",
    "edit_commit_ns",
    "publication_ns",
    "commit_drop_ns",
    "total_ns",
)
QUANTILES = (0.50, 0.95, 0.99)
WINDOWS = tuple((start, start + 10) for start in range(0, 60, 10))
EXPECTED_CASES = (
    "medium-one-percent",
    "dense-sparse-one-percent",
    "medium-one-edit",
)


class EvidenceError(RuntimeError):
    pass


def read_json(path: Path) -> Any:
    try:
        with path.open("r", encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, json.JSONDecodeError) as exc:
        raise EvidenceError(f"cannot read JSON {path}: {exc}") from exc


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as exc:
        raise EvidenceError(f"cannot hash {path}: {exc}") from exc
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    try:
        stat = path.stat()
    except OSError as exc:
        raise EvidenceError(f"cannot stat {path}: {exc}") from exc
    return {
        "path": str(path),
        "bytes": stat.st_size,
        "sha256": sha256(path),
    }


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def nearest_rank(values: Iterable[int | float], q: float) -> int | float:
    ordered = sorted(values)
    require(bool(ordered), "quantile requested for an empty sample")
    rank = max(1, math.ceil(q * len(ordered)))
    return ordered[rank - 1]


def number(value: int | float) -> int | float:
    """Keep integral counters compact while preserving floating means."""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def quantiles(values: Iterable[int | float]) -> dict[str, int | float]:
    values = list(values)
    return {
        "p50": number(nearest_rank(values, 0.50)),
        "p95": number(nearest_rank(values, 0.95)),
        "p99": number(nearest_rank(values, 0.99)),
    }


def range_summary(values: Iterable[int | float]) -> dict[str, int | float]:
    values = list(values)
    return {
        "min": number(min(values)),
        "max": number(max(values)),
        "mean": number(statistics.fmean(values)),
    }


def median(values: Iterable[int | float]) -> int | float:
    return number(statistics.median(list(values)))


def pearson(xs: list[float], ys: list[float]) -> float | None:
    require(len(xs) == len(ys), "correlation vectors have different lengths")
    if len(xs) < 2:
        return None
    mx = statistics.fmean(xs)
    my = statistics.fmean(ys)
    dx = [x - mx for x in xs]
    dy = [y - my for y in ys]
    denominator = math.sqrt(sum(x * x for x in dx) * sum(y * y for y in dy))
    if denominator == 0:
        return None
    return number(sum(x * y for x, y in zip(dx, dy)) / denominator)


def rank(values: list[float]) -> list[float]:
    ordered = sorted((value, index) for index, value in enumerate(values))
    result = [0.0] * len(values)
    cursor = 0
    while cursor < len(ordered):
        end = cursor + 1
        while end < len(ordered) and ordered[end][0] == ordered[cursor][0]:
            end += 1
        average = (cursor + 1 + end) / 2.0
        for _, index in ordered[cursor:end]:
            result[index] = average
        cursor = end
    return result


def spearman(xs: list[float], ys: list[float]) -> float | None:
    return pearson(rank(xs), rank(ys))


def iso_timestamp(value: str) -> float:
    try:
        return datetime.fromisoformat(value).timestamp()
    except ValueError as exc:
        raise EvidenceError(f"invalid ISO timestamp {value!r}") from exc


def phase_quantiles(samples: list[dict[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for phase in PHASES:
        values = [sample[phase] for sample in samples]
        require(all(isinstance(value, (int, float)) and value >= 0 for value in values),
                f"invalid values in phase {phase}")
        result[phase] = {
            "quantiles_ns": quantiles(values),
            "range_ns": range_summary(values),
            "sample_window_medians_ns": [
                median(values[start:end]) for start, end in WINDOWS
            ],
        }
    return result


def validate_child(
    packet_root: Path,
    expected: dict[str, Any],
    report_hash: str,
) -> dict[str, Any]:
    child_id = expected["child_id"]
    child_root = packet_root / "children" / child_id
    require(child_root.is_dir(), f"missing child directory {child_root}")
    start_path = child_root / "start.json"
    launch_path = child_root / "launch.json"
    terminal_path = child_root / "terminal.json"
    validation_path = child_root / "validation.json"
    report_path = child_root / "output" / "report.json"
    for path in (start_path, launch_path, terminal_path, validation_path, report_path):
        require(path.is_file(), f"missing child artifact {path}")
    start = read_json(start_path)
    launch = read_json(launch_path)
    terminal = read_json(terminal_path)
    validation = read_json(validation_path)
    report = read_json(report_path)
    actual = start.get("child")
    require(actual == expected, f"plan/start child mismatch for {child_id}")
    require(launch.get("observed_affinity") == [12, 13, 14, 15],
            f"affinity mismatch for {child_id}")
    require(launch.get("expected_affinity") == [12, 13, 14, 15],
            f"expected affinity mismatch for {child_id}")
    require(launch.get("affinity_verified_before_wait") is True,
            f"affinity verification missing for {child_id}")
    require(terminal.get("status") == "completed" and terminal.get("return_code") == 0,
            f"child not successful for {child_id}")
    require(terminal.get("terminal_before_report_parse") is True,
            f"terminal-before-parse missing for {child_id}")
    require(validation.get("status") == "passed", f"validation failed for {child_id}")
    require(validation.get("report", {}).get("sha256") == report_hash,
            f"validation/report hash mismatch for {child_id}")
    require(sha256(report_path) == report_hash, f"report bytes changed for {child_id}")
    require(report.get("schema") == "litchi.xlsx-managed-compression-benchmark.v1",
            f"unexpected report schema for {child_id}")
    require(report.get("status") == "completed", f"report incomplete for {child_id}")
    case = report.get("case")
    require(isinstance(case, dict), f"missing case in report for {child_id}")
    samples = case.get("samples")
    require(isinstance(samples, list) and len(samples) == 60,
            f"expected 60 samples for {child_id}")
    require([sample.get("sample") for sample in samples] == list(range(60)),
            f"sample sequence mismatch for {child_id}")
    require(report.get("warmup_iterations") == 10 and report.get("measured_samples") == 60,
            f"sample contract mismatch for {child_id}")
    require(case.get("id") == expected["case_id"].replace("-", "-") + "-workers1",
            f"case id mismatch for {child_id}")
    return {
        "expected": expected,
        "start": start,
        "launch": launch,
        "terminal": terminal,
        "validation": validation,
        "report": report,
        "report_path": report_path,
        "samples": samples,
    }


def summary_stats(samples: list[dict[str, Any]]) -> dict[str, Any]:
    return phase_quantiles(samples)


def case_role_rows(rows: list[dict[str, Any]], case_id: str, role: str) -> list[dict[str, Any]]:
    return [row for row in rows if row["case_id"] == case_id and row["binary_role"] == role]


def aggregate_sample_windows(rows: list[dict[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for phase in PHASES:
        sample_values = [
            [sample[phase] for sample in row["samples"]] for row in rows
        ]
        require(all(len(values) == 60 for values in sample_values), "sample count mismatch")
        medians_by_sample = [median(values[index] for values in sample_values) for index in range(60)]
        pooled = [value for values in sample_values for value in values]
        result[phase] = {
            "sample_index_medians_ns": medians_by_sample,
            "window_medians_ns": [median(medians_by_sample[start:end]) for start, end in WINDOWS],
            "window_pooled_p95_ns": [
                number(nearest_rank([values[index] for values in sample_values for index in range(start, end)], 0.95))
                for start, end in WINDOWS
            ],
            "pooled_sample_quantiles_ns": quantiles(pooled),
        }
    return result


def phase_correlation(rows: list[dict[str, Any]]) -> dict[str, Any]:
    vectors = {
        phase: [sample[phase] for row in rows for sample in row["samples"]]
        for phase in PHASES
    }
    matrix: dict[str, dict[str, float | None]] = {}
    for left in PHASES:
        matrix[left] = {}
        for right in PHASES:
            matrix[left][right] = pearson(vectors[left], vectors[right])
    residuals = [
        sample["total_ns"] - sum(sample[phase] for phase in PHASES if phase != "total_ns")
        for row in rows for sample in row["samples"]
    ]
    ingress_fraction = [
        sample["ingress_ns"] / sample["total_ns"]
        for row in rows for sample in row["samples"]
    ]
    return {
        "observation_count": len(vectors[PHASES[0]]),
        "pearson": matrix,
        "total_minus_phase_sum_ns": {
            "min": number(min(residuals)),
            "max": number(max(residuals)),
            "median": median(residuals),
            "mean": number(statistics.fmean(residuals)),
        },
        "ingress_fraction_of_total": {
            "quantiles": {
                key: number(value)
                for key, value in quantiles(ingress_fraction).items()
            },
            "min": number(min(ingress_fraction)),
            "max": number(max(ingress_fraction)),
        },
    }


def child_order_summary(rows: list[dict[str, Any]]) -> dict[str, Any]:
    rows = sorted(rows, key=lambda row: row["start_epoch_seconds"])
    for order, row in enumerate(rows):
        row["global_start_order"] = order
    by_case: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        by_case[row["case_id"]].append(row)
    for case_rows in by_case.values():
        for order, row in enumerate(case_rows):
            row["case_start_order"] = order
    trends: dict[str, Any] = {}
    for case_id, case_rows in by_case.items():
        case_trends: dict[str, Any] = {}
        for role in ("baseline", "candidate"):
            role_rows = [row for row in case_rows if row["binary_role"] == role]
            role_trends: dict[str, Any] = {}
            for phase in PHASES:
                orders = [float(row["case_start_order"]) for row in role_rows]
                p99s = [float(row["stats"][phase]["quantiles_ns"]["p99"]) for row in role_rows]
                p95s = [float(row["stats"][phase]["quantiles_ns"]["p95"]) for row in role_rows]
                p50s = [float(row["stats"][phase]["quantiles_ns"]["p50"]) for row in role_rows]
                role_trends[phase] = {
                    "child_count": len(role_rows),
                    "spearman_case_start_order_vs_child_p50": spearman(orders, p50s),
                    "spearman_case_start_order_vs_child_p95": spearman(orders, p95s),
                    "spearman_case_start_order_vs_child_p99": spearman(orders, p99s),
                    "child_order_values": [
                        {
                            "child_id": row["child_id"],
                            "case_start_order": row["case_start_order"],
                            "block": row["block"],
                            "position": row["position"],
                            "p50_ns": row["stats"][phase]["quantiles_ns"]["p50"],
                            "p95_ns": row["stats"][phase]["quantiles_ns"]["p95"],
                            "p99_ns": row["stats"][phase]["quantiles_ns"]["p99"],
                        }
                        for row in role_rows
                    ],
                }
            case_trends[role] = role_trends
        trends[case_id] = case_trends
    return {
        "children_in_start_order": [
            {
                "child_id": row["child_id"],
                "case_id": row["case_id"],
                "block": row["block"],
                "position": row["position"],
                "binary_role": row["binary_role"],
                "start_utc": row["start_utc"],
                "global_start_order": row["global_start_order"],
                "case_start_order": row["case_start_order"],
            }
            for row in rows
        ],
        "case_trends": trends,
    }


def raw_tail(rows: list[dict[str, Any]], limit: int = 30) -> dict[str, Any]:
    observations = []
    for row in rows:
        for sample in row["samples"]:
            observations.append({
                "ingress_ns": sample["ingress_ns"],
                "sample": sample["sample"],
                "child_id": row["child_id"],
                "case_id": row["case_id"],
                "binary_role": row["binary_role"],
                "block": row["block"],
                "position": row["position"],
                "global_start_order": row["global_start_order"],
                "edit_commit_ns": sample["edit_commit_ns"],
                "publication_ns": sample["publication_ns"],
                "commit_drop_ns": sample["commit_drop_ns"],
                "total_ns": sample["total_ns"],
            })
    observations.sort(key=lambda item: (-item["ingress_ns"], item["child_id"], item["sample"]))
    top = observations[:limit]
    by_case_role: dict[str, Any] = {}
    for case_id in EXPECTED_CASES:
        case_result: dict[str, Any] = {}
        for role in ("baseline", "candidate"):
            role_observations = [
                item for item in observations
                if item["case_id"] == case_id and item["binary_role"] == role
            ]
            case_result[role] = {
                "observation_count": len(role_observations),
                "top_10_ingress_samples_ns": role_observations[:10],
                "top_sample_index_counts": {
                    str(index): sum(item["sample"] == index for item in role_observations[:18])
                    for index in sorted({item["sample"] for item in role_observations[:18]})
                },
                "top_18_sample_index_range": {
                    "min": min(item["sample"] for item in role_observations[:18]),
                    "max": max(item["sample"] for item in role_observations[:18]),
                },
            }
        by_case_role[case_id] = case_result
    return {
        "top_ingress_observations": top,
        "top_5_percent_definition": "18 of 360 observations per case and role",
        "by_case_role": by_case_role,
    }


def build_summary(root: Path, plan_path: Path) -> dict[str, Any]:
    plan = read_json(plan_path)
    require(plan.get("schema") == "litchi.performance.0964.inline-number-screen-plan.v1",
            "unexpected plan schema")
    require(plan.get("plan_revision") == "epoch2-root-final-v1", "unexpected plan revision")
    require(plan.get("status") == "authorized" and plan.get("root_go", {}).get("status") == "GO",
            "plan is not the authorized epoch2 GO plan")
    require(plan.get("scope", {}).get("capture_epoch") == "epoch2-affinity-hardened",
            "wrong capture epoch")
    native_root = root / "capture-native-epoch2-v1"
    require(native_root.is_dir(), f"missing native packet {native_root}")
    packet_path = native_root / "completed-packet.json"
    start_packet_path = native_root / "capture-start.json"
    terminal_packet_path = native_root / "capture-terminal.json"
    packet = read_json(packet_path)
    capture_start = read_json(start_packet_path)
    capture_terminal = read_json(terminal_packet_path)
    require(packet.get("schema") == "litchi.performance.0964.capture.v1.completed-packet",
            "unexpected completed packet schema")
    require(packet.get("lane") == "native" and packet.get("status") == "completed",
            "native packet is not completed")
    require(capture_start.get("lane") == "native" and capture_start.get("serial") is True,
            "native capture start contract mismatch")
    require(capture_start.get("plan", {}).get("path") == str(plan_path),
            "capture start is bound to a different plan path")
    require(capture_start.get("plan", {}).get("sha256") == sha256(plan_path),
            "capture start is bound to a different plan hash")
    require(capture_terminal.get("status") == "completed" and capture_terminal.get("terminal_before_parse") is True,
            "native capture terminal contract mismatch")
    require(capture_terminal.get("capture_start", {}).get("sha256") == sha256(start_packet_path),
            "capture terminal is bound to a different capture start")
    require(capture_terminal.get("children_expected") == 36 and capture_terminal.get("children_completed") == 36,
            "native child count mismatch")
    expected_children = plan.get("lanes", {}).get("native", {}).get("children", [])
    require(len(expected_children) == 36, "plan does not bind 36 native children")
    packet_children = packet.get("children", [])
    packet_by_id = {child.get("child_id"): child for child in packet_children}
    require(len(packet_by_id) == 36, "packet does not contain exactly 36 native children")
    rows: list[dict[str, Any]] = []
    report_bindings: list[dict[str, Any]] = []
    for expected in expected_children:
        child_id = expected["child_id"]
        bound = packet_by_id.get(child_id)
        require(bound is not None, f"packet missing planned child {child_id}")
        report_path = native_root / "children" / child_id / "output" / "report.json"
        report_hash = sha256(report_path)
        validation_path = native_root / "children" / child_id / "validation.json"
        require(bound.get("validation", {}).get("sha256") == sha256(validation_path),
                f"packet/validation binding mismatch for {child_id}")
        record = validate_child(native_root, expected, report_hash)
        child = expected
        report = record["report"]
        row = {
            "child_id": child_id,
            "case_id": child["case_id"],
            "shape": child["shape"],
            "workload": child["workload"],
            "binary_role": child["binary_role"],
            "role": child["role"],
            "block": child["block"],
            "position": child["position"],
            "start_utc": record["start"]["started_utc"],
            "start_epoch_seconds": iso_timestamp(record["start"]["started_utc"]),
            "samples": report["case"]["samples"],
            "stats": summary_stats(report["case"]["samples"]),
            "report": descriptor(report_path),
            "validation": descriptor(native_root / "children" / child_id / "validation.json"),
        }
        rows.append(row)
        report_bindings.append(row["report"] | {"child_id": child_id})
    require({row["case_id"] for row in rows} == set(EXPECTED_CASES), "unexpected case set")
    require(all(sum(row["case_id"] == case for row in rows) == 12 for case in EXPECTED_CASES),
            "each case must have 12 children")
    require(all(sum(row["case_id"] == case and row["binary_role"] == role for row in rows) == 6
                for case in EXPECTED_CASES for role in ("baseline", "candidate")),
            "each case must have six children per binary role")

    rows.sort(key=lambda row: row["start_epoch_seconds"])
    order = child_order_summary(rows)
    pair_rows: list[dict[str, Any]] = []
    by_case_block: dict[tuple[str, int], list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        by_case_block[(row["case_id"], row["block"])].append(row)
    for case_id in EXPECTED_CASES:
        for block in range(3):
            block_rows = sorted(by_case_block[(case_id, block)], key=lambda row: row["position"])
            require([row["position"] for row in block_rows] == [0, 1, 2, 3],
                    f"frozen positions missing for {case_id} block {block}")
            for pair_index, pair in enumerate((block_rows[:2], block_rows[2:])):
                baseline = next(row for row in pair if row["binary_role"] == "baseline")
                candidate = next(row for row in pair if row["binary_role"] == "candidate")
                phases: dict[str, Any] = {}
                for phase in PHASES:
                    baseline_values = [sample[phase] for sample in baseline["samples"]]
                    candidate_values = [sample[phase] for sample in candidate["samples"]]
                    baseline_q = quantiles(baseline_values)
                    candidate_q = quantiles(candidate_values)
                    deltas = [c - b for b, c in zip(baseline_values, candidate_values)]
                    ratios = [c / b if b else None for b, c in zip(baseline_values, candidate_values)]
                    ratio_values = [value for value in ratios if value is not None]
                    phases[phase] = {
                        "absolute_quantiles_ns": {
                            "baseline": baseline_q,
                            "candidate": candidate_q,
                        },
                        "quantile_delta_candidate_minus_baseline_ns": {
                            key: number(candidate_q[key] - baseline_q[key])
                            for key in ("p50", "p95", "p99")
                        },
                        "quantile_ratio_candidate_over_baseline": {
                            key: number(candidate_q[key] / baseline_q[key])
                            for key in ("p50", "p95", "p99")
                        },
                        "paired_raw_sample_delta_ns": quantiles(deltas),
                        "paired_raw_sample_ratio": quantiles(ratio_values),
                    }
                pair_rows.append({
                    "case_id": case_id,
                    "block": block,
                    "pair_index": pair_index,
                    "baseline_child": baseline["child_id"],
                    "candidate_child": candidate["child_id"],
                    "scheduled_positions": [row["position"] for row in pair],
                    "scheduled_binary_roles": [row["binary_role"] for row in pair],
                    "phases": phases,
                })

    case_summary: dict[str, Any] = {}
    phase_correlations: dict[str, Any] = {}
    for case_id in EXPECTED_CASES:
        case_rows = [row for row in rows if row["case_id"] == case_id]
        case_result: dict[str, Any] = {}
        for role in ("baseline", "candidate"):
            role_rows = case_role_rows(rows, case_id, role)
            role_result: dict[str, Any] = {}
            for phase in PHASES:
                values_by_quantile = {
                    key: [row["stats"][phase]["quantiles_ns"][key] for row in role_rows]
                    for key in ("p50", "p95", "p99")
                }
                role_result[phase] = {
                    "child_quantile_distribution_ns": {
                        key: {
                            "min": min(values),
                            "median": median(values),
                            "max": max(values),
                        }
                        for key, values in values_by_quantile.items()
                    },
                    "sample_windows": aggregate_sample_windows(role_rows)[phase],
                }
            case_result[role] = role_result
            phase_correlations[f"{case_id}:{role}"] = phase_correlation(role_rows)
        case_summary[case_id] = case_result
        phase_correlations[f"{case_id}:all_roles"] = phase_correlation(case_rows)

    tail = raw_tail(rows)
    bindings = {
        "plan": descriptor(plan_path),
        "capture_start": descriptor(start_packet_path),
        "capture_terminal": descriptor(terminal_packet_path),
        "completed_packet": descriptor(packet_path),
        "report_count": len(report_bindings),
        "reports": sorted(report_bindings, key=lambda item: item["child_id"]),
    }
    return {
        "schema": "litchi.performance.0965.ingress-evidence.v1",
        "status": "descriptive_read_only",
        "decision": {
            "adoption_allowed": False,
            "screen_reopened": False,
            "causal_claim": False,
            "purpose": "diagnostic design for a future fresh ingress-tail experiment",
        },
        "scope": {
            "source": "valid 0964 epoch2 native raw reports only",
            "lane": "native",
            "children": len(rows),
            "cases": list(EXPECTED_CASES),
            "children_per_case": 12,
            "children_per_role_per_case": 6,
            "samples_per_child": 60,
            "warmup_iterations": 10,
            "phases": list(PHASES),
            "quantile_method": "nearest-rank; rank=ceil(q*n), one-indexed",
            "nearest_rank_for_n60": {"p50": 30, "p95": 57, "p99": 60},
            "pooled_samples_used_for_decision": False,
            "allocator_reports_read": False,
            "workloads_rerun": False,
        },
        "bindings": bindings,
        "child_summaries": [
            {
                key: row[key]
                for key in (
                    "child_id", "case_id", "shape", "workload", "binary_role", "role",
                    "block", "position", "start_utc", "global_start_order", "case_start_order",
                    "stats", "report", "validation",
                )
            }
            for row in sorted(rows, key=lambda row: row["global_start_order"])
        ],
        "paired_legs": pair_rows,
        "case_role_summary": case_summary,
        "launch_order": order,
        "phase_correlations": phase_correlations,
        "ingress_tail": tail,
        "interpretation_limits": [
            "Each native child has 60 measured samples after 10 untimed warmups; nearest-rank p99 is the maximum sample in that child.",
            "Six paired legs and their quantiles are descriptive; no confidence interval or pooled-sample effect is computed.",
            "Launch order, sample index, phase correlation, and tail co-occurrence are diagnostics only and do not establish cause.",
            "The existing 0964 safety gate remains binding; this evidence does not relax or requalify the rejected screen.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True,
                        help="external 0964 evidence root")
    parser.add_argument("--plan", type=Path, required=True,
                        help="bound epoch2 final plan JSON")
    parser.add_argument("--output", type=Path, required=True,
                        help="new JSON output path")
    args = parser.parse_args()
    try:
        result = build_summary(args.root.resolve(), args.plan.resolve())
        args.output.parent.mkdir(parents=True, exist_ok=True)
        with args.output.open("w", encoding="utf-8", newline="\n") as stream:
            json.dump(result, stream, indent=2, sort_keys=True)
            stream.write("\n")
    except EvidenceError as exc:
        parser.error(str(exc))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
