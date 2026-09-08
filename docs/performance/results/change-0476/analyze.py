#!/usr/bin/env python3
"""Derive reproducible paired-arm evidence for the 0476 experiment.

This module reads only retained JSON artifacts in its own bundle.  It validates
the 24 formal PPTX lanes, preserves normal and allocator measurements as
separate observations, and emits comparisons whose scope is explicitly
descriptive.  A timing difference is a review flag; it is never presented as
a registered latency result.  Counter and transport-guard lanes are retained
in the summary but are not forced through the PPTX report schema.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

import report_checks

ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-0476-evidence-analysis-v1"
PROTOCOL_SCHEMA = "litchi-0476-protocol-v1"
MODES = ("normal", "allocator")
SHAPES = ("tiny", "medium", "large")
ARMS = ("control", "candidate")
REPEATS = ("R1", "R2")
FORMAL_COUNT = 24
REVIEW_PERCENT = 5.0
CASE = report_checks.CASE
SLIDES = report_checks.SHAPES


class AnalysisError(ValueError):
    """The retained evidence cannot support the declared comparison."""


def fail(message: str) -> None:
    raise AnalysisError(message)


def reject_constant(value: str) -> Any:
    raise AnalysisError(f"non-finite JSON value {value}")


def reject_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise AnalysisError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_pairs, parse_constant=reject_constant)
    except (OSError, UnicodeError, json.JSONDecodeError, AnalysisError) as error:
        raise AnalysisError(f"cannot read {path}: {error}") from error


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise AnalysisError(f"cannot canonicalize JSON: {error}") from error


def sha256_file(path: Path) -> tuple[str, int]:
    digest = hashlib.sha256()
    size = 0
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                size += len(block)
    except OSError as error:
        raise AnalysisError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest(), size


def binding(path: Path, root: Path) -> dict[str, Any]:
    digest, size = sha256_file(path)
    try:
        relative = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        relative = path.name
    return {"path": relative, "sha256": digest, "bytes": size}


def _field(row: Mapping[str, Any], *names: str, default: Any = None) -> Any:
    for name in names:
        if name in row:
            return row[name]
    return default


def lane_arm(row: Mapping[str, Any]) -> str | None:
    value = _field(row, "arm", "variant", "implementation", "build_arm")
    return value if value in ARMS else None


def lane_kind(row: Mapping[str, Any]) -> str:
    value = _field(row, "kind", "lane_kind", "measurement", default=row.get("suite", "formal"))
    return value.lower() if isinstance(value, str) else "formal"


def is_formal(row: Mapping[str, Any]) -> bool:
    return (
        lane_kind(row) in {"formal", "main", "pptx", "measurement"}
        and row.get("repeat") in REPEATS
        and row.get("mode") in MODES
        and row.get("shape") in SHAPES
        and lane_arm(row) in ARMS
    )


def _case_order(rows: Sequence[Mapping[str, Any]]) -> list[tuple[str, str]]:
    expected = [(mode, shape) for mode in MODES for shape in SHAPES]
    seen: list[tuple[str, str]] = []
    for row in rows:
        pair = (row.get("mode"), row.get("shape"))
        if pair not in expected:
            fail(f"unknown formal mode/shape {pair!r}")
        if not seen or seen[-1] != pair:
            seen.append(pair)
    if seen != expected:
        fail(f"R1 formal case order is {seen!r}, expected {expected!r}")
    return seen


def protocol_order(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    if protocol.get("schema") != PROTOCOL_SCHEMA:
        fail("protocol schema differs")
    for key, expected in (
        ("samples", 30),
        ("warmups", 3),
        ("workers", 1),
        ("cpu", 2),
        ("pilot_samples", 1),
        ("pilot_warmups", 0),
        ("large_requested_bytes_reduction_required_percent", 95),
        ("normal_regression_review_threshold_percent", 5),
        ("normal_repeat_review_threshold_percent", 5),
        ("peak_rss_regression_review_threshold_percent", 5),
    ):
        if protocol.get(key) != expected:
            fail(f"protocol.{key} differs")
    rows: list[dict[str, Any]] = []
    suites = (("order", "main"), ("pilot_order", "pilot"), ("counter_order", "counter"), ("guard_order", "guard"))
    for key, suite in suites:
        raw = protocol.get(key, [])
        if not isinstance(raw, list):
            fail(f"protocol.{key} must be an array")
        for index, value in enumerate(raw):
            if not isinstance(value, dict):
                fail(f"protocol.{key}[{index}] must be an object")
            row = dict(value)
            if not isinstance(row.get("lane"), str) or not row["lane"]:
                fail(f"protocol.{key}[{index}].lane is missing")
            row["suite"] = suite
            if suite == "guard":
                row.setdefault("guard_selectors", list(protocol.get("guard_selectors", [])))
            rows.append(row)
    formal = [row for row in rows if is_formal(row)]
    if len(formal) != FORMAL_COUNT:
        fail(f"protocol must contain exactly {FORMAL_COUNT} formal lanes, found {len(formal)}")
    lanes = [row["lane"] for row in rows]
    if len(set(lanes)) != len(lanes):
        fail("protocol lane names are not unique")
    for repeat in REPEATS:
        repeat_rows = [row for row in formal if row["repeat"] == repeat]
        if len(repeat_rows) != 12:
            fail(f"{repeat} must contain twelve formal lanes")
    r1 = [row for row in formal if row["repeat"] == "R1"]
    r2 = [row for row in formal if row["repeat"] == "R2"]
    _case_order(r1[:])
    # The first arm of every R1 case is control and the second is candidate.
    for index in range(0, 12, 2):
        pair = r1[index:index + 2]
        if [(item.get("mode"), item.get("shape")) for item in pair] != [(pair[0].get("mode"), pair[0].get("shape"))] * 2:
            fail("R1 lanes must be paired by mode and shape")
        if [lane_arm(item) for item in pair] != list(ARMS):
            fail("R1 lane pairs must be control then candidate")
    # Reversing the complete R1 stream naturally reverses each control/candidate
    # pair as well as the case order, yielding candidate then control.
    expected_r2 = [(row["mode"], row["shape"], lane_arm(row)) for row in reversed(r1)]
    actual_r2 = [(row["mode"], row["shape"], lane_arm(row)) for row in r2]
    if actual_r2 != expected_r2:
        fail("R2 formal lanes must reverse R1 and swap arm order within each case")
    return rows


def formal_rows(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    return [row for row in protocol_order(protocol) if is_formal(row)]


def lane_map(protocol: Mapping[str, Any]) -> dict[str, dict[str, Any]]:
    return {row["lane"]: row for row in protocol_order(protocol)}


def lane_path(root: Path, row: Mapping[str, Any]) -> Path:
    lane = row["lane"]
    value = _field(row, "directory", "output_directory", "capture_directory")
    if isinstance(value, str) and value:
        candidate = root / value if not Path(value).is_absolute() else Path(value)
    else:
        base = "pilots" if row.get("suite") == "pilot" else "captures"
        candidate = root / base / lane
        if not candidate.is_dir():
            candidate = root / lane
    return candidate


def validate_report(report: Mapping[str, Any], row: Mapping[str, Any], path: str) -> dict[str, Any]:
    try:
        return report_checks.validate_report(report, row, path=path, samples=int(row.get("samples", report_checks.SAMPLES)), warmups=int(row.get("warmups", report_checks.WARMUPS)), workers=int(row.get("workers", report_checks.WORKERS)), case=str(row.get("selector", CASE)))
    except (report_checks.ReportError, KeyError, TypeError, ValueError) as error:
        raise AnalysisError(str(error)) from error


def _read_lane(root: Path, row: Mapping[str, Any]) -> dict[str, Any]:
    directory = lane_path(root, row)
    report_path = directory / "report.json"
    if not report_path.is_file():
        fail(f"{row['lane']}: report.json is missing at {report_path}")
    report = read_json(report_path)
    checked = validate_report(report, row, str(report_path))
    resource_path = directory / "resource.log"
    if not resource_path.is_file():
        fail(f"{row['lane']}: resource.log is missing")
    try:
        rss_lines = [line for line in resource_path.read_text(encoding="utf-8").splitlines() if line.lstrip().startswith("Maximum resident set size (kbytes):")]
    except (OSError, UnicodeError) as error:
        fail(f"{row['lane']}: cannot read resource.log: {error}")
    if len(rss_lines) != 1:
        fail(f"{row['lane']}: resource.log must contain one maximum RSS value")
    try:
        rss = int(rss_lines[0].split(":", 1)[1].strip())
    except ValueError as error:
        fail(f"{row['lane']}: resource.log RSS is not an integer")
    if rss < 0:
        fail(f"{row['lane']}: resource.log RSS is negative")
    checked["peak_rss_kib"] = rss
    return checked


def _identity(checked: Mapping[str, Any], lane: Mapping[str, Any]) -> dict[str, Any]:
    identity = dict(checked.get("identity", {}))
    identity["arm"] = lane_arm(lane)
    identity["repeat"] = lane.get("repeat")
    identity["mode"] = lane.get("mode")
    identity["shape"] = lane.get("shape")
    return identity


def _equal_identity(left: Mapping[str, Any], right: Mapping[str, Any]) -> bool:
    # Source metadata can contain process-local paths; parity is based on the
    # deterministic archive, sink counters and writer source projection.
    keys = {key for key in left if key.startswith("output_sha256") or key.startswith("corpus") or key.startswith("sink.") or key == "source"}
    keys &= set(right)
    return bool(keys) and all(left[key] == right[key] for key in keys)


def _stats(values: Sequence[int | float], label: str) -> dict[str, Any]:
    return report_checks.vector_stats(values, label)


def _mean(checked: Mapping[str, Any]) -> float:
    return float(checked["elapsed_summary"]["mean"])


def _elapsed_statistics(checked: Mapping[str, Any]) -> dict[str, Any]:
    summary = checked.get("elapsed_summary")
    if not isinstance(summary, Mapping):
        fail(f"{checked.get('lane', '<lane>')}: elapsed summary is absent")
    return {
        key: copy.deepcopy(summary[key])
        for key in ("min", "p50", "p95", "p99", "max", "mean", "standard_deviation", "confidence_interval_95")
    }


def _regression_observation(control: Mapping[str, Any], candidate: Mapping[str, Any], *, absolute: bool) -> dict[str, Any]:
    metrics: dict[str, Any] = {}
    for key in ("mean", "p50", "p95", "p99"):
        delta = report_checks.percent_delta(float(control[key]), float(candidate[key]))
        metrics[key] = {
            "candidate_vs_control_percent": delta,
            "exceeds_review_threshold": abs(delta) > REVIEW_PERCENT if absolute else delta > REVIEW_PERCENT,
        }
    return {
        "threshold_percent": REVIEW_PERCENT,
        "direction": "absolute_drift" if absolute else "candidate_regression_only",
        "metrics": metrics,
        "any_exceeds_review_threshold": any(item["exceeds_review_threshold"] for item in metrics.values()),
    }


def _allocator_values(checked: Mapping[str, Any], field: str) -> list[int]:
    vectors = checked.get("operation", {}).get("allocation")
    if not isinstance(vectors, dict) or field not in vectors:
        fail(f"{checked.get('lane', '<lane>')}: allocator vector {field} is absent")
    return vectors[field]


def _allocator_peak_delta(checked: Mapping[str, Any]) -> list[int]:
    """Return the operation peak above the per-sample live baseline.

    ``region_peak_live_bytes`` is an absolute process-region counter.  Its
    baseline can differ between fresh child processes, so the paired peak
    observation is the per-sample delta from ``live_bytes_before``.  The raw
    vectors remain in the allocation projection for auditability.
    """

    region = _allocator_values(checked, "region_peak_live_bytes")
    before = _allocator_values(checked, "live_bytes_before")
    if len(region) != len(before):
        fail(f"{checked.get('lane', '<lane>')}: allocator peak vectors are not aligned")
    return [peak - baseline for peak, baseline in zip(region, before)]


def _pair_summary(control: Mapping[str, Any], candidate: Mapping[str, Any], mode: str, shape: str, repeat: str) -> dict[str, Any]:
    result: dict[str, Any] = {
        "repeat": repeat,
        "mode": mode,
        "shape": shape,
        "normal_latency_scope": "descriptive_paired_observation" if mode == "normal" else "not_a_latency_claim",
    }
    result["elapsed_ns"] = report_checks.compare_vectors(control["elapsed"], candidate["elapsed"], f"{repeat}.{mode}.{shape}.elapsed_ns")
    result["elapsed_statistics"] = {
        "control": _elapsed_statistics(control),
        "candidate": _elapsed_statistics(candidate),
    }
    result["peak_rss_kib"] = report_checks.compare_vectors([control["peak_rss_kib"]], [candidate["peak_rss_kib"]], f"{repeat}.{mode}.{shape}.peak_rss_kib")
    if mode == "allocator":
        metrics: dict[str, Any] = {}
        for field in report_checks.ALLOCATOR_FIELDS:
            metrics[field] = report_checks.compare_vectors(_allocator_values(control, field), _allocator_values(candidate, field), f"{repeat}.{mode}.{shape}.allocation.{field}")
        # Keep the absolute counters above, while using the baseline-adjusted
        # operation peak for the paired review flag.
        metrics["live_bytes_delta"] = report_checks.compare_vectors(
            [after - before for after, before in zip(_allocator_values(control, "live_bytes_after"), _allocator_values(control, "live_bytes_before"))],
            [after - before for after, before in zip(_allocator_values(candidate, "live_bytes_after"), _allocator_values(candidate, "live_bytes_before"))],
            f"{repeat}.{mode}.{shape}.allocation.live_bytes_delta",
        )
        metrics["peak_live_bytes_delta"] = report_checks.compare_vectors(
            _allocator_peak_delta(control),
            _allocator_peak_delta(candidate),
            f"{repeat}.{mode}.{shape}.allocation.peak_live_bytes_delta",
        )
        result["allocation"] = metrics
        control_requested = float(metrics["allocated_bytes"]["control"]["mean"])
        candidate_requested = float(metrics["allocated_bytes"]["candidate"]["mean"])
        reduction = 0.0 if control_requested == 0 else (1.0 - candidate_requested / control_requested) * 100.0
        result["requested_bytes_reduction_percent"] = reduction
        control_samples = metrics["allocated_bytes"]["control"]["samples"]
        candidate_samples = metrics["allocated_bytes"]["candidate"]["samples"]
        sample_reductions = [0.0 if control_value == 0 else (1.0 - candidate_value / control_value) * 100.0 for control_value, candidate_value in zip(control_samples, candidate_samples)]
        result["requested_bytes_reduction_min_percent"] = min(sample_reductions)
        result["requested_bytes_reduction_max_percent"] = max(sample_reductions)
    if mode == "normal":
        result["timing_regression"] = _regression_observation(
            result["elapsed_statistics"]["control"], result["elapsed_statistics"]["candidate"], absolute=False
        )
        result["timing_review_required"] = result["timing_regression"]["any_exceeds_review_threshold"]
    else:
        result["timing_regression"] = {
            "status": "not_a_latency_claim",
            "threshold_percent": REVIEW_PERCENT,
            "direction": "descriptive_only",
        }
        result["timing_review_required"] = False
    rss_delta = float(result["peak_rss_kib"]["candidate_vs_control_percent"])
    result["peak_rss_regression"] = {
        "candidate_vs_control_percent": rss_delta,
        "threshold_percent": REVIEW_PERCENT,
        "direction": "candidate_regression_only",
        "exceeds_review_threshold": rss_delta > REVIEW_PERCENT,
    }
    result["peak_rss_review_required"] = result["peak_rss_regression"]["exceeds_review_threshold"]
    if mode == "allocator":
        operation_peak_delta = float(result["allocation"]["peak_live_bytes_delta"]["candidate_vs_control_percent"])
        result["operation_peak_regression"] = {
            "candidate_vs_control_percent": operation_peak_delta,
            "threshold_percent": REVIEW_PERCENT,
            "direction": "candidate_regression_only",
            "exceeds_review_threshold": operation_peak_delta > REVIEW_PERCENT,
        }
        result["operation_peak_review_required"] = result["operation_peak_regression"]["exceeds_review_threshold"]
    else:
        result["operation_peak_regression"] = {"status": "not_applicable", "threshold_percent": REVIEW_PERCENT}
        result["operation_peak_review_required"] = False
    result["review_required"] = result["timing_review_required"] or result["peak_rss_review_required"] or result["operation_peak_review_required"]
    return result


def _repeat_summary(rows: Mapping[tuple[str, str, str, str], Mapping[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for arm in ARMS:
        for mode in MODES:
            for shape in SHAPES:
                first = rows[("R1", arm, mode, shape)]
                second = rows[("R2", arm, mode, shape)]
                delta = report_checks.percent_delta(_mean(first), _mean(second))
                rss_delta = report_checks.percent_delta(float(first["peak_rss_kib"]), float(second["peak_rss_kib"]))
                elapsed_drift = (
                    _regression_observation(_elapsed_statistics(first), _elapsed_statistics(second), absolute=True)
                    if mode == "normal"
                    else {"status": "not_a_latency_claim", "threshold_percent": REVIEW_PERCENT, "direction": "descriptive_only"}
                )
                result[f"{arm}-{mode}-{shape}"] = {
                    "r1_mean_ns": _mean(first),
                    "r2_mean_ns": _mean(second),
                    "signed_percent": delta,
                    "exceeds_review_threshold": elapsed_drift.get("any_exceeds_review_threshold", False),
                    "elapsed_drift": elapsed_drift,
                    "r1_peak_rss_kib": first["peak_rss_kib"],
                    "r2_peak_rss_kib": second["peak_rss_kib"],
                    "peak_rss_signed_percent": rss_delta,
                    "peak_rss_exceeds_review_threshold": rss_delta > REVIEW_PERCENT,
                }
                if mode == "allocator":
                    r1_peak = report_checks.vector_stats(_allocator_peak_delta(first), f"{arm}-{mode}-{shape}.R1.peak_live_bytes_delta")
                    r2_peak = report_checks.vector_stats(_allocator_peak_delta(second), f"{arm}-{mode}-{shape}.R2.peak_live_bytes_delta")
                    peak_delta = report_checks.percent_delta(float(r1_peak["mean"]), float(r2_peak["mean"]))
                    result[f"{arm}-{mode}-{shape}"].update({
                        "r1_operation_peak_delta_bytes": r1_peak,
                        "r2_operation_peak_delta_bytes": r2_peak,
                        "operation_peak_signed_percent_r1_to_r2": peak_delta,
                        "operation_peak_exceeds_review_threshold": abs(peak_delta) > REVIEW_PERCENT,
                    })
    return result


def _auxiliary(root: Path, rows: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for row in rows:
        directory = lane_path(root, row)
        report = directory / "report.json"
        item: dict[str, Any] = {
            "lane": row["lane"],
            "kind": lane_kind(row),
            "selector": _field(row, "selector", "case"),
            "arm": lane_arm(row),
            "repeat": row.get("repeat"),
            "mode": row.get("mode"),
            "shape": row.get("shape"),
            "report": binding(report, root) if report.is_file() else {"path": str(report), "missing": True},
        }
        result.append(item)
    return result


def _semantic_projection(result: Mapping[str, Any]) -> dict[str, Any]:
    """Keep deterministic corpus/source/sink/metric scalars, drop samples."""

    def strip_vectors(value: Any) -> Any:
        if isinstance(value, dict):
            return {
                key: strip_vectors(child)
                for key, child in value.items()
                if key not in {"values", "sample_indices"}
            }
        if isinstance(value, list):
            return [strip_vectors(child) for child in value]
        return value

    projection = copy.deepcopy(dict(result))
    projection.pop("elapsed_ns", None)
    projection.pop("process", None)
    projection.pop("allocation", None)
    if "operation_metrics" in projection:
        projection["operation_metrics"] = strip_vectors(projection["operation_metrics"])
    return projection


def _auxiliary_rss(root: Path, row: Mapping[str, Any]) -> int:
    path = lane_path(root, row) / "resource.log"
    if not path.is_file():
        fail(f"{row['lane']}: resource.log is missing")
    try:
        lines = [line for line in path.read_text(encoding="utf-8").splitlines() if line.lstrip().startswith("Maximum resident set size (kbytes):")]
    except (OSError, UnicodeError) as error:
        fail(f"{row['lane']}: cannot read resource.log: {error}")
    if len(lines) != 1:
        fail(f"{row['lane']}: resource.log must contain one maximum RSS value")
    try:
        value = int(lines[0].split(":", 1)[1].strip())
    except ValueError as error:
        fail(f"{row['lane']}: resource.log RSS is not an integer")
    if value < 0:
        fail(f"{row['lane']}: resource.log RSS is negative")
    return value


def _auxiliary_pair(
    measurements: Mapping[str, Mapping[str, Mapping[str, Any]]],
    rss: Mapping[str, int],
    control_lane: str,
    candidate_lane: str,
    *,
    label: str,
    absolute_drift: bool = False,
) -> dict[str, Any]:
    control = measurements[control_lane]
    candidate = measurements[candidate_lane]
    if set(control) != set(candidate):
        fail(f"{label}: result identity sets differ")
    rows: list[dict[str, Any]] = []
    for key in sorted(control):
        left = control[key]
        right = candidate[key]
        comparison = report_checks.compare_vectors(left["samples"], right["samples"], f"{label}.{key}.elapsed_ns")
        regression = _regression_observation(left["stats"], right["stats"], absolute=absolute_drift)
        rows.append({"key": key, "elapsed_ns": comparison, "regression": regression})
    rss_delta = report_checks.percent_delta(float(rss[control_lane]), float(rss[candidate_lane]))
    rss_comparison = {
        "control_kib": rss[control_lane],
        "candidate_kib": rss[candidate_lane],
        "candidate_vs_control_percent": rss_delta,
        "threshold_percent": REVIEW_PERCENT,
        "direction": "absolute_drift" if absolute_drift else "candidate_regression_only",
        "exceeds_review_threshold": abs(rss_delta) > REVIEW_PERCENT if absolute_drift else rss_delta > REVIEW_PERCENT,
    }
    return {
        "control_lane": control_lane,
        "candidate_lane": candidate_lane,
        "keys": sorted(control),
        "rows": rows,
        "peak_rss": rss_comparison,
    }


def _auxiliary_comparisons(root: Path, all_rows: Sequence[Mapping[str, Any]], protocol: Mapping[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for suite in ("counter", "guard"):
        rows = [row for row in all_rows if row.get("suite") == suite]
        if not rows:
            continue
        expected_arms = ["control", "candidate", "candidate", "control"]
        if len(rows) != 4:
            fail(f"{suite} suite must contain four ABBA lanes")
        actual_arms = [lane_arm(row) for row in rows]
        if actual_arms != expected_arms:
            fail(f"{suite} suite arm order must be control/candidate/candidate/control")
        lane_data: list[dict[str, Any]] = []
        projections: dict[str, dict[str, Any]] = {}
        measurements: dict[str, dict[str, dict[str, Any]]] = {}
        rss: dict[str, int] = {}
        for row in rows:
            path = lane_path(root, row) / "report.json"
            if not path.is_file():
                fail(f"{row['lane']}: auxiliary report is missing")
            report = read_json(path)
            raw_results = report.get("results") if isinstance(report, dict) else None
            if not isinstance(raw_results, list):
                fail(f"{row['lane']}: auxiliary results are missing")
            expected_count = 1 if suite == "counter" else 10
            if len(raw_results) != expected_count:
                fail(f"{row['lane']}: expected {expected_count} auxiliary rows")
            if suite == "guard":
                try:
                    report_checks.validate_guard_report(report, protocol.get("guard_selectors", []), path=str(path))
                except report_checks.ReportError as error:
                    fail(f"{row['lane']}: {error}")
            entries: list[dict[str, Any]] = []
            seen_keys: set[str] = set()
            for index, raw in enumerate(raw_results):
                if not isinstance(raw, dict):
                    fail(f"{row['lane']}.results[{index}] is malformed")
                samples, _, stats = report_checks.check_elapsed(raw.get("elapsed_ns"), f"{row['lane']}.results[{index}].elapsed_ns", int(row.get("samples", report_checks.SAMPLES)))
                key = f"{raw.get('case')}::{raw.get('corpus', {}).get('shape') if isinstance(raw.get('corpus'), dict) else None}"
                if key in seen_keys:
                    fail(f"{row['lane']}: duplicate auxiliary result identity {key}")
                seen_keys.add(key)
                projection = _semantic_projection(raw)
                entries.append({"key": key, "elapsed_ns": stats, "samples": samples, "projection": projection})
            entries.sort(key=lambda item: item["key"])
            if suite == "guard":
                selectors = protocol.get("guard_selectors")
                if not isinstance(selectors, list) or not selectors:
                    fail("protocol.guard_selectors is missing")
                expected_keys = {f"{selector}::{shape}" for selector in selectors for shape in str(row.get("shape", "")).split(",") if shape}
                if seen_keys != expected_keys:
                    fail(f"{row['lane']}: guard result identities differ from declared selectors/shapes")
            projections[row["lane"]] = {entry["key"]: entry["projection"] for entry in entries}
            measurements[row["lane"]] = {entry["key"]: {"samples": entry["samples"], "stats": entry["elapsed_ns"]} for entry in entries}
            rss[row["lane"]] = _auxiliary_rss(root, row)
            lane_summary: dict[str, Any] = {"lane": row["lane"], "arm": lane_arm(row), "peak_rss_kib": rss[row["lane"]], "rows": [{"key": entry["key"], "elapsed_ns": entry["elapsed_ns"]} for entry in entries]}
            if suite == "counter":
                counter_path = lane_path(root, row) / "counters.csv"
                if not counter_path.is_file():
                    fail(f"{row['lane']}: counters.csv is missing")
                events = protocol.get("counter_events")
                if not isinstance(events, str) or not events:
                    fail("protocol.counter_events is missing")
                try:
                    parsed = report_checks.parse_counter_text(counter_path.read_text(encoding="utf-8"), [event for event in events.split(",") if event])
                except (OSError, UnicodeError, report_checks.ReportError) as error:
                    fail(f"{row['lane']}: invalid counters.csv: {error}")
                lane_summary["counters"] = parsed
            lane_data.append(lane_summary)
        baseline = projections[rows[0]["lane"]]
        for row in rows[1:]:
            if projections[row["lane"]] != baseline:
                fail(f"{suite}: semantic/source/sink scalar projection differs between ABBA lanes")
        suite_result: dict[str, Any] = {"lanes": lane_data, "semantic_source_sink_parity": True, "scope": "all retained auxiliary report rows"}
        if suite == "guard":
            suite_result["pairs"] = {
                "G1_to_G2": _auxiliary_pair(measurements, rss, "G1-control", "G2-candidate", label="guard.G1_to_G2"),
                "G4_to_G3": _auxiliary_pair(measurements, rss, "G4-control", "G3-candidate", label="guard.G4_to_G3"),
            }
            suite_result["repeat_drift"] = {
                "G1_to_G4": _auxiliary_pair(measurements, rss, "G1-control", "G4-control", label="guard.G1_to_G4", absolute_drift=True),
                "G2_to_G3": _auxiliary_pair(measurements, rss, "G2-candidate", "G3-candidate", label="guard.G2_to_G3", absolute_drift=True),
            }
        result[suite] = suite_result
    return result


def analyze(root: Path = ROOT) -> dict[str, Any]:
    protocol_path = root / "protocol.json"
    protocol = read_json(protocol_path)
    all_rows = protocol_order(protocol)
    formals = [row for row in all_rows if is_formal(row)]
    checked: dict[tuple[str, str, str, str], dict[str, Any]] = {}
    identities: dict[str, dict[str, Any]] = {}
    for row in formals:
        key = (str(row["repeat"]), str(lane_arm(row)), str(row["mode"]), str(row["shape"]))
        if key in checked:
            fail(f"duplicate formal lane identity {key!r}")
        value = _read_lane(root, row)
        value["lane"] = row["lane"]
        checked[key] = value
        identities[row["lane"]] = _identity(value, row)
    expected_keys = {(repeat, arm, mode, shape) for repeat in REPEATS for arm in ARMS for mode in MODES for shape in SHAPES}
    if set(checked) != expected_keys:
        fail("formal lane identities are incomplete")

    parity: list[dict[str, Any]] = []
    pairs: dict[str, dict[str, Any]] = {}
    for repeat in REPEATS:
        for mode in MODES:
            for shape in SHAPES:
                control = checked[(repeat, "control", mode, shape)]
                candidate = checked[(repeat, "candidate", mode, shape)]
                left = identities[control["lane"]]
                right = identities[candidate["lane"]]
                equal = _equal_identity(left, right)
                if not equal:
                    fail(f"{repeat}.{mode}.{shape}: control/candidate output or writer identity differs")
                parity.append({"repeat": repeat, "mode": mode, "shape": shape, "equal": equal, "identity": {key: left[key] for key in left if key not in {"arm", "repeat", "mode", "shape"}}})
                pairs[f"{repeat}-{mode}-{shape}"] = _pair_summary(control, candidate, mode, shape, repeat)

    # Different shapes intentionally have different archive identities.  The
    # parity invariant is therefore per shape: every arm, mode, and repeat
    # must retain the same output/source projection for that shape.
    for shape in SHAPES:
        shape_identities = [item["identity"] for item in parity if item["shape"] == shape]
        if shape_identities and any(item != shape_identities[0] for item in shape_identities[1:]):
            fail(f"formal lanes for {shape} do not retain one deterministic output/source identity")

    repeat = _repeat_summary(checked)
    large_reductions = [pairs[f"{repeat_name}-allocator-large"]["requested_bytes_reduction_percent"] for repeat_name in REPEATS]
    large_min_reductions = [pairs[f"{repeat_name}-allocator-large"]["requested_bytes_reduction_min_percent"] for repeat_name in REPEATS]
    reduction_mean = math.fsum(large_reductions) / len(large_reductions)
    acceptance = protocol.get("acceptance", {})
    if acceptance is not None and not isinstance(acceptance, (dict, str)):
        fail("protocol.acceptance must be an object or descriptive string")
    if isinstance(acceptance, dict):
        threshold = acceptance.get("requested_bytes_reduction_percent", acceptance.get("minimum_large_requested_bytes_reduction_percent", protocol.get("large_requested_bytes_reduction_required_percent", 95.0)))
    else:
        threshold = protocol.get("large_requested_bytes_reduction_required_percent", 95.0)
    if isinstance(threshold, bool) or not isinstance(threshold, (int, float)) or not math.isfinite(float(threshold)):
        fail("protocol acceptance reduction threshold is invalid")
    failed_allocations = []
    for key, item in checked.items():
        if key[2] == "allocator":
            failed_allocations.extend(_allocator_values(item, "failed_allocation_calls"))
    summary = {
        "schema": SCHEMA,
        "protocol": binding(protocol_path, root),
        "formal_lanes": [
            {
                "lane": row["lane"],
                "repeat": row["repeat"],
                "arm": lane_arm(row),
                "mode": row["mode"],
                "shape": row["shape"],
                "report": binding(lane_path(root, row) / "report.json", root),
                "identity": identities[row["lane"]],
            }
            for row in formals
        ],
        "auxiliary_lanes": _auxiliary(root, [row for row in all_rows if not is_formal(row)]),
        "auxiliary_comparisons": _auxiliary_comparisons(root, [row for row in all_rows if not is_formal(row)], protocol),
        "pairs": pairs,
        "normal_repeat_drift": repeat,
        "output_parity": parity,
        "acceptance": {
            "large_requested_bytes_reduction_threshold_percent": float(threshold),
            "R1_percent": large_reductions[0],
            "R2_percent": large_reductions[1],
            "R1_min_sample_percent": large_min_reductions[0],
            "R2_min_sample_percent": large_min_reductions[1],
            "mean_percent": reduction_mean,
            "passed": all(value >= float(threshold) for value in large_reductions),
            "all_samples_passed": all(value >= float(threshold) for value in large_min_reductions),
            "scope": "allocator requested_bytes arithmetic across matched large PPTX lanes",
        },
        "failed_allocation_calls": {
            "maximum": max(failed_allocations) if failed_allocations else 0,
            "all_zero": not any(failed_allocations),
            "scope": "formal allocator samples",
        },
        "claims": [
            "Control/candidate comparisons are descriptive paired observations on deterministic output-identical formal lanes.",
            "Normal elapsed comparisons carry no registered-latency claim; values over five percent require review.",
            "Allocator requested-byte reduction is scoped to the operation allocator counter and does not establish constant memory or RSS behavior.",
            "Allocator operation-peak review uses region_peak_live_bytes minus live_bytes_before per sample; the absolute vectors are retained for audit.",
            "Peak RSS is a whole-process /usr/bin/time observation; paired differences over five percent are review flags.",
            "Profiler and auxiliary selector lanes are retained for transport guards and are not folded into the PPTX latency comparison.",
        ],
    }
    return summary


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args(argv)
    result = analyze(args.root)
    output = args.root / "summary.json"
    if args.check:
        if not output.is_file() or read_json(output) != result:
            raise SystemExit("summary.json does not replay exactly")
        print("0476 paired summary replay passed")
    else:
        output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print("0476 paired summary written")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
