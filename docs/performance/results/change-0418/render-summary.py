#!/usr/bin/env python3
"""Render the verified 0418 ABBA evidence without making a pooled claim.

The verifier remains the authority for raw reports, catalogs, compressed
report fallbacks, corpus identities, binary identities, and the published
projections.  This renderer replays that verifier before consuming the
projections, then writes a readable table and a deterministic uncertainty
description.  It deliberately keeps the four ABBA legs separate.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import sys
from pathlib import Path
from typing import Any


SELECTORS = (
    "pptx_cross_copy_media_rich_lifecycle",
    "pptx_cross_copy_plain_lifecycle",
    "pptx_cross_copy_plain",
    "pptx_cross_copy_media_rich",
)
LEGS = ("A1", "B1", "B2", "A2")
LEG_KEYS = {leg: leg.lower() for leg in LEGS}
STATISTICS = ("p50", "mean", "p95", "p99")
PAIRS = (
    ("a1_to_b1", "A1", "B1"),
    ("a2_to_b2", "A2", "B2"),
)
NONCLAIMABLE_SNAPSHOTS = (
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
)
REVIEW_THRESHOLD_PERCENT = 5
BOOTSTRAP_RESAMPLES = 1_000


class RenderError(ValueError):
    """The verified projections are missing or inconsistent."""


def _fail(path: str, message: str) -> None:
    raise RenderError(f"{path}: {message}")


def _strict_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    value: dict[str, Any] = {}
    for key, item in pairs:
        if key in value:
            raise ValueError(f"duplicate JSON key {key!r}")
        value[key] = item
    return value


def _reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON value {value!r}")


def _load_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_strict_pairs,
            parse_constant=_reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        _fail(str(path), f"cannot read strict JSON: {error}")
    raise AssertionError("unreachable")


def _object(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        _fail(path, "must be an object")
    return value


def _number(value: Any, path: str) -> float | int:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        _fail(path, "must be a number")
    if not math.isfinite(float(value)):
        _fail(path, "must be finite")
    return value


def _integer(value: Any, path: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        _fail(path, f"must be an integer >= {minimum}")
    return value


def _load_verified_projections(
    root: Path,
) -> tuple[dict[str, Any], dict[str, Any], Any]:
    """Run the frozen verifier in replay mode before reading its outputs."""
    verifier_path = root / "verify.py"
    if not verifier_path.is_file():
        _fail(str(verifier_path), "frozen verifier is missing")
    spec = importlib.util.spec_from_file_location(
        "change0418_render_verifier", verifier_path
    )
    if spec is None or spec.loader is None:
        _fail(str(verifier_path), "cannot load frozen verifier")
    verifier = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = verifier
    try:
        spec.loader.exec_module(verifier)
        verifier.verify(root, write=False)
    except Exception as error:
        raise RenderError(f"0418 verifier replay failed: {error}") from error
    summary = _object(_load_json(root / "summary.json"), "summary.json")
    allocation = _object(
        _load_json(root / "allocation-metrics.json"),
        "allocation-metrics.json",
    )
    return summary, allocation, verifier


def _load_formal_samples(
    root: Path, summary: dict[str, Any], verifier: Any
) -> dict[str, dict[str, list[int]]]:
    """Read the retained normal reports after verifier replay.

    The published ABBA projection intentionally keeps scalar statistics rather
    than raw timing vectors.  Re-reading the formal reports here keeps the
    renderer's uncertainty output tied to those retained vectors, including a
    ``.zst``/``.zstd`` sidecar when the uncompressed path is absent.
    """
    evidence = _object(summary.get("correctness_evidence"), "summary.json.correctness_evidence")
    formal_runs = evidence.get("formal_runs")
    if not isinstance(formal_runs, list):
        _fail("summary.json.correctness_evidence.formal_runs", "must be a list")
    samples: dict[str, dict[str, list[int]]] = {selector: {} for selector in SELECTORS}
    seen: set[tuple[str, str]] = set()
    for index, evidence_row in enumerate(formal_runs):
        row = _object(
            evidence_row,
            f"summary.json.correctness_evidence.formal_runs[{index}]",
        )
        if row.get("phase") != "normal":
            continue
        selector = row.get("selector")
        leg = row.get("leg")
        if selector not in SELECTORS or leg not in LEGS:
            _fail(
                f"summary.json.correctness_evidence.formal_runs[{index}]",
                "has an unknown normal selector or ABBA leg",
            )
        key = (selector, leg)
        if key in seen:
            _fail(
                f"summary.json.correctness_evidence.formal_runs[{index}]",
                "duplicates a normal selector/leg report",
            )
        seen.add(key)
        report_path = row.get("report")
        if not isinstance(report_path, str) or not report_path:
            _fail(
                f"summary.json.correctness_evidence.formal_runs[{index}].report",
                "must be a non-empty relative report path",
            )
        label = f"formal normal {selector} {leg}"
        try:
            artifact = verifier.Artifact(root, report_path, label)
            report = _object(artifact.json(label), label)
        except Exception as error:
            raise RenderError(f"{label}: cannot read retained report: {error}") from error
        expected_digest = row.get("report_sha256")
        if isinstance(expected_digest, str) and artifact.sha256 != expected_digest:
            _fail(f"{label}.report_sha256", "does not match retained report bytes")
        result_list = report.get("results")
        if not isinstance(result_list, list):
            _fail(f"{label}.results", "must be a list")
        matching = [
            _object(item, f"{label}.results[{result_index}]")
            for result_index, item in enumerate(result_list)
            if isinstance(item, dict) and item.get("case") == selector
        ]
        if len(matching) != 1:
            _fail(f"{label}.results", "must contain one result for the selector")
        elapsed = _object(matching[0].get("elapsed_ns"), f"{label}.elapsed_ns")
        values = elapsed.get("samples")
        if not isinstance(values, list) or not values:
            _fail(f"{label}.elapsed_ns.samples", "must be a non-empty vector")
        integer_values = [
            _integer(value, f"{label}.elapsed_ns.samples[{value_index}]")
            for value_index, value in enumerate(values)
        ]
        if integer_values != sorted(integer_values):
            _fail(f"{label}.elapsed_ns.samples", "must be sorted ascending")
        samples[selector][leg] = integer_values
    if seen != {(selector, leg) for selector in SELECTORS for leg in LEGS}:
        _fail(
            "summary.json.correctness_evidence.formal_runs",
            "must contain exactly one normal report for each selector and ABBA leg",
        )
    return samples


def _validate_projections(
    summary: dict[str, Any],
    allocation: dict[str, Any],
    raw_samples: dict[str, dict[str, list[int]]],
) -> dict[str, dict[str, Any]]:
    if summary.get("schema_version") != 1 or summary.get("change") != "0418":
        _fail("summary.json", "schema_version/change does not identify 0418")
    if summary.get("timing_status") != (
        "normal ABBA elapsed statistics validated from retained samples"
    ):
        _fail("summary.json.timing_status", "unexpected timing status")
    if summary.get("allocator_timing_status") != "excluded":
        _fail("summary.json.allocator_timing_status", "allocator latency is not excluded")
    protocol = _object(summary.get("protocol"), "summary.json.protocol")
    if protocol.get("order") != list(LEGS):
        _fail("summary.json.protocol.order", "must retain A1/B1/B2/A2 order")
    if protocol.get("roles") != ["control", "candidate", "candidate", "control"]:
        _fail("summary.json.protocol.roles", "must match the ABBA roles")
    if protocol.get("latency_drift_ceilings_percent") != {
        name: REVIEW_THRESHOLD_PERCENT for name in STATISTICS
    }:
        _fail(
            "summary.json.protocol.latency_drift_ceilings_percent",
            "must retain the five percent ceilings",
        )
    capture = _object(summary.get("capture"), "summary.json.capture")
    if capture.get("selectors") != list(SELECTORS):
        _fail("summary.json.capture.selectors", "selector order does not match protocol")
    normal_counts = _object(
        protocol.get("normal_counts"), "summary.json.protocol.normal_counts"
    )
    normal_abba = _object(summary.get("normal_abba"), "summary.json.normal_abba")
    if set(normal_abba) != set(SELECTORS):
        _fail("summary.json.normal_abba", "must contain exactly the four selectors")

    if allocation.get("schema_version") != 1 or allocation.get("change") != "0418":
        _fail("allocation-metrics.json", "schema_version/change does not identify 0418")
    if allocation.get("timing_status") != "observational_only":
        _fail("allocation-metrics.json.timing_status", "unexpected allocation status")
    if allocation.get("latency_comparison") != "excluded":
        _fail("allocation-metrics.json.latency_comparison", "must exclude allocator latency")
    if allocation.get("review_threshold_percent") != REVIEW_THRESHOLD_PERCENT:
        _fail(
            "allocation-metrics.json.review_threshold_percent",
            "must be five percent",
        )
    if set(allocation.get("nonclaimable_metrics", [])) != set(NONCLAIMABLE_SNAPSHOTS):
        _fail(
            "allocation-metrics.json.nonclaimable_metrics",
            "must retain the literal live/peak snapshot boundary",
        )
    if allocation.get("validation", {}).get("no_allocator_latency_claim") is not True:
        _fail(
            "allocation-metrics.json.validation.no_allocator_latency_claim",
            "must be true",
        )
    matched = _object(
        allocation.get("matched_pairs"), "allocation-metrics.json.matched_pairs"
    )
    if set(matched) != set(SELECTORS):
        _fail("allocation-metrics.json.matched_pairs", "must contain the four selectors")

    rows: dict[str, dict[str, Any]] = {}
    for selector in SELECTORS:
        expected_count = _integer(
            _object(normal_counts.get(selector), f"normal_counts.{selector}").get("samples"),
            f"normal_counts.{selector}.samples",
            minimum=1,
        )
        summary_selector = _object(
            normal_abba[selector], f"summary.json.normal_abba.{selector}"
        )
        result_list = summary_selector.get("results")
        if not isinstance(result_list, list) or len(result_list) != 1:
            _fail(
                f"summary.json.normal_abba.{selector}.results",
                "must contain one selected corpus result",
            )
        result = _object(result_list[0], f"summary.json.normal_abba.{selector}.results[0]")
        if result.get("case") != selector:
            _fail(f"summary.json.normal_abba.{selector}.results[0].case", "selector mismatch")
        elapsed = _object(
            result.get("elapsed_ns"),
            f"summary.json.normal_abba.{selector}.results[0].elapsed_ns",
        )
        if elapsed.get("sample_count") != expected_count:
            _fail(
                f"summary.json.normal_abba.{selector}.elapsed_ns.sample_count",
                "does not match the protocol",
            )
        legs = _object(
            elapsed.get("legs_ns"),
            f"summary.json.normal_abba.{selector}.elapsed_ns.legs_ns",
        )
        for leg in LEGS:
            key = LEG_KEYS[leg]
            leg_stats = _object(legs.get(key), f"{selector}.{key}")
            integer_samples = raw_samples[selector][leg]
            if len(integer_samples) != expected_count:
                _fail(f"{selector}.{key}.samples", "sample vector length mismatch")
            if leg_stats.get("sample_count") != expected_count:
                _fail(f"{selector}.{key}.sample_count", "does not match the protocol")
            for statistic in STATISTICS:
                _number(leg_stats.get(statistic), f"{selector}.{key}.{statistic}")
            recomputed = {
                statistic: _sample_statistic(integer_samples, statistic)
                for statistic in STATISTICS
            }
            for statistic in ("p50", "p95", "p99"):
                if leg_stats[statistic] != recomputed[statistic]:
                    _fail(
                        f"{selector}.{key}.{statistic}",
                        "does not match the retained raw vector",
                    )
            if not math.isclose(
                float(leg_stats["mean"]), float(recomputed["mean"]),
                rel_tol=1e-12, abs_tol=1e-6,
            ):
                _fail(
                    f"{selector}.{key}.mean",
                    "does not match the retained raw vector",
                )
        for field in (
            "candidate_reduction_percent",
            "same_implementation_drift_percent",
        ):
            value = _object(elapsed.get(field), f"{selector}.{field}")
            for name in ("a1_to_b1", "a2_to_b2") if field.startswith("candidate") else ("control", "candidate"):
                values = _object(value.get(name), f"{selector}.{field}.{name}")
                for statistic in STATISTICS:
                    _number(values.get(statistic), f"{selector}.{field}.{name}.{statistic}")
        accepted = elapsed.get("accepted_statistics")
        rejected = elapsed.get("rejected_statistics")
        if not isinstance(accepted, list) or not all(
            isinstance(item, str) and item in STATISTICS for item in accepted
        ):
            _fail(f"{selector}.accepted_statistics", "must contain known statistic names")
        if (
            not isinstance(rejected, dict)
            or set(accepted) & set(rejected)
            or set(accepted) | set(rejected) != set(STATISTICS)
        ):
            _fail(f"{selector}.rejected_statistics", "must be a disjoint object")
        for statistic, reason in rejected.items():
            if statistic not in STATISTICS or not isinstance(reason, str) or not reason:
                _fail(f"{selector}.rejected_statistics", "has an invalid entry")
        rows[selector] = {
            "result": result,
            "elapsed": elapsed,
            "legs": legs,
            "raw_samples": raw_samples[selector],
        }

    return rows


def _midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def _nearest_rank(values: list[int], percentile: int) -> int:
    index = min((percentile * len(values) + 99) // 100 - 1, len(values) - 1)
    return values[index]


def _sample_statistic(values: list[int], statistic: str) -> int | float:
    ordered = sorted(values)
    if statistic == "p50":
        return _midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2])
    if statistic == "mean":
        return math.fsum(values) / len(values)
    if statistic == "p95":
        return _nearest_rank(ordered, 95)
    if statistic == "p99":
        return _nearest_rank(ordered, 99)
    raise AssertionError(f"unknown statistic {statistic}")


def _median_interval(values: list[int]) -> dict[str, Any] | None:
    """Return the narrowest exact distribution-free IID 95% median interval."""
    count = len(values)
    total = 1 << count
    tail = 0
    selected: int | None = None
    for rank in range(1, count // 2 + 1):
        tail += math.comb(count, rank - 1)
        if (total - 2 * tail) * 20 >= total * 19:
            selected = rank
    if selected is None:
        return None
    ordered = sorted(values)
    return {
        "method": "exact_distribution_free_iid_95_percent_median_order_statistic",
        "assumption": "IID samples within one isolated process",
        "lower_rank": selected,
        "upper_rank": count - selected + 1,
        "lower_ns": ordered[selected - 1],
        "upper_ns": ordered[count - selected],
    }


class _DeterministicRng:
    """Small fixed PRNG so bootstrap output does not depend on Python version."""

    def __init__(self, seed: int) -> None:
        self.state = seed & ((1 << 64) - 1) or 1

    def next_u64(self) -> int:
        self.state = (
            self.state * 6364136223846793005 + 1442695040888963407
        ) & ((1 << 64) - 1)
        return self.state

    def randbelow(self, bound: int) -> int:
        limit = ((1 << 64) // bound) * bound
        while True:
            value = self.next_u64()
            if value < limit:
                return value % bound


def _seed(selector: str, leg: str) -> int:
    digest = hashlib.sha256(
        f"change-0418\0{selector}\0{leg}".encode("utf-8")
    ).digest()
    return int.from_bytes(digest[:8], "big") or 1


def _percentile(sorted_values: list[int | float], probability: float) -> int | float:
    rank = max(1, math.ceil(probability * len(sorted_values)))
    return sorted_values[min(rank - 1, len(sorted_values) - 1)]


def _bootstrap_intervals(
    values: list[int], selector: str, leg: str
) -> dict[str, Any]:
    rng = _DeterministicRng(_seed(selector, leg))
    estimates: dict[str, list[int | float]] = {
        statistic: [] for statistic in STATISTICS
    }
    for _ in range(BOOTSTRAP_RESAMPLES):
        resample = [values[rng.randbelow(len(values))] for _ in values]
        resample.sort()
        estimates["p50"].append(
            _midpoint(resample[(len(resample) - 1) // 2], resample[len(resample) // 2])
        )
        estimates["mean"].append(math.fsum(resample) / len(resample))
        estimates["p95"].append(_nearest_rank(resample, 95))
        estimates["p99"].append(_nearest_rank(resample, 99))
    intervals: dict[str, Any] = {}
    for statistic in STATISTICS:
        ordered = sorted(estimates[statistic])
        intervals[statistic] = {
            "method": "deterministic_percentile_bootstrap_95",
            "assumption": "IID samples within one isolated process",
            "resamples": BOOTSTRAP_RESAMPLES,
            "seed_hex": f"{_seed(selector, leg):016x}",
            "lower_ns": _percentile(ordered, 0.025),
            "upper_ns": _percentile(ordered, 0.975),
        }
    return intervals


def _normal_uncertainty(rows: dict[str, dict[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for selector in SELECTORS:
        elapsed = rows[selector]["elapsed"]
        legs = rows[selector]["legs"]
        leg_result: dict[str, Any] = {}
        for leg in LEGS:
            key = LEG_KEYS[leg]
            stats = _object(legs[key], f"{selector}.{key}")
            values = rows[selector]["raw_samples"][leg]
            bootstrap = _bootstrap_intervals(values, selector, leg)
            statistics: dict[str, Any] = {}
            for statistic in STATISTICS:
                observed = stats[statistic]
                if statistic == "mean":
                    producer_ci = _object(
                        stats.get("confidence_interval_95"),
                        f"{selector}.{key}.confidence_interval_95",
                    )
                    statistics[statistic] = {
                        "observed_ns": observed,
                        "producer_student_t_interval_95": {
                            "method": producer_ci.get("method"),
                            "lower_ns": _number(
                                producer_ci.get("lower"),
                                f"{selector}.{key}.confidence_interval_95.lower",
                            ),
                            "upper_ns": _number(
                                producer_ci.get("upper"),
                                f"{selector}.{key}.confidence_interval_95.upper",
                            ),
                        },
                        "bootstrap_percentile_interval_95": bootstrap[statistic],
                    }
                elif statistic == "p50":
                    statistics[statistic] = {
                        "observed_ns": observed,
                        "exact_iid_order_statistic_interval_95": _median_interval(values),
                        "bootstrap_percentile_interval_95": bootstrap[statistic],
                    }
                else:
                    statistics[statistic] = {
                        "observed_ns": observed,
                        "bootstrap_percentile_interval_95": bootstrap[statistic],
                    }
            leg_result[leg] = {
                "sample_count": len(values),
                "statistics": statistics,
            }
        result[selector] = {
            "legs": leg_result,
            "paired_changes_percent": elapsed["candidate_reduction_percent"],
            "same_revision_drift_percent": elapsed["same_implementation_drift_percent"],
            "drift_ceiling_percent": elapsed["drift_ceiling_percent"],
            "accepted_statistics": elapsed["accepted_statistics"],
            "rejected_statistics": elapsed["rejected_statistics"],
            "adverse_both_statistics": elapsed["adverse_both_statistics"],
        }
    return result


def _fmt_ns(value: int | float) -> str:
    return f"{float(value) / 1_000_000.0:.6f}"


def _fmt_percent(value: Any) -> str:
    if value is None:
        return "unavailable"
    return f"{float(value):+.3f}%"


def _fmt_count(value: Any) -> str:
    if value is None:
        return "unavailable"
    return f"{int(value):,}"


def _fmt_delta(metric: dict[str, Any]) -> str:
    if metric.get("status") != "measured":
        return "unavailable"
    return (
        f"{_fmt_count(metric.get('control_sum'))} → "
        f"{_fmt_count(metric.get('candidate_sum'))} "
        f"({_fmt_percent(metric.get('candidate_minus_control_percent'))})"
        + (" REVIEW" if metric.get("review_required") else "")
    )


def _fmt_rss(pair: dict[str, Any]) -> str:
    return (
        f"{_fmt_count(pair.get('control_kib'))} → "
        f"{_fmt_count(pair.get('candidate_kib'))} KiB "
        f"({_fmt_percent(pair.get('candidate_minus_control_percent'))})"
        + (" REVIEW" if pair.get("review_required") else "")
    )


def _markdown_table(
    rows: list[list[str]], headers: list[str]
) -> list[str]:
    def cell(value: str) -> str:
        return value.replace("|", "\\|").replace("\n", " ")

    lines = [
        "| " + " | ".join(cell(value) for value in headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    lines.extend(
        "| " + " | ".join(cell(value) for value in row) + " |" for row in rows
    )
    return lines


def _render_table(
    summary: dict[str, Any],
    allocation: dict[str, Any],
    rows: dict[str, dict[str, Any]],
) -> str:
    lines = [
        "# Change 0418 performance evidence",
        "",
        "The frozen verifier was replayed before this table was rendered. It validated",
        "the retained raw reports, catalogs, corpus bindings, binaries, and the",
        "published projections; a missing raw report may be represented by its zstd",
        "fallback. Four process-isolated ABBA legs are shown separately for each",
        "selector. These are within-process samples and matched ABBA repeat",
        "observations; they make no host-population uncertainty or speedup claim.",
        "",
        "`p50` is the harness integer-nanosecond midpoint (floored for an even",
        "sample count); `p95` and `p99` are nearest-rank values. Latency cells are",
        "milliseconds. A positive paired percentage means the candidate elapsed",
        "value is lower than the matched control value under the verifier formula.",
        "",
        "## Normal latency by ABBA leg",
        "",
    ]
    metric_rows: list[list[str]] = []
    for selector in SELECTORS:
        elapsed = rows[selector]["elapsed"]
        legs = rows[selector]["legs"]
        changes = _object(
            elapsed["candidate_reduction_percent"],
            f"{selector}.candidate_reduction_percent",
        )
        drift = _object(
            elapsed["same_implementation_drift_percent"],
            f"{selector}.same_implementation_drift_percent",
        )
        for statistic in STATISTICS:
            metric_rows.append([
                selector,
                statistic,
                *[
                    _fmt_ns(_object(legs[LEG_KEYS[leg]], f"{selector}.{leg}")[statistic])
                    for leg in LEGS
                ],
                _fmt_percent(_object(changes["a1_to_b1"], "a1_to_b1")[statistic]),
                _fmt_percent(_object(changes["a2_to_b2"], "a2_to_b2")[statistic]),
                _fmt_percent(_object(drift["control"], "control")[statistic]),
                _fmt_percent(_object(drift["candidate"], "candidate")[statistic]),
            ])
    lines.extend(_markdown_table(
        metric_rows,
        [
            "Selector", "Statistic", "A1 control", "B1 candidate",
            "B2 candidate", "A2 control", "A1→B1 change", "A2→B2 change",
            "Control drift", "Candidate drift",
        ],
    ))
    lines.extend(["", "## ABBA decisions", ""])
    decision_rows: list[list[str]] = []
    for selector in SELECTORS:
        elapsed = rows[selector]["elapsed"]
        rejected = elapsed["rejected_statistics"]
        rejected_text = "; ".join(
            f"{name}: {reason}" for name, reason in rejected.items()
        ) or "none"
        decision_rows.append([
            selector,
            ", ".join(elapsed["accepted_statistics"]) or "none",
            rejected_text,
            ", ".join(elapsed["adverse_both_statistics"]) or "none",
        ])
    lines.extend(_markdown_table(
        decision_rows,
        ["Selector", "Accepted statistics", "Rejected statistics", "Adverse in both pairs"],
    ))
    lines.extend([
        "",
        "## Allocation and RSS guards",
        "",
        "Allocator latency is excluded. Allocation fields are process-leg aggregate",
        "observations from the allocator lane; phase selectors expose unavailable",
        "allocation attribution. `allocated_bytes` and whole-process RSS pairs marked",
        "REVIEW exceed the five percent review threshold. RSS is GNU `time -v` maximum",
        "resident set size for the complete process, including setup and warmups.",
        "Resource percentages use candidate-minus-control, so positive values mean",
        "more candidate resource use.",
        "",
    ])
    resource_rows: list[list[str]] = []
    allocation_pairs = _object(
        allocation["matched_pairs"], "allocation-metrics.json.matched_pairs"
    )
    rss = _object(allocation["rss"], "allocation-metrics.json.rss")
    rss_lanes = _object(rss["lanes"], "allocation-metrics.json.rss.lanes")
    for selector in SELECTORS:
        selector_pairs = _object(
            allocation_pairs[selector], f"allocation-metrics.json.matched_pairs.{selector}"
        )
        normal_rss = _object(
            rss_lanes["normal"][selector],
            f"allocation-metrics.json.rss.lanes.normal.{selector}",
        )
        allocator_rss = _object(
            rss_lanes["allocator"][selector],
            f"allocation-metrics.json.rss.lanes.allocator.{selector}",
        )
        for pair_name, _, _ in PAIRS:
            pair = _object(selector_pairs[pair_name], f"{selector}.{pair_name}")
            metrics = _object(pair["metrics"], f"{selector}.{pair_name}.metrics")
            normal_pair = _object(
                normal_rss[pair_name], f"normal RSS {selector}.{pair_name}"
            )
            allocator_pair = _object(
                allocator_rss[pair_name], f"allocator RSS {selector}.{pair_name}"
            )
            resource_rows.append([
                selector,
                pair_name,
                _fmt_delta(_object(metrics["allocation_calls"], "allocation_calls")),
                _fmt_delta(_object(metrics["allocated_bytes"], "allocated_bytes")),
                _fmt_rss(normal_pair),
                _fmt_rss(allocator_pair),
            ])
    lines.extend(_markdown_table(
        resource_rows,
        [
            "Selector", "Pair", "Allocation calls (control → candidate)",
            "Allocated bytes (control → candidate)", "Normal RSS", "Allocator RSS",
        ],
    ))
    lines.extend([
        "",
        "The retained `live_bytes_before`, `live_bytes_after`,",
        "`peak_live_bytes_before`, and `peak_live_bytes_after` values are literal",
        "snapshots. They are audited as snapshots and are not interpreted as an",
        "operation-local peak or a full output-retention memory measurement.",
        "",
        "Deterministic within-process uncertainty details, including exact IID median",
        "order-statistic intervals and seeded percentile bootstrap intervals, are in",
        "[`uncertainty.json`](uncertainty.json). The full raw vectors remain in the",
        "validated reports; [`summary.json`](summary.json) retains their checked statistics.",
        "",
        f"Verifier source: [`verify.py`](verify.py). Rendered from {summary['change']}.",
    ])
    return "\n".join(lines) + "\n"


def _write_json(path: Path, value: dict[str, Any]) -> None:
    try:
        payload = json.dumps(
            value, indent=2, sort_keys=True, ensure_ascii=False, allow_nan=False
        ) + "\n"
    except (TypeError, ValueError, OverflowError) as error:
        _fail(str(path), f"cannot serialize deterministic JSON: {error}")
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(payload, encoding="utf-8")
    temporary.replace(path)


def _write_text(path: Path, text: str) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(text, encoding="utf-8")
    temporary.replace(path)


def _uncertainty_document(
    root: Path,
    summary: dict[str, Any],
    allocation: dict[str, Any],
    normal: dict[str, Any],
) -> dict[str, Any]:
    provenance = {
        "summary": {
            "path": "summary.json",
            "sha256": hashlib.sha256(
                (root / "summary.json").read_bytes()
            ).hexdigest(),
        },
        "allocation": {
            "path": "allocation-metrics.json",
            "sha256": hashlib.sha256(
                (root / "allocation-metrics.json").read_bytes()
            ).hexdigest(),
        },
    }
    return {
        "schema_version": 1,
        "change": "0418",
        "status": "descriptive_uncertainty_only",
        "scope": {
            "latency": (
                "retained samples within each fresh process leg after warmups; "
                "A1/B1/B2/A2 are kept separate"
            ),
            "abba": "paired directions and same-revision repeat drift; no pooled speedup claim",
            "allocation": "allocator-lane process-leg aggregate vectors; allocator latency excluded",
            "rss": "whole-process GNU time -v maximum resident set size, including setup and warmups",
            "population_claim": "none; intervals do not model shared-host or host-population uncertainty",
        },
        "statistics": {
            "p50": "harness integer-nanosecond midpoint; exact IID order-statistic interval plus bootstrap diagnostic",
            "mean": "producer Student-t interval plus deterministic percentile bootstrap interval",
            "p95": "nearest-rank statistic with deterministic percentile bootstrap interval",
            "p99": "nearest-rank statistic with deterministic percentile bootstrap interval",
        },
        "bootstrap": {
            "resamples": BOOTSTRAP_RESAMPLES,
            "algorithm": "fixed 64-bit linear-congruential generator; nearest-rank percentile endpoints at 2.5% and 97.5%",
            "seed_derivation": "SHA-256(change-0418, selector, ABBA leg), first eight bytes",
        },
        "provenance": provenance,
        "normal_abba": normal,
        "allocation_projection": {
            "path": "allocation-metrics.json",
            "timing_status": allocation["timing_status"],
            "latency_comparison": allocation["latency_comparison"],
            "review_threshold_percent": allocation["review_threshold_percent"],
            "nonclaimable_metrics": allocation["nonclaimable_metrics"],
            "validation": allocation["validation"],
        },
        "replay": {
            "summary_timing_status": summary["timing_status"],
            "raw_vectors_retained": True,
            "compressed_report_fallback_supported_by_verifier": True,
        },
    }


def render(root: Path) -> None:
    root = root.resolve()
    summary, allocation, verifier = _load_verified_projections(root)
    raw_samples = _load_formal_samples(root, summary, verifier)
    rows = _validate_projections(summary, allocation, raw_samples)
    normal = _normal_uncertainty(rows)
    uncertainty = _uncertainty_document(root, summary, allocation, normal)
    _write_text(root / "result-table.md", _render_table(summary, allocation, rows))
    _write_json(root / "uncertainty.json", uncertainty)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="0418 evidence bundle root (default: this script's directory)",
    )
    args = parser.parse_args()
    try:
        render(args.root)
    except (RenderError, OSError) as error:
        print(f"0418 render failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps({
        "status": "pass",
        "change": "0418",
        "outputs": ["result-table.md", "uncertainty.json"],
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
