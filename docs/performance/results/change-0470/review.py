#!/usr/bin/env python3
"""Recompute descriptive full-guard flags and paired normal-process RSS drift."""

from __future__ import annotations

import json
import math
import copy
from pathlib import Path
from typing import Any, Mapping

import analyze
import verify


ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-0470-review-v1"
FULL_LANES = ("A-full", "B-full")
NORMAL_LANES = ("A1", "B1", "B2", "A2")
FULL_SAMPLES = 15
FULL_WARMUPS = 3
RSS_PAIRS = (
    ("control_repetition", "A1", "A2"),
    ("candidate_repetition", "B1", "B2"),
    ("a1_control_to_b1_candidate", "A1", "B1"),
    ("a2_control_to_b2_candidate", "A2", "B2"),
)


def require(condition: bool, message: str) -> None:
    if not condition:
        raise verify.VerificationError(message)


def _mean(samples: list[int]) -> float:
    require(bool(samples), "full guard elapsed sample vector is empty")
    return math.fsum(samples) / len(samples)


def _latencies(
    row: Mapping[str, Any], label: str, samples: int, comparator: Any
) -> dict[str, float]:
    values = verify.obj(row.get("elapsed_ns"), f"{label}.elapsed_ns")
    raw = values.get("samples")
    if not isinstance(raw, list) or len(raw) != samples:
        raise verify.VerificationError(f"{label}.elapsed_ns.samples cardinality differs")
    if any(isinstance(value, bool) or not isinstance(value, int) or value <= 0 for value in raw):
        raise verify.VerificationError(f"{label}.elapsed_ns.samples are malformed")
    # _latencies validates the recorded p50/p95/p99 against the raw vectors.
    percentiles = comparator._latencies(dict(row), label, samples)
    return {"mean": _mean(raw), **percentiles}


def _delta_percent(baseline: float, candidate: float) -> float:
    if baseline == 0:
        return 0.0 if candidate == 0 else math.inf
    return (candidate / baseline - 1.0) * 100.0


def _metadata(report: Mapping[str, Any], lane: str) -> dict[str, Any]:
    environment = verify.obj(report.get("environment"), f"{lane}.environment")
    binary = verify.obj(report.get("binary_identity"), f"{lane}.binary_identity")
    configuration = verify.obj(report.get("configuration"), f"{lane}.configuration")
    return {
        "schema_version": report.get("schema_version"),
        "tool": report.get("tool"),
        "revision": environment.get("git_revision"),
        "worktree_dirty": environment.get("git_worktree_dirty"),
        "binary_sha256": binary.get("binary_sha256"),
        "binary_bytes": binary.get("binary_bytes"),
        "samples_per_case": configuration.get("samples_per_case"),
        "warmup_iterations_per_case": configuration.get("warmup_iterations_per_case"),
    }


def _rss(root: Path, lane: str) -> dict[str, Any]:
    return analyze.parse_gnu_rss(
        root / lane / "resource.log",
        root=root,
        latency_comparison="descriptive_only",
    )


def _rss_drift(rss: Mapping[str, Mapping[str, Any]], name: str, baseline_lane: str, current_lane: str) -> dict[str, Any]:
    baseline = rss[baseline_lane]
    current = rss[current_lane]
    baseline_bytes = int(baseline["maximum_resident_set_bytes"])
    current_bytes = int(current["maximum_resident_set_bytes"])
    delta = _delta_percent(float(baseline_bytes), float(current_bytes))
    return {
        "name": name,
        "baseline_lane": baseline_lane,
        "current_lane": current_lane,
        "baseline_kib": baseline["maximum_resident_set_kib"],
        "current_kib": current["maximum_resident_set_kib"],
        "baseline_bytes": baseline_bytes,
        "current_bytes": current_bytes,
        "delta_percent": delta if math.isfinite(delta) else None,
        "delta_is_infinite": math.isinf(delta),
    }


def evaluate(root: Path = ROOT) -> dict[str, Any]:
    root = root.resolve()
    protocol = verify.verify_protocol(root)
    policy = verify.read_json(root / "report-policy.json", "report-policy.json")
    comparator, comparator_path = analyze.load_comparator(root)
    bindings = {role: verify.verify_role_binding(root, role) for role in ("control", "candidate")}

    full_reports = [verify.read_json(root / lane / "report.json", f"{lane}/report.json") for lane in FULL_LANES]
    verify.verify_report_metadata(
        full_reports[0], "A-full.report", bindings["control"], FULL_SAMPLES, FULL_WARMUPS
    )
    verify.verify_report_metadata(
        full_reports[1], "B-full.report", bindings["candidate"], FULL_SAMPLES, FULL_WARMUPS
    )

    # This is the unchanged canonical full guard.  Its result is retained only
    # as a compact identity/status summary; raw reports remain the source for
    # the descriptive mean and percentile flags below.
    canonical_guard = analyze.compare_guard(
        full_reports[0],
        full_reports[1],
        policy,
        root=root,
        report_paths=[root / lane / "report.json" for lane in FULL_LANES],
    )
    comparison = verify.obj(canonical_guard.get("comparison"), "canonical full guard comparison")
    canonical_status = comparison.get("status")
    require(canonical_status in {"pass", "regression"}, "canonical full guard status is invalid")

    before_rows = verify._report_rows(full_reports[0], "A-full.report")
    after_rows = verify._report_rows(full_reports[1], "B-full.report")
    require(set(before_rows) == set(after_rows), "full guard row identity differs")
    flags: list[dict[str, Any]] = []
    for key in sorted(before_rows):
        before = before_rows[key]
        after = after_rows[key]
        case, corpus_bytes = key
        require(verify.canonical(before["corpus"]) == corpus_bytes, f"A-full.{case}: corpus identity differs")
        require(verify.canonical(after["corpus"]) == corpus_bytes, f"B-full.{case}: corpus identity differs")
        baseline_stats = _latencies(before, f"A-full.{case}", FULL_SAMPLES, comparator)
        candidate_stats = _latencies(after, f"B-full.{case}", FULL_SAMPLES, comparator)
        for metric in ("mean", "p50", "p95", "p99"):
            baseline = baseline_stats[metric]
            candidate = candidate_stats[metric]
            delta = _delta_percent(baseline, candidate)
            if delta > 5.0:
                flags.append(
                    {
                        "case": case,
                        "corpus": before["corpus"],
                        "metric": metric,
                        "baseline_ns": baseline,
                        "candidate_ns": candidate,
                        "delta_percent": delta if math.isfinite(delta) else None,
                        "delta_is_infinite": math.isinf(delta),
                    }
                )

    normal_reports = {
        lane: verify.read_json(root / lane / "report.json", f"{lane}/report.json")
        for lane in NORMAL_LANES
    }
    for lane in NORMAL_LANES:
        role = "control" if lane in ("A1", "A2") else "candidate"
        verify.verify_report_metadata(
            normal_reports[lane],
            f"{lane}.report",
            bindings[role],
            int(protocol["probe_samples"]),
            int(protocol["probe_warmups"]),
            cases=protocol["probe_cases"],
            shapes=protocol["probe_shapes"],
        )
    rss = {lane: _rss(root, lane) for lane in NORMAL_LANES}
    paired_rss = {
        name: _rss_drift(rss, name, baseline_lane, current_lane)
        for name, baseline_lane, current_lane in RSS_PAIRS
    }

    return {
        "schema": SCHEMA,
        "scope": "Descriptive full-guard raw-sample flags above +5%; no registered latency or RSS claim",
        "threshold_percent": 5.0,
        "full_guard": {
            "rows": len(before_rows),
            "samples": FULL_SAMPLES,
            "warmups": FULL_WARMUPS,
            "canonical_status": canonical_status,
            "canonical_summary": comparison.get("summary"),
            "flags_above_threshold": flags,
        },
        "full_guard_report_metadata": {
            lane: _metadata(report, lane) for lane, report in zip(FULL_LANES, full_reports)
        },
        "normal_process_rss": rss,
        "paired_normal_rss_drift": paired_rss,
        "inputs": {
            "full_reports": [analyze.file_binding(root / lane / "report.json", root) for lane in FULL_LANES],
            "normal_reports": [analyze.file_binding(root / lane / "report.json", root) for lane in NORMAL_LANES],
            "normal_resource_logs": [analyze.file_binding(root / lane / "resource.log", root) for lane in NORMAL_LANES],
            "policy": analyze.file_binding(root / "report-policy.json", root),
            "protocol": analyze.file_binding(root / "protocol.json", root),
            "analyzer": analyze.file_binding(root / "analyze.py", root),
            "verifier": analyze.file_binding(root / "verify.py", root),
            "comparator": analyze.file_binding(comparator_path, root),
        },
    }


def _without_verifier_binding(value: Mapping[str, Any], label: str) -> dict[str, Any]:
    """Project a review summary for the one permitted helper-binding change."""

    projected = copy.deepcopy(dict(value))
    inputs = verify.obj(projected.get("inputs"), f"{label}.inputs")
    if "verifier" not in inputs:
        verify.fail(f"{label}.inputs.verifier: binding is missing")
    del inputs["verifier"]
    return projected


def historical_metadata_evolution(
    historical: Mapping[str, Any], current: Mapping[str, Any]
) -> dict[str, Any]:
    """Validate the sealed review against the current helper output.

    The original review summary is frozen by ``guard-protocol.json``.  The
    verifier grew additional integration checks afterward, so only its input
    file binding may evolve; every report, statistic, oracle, policy, and
    protocol field remains covered by the projected equality check.
    """

    historical_inputs = verify.obj(historical.get("inputs"), "historical review.inputs")
    current_inputs = verify.obj(current.get("inputs"), "current review.inputs")
    historical_verifier = copy.deepcopy(
        verify.obj(historical_inputs.get("verifier"), "historical review.inputs.verifier")
    )
    current_verifier = copy.deepcopy(
        verify.obj(current_inputs.get("verifier"), "current review.inputs.verifier")
    )
    if historical_verifier.get("path") != current_verifier.get("path"):
        verify.fail("historical review: verifier binding path differs")
    if not verify.canonical_equal(
        _without_verifier_binding(historical, "historical review"),
        _without_verifier_binding(current, "current review"),
    ):
        verify.fail(
            "historical review: semantic payload differs outside inputs.verifier"
        )
    return {
        "allowed_field": "inputs.verifier",
        "reason": (
            "The historical review summary was sealed before the verifier's "
            "later integration checks; only this helper metadata binding evolved."
        ),
        "historical_verifier": historical_verifier,
        "current_verifier": current_verifier,
    }


def final_review_summary(root: Path = ROOT, current: Mapping[str, Any] | None = None) -> dict[str, Any]:
    """Return the current review with an explicit frozen-history note."""

    root = root.resolve()
    if current is None:
        current = evaluate(root)
    historical = verify.read_json(root / "review-summary.json", "historical full guard review summary")
    result = copy.deepcopy(dict(current))
    result["historical_metadata_evolution"] = historical_metadata_evolution(historical, current)
    return result


if __name__ == "__main__":
    result = evaluate()
    final = final_review_summary(ROOT, result)
    (ROOT / "final-review-summary.json").write_text(
        json.dumps(final, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    print(json.dumps({"status": "pass", "flags": len(result["full_guard"]["flags_above_threshold"])}, sort_keys=True))
