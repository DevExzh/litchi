#!/usr/bin/env python3
"""Validate and decide the 0722 paired DOCX structural-scan pilot.

The analyzer rechecks every child receipt, delegates the detailed DOCX report
and allocation schema checks to the reviewed 0709 validators, and compares
paired native observations plus the first-cycle allocator gauges.  It emits a
decision even when a measured performance gate fails, so a negative result
remains reviewable evidence.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import statistics
import sys
from typing import Any

sys.dont_write_bytecode = True


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import pilot as CAPTURE  # noqa: E402


def load_module(path: Path, name: str) -> Any:
    original_path = list(sys.path)
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise AssertionError(f"cannot load validator: {path}")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    finally:
        # The historical 0713 analyzer imports its sibling ``capture`` module
        # by basename and prepends that packet directory to sys.path.  Keep
        # that private module binding, but do not let its path affect the
        # packet's explicit 0722 custody loader.
        sys.path[:] = original_path
    return module


LEGACY = load_module(REPO / "docs/performance/results/change-0709/analyze.py",
                     "ordinary_save_0709_for_0722")
PARITY = load_module(REPO / "docs/performance/results/change-0713/analyze.py",
                     "pair_normalization_0713_for_0722")

PHASES = CAPTURE.PHASES
LANES = CAPTURE.LANES
STAGES = CAPTURE.STAGES
ALLOCATOR_STAGES = CAPTURE.ALLOCATOR_STAGES
PAIR_STAGES = {
    "pair-1": ("baseline-A1", "candidate-B1"),
    "pair-2": ("baseline-A2", "candidate-B2"),
    "pair-3": ("baseline-A3", "candidate-B3"),
    "pair-4": ("baseline-A4", "candidate-B4"),
}
ALLOCATION_FIELDS = {
    "request_count": "allocation_calls",
    "requested_bytes": "allocated_bytes",
}
ALLOCATION_DERIVED_FIELDS = {
    "peak_above_start": ("region_peak_live_bytes", "live_bytes_before"),
    "net_live": ("live_bytes_after", "live_bytes_before"),
}
READ_CONTROL_IDS = (
    "generated-medium-list-paragraphs",
    "pinned-media-eager-paragraph-count",
)
PROCESS_DELTA_FIELDS = (
    "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
    "syscr", "syscw", "minor_faults", "major_faults", "user_cpu_ticks",
    "system_cpu_ticks", "clock_ticks_per_second", "voluntary_context_switches",
    "nonvoluntary_context_switches", "rss_bytes", "peak_rss_bytes",
)
EXPECTED_PROCESS_PROBE_SCOPE = (
    "ordinary_save_phase_interval_same_process_counters_including_procfs_probe_overhead"
)
EXPECTED_PROCESS_PROBE_LATENCY = "diagnostic_only_procfs_probe_instrumentation_latency"
EXPECTED_PROCESS_PROBE_CONTROL_SCOPE = (
    "fixed_32_empty_adjacent_procfs_snapshot_pairs_acquired_before_warmups_and_never_subtracted"
)
HEX = set("0123456789abcdef")


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing evidence file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def check_hex(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def finite(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def source_delta(baseline: dict[str, str], candidate: dict[str, str]) -> list[str]:
    return sorted(name for name in set(baseline) | set(candidate)
                  if baseline.get(name) != candidate.get(name))


def cleanup_witnesses() -> list[dict[str, Any]]:
    return CAPTURE.cleanup_witnesses()


def validate_sources(plan: dict[str, Any]) -> tuple[
    dict[str, str], dict[str, str], list[str], dict[str, str] | None, dict[str, Any] | None
]:
    baseline, candidate, changed = CAPTURE.source_pair()
    require(set(changed) <= set(plan["source_delta_allowlist"]),
            f"source delta exceeds allowlist: {changed}")
    current = CAPTURE.source_census()
    final_path = HERE / "source-final.json"
    disposition_path = HERE / "disposition.json"
    require(final_path.exists() == disposition_path.exists(),
            "source-final.json and disposition.json must be created together")
    if final_path.exists():
        final_source = read(final_path)
        disposition = read(disposition_path)
        require(isinstance(final_source, dict) and final_source,
                "source-final.json is not a raw source map")
        require(isinstance(disposition, dict)
                and set(disposition) == {"retained", "final_source"},
                "disposition.json must contain retained and final_source")
        final_label = disposition.get("final_source")
        require(final_label in {"baseline", "candidate"},
                "disposition final_source must be baseline or candidate")
        expected = baseline if final_label == "baseline" else candidate
        require(final_source == expected,
                "source-final.json does not match its declared final source")
        require(current == final_source,
                "current checkout does not match source-final.json")
        require(isinstance(disposition.get("retained"), bool)
                and disposition.get("retained") == (final_label == "candidate"),
                "disposition retained flag is inconsistent with final_source")
        return baseline, candidate, changed, final_source, disposition
    require(current == candidate, "current checkout is not the frozen live candidate source")
    return baseline, candidate, changed, None, None


def validate_constraints() -> str:
    path = HERE / "constraints.json"
    constraints = read(path)
    require(isinstance(constraints, dict), "constraints is not an object")
    for name, expected in constraints.items():
        target = REPO / name
        require(target.is_file() and sha(target) == expected,
                f"constraint changed: {name}")
    return sha(path)


def validate_builds(source_label: str, lane: str) -> dict[str, Any]:
    return CAPTURE.build_info(source_label, lane)


def expected_command(plan: dict[str, Any], build: dict[str, Any],
                     corpus: dict[str, Any], phase: str, name: str) -> list[str]:
    return CAPTURE.expected_command(plan, build, corpus, phase, name)


def fixture_binding(corpus: dict[str, Any]) -> dict[str, Any] | None:
    _, binding = CAPTURE.fixture_info(corpus)
    return binding


def validate_read_controls(
    plan: dict[str, Any], disposition: dict[str, Any],
    baseline: dict[str, str], candidate: dict[str, str],
) -> dict[str, Any]:
    """Validate the independent read-control decision used for finalization.

    This intentionally reads the control analyzer's direct result rather than
    the primary ``analysis.json``.  A restored baseline is therefore justified
    by an independently source-bound failed control gate, without creating an
    analysis-output dependency cycle.
    """

    path = HERE / "read-controls-analysis.json"
    value = read(path)
    require(value.get("schema_version") == 1
            and value.get("packet") == "change-0722-docx-read-controls"
            and value.get("revision") == plan["revision"],
            "read-control analysis identity changed")
    scope = value.get("scope")
    require(isinstance(scope, dict)
            and scope.get("stages") == list(STAGES)
            and scope.get("controls") == list(READ_CONTROL_IDS)
            and scope.get("capture_live_source") == "candidate"
            and scope.get("final_source") == disposition["final_source"],
            "read-control analysis scope or final source changed")
    require(value.get("plan_sha256") == sha(HERE / "read-controls-plan.json")
            and value.get("capture_script_sha256") == sha(HERE / "read-controls.py")
            and value.get("analysis_script_sha256")
            == sha(HERE / "read-controls-analyze.py")
            and value.get("constraints_sha256") == sha(HERE / "constraints.json"),
            "read-control analysis binding changed")
    decision = value.get("decision")
    require(isinstance(decision, dict), "read-control decision is missing")
    hard_gates = decision.get("hard_gates")
    require(isinstance(hard_gates, list) and hard_gates,
            "read-control hard gates are missing")
    require(all(isinstance(item, dict) and isinstance(item.get("pass"), bool)
                for item in hard_gates),
            "read-control hard gate result is malformed")
    all_gates_pass = all(item["pass"] for item in hard_gates)
    require(decision.get("all_hard_gates_pass") == all_gates_pass
            and decision.get("accepted") == all_gates_pass,
            "read-control decision is inconsistent with its actual gates")
    require(decision.get("deterministic_output_parity_pass") is True,
            "read-control semantic parity did not pass")
    verification = value.get("verification")
    require(isinstance(verification, dict)
            and verification.get("final_source_matches_checkout") is True
            and verification.get("final_source_label") == disposition["final_source"],
            "read-control final source witness is missing")
    manifests = value.get("source_manifests")
    require(isinstance(manifests, dict), "read-control source manifests are missing")
    for label, source in (("baseline", baseline), ("candidate", candidate)):
        item = manifests.get(label)
        require(isinstance(item, dict)
                and item.get("sha256") == sha(HERE / f"source-{label}.json")
                and item.get("entry_count") == len(source),
                f"read-control {label} source manifest binding changed")
    return {
        "path": path.name,
        "sha256": sha(path),
        "analysis_script_sha256": sha(HERE / "read-controls-analyze.py"),
        "accepted": bool(decision["accepted"]),
        "all_hard_gates_pass": all_gates_pass,
        "failed_hard_gate_count": sum(not item["pass"] for item in hard_gates),
        "deterministic_output_parity_pass": True,
    }


def validate_receipt(plan: dict[str, Any], job: dict[str, Any], build: dict[str, Any],
                     baseline: dict[str, str], candidate: dict[str, str],
                     plan_sha: str, script_sha: str,
                     constraints_sha: str) -> dict[str, Any]:
    name = job["name"]
    receipt = read(HERE / f"{name}.receipt.json")
    metadata = CAPTURE.stage_metadata(plan, job["stage"])
    expected_fields = {
        "schema_version": 1,
        "packet": plan["packet"],
        "name": name,
        "stage": job["stage"],
        "source_label": metadata["source"],
        "build_label": metadata["build"],
        "pair": metadata["pair"],
        "cycle": metadata["cycle"],
        "lane": job["lane"],
        "stage_order": metadata["order"],
        "stage_order_index": job["stage_order_index"],
        "corpus_id": job["corpus"]["id"],
        "corpus_label": job["corpus"]["label"],
        "corpus_origin": job["corpus"]["origin"],
        "phase": job["phase"],
        "case": job["case"],
        "samples": plan[job["lane"]]["samples"],
        "warmup": plan[job["lane"]]["warmup"],
        "cpu": plan["cpu"],
        "fresh_child_process": True,
    }
    for key, expected in expected_fields.items():
        require(receipt.get(key) == expected, f"{name}: receipt {key} changed")
    require(receipt.get("exit_code") == 0, f"{name}: child failed")
    require(receipt.get("binary_path") == build["binary"]
            and receipt.get("binary_sha256") == build["binary_sha256"]
            and receipt.get("binary_bytes") == build["binary_bytes"],
            f"{name}: binary binding changed")
    require(receipt.get("build_record") == f"build-{metadata['build']}.json"
            and receipt.get("build_record_sha256") == build["record_sha256"],
            f"{name}: build record binding changed")
    require(receipt.get("build_source_manifest") == build["source_manifest"]
            and receipt.get("build_source_manifest_sha256") == build["source_manifest_sha256"],
            f"{name}: binary source manifest binding changed")
    require(receipt.get("plan_sha256") == plan_sha, f"{name}: plan binding changed")
    require(receipt.get("script_sha256") == script_sha, f"{name}: capture script changed")
    require(receipt.get("constraints_sha256") == constraints_sha,
            f"{name}: constraints binding changed")
    check_hex(receipt.get("binary_sha256"), f"{name}.binary_sha256")
    finite(receipt.get("seconds"), f"{name}.seconds")
    require(receipt["seconds"] >= 0.0, f"{name}.seconds is negative")

    expected_binary_source = baseline if metadata["source"] == "baseline" else candidate
    binary_source = receipt.get("binary_source")
    require(isinstance(binary_source, dict), f"{name}: binary source record is missing")
    require(binary_source == receipt.get("retained_binary_source"),
            f"{name}: binary source aliases disagree")
    require(binary_source.get("manifest") == build["source_manifest"]
            and binary_source.get("manifest_sha256") == build["source_manifest_sha256"]
            and binary_source.get("source_census_sha256") == digest(expected_binary_source)
            and binary_source.get("source_entry_count") == len(expected_binary_source),
            f"{name}: frozen binary source custody changed")

    live_source = receipt.get("live_checkout_source")
    require(isinstance(live_source, dict)
            and live_source == receipt.get("current_checkout_source"),
            f"{name}: live source aliases disagree")
    require(live_source.get("manifest") == "source-candidate.json"
            and live_source.get("manifest_sha256") == sha(HERE / "source-candidate.json")
            and live_source.get("source_census_sha256") == digest(candidate)
            and live_source.get("source_entry_count") == len(candidate)
            and live_source.get("before_sha256") == digest(candidate)
            and live_source.get("after_sha256") == digest(candidate)
            and live_source.get("recensus_before") is True
            and live_source.get("recensus_after") is True
            and live_source.get("unchanged_during_child") is True,
            f"{name}: live candidate source custody changed")
    for relation_name in ("relation_before", "relation_after"):
        relation = live_source.get(relation_name)
        require(isinstance(relation, dict) and relation.get("mode") == "exact"
                and relation.get("changed_paths") == []
                and relation.get("expected_sha256") == digest(candidate)
                and relation.get("current_sha256") == digest(candidate),
                f"{name}: {relation_name} source recensus changed")

    fixture = receipt.get("fixture")
    require(isinstance(fixture, dict), f"{name}: fixture receipt is missing")
    planned = fixture_binding(job["corpus"])
    if planned is None:
        require(fixture.get("plan_path") is None
                and fixture.get("plan_sha256") is None
                and fixture.get("before") is None
                and fixture.get("after") is None,
                f"{name}: generated fixture binding changed")
    else:
        require(fixture.get("plan_path") == planned["path"]
                and fixture.get("plan_sha256") == planned["sha256"]
                and fixture.get("before") == planned
                and fixture.get("after") == planned,
                f"{name}: real fixture custody changed")
    require(receipt.get("command") == expected_command(
        plan, build, job["corpus"], job["phase"], name),
            f"{name}: command changed")
    environment = receipt.get("environment")
    require(environment == plan["environment"], f"{name}: child environment changed")
    artifacts = receipt.get("artifacts")
    expected_artifacts = {f"{name}.json", f"{name}.stdout", f"{name}.stderr"}
    require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts,
            f"{name}: artifact inventory changed")
    for filename, file_digest in artifacts.items():
        check_hex(file_digest, f"{name}:{filename}")
        target = HERE / filename
        require(target.is_file() and not target.is_symlink() and sha(target) == file_digest,
                f"{name}: artifact digest changed: {filename}")
    return receipt


def validate_process_delta(value: Any, label: str) -> dict[str, int] | None:
    if value is None:
        return None
    require(isinstance(value, dict) and set(value) == set(PROCESS_DELTA_FIELDS),
            f"{label}: process delta fields changed")
    result: dict[str, int] = {}
    for field in PROCESS_DELTA_FIELDS:
        nonnegative_int(value[field], f"{label}.{field}")
        result[field] = value[field]
    positive_int(result["clock_ticks_per_second"], f"{label}.clock_ticks_per_second")
    return result


def process_trace_summary(result: dict[str, Any], job: dict[str, Any]) -> dict[str, Any]:
    """Retain optional 0717 process-probe traces as diagnostic evidence only."""

    ordinary = result.get("source", {}).get("ordinary_save")
    require(isinstance(ordinary, dict), f"{job['name']}: ordinary-save evidence is missing")
    probe = ordinary.get("process_probe")
    if probe is None:
        return {"status": "absent", "raw_sha256": None}
    require(isinstance(probe, dict), f"{job['name']}: process probe is malformed")
    expected_keys = {
        "phase", "timing_scope", "scope", "latency_claim", "control_scope",
        "fixed_count", "empty_adjacent_snapshot_controls", "sample_deltas",
    }
    require(set(probe) == expected_keys, f"{job['name']}: process probe fields changed")
    require(probe.get("scope") == EXPECTED_PROCESS_PROBE_SCOPE
            and probe.get("latency_claim") == EXPECTED_PROCESS_PROBE_LATENCY
            and probe.get("control_scope") == EXPECTED_PROCESS_PROBE_CONTROL_SCOPE
            and probe.get("fixed_count") == 32,
            f"{job['name']}: process probe identity changed")
    require(probe.get("phase") == LEGACY.PHASE_LABELS[job["phase"]]
            and probe.get("timing_scope") == LEGACY.PHASE_TIMING_SCOPES[job["phase"]],
            f"{job['name']}: process probe phase scope changed")
    controls = probe["empty_adjacent_snapshot_controls"]
    samples = probe["sample_deltas"]
    require(isinstance(controls, list) and len(controls) == 32,
            f"{job['name']}: process probe control count changed")
    require(isinstance(samples, list) and len(samples) == job["samples"],
            f"{job['name']}: process probe sample count changed")
    checked_controls = [validate_process_delta(value, f"{job['name']}.control[{index}]")
                        for index, value in enumerate(controls)]
    checked_samples = [validate_process_delta(value, f"{job['name']}.sample[{index}]")
                       for index, value in enumerate(samples)]
    present = [value is not None for value in checked_samples]
    require(all(present) or not any(present),
            f"{job['name']}: process probe sample availability is asymmetric")
    return {
        "status": "measured" if all(present) else "unavailable",
        "raw_sha256": digest(probe),
        "fixed_count": len(checked_controls),
        "available_controls": sum(value is not None for value in checked_controls),
        "available_samples": sum(value is not None for value in checked_samples),
        "controls_subtracted": False,
        "interpretation": "0717 process counters are retained as separate diagnostics; no latency or causal claim is made.",
    }


def normalized_result(result: dict[str, Any]) -> dict[str, Any]:
    """Use the reviewed 0713 parity normalization plus the 0717 probe field."""

    value = PARITY.normalized_result(result)
    ordinary = value.get("source", {}).get("ordinary_save")
    if isinstance(ordinary, dict):
        ordinary.pop("process_probe", None)
    return value


def allocation_summary(result: dict[str, Any], label: str) -> dict[str, Any]:
    envelope = result.get("operation_metrics", {}).get("allocation")
    require(isinstance(envelope, dict), f"{label}: allocation envelope is missing")
    values: dict[str, list[int]] = {}
    for field in (
        "allocation_calls",
        "allocated_bytes",
        "live_bytes_before",
        "live_bytes_after",
        "region_peak_live_bytes",
    ):
        metric = envelope.get(field)
        require(isinstance(metric, dict) and metric.get("status") == "measured",
                f"{label}: allocation {field} is not measured")
        vector = metric.get("values")
        require(isinstance(vector, list), f"{label}: allocation {field} vector is missing")
        require(all(isinstance(value, int) and not isinstance(value, bool)
                    for value in vector),
                f"{label}: allocation {field} vector contains a non-integer")
        values[field] = list(vector)
    sample_count = len(values["allocation_calls"])
    require(sample_count > 0, f"{label}: allocation vectors are empty")
    require(all(len(vector) == sample_count for vector in values.values()),
            f"{label}: allocation vector lengths differ")
    require(all(region >= before and region >= after
                for region, before, after in zip(
                    values["region_peak_live_bytes"],
                    values["live_bytes_before"],
                    values["live_bytes_after"],
                )), f"{label}: region peak invariant changed")
    values["peak_above_start"] = [
        peak - before
        for peak, before in zip(
            values["region_peak_live_bytes"], values["live_bytes_before"]
        )
    ]
    values["net_live"] = [
        after - before
        for after, before in zip(
            values["live_bytes_after"], values["live_bytes_before"]
        )
    ]
    return {
        "sample_count": sample_count,
        "values": values,
        "stats": {
            field: LEGACY.integer_stats(vector, f"{label}/{field}")
            for field, vector in values.items()
        },
        "diagnostic_note": "allocation request counts, bytes, peak above start, and net-live are per-operation observations; values are never summed across phases or samples",
    }


def validate_child(plan: dict[str, Any], job: dict[str, Any], build: dict[str, Any],
                   baseline: dict[str, str], candidate: dict[str, str],
                   plan_sha: str, script_sha: str,
                   constraints_sha: str) -> dict[str, Any]:
    receipt = validate_receipt(plan, job, build, baseline, candidate,
                               plan_sha, script_sha, constraints_sha)
    report = read(HERE / f"{job['name']}.json")
    build_identity = {
        "binary_sha256": build["binary_sha256"],
        "binary_bytes": build["binary_bytes"],
    }
    LEGACY.check_report_metadata(
        report, build_identity, job["lane"], job["case"],
        job["samples"], job["warmup"], job["name"],
    )
    result = report["results"][0]
    require(result.get("case") == job["case"], f"{job['name']}: result case changed")
    elapsed, elapsed_stats = LEGACY.validate_elapsed(
        result.get("elapsed_ns"), job["samples"], job["name"])
    ordinary = LEGACY.validate_ordinary_save(
        result, job["corpus"], job["phase"], job["lane"],
        job["samples"], job["name"],
    )
    LEGACY.validate_operation_metrics(
        result.get("operation_metrics"), result["elapsed_ns"],
        job["samples"], job["lane"], job["name"],
    )
    return {
        **job,
        "receipt": receipt,
        "report": report,
        "result": result,
        "ordinary": ordinary,
        "elapsed": elapsed,
        "elapsed_stats": elapsed_stats,
        "allocation": allocation_summary(result, job["name"])
        if job["lane"] == "allocator" else None,
        "process_trace": process_trace_summary(result, job),
        "normalized": normalized_result(result),
    }


def percent_delta(candidate: float, baseline: float) -> float:
    require(baseline > 0.0, "baseline metric must be positive")
    return (candidate / baseline - 1.0) * 100.0


def comparison(candidate: float, baseline: float) -> dict[str, float]:
    delta = percent_delta(candidate, baseline)
    return {
        "baseline": float(baseline),
        "candidate": float(candidate),
        "delta_percent": delta,
        "improvement_percent": -delta,
    }


def allocation_comparison(candidate: float, baseline: float, *, signed: bool) -> dict[str, Any]:
    """Compare allocation gauges while making a zero baseline explicit.

    A zero baseline has no meaningful percentage denominator.  Such a gate is
    therefore satisfied only by exact equality; a nonzero candidate remains a
    visible failed gate instead of aborting analysis.
    """

    candidate = float(candidate)
    baseline = float(baseline)
    if baseline == 0.0:
        return {
            "baseline": baseline,
            "candidate": candidate,
            "absolute_delta": candidate - baseline,
            "delta_percent": None,
            "improvement_percent": None,
            "baseline_zero": True,
            "baseline_equality": candidate == baseline,
        }
    delta = (candidate - baseline) / (abs(baseline) if signed else baseline) * 100.0
    return {
        "baseline": baseline,
        "candidate": candidate,
        "absolute_delta": candidate - baseline,
        "delta_percent": delta,
        "improvement_percent": -delta,
        "baseline_zero": False,
        "baseline_equality": None,
    }


def compare_elapsed(candidate: dict[str, Any], baseline: dict[str, Any]) -> dict[str, Any]:
    return {
        metric: comparison(float(candidate[metric]), float(baseline[metric]))
        for metric in ("p50", "mean", "p95", "p99")
    }


def compare_allocation(candidate: dict[str, Any], baseline: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for label, raw_field in ALLOCATION_FIELDS.items():
        result[label] = {
            metric: allocation_comparison(
                float(candidate["stats"][raw_field][metric]),
                float(baseline["stats"][raw_field][metric]),
                signed=False,
            )
            for metric in ("p50", "mean", "p95", "p99")
        }
        result[label]["source_field"] = raw_field
    for label, (candidate_field, baseline_field) in ALLOCATION_DERIVED_FIELDS.items():
        result[label] = {
            metric: allocation_comparison(
                float(candidate["stats"][label][metric]),
                float(baseline["stats"][label][metric]),
                signed=label == "net_live",
            )
            for metric in ("p50", "mean", "p95", "p99")
        }
        result[label]["source_fields"] = {
            "minuend": candidate_field,
            "subtrahend": baseline_field,
        }
        result[label]["source_field"] = f"{candidate_field} - {baseline_field}"
    return result


def allocation_gate_passes(
    comparison_row: dict[str, Any], label: str,
    thresholds: dict[str, Any],
) -> bool:
    """Return the hard-gate result for one allocator metric observation.

    Zero baselines are exact-equality gates because their percentage delta is
    undefined.  Net-live is a signed gauge and is accepted only when the
    candidate does not increase live bytes; the configured zero-percent
    threshold documents that rule but is not used as a denominator test.
    """

    require(label in ALLOCATION_FIELDS or label in ALLOCATION_DERIVED_FIELDS,
            f"unknown allocator gate label: {label}")
    require(isinstance(comparison_row, dict),
            f"{label}: allocator comparison row is not an object")
    require(isinstance(comparison_row.get("baseline_zero"), bool),
            f"{label}: allocator baseline-zero marker is missing")
    require(isinstance(comparison_row.get("baseline_equality"), (bool, type(None))),
            f"{label}: allocator baseline-equality marker is malformed")
    if comparison_row["baseline_zero"]:
        return comparison_row["baseline_equality"] is True

    if label == "net_live":
        require(thresholds.get("allocation_net_live_regression_percent") == 0,
                "net_live gate threshold must be zero")
        return comparison_row["candidate"] <= comparison_row["baseline"]

    threshold = thresholds.get("allocation_regression_percent")
    finite(threshold, "allocation_regression_percent")
    observed = comparison_row.get("delta_percent")
    require(observed is not None, f"{label}: nonzero allocator baseline has no delta")
    finite(observed, f"{label}.delta_percent")
    return observed <= threshold


def repeat_diagnostics(rows: dict[tuple[str, str, str, str], dict[str, Any]],
                       plan: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    diagnostics: list[dict[str, Any]] = []
    flags: list[dict[str, Any]] = []
    for lane in LANES:
        lane_stages = STAGES if lane == "native" else ALLOCATOR_STAGES
        for source_label in ("baseline", "candidate"):
            source_stages = [stage for stage in lane_stages
                             if CAPTURE.stage_metadata(plan, stage)["source"] == source_label]
            for corpus in plan["corpora"]:
                for phase in PHASES:
                    selected = [rows[(stage, lane, corpus["id"], phase)]
                                for stage in source_stages]
                    stats = [row["elapsed_stats"] for row in selected]
                    metric_rows: dict[str, Any] = {}
                    for metric in ("p50", "mean"):
                        values = [float(item[metric]) for item in stats]
                        require(values and min(values) > 0.0,
                                f"{lane}/{source_label}/{corpus['id']}/{phase}/{metric} invalid")
                        spread = (max(values) - min(values)) * 100.0 / min(values)
                        metric_rows[metric] = {
                            "values": values,
                            "spread_percent": spread,
                            "flag_over_5_percent": spread > plan["thresholds"]["repeat_drift_flag_percent"],
                        }
                        if spread > plan["thresholds"]["repeat_drift_flag_percent"]:
                            flags.append({
                                "lane": lane,
                                "source": source_label,
                                "corpus_id": corpus["id"],
                                "phase": phase,
                                "metric": metric,
                                "spread_percent": spread,
                                "flag_over_5_percent": True,
                            })
                    diagnostics.append({
                        "lane": lane,
                        "source": source_label,
                        "corpus_id": corpus["id"],
                        "phase": phase,
                        "stages": source_stages,
                        "metrics": metric_rows,
                    })
    return diagnostics, flags


def child_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    order = 0
    for stage in STAGES:
        lanes = ["native"] + (["allocator"] if stage in ALLOCATOR_STAGES else [])
        for lane in lanes:
            for corpus, phase in CAPTURE.ordered_jobs(plan, stage):
                result.append({
                    "stage": stage,
                    "lane": lane,
                    "corpus": corpus,
                    "phase": phase,
                    "case": CAPTURE.phase_case(corpus, phase),
                    "samples": plan[lane]["samples"],
                    "warmup": plan[lane]["warmup"],
                    "stage_order_index": CAPTURE.ordered_jobs(plan, stage).index((corpus, phase)),
                    "name": CAPTURE.child_name(stage, lane, corpus, phase),
                    "global_order_index": order,
                })
                order += 1
    return result


def validate_inventory(plan: dict[str, Any]) -> list[dict[str, Any]]:
    jobs = child_jobs(plan)
    expected_names = {job["name"] for job in jobs}
    receipt_paths = tuple(HERE.glob("*.receipt.json"))
    actual_receipts = {
        path.name[:-len(".receipt.json")]
        for path in receipt_paths
        if path.is_file() and not path.is_symlink()
        and path.name[:-len(".receipt.json")] in expected_names
    }
    primary_prefixes = tuple(
        f"{stage}-{lane}-"
        for stage in STAGES
        for lane in (("native", "allocator")
                     if stage in ALLOCATOR_STAGES else ("native",))
    )
    unknown_primary_receipts = sorted(
        path.name for path in receipt_paths
        if path.name.startswith(primary_prefixes)
        and path.name[:-len(".receipt.json")] not in expected_names
    )
    require(not unknown_primary_receipts,
            f"unknown primary receipt(s) retained: {unknown_primary_receipts}")
    require(actual_receipts == expected_names,
            f"receipt inventory differs; missing={sorted(expected_names - actual_receipts)} "
            f"extra={sorted(actual_receipts - expected_names)}")
    expected_suffixes = (".json", ".stdout", ".stderr", ".receipt.json")
    expected_artifacts = {f"{job['name']}{suffix}"
                          for job in jobs for suffix in expected_suffixes}
    for path in HERE.iterdir():
        if not path.is_file():
            continue
        if any(path.name.startswith(f"{name}.") for name in expected_names):
            require(path.name in expected_artifacts,
                    f"unexpected child artifact retained: {path.name}")
    return jobs


def build_pair_output(rows: dict[tuple[str, str, str, str], dict[str, Any]],
                      plan: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]], list[dict[str, Any]]]:
    paired: dict[str, Any] = {}
    hard_gates: list[dict[str, Any]] = []
    tail_rows: list[dict[str, Any]] = []
    for pair, (baseline_stage, candidate_stage) in PAIR_STAGES.items():
        pair_data: dict[str, Any] = {
            "baseline_stage": baseline_stage,
            "candidate_stage": candidate_stage,
            "corpora": {},
        }
        for corpus in plan["corpora"]:
            cid = corpus["id"]
            pair_data["corpora"][cid] = {}
            for phase in PHASES:
                baseline_row = rows[(baseline_stage, "native", cid, phase)]
                candidate_row = rows[(candidate_stage, "native", cid, phase)]
                native = compare_elapsed(candidate_row["elapsed_stats"],
                                         baseline_row["elapsed_stats"])
                pair_data["corpora"][cid][phase] = {"native": native}
                for metric in ("p50", "mean"):
                    if phase == "edit":
                        passed = (native[metric]["improvement_percent"]
                                  >= plan["thresholds"]["edit_improvement_percent"])
                        hard_gates.append({
                            "name": f"{pair}/{cid}/edit/{metric}/improvement",
                            "pass": passed,
                            "threshold_percent": plan["thresholds"]["edit_improvement_percent"],
                            "observed_improvement_percent": native[metric]["improvement_percent"],
                        })
                    else:
                        passed = (native[metric]["delta_percent"]
                                  <= plan["thresholds"]["lifecycle_regression_percent"])
                        hard_gates.append({
                            "name": f"{pair}/{cid}/lifecycle/{metric}/nonregression",
                            "pass": passed,
                            "threshold_percent": plan["thresholds"]["lifecycle_regression_percent"],
                            "observed_delta_percent": native[metric]["delta_percent"],
                        })
                for metric in ("p95", "p99"):
                    tail_rows.append({
                        "pair": pair,
                        "corpus_id": cid,
                        "phase": phase,
                        "metric": metric,
                        **native[metric],
                        "flag_over_5_percent": native[metric]["delta_percent"]
                        > plan["thresholds"]["tail_flag_percent"],
                    })
                if pair in {"pair-1", "pair-2"}:
                    baseline_alloc = rows[(baseline_stage, "allocator", cid, phase)]["allocation"]
                    candidate_alloc = rows[(candidate_stage, "allocator", cid, phase)]["allocation"]
                    allocation = compare_allocation(candidate_alloc, baseline_alloc)
                    pair_data["corpora"][cid][phase]["allocator"] = {
                        "allocation": allocation,
                        "baseline_diagnostics": baseline_alloc,
                        "candidate_diagnostics": candidate_alloc,
                    }
                    for label, metrics in allocation.items():
                        for metric in ("p50", "mean"):
                            comparison_row = metrics[metric]
                            observed = comparison_row["delta_percent"]
                            baseline_zero = comparison_row["baseline_zero"]
                            passed = allocation_gate_passes(
                                comparison_row, label, plan["thresholds"]
                            )
                            threshold = (
                                plan["thresholds"]["allocation_net_live_regression_percent"]
                                if label == "net_live"
                                else plan["thresholds"]["allocation_regression_percent"]
                            )
                            source = metrics.get("source_field", metrics.get("source_fields"))
                            hard_gates.append({
                                "name": f"{pair}/{cid}/{phase}/allocator/{label}/{metric}/nonregression",
                                "pass": passed,
                                "threshold_percent": threshold,
                                "observed_delta_percent": observed,
                                "observed_absolute_delta": metrics[metric]["absolute_delta"],
                                "baseline_zero": baseline_zero,
                                "baseline_equality": metrics[metric]["baseline_equality"],
                                "rule": ("candidate equals baseline when baseline is zero; otherwise candidate <= baseline"
                                          if label == "net_live" else
                                          "candidate equals baseline when baseline is zero; otherwise percentage regression <= threshold"),
                                "source_field": source,
                            })
                        for metric in ("p95", "p99"):
                            observed = metrics[metric]["delta_percent"]
                            tail_rows.append({
                                "lane": "allocator",
                                "pair": pair,
                                "corpus_id": cid,
                                "phase": phase,
                                "metric": f"{label}/{metric}",
                                **metrics[metric],
                                "source_field": metrics.get("source_field", metrics.get("source_fields")),
                                "flag_over_5_percent": (
                                    observed is not None
                                    and observed > plan["thresholds"]["tail_flag_percent"]
                                ),
                            })
        paired[pair] = pair_data
    return paired, hard_gates, tail_rows


def build_output(plan: dict[str, Any], baseline: dict[str, str],
                 candidate: dict[str, str], changed: list[str],
                 constraints_sha: str, rows: dict[tuple[str, str, str, str], dict[str, Any]],
                 builds: dict[str, Any], plan_sha: str, script_sha: str,
                 final_source: dict[str, str] | None,
                 disposition: dict[str, Any] | None) -> dict[str, Any]:
    parity_digests: dict[str, str] = {}
    parity_failures: list[dict[str, Any]] = []
    witnesses: dict[str, list[dict[str, Any]]] = {}
    for corpus in plan["corpora"]:
        for phase in PHASES:
            key = f"{corpus['id']}/{phase}"
            reference = rows[("baseline-A1", "native", corpus["id"], phase)]["normalized"]
            parity_digests[key] = digest(reference)
            witnesses[key] = []
            for row in rows.values():
                if row["corpus"]["id"] != corpus["id"] or row["phase"] != phase:
                    continue
                if row["normalized"] != reference:
                    parity_failures.append({
                        "key": key, "stage": row["stage"], "lane": row["lane"],
                    })
                witnesses[key].append({
                    "stage": row["stage"],
                    "lane": row["lane"],
                    "report": f"{row['name']}.json",
                    "report_sha256": sha(HERE / f"{row['name']}.json"),
                    "normalized_sha256": digest(row["normalized"]),
                    "published_sha256": row["ordinary"]["corpus"]["published_sha256"],
                })
    require(not parity_failures, f"deterministic output parity failed: {parity_failures}")
    paired, hard_gates, tail_rows = build_pair_output(rows, plan)
    primary_accepted = bool(hard_gates) and all(item["pass"] for item in hard_gates)
    read_controls = None
    if disposition is not None:
        read_controls = validate_read_controls(plan, disposition, baseline, candidate)
        if disposition["final_source"] == "baseline":
            require(disposition["retained"] is False,
                    "baseline final disposition requires an explicit rejected decision")
            if primary_accepted:
                require(read_controls["accepted"] is False
                        and read_controls["failed_hard_gate_count"] > 0,
                        "accepted primary gates require a failed read-control gate for baseline restoration")
        else:
            require(disposition["retained"] is True and primary_accepted
                    and read_controls["accepted"] is True,
                    "candidate final disposition requires an accepted decision")
    accepted = primary_accepted and (
        read_controls is None or read_controls["accepted"]
    )
    if disposition is not None:
        if disposition["final_source"] == "baseline":
            require(not accepted,
                    "baseline final disposition requires the aggregate decision to reject")
        else:
            require(accepted,
                    "candidate final disposition requires the aggregate decision to accept")
    repeat_rows, repeat_flags = repeat_diagnostics(rows, plan)
    process_traces = {
        "/".join((stage, lane, cid, phase)): row["process_trace"]
        for (stage, lane, cid, phase), row in sorted(rows.items())
    }
    raw_statistics = {
        "/".join((stage, lane, cid, phase)): {
            "elapsed_ns": row["elapsed_stats"],
            "allocation": row["allocation"],
        }
        for (stage, lane, cid, phase), row in sorted(rows.items())
    }
    stable_builds = copy.deepcopy(builds)
    for build in stable_builds.values():
        # The live binary may later be replaced by an exact cleanup witness.
        # Keep that observation out of the replayed decision bytes while still
        # validating it on every analysis pass.
        build["binary_custody"] = "validated-live-or-exact-cleanup-witness"
    return {
        "schema_version": 1,
        "packet": plan["packet"],
        "revision": plan["revision"],
        "performance_claim": "paired pilot decision only; no hardware, RSS, cold-cache, throughput, or scaling claim",
        "scope": {
            "corpora": [item["id"] for item in plan["corpora"]],
            "phases": list(PHASES),
            "stages": list(STAGES),
            "native_stages": len(STAGES),
            "allocator_stages": len(ALLOCATOR_STAGES),
            "children": len(rows),
            "native_children": len(STAGES) * len(plan["corpora"]) * len(PHASES),
            "allocator_children": len(ALLOCATOR_STAGES) * len(plan["corpora"]) * len(PHASES),
            "native_samples": plan["native"]["samples"],
            "allocator_samples": plan["allocator"]["samples"],
        },
        "raw_sample_statistics": raw_statistics,
        "process_probe_diagnostics": {
            "rows": process_traces,
            "latency_gate_used": False,
            "controls_subtracted": False,
            "note": "0717 process-probe/new-trace fields, when present, are separate diagnostics and are excluded only from normalized parity envelopes.",
        },
        "normalized_output_parity": {
            "verified": True,
            "reference": "baseline-A1/native",
            "keys": parity_digests,
            "stable_output_witnesses": witnesses,
            "excluded_fields": [
                "result.elapsed_ns",
                "result.operation_metrics",
                "result.source.ordinary_save.published_sha256",
                "result.source.ordinary_save.edit_outcome_sha256",
                "result.source.ordinary_save.process_probe",
            ],
            "note": "Scalar publication hashes, decoded manifests, phase identity, edit outcomes, and all deterministic non-timing evidence remain compared.",
        },
        "paired_comparisons": paired,
        "review_flags": {
            "tail_comparisons": tail_rows,
            "native_tail_regressions_over_5_percent": [
                item for item in tail_rows
                if item.get("lane", "native") == "native" and item["flag_over_5_percent"]
            ],
            "allocator_tail_regressions_over_5_percent": [
                item for item in tail_rows
                if item.get("lane") == "allocator" and item["flag_over_5_percent"]
            ],
            "within_variant_repeat_diagnostics": repeat_rows,
            "within_variant_repeat_drift_over_5_percent": repeat_flags,
            "note": "Review flags remain visible and never exclude samples, stages, pairs, or corpora from hard gates.",
        },
        "decision": {
            "hard_gates": hard_gates,
            "primary_hard_gates_pass": primary_accepted,
            "read_controls_pass": (None if read_controls is None else read_controls["accepted"]),
            "all_hard_gates_pass": accepted,
            "deterministic_output_parity_pass": True,
            "accepted": accepted,
            "acceptance_rule": "Accept only when normalized output parity, every per-pair primary gate, and the independent read-control decision pass when finalized.",
            "pooled_gates_used": False,
        },
        "verification": {
            "current_source_matches_source_candidate": (final_source is None or final_source == candidate),
            "source_delta_allowlist_verified": True,
            "source_delta_paths": changed,
            "child_receipts_verified": len(rows),
            "strict_0709_validators_used": [
                "check_report_metadata", "validate_elapsed",
                "validate_operation_metrics", "validate_ordinary_save",
            ],
            "0713_normalized_parity_used": True,
            "0717_process_probe_schema_checked_when_present": True,
            "all_samples_retained": True,
            "all_pairs_gated_independently": True,
            "fresh_child_per_corpus_phase": True,
            "allocation_request_count_requested_bytes_and_peak_above_start_gated": True,
            "allocation_net_live_no_increase_gated": True,
            "allocation_zero_baseline_requires_exact_equality": True,
            "allocation_net_live_and_peak_above_start_are_not_summed": True,
            "binary_live_or_cleanup_witness_verified": True,
            "source_maps_retained_once": True,
            "final_source_witness_verified": final_source is not None,
            "read_controls_analysis_verified": read_controls is not None,
        },
        "source": {
            "baseline_manifest": "source-baseline.json",
            "baseline_sha256": sha(HERE / "source-baseline.json"),
            "baseline_entry_count": len(baseline),
            "candidate_manifest": "source-candidate.json",
            "candidate_sha256": sha(HERE / "source-candidate.json"),
            "candidate_entry_count": len(candidate),
            "changed_paths": changed,
            "allowlist": plan["source_delta_allowlist"],
            "final_manifest": "source-final.json" if final_source is not None else None,
            "final_sha256": (sha(HERE / "source-final.json")
                             if final_source is not None else None),
            "final_entry_count": (len(final_source) if final_source is not None else None),
        },
        "builds": stable_builds,
        "final_disposition": disposition,
        "read_controls": read_controls,
        "freeze_bindings": {
            "plan_sha256": plan_sha,
            "capture_script_sha256": script_sha,
            "analyzer_sha256": sha(Path(__file__).resolve()),
            "analysis_script_sha256": sha(Path(__file__).resolve()),
            "constraints_sha256": constraints_sha,
        },
        "limits": plan["claims"]["limitations"],
    }


def analyze() -> dict[str, Any]:
    plan = CAPTURE.load_plan()
    constraints_sha = validate_constraints()
    baseline, candidate, changed, final_source, disposition = validate_sources(plan)
    builds: dict[str, Any] = {}
    for source_label in ("baseline", "candidate"):
        for lane in LANES:
            builds[f"{source_label}/{lane}"] = validate_builds(source_label, lane)
    plan_sha = sha(HERE / "plan.json")
    script_sha = sha(HERE / "pilot.py")
    jobs = validate_inventory(plan)
    rows: dict[tuple[str, str, str, str], dict[str, Any]] = {}
    for job in jobs:
        metadata = CAPTURE.stage_metadata(plan, job["stage"])
        build = builds[f"{metadata['build']}/{job['lane']}"]
        row = validate_child(plan, job, build, baseline, candidate,
                             plan_sha, script_sha, constraints_sha)
        key = (job["stage"], job["lane"], job["corpus"]["id"], job["phase"])
        require(key not in rows, f"duplicate child identity: {key}")
        rows[key] = row
    expected_count = len(STAGES) * 4 + len(ALLOCATOR_STAGES) * 4
    require(len(rows) == expected_count,
            f"expected {expected_count} isolated children, found {len(rows)}")
    return build_output(plan, baseline, candidate, changed, constraints_sha,
                        rows, builds, plan_sha, script_sha, final_source,
                        disposition)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", default=str(HERE / "analysis.json"))
    parser.add_argument("--check", action="store_true",
                        help="recompute and compare an existing analysis artifact")
    args = parser.parse_args()
    try:
        value = analyze()
        encoded = (json.dumps(value, indent=2) + "\n").encode()
        output = Path(args.output)
        if args.check:
            require(output.is_file() and not output.is_symlink(),
                    f"missing analysis output: {output}")
            require(output.read_bytes() == encoded,
                    f"analysis output is stale: {output}")
            print(f"verified {output}")
        else:
            require(not output.exists() and not output.is_symlink(),
                    f"refusing to replace {output}")
            output.write_bytes(encoded)
            print(f"verified {len(child_jobs(CAPTURE.load_plan()))} job identities; wrote {output}")
    except (AssertionError, KeyError, OSError, RuntimeError, ValueError) as error:
        print(f"analysis failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
