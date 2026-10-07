#!/usr/bin/env python3
"""Fail-closed offline reducer for the sealed 0959 v8 final-plan packet.

This program only reads an already sealed native packet and an already sealed
allocator packet.  It starts no workload and performs no build.  It validates
the v4 schedule, terminal-first child custody, artifact descriptors, report
oracles, and the plan binding before deriving paired summaries.  The native
bootstrap unit is the six block geometric means.  The allocator bootstrap unit
is the three paired repeats.  Both use 10,000 deterministic resamples with
seed 959.

A plan may be checked prospectively with ``--check-plan``.  Reduction requires
an authorized plan and complete lane packets; pending bindings are never
silently treated as a completed comparison.  This final-plan adapter preserves
the v2 report and gate logic, while accepting the final plan's equivalent
``candidate_manifest``/``candidate_freeze`` source binding names.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
from pathlib import Path
import random
import statistics
import sys
from typing import Any, Iterable

ROOT = Path("/home/zhuhe/litchi-performance-0959").resolve()
PLAN_SCHEMA = "litchi.performance.0959.inline-number-comparison-plan.v4"
REDUCER_SCHEMA = "litchi.performance.0959.reducer.v3.final-plan-adapter"
RUNNER_SCHEMA = "litchi.performance.0959.capture.v4"
CPU_MASK = [12, 13, 14, 15]
DEADLINE = 900
NATIVE_COUNT = 72
ALLOCATOR_COUNT = 18
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_SEED = 959
SAFETY = 1.05
IMPROVEMENT = 0.95
CI_UPPER_FOR_IMPROVEMENT = 1.0
REPORT_LIMIT = 8 * 1024 * 1024
JSON_LIMIT = 64 * 1024 * 1024
MAX_SINK_WRITE = 65_536
NATIVE_PHASES = (
    "ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns",
)
NATIVE_PRIMARY = ("edit_commit_ns", "total_ns")
NATIVE_SECONDARY = ("ingress_ns", "publication_ns", "commit_drop_ns")
ALLOC_PHASES = (
    "setup", "ingress", "edit_commit", "publication", "commit_drop",
    "worker_drain", "core_combined", "core_combined_with_drain", "postdrop",
)
ALLOC_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
ALLOC_METRICS = (
    "allocated_bytes", "allocation_calls", "region_peak_live_bytes",
    "core_combined_with_drain.allocated_bytes", "max_rss_kib",
)
TIME_METRICS = (
    "max_rss_kib", "user_seconds", "system_seconds", "major_page_faults",
    "minor_page_faults", "voluntary_context_switches", "involuntary_context_switches",
)


class ReductionError(ValueError):
    """The sealed packet is not admissible under the frozen v4 contract."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReductionError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_json(path: Path, *, limit: int = JSON_LIMIT) -> Any:
    require(not path.is_symlink() and path.is_file(), f"regular JSON file required: {path}")
    require(path.stat().st_size <= limit, f"JSON file exceeds limit: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ReductionError(f"invalid JSON {path}: {exc}") from exc


def read_json_bytes(path: Path, *, limit: int = JSON_LIMIT) -> tuple[Any, bytes]:
    require(not path.is_symlink() and path.is_file(), f"regular JSON file required: {path}")
    raw = path.read_bytes()
    require(len(raw) <= limit, f"JSON file exceeds limit: {path}")
    try:
        return json.loads(raw.decode("utf-8")), raw
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ReductionError(f"invalid JSON {path}: {exc}") from exc


def write_new(path: Path, value: Any) -> None:
    path = path.resolve()
    require(not path.exists(), f"refusing to overwrite output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    raw = (json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n").encode()
    with path.open("xb") as stream:
        stream.write(raw)
        stream.flush()
        os.fsync(stream.fileno())


def descriptor(path: Path) -> dict[str, Any]:
    require(not path.is_symlink() and path.is_file(), f"regular artifact required: {path}")
    mode = path.stat().st_mode & 0o777
    return {
        "path": str(path.resolve()),
        "bytes": path.stat().st_size,
        "mode": mode,
        "executable": bool(mode & 0o111),
        "sha256": sha256(path),
    }


def check_descriptor(value: Any, expected: Path, label: str, *, executable: bool = False) -> dict[str, Any]:
    """Check a descriptor against the packet's expected physical artifact.

    Capture descriptors contain absolute paths.  The physical packet can be
    copied for replay, so the declared path is treated as provenance while the
    known packet-relative location is the file that is read and verified.
    """
    require(isinstance(value, dict), f"{label}: descriptor missing")
    require(isinstance(value.get("path"), str), f"{label}.path missing")
    actual = descriptor(expected)
    for key in ("bytes", "mode", "sha256"):
        require(value.get(key) == actual[key], f"{label}.{key} differs from physical artifact")
    if "executable" in value:
        require(value.get("executable") == actual["executable"], f"{label}.executable differs from physical artifact")
    if executable:
        require(value.get("executable") is True and actual["executable"], f"{label}: executable binding missing")
    return actual


def external_descriptor(value: Any, label: str, *, executable: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and Path(raw_path).is_absolute(), f"{label}.path must be absolute")
    path = Path(raw_path)
    actual = descriptor(path)
    for key in ("path", "bytes", "sha256"):
        require(value.get(key) == actual[key], f"{label}.{key} differs from bound artifact")
    if executable:
        require(value.get("executable") is True and actual["executable"], f"{label}: executable binding missing")
    return actual


def integer(value: Any, label: str, *, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{label}: expected integer >= {minimum}")
    return value


def number(value: Any, label: str, *, positive: bool = False) -> float:
    require(type(value) in (int, float) and not isinstance(value, bool), f"{label}: expected number")
    result = float(value)
    require(math.isfinite(result), f"{label}: expected finite number")
    require(result > 0 if positive else result >= 0, f"{label}: expected {'positive' if positive else 'nonnegative'} number")
    return result


def nearest_rank(values: Iterable[float], percentile: int) -> float:
    ordered = sorted(float(x) for x in values)
    require(ordered, "cannot calculate a percentile of an empty sequence")
    if percentile == 50:
        return float(statistics.median(ordered))
    index = max(1, math.ceil(percentile * len(ordered) / 100)) - 1
    return ordered[min(index, len(ordered) - 1)]


def nearest_rank_fraction(values: Iterable[float], fraction: float) -> float:
    ordered = sorted(float(x) for x in values)
    require(ordered, "cannot calculate a fractional percentile of an empty sequence")
    require(0.0 < fraction <= 1.0, "fractional percentile is outside (0,1]")
    index = max(1, math.ceil(fraction * len(ordered))) - 1
    return ordered[min(index, len(ordered) - 1)]


def stats(values: Iterable[float]) -> dict[str, float]:
    values = [number(value, "statistic input") for value in values]
    require(values, "empty metric vector")
    return {
        "mean": float(statistics.fmean(values)),
        "p50": nearest_rank(values, 50),
        "p95": nearest_rank(values, 95),
        "p99": nearest_rank(values, 99),
        "min": float(min(values)),
        "max": float(max(values)),
        "count": len(values),
    }


def geometric_mean(values: Iterable[float]) -> float:
    values = [number(value, "geometric-mean input") for value in values]
    require(values, "empty geometric-mean vector")
    # A zero numerator with a positive baseline is a valid zero ratio.  Only
    # a zero denominator is undefined under the protocol; do not turn a valid
    # zero result into a bootstrap failure.
    if any(value == 0 for value in values):
        return 0.0
    return math.exp(sum(math.log(value) for value in values) / len(values))


def ratio(numerator: float, denominator: float, label: str) -> dict[str, Any]:
    numerator = number(numerator, f"{label}.numerator")
    denominator = number(denominator, f"{label}.denominator")
    if denominator == 0:
        return {"status": "undefined", "reason": "zero_denominator", "numerator": numerator, "denominator": denominator}
    value = numerator / denominator
    require(math.isfinite(value) and value >= 0, f"{label}: ratio is not finite")
    return {"status": "defined", "value": value, "numerator": numerator, "denominator": denominator}


def ratio_value(value: dict[str, Any], label: str) -> float | None:
    require(isinstance(value, dict), f"{label}: malformed ratio")
    if value.get("status") == "undefined":
        require(value.get("reason") == "zero_denominator", f"{label}: unknown undefined reason")
        return None
    require(value.get("status") == "defined", f"{label}: invalid ratio status")
    return number(value.get("value"), f"{label}.value")


def bootstrap(values: list[float]) -> dict[str, Any]:
    """Bootstrap geometric means over the frozen block/repeat units."""
    require(values, "empty bootstrap vector")
    if any(value < 0 or not math.isfinite(value) for value in values):
        return {
            "status": "undefined", "reason": "negative_or_nonfinite_ratio",
            "unit_count": len(values), "resamples": BOOTSTRAP_RESAMPLES,
            "seed": BOOTSTRAP_SEED,
        }
    point = geometric_mean(values)
    rng = random.Random(BOOTSTRAP_SEED)
    sampled: list[float] = []
    n = len(values)
    for _ in range(BOOTSTRAP_RESAMPLES):
        sampled.append(geometric_mean(values[rng.randrange(n)] for _ in range(n)))
    sampled.sort()
    return {
        "status": "defined",
        "point": point,
        "ci_lower": nearest_rank_fraction(sampled, 0.025),
        "ci_upper": nearest_rank_fraction(sampled, 0.975),
        "ci_method": "nearest-rank two-sided percentile bootstrap; 2.5%/97.5% endpoints",
        "unit": "six native block geometric means or three allocator paired repeats",
        "unit_count": n,
        "resamples": BOOTSTRAP_RESAMPLES,
        "seed": BOOTSTRAP_SEED,
    }


def check_plan_shape(plan: dict[str, Any], *, require_authorized: bool = False) -> None:
    require(plan.get("schema") == PLAN_SCHEMA, "plan schema is not v4")
    if require_authorized:
        require(plan.get("status") == "authorized", "plan is not authorized")
    cases = plan.get("cases")
    require(isinstance(cases, list) and len(cases) == 3, "plan must contain three cases")
    case_ids = [case.get("case_id") for case in cases]
    require(case_ids == ["medium-one-percent", "dense-sparse-one-percent", "medium-one-edit"], "case order changed")
    comparison = plan.get("comparison")
    require(isinstance(comparison, dict), "comparison section missing")
    expected_blocks = [
        {"id": 0, "order": ["A", "B", "B", "A"]},
        {"id": 1, "order": ["B", "A", "A", "B"]},
        {"id": 2, "order": ["A", "B", "B", "A"]},
        {"id": 3, "order": ["B", "A", "A", "B"]},
        {"id": 4, "order": ["A", "B", "B", "A"]},
        {"id": 5, "order": ["B", "A", "A", "B"]},
    ]
    expected_allocator = [
        {"id": 0, "order": ["A", "B"]},
        {"id": 1, "order": ["B", "A"]},
        {"id": 2, "order": ["A", "B"]},
    ]
    require(comparison.get("blocks") == expected_blocks, "native ABBA/BAAB schedule changed")
    require(comparison.get("allocator_blocks") == expected_allocator, "allocator schedule changed")
    require(comparison.get("block_count") == 6 and comparison.get("pairs_per_case_native") == 12, "native pair counts changed")
    require(comparison.get("pairs_per_case_allocator") == 3 and comparison.get("allocator_repeats_per_role") == 3, "allocator pair counts changed")
    thresholds = plan.get("gates", {}).get("review_thresholds", {})
    require(thresholds.get("bootstrap_resamples") == BOOTSTRAP_RESAMPLES, "bootstrap resample count changed")
    require(thresholds.get("bootstrap_seed") == BOOTSTRAP_SEED, "bootstrap seed changed")
    require(thresholds.get("ratio_direction") == "candidate_over_baseline; lower is better", "ratio direction changed")
    require(float(thresholds.get("safety_ratio_upper")) == SAFETY, "safety threshold changed")
    require(float(thresholds.get("improvement_point_ratio")) == IMPROVEMENT, "improvement threshold changed")
    lanes = plan.get("lanes")
    require(isinstance(lanes, dict) and set(lanes) >= {"native", "allocator"}, "lanes missing")
    for lane, count, warmup, samples, schema in (
        ("native", NATIVE_COUNT, 20, 200, "litchi.xlsx-managed-compression-benchmark.v1"),
        ("allocator", ALLOCATOR_COUNT, 2, 10, "litchi.xlsx-managed-compression-allocation.v1"),
    ):
        spec = lanes[lane]
        require(spec.get("child_count") == count, f"{lane} child count changed")
        require(spec.get("warmup_iterations") == warmup and spec.get("measured_samples") == samples, f"{lane} sample schedule changed")
        require(spec.get("report_schema") == schema, f"{lane} report schema changed")
        execution = spec.get("execution")
        require(isinstance(execution, dict), f"{lane} execution missing")
        require(execution.get("cpu_mask") == CPU_MASK, f"{lane} CPU mask changed")
        require(execution.get("deadline_seconds") == DEADLINE, f"{lane} deadline changed")
        require(execution.get("max_parallel_children") == 1, f"{lane} children are not serial")
        require(execution.get("terminal_before_parse") is True, f"{lane} terminal-first custody disabled")
        require(execution.get("retry_or_replacement") is False, f"{lane} retry policy changed")
        rows = spec.get("children")
        require(isinstance(rows, list) and len(rows) == count, f"{lane} child schedule missing")
        expected = expected_rows(plan, lane)
        require(rows == expected, f"{lane} children differ from plan schedule")
    report_contract = plan.get("report_contract")
    require(isinstance(report_contract, dict), "report contract missing")
    require(report_contract.get("native", {}).get("raw_sample_count") == 200, "native report sample contract changed")
    require(report_contract.get("allocator", {}).get("raw_sample_count") == 10, "allocator report sample contract changed")


def expected_rows(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    cases = plan["cases"]
    blocks = plan["comparison"]["blocks" if lane == "native" else "allocator_blocks"]
    spec = plan["lanes"][lane]
    rows: list[dict[str, Any]] = []
    for case in cases:
        for block in blocks:
            for position, role in enumerate(block["order"]):
                rows.append({
                    "binary_role": "baseline" if role == "A" else "candidate",
                    "block": block["id"],
                    "case_id": case["case_id"],
                    "child_id": f"{lane}__{case['case_id']}__b{block['id']:02d}__p{position}",
                    "comparison_id": "baseline-vs-candidate",
                    "cpu_mask": CPU_MASK,
                    "lane": lane,
                    "measured_samples": spec["measured_samples"],
                    "position": position,
                    "role": role,
                    "shape": case["shape"],
                    "warmup_iterations": spec["warmup_iterations"],
                    "workers": 1,
                    "workload": case["workload"],
                })
    return rows


def validate_plan_bindings(plan: dict[str, Any], plan_path: Path, helper_path: Path) -> None:
    """Validate bindings which are only meaningful once a packet is complete."""
    require(plan.get("status") == "authorized", "reduction requires an authorized plan")
    source = plan.get("source", {})
    # Final plan v8 retains the same bound files under the shorter names.  The
    # v2 reducer used the earlier names; accept only this explicit, equivalent
    # descriptor mapping and continue to validate the descriptor bytes below.
    source_aliases = {
        "candidate_source_manifest": "candidate_manifest",
        "candidate_source_freeze": "candidate_freeze",
    }
    for legacy_name, final_name in source_aliases.items():
        value = source.get(legacy_name)
        if value is None:
            value = source.get(final_name)
        require(isinstance(value, dict) and value.get("status") != "pending_root_binding", f"source.{legacy_name} is pending")
    for lane in ("native", "allocator"):
        for role in ("baseline_binary", "candidate_binary"):
            value = plan["lanes"][lane].get(role)
            require(isinstance(value, dict) and value.get("status") != "pending_root_binding", f"{lane}.{role} is pending")
            external_descriptor(value, f"plan.lanes.{lane}.{role}", executable=True)
        for receipt in ("baseline_receipt", "candidate_receipt"):
            value = plan["lanes"][lane].get(receipt)
            if value is not None:
                require(isinstance(value, dict), f"{lane}.{receipt} malformed")
    helpers = plan.get("helpers")
    require(isinstance(helpers, dict), "plan helper bindings missing")
    reducer_binding = helpers.get("reducer")
    require(isinstance(reducer_binding, dict) and reducer_binding.get("status") != "pending_root_binding", "plan reducer binding is pending")
    actual = descriptor(helper_path)
    # The authorized plan intentionally binds the frozen v2 reducer.  This
    # versioned adapter is the post-capture compatibility correction, so the
    # plan binding must still match the immutable v2 descriptor exactly.
    frozen_v2 = ROOT / "tools" / "reduce_0959_v2.py"
    frozen_v2_descriptor = descriptor(frozen_v2)
    require(
        all(reducer_binding.get(key) == frozen_v2_descriptor[key] for key in ("path", "bytes", "sha256")),
        "plan reducer binding differs from frozen v2 helper",
    )
    require(actual["path"] != frozen_v2_descriptor["path"] and actual["sha256"] != frozen_v2_descriptor["sha256"], "adapter must be a distinct versioned helper")
    for key in ("capture", "controls"):
        value = helpers.get(key)
        require(isinstance(value, dict) and value.get("status") != "pending_root_binding", f"plan helper {key} is pending")
    # Source manifests and build receipts are external descriptors in a final plan.
    for key in ("baseline_manifest", "baseline_receipt", "baseline_runtime_manifest", "candidate_manifest", "candidate_freeze"):
        value = source.get(key)
        if isinstance(value, dict) and value.get("status") != "pending_root_binding":
            if "path" in value:
                external_descriptor(value, f"plan.source.{key}")
    require(isinstance(plan.get("host"), dict) and plan["host"].get("status") != "pending_root_binding", "host binding is pending")
    require(isinstance(plan.get("root_go"), dict) and plan["root_go"].get("status") != "pending_root_binding", "root GO is pending")


def expected_case(plan: dict[str, Any], case_id: str) -> dict[str, Any]:
    for case in plan["cases"]:
        if case.get("case_id") == case_id:
            return case
    raise ReductionError(f"unknown case {case_id}")


def validate_common_report(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any], lane: str) -> dict[str, Any]:
    spec = plan["lanes"][lane]
    require(report.get("schema") == spec["report_schema"], f"{row['child_id']}: report schema changed")
    require(report.get("status") == "completed", f"{row['child_id']}: report not completed")
    case = report.get("case")
    require(isinstance(case, dict), f"{row['child_id']}: case missing")
    # Native serializes the common case identity and sample schedule at the
    # report root.  The allocator serializer keeps those fields under case;
    # the retained 0957 allocator reports are the compatibility fixture for
    # this distinction.  Never infer absent top-level allocator fields.
    if lane == "native":
        require(report.get("shape") == row["shape"] and report.get("workload") == row["workload"], f"{row['child_id']}: native report case differs")
        require(report.get("workers") == 1, f"{row['child_id']}: native report workers changed")
        require(report.get("warmup_iterations") == spec["warmup_iterations"] and report.get("measured_samples") == spec["measured_samples"], f"{row['child_id']}: native report sample schedule changed")
    expected = expected_case(plan, row["case_id"])
    expected_case_fields = {
        "id": f"{expected['shape']}-{expected['workload']}-workers1",
        "shape": expected["shape"],
        "workload": expected["workload"],
        "workers": 1,
        "update_count": expected["update_count"],
        "source_bytes": expected["source_bytes"],
        "source_sha256": expected["source_sha256"],
        "output_bytes": expected["output_bytes"],
        "output_sha256": expected["output_sha256"],
        "semantic_sha256": expected["semantic_sha256"],
        "expected_cpu_tasks": expected["expected_cpu_tasks"],
        "admission_expected_worker_permits": 0,
        "admission_observed_worker_permits": 0,
        "admission_cpu_tasks_used": expected["expected_cpu_tasks"],
        "preflight_actual_commit_inverse_verified": True,
        "context_limits": expected["context_limits"],
        "sink_capacity_bytes": expected["context_limits"]["output_bytes"],
        "sink_max_write_bytes": MAX_SINK_WRITE,
        "warmup_iterations": spec["warmup_iterations"],
        "measured_samples": spec["measured_samples"],
    }
    for key, value in expected_case_fields.items():
        require(case.get(key) == value, f"{row['child_id']}: case.{key} oracle changed")
    require(case.get("id") == expected_case_fields["id"], f"{row['child_id']}: case id changed")
    require(case.get("sink_max_write_bytes") == MAX_SINK_WRITE, f"{row['child_id']}: sink write limit changed")
    require(case.get("admission_observed_worker_permits") == 0 and case.get("admission_expected_worker_permits") == 0, f"{row['child_id']}: worker permit oracle changed")
    require(case.get("preflight_actual_commit_inverse_verified") is True, f"{row['child_id']}: inverse admission failed")
    return case


def validate_native_report(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    case = validate_common_report(report, row, plan, "native")
    scope = {
        "timing_scope": "with fresh finite managed context prepared outside the clock, the timer begins at SourceBackedEditor construction and covers edit/commit, source-backed publication, returned snapshot drop, and commit/package worker-pool teardown; precomputed source/cache/context setup and bounded sink allocation are outside",
        "phase_scope": "ingress_ns is the direct SourceBackedEditor constructor; edit_commit_ns is selector staging, cell sets, and MultiCommit::commit; publication_ns is publish_multi_commit_to_stream plus returned MultiSnapshot drop; commit_drop_ns is MultiCommit drop; total_ns starts at the constructor and ends after commit teardown",
        "setup_scope": "deterministic corpus and expected output construction, frozen source/output/semantic oracles, SourceCacheLimits and output ceiling, one untimed admission preflight, inverse patch, fresh OwnedSource clone, finite managed context, and sink capacity reservation are outside measured phases",
        "verification_scope": "exact output bytes, immutable fixture source hash, output/semantic hashes, workbook verification, sink bounds, CpuTasks/OutputBytes counters, and post-drop Memory/Workers/IoConcurrency/Objects checks run outside timers; the untimed admission preflight verifies an actual managed-commit inverse (preflight_actual_commit_inverse_verified=true), while each timed sample applies the precomputed inverse patch only as a fixture to its published bytes (fixture_inverse_restored_exact_source)",
        "performance_claim": "none; raw phase timings describe this explicit in-memory harness and do not claim OS thread counts, speedup, allocation behavior, or physical I/O",
    }
    for key, value in scope.items():
        require(report.get(key) == value, f"{row['child_id']}: report {key} scope changed")
    samples = case.get("samples")
    require(isinstance(samples, list) and len(samples) == 200, f"{row['child_id']}: native samples missing")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("sample") == index, f"{row['child_id']}: native sample index changed")
        values = [integer(sample.get(name), f"{row['child_id']}.sample[{index}].{name}") for name in NATIVE_PHASES]
        require(sum(values[:4]) <= values[-1], f"{row['child_id']}: phase sum exceeds total")
        for key in ("output_bytes", "output_bytes_used"):
            require(sample.get(key) == expected_case(plan, row["case_id"])["output_bytes"], f"{row['child_id']}: {key} oracle changed")
        require(sample.get("cpu_tasks_used") == expected_case(plan, row["case_id"])["expected_cpu_tasks"], f"{row['child_id']}: CpuTasks changed")
        require(integer(sample.get("sink_write_calls"), "sink_write_calls", minimum=1) >= 1, f"{row['child_id']}: sink calls missing")
        require(0 < integer(sample.get("sink_largest_write"), "sink_largest_write") <= MAX_SINK_WRITE, f"{row['child_id']}: sink bound changed")
        require(sample.get("output_sha256") == case["output_sha256"] and sample.get("semantic_sha256") == case["semantic_sha256"], f"{row['child_id']}: output oracle changed")
        require(sample.get("fixture_inverse_restored_exact_source") is True, f"{row['child_id']}: inverse restoration failed")
        resources = sample.get("resources_after_drop")
        require(isinstance(resources, dict) and all(resources.get(key) == 0 for key in ("memory", "workers", "io_concurrency", "objects")), f"{row['child_id']}: resources after drop not zero")
    return case


def validate_alloc_sample(sample: Any, label: str) -> dict[str, Any]:
    require(isinstance(sample, dict) and sample.get("status") == "measured", f"{label}: sample not measured")
    require(sample.get("scope") == "operation_global_system_allocator", f"{label}: allocator scope changed")
    for field in ALLOC_FIELDS:
        integer(sample.get(field), f"{label}.{field}")
    require(sample["failed_allocation_calls"] == 0, f"{label}: failed allocation recorded")
    require(sample["region_peak_live_bytes"] >= max(sample["live_bytes_before"], sample["live_bytes_after"]), f"{label}: region peak below endpoint")
    return sample


def combine_alloc(samples: list[dict[str, Any]], label: str) -> dict[str, Any]:
    require(samples, f"{label}: no phases")
    for left, right in zip(samples, samples[1:]):
        require(left["live_bytes_after"] == right["live_bytes_before"], f"{label}: non-adjacent live boundaries")
    result: dict[str, Any] = {
        "status": "measured",
        "scope": "operation_global_system_allocator",
        "live_bytes_before": samples[0]["live_bytes_before"],
        "live_bytes_after": samples[-1]["live_bytes_after"],
        "peak_live_bytes_before": samples[0]["peak_live_bytes_before"],
        "peak_live_bytes_after": samples[-1]["peak_live_bytes_after"],
        "region_peak_live_bytes": max(item["region_peak_live_bytes"] for item in samples),
    }
    for field in ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes"):
        result[field] = sum(item[field] for item in samples)
    return result


def validate_allocator_report(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    case = validate_common_report(report, row, plan, "allocator")
    expected_scope = {
        "allocator_identity": "CountingSystemAllocator(std::alloc::System)",
        "instrumentation_identity": "system_allocator_operation_scoped",
        "counter_revision": "serialized_region_peak_v3",
        "allocation_scope": "global system-allocator callbacks serialized by allocation_metrics; process-wide across threads; observer peaks are not RSS, unique heap, per-thread bytes, OS-thread counts, or timing",
        "thread_observation_scope": "calibration uses a separate measured allocator region; pre-sample /proc/self/status probes are outside operation regions; the post-core fixed-buffer drain is the fifth in-region segment; observed process thread counts do not establish admitted-worker/thread mapping or task overlap",
        "timing_scope": "none; this report contains allocator regions and diagnostic quiescence observations only",
        "performance_claim": "none; allocator counters do not claim latency, throughput, RSS, OS-thread width, worker/task overlap, speedup, or adoption; procfs Threads values are quiescence observations only",
        "peak_semantics": "observer-ordered process-live high-water evidence, not RSS, unique heap, per-thread bytes, or OS-thread width",
    }
    for key, value in expected_scope.items():
        require(report.get(key) == value, f"{row['child_id']}: allocator {key} scope changed")
    calibration = report.get("calibration")
    require(isinstance(calibration, dict), f"{row['child_id']}: calibration missing")
    for key, value in {"status": "pass", "polls": 100, "thread_count_first": 1, "thread_count_last": 1, "allocation_calls": 0, "deallocation_calls": 0, "reallocation_calls": 0, "failed_allocation_calls": 0, "no_allocator_callbacks": True, "poll_sleep_milliseconds": 1}.items():
        require(calibration.get(key) == value, f"{row['child_id']}: calibration {key} changed")
    rows = case.get("rows")
    require(isinstance(rows, list) and len(rows) == 10, f"{row['child_id']}: allocator rows missing")
    for index, item in enumerate(rows):
        require(isinstance(item, dict) and item.get("sample") == index, f"{row['child_id']}: allocator row index changed")
        expected = expected_case(plan, row["case_id"])
        for key, value in {
            "context_limits": expected["context_limits"], "workers_requested": 1,
            "expected_worker_permits": 0, "admitted_worker_permits": 0,
            "expected_cpu_tasks": expected["expected_cpu_tasks"], "cpu_tasks_used": expected["expected_cpu_tasks"],
            "output_bytes": expected["output_bytes"], "output_bytes_used": expected["output_bytes"],
            "source_sha256": expected["source_sha256"], "output_sha256": expected["output_sha256"],
            "semantic_sha256": expected["semantic_sha256"], "fixture_inverse_restored_exact_source": True,
        }.items():
            require(item.get(key) == value, f"{row['child_id']}: allocator row {index} {key} oracle changed")
        require(integer(item.get("sink_write_calls"), "sink_write_calls", minimum=1) >= 1, f"{row['child_id']}: sink calls missing")
        require(0 < integer(item.get("sink_largest_write"), "sink_largest_write") <= MAX_SINK_WRITE, f"{row['child_id']}: sink bound changed")
        resources = item.get("resources_after_core")
        require(isinstance(resources, dict) and all(resources.get(key) == 0 for key in ("memory", "workers", "io_concurrency", "objects")), f"{row['child_id']}: resources after core not zero")
        phases = [validate_alloc_sample(item.get(name), f"{row['child_id']}.row[{index}].{name}") for name in ALLOC_PHASES]
        core = combine_alloc(phases[1:5], f"{row['child_id']}.row[{index}].core")
        core_drain = combine_alloc(phases[1:6], f"{row['child_id']}.row[{index}].core_combined_with_drain")
        require(item["core_combined"] == core, f"{row['child_id']}: core conservation failed")
        require(item["core_combined_with_drain"] == core_drain, f"{row['child_id']}: drained core conservation failed")
        integer(item.get("max_setup_core_peak_above_setup_live_before"), "max_setup_core_peak_above_setup_live_before")
        integer(item.get("max_setup_drained_core_peak_above_setup_live_before"), "max_setup_drained_core_peak_above_setup_live_before")
    binary_identity = report.get("binary_identity")
    require(isinstance(binary_identity, dict), f"{row['child_id']}: binary identity missing")
    bound_binary = plan["lanes"]["allocator"]["baseline_binary" if row["role"] == "A" else "candidate_binary"]
    actual_binary = external_descriptor(bound_binary, f"allocator.{row['role']}.binary", executable=True)
    for key in ("path", "bytes", "sha256"):
        require(binary_identity.get(key) == actual_binary[key], f"{row['child_id']}: allocator binary identity {key} changed")
    return case


def parse_time_metrics(validation: dict[str, Any], child_id: str) -> dict[str, float | int]:
    value = validation.get("time_metrics")
    require(isinstance(value, dict), f"{child_id}: time_metrics missing")
    result: dict[str, float | int] = {}
    for key in TIME_METRICS:
        raw = value.get(key)
        if key in ("user_seconds", "system_seconds"):
            require(type(raw) in (int, float) and math.isfinite(float(raw)) and float(raw) >= 0, f"{child_id}: {key} invalid")
            result[key] = float(raw)
        else:
            result[key] = integer(raw, f"{child_id}.{key}")
    return result


def expected_packet_path(packet: Path, child_id: str, name: str) -> Path:
    return packet / "children" / child_id / name


def validate_lane_packet(packet: Path, plan_path: Path, plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    require(packet.is_dir() and not packet.is_symlink(), f"{lane} packet directory missing")
    lane_spec = plan["lanes"][lane]
    rows = expected_rows(plan, lane)
    start_path = packet / "capture-start.json"
    terminal_path = packet / "capture-terminal.json"
    completed_path = packet / "completed-packet.json"
    start = read_json(start_path)
    require(start.get("schema") == f"{RUNNER_SCHEMA}.capture-start", f"{lane}: capture-start schema changed")
    require(start.get("lane") == lane and start.get("children_expected") == len(rows) and start.get("serial") is True, f"{lane}: capture-start schedule changed")
    plan_desc = start.get("plan")
    actual_plan = descriptor(plan_path)
    require(isinstance(plan_desc, dict) and plan_desc.get("bytes") == actual_plan["bytes"] and plan_desc.get("sha256") == actual_plan["sha256"], f"{lane}: capture-start plan binding changed")
    runner_desc = start.get("runner")
    require(isinstance(runner_desc, dict), f"{lane}: runner descriptor missing")
    if isinstance(runner_desc.get("path"), str) and Path(runner_desc["path"]).is_file():
        external_descriptor(runner_desc, f"{lane}.capture-start.runner", executable=True)
    terminal = read_json(terminal_path)
    require(terminal.get("schema") == f"{RUNNER_SCHEMA}.capture-terminal", f"{lane}: capture-terminal schema changed")
    require(terminal.get("status") == "completed" and terminal.get("lane") == lane, f"{lane}: capture terminal is not completed")
    require(terminal.get("children_started") == len(rows) and terminal.get("children_expected") == len(rows), f"{lane}: capture child counts incomplete")
    require(terminal.get("terminal_before_parse") is True and terminal.get("retry_or_replacement") is False, f"{lane}: capture custody changed")
    check_descriptor(terminal.get("capture_start"), start_path, f"{lane}.capture-terminal.capture_start")
    completed = read_json(completed_path)
    require(completed.get("schema") == f"{RUNNER_SCHEMA}.completed-packet" and completed.get("status") == "completed", f"{lane}: completed-packet is not completed")
    completed_children = completed.get("children")
    require(isinstance(completed_children, list) and [item.get("child_id") for item in completed_children] == [row["child_id"] for row in rows], f"{lane}: completed child manifest differs")
    for item in completed_children:
        check_descriptor(item.get("validation"), packet / "children" / item["child_id"] / "validation.json", f"{lane}.completed-packet.{item['child_id']}.validation")
    check_descriptor(completed.get("capture_terminal"), terminal_path, f"{lane}.completed-packet.capture_terminal")
    result: list[dict[str, Any]] = []
    for expected in rows:
        child_id = expected["child_id"]
        child = packet / "children" / child_id
        require(child.is_dir() and not child.is_symlink(), f"{lane}/{child_id}: child directory missing")
        start_child = read_json(child / "start.json")
        require(start_child.get("schema") == f"{RUNNER_SCHEMA}.child-start" and start_child.get("child") == expected, f"{child_id}: start schedule differs")
        require(start_child.get("cwd") == lane_spec["execution"]["workload_cwd"], f"{child_id}: workload cwd changed")
        require(start_child.get("deadline_seconds") == DEADLINE, f"{child_id}: start deadline changed")
        expected_environment = dict(lane_spec["execution"]["environment"])
        approval_key = expected_environment.pop("approval_variable")
        expected_environment[approval_key] = "1"
        require(start_child.get("environment") == expected_environment, f"{child_id}: environment changed")
        binary = lane_spec["baseline_binary"] if expected["role"] == "A" else lane_spec["candidate_binary"]
        require(isinstance(binary, dict) and binary.get("status") != "pending_root_binding", f"{child_id}: binary binding is pending")
        expected_binary = external_descriptor(binary, f"{lane}.{expected['role']}.binary", executable=True)
        check_descriptor(start_child.get("binary"), Path(expected_binary["path"]), f"{child_id}.start.binary", executable=True)
        output_path = child / "output"
        time_path = child / "time-v.txt"
        expected_workload_argv = [
            "/usr/bin/taskset", "--cpu-list", "12-15", expected_binary["path"],
            "--shape", expected["shape"], "--workload", expected["workload"],
            "--workers", "1", "--warmup", str(expected["warmup_iterations"]),
            "--samples", str(expected["measured_samples"]), "--output", str(output_path),
        ]
        expected_wrapped_argv = [
            "/usr/bin/taskset", "--cpu-list", "12-15", "/usr/bin/time", "-v", "-o", str(time_path), "--",
            expected_binary["path"], "--shape", expected["shape"], "--workload", expected["workload"],
            "--workers", "1", "--warmup", str(expected["warmup_iterations"]),
            "--samples", str(expected["measured_samples"]), "--output", str(output_path),
        ]
        require(start_child.get("workload_argv") == expected_workload_argv, f"{child_id}: workload argv differs from plan/binary binding")
        require(start_child.get("argv") == expected_wrapped_argv, f"{child_id}: wrapped argv differs from plan/binary binding")
        launch = read_json(child / "launch.json")
        require(launch.get("schema") == f"{RUNNER_SCHEMA}.child-launch" and launch.get("observed_affinity") == CPU_MASK, f"{child_id}: observed affinity changed")
        terminal_child = read_json(child / "terminal.json")
        require(terminal_child.get("schema") == f"{RUNNER_SCHEMA}.child-terminal", f"{child_id}: terminal schema changed")
        require(terminal_child.get("status") == "completed" and terminal_child.get("return_code") == 0 and terminal_child.get("timed_out") is False and terminal_child.get("process_group_clear") is True, f"{child_id}: child did not complete cleanly")
        validation = read_json(child / "validation.json")
        require(validation.get("schema") == f"{RUNNER_SCHEMA}.child-validation" and validation.get("status") == "passed", f"{child_id}: validation did not pass")
        require(validation.get("child") == expected, f"{child_id}: validation schedule differs")
        check_descriptor(validation.get("terminal"), child / "terminal.json", f"{child_id}.validation.terminal")
        report_path = output_path / "report.json"
        check_descriptor(validation.get("report"), report_path, f"{child_id}.validation.report")
        check_descriptor(validation.get("time"), time_path, f"{child_id}.validation.time")
        time_metrics = parse_time_metrics(validation, child_id)
        post = read_json(child / "post-terminal-artifacts.json")
        require(post.get("schema") == f"{RUNNER_SCHEMA}.post-terminal-artifacts", f"{child_id}: post-terminal schema changed")
        check_descriptor(post.get("terminal"), child / "terminal.json", f"{child_id}.post.terminal")
        check_descriptor(post.get("stdout"), child / "stdout.log", f"{child_id}.post.stdout")
        check_descriptor(post.get("stderr"), child / "stderr.log", f"{child_id}.post.stderr")
        check_descriptor(post.get("time"), time_path, f"{child_id}.post.time")
        check_descriptor(post.get("report"), report_path, f"{child_id}.post.report")
        check_descriptor(post.get("validation"), child / "validation.json", f"{child_id}.post.validation")
        report = read_json(report_path, limit=REPORT_LIMIT)
        if lane == "native":
            validate_native_report(report, expected, plan)
        else:
            validate_allocator_report(report, expected, plan)
        result.append({"row": expected, "report": report, "time": time_metrics, "child": child_id, "child_path": child})
    return result


def metric_stats_native(record: dict[str, Any], metric: str) -> dict[str, float]:
    samples = record["report"]["case"]["samples"]
    return stats(sample[metric] for sample in samples)


def native_child_metrics(records: list[dict[str, Any]]) -> dict[str, dict[str, float]]:
    result: dict[str, dict[str, float]] = {}
    for record in records:
        metrics = {phase: metric_stats_native(record, phase) for phase in NATIVE_PHASES}
        metrics["max_rss_kib"] = {"value": float(record["time"]["max_rss_kib"]), "count": 1}
        result[record["child"]] = metrics
    return result


def allocator_child_metrics(record: dict[str, Any]) -> dict[str, Any]:
    rows = record["report"]["case"]["rows"]
    phase_summary: dict[str, dict[str, dict[str, float]]] = {}
    for phase in ALLOC_PHASES:
        phase_summary[phase] = {field: stats(row[phase][field] for row in rows) for field in ALLOC_FIELDS}
    derived = {
        "max_setup_core_peak_above_setup_live_before": stats(row["max_setup_core_peak_above_setup_live_before"] for row in rows),
        "max_setup_drained_core_peak_above_setup_live_before": stats(row["max_setup_drained_core_peak_above_setup_live_before"] for row in rows),
    }
    core_rows = [row["core_combined_with_drain"] for row in rows]
    selected = {
        "allocated_bytes": {"value": float(statistics.fmean(row["allocated_bytes"] for row in core_rows)), "aggregation": "mean over ten core_combined_with_drain rows"},
        "allocation_calls": {"value": float(statistics.fmean(row["allocation_calls"] for row in core_rows)), "aggregation": "mean over ten core_combined_with_drain rows"},
        "region_peak_live_bytes": {"value": float(max(row["region_peak_live_bytes"] for row in core_rows)), "aggregation": "maximum over ten core_combined_with_drain rows"},
        "core_combined_with_drain.allocated_bytes": {"value": float(statistics.fmean(row["allocated_bytes"] for row in core_rows)), "aggregation": "mean over ten core_combined_with_drain rows"},
        "max_rss_kib": {"value": float(record["time"]["max_rss_kib"]), "aggregation": "whole-process /usr/bin/time -v max RSS; separate from phase peaks"},
    }
    return {"phase_summary": phase_summary, "derived_peak_increments": derived, "selected": selected, "row_count": len(rows)}


def child_metric_value(metrics: dict[str, Any], metric: str) -> float:
    if metric == "max_rss_kib":
        return number(metrics[metric]["value"], metric)
    phase, quantile = metric.rsplit(".", 1)
    require(phase in NATIVE_PHASES and quantile in ("p50", "p95", "p99"), f"unknown native metric {metric}")
    return number(metrics[phase][quantile], metric)


def pair_ratio_metrics(a: dict[str, Any], b: dict[str, Any], metric_names: Iterable[str]) -> dict[str, Any]:
    return {metric: ratio(child_metric_value(b, metric), child_metric_value(a, metric), metric) for metric in metric_names}


def all_ratio_stats(pairs: list[dict[str, Any]], metric: str) -> dict[str, Any]:
    values: list[float] = []
    undefined: list[dict[str, Any]] = []
    for pair in pairs:
        value = pair["ratios"][metric]
        unit_id = pair.get("pair_id", f"repeat-{pair.get('repeat')}")
        parsed = ratio_value(value, f"{unit_id}.{metric}")
        if parsed is None:
            undefined.append({"unit_id": unit_id, "ratio": value})
        else:
            values.append(parsed)
    if undefined:
        return {"status": "undefined", "reason": "zero_denominator", "undefined_units": undefined, "unit_count": len(pairs)}
    return bootstrap(values)


def gate_assessment(aggregate: dict[str, Any]) -> dict[str, Any]:
    defined = aggregate.get("status") == "defined"
    point = aggregate.get("point") if defined else None
    upper = aggregate.get("ci_upper") if defined else None
    return {
        "safety_pass": bool(defined and upper <= SAFETY),
        "improvement_pass": bool(defined and point <= IMPROVEMENT and upper < CI_UPPER_FOR_IMPROVEMENT),
        "point": point,
        "ci_upper": upper,
    }


def build_native_case(records: list[dict[str, Any]], plan: dict[str, Any], case_id: str) -> dict[str, Any]:
    case_records = [record for record in records if record["row"]["case_id"] == case_id]
    require(len(case_records) == 24, f"native {case_id}: expected 24 children")
    by_block: dict[int, list[dict[str, Any]]] = {}
    for record in case_records:
        by_block.setdefault(record["row"]["block"], []).append(record)
    require(sorted(by_block) == list(range(6)) and all(len(v) == 4 for v in by_block.values()), f"native {case_id}: block coverage changed")
    child_metrics = native_child_metrics(case_records)
    metric_names = [f"{phase}.{quantile}" for phase in NATIVE_PHASES for quantile in ("p50", "p95", "p99")]
    metric_names.append("max_rss_kib")
    pairs: list[dict[str, Any]] = []
    blocks: list[dict[str, Any]] = []
    for block_id in range(6):
        block_records = sorted(by_block[block_id], key=lambda r: r["row"]["position"])
        require([r["row"]["role"] for r in block_records] == plan["comparison"]["blocks"][block_id]["order"], f"native {case_id} block {block_id}: role order changed")
        leg_data: list[dict[str, Any]] = []
        leg_ratios: dict[str, list[float]] = {metric: [] for metric in metric_names}
        for leg, (left, right) in enumerate(((block_records[0], block_records[1]), (block_records[2], block_records[3]))):
            if left["row"]["role"] == "A":
                a, b = left, right
            else:
                b, a = left, right
            pair_id = f"b{block_id:02d}-leg{leg}"
            ratios = pair_ratio_metrics(child_metrics[a["child"]], child_metrics[b["child"]], metric_names)
            for metric in metric_names:
                value = ratio_value(ratios[metric], f"{pair_id}.{metric}")
                if value is not None:
                    leg_ratios[metric].append(value)
            pairs.append({"pair_id": pair_id, "block": block_id, "leg": leg, "baseline_child": a["child"], "candidate_child": b["child"], "ratios": ratios})
            leg_data.append(pairs[-1])
        block_ratios: dict[str, Any] = {}
        for metric in metric_names:
            values = leg_ratios[metric]
            block_ratios[metric] = {"status": "defined", "value": geometric_mean(values)} if len(values) == 2 else {"status": "undefined", "reason": "zero_denominator"}
        blocks.append({"block": block_id, "legs": leg_data, "ratios": block_ratios})
    metrics: dict[str, Any] = {}
    for metric in metric_names:
        block_values = [block["ratios"][metric]["value"] for block in blocks if block["ratios"][metric].get("status") == "defined"]
        aggregate = bootstrap(block_values) if len(block_values) == 6 else {"status": "undefined", "reason": "zero_denominator", "unit_count": len(block_values)}
        legs = [pair["ratios"][metric] for pair in pairs]
        block_flags = [block["block"] for block in blocks if block["ratios"][metric].get("status") == "defined" and block["ratios"][metric]["value"] > SAFETY]
        leg_flags = [pair["pair_id"] for pair in pairs if pair["ratios"][metric].get("status") == "defined" and pair["ratios"][metric]["value"] > SAFETY]
        aggregate_flag = aggregate.get("status") == "defined" and aggregate.get("point", math.inf) > SAFETY
        metrics[metric] = {"aggregate": aggregate, "gate": gate_assessment(aggregate), "leg_ratios": legs, "block_ratios": [block["ratios"][metric] for block in blocks], "flags": {"leg_gt_1_05": leg_flags, "block_gt_1_05": block_flags, "aggregate_gt_1_05": bool(aggregate_flag)}}
    role_stats = {}
    for role in ("A", "B"):
        role_records = [record for record in case_records if record["row"]["role"] == role]
        role_stats[role] = {metric: stats(metric_stats_native(record, metric)["p50"] for record in role_records) for metric in NATIVE_PHASES}
        role_stats[role]["max_rss_kib"] = stats(record["time"]["max_rss_kib"] for record in role_records)
    phase_metrics = {
        phase: {quantile: metrics[f"{phase}.{quantile}"] for quantile in ("p50", "p95", "p99")}
        for phase in NATIVE_PHASES
    }
    return {
        "child_count": len(case_records),
        "role_child_summary": role_stats,
        "pairs": pairs,
        "blocks": blocks,
        "metrics": metrics,
        "phase_metrics": phase_metrics,
        "primary_metrics": list(NATIVE_PRIMARY),
        "secondary_metrics": list(NATIVE_SECONDARY),
        "publication_control": phase_metrics["publication_ns"],
        "rss": metrics["max_rss_kib"],
    }


def build_allocator_case(records: list[dict[str, Any]], plan: dict[str, Any], case_id: str) -> dict[str, Any]:
    case_records = [record for record in records if record["row"]["case_id"] == case_id]
    require(len(case_records) == 6, f"allocator {case_id}: expected 6 children")
    by_block: dict[int, list[dict[str, Any]]] = {}
    for record in case_records:
        by_block.setdefault(record["row"]["block"], []).append(record)
    require(sorted(by_block) == [0, 1, 2] and all(len(v) == 2 for v in by_block.values()), f"allocator {case_id}: repeat coverage changed")
    child_metrics = {record["child"]: allocator_child_metrics(record) for record in case_records}
    pairs: list[dict[str, Any]] = []
    for block_id in range(3):
        block_records = sorted(by_block[block_id], key=lambda r: r["row"]["position"])
        require([r["row"]["role"] for r in block_records] == plan["comparison"]["allocator_blocks"][block_id]["order"], f"allocator {case_id} block {block_id}: role order changed")
        left, right = block_records
        if left["row"]["role"] == "A":
            a, b = left, right
        else:
            b, a = left, right
        ratios: dict[str, Any] = {}
        for metric in ALLOC_METRICS:
            ratios[metric] = ratio(child_metrics[b["child"]]["selected"][metric]["value"], child_metrics[a["child"]]["selected"][metric]["value"], metric)
        pairs.append({"repeat": block_id, "baseline_child": a["child"], "candidate_child": b["child"], "ratios": ratios})
    metrics: dict[str, Any] = {}
    for metric in ALLOC_METRICS:
        aggregate = all_ratio_stats(pairs, metric)
        flags = [pair["repeat"] for pair in pairs if pair["ratios"][metric].get("status") == "defined" and pair["ratios"][metric]["value"] > SAFETY]
        metrics[metric] = {"aggregate": aggregate, "gate": gate_assessment(aggregate), "repeat_ratios": [pair["ratios"][metric] for pair in pairs], "flags": {"leg_gt_1_05": flags, "block_gt_1_05": flags, "repeat_gt_1_05": flags, "aggregate_gt_1_05": aggregate.get("status") == "defined" and aggregate.get("point", math.inf) > SAFETY}}
    descriptive = {
        record["child"]: {
            "case": case_id,
            "role": record["row"]["role"],
            "block": record["row"]["block"],
            "phase_summary": child_metrics[record["child"]]["phase_summary"],
            "derived_peak_increments": child_metrics[record["child"]]["derived_peak_increments"],
            "selected": child_metrics[record["child"]]["selected"],
            "row_count": child_metrics[record["child"]]["row_count"],
        }
        for record in case_records
    }
    return {"child_count": len(case_records), "descriptive_children": descriptive, "pairs": pairs, "metrics": metrics, "phase_metrics": list(ALLOC_FIELDS), "phase_names": list(ALLOC_PHASES), "process_rss_scope": "whole-process max_rss_kib from /usr/bin/time -v; never a phase peak"}


def reduce_packets(plan_path: Path, native_packet: Path, allocator_packet: Path, output: Path, helper_path: Path) -> dict[str, Any]:
    plan = read_json(plan_path)
    require(isinstance(plan, dict), "plan root is not an object")
    check_plan_shape(plan, require_authorized=True)
    validate_plan_bindings(plan, plan_path, helper_path)
    native_records = validate_lane_packet(native_packet, plan_path, plan, "native")
    allocator_records = validate_lane_packet(allocator_packet, plan_path, plan, "allocator")
    native_cases = {case["case_id"]: build_native_case(native_records, plan, case["case_id"]) for case in plan["cases"]}
    allocator_cases = {case["case_id"]: build_allocator_case(allocator_records, plan, case["case_id"]) for case in plan["cases"]}
    native_terminal = descriptor(native_packet / "capture-terminal.json")
    allocator_terminal = descriptor(allocator_packet / "capture-terminal.json")
    plan_desc = descriptor(plan_path)
    native_gate_metrics = [f"{phase}.{quantile}" for phase in NATIVE_PRIMARY for quantile in ("p50", "p95")]
    native_gate_metrics.append("max_rss_kib")
    allocator_gate_metrics = list(ALLOC_METRICS)
    gates = {
        "custody": True,
        "native_gate_metrics": native_gate_metrics,
        "allocator_gate_metrics": allocator_gate_metrics,
        "native": {case_id: {metric: native_cases[case_id]["metrics"][metric]["gate"] for metric in native_gate_metrics} for case_id in native_cases},
        "allocator": {case_id: {metric: allocator_cases[case_id]["metrics"][metric]["gate"] for metric in allocator_gate_metrics} for case_id in allocator_cases},
    }
    gates["native_safety"] = {case_id: all(value["safety_pass"] for value in metrics.values()) for case_id, metrics in gates["native"].items()}
    gates["allocator_safety"] = {case_id: all(value["safety_pass"] for value in metrics.values()) for case_id, metrics in gates["allocator"].items()}
    gates["all_safety"] = all(gates["native_safety"].values()) and all(gates["allocator_safety"].values())
    result = {
        "schema": REDUCER_SCHEMA,
        "status": "reduced",
        "scope": "offline reduction of sealed 0959 v4 native and allocator packets; no workload or build launched",
        "plan": plan_desc,
        "packets": {"native": native_terminal, "allocator": allocator_terminal},
        "schedule": {"native_children": len(native_records), "allocator_children": len(allocator_records), "native_blocks_per_case": 6, "allocator_repeats_per_case": 3},
        "bootstrap": {"resamples": BOOTSTRAP_RESAMPLES, "seed": BOOTSTRAP_SEED, "unit_native": "six block geometric means", "unit_allocator": "three paired repeats", "ci": "nearest-rank two-sided percentile bootstrap; 2.5%/97.5% endpoints"},
        "native": {"cases": native_cases, "children": [{"child_id": r["child"], "case_id": r["row"]["case_id"], "role": r["row"]["role"], "block": r["row"]["block"], "time": r["time"], "report_sha256": sha256(r["child_path"] / "output" / "report.json")} for r in native_records]},
        "allocator": {"cases": allocator_cases, "children": [{"child_id": r["child"], "case_id": r["row"]["case_id"], "role": r["row"]["role"], "block": r["row"]["block"], "time": r["time"]} for r in allocator_records]},
        "gates": gates,
        "limitations": [
            "Allocator phase summaries use serialized region counters and observer peaks; max_rss_kib is retained separately as whole-process evidence.",
            "Native p99, publication, ingress, and commit-drop values are retained as review controls; primary timing gates are edit_commit and total p50/p95.",
            "No adoption or production claim is made by this reducer.",
        ],
    }
    write_new(output, result)
    return result


def self_test(output: Path) -> dict[str, Any]:
    """Run calculation and custody controls without creating a workload packet."""
    checks: list[dict[str, Any]] = []
    def check(name: str, fn, expected: str = "pass") -> None:
        try:
            fn()
        except Exception as exc:
            checks.append({"name": name, "status": "unexpected_failure", "error": str(exc)})
        else:
            checks.append({"name": name, "status": expected})
    check("quantiles_even_median", lambda: require(stats([1, 2, 3, 4])["p50"] == 2.5, "median changed"))
    check("native_bootstrap_deterministic", lambda: require(bootstrap([0.9, 1.0, 1.1])["point"] == bootstrap([0.9, 1.0, 1.1])["point"], "bootstrap is not deterministic"))
    check("zero_denominator_is_explicit", lambda: require(ratio_value(ratio(1, 0, "negative"), "negative") is None, "zero denominator was not retained"))
    def undefined_bootstrap():
        value = bootstrap([1.0, 0.0])
        require(value.get("status") == "defined" and value.get("point") == 0.0, "zero numerator was not retained as a valid ratio")
    check("zero_numerator_bootstrap_is_valid", undefined_bootstrap)
    def negative_bootstrap():
        value = bootstrap([1.0, -1.0])
        require(value.get("status") == "undefined", "negative ratio did not fail closed")
    check("negative_ratio_bootstrap_fails_closed", negative_bootstrap)
    def nonadjacent_allocator_region():
        base = {
            "live_bytes_before": 0, "live_bytes_after": 1,
            "peak_live_bytes_before": 0, "peak_live_bytes_after": 1,
            "region_peak_live_bytes": 1,
            "allocation_calls": 1, "deallocation_calls": 0,
            "reallocation_calls": 0, "failed_allocation_calls": 0,
            "allocated_bytes": 1, "deallocated_bytes": 0,
        }
        left = {"status": "measured", "scope": "operation_global_system_allocator", **base}
        right = {"status": "measured", "scope": "operation_global_system_allocator", **{**base, "live_bytes_before": 99}}
        try:
            combine_alloc([left, right], "negative")
        except ReductionError:
            return
        raise AssertionError("non-adjacent allocator phases were accepted")
    check("allocator_live_boundary_rejected", nonadjacent_allocator_region)
    def missing_terminal():
        with tempfile_dir() as packet:
            (packet / "capture-start.json").write_text("{}")
            try:
                validate_lane_packet(packet, Path("/missing-plan"), {"lanes": {"native": {"execution": {}}}, "cases": [], "comparison": {"blocks": []}}, "native")
            except ReductionError:
                return
            raise AssertionError("missing capture terminal was accepted")
    check("missing_capture_terminal_rejected", missing_terminal, expected="pass")
    def wrong_child_count():
        plan = {"comparison": {"blocks": [], "allocator_blocks": []}}
        # The static check is intentionally exercised independently from packet I/O.
        try:
            check_plan_shape({"schema": PLAN_SCHEMA, "status": "prospective_candidate_pending", "cases": [], "comparison": plan["comparison"], "lanes": {}, "gates": {}}, require_authorized=False)
        except ReductionError:
            return
        raise AssertionError("wrong child schedule accepted")
    check("wrong_schedule_rejected", wrong_child_count)
    result = {"schema": "litchi.performance.0959.reducer-controls.v2", "status": "passed" if all(item["status"] == "pass" for item in checks) else "failed", "scope": "synthetic reducer and custody controls; no workload or build", "checks": checks}
    write_new(output, result)
    return result


class tempfile_dir:
    """Tiny context manager kept local so controls do not import a runner."""
    def __enter__(self):
        import tempfile
        self._value = tempfile.TemporaryDirectory(prefix="reducer-0959-")
        return Path(self._value.name)
    def __exit__(self, exc_type, exc, tb):
        self._value.cleanup()
        return False


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--packet", type=Path, help="parent containing native/ and allocator/ lane packets")
    parser.add_argument("--native-packet", type=Path)
    parser.add_argument("--allocator-packet", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--helper", type=Path, default=Path(__file__))
    parser.add_argument("--check-plan", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args(argv)
    try:
        if args.self_test:
            result = self_test(args.output)
        else:
            require(args.plan is not None, "--plan is required")
            plan_path = args.plan.resolve()
            plan = read_json(plan_path)
            require(isinstance(plan, dict), "plan root is not an object")
            if args.check_plan:
                check_plan_shape(plan, require_authorized=False)
                print(json.dumps({"status": "plan-shape-pass", "schema": plan.get("schema"), "plan_sha256": sha256(plan_path)}, sort_keys=True))
                return 0
            if args.packet is not None:
                root = args.packet.resolve()
                native_packet = root / "native"
                allocator_packet = root / "allocator"
            else:
                require(args.native_packet is not None and args.allocator_packet is not None, "--packet or both lane packet paths are required")
                native_packet = args.native_packet.resolve()
                allocator_packet = args.allocator_packet.resolve()
            result = reduce_packets(plan_path, native_packet, allocator_packet, args.output.resolve(), args.helper.resolve())
        print(json.dumps({"status": result["status"], "output": str(args.output.resolve())}, sort_keys=True))
        return 0 if result["status"] in ("reduced", "passed") else 1
    except (ReductionError, OSError, ValueError) as exc:
        print(f"reduce_0959_v1: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
