#!/usr/bin/env python3
"""Strict synthetic controls for the 0957 v2 allocator contract.

These controls never execute a benchmark.  They exercise the same v2 report
validator used by the runner/reducer, including phase adjacency, all additive
counters, exact derived peaks, context/admission oracles, drain custody, and
resource/sink checks.  A separate real tiny preflight remains mandatory.
"""
from __future__ import annotations

import copy
import json
import tempfile
from pathlib import Path

from common_0957_v2 import (
    ContractError,
    PHASES,
    exclusive_json,
    file_descriptor,
    load_json,
    validate_plan_schedule,
    validate_report,
)

ROOT = Path(__file__).resolve().parents[1]
PLAN = ROOT / "review" / "plan-schedule-v5-v2-draft.json"
OUT = ROOT / "review" / "controls-v2-final-result-v7.json"


def sample(before: int, after: int, calls: int, allocated: int, deallocated: int = 0) -> dict[str, object]:
    return {
        "status": "measured",
        "scope": "operation_global_system_allocator",
        "allocation_calls": calls,
        "deallocation_calls": deallocated,
        "reallocation_calls": 0,
        "failed_allocation_calls": 0,
        "allocated_bytes": allocated,
        "deallocated_bytes": deallocated,
        "live_bytes_before": before,
        "live_bytes_after": after,
        "peak_live_bytes_before": before,
        "peak_live_bytes_after": max(before, after),
        "region_peak_live_bytes": max(before, after),
    }


def combine(parts: list[dict[str, object]]) -> dict[str, object]:
    result = {
        "status": "measured",
        "scope": "operation_global_system_allocator",
        "live_bytes_before": parts[0]["live_bytes_before"],
        "live_bytes_after": parts[-1]["live_bytes_after"],
        "peak_live_bytes_before": parts[0]["peak_live_bytes_before"],
        "peak_live_bytes_after": parts[-1]["peak_live_bytes_after"],
        "region_peak_live_bytes": max(part["region_peak_live_bytes"] for part in parts),
    }
    for field in (
        "allocation_calls",
        "deallocation_calls",
        "reallocation_calls",
        "failed_allocation_calls",
        "allocated_bytes",
        "deallocated_bytes",
    ):
        result[field] = sum(part[field] for part in parts)
    return result


def report(plan: dict[str, object], child: dict[str, object]) -> dict[str, object]:
    case = next(item for item in plan["cases"] if item["case_id"] == child["case_id"])
    identity = plan["binary"]
    context = case["context_limits_by_workers"][str(child["workers"])]
    admission = case["admission_by_workers"][str(child["workers"])]
    ingress = sample(100, 110, 1, 10)
    edit_commit = sample(110, 120, 2, 20)
    publication = sample(120, 130, 3, 30)
    commit_drop = sample(130, 140, 4, 40, 4)
    worker_drain = sample(140, 150, 5, 50, 5)
    core = combine([ingress, edit_commit, publication, commit_drop])
    core_with_drain = combine([ingress, edit_commit, publication, commit_drop, worker_drain])
    rows = []
    for index in range(10):
        rows.append({
            "sample": index,
            "setup": sample(100, 100, 0, 0),
            "ingress": copy.deepcopy(ingress),
            "edit_commit": copy.deepcopy(edit_commit),
            "publication": copy.deepcopy(publication),
            "commit_drop": copy.deepcopy(commit_drop),
            "worker_drain": copy.deepcopy(worker_drain),
            "core_combined": copy.deepcopy(core),
            "core_combined_with_drain": copy.deepcopy(core_with_drain),
            "postdrop": sample(150, 100, 0, 0),
            "max_setup_core_peak_above_setup_live_before": 40,
            "max_setup_drained_core_peak_above_setup_live_before": 50,
            "context_limits": copy.deepcopy(context),
            "workers_requested": child["workers"],
            "expected_worker_permits": admission["admission_expected_worker_permits"],
            "admitted_worker_permits": admission["admission_observed_worker_permits"],
            "expected_cpu_tasks": case["expected_cpu_tasks"],
            "cpu_tasks_used": case["expected_cpu_tasks"],
            "output_bytes": case["output_bytes"],
            "output_bytes_used": case["output_bytes"],
            "sink_write_calls": 1,
            "sink_largest_write": 64,
            "resources_after_core": {"memory": 0, "workers": 0, "io_concurrency": 0, "objects": 0},
            "source_sha256": case["source_sha256"],
            "output_sha256": case["output_sha256"],
            "semantic_sha256": case["semantic_sha256"],
            "fixture_inverse_restored_exact_source": True,
            "pre_sample_threads": {
                "first_thread_count": 1, "final_thread_count": 1, "polls": 1,
                "drained": True, "deadline_seconds": 2, "poll_limit": 10000,
                "poll_sleep_milliseconds": 1,
            },
            "post_core_threads": {
                "first_thread_count": 2, "final_thread_count": 1, "polls": 2,
                "drained": True, "deadline_seconds": 2, "poll_limit": 10000,
                "poll_sleep_milliseconds": 1,
            },
        })
    calibration = {
        "status": "pass",
        "polls": 100,
        "thread_count_first": 1,
        "thread_count_last": 1,
        "allocation_calls": 0,
        "deallocation_calls": 0,
        "reallocation_calls": 0,
        "failed_allocation_calls": 0,
        "no_allocator_callbacks": True,
        "poll_sleep_milliseconds": 1,
        "thread_scope": plan["report_contract"]["thread_observation_scope"],
    }
    return {
        "schema": plan["report_contract"]["schema"],
        "status": "completed",
        "binary_identity": {
            "path": identity["path"], "bytes": identity["bytes"], "sha256": identity["sha256"],
        },
        "allocator_identity": plan["report_contract"]["allocator_identity"],
        "instrumentation_identity": plan["report_contract"]["instrumentation_identity"],
        "counter_revision": plan["report_contract"]["counter_revision"],
        "allocation_scope": plan["report_contract"]["allocation_scope"],
        "thread_observation_scope": plan["report_contract"]["thread_observation_scope"],
        "timing_scope": plan["report_contract"]["timing_scope"],
        "performance_claim": plan["report_contract"]["performance_claim"],
        "peak_semantics": plan["report_contract"]["peak_semantics"],
        "calibration": calibration,
        "case": {
            "id": child["case_id"] + "-workers" + str(child["workers"]),
            "shape": case["shape"],
            "workload": case["workload"],
            "workers": child["workers"],
            "update_count": case["update_count"],
            "source_bytes": case["source_bytes"],
            "source_sha256": case["source_sha256"],
            "output_bytes": case["output_bytes"],
            "output_sha256": case["output_sha256"],
            "semantic_sha256": case["semantic_sha256"],
            "expected_cpu_tasks": case["expected_cpu_tasks"],
            "admission_expected_worker_permits": admission["admission_expected_worker_permits"],
            "admission_observed_worker_permits": admission["admission_observed_worker_permits"],
            "admission_cpu_tasks_used": admission["admission_cpu_tasks_used"],
            "preflight_actual_commit_inverse_verified": True,
            "context_limits": copy.deepcopy(context),
            "sink_capacity_bytes": case["sink_capacity_bytes"],
            "sink_max_write_bytes": case["sink_max_write_bytes"],
            "warmup_iterations": 2,
            "measured_samples": 10,
            "rows": rows,
        },
    }


def expect_reject(outcomes: dict[str, str], name: str, value: dict[str, object], plan: dict[str, object], child: dict[str, object]) -> None:
    try:
        validate_report(value, plan, child)
    except ContractError:
        outcomes[name] = "rejected"
    else:
        outcomes[name] = "unexpected_pass"


def main() -> int:
    plan = load_json(PLAN)
    children = validate_plan_schedule(plan)
    child = children[0]
    positive = report(plan, child)
    outcomes: dict[str, str] = {}
    try:
        validate_report(positive, plan, child)
        outcomes["positive"] = "passed"
    except Exception as exc:
        outcomes["positive"] = f"failed: {exc}"

    missing_pre = copy.deepcopy(positive)
    del missing_pre["case"]["rows"][0]["pre_sample_threads"]
    expect_reject(outcomes, "missing_pre_sample_threads", missing_pre, plan, child)

    bad_failed = copy.deepcopy(positive)
    bad_failed["case"]["rows"][0]["publication"]["failed_allocation_calls"] = 1
    expect_reject(outcomes, "failed_allocation_nonzero", bad_failed, plan, child)

    bad_reallocation = copy.deepcopy(positive)
    bad_reallocation["case"]["rows"][0]["core_combined"]["reallocation_calls"] += 1
    expect_reject(outcomes, "reallocation_conservation", bad_reallocation, plan, child)

    bad_deallocation = copy.deepcopy(positive)
    bad_deallocation["case"]["rows"][0]["core_combined"]["deallocation_calls"] += 1
    expect_reject(outcomes, "deallocation_conservation", bad_deallocation, plan, child)

    bad_allocated = copy.deepcopy(positive)
    bad_allocated["case"]["rows"][0]["core_combined"]["allocated_bytes"] += 1
    expect_reject(outcomes, "allocated_bytes_conservation", bad_allocated, plan, child)

    bad_deallocated = copy.deepcopy(positive)
    bad_deallocated["case"]["rows"][0]["core_combined"]["deallocated_bytes"] += 1
    expect_reject(outcomes, "deallocated_bytes_conservation", bad_deallocated, plan, child)

    bad_adjacency = copy.deepcopy(positive)
    bad_adjacency["case"]["rows"][0]["edit_commit"]["live_bytes_before"] += 1
    expect_reject(outcomes, "live_adjacency", bad_adjacency, plan, child)

    bad_combined_peak = copy.deepcopy(positive)
    bad_combined_peak["case"]["rows"][0]["core_combined"]["region_peak_live_bytes"] += 1
    expect_reject(outcomes, "combined_peak_maximum", bad_combined_peak, plan, child)

    bad_combined_highwater = copy.deepcopy(positive)
    bad_combined_highwater["case"]["rows"][0]["core_combined"]["peak_live_bytes_after"] += 1
    expect_reject(outcomes, "combined_highwater_maximum", bad_combined_highwater, plan, child)

    bad_peak = copy.deepcopy(positive)
    bad_peak["case"]["rows"][0]["max_setup_drained_core_peak_above_setup_live_before"] += 1
    expect_reject(outcomes, "derived_peak_formula", bad_peak, plan, child)

    bad_context = copy.deepcopy(positive)
    bad_context["case"]["rows"][0]["context_limits"]["memory"] += 1
    expect_reject(outcomes, "context_limits", bad_context, plan, child)

    bad_admission = copy.deepcopy(positive)
    bad_admission["case"]["rows"][0]["admitted_worker_permits"] += 1
    expect_reject(outcomes, "admission_permits", bad_admission, plan, child)

    bad_cpu = copy.deepcopy(positive)
    bad_cpu["case"]["rows"][0]["cpu_tasks_used"] += 1
    expect_reject(outcomes, "cpu_task_counter", bad_cpu, plan, child)

    bad_output_hash = copy.deepcopy(positive)
    bad_output_hash["case"]["rows"][0]["output_sha256"] = "0" * 64
    expect_reject(outcomes, "output_hash", bad_output_hash, plan, child)

    bad_output_counter = copy.deepcopy(positive)
    bad_output_counter["case"]["rows"][0]["output_bytes_used"] += 1
    expect_reject(outcomes, "output_bytes_counter", bad_output_counter, plan, child)

    bad_resources = copy.deepcopy(positive)
    bad_resources["case"]["rows"][0]["resources_after_core"]["objects"] = 1
    expect_reject(outcomes, "resource_snapshot", bad_resources, plan, child)

    bad_sink = copy.deepcopy(positive)
    bad_sink["case"]["rows"][0]["sink_largest_write"] = 65537
    expect_reject(outcomes, "sink_bound", bad_sink, plan, child)

    bad_calibration = copy.deepcopy(positive)
    bad_calibration["calibration"]["polls"] = 99
    expect_reject(outcomes, "calibration_polls", bad_calibration, plan, child)

    bad_calibration_counter = copy.deepcopy(positive)
    bad_calibration_counter["calibration"]["deallocation_calls"] = 1
    expect_reject(outcomes, "calibration_serialized_counter", bad_calibration_counter, plan, child)

    bad_post = copy.deepcopy(positive)
    del bad_post["case"]["rows"][0]["post_core_threads"]
    expect_reject(outcomes, "missing_post_core_threads", bad_post, plan, child)

    bad_poll_bound = copy.deepcopy(positive)
    bad_poll_bound["case"]["rows"][0]["post_core_threads"]["polls"] = 10001
    expect_reject(outcomes, "thread_poll_bound", bad_poll_bound, plan, child)

    with tempfile.TemporaryDirectory(prefix="litchi-0957-v2-json-controls-") as directory:
        duplicate = Path(directory) / "duplicate.json"
        duplicate.write_text('{"a":1,"a":2}\n')
        try:
            load_json(duplicate)
            outcomes["duplicate_json_key"] = "unexpected_pass"
        except ContractError:
            outcomes["duplicate_json_key"] = "rejected"

    result = {
        "schema": "litchi.performance.0957.controls.v2",
        "status": "passed",
        "plan": file_descriptor(PLAN),
        "validator": "common_0957_v2.validate_report",
        "outcomes": outcomes,
        "synthetic_only": True,
        "requires_real_tiny_preflight_before_formal": True,
    }
    if outcomes.get("positive") != "passed" or any(
        value != "rejected" for key, value in outcomes.items() if key != "positive"
    ):
        result["status"] = "failed"
    exclusive_json(OUT, result)
    print(json.dumps(result, sort_keys=True))
    return 0 if result["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
