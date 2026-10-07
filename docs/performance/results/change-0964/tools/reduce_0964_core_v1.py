#!/usr/bin/env python3
"""Small, fail-closed offline reducer for the 0964 screening packet.

The screening lane is deliberately descriptive.  It has three native blocks
per case (60 samples, 10 warmups) and three allocator pairs per case.  A
screening signal never authorizes adoption; an encouraging or ambiguous result
requires the full six-block/200-sample protocol.  Packet custody and final-plan
bindings are checked before any report is parsed.

The report and allocator validators are imported from the frozen 0959 adapter
so their oracle, scope, and phase-conservation checks have one implementation.
Only the screening schedule and its non-bootstrapped descriptive summaries are
implemented here.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import statistics
import sys
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0964").resolve()
BASE_REDUCER = Path("/home/zhuhe/litchi-performance-0959/tools/reduce_0959_v3_finalplan.py").resolve()
PLAN_SCHEMA = "litchi.performance.0964.inline-number-screen-plan.v1"
NATIVE_COUNT = 36
ALLOCATOR_COUNT = 18
NATIVE_WARMUPS = 10
NATIVE_SAMPLES = 60
ALLOCATOR_WARMUPS = 2
ALLOCATOR_SAMPLES = 10
CPU_MASK = [12, 13, 14, 15]
DEADLINE = 900
SCREEN_SAFETY = 1.05
MAX_SINK_WRITE = 65_536
NATIVE_PHASES = ("ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns")


class ReductionError(ValueError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReductionError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    value = path.stat()
    return {"path": str(path.resolve()), "bytes": value.st_size, "mode": value.st_mode & 0o777, "sha256": sha256(path)}


def load_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as stream:
        return json.load(stream)


def load_base():
    spec = importlib.util.spec_from_file_location("reducer_0959_frozen", BASE_REDUCER)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load frozen reducer: {BASE_REDUCER}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


base = load_base()


def expected_cases(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {case["case_id"]: case for case in plan["cases"]}


def check_screen_plan(plan: dict[str, Any]) -> None:
    require(plan.get("schema") == PLAN_SCHEMA, "plan schema is not 0964 screening v1")
    require(plan.get("status") == "authorized", "screening reduction requires an authorized plan")
    cases = plan.get("cases")
    require(isinstance(cases, list) and [case.get("case_id") for case in cases] == ["medium-one-percent", "dense-sparse-one-percent", "medium-one-edit"], "screening case order changed")
    for case in cases:
        require(case.get("shape") in {"medium", "dense-sparse"}, f"unknown shape: {case.get('shape')}")
        require(case.get("workload") in {"one-percent", "one-edit"}, f"unknown workload: {case.get('workload')}")
        require(isinstance(case.get("context_limits"), dict), f"{case['case_id']}: context limits missing")
    comparison = plan.get("comparison")
    require(isinstance(comparison, dict), "comparison section missing")
    blocks = comparison.get("blocks")
    allocator_blocks = comparison.get("allocator_blocks")
    require(isinstance(blocks, list) and len(blocks) == 3 and [item.get("id") for item in blocks] == [0, 1, 2], "screening native blocks must be exactly three")
    require(isinstance(allocator_blocks, list) and len(allocator_blocks) == 3 and [item.get("id") for item in allocator_blocks] == [0, 1, 2], "screening allocator blocks must be exactly three")
    for item in blocks:
        require(item.get("order") in (["A", "B", "B", "A"], ["B", "A", "A", "B"]), "native block is not ABBA/BAAB")
    for item in allocator_blocks:
        require(item.get("order") in (["A", "B"], ["B", "A"]), "allocator block is not AB/BA")
    require(comparison.get("block_count") == 3, "native block count changed")
    require(comparison.get("pairs_per_case_native") == 6, "screening native pair count changed")
    require(comparison.get("pairs_per_case_allocator") == 3, "screening allocator pair count changed")
    require(comparison.get("screening_only") is True, "screening-only marker missing")
    require(comparison.get("full_followup_required") is True, "full follow-up requirement missing")
    lanes = plan.get("lanes")
    require(isinstance(lanes, dict) and set(lanes) >= {"native", "allocator"}, "lanes missing")
    for lane, count, warmup, samples, schema in (
        ("native", NATIVE_COUNT, NATIVE_WARMUPS, NATIVE_SAMPLES, "litchi.xlsx-managed-compression-benchmark.v1"),
        ("allocator", ALLOCATOR_COUNT, ALLOCATOR_WARMUPS, ALLOCATOR_SAMPLES, "litchi.xlsx-managed-compression-allocation.v1"),
    ):
        spec = lanes[lane]
        require(spec.get("child_count") == count, f"{lane} child count changed")
        require(spec.get("warmup_iterations") == warmup and spec.get("measured_samples") == samples, f"{lane} sample schedule changed")
        require(spec.get("report_schema") == schema, f"{lane} report schema changed")
        execution = spec.get("execution")
        require(isinstance(execution, dict), f"{lane} execution missing")
        require(execution.get("cpu_mask") == CPU_MASK, f"{lane} CPU mask changed")
        require(execution.get("deadline_seconds") == DEADLINE, f"{lane} deadline changed")
        require(execution.get("max_parallel_children") == 1 and execution.get("retry_or_replacement") is False, f"{lane} serial/retry policy changed")
        require(execution.get("terminal_before_parse") is True, f"{lane} terminal-first policy changed")
        rows = spec.get("children")
        require(isinstance(rows, list) and len(rows) == count, f"{lane} child schedule missing")
        require(rows == base.expected_rows(plan, lane), f"{lane} child schedule differs from plan")
    runner = plan.get("runner")
    require(isinstance(runner, dict), "runner binding missing")
    require(isinstance(runner.get("schema"), str) and runner["schema"], "runner schema missing")
    require(runner.get("deadline_seconds") == DEADLINE and runner.get("cpu_mask") == CPU_MASK, "runner custody values changed")
    gates = plan.get("gates")
    require(isinstance(gates, dict) and gates.get("adoption_allowed") is False and gates.get("full_followup_required") is True, "screening adoption/follow-up gates changed")


def validate_plan_bindings(plan: dict[str, Any], plan_path: Path, helper_path: Path) -> None:
    source = plan.get("source")
    require(isinstance(source, dict), "source binding missing")
    for key in ("baseline_manifest", "candidate_manifest", "baseline_receipt", "candidate_freeze"):
        value = source.get(key)
        require(isinstance(value, dict) and value.get("status") != "pending_root_binding", f"source.{key} is pending")
        base.external_descriptor(value, f"plan.source.{key}")
    for lane in ("native", "allocator"):
        for role in ("baseline_binary", "candidate_binary"):
            value = plan["lanes"][lane].get(role)
            require(isinstance(value, dict) and value.get("status") != "pending_root_binding", f"{lane}.{role} is pending")
            base.external_descriptor(value, f"plan.lanes.{lane}.{role}", executable=True)
    helpers = plan.get("helpers")
    require(isinstance(helpers, dict), "helper bindings missing")
    reducer = helpers.get("reducer")
    actual = descriptor(helper_path)
    require(isinstance(reducer, dict) and all(reducer.get(key) == actual[key] for key in ("path", "bytes", "sha256")), "plan reducer binding differs")
    for key in ("capture", "controls"):
        value = helpers.get(key)
        require(isinstance(value, dict) and value.get("status") != "pending_root_binding", f"plan helper {key} is pending")
        if "path" in value:
            base.external_descriptor(value, f"plan.helpers.{key}")
    require(isinstance(plan.get("host"), dict) and plan["host"].get("status") != "pending_root_binding", "host binding is pending")
    require(isinstance(plan.get("root_go"), dict) and plan["root_go"].get("status") != "pending_root_binding", "root GO binding is pending")
    # The packet start records must point to this exact plan descriptor; this
    # is checked again by base.validate_lane_packet before any report parse.
    require(plan_path.is_file(), "plan file missing")


NATIVE_TIMING_SCOPE = "with fresh finite managed context prepared outside the clock, the timer begins at SourceBackedEditor construction and covers edit/commit, source-backed publication, returned snapshot drop, and commit/package worker-pool teardown; precomputed source/cache/context setup and bounded sink allocation are outside"
NATIVE_PHASE_SCOPE = "ingress_ns is the direct SourceBackedEditor constructor; edit_commit_ns is selector staging, cell sets, and MultiCommit::commit; publication_ns is publish_multi_commit_to_stream plus returned MultiSnapshot drop; commit_drop_ns is MultiCommit drop; total_ns starts at the constructor and ends after commit teardown"
NATIVE_SETUP_SCOPE = "deterministic corpus and expected output construction, frozen source/output/semantic oracles, SourceCacheLimits and output ceiling, one untimed admission preflight, inverse patch, fresh OwnedSource clone, finite managed context, and sink capacity reservation are outside measured phases"
NATIVE_VERIFICATION_SCOPE = "exact output bytes, immutable fixture source hash, output/semantic hashes, workbook verification, sink bounds, CpuTasks/OutputBytes counters, and post-drop Memory/Workers/IoConcurrency/Objects checks run outside timers; the untimed admission preflight verifies an actual managed-commit inverse (preflight_actual_commit_inverse_verified=true), while each timed sample applies the precomputed inverse patch only as a fixture to its published bytes (fixture_inverse_restored_exact_source)"
NATIVE_PERFORMANCE_CLAIM = "none; raw phase timings describe this explicit in-memory harness and do not claim OS thread counts, speedup, allocation behavior, or physical I/O"


def validate_native_screen_report(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    case = base.validate_common_report(report, row, plan, "native")
    for key, value in {"timing_scope": NATIVE_TIMING_SCOPE, "phase_scope": NATIVE_PHASE_SCOPE, "setup_scope": NATIVE_SETUP_SCOPE, "verification_scope": NATIVE_VERIFICATION_SCOPE, "performance_claim": NATIVE_PERFORMANCE_CLAIM}.items():
        base.require(report.get(key) == value, f"{row['child_id']}: report {key} scope changed")
    samples = case.get("samples")
    base.require(isinstance(samples, list) and len(samples) == NATIVE_SAMPLES, f"{row['child_id']}: native screening samples missing")
    expected = base.expected_case(plan, row["case_id"])
    for index, sample in enumerate(samples):
        base.require(isinstance(sample, dict) and sample.get("sample") == index, f"{row['child_id']}: native sample index changed")
        values = [base.integer(sample.get(name), f"{row['child_id']}.sample[{index}].{name}") for name in NATIVE_PHASES]
        base.require(sum(values[:4]) <= values[-1], f"{row['child_id']}: phase sum exceeds total")
        for key in ("output_bytes", "output_bytes_used"):
            base.require(sample.get(key) == expected["output_bytes"], f"{row['child_id']}: {key} oracle changed")
        base.require(sample.get("cpu_tasks_used") == expected["expected_cpu_tasks"], f"{row['child_id']}: CpuTasks changed")
        base.require(base.integer(sample.get("sink_write_calls"), "sink_write_calls", minimum=1) >= 1, f"{row['child_id']}: sink calls missing")
        base.require(0 < base.integer(sample.get("sink_largest_write"), "sink_largest_write") <= MAX_SINK_WRITE, f"{row['child_id']}: sink bound changed")
        base.require(sample.get("output_sha256") == case["output_sha256"] and sample.get("semantic_sha256") == case["semantic_sha256"], f"{row['child_id']}: output oracle changed")
        base.require(sample.get("fixture_inverse_restored_exact_source") is True, f"{row['child_id']}: inverse restoration failed")
        resources = sample.get("resources_after_drop")
        base.require(isinstance(resources, dict) and all(resources.get(key) == 0 for key in ("memory", "workers", "io_concurrency", "objects")), f"{row['child_id']}: resources after drop not zero")
    return case


def screen_native_case(records: list[dict[str, Any]], plan: dict[str, Any], case_id: str) -> dict[str, Any]:
    case_records = [record for record in records if record["row"]["case_id"] == case_id]
    base.require(len(case_records) == 12, f"native {case_id}: expected 12 screening children")
    by_block: dict[int, list[dict[str, Any]]] = {}
    for record in case_records:
        by_block.setdefault(record["row"]["block"], []).append(record)
    base.require(sorted(by_block) == [0, 1, 2] and all(len(rows) == 4 for rows in by_block.values()), f"native {case_id}: screening block coverage changed")
    metrics_by_child = {}
    for record in case_records:
        metrics_by_child[record["child"]] = {phase: base.metric_stats_native(record, phase) for phase in NATIVE_PHASES}
        metrics_by_child[record["child"]]["max_rss_kib"] = {"value": float(record["time"]["max_rss_kib"]), "count": 1}
    metric_names = [f"{phase}.{quantile}" for phase in NATIVE_PHASES for quantile in ("p50", "p95", "p99")] + ["max_rss_kib"]
    pairs = []
    for block_id in range(3):
        rows = sorted(by_block[block_id], key=lambda record: record["row"]["position"])
        base.require([record["row"]["role"] for record in rows] == plan["comparison"]["blocks"][block_id]["order"], f"native {case_id} block {block_id}: role order changed")
        for leg, (left, right) in enumerate(((rows[0], rows[1]), (rows[2], rows[3]))):
            if left["row"]["role"] == "A":
                a, b = left, right
            else:
                b, a = left, right
            ratios = base.pair_ratio_metrics(metrics_by_child[a["child"]], metrics_by_child[b["child"]], metric_names)
            pairs.append({"pair_id": f"b{block_id:02d}-leg{leg}", "block": block_id, "ratios": ratios, "baseline_child": a["child"], "candidate_child": b["child"]})
    metrics = {}
    for metric in metric_names:
        values = [base.ratio_value(pair["ratios"][metric], f"{pair['pair_id']}.{metric}") for pair in pairs]
        defined = [value for value in values if value is not None]
        metrics[metric] = {
            "screening_statistic": "median of six paired screening legs",
            "pair_count": len(defined),
            "median": statistics.median(defined) if defined else None,
            "min": min(defined) if defined else None,
            "max": max(defined) if defined else None,
            "undefined_count": len(values) - len(defined),
            "review_flag_pair_gt_1_05": sum(value > SCREEN_SAFETY for value in defined),
        }
    return {"child_count": len(case_records), "pairs": pairs, "metrics": metrics, "screening_only": True, "full_followup_required": True}


def screen_allocator_case(records: list[dict[str, Any]], plan: dict[str, Any], case_id: str) -> dict[str, Any]:
    case_records = [record for record in records if record["row"]["case_id"] == case_id]
    base.require(len(case_records) == 6, f"allocator {case_id}: expected 6 screening children")
    by_block: dict[int, list[dict[str, Any]]] = {}
    for record in case_records:
        by_block.setdefault(record["row"]["block"], []).append(record)
    base.require(sorted(by_block) == [0, 1, 2] and all(len(rows) == 2 for rows in by_block.values()), f"allocator {case_id}: screening repeat coverage changed")
    children = {record["child"]: base.allocator_child_metrics(record) for record in case_records}
    metric_names = list(base.ALLOC_METRICS)
    pairs = []
    for block_id in range(3):
        rows = sorted(by_block[block_id], key=lambda record: record["row"]["position"])
        base.require([record["row"]["role"] for record in rows] == plan["comparison"]["allocator_blocks"][block_id]["order"], f"allocator {case_id} block {block_id}: order changed")
        left, right = rows
        if left["row"]["role"] == "A":
            a, b = left, right
        else:
            b, a = left, right
        ratios = {metric: base.ratio(children[b["child"]]["selected"][metric]["value"], children[a["child"]]["selected"][metric]["value"], metric) for metric in metric_names}
        pairs.append({"repeat": block_id, "ratios": ratios, "baseline_child": a["child"], "candidate_child": b["child"]})
    metrics = {}
    for metric in metric_names:
        values = [base.ratio_value(pair["ratios"][metric], f"repeat-{pair['repeat']}.{metric}") for pair in pairs]
        defined = [value for value in values if value is not None]
        metrics[metric] = {"screening_statistic": "median of three paired repeats", "pair_count": len(defined), "median": statistics.median(defined) if defined else None, "min": min(defined) if defined else None, "max": max(defined) if defined else None, "undefined_count": len(values) - len(defined), "review_flag_pair_gt_1_05": sum(value > SCREEN_SAFETY for value in defined)}
    return {"child_count": len(case_records), "pairs": pairs, "metrics": metrics, "screening_only": True, "full_followup_required": True}


def write_new(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o664)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
    except BaseException:
        try:
            os.close(fd)
        except OSError:
            pass
        raise


def reduce_screen(plan_path: Path, native_packet: Path, allocator_packet: Path, output: Path, helper_path: Path) -> dict[str, Any]:
    plan = load_json(plan_path)
    require(isinstance(plan, dict), "plan root is not an object")
    check_screen_plan(plan)
    validate_plan_bindings(plan, plan_path, helper_path)
    base.RUNNER_SCHEMA = plan["runner"]["schema"]
    base.DEADLINE = DEADLINE
    base.CPU_MASK = CPU_MASK
    base.MAX_SINK_WRITE = MAX_SINK_WRITE
    base.validate_native_report = validate_native_screen_report
    native_records = base.validate_lane_packet(native_packet.resolve(), plan_path.resolve(), plan, "native")
    allocator_records = base.validate_lane_packet(allocator_packet.resolve(), plan_path.resolve(), plan, "allocator")
    native_cases = {case["case_id"]: screen_native_case(native_records, plan, case["case_id"]) for case in plan["cases"]}
    allocator_cases = {case["case_id"]: screen_allocator_case(allocator_records, plan, case["case_id"]) for case in plan["cases"]}
    result = {
        "schema": "litchi.performance.0964.screen-reduction.v1",
        "status": "screened",
        "scope": "offline descriptive screening reduction; no adoption or production claim",
        "plan": descriptor(plan_path),
        "packets": {"native": descriptor(native_packet / "capture-terminal.json"), "allocator": descriptor(allocator_packet / "capture-terminal.json")},
        "schedule": {"native_children": len(native_records), "allocator_children": len(allocator_records), "native_blocks_per_case": 3, "native_samples": 60, "native_warmups": 10, "allocator_repeats_per_case": 3, "allocator_rows": 10, "allocator_warmups": 2},
        "screening_policy": {"adoption_allowed": False, "full_followup_required": True, "safety_review_ratio": SCREEN_SAFETY, "native_statistic": "median of six paired screening legs", "allocator_statistic": "median of three paired screening repeats"},
        "native": {"cases": native_cases},
        "allocator": {"cases": allocator_cases},
        "limitations": ["Screening has three native blocks and 60 samples, so it cannot qualify a full six-block/200-sample result.", "Ratios and review flags are descriptive; no bootstrap or adoption decision is emitted.", "Allocator max RSS remains whole-process evidence and is not a phase peak."],
    }
    write_new(output, result)
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--native-packet", type=Path, required=True)
    parser.add_argument("--allocator-packet", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--helper", type=Path, default=Path(__file__))
    args = parser.parse_args(argv)
    try:
        result = reduce_screen(args.plan.resolve(), args.native_packet.resolve(), args.allocator_packet.resolve(), args.output.resolve(), args.helper.resolve())
        print(json.dumps({"status": result["status"], "output": str(args.output.resolve())}, sort_keys=True))
        return 0
    except (ReductionError, OSError, ValueError) as exc:
        print(f"reduce_0964_core_v1: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
