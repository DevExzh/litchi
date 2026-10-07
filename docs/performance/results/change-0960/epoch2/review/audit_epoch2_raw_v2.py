#!/usr/bin/env python3
"""Independent epoch2 custody/oracle/median audit.

Reads only the final plan, sealed packets, raw child reports, validation
metadata, and existing reduction/selection outputs for comparison. It does
not import or call the frozen reducer for arithmetic.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import statistics
import sys
from pathlib import Path
from typing import Any

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("root", nargs="?", default="/tmp/litchi-performance-0960-epoch2", help="epoch root; retained v1 positional binding")
    parser.add_argument("--output", type=Path, default=None, help="optional output JSON path; default preserves the v1 output path")
    return parser.parse_args()

ARGS = parse_args()
ROOT = Path(ARGS.root).resolve()
OUTPUT_PATH = ARGS.output.expanduser().resolve() if ARGS.output is not None else None
PLAN_PATH = ROOT / "plan-final-v1.json"
NATIVE_PACKET = ROOT / "capture-native-v1"
ALLOCATOR_PACKET = ROOT / "capture-allocator-v1"
REDUCED_PATH = ROOT / "analysis-v1/screen-v1.json"
SELECTION_PATH = ROOT / "analysis-v1/selection-v1.json"
CPU_MASK = [12, 13, 14, 15]
RUNNER_SCHEMA = "litchi.performance.0960.capture.v1"
NATIVE_BLOCKS = [["A", "B", "B", "A"], ["B", "A", "A", "B"], ["A", "B", "B", "A"]]
ALLOCATOR_BLOCKS = [["A", "B"], ["B", "A"], ["A", "B"]]
NATIVE_PHASES = ["ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns"]
ALLOC_PHASES = ["setup", "ingress", "edit_commit", "publication", "commit_drop", "worker_drain", "core_combined", "core_combined_with_drain", "postdrop"]
ALLOC_FIELDS = ["allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes"]
ALLOC_METRICS = ["allocated_bytes", "allocation_calls", "region_peak_live_bytes", "core_combined_with_drain.allocated_bytes", "max_rss_kib"]
PRIMARY = ["edit_commit_ns.p50", "edit_commit_ns.p95", "total_ns.p50", "total_ns.p95"]

class AuditError(RuntimeError):
    pass

def require(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)

def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    return json.loads(path.read_text(encoding="utf-8"))

def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()

def descriptor(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing regular artifact: {path}")
    st = path.stat()
    return {"path": str(path.resolve()), "bytes": st.st_size, "mode": st.st_mode & 0o777, "executable": bool(st.st_mode & 0o111), "sha256": digest(path)}

def check_descriptor(value: Any, path: Path, label: str, executable: bool | None = None) -> None:
    actual = descriptor(path)
    require(isinstance(value, dict), f"{label}: descriptor missing")
    for key in ("path", "bytes", "mode", "executable", "sha256"):
        if key in value:
            require(value[key] == actual[key], f"{label}.{key} differs")
    if executable is not None:
        require(actual["executable"] is executable, f"{label}: executable mismatch")

def finite_number(value: Any, label: str, positive: bool = False) -> float:
    require(type(value) in (int, float) and not isinstance(value, bool), f"{label}: not numeric")
    x = float(value)
    require(math.isfinite(x) and (x > 0 if positive else x >= 0), f"{label}: invalid numeric value")
    return x

def nonnegative_int(value: Any, label: str) -> int:
    require(type(value) is int and value >= 0, f"{label}: invalid integer")
    return value

def expected_rows(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    blocks = NATIVE_BLOCKS if lane == "native" else ALLOCATOR_BLOCKS
    spec = plan["lanes"][lane]
    rows = []
    for case in plan["cases"]:
        for block, order in enumerate(blocks):
            for position, role in enumerate(order):
                rows.append({
                    "binary_role": "baseline" if role == "A" else "candidate",
                    "block": block,
                    "case_id": case["case_id"],
                    "child_id": f"{lane}__{case['case_id']}__b{block:02d}__p{position}",
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

def nearest(values: list[float], percentile: int) -> float:
    ordered = sorted(values)
    require(ordered, "empty percentile input")
    if percentile == 50:
        return float(statistics.median(ordered))
    index = max(1, math.ceil(percentile * len(ordered) / 100)) - 1
    return float(ordered[min(index, len(ordered) - 1)])

def stats(values: list[float]) -> dict[str, Any]:
    require(values, "empty statistic input")
    return {"p50": nearest(values, 50), "p95": nearest(values, 95), "p99": nearest(values, 99), "min": float(min(values)), "max": float(max(values)), "count": len(values)}

def ratio(numerator: float, denominator: float) -> dict[str, Any]:
    finite_number(numerator, "ratio numerator")
    finite_number(denominator, "ratio denominator")
    if denominator == 0:
        return {"status": "undefined", "reason": "zero_denominator", "numerator": numerator, "denominator": denominator}
    value = numerator / denominator
    require(math.isfinite(value) and value >= 0, "ratio invalid")
    return {"status": "defined", "value": value, "numerator": numerator, "denominator": denominator}

def ratio_value(value: dict[str, Any]) -> float | None:
    if value.get("status") == "undefined":
        require(value.get("reason") == "zero_denominator", "unknown undefined ratio")
        return None
    require(value.get("status") == "defined", "invalid ratio status")
    return finite_number(value.get("value"), "ratio value")

def metric_summary(values: list[float]) -> dict[str, Any]:
    defined = [x for x in values if x is not None]
    return {"screening_statistic": "median of six paired screening legs", "pair_count": len(defined), "median": float(statistics.median(defined)) if defined else None, "min": float(min(defined)) if defined else None, "max": float(max(defined)) if defined else None, "undefined_count": len(values) - len(defined), "review_flag_pair_gt_1_05": sum(x > 1.05 for x in defined)}

def common_case(report_case: dict[str, Any], row: dict[str, Any], expected: dict[str, Any]) -> None:
    wanted = {
        "id": f"{expected["shape"]}-{expected["workload"]}-workers1", "shape": expected["shape"], "workload": expected["workload"], "workers": 1,
        "update_count": expected["update_count"], "source_bytes": expected["source_bytes"], "source_sha256": expected["source_sha256"],
        "output_bytes": expected["output_bytes"], "output_sha256": expected["output_sha256"], "semantic_sha256": expected["semantic_sha256"],
        "expected_cpu_tasks": expected["expected_cpu_tasks"], "admission_expected_worker_permits": expected["expected_worker_permits"],
        "admission_observed_worker_permits": 0, "admission_cpu_tasks_used": expected["expected_cpu_tasks"],
        "preflight_actual_commit_inverse_verified": True, "context_limits": expected["context_limits"],
        "sink_capacity_bytes": expected["context_limits"]["output_bytes"], "sink_max_write_bytes": 65536,
        "warmup_iterations": row["warmup_iterations"], "measured_samples": row["measured_samples"],
    }
    for key, value in wanted.items():
        require(report_case.get(key) == value, f"{row['child_id']}: case oracle {key} differs")

def check_time(validation: dict[str, Any], child: str) -> int:
    metrics = validation.get("time_metrics")
    require(isinstance(metrics, dict), f"{child}: time metrics missing")
    rss = metrics.get("max_rss_kib")
    nonnegative_int(rss, f"{child}.max_rss_kib")
    require(rss > 0, f"{child}.max_rss_kib not positive")
    return rss

def check_capture(plan: dict[str, Any], lane: str, packet: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    rows = expected_rows(plan, lane)
    lane_plan = plan["lanes"][lane]
    start_path = packet / "capture-start.json"
    terminal_path = packet / "capture-terminal.json"
    completed_path = packet / "completed-packet.json"
    start = read(start_path)
    require(start.get("schema") == RUNNER_SCHEMA + ".capture-start", f"{lane}: capture-start schema")
    require(start.get("lane") == lane and start.get("serial") is True, f"{lane}: capture start schedule")
    require(start.get("children_expected") == len(rows) and start.get("retry_or_replacement") is False, f"{lane}: capture start counts/policy")
    require(start.get("plan", {}).get("sha256") == digest(PLAN_PATH), f"{lane}: plan binding")
    check_descriptor(start.get("plan"), PLAN_PATH, f"{lane}.capture-start.plan")
    check_descriptor(start.get("runner"), Path(plan["helpers"]["capture"]["path"]), f"{lane}.capture-start.runner", executable=True)
    terminal = read(terminal_path)
    require(terminal.get("schema") == RUNNER_SCHEMA + ".capture-terminal", f"{lane}: terminal schema")
    require(terminal.get("status") == "completed" and terminal.get("children_started") == len(rows) and terminal.get("children_expected") == len(rows) and terminal.get("children_completed") == len(rows), f"{lane}: terminal counts/status")
    require(terminal.get("terminal_before_parse") is True and terminal.get("retry_or_replacement") is False, f"{lane}: terminal policy")
    check_descriptor(terminal.get("capture_start"), start_path, f"{lane}.terminal.capture_start")
    completed = read(completed_path)
    require(completed.get("schema") == RUNNER_SCHEMA + ".completed-packet" and completed.get("status") == "completed", f"{lane}: completed packet status")
    require([x.get("child_id") for x in completed.get("children", [])] == [x["child_id"] for x in rows], f"{lane}: completed child order")
    check_descriptor(completed.get("capture_terminal"), terminal_path, f"{lane}.completed.capture_terminal")
    child_dirs = sorted(p.name for p in (packet / "children").iterdir() if p.is_dir())
    require(child_dirs == sorted(x["child_id"] for x in rows), f"{lane}: child directory set differs")
    records = []
    for row in rows:
        child = packet / "children" / row["child_id"]
        for name in ("start.json", "launch.json", "terminal.json", "validation.json", "post-terminal-artifacts.json"):
            require((child / name).is_file() and not (child / name).is_symlink(), f"{row['child_id']}: missing {name}")
        start_child = read(child / "start.json")
        require(start_child.get("schema") == RUNNER_SCHEMA + ".child-start" and start_child.get("child") == row, f"{row['child_id']}: child schedule")
        require(start_child.get("cwd") == lane_plan["execution"]["workload_cwd"] and start_child.get("deadline_seconds") == 900 and start_child.get("output_absent_before_launch") is True, f"{row['child_id']}: cwd/deadline/output custody")
        expected_env = {k: str(v) for k, v in lane_plan["execution"]["environment"].items() if k != "approval_variable"}
        expected_env[lane_plan["execution"]["environment"]["approval_variable"]] = "1"
        require(start_child.get("environment") == expected_env, f"{row['child_id']}: environment")
        binary_key = "baseline_binary" if row["role"] == "A" else "candidate_binary"
        binary = lane_plan[binary_key]
        binary_path = Path(binary["path"]).resolve()
        check_descriptor(start_child.get("binary"), binary_path, f"{row['child_id']}.binary", executable=True)
        output_path = child / "output"
        time_path = child / "time-v.txt"
        expected_workload = ["/usr/bin/taskset", "--cpu-list", "12-15", str(binary_path), "--shape", row["shape"], "--workload", row["workload"], "--workers", "1", "--warmup", str(row["warmup_iterations"]), "--samples", str(row["measured_samples"]), "--output", str(output_path)]
        expected_argv = ["/usr/bin/taskset", "--cpu-list", "12-15", "/usr/bin/time", "-v", "-o", str(time_path), "--", str(binary_path), "--shape", row["shape"], "--workload", row["workload"], "--workers", "1", "--warmup", str(row["warmup_iterations"]), "--samples", str(row["measured_samples"]), "--output", str(output_path)]
        require(start_child.get("workload_argv") == expected_workload and start_child.get("argv") == expected_argv, f"{row['child_id']}: argv custody")
        launch = read(child / "launch.json")
        require(launch.get("schema") == RUNNER_SCHEMA + ".child-launch" and launch.get("observed_affinity") == CPU_MASK, f"{row['child_id']}: affinity")
        terminal_child = read(child / "terminal.json")
        require(terminal_child.get("schema") == RUNNER_SCHEMA + ".child-terminal" and terminal_child.get("status") == "completed" and terminal_child.get("return_code") == 0 and terminal_child.get("timed_out") is False and terminal_child.get("process_group_clear") is True and terminal_child.get("terminal_before_report_parse") is True, f"{row['child_id']}: terminal custody")
        run = terminal_child.get("run", {})
        require(run.get("reaped") is True and run.get("process_group_clear") is True, f"{row['child_id']}: process cleanup")
        validation = read(child / "validation.json")
        require(validation.get("schema") == RUNNER_SCHEMA + ".child-validation" and validation.get("status") == "passed" and validation.get("child") == row, f"{row['child_id']}: validation")
        check_descriptor(validation.get("terminal"), child / "terminal.json", f"{row['child_id']}.validation.terminal")
        report_path = output_path / "report.json"
        check_descriptor(validation.get("report"), report_path, f"{row['child_id']}.validation.report")
        check_descriptor(validation.get("time"), time_path, f"{row['child_id']}.validation.time")
        rss = check_time(validation, row["child_id"])
        post = read(child / "post-terminal-artifacts.json")
        require(post.get("schema") == RUNNER_SCHEMA + ".post-terminal-artifacts", f"{row['child_id']}: post schema")
        for key, path in (("terminal", child / "terminal.json"), ("stdout", child / "stdout.log"), ("stderr", child / "stderr.log"), ("time", time_path), ("report", report_path), ("validation", child / "validation.json")):
            check_descriptor(post.get(key), path, f"{row['child_id']}.post.{key}")
        records.append({"row": row, "child": row["child_id"], "child_path": child, "report": read(report_path), "validation": validation, "max_rss_kib": rss})
    return records, {"start": descriptor(start_path), "terminal": descriptor(terminal_path), "completed": descriptor(completed_path)}

def check_native_report(record: dict[str, Any], expected: dict[str, Any]) -> dict[str, Any]:
    row, report = record["row"], record["report"]
    require(report.get("schema") == "litchi.xlsx-managed-compression-benchmark.v1" and report.get("status") == "completed", f"{row['child_id']}: native report status/schema")
    common_case(report.get("case", {}), row, expected)
    contract = json.loads(PLAN_PATH.read_text())["report_contract"]["native"]
    for scope_key in ("timing_scope", "phase_scope", "setup_scope", "verification_scope"):
        require(isinstance(report.get(scope_key), str) and report[scope_key], f"{row['child_id']}: native report scope {scope_key} missing")
    require(all(phase in report["phase_scope"] for phase in ("ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns")), f"{row['child_id']}: native phase scope incomplete")
    require(report.get("performance_claim") == contract["performance_claim"], f"{row['child_id']}: native report scope performance_claim")
    case = report["case"]
    samples = case.get("samples")
    require(isinstance(samples, list) and len(samples) == 60, f"{row['child_id']}: native sample count")
    for i, sample in enumerate(samples):
        require(sample.get("sample") == i, f"{row['child_id']}: native sample index")
        values = [nonnegative_int(sample.get(phase), f"{row['child_id']}.{phase}") for phase in NATIVE_PHASES]
        require(sum(values[:4]) <= values[-1], f"{row['child_id']}: native phase total")
        require(sample.get("output_bytes") == expected["output_bytes"] and sample.get("output_bytes_used") == expected["output_bytes"], f"{row['child_id']}: output bytes oracle")
        require(sample.get("cpu_tasks_used") == expected["expected_cpu_tasks"], f"{row['child_id']}: CPU task oracle")
        require(type(sample.get("sink_write_calls")) is int and sample["sink_write_calls"] > 0 and 0 < sample.get("sink_largest_write", 0) <= 65536, f"{row['child_id']}: sink oracle")
        require(sample.get("output_sha256") == expected["output_sha256"] and sample.get("semantic_sha256") == expected["semantic_sha256"] and sample.get("fixture_inverse_restored_exact_source") is True, f"{row['child_id']}: output/semantic/inverse oracle")
        resources = sample.get("resources_after_drop")
        require(isinstance(resources, dict) and all(resources.get(k) == 0 for k in ("memory", "workers", "io_concurrency", "objects")), f"{row['child_id']}: native resources")
    metrics = {}
    for phase in NATIVE_PHASES:
        s = stats([float(sample[phase]) for sample in samples])
        for q in ("p50", "p95", "p99"):
            metrics[f"{phase}.{q}"] = s[q]
    metrics["max_rss_kib"] = float(record["max_rss_kib"])
    return metrics

def check_alloc_phase(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict) and value.get("status") == "measured" and value.get("scope") == "operation_global_system_allocator", f"{label}: phase status/scope")
    for field in ALLOC_FIELDS:
        nonnegative_int(value.get(field), f"{label}.{field}")
    require(value["failed_allocation_calls"] == 0 and value["region_peak_live_bytes"] >= max(value["live_bytes_before"], value["live_bytes_after"]), f"{label}: phase safety")
    return value

def combine(phases: list[dict[str, Any]], label: str) -> dict[str, Any]:
    require(phases, f"{label}: empty")
    for left, right in zip(phases, phases[1:]):
        require(left["live_bytes_after"] == right["live_bytes_before"], f"{label}: live boundary")
    result = {"status": "measured", "scope": "operation_global_system_allocator", "live_bytes_before": phases[0]["live_bytes_before"], "live_bytes_after": phases[-1]["live_bytes_after"], "peak_live_bytes_before": phases[0]["peak_live_bytes_before"], "peak_live_bytes_after": phases[-1]["peak_live_bytes_after"], "region_peak_live_bytes": max(x["region_peak_live_bytes"] for x in phases)}
    for field in ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes"):
        result[field] = sum(x[field] for x in phases)
    return result

def check_allocator_report(record: dict[str, Any], expected: dict[str, Any], plan: dict[str, Any]) -> dict[str, float]:
    row, report = record["row"], record["report"]
    require(report.get("schema") == "litchi.xlsx-managed-compression-allocation.v1" and report.get("status") == "completed", f"{row['child_id']}: allocator report status/schema")
    common_case(report.get("case", {}), row, expected)
    contract = json.loads(PLAN_PATH.read_text())["report_contract"]["allocator"]
    for key in ("allocation_scope", "allocator_identity", "instrumentation_identity", "counter_revision", "performance_claim"):
        require(report.get(key) == contract[key], f"{row['child_id']}: allocator report scope {key}")
    require(report.get("timing_scope") == "none; this report contains allocator regions and diagnostic quiescence observations only", f"{row['child_id']}: allocator timing scope")
    expected_binary = plan["lanes"]["allocator"]["baseline_binary" if row["role"] == "A" else "candidate_binary"]
    identity = report.get("binary_identity")
    require(isinstance(identity, dict) and identity.get("path") == expected_binary["path"] and identity.get("bytes") == expected_binary["bytes"] and identity.get("sha256") == expected_binary["sha256"], f"{row['child_id']}: allocator binary identity")
    calibration = report.get("calibration")
    require(isinstance(calibration, dict) and calibration.get("status") == "pass" and calibration.get("polls") == 100 and calibration.get("thread_count_first") == 1 and calibration.get("thread_count_last") == 1 and calibration.get("allocation_calls") == 0 and calibration.get("deallocation_calls") == 0 and calibration.get("reallocation_calls") == 0 and calibration.get("failed_allocation_calls") == 0 and calibration.get("no_allocator_callbacks") is True, f"{row['child_id']}: allocator calibration")
    case = report["case"]
    rows = case.get("rows")
    require(isinstance(rows, list) and len(rows) == 10, f"{row['child_id']}: allocator row count")
    for i, item in enumerate(rows):
        require(item.get("sample") == i, f"{row['child_id']}: allocator row index")
        wanted = {"context_limits": expected["context_limits"], "workers_requested": 1, "expected_worker_permits": 0, "admitted_worker_permits": 0, "expected_cpu_tasks": expected["expected_cpu_tasks"], "cpu_tasks_used": expected["expected_cpu_tasks"], "output_bytes": expected["output_bytes"], "output_bytes_used": expected["output_bytes"], "source_sha256": expected["source_sha256"], "output_sha256": expected["output_sha256"], "semantic_sha256": expected["semantic_sha256"], "fixture_inverse_restored_exact_source": True}
        for key, value in wanted.items(): require(item.get(key) == value, f"{row['child_id']}: allocator row {i} oracle {key}")
        require(type(item.get("sink_write_calls")) is int and item["sink_write_calls"] > 0 and 0 < item.get("sink_largest_write", 0) <= 65536, f"{row['child_id']}: allocator sink oracle")
        resources = item.get("resources_after_core")
        require(isinstance(resources, dict) and all(resources.get(k) == 0 for k in ("memory", "workers", "io_concurrency", "objects")), f"{row['child_id']}: allocator resources")
        phases = [check_alloc_phase(item.get(name), f"{row['child_id']}.row[{i}].{name}") for name in ALLOC_PHASES]
        require(item["core_combined"] == combine(phases[1:5], f"{row['child_id']}.row[{i}].core"), f"{row['child_id']}: core conservation")
        require(item["core_combined_with_drain"] == combine(phases[1:6], f"{row['child_id']}.row[{i}].core+drain"), f"{row['child_id']}: drained core conservation")
        nonnegative_int(item.get("max_setup_core_peak_above_setup_live_before"), f"{row['child_id']}: setup peak diagnostic")
        nonnegative_int(item.get("max_setup_drained_core_peak_above_setup_live_before"), f"{row['child_id']}: drained setup peak diagnostic")
    core = [item["core_combined_with_drain"] for item in rows]
    return {
        "allocated_bytes": float(statistics.fmean(item["allocated_bytes"] for item in core)),
        "allocation_calls": float(statistics.fmean(item["allocation_calls"] for item in core)),
        "region_peak_live_bytes": float(max(item["region_peak_live_bytes"] for item in core)),
        "core_combined_with_drain.allocated_bytes": float(statistics.fmean(item["allocated_bytes"] for item in core)),
        "max_rss_kib": float(record["max_rss_kib"]),
    }

def pair_records(records: list[dict[str, Any]], lane: str, case_id: str, metrics_by_child: dict[str, dict[str, float]]) -> list[dict[str, Any]]:
    blocks = NATIVE_BLOCKS if lane == "native" else ALLOCATOR_BLOCKS
    selected = [r for r in records if r["row"]["case_id"] == case_id]
    require(len(selected) == (12 if lane == "native" else 6), f"{lane}/{case_id}: child count")
    by_block = {}
    for r in selected: by_block.setdefault(r["row"]["block"], []).append(r)
    pairs = []
    for block_id, order in enumerate(blocks):
        block = sorted(by_block[block_id], key=lambda r: r["row"]["position"])
        require([r["row"]["role"] for r in block] == order, f"{lane}/{case_id}/b{block_id}: role order")
        legs = [(block[0], block[1]), (block[2], block[3])] if lane == "native" else [(block[0], block[1])]
        for leg, (left, right) in enumerate(legs):
            a, b = (left, right) if left["row"]["role"] == "A" else (right, left)
            ratios = {metric: ratio(metrics_by_child[b["child"]][metric], metrics_by_child[a["child"]][metric]) for metric in metrics_by_child[b["child"]]}
            pairs.append({"pair_id": f"b{block_id:02d}-leg{leg}" if lane == "native" else f"repeat-{block_id}", "block": block_id, "baseline_child": a["child"], "candidate_child": b["child"], "ratios": ratios})
    return pairs

def summarize_pairs(pairs: list[dict[str, Any]], metric_names: list[str], statistic: str) -> dict[str, Any]:
    result = {}
    for metric in metric_names:
        vals = [ratio_value(p["ratios"][metric]) for p in pairs]
        defined = [v for v in vals if v is not None]
        result[metric] = {"screening_statistic": statistic, "pair_count": len(defined), "median": float(statistics.median(defined)) if defined else None, "min": float(min(defined)) if defined else None, "max": float(max(defined)) if defined else None, "undefined_count": len(vals) - len(defined), "review_flag_pair_gt_1_05": sum(v > 1.05 for v in defined), "pair_values": vals}
    return result

def close(a: Any, b: Any) -> bool:
    if isinstance(a, float) or isinstance(b, float):
        try: return math.isclose(float(a), float(b), rel_tol=1e-14, abs_tol=1e-15)
        except (TypeError, ValueError): return False
    return a == b

def compare_reduction(own: dict[str, Any], reduced: dict[str, Any]) -> list[dict[str, Any]]:
    mismatches=[]
    for lane in ("native", "allocator"):
        for case, value in own[lane]["cases"].items():
            expected = reduced[lane]["cases"][case]["metrics"]
            for metric, summary in value["metrics"].items():
                for key in ("screening_statistic", "pair_count", "median", "min", "max", "undefined_count", "review_flag_pair_gt_1_05"):
                    if not close(summary[key], expected[metric].get(key)):
                        mismatches.append({"lane":lane,"case":case,"metric":metric,"field":key,"own":summary[key],"reducer":expected[metric].get(key)})
    return mismatches

def main() -> int:
    plan=read(PLAN_PATH)
    require(plan.get("status") == "authorized" and plan.get("measurements_started") is False, "final plan not authorized or measurements marker changed")
    require(plan.get("cpu_mask") == CPU_MASK, "plan CPU mask changed")
    all_records={}
    custody={}
    for lane, packet in (("native",NATIVE_PACKET),("allocator",ALLOCATOR_PACKET)):
        all_records[lane], custody[lane]=check_capture(plan,lane,packet)
    expected={c["case_id"]:c for c in plan["cases"]}
    native_metric_children={}; allocator_metric_children={}
    native_oracle_children=allocator_oracle_children=0
    for r in all_records["native"]:
        native_metric_children[r["child"]]=check_native_report(r,expected[r["row"]["case_id"]]); native_oracle_children+=1
    for r in all_records["allocator"]:
        allocator_metric_children[r["child"]]=check_allocator_report(r,expected[r["row"]["case_id"]],plan); allocator_oracle_children+=1
    own={"native":{"cases":{}},"allocator":{"cases":{}}}
    for case in expected:
        pairs=pair_records(all_records["native"],"native",case,native_metric_children)
        metrics=list(native_metric_children[pairs[0]["baseline_child"]].keys())
        own["native"]["cases"][case]={"child_count":12,"pair_count":len(pairs),"pairs":pairs,"metrics":summarize_pairs(pairs,metrics,"median of six paired screening legs")}
        pairs=pair_records(all_records["allocator"],"allocator",case,allocator_metric_children)
        own["allocator"]["cases"][case]={"child_count":6,"pair_count":len(pairs),"pairs":pairs,"metrics":summarize_pairs(pairs,ALLOC_METRICS,"median of three paired repeats")}
    reduced=read(REDUCED_PATH)
    mismatches=compare_reduction(own,reduced)
    signals=[]; stops=[]
    for lane in ("native","allocator"):
        for case, value in own[lane]["cases"].items():
            required=PRIMARY+["max_rss_kib"] if lane=="native" else ALLOC_METRICS
            for metric in required:
                m=value["metrics"][metric]; valid=m["pair_count"]==(6 if lane=="native" else 3) and m["undefined_count"]==0 and isinstance(m["median"],(int,float)) and math.isfinite(m["median"]) and m["median"]>0
                if not valid: stops.append(f"{lane}/{case}/{metric}: invalid or undefined")
                if lane=="native" and metric in PRIMARY and valid and m["median"]>1.05: stops.append(f"{lane}/{case}/{metric}: primary median exceeds 1.05")
                if (lane=="native" or metric=="region_peak_live_bytes") and valid and m["median"]<=0.95: signals.append({"metric":f"{lane}/{case}/{metric}","median":m["median"]})
    if not signals: stops.append("No qualifying metric among signal_any")
    selection={"decision":"reject_screen" if stops else "require_full_confirmation_and_regression_review","signals":signals,"stops":stops,"adoption_allowed":False,"statistical_effect_claim":False}
    existing_selection=read(SELECTION_PATH)
    require(existing_selection.get("decision")==selection["decision"], "selector decision differs")
    require(existing_selection.get("signals")==selection["signals"], "selector signal list differs")
    require(existing_selection.get("stops")==selection["stops"], "selector stop list differs")
    result={
      "schema":"litchi.performance.0960.epoch2-independent-raw-audit.v1",
      "status":"passed" if not mismatches else "failed",
      "scope":"independent raw packet custody, oracle, allocator conservation, and descriptive median audit; no adoption claim",
      "plan":descriptor(PLAN_PATH),
      "packets":custody,
      "custody_summary":{"native_children":len(all_records["native"]),"allocator_children":len(all_records["allocator"]),"native_oracles_passed":native_oracle_children,"allocator_oracles_passed":allocator_oracle_children,"no_retries_or_replacements":True,"pooled_samples":False,"source_input_oracles_passed":True,"process_group_cleanup_passed":True},
      "native":{"cases":own["native"]["cases"],"metrics_included":"all five phases at p50/p95/p99 plus whole-process max_rss_kib; review flags retained"},
      "allocator":{"cases":own["allocator"]["cases"],"metric_semantics":{"allocated_bytes":"mean core_combined_with_drain traffic bytes over ten rows","allocation_calls":"mean core_combined_with_drain allocation-call counter over ten rows","core_combined_with_drain.allocated_bytes":"same traffic-byte aggregation as allocated_bytes","region_peak_live_bytes":"maximum observer-ordered live high-water over ten core_combined_with_drain rows; not traffic","max_rss_kib":"whole-process /usr/bin/time -v maximum RSS; not allocator phase peak"}},
      "reduction_comparison":{"reducer_output":descriptor(REDUCED_PATH),"mismatch_count":len(mismatches),"mismatches":mismatches},
      "selection_comparison":{"selection_output":descriptor(SELECTION_PATH),"independent":selection,"matched":True},
      "verdict":"screen rejected by the bound rule because no qualifying signal was present; no full confirmation is indicated; no adoption or statistical effect claim",
    }
    out=OUTPUT_PATH if OUTPUT_PATH is not None else ROOT/"review/independent-raw-audit-v1.json"
    require(not out.exists(), f"output already exists: {out}")
    out.write_text(json.dumps(result,indent=2,sort_keys=True)+"\n",encoding="utf-8")
    print(json.dumps({"status":result["status"],"output":str(out),"output_sha256":digest(out),"mismatch_count":len(mismatches),"decision":selection["decision"]},sort_keys=True))
    return 0 if not mismatches else 2

if __name__ == "__main__":
    try: raise SystemExit(main())
    except AuditError as exc:
        print(f"audit_epoch2_raw_v1: {exc}", file=sys.stderr)
        raise SystemExit(2)
