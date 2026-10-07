#!/usr/bin/env python3
"""Independent epoch2 0964 raw custody/oracle/median audit.

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
    parser.add_argument("root", nargs="?", default="/home/zhuhe/litchi-performance-0964", help="0964 evidence root")
    parser.add_argument("--output", type=Path, default=None, help="optional output JSON path; defaults below the evidence root")
    return parser.parse_args()

ARGS = parse_args()
ROOT = Path(ARGS.root).resolve()
OUTPUT_PATH = ARGS.output.expanduser().resolve() if ARGS.output is not None else None
PLAN_PATH = ROOT / "protocol/plan-final-epoch2-v1.json"
NATIVE_PACKET = ROOT / "capture-native-epoch2-v1"
ALLOCATOR_PACKET = ROOT / "capture-allocator-epoch2-v1"
REDUCED_PATH = ROOT / "analysis-epoch2/screen-v1.json"
SELECTION_PATH = ROOT / "analysis-epoch2/selection-v1.json"
CPU_MASK = [12, 13, 14, 15]
RUNNER_SCHEMA = "litchi.performance.0964.capture.v1"
NATIVE_BLOCKS = [["A", "B", "B", "A"], ["B", "A", "A", "B"], ["A", "B", "B", "A"]]
ALLOCATOR_BLOCKS = [["A", "B"], ["B", "A"], ["A", "B"]]
NATIVE_PHASES = ["ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns"]
ALLOC_PHASES = ["setup", "ingress", "edit_commit", "publication", "commit_drop", "worker_drain", "core_combined", "core_combined_with_drain", "postdrop"]
ALLOC_FIELDS = ["allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes"]
ALLOC_METRICS = ["allocated_bytes", "allocation_calls", "region_peak_live_bytes", "core_combined_with_drain.allocated_bytes", "max_rss_kib"]
PRIMARY = ["edit_commit_ns.p50", "edit_commit_ns.p95", "total_ns.p50", "total_ns.p95"]
TAIL = ["ingress_ns.p99", "edit_commit_ns.p99", "publication_ns.p99", "commit_drop_ns.p99", "total_ns.p99"]
NATIVE_REQUIRED = PRIMARY + TAIL + ["max_rss_kib"]
SIGNAL_THRESHOLD = 0.95
SAFETY_THRESHOLD = 1.05

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

def plan_contract(plan_path: Path, lane: str) -> dict[str, Any]:
    value = read(plan_path)
    contract = value.get("report_contract", {}).get(lane)
    require(isinstance(contract, dict), f"report contract missing for {lane}")
    return contract

def check_plan_descriptor(value: Any, label: str, executable: bool | None = None) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    path_value = value.get("path")
    require(isinstance(path_value, str) and path_value, f"{label}: descriptor path missing")
    path = Path(path_value)
    check_descriptor(value, path, label, executable)
    return path.resolve()

def manifest_entries(path: Path, label: str) -> dict[str, tuple[int, int, str]]:
    value = read(path)
    require(value.get("schema", "").startswith("litchi.performance.0964.source-manifest."), f"{label}: manifest schema")
    files = value.get("files")
    require(isinstance(files, list) and value.get("file_count") == len(files), f"{label}: manifest file census")
    result: dict[str, tuple[int, int, str]] = {}
    for item in files:
        require(isinstance(item, dict), f"{label}: manifest entry malformed")
        name = item.get("path")
        require(isinstance(name, str) and name and not Path(name).is_absolute() and ".." not in Path(name).parts and "\\" not in name, f"{label}: unsafe manifest path")
        require(name not in result, f"{label}: duplicate manifest path {name}")
        size, mode, sha = item.get("bytes"), item.get("mode"), item.get("sha256")
        require(type(size) is int and size >= 0 and type(mode) is int and mode >= 0 and isinstance(sha, str) and len(sha) == 64, f"{label}: invalid manifest metadata {name}")
        result[name] = (size, mode, sha)
    return result

def check_source_binding(plan: dict[str, Any]) -> dict[str, Any]:
    source = plan.get("source")
    require(isinstance(source, dict), "plan source binding missing")
    descriptor_keys = ("baseline_manifest", "baseline_runtime_manifest", "baseline_receipt", "candidate_allowlist_file", "candidate_freeze", "candidate_manifest", "mode_correction", "source_export", "source_start")
    paths: dict[str, Path] = {}
    for key in descriptor_keys:
        paths[key] = check_plan_descriptor(source.get(key), f"source.{key}")
    baseline = manifest_entries(paths["baseline_manifest"], "baseline manifest")
    candidate = manifest_entries(paths["candidate_manifest"], "candidate manifest")
    changed = sorted(name for name in set(baseline) | set(candidate) if baseline.get(name) != candidate.get(name))
    allowlist = source.get("candidate_change_allowlist")
    require(isinstance(allowlist, list) and all(isinstance(name, str) for name in allowlist), "source candidate allowlist missing")
    require(changed == sorted(allowlist), f"source manifest diff changed paths: {changed}")
    allowlist_value = read(paths["candidate_allowlist_file"])
    require(allowlist_value.get("status") == "bound_to_current_candidate_source_snapshot", "candidate test allowlist status")
    require(allowlist_value.get("paths") == sorted([name for name in allowlist if name.endswith("_tests.rs")]), "candidate test allowlist paths")
    require(allowlist_value.get("source_manifest", {}).get("sha256") == digest(paths["candidate_manifest"]), "candidate test allowlist manifest binding")
    freeze = read(paths["candidate_freeze"])
    require(freeze.get("status") == "pass" and freeze.get("build_or_workload_launched") is False, "candidate freeze status")
    require(sorted(freeze.get("changed_paths", [])) == sorted(allowlist), "candidate freeze changed paths")
    require(sorted(freeze.get("production_paths", [])) == sorted([name for name in allowlist if not name.endswith("_tests.rs")]), "candidate freeze production paths")
    return {"baseline_manifest": descriptor(paths["baseline_manifest"]), "candidate_manifest": descriptor(paths["candidate_manifest"]), "changed_paths": changed, "allowlist": sorted(allowlist), "test_allowlist": descriptor(paths["candidate_allowlist_file"])}

def check_root_go_binding(plan: dict[str, Any]) -> dict[str, Any]:
    root_go = plan.get("root_go")
    require(isinstance(root_go, dict) and root_go.get("status") == "GO", "root GO binding missing")
    receipt_path = check_plan_descriptor(root_go.get("receipt"), "root_go.receipt")
    receipt = read(receipt_path)
    require(receipt.get("schema") is None or isinstance(receipt.get("schema"), str), "root GO receipt schema malformed")
    require(receipt.get("status") == "GO", "root GO receipt status changed")
    bindings = receipt.get("bindings")
    require(isinstance(bindings, list) and bindings, "root GO receipt bindings missing")
    checked = []
    for index, value in enumerate(bindings):
        require(isinstance(value, dict) and isinstance(value.get("path"), str), f"root_go.binding[{index}] malformed")
        path = Path(value["path"])
        actual = descriptor(path)
        require(value.get("bytes") == actual["bytes"] and value.get("sha256") == actual["sha256"], f"root_go.binding[{index}] differs")
        checked.append(actual)
    return {"receipt": descriptor(receipt_path), "bindings": checked}

def check_epoch2_affinity_binding(plan: dict[str, Any]) -> dict[str, Any]:
    hardening = plan.get("epoch2_affinity_hardening")
    require(isinstance(hardening, dict), "epoch2 affinity hardening binding missing")
    require(hardening.get("exact_process_affinity_getter") == "os.sched_getaffinity(pid)", "epoch2 affinity getter binding changed")
    require(hardening.get("explicit_taskset_argv_preserved") is True, "epoch2 taskset argv preservation binding changed")
    require(hardening.get("fresh_54_child_schedule_required") is True, "epoch2 fresh schedule binding changed")
    require(hardening.get("launch_mismatch") == "fail before wait/report parsing; no retry or replacement", "epoch2 mismatch policy changed")
    require(hardening.get("parent_preexec_affinity") == CPU_MASK, "epoch2 parent preexec affinity changed")
    require(hardening.get("prior_epoch_capture_status") == "invalid_due_to_launch_affinity", "epoch1 exclusion rationale changed")

    capture = check_plan_descriptor(plan.get("helpers", {}).get("capture"), "helpers.capture", executable=True)
    runner = check_plan_descriptor(hardening.get("runner"), "epoch2_affinity_hardening.runner", executable=True)
    require(runner == capture, "epoch2 affinity runner differs from capture helper")
    require(hardening.get("runner") == plan["helpers"]["capture"], "epoch2 affinity runner descriptor differs from helper binding")

    controls_path = check_plan_descriptor(hardening.get("controls"), "epoch2_affinity_hardening.controls")
    require(hardening.get("controls") == plan.get("helpers", {}).get("affinity_controls"), "epoch2 affinity control descriptor differs from helper binding")
    controls = read(controls_path)
    require(controls.get("schema") == "litchi.performance.0964.affinity-launch-controls.v1", "epoch2 affinity controls schema changed")
    require(controls.get("status") == "pass" and controls.get("workload_started") is False, "epoch2 affinity controls status/workload marker changed")
    require(controls.get("cpu_mask") == CPU_MASK, "epoch2 affinity controls CPU mask changed")
    expected_command = ["/usr/bin/taskset", "--cpu-list", "12-15", "/usr/bin/true"]
    require(controls.get("command") == expected_command, "epoch2 affinity control command changed")

    positive = controls.get("positive")
    require(isinstance(positive, dict), "epoch2 positive affinity control missing")
    require(positive.get("command") == expected_command and positive.get("requested_preexec_affinity") == CPU_MASK, "epoch2 positive control request changed")
    require(positive.get("observed_affinity") == CPU_MASK and positive.get("guard_passed") is True and positive.get("returncode") == 0, "epoch2 positive affinity control failed")
    positive_cleanup = positive.get("cleanup")
    require(isinstance(positive_cleanup, dict) and positive_cleanup.get("actions") == [] and positive_cleanup.get("reaped") is True and positive_cleanup.get("process_group_clear") is True, "epoch2 positive cleanup control failed")

    negative = controls.get("negative")
    require(isinstance(negative, dict), "epoch2 negative affinity control missing")
    require(negative.get("command") == expected_command and negative.get("requested_preexec_affinity") == [12], "epoch2 negative control request changed")
    require(negative.get("observed_affinity") != CPU_MASK and negative.get("guard_rejected") is True, "epoch2 negative mismatch guard did not reject")
    negative_cleanup = negative.get("cleanup")
    require(isinstance(negative_cleanup, dict) and negative_cleanup.get("reaped") is True and negative_cleanup.get("process_group_clear") is True, "epoch2 negative cleanup control failed")
    return {"runner": descriptor(runner), "controls": descriptor(controls_path), "status": controls["status"], "workload_started": controls["workload_started"]}

def check_descriptor(value: Any, path: Path, label: str, executable: bool | None = None) -> None:
    actual = descriptor(path)
    require(isinstance(value, dict), f"{label}: descriptor missing")
    require(all(key in value for key in ("path", "bytes", "mode", "executable", "sha256")), f"{label}: descriptor incomplete")
    for key in ("path", "bytes", "mode", "executable", "sha256"):
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
        "expected_cpu_tasks": expected["expected_cpu_tasks"], "admission_expected_worker_permits": expected.get("expected_worker_permits", 0),
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
        require(launch.get("schema") == RUNNER_SCHEMA + ".child-launch" and launch.get("observed_affinity") == CPU_MASK and launch.get("expected_affinity") == CPU_MASK and launch.get("affinity_verified_before_wait") is True, f"{row['child_id']}: affinity")
        terminal_child = read(child / "terminal.json")
        require(terminal_child.get("schema") == RUNNER_SCHEMA + ".child-terminal" and terminal_child.get("status") == "completed" and terminal_child.get("return_code") == 0 and terminal_child.get("timed_out") is False and terminal_child.get("process_group_clear") is True and terminal_child.get("terminal_before_report_parse") is True, f"{row['child_id']}: terminal custody")
        run = terminal_child.get("run", {})
        require(run.get("reaped") is True and run.get("process_group_clear") is True and run.get("affinity_verified") is True, f"{row['child_id']}: process cleanup/affinity verification")
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
    contract = plan_contract(PLAN_PATH, "native")
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
    contract = plan_contract(PLAN_PATH, "allocator")
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

def independent_selection(own: dict[str, Any]) -> dict[str, Any]:
    """Recompute the bound selector without importing the reducer or selector."""
    stops: list[str] = []
    signals: list[dict[str, Any]] = []
    observed: dict[str, dict[str, dict[str, Any]]] = {}
    for lane, expected_pairs, required in (
        ("native", 6, NATIVE_REQUIRED),
        ("allocator", 3, ALLOC_METRICS),
    ):
        cases = own[lane]["cases"]
        require(set(cases) == set(c["case_id"] for c in read(PLAN_PATH)["cases"]), f"{lane}: case set changed")
        observed[lane] = {}
        for case, value in cases.items():
            metrics = value.get("metrics", {})
            observed[lane][case] = {}
            for key in required:
                metric = metrics.get(key)
                observed[lane][case][key] = metric
                label = f"{lane}/{case}/{key}"
                valid = (
                    isinstance(metric, dict)
                    and metric.get("pair_count") == expected_pairs
                    and metric.get("undefined_count") == 0
                    and type(metric.get("median")) in (int, float)
                    and math.isfinite(float(metric["median"]))
                    and float(metric["median"]) > 0
                )
                if not valid:
                    stops.append(label + ": invalid or undefined")
                    continue
                median = float(metric["median"])
                hard_safety = (
                    lane == "native" and (key in PRIMARY or key in TAIL or key == "max_rss_kib")
                ) or (
                    lane == "allocator" and key in {"region_peak_live_bytes", "max_rss_kib"}
                )
                if hard_safety and median > SAFETY_THRESHOLD:
                    stops.append(label + ": safety median exceeds 1.05")
                if median <= SIGNAL_THRESHOLD and (
                    (lane == "native" and (key in PRIMARY or key == "max_rss_kib"))
                    or (lane == "allocator" and key == "region_peak_live_bytes")
                ):
                    signals.append({"metric": label, "median": median, "threshold": SIGNAL_THRESHOLD})
    if not signals:
        stops.append("No qualifying metric among native primary/native RSS/allocator region peak")
    return {
        "decision": "reject_screen" if stops else "require_full_confirmation_and_regression_review",
        "adoption_allowed": False,
        "statistical_effect_claim": False,
        "signal_threshold": SIGNAL_THRESHOLD,
        "safety_review_threshold": SAFETY_THRESHOLD,
        "signal_comparator": "candidate_over_baseline <= threshold",
        "safety_comparator": "candidate_over_baseline > threshold",
        "stops": stops,
        "signals": signals,
        "required_metrics": observed,
        "requested_bytes_or_allocation_calls_signal": False,
    }

def compare_selection(own: dict[str, Any], existing: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    independent = independent_selection(own)
    mismatches: list[dict[str, Any]] = []
    for key in ("decision", "adoption_allowed", "statistical_effect_claim", "signal_threshold", "safety_review_threshold", "signal_comparator", "safety_comparator", "stops", "signals", "requested_bytes_or_allocation_calls_signal"):
        if existing.get(key) != independent[key]:
            mismatches.append({"field": key, "independent": independent[key], "selector": existing.get(key)})
    return independent, mismatches

def main() -> int:
    plan=read(PLAN_PATH)
    require(plan.get("schema") == "litchi.performance.0964.inline-number-screen-plan.v1", "final plan schema changed")
    require(plan.get("status") == "authorized" and plan.get("measurements_started") is False, "final plan not authorized or measurements marker changed")
    require(plan.get("cpu_mask") == CPU_MASK, "plan CPU mask changed")
    require(plan.get("root_go", {}).get("status") == "GO", "root GO is not bound")
    root_go_binding = check_root_go_binding(plan)
    affinity_binding = check_epoch2_affinity_binding(plan)
    source_binding = check_source_binding(plan)
    for lane in ("native", "allocator"):
        rows = expected_rows(plan, lane)
        require(plan["lanes"][lane].get("children") == rows, f"{lane}: plan schedule differs from independent schedule")
        for role in ("baseline_binary", "candidate_binary"):
            check_plan_descriptor(plan["lanes"][lane].get(role), f"{lane}.{role}", executable=True)
    capture_helper = check_plan_descriptor(plan.get("helpers", {}).get("capture"), "helpers.capture", executable=True)
    require(capture_helper == Path(plan["helpers"]["capture"]["path"]).resolve(), "capture helper path binding changed")
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
    existing_selection=read(SELECTION_PATH)
    require(existing_selection.get("schema") == "litchi.performance.0964.root-screen-selection.v1", "selector schema changed")
    selection, selection_mismatches = compare_selection(own, existing_selection)
    result={
      "schema":"litchi.performance.0964.independent-epoch2-raw-audit.v1",
      "status":"passed" if not mismatches and not selection_mismatches else "failed",
      "scope":"independent raw packet custody, oracle, allocator conservation, and descriptive 0964 paired median audit; no adoption claim",
      "plan":descriptor(PLAN_PATH),
      "root_go_binding":root_go_binding,
      "affinity_binding":affinity_binding,
      "source_binding":source_binding,
      "packets":custody,
      "custody_summary":{"native_children":len(all_records["native"]),"allocator_children":len(all_records["allocator"]),"native_oracles_passed":native_oracle_children,"allocator_oracles_passed":allocator_oracle_children,"no_retries_or_replacements":True,"pooled_samples":False,"source_input_oracles_passed":True,"process_group_cleanup_passed":True},
      "native":{"cases":own["native"]["cases"],"metrics_included":"all five phases at p50/p95/p99 plus whole-process max_rss_kib; review flags retained"},
      "allocator":{"cases":own["allocator"]["cases"],"metric_semantics":{"allocated_bytes":"mean core_combined_with_drain traffic bytes over ten rows","allocation_calls":"mean core_combined_with_drain allocation-call counter over ten rows","core_combined_with_drain.allocated_bytes":"same traffic-byte aggregation as allocated_bytes","region_peak_live_bytes":"maximum observer-ordered live high-water over ten core_combined_with_drain rows; not traffic","max_rss_kib":"whole-process /usr/bin/time -v maximum RSS; not allocator phase peak"}},
      "reduction_comparison":{"reducer_output":descriptor(REDUCED_PATH),"mismatch_count":len(mismatches),"mismatches":mismatches},
      "selection_comparison":{"selection_output":descriptor(SELECTION_PATH),"independent":selection,"mismatch_count":len(selection_mismatches),"mismatches":selection_mismatches,"matched":not selection_mismatches},
      "verdict":(
          "screen rejected by the bound rule; no full confirmation is indicated; no adoption or statistical effect claim"
          if selection["decision"] == "reject_screen" else
          "screen passed selection gate; independent full confirmation remains required; no adoption or statistical effect claim"
      ),
    }
    out=OUTPUT_PATH if OUTPUT_PATH is not None else ROOT/"review/independent-epoch2-raw-audit-v1.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    require(not out.exists(), f"output already exists: {out}")
    out.write_text(json.dumps(result,indent=2,sort_keys=True)+"\n",encoding="utf-8")
    print(json.dumps({"status":result["status"],"output":str(out),"output_sha256":digest(out),"mismatch_count":len(mismatches),"decision":selection["decision"]},sort_keys=True))
    return 0 if not mismatches and not selection_mismatches else 2

if __name__ == "__main__":
    try: raise SystemExit(main())
    except AuditError as exc:
        print(f"audit_screen_0964_v2: {exc}", file=sys.stderr)
        raise SystemExit(2)
