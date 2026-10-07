#!/usr/bin/env python3
"""Epoch-2 0964 screening runner with launch-time affinity hardening.

Check-plan imports the byte-bound 0959 reduction primitives and validates the
fixed schedule. Capture requires an explicit GO and approval variable. It is
serial, fresh-process, fail-stop, terminal-first, and never pre-creates the
benchmark output directory.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import signal
import subprocess
import sys
import time
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0964").resolve()
PLAN_PATH = ROOT / "protocol/plan-epoch2-0964-v1.json"
PLAN_SCHEMA = "litchi.performance.0964.inline-number-screen-plan.v1"
RUNNER_SCHEMA = "litchi.performance.0964.capture.v1"
APPROVAL_ENV = "LITCHI_PERF_0964_CAPTURE_APPROVED"
CPU_MASK = [12, 13, 14, 15]
DEADLINE = 900
FIXED_ENV = {"PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C", "RUST_BACKTRACE": "0", "TZ": "UTC"}
BASE_REDUCER = Path("/home/zhuhe/litchi-performance-0964/tools/reduce_0964_v1.py").resolve()
AFFINITY_CONTROL_COMMAND = ["/usr/bin/taskset", "--cpu-list", "12-15", "/usr/bin/true"]
CASES = ("medium-one-percent", "dense-sparse-one-percent", "medium-one-edit")
NATIVE_BLOCKS = (("A", "B", "B", "A"), ("B", "A", "A", "B"), ("A", "B", "B", "A"))
ALLOCATOR_BLOCKS = (("A", "B"), ("B", "A"), ("A", "B"))
ALLOC_PHASES = ("setup", "ingress", "edit_commit", "publication", "commit_drop", "worker_drain", "core_combined", "core_combined_with_drain", "postdrop")
ALLOC_FIELDS = ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes")
TIME_KEYS = {
    "Maximum resident set size (kbytes)": "max_rss_kib",
    "User time (seconds)": "user_seconds",
    "System time (seconds)": "system_seconds",
    "Major (requiring I/O) page faults": "major_page_faults",
    "Minor (reclaiming a frame) page faults": "minor_page_faults",
    "Voluntary context switches": "voluntary_context_switches",
    "Involuntary context switches": "involuntary_context_switches",
}


class ScreenError(RuntimeError):
    pass


class ScreenInterrupted(ScreenError):
    """Raised by the parent signal handlers so capture can seal its packet."""


def require(ok: bool, message: str) -> None:
    if not ok:
        raise ScreenError(message)


def affinity_setter(mask: list[int]):
    """Return a single-threaded child preexec hook for the frozen CPU mask."""
    frozen = set(mask)

    def set_affinity() -> None:
        os.sched_setaffinity(0, frozen)

    return set_affinity


def exact_process_affinity(pid: int) -> list[int]:
    """Read the launched process affinity from the kernel, fail-closed."""
    try:
        return sorted(os.sched_getaffinity(pid))
    except (OSError, ProcessLookupError) as exc:
        raise ScreenError(f"cannot read launched process affinity for pid {pid}: {exc}") from exc


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="microseconds")


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def desc(path: Path, executable: bool | None = None) -> dict[str, Any]:
    path = path.expanduser().resolve()
    require(path.is_file() and not path.is_symlink(), f"regular file required: {path}")
    is_exec = bool(path.stat().st_mode & 0o111)
    if executable is not None:
        require(is_exec is executable, f"executable mismatch: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "mode": path.stat().st_mode & 0o777, "executable": is_exec, "sha256": digest(path)}


def bound(value: Any, label: str, executable: bool | None = None) -> dict[str, Any]:
    require(isinstance(value, dict) and isinstance(value.get("path"), str), f"{label} binding missing")
    actual = desc(Path(value["path"]), executable)
    for key in ("path", "bytes", "mode", "executable", "sha256"):
        if key in value:
            require(value[key] == actual[key], f"{label}.{key} changed")
    return actual


def read_json(path: Path) -> Any:
    path = path.expanduser().resolve()
    require(path.is_file() and not path.is_symlink(), f"JSON missing: {path}")
    require(path.stat().st_size <= 64 * 1024 * 1024, f"JSON too large: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ScreenError(f"invalid JSON {path}: {exc}") from exc


def write_once(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(), f"artifact already exists: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write((json.dumps(value, indent=2, sort_keys=True) + "\n").encode())
        stream.flush()
        os.fsync(stream.fileno())


def rows_for(lane: str, plan: dict[str, Any]) -> list[dict[str, Any]]:
    blocks = NATIVE_BLOCKS if lane == "native" else ALLOCATOR_BLOCKS
    spec = plan["lanes"][lane]
    rows: list[dict[str, Any]] = []
    for case in plan["cases"]:
        for block, order in enumerate(blocks):
            for position, role in enumerate(order):
                rows.append({
                    "child_id": f"{lane}__{case['case_id']}__b{block:02d}__p{position}",
                    "lane": lane, "case_id": case["case_id"], "comparison_id": "baseline-vs-candidate", "role": role,
                    "binary_role": "baseline" if role == "A" else "candidate",
                    "block": block, "position": position, "shape": case["shape"],
                    "workload": case["workload"], "workers": 1,
                    "warmup_iterations": spec["warmup_iterations"],
                    "measured_samples": spec["measured_samples"], "cpu_mask": CPU_MASK,
                })
    return rows


def reduction_self_test(value: Any) -> dict[str, Any]:
    binding = bound(value, "0959 reducer")
    sys.dont_write_bytecode = True
    spec = importlib.util.spec_from_file_location("reducer_0959_math", binding["path"])
    require(spec is not None and spec.loader is not None, "cannot import 0959 reducer")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    for name in ("nearest_rank", "geometric_mean", "ratio", "bootstrap"):
        require(callable(getattr(module, name, None)), f"0959 reducer primitive missing: {name}")
    result = module.bootstrap([0.9, 1.0, 1.1])
    require(result.get("status") == "defined" and result.get("unit_count") == 3, "0959 bootstrap self-test failed")
    require(module.ratio(0.9, 1.0, "screen")["status"] == "defined", "0959 ratio self-test failed")
    return {"binding": binding, "primitives": ["nearest_rank", "geometric_mean", "ratio", "bootstrap"], "self_test": "passed"}


def load_screen_reducer() -> Any:
    sys.dont_write_bytecode = True
    spec = importlib.util.spec_from_file_location("reducer_0964_screen", BASE_REDUCER)
    require(spec is not None and spec.loader is not None, f"cannot import 0964 reducer: {BASE_REDUCER}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def manifest_entries(path: Path, label: str) -> dict[str, tuple[Any, Any, Any]]:
    value = read_json(path)
    require(isinstance(value, dict) and isinstance(value.get("files"), list), f"{label} source census missing")
    result: dict[str, tuple[Any, Any, Any]] = {}
    for item in value["files"]:
        require(isinstance(item, dict) and isinstance(item.get("path"), str), f"{label} census entry malformed")
        name = item["path"]
        require(name not in result, f"{label} census duplicate: {name}")
        result[name] = (item.get("bytes"), item.get("mode"), item.get("sha256"))
    require(value.get("file_count") == len(result), f"{label} census count mismatch")
    return result


def validate_candidate_manifest_diff(plan: dict[str, Any]) -> list[str]:
    source = plan["source"]
    before = manifest_entries(Path(source["baseline_manifest"]["path"]), "baseline")
    after = manifest_entries(Path(source["candidate_manifest"]["path"]), "candidate")
    changed = sorted(name for name in set(before) | set(after) if before.get(name) != after.get(name))
    allowlist = plan.get("candidate_change", {}).get("expected_allowlist")
    require(isinstance(allowlist, list) and all(isinstance(name, str) for name in allowlist), "candidate allowlist missing")
    require(changed == sorted(allowlist), f"candidate source diff changed paths: {changed}")
    return changed


def validate_plan(plan: dict[str, Any], capture: bool = False) -> tuple[dict[str, Any], dict[str, Any]]:
    require(plan.get("schema") == PLAN_SCHEMA, "plan schema mismatch")
    require(plan.get("status") in {"prospective", "authorized"}, "plan status invalid")
    reducer = load_screen_reducer()
    # The frozen reducer owns the schedule contract.  It requires an
    # authorized status, so prospective plans are checked through a temporary
    # in-memory status substitution; no plan file is modified.
    schedule_plan = dict(plan)
    schedule_plan["status"] = "authorized"
    try:
        # reduce_0964_v1 owns the adapter; its v1 member owns the frozen 0964
        # schedule checker while base remains the imported 0959 report core.
        reducer.v1.check_screen_plan(schedule_plan)
    except Exception as exc:
        raise ScreenError(f"frozen 0964 schedule contract rejected plan: {exc}") from exc
    require(plan.get("runner", {}).get("schema") == RUNNER_SCHEMA, "runner schema changed")
    require(plan.get("runner", {}).get("deadline_seconds") == DEADLINE, "runner deadline changed")
    require(plan.get("runner", {}).get("cpu_mask") == CPU_MASK, "runner CPU mask changed")
    require(plan.get("capture_authorization") == "root_GO_required", "root GO policy missing")
    require(plan.get("measurements_started") is False, "measurements already started")
    require([case.get("case_id") for case in plan.get("cases", [])] == list(CASES), "case order changed")
    for lane, count, warmup, samples in (("native", 36, 10, 60), ("allocator", 18, 2, 10)):
        spec = plan["lanes"][lane]
        rows = reducer.base.expected_rows(plan, lane)
        require(spec.get("children") == rows and spec.get("child_count") == count, f"{lane} schedule differs from frozen reducer rows")
        execution = spec.get("execution")
        require(isinstance(execution, dict), f"{lane} execution missing")
        require(execution.get("cpu_mask") == CPU_MASK and execution.get("deadline_seconds") == DEADLINE, f"{lane} execution custody changed")
        require(execution.get("max_parallel_children") == 1 and execution.get("retry_or_replacement") is False, f"{lane} serial/retry policy changed")
        require(execution.get("terminal_before_parse") is True, f"{lane} terminal-first policy changed")
        require(execution.get("environment", {}).get("approval_variable") == APPROVAL_ENV, f"{lane} approval binding changed")
        require(spec.get("warmup_iterations") == warmup and spec.get("measured_samples") == samples, f"{lane} sample contract changed")

    source = plan.get("source")
    require(isinstance(source, dict), "source binding missing")
    bound(source.get("baseline_manifest"), "source baseline manifest")
    bound(source.get("baseline_receipt"), "source baseline receipt")
    for key in ("candidate_manifest", "candidate_freeze"):
        value = source.get(key)
        if isinstance(value, dict) and value.get("status") == "pending_root_binding":
            require(not capture, f"capture requires source.{key}")
        else:
            bound(value, f"source {key}")
    for lane in ("native", "allocator"):
        for role in ("baseline_binary", "candidate_binary"):
            value = plan["lanes"][lane].get(role)
            if role == "candidate_binary" and isinstance(value, dict) and value.get("status") == "pending_root_binding":
                require(not capture, f"capture requires {lane}.{role}")
            else:
                bound(value, f"{lane}.{role}", True)

    helpers = plan.get("helpers")
    require(isinstance(helpers, dict), "helper bindings missing")
    for key in ("reducer", "capture", "controls"):
        value = helpers.get(key)
        require(isinstance(value, dict) and value.get("status") != "pending_root_binding", f"helper {key} is pending")
        bound(value, f"helper {key}", True if key == "capture" else None)
    running = desc(Path(__file__), True)
    require(helpers["capture"]["sha256"] == running["sha256"], "running capture helper differs from plan binding")
    hardening = plan.get("epoch2_affinity_hardening")
    require(isinstance(hardening, dict), "epoch2 affinity hardening binding missing")
    require(hardening.get("parent_preexec_affinity") == CPU_MASK, "epoch2 parent affinity binding changed")
    require(hardening.get("exact_process_affinity_getter") == "os.sched_getaffinity(pid)", "epoch2 affinity getter binding changed")
    require(hardening.get("explicit_taskset_argv_preserved") is True, "epoch2 taskset binding changed")
    require(hardening.get("fresh_54_child_schedule_required") is True, "epoch2 fresh schedule binding changed")
    require(hardening.get("prior_epoch_capture_status") == "invalid_due_to_launch_affinity", "epoch1 invalidation binding changed")
    affinity_controls = helpers.get("affinity_controls")
    require(isinstance(affinity_controls, dict), "epoch2 affinity controls binding missing")
    control_binding = bound(affinity_controls, "epoch2 affinity controls")
    require(hardening.get("controls") == affinity_controls, "epoch2 affinity controls descriptor changed")
    control = read_json(Path(control_binding["path"]))
    require(control.get("schema") == "litchi.performance.0964.affinity-launch-controls.v1", "epoch2 affinity controls schema changed")
    require(control.get("status") == "pass" and control.get("workload_started") is False, "epoch2 affinity controls did not pass before workload")
    require(control.get("cpu_mask") == CPU_MASK and control.get("command") == AFFINITY_CONTROL_COMMAND, "epoch2 affinity controls command changed")
    require(control.get("runner") == running, "epoch2 affinity controls runner binding differs")
    positive = control.get("positive")
    require(isinstance(positive, dict), "epoch2 positive affinity control missing")
    require(positive.get("requested_preexec_affinity") == CPU_MASK and positive.get("observed_affinity") == CPU_MASK, "epoch2 positive affinity control mask changed")
    require(positive.get("guard_passed") is True and positive.get("returncode") == 0, "epoch2 positive affinity control failed")
    require(positive.get("cleanup", {}).get("reaped") is True and positive.get("cleanup", {}).get("process_group_clear") is True, "epoch2 positive affinity cleanup failed")
    negative = control.get("negative")
    require(isinstance(negative, dict), "epoch2 negative affinity control missing")
    require(negative.get("requested_preexec_affinity") == [CPU_MASK[0]] and negative.get("observed_affinity") != CPU_MASK, "epoch2 negative affinity control did not exercise mismatch")
    require(negative.get("guard_rejected") is True and isinstance(negative.get("guard_error"), str), "epoch2 negative affinity control did not reject")
    require(negative.get("cleanup", {}).get("reaped") is True and negative.get("cleanup", {}).get("process_group_clear") is True, "epoch2 negative affinity cleanup failed")

    reused = plan.get("reused_0959", {})
    reduction = reduction_self_test(reused.get("reducer"))
    validator = bound(reused.get("capture_validator"), "0959 capture validator", True)
    if capture:
        require(plan.get("status") == "authorized", "capture requires status=authorized")
        require(plan.get("root_go", {}).get("status") == "GO", "capture requires root GO")
        require(os.environ.get(APPROVAL_ENV) == "1", f"{APPROVAL_ENV}=1 is required")
        host = plan.get("host")
        bound(host, "host receipt")
        validate_candidate_manifest_diff(plan)
        root_go = plan["root_go"]
        if "receipt" in root_go:
            go_desc = bound(root_go["receipt"], "root GO receipt")
            go_value = read_json(Path(go_desc["path"]))
            require(isinstance(go_value, dict) and go_value.get("status") in {"GO", "authorized", "passed"}, "root GO receipt is not authorized")
    return plan, {"plan": desc(Path(plan["_plan_path"])), "runner": running, "reduction": reduction, "validator": validator, "rows": {"native": rows_for("native", plan), "allocator": rows_for("allocator", plan)}}


def case_for(plan: dict[str, Any], case_id: str) -> dict[str, Any]:
    return next(case for case in plan["cases"] if case["case_id"] == case_id)


def integer(value: Any, label: str, positive: bool = False) -> int:
    require(type(value) is int and (value > 0 if positive else value >= 0), f"{label} invalid")
    return value


def parse_time(path: Path) -> dict[str, float | int]:
    require(path.is_file() and not path.is_symlink(), f"time output missing: {path}")
    values: dict[str, float | int] = {}
    for line in path.read_text(errors="replace").splitlines():
        if ":" not in line:
            continue
        key, raw = line.split(":", 1)
        name = TIME_KEYS.get(key.strip())
        if name is None:
            continue
        token = raw.strip().split()[0] if raw.strip() else ""
        try:
            value: float | int = float(token) if name in {"user_seconds", "system_seconds"} else int(token)
        except ValueError as exc:
            raise ScreenError(f"invalid time metric {name}") from exc
        require(math.isfinite(float(value)) and value >= 0, f"invalid time metric {name}")
        values[name] = value
    require(set(TIME_KEYS.values()) <= values.keys(), "complete /usr/bin/time metric set missing")
    require(type(values.get("max_rss_kib")) is int and values["max_rss_kib"] > 0, "positive max RSS metric missing")
    return values


def combine(values: list[dict[str, Any]], label: str) -> dict[str, Any]:
    require(values, f"{label} empty")
    for left, right in zip(values, values[1:]):
        require(left["live_bytes_after"] == right["live_bytes_before"], f"{label} live boundary changed")
    out = {"status": "measured", "scope": values[0]["scope"], "live_bytes_before": values[0]["live_bytes_before"], "live_bytes_after": values[-1]["live_bytes_after"], "peak_live_bytes_before": values[0]["peak_live_bytes_before"], "peak_live_bytes_after": values[-1]["peak_live_bytes_after"], "region_peak_live_bytes": max(x["region_peak_live_bytes"] for x in values)}
    for key in ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes"):
        out[key] = sum(x[key] for x in values)
    return out


def validate_report(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any]) -> None:
    contract = plan["report_contract"][row["lane"]]
    require(report.get("schema") == contract["schema"] and report.get("status") == "completed", f"{row['child_id']} report status/schema changed")
    case_cfg = case_for(plan, row["case_id"])
    case = report.get("case")
    require(isinstance(case, dict), f"{row['child_id']} case missing")
    if row["lane"] == "native":
        for key, expected in (("shape", row["shape"]), ("workload", row["workload"]), ("workers", 1), ("warmup_iterations", 10), ("measured_samples", 60)):
            require(report.get(key) == expected, f"{row['child_id']} native root {key} changed")
        samples = case.get("samples")
        require(isinstance(samples, list) and len(samples) == 60, f"{row['child_id']} native sample count changed")
    else:
        require(all(key not in report for key in ("shape", "workload", "workers", "warmup_iterations", "measured_samples")), f"{row['child_id']} allocator root schema changed")
        samples = case.get("rows")
        require(isinstance(samples, list) and len(samples) == 10, f"{row['child_id']} allocator sample count changed")
        for key in ("allocator_identity", "instrumentation_identity", "counter_revision", "allocation_scope"):
            if key in contract:
                require(report.get(key) == contract[key], f"{row['child_id']} allocator scope {key} changed")
    expected = {"id": f"{row['shape']}-{row['workload']}-workers1", "shape": row["shape"], "workload": row["workload"], "workers": 1, "update_count": case_cfg["update_count"], "source_bytes": case_cfg["source_bytes"], "source_sha256": case_cfg["source_sha256"], "output_bytes": case_cfg["output_bytes"], "output_sha256": case_cfg["output_sha256"], "semantic_sha256": case_cfg["semantic_sha256"], "expected_cpu_tasks": case_cfg["expected_cpu_tasks"], "admission_expected_worker_permits": 0, "admission_observed_worker_permits": 0, "admission_cpu_tasks_used": case_cfg["expected_cpu_tasks"], "preflight_actual_commit_inverse_verified": True, "context_limits": case_cfg["context_limits"], "sink_capacity_bytes": case_cfg["context_limits"]["output_bytes"], "sink_max_write_bytes": 65536, "warmup_iterations": row["warmup_iterations"], "measured_samples": row["measured_samples"]}
    for key, expected_value in expected.items():
        require(case.get(key) == expected_value, f"{row['child_id']} case oracle {key} changed")
    require(report.get("performance_claim") == contract["performance_claim"], f"{row['child_id']} performance claim changed")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("sample") == index, f"{row['child_id']} sample index changed")
        if row["lane"] == "native":
            values = [integer(sample.get(name), f"{row['child_id']} {name}", True) for name in ("ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns")]
            require(sum(values[:4]) <= values[4], f"{row['child_id']} phase sum exceeds total")
            for key, wanted in (("output_bytes", case_cfg["output_bytes"]), ("output_bytes_used", case_cfg["output_bytes"]), ("cpu_tasks_used", case_cfg["expected_cpu_tasks"]), ("output_sha256", case_cfg["output_sha256"]), ("semantic_sha256", case_cfg["semantic_sha256"])):
                require(sample.get(key) == wanted, f"{row['child_id']} native oracle {key} changed")
            require(sample.get("fixture_inverse_restored_exact_source") is True, f"{row['child_id']} inverse oracle failed")
            require(type(sample.get("sink_write_calls")) is int and sample["sink_write_calls"] > 0, f"{row['child_id']} native sink calls missing")
            require(type(sample.get("sink_largest_write")) is int and 0 < sample["sink_largest_write"] <= 65536, f"{row['child_id']} native sink bound changed")
            resources = sample.get("resources_after_drop")
        else:
            for key, wanted in (("context_limits", case_cfg["context_limits"]), ("workers_requested", 1), ("expected_worker_permits", 0), ("admitted_worker_permits", 0), ("expected_cpu_tasks", case_cfg["expected_cpu_tasks"]), ("cpu_tasks_used", case_cfg["expected_cpu_tasks"]), ("output_bytes", case_cfg["output_bytes"]), ("output_bytes_used", case_cfg["output_bytes"]), ("source_sha256", case_cfg["source_sha256"]), ("output_sha256", case_cfg["output_sha256"]), ("semantic_sha256", case_cfg["semantic_sha256"])):
                require(sample.get(key) == wanted, f"{row['child_id']} allocator oracle {key} changed")
            require(sample.get("fixture_inverse_restored_exact_source") is True, f"{row['child_id']} inverse oracle failed")
            require(type(sample.get("sink_write_calls")) is int and sample["sink_write_calls"] > 0, f"{row['child_id']} allocator sink calls missing")
            require(type(sample.get("sink_largest_write")) is int and 0 < sample["sink_largest_write"] <= 65536, f"{row['child_id']} allocator sink bound changed")
            resources = sample.get("resources_after_core")
            phases = []
            for phase in ALLOC_PHASES:
                value = sample.get(phase)
                require(isinstance(value, dict) and value.get("status") == "measured", f"{row['child_id']} {phase} missing")
                for field in ALLOC_FIELDS:
                    integer(value.get(field), f"{row['child_id']} {phase}.{field}")
                require(value["failed_allocation_calls"] == 0 and value["region_peak_live_bytes"] >= max(value["live_bytes_before"], value["live_bytes_after"]), f"{row['child_id']} {phase} counter safety failed")
                phases.append(value)
            require(phases[6] == combine(phases[1:5], "core") and phases[7] == combine(phases[1:6], "core+drain"), f"{row['child_id']} allocation conservation failed")
            integer(sample.get("max_setup_core_peak_above_setup_live_before"), f"{row['child_id']} core peak diagnostic")
            integer(sample.get("max_setup_drained_core_peak_above_setup_live_before"), f"{row['child_id']} drained peak diagnostic")
            for thread_label in ("pre_sample_threads", "post_core_threads"):
                threads = sample.get(thread_label)
                require(isinstance(threads, dict) and type(threads.get("first_thread_count")) is int and threads["first_thread_count"] >= 1 and threads.get("final_thread_count") == 1 and threads.get("drained") is True, f"{row['child_id']} {thread_label} drain contract changed")
        require(isinstance(resources, dict) and all(resources.get(name) == 0 for name in ("memory", "workers", "io_concurrency", "objects")), f"{row['child_id']} resource cleanup failed")


def group_clear(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return True
    except PermissionError:
        return False
    return False


def cleanup(process: subprocess.Popen[bytes] | None, pgid: int | None) -> dict[str, Any]:
    actions: list[str] = []
    if process is None:
        return {"actions": actions, "reaped": True, "process_group_clear": True}
    if pgid is None:
        # start_new_session makes the child PID its process-group ID.  Keep
        # that identity even after the parent has reaped and getpgid(pid)
        # consequently raises while a descendant remains in the group.
        pgid = process.pid
        try:
            os.getpgid(process.pid)
        except ProcessLookupError:
            pass
    if pgid is not None and not group_clear(pgid):
        try:
            os.killpg(pgid, signal.SIGTERM)
            actions.append("SIGTERM")
        except ProcessLookupError:
            pass
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(pgid, signal.SIGKILL)
                actions.append("SIGKILL")
            except ProcessLookupError:
                pass
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        return {"actions": actions, "reaped": False, "process_group_clear": False}
    end = time.monotonic() + 5
    while pgid is not None and time.monotonic() < end and not group_clear(pgid):
        time.sleep(0.05)
    return {"actions": actions, "reaped": True, "process_group_clear": pgid is None or group_clear(pgid)}


def run_process(argv: list[str], cwd: Path, child: Path, environment: dict[str, str], launched_ids: list[str] | None = None, child_id: str | None = None) -> dict[str, Any]:
    process: subprocess.Popen[bytes] | None = None
    pgid: int | None = None
    error: str | None = None
    timed_out = False
    returncode: int | None = None
    observed_affinity: list[int] | None = None
    started_utc = now()
    try:
        with (child / "stdout.log").open("xb") as stdout, (child / "stderr.log").open("xb") as stderr:
            # Capture this before Popen: the receipt describes launch time, not
            # the time at which a long-running child finally reaped.
            started_utc = now()
            # Popen waits for the child exec handshake, so this preexec hook
            # completes before the parent can inspect the launched process.
            # Keep the explicit taskset in argv as an independent workload
            # binding; the parent-side check catches launch drift before wait.
            process = subprocess.Popen(
                argv,
                cwd=str(cwd),
                env=environment,
                stdin=subprocess.DEVNULL,
                stdout=stdout,
                stderr=stderr,
                start_new_session=True,
                preexec_fn=affinity_setter(CPU_MASK),
            )
            if launched_ids is not None and child_id is not None and child_id not in launched_ids:
                launched_ids.append(child_id)
            # Set the cleanup identity before any post-Popen observation.  A
            # fast parent exit must not erase the group ID needed to reap
            # descendants.
            pgid = process.pid
            try:
                os.getpgid(process.pid)
            except ProcessLookupError:
                pass
            observed_affinity = exact_process_affinity(process.pid)
            require(observed_affinity == CPU_MASK, f"child affinity differs from frozen mask: {observed_affinity}")
            affinity_verified = True
            try:
                returncode = process.wait(timeout=DEADLINE)
            except subprocess.TimeoutExpired:
                timed_out = True
                clean = cleanup(process, pgid)
                returncode = process.returncode
            else:
                clean = cleanup(process, pgid)
    except BaseException as exc:
        error = f"{type(exc).__name__}: {exc}"
        if process is not None and launched_ids is not None and child_id is not None and child_id not in launched_ids:
            launched_ids.append(child_id)
        clean = cleanup(process, pgid)
        if process is not None:
            returncode = process.returncode
    return {"started_utc": started_utc, "terminal_utc": now(), "pid": process.pid if process else None, "process_group_id": pgid, "returncode": returncode, "timed_out": timed_out, "launch_error": error, "cleanup": clean, "reaped": clean["reaped"], "process_group_clear": clean["process_group_clear"], "observed_affinity": observed_affinity, "affinity_verified": error is None and observed_affinity == CPU_MASK}


def launch_affinity_control(mask: list[int]) -> tuple[subprocess.Popen[bytes], list[int], dict[str, Any]]:
    """Launch only taskset->/usr/bin/true and observe the parent affinity."""
    process = subprocess.Popen(
        AFFINITY_CONTROL_COMMAND,
        cwd="/tmp",
        env=FIXED_ENV,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
        preexec_fn=affinity_setter(mask),
    )
    observed = exact_process_affinity(process.pid)
    return process, observed, {"pid": process.pid, "command": list(AFFINITY_CONTROL_COMMAND), "requested_preexec_affinity": list(mask)}


def run_affinity_controls() -> dict[str, Any]:
    """Run positive and deliberate-mismatch launch guards before any workload."""
    wrong_mask = [CPU_MASK[0]]
    positive_process: subprocess.Popen[bytes] | None = None
    negative_process: subprocess.Popen[bytes] | None = None
    positive: dict[str, Any] = {}
    negative: dict[str, Any] = {}
    try:
        positive_process, observed, positive = launch_affinity_control(CPU_MASK)
        require(observed == CPU_MASK, f"positive control observed affinity {observed}")
        positive["observed_affinity"] = observed
        positive["guard_passed"] = True
        positive["returncode"] = positive_process.wait(timeout=5)
        require(positive["returncode"] == 0, f"positive control return code {positive['returncode']}")
        positive["cleanup"] = cleanup(positive_process, positive_process.pid)
        positive_process = None

        negative_process, observed, negative = launch_affinity_control(wrong_mask)
        negative["observed_affinity"] = observed
        try:
            require(observed == CPU_MASK, f"negative control observed affinity {observed}")
        except ScreenError as exc:
            negative["guard_rejected"] = True
            negative["guard_error"] = str(exc)
        require(negative.get("guard_rejected") is True, "negative control was not rejected")
        negative["cleanup"] = cleanup(negative_process, negative_process.pid)
        negative_process = None
    finally:
        if positive_process is not None:
            cleanup(positive_process, positive_process.pid)
        if negative_process is not None:
            cleanup(negative_process, negative_process.pid)
    return {
        "schema": "litchi.performance.0964.affinity-launch-controls.v1",
        "status": "pass",
        "runner": desc(Path(__file__), True),
        "cpu_mask": CPU_MASK,
        "command": list(AFFINITY_CONTROL_COMMAND),
        "positive": positive,
        "negative": negative,
        "workload_started": False,
    }


def child_environment(plan: dict[str, Any], lane: str) -> dict[str, str]:
    configured = plan["lanes"][lane]["execution"]["environment"]
    environment = {key: str(value) for key, value in configured.items() if key != "approval_variable"}
    environment[configured["approval_variable"]] = "1"
    return environment


def binary_for(plan: dict[str, Any], lane: str, row: dict[str, Any]) -> Path:
    key = "baseline_binary" if row["binary_role"] == "baseline" else "candidate_binary"
    return Path(plan["lanes"][lane][key]["path"]).resolve()


def command_for(plan: dict[str, Any], row: dict[str, Any], child: Path) -> tuple[list[str], list[str], Path, Path, dict[str, str]]:
    binary = binary_for(plan, row["lane"], row)
    cwd = Path(plan["lanes"][row["lane"]]["execution"]["workload_cwd"])
    output = child / "output"
    time_file = child / "time-v.txt"
    require(not output.exists() and not output.is_symlink(), "benchmark output must be absent before launch")
    direct = [
        "/usr/bin/taskset", "--cpu-list", "12-15", str(binary),
        "--shape", row["shape"], "--workload", row["workload"],
        "--workers", "1", "--warmup", str(row["warmup_iterations"]),
        "--samples", str(row["measured_samples"]), "--output", str(output),
    ]
    wrapped = [
        "/usr/bin/taskset", "--cpu-list", "12-15", "/usr/bin/time", "-v",
        "-o", str(time_file), "--", *direct[3:],
    ]
    return wrapped, direct, cwd, output, child_environment(plan, row["lane"])


def run_child(plan: dict[str, Any], row: dict[str, Any], packet: Path, launched_ids: list[str]) -> None:
    child = packet / "children" / row["child_id"]
    child.mkdir(parents=True, exist_ok=False)
    argv, workload_argv, cwd, output, environment = command_for(plan, row, child)
    binary = binary_for(plan, row["lane"], row)
    binary_key = "baseline_binary" if row["binary_role"] == "baseline" else "candidate_binary"
    # Recheck the authorized role binding for every child immediately before
    # the launch receipt is written.  A plan-wide preflight cannot detect a
    # binary replacement between children.
    binary_binding = bound(plan["lanes"][row["lane"]][binary_key], f"{row['child_id']} {binary_key}", True)
    require(binary_binding["path"] == str(binary), f"{row['child_id']} binary path changed")
    start = {
        "schema": RUNNER_SCHEMA + ".child-start",
        "index": row["block"] * 4 + row["position"],
        "child": row,
        "binary": binary_binding,
        "argv": argv,
        "workload_argv": workload_argv,
        "cwd": str(cwd),
        "environment": environment,
        "deadline_seconds": DEADLINE,
        "output_absent_before_launch": True,
        "started_utc": now(),
    }
    write_once(child / "start.json", start)
    run = run_process(argv, cwd, child, environment, launched_ids, row["child_id"])
    ok = run["launch_error"] is None and run["returncode"] == 0 and not run["timed_out"] and run["reaped"] and run["process_group_clear"]
    write_once(child / "launch.json", {"schema": RUNNER_SCHEMA + ".child-launch", "pid": run["pid"], "observed_affinity": run["observed_affinity"], "expected_affinity": CPU_MASK, "affinity_verified_before_wait": run["affinity_verified"], "launched_utc": run["started_utc"]})
    terminal = {
        "schema": RUNNER_SCHEMA + ".child-terminal",
        "index": start["index"],
        "child_id": row["child_id"],
        "status": "completed" if ok else "failed",
        "return_code": run["returncode"],
        "timed_out": run["timed_out"],
        "process_group_clear": run["process_group_clear"],
        "pid": run["pid"],
        "run": run,
        "output": {"path": str(output), "exists_after_terminal": output.is_dir() and not output.is_symlink()},
        "terminal_before_report_parse": True,
        "ended_utc": now(),
    }
    write_once(child / "terminal.json", terminal)
    require(ok, f"{row['child_id']} process failed")
    try:
        report = output / "report.json"
        require(report.is_file() and not report.is_symlink(), "report missing")
        validate_report(read_json(report), row, plan)
        require((child / "time-v.txt").is_file(), "time output missing")
        validation = {"schema": RUNNER_SCHEMA + ".child-validation", "status": "passed", "index": start["index"], "child": row, "terminal": desc(child / "terminal.json"), "report": desc(report), "time": desc(child / "time-v.txt"), "time_metrics": parse_time(child / "time-v.txt")}
    except BaseException as exc:
        validation = {"schema": RUNNER_SCHEMA + ".child-validation", "status": "failed", "index": start["index"], "child": row, "terminal": desc(child / "terminal.json"), "error": f"{type(exc).__name__}: {exc}"}
        write_once(child / "validation.json", validation)
        raise
    write_once(child / "validation.json", validation)
    write_once(
        child / "post-terminal-artifacts.json",
        {
            "schema": RUNNER_SCHEMA + ".post-terminal-artifacts",
            "terminal": desc(child / "terminal.json"),
            "stdout": desc(child / "stdout.log"),
            "stderr": desc(child / "stderr.log"),
            "time": desc(child / "time-v.txt"),
            "report": desc(report),
            "validation": desc(child / "validation.json"),
        },
    )


def _raise_for_signal(signum: int, _frame: Any) -> None:
    raise ScreenInterrupted(f"received signal {signum}")


def _install_signal_handlers() -> dict[int, Any]:
    previous: dict[int, Any] = {}
    for signum in (signal.SIGINT, signal.SIGTERM):
        previous[signum] = signal.getsignal(signum)
        signal.signal(signum, _raise_for_signal)
    return previous


def _restore_signal_handlers(previous: dict[int, Any]) -> None:
    for signum, handler in previous.items():
        signal.signal(signum, handler)


def capture(plan: dict[str, Any], base: dict[str, Any], lane: str, packet: Path) -> int:
    root = packet.resolve()
    require(not root.exists() and not root.is_symlink(), "capture root already exists")
    root.mkdir(parents=True, exist_ok=False)
    (root / "children").mkdir()
    rows = base["rows"][lane]
    completed: list[str] = []
    launched: list[str] = []
    error: str | None = None
    previous_handlers = _install_signal_handlers()
    try:
        write_once(root / "capture-start.json", {"schema": RUNNER_SCHEMA + ".capture-start", "lane": lane, "children_expected": len(rows), "serial": True, "retry_or_replacement": False, "plan": base["plan"], "runner": desc(Path(__file__), True), "started_utc": now()})
        for row in rows:
            try:
                run_child(plan, row, root, launched)
                completed.append(row["child_id"])
            except BaseException as exc:
                error = f"{row['child_id']}: {type(exc).__name__}: {exc}"
                break
    except BaseException as exc:
        error = f"capture: {type(exc).__name__}: {exc}"
    finally:
        # The terminal receipt is written even when the parent receives SIGINT
        # or SIGTERM between children.  launched is populated at Popen time;
        # completed intentionally counts only validated children.
        status = "completed" if error is None and len(completed) == len(rows) else "failed"
        terminal = {"schema": RUNNER_SCHEMA + ".capture-terminal", "status": status, "lane": lane, "children_started": len(launched), "children_expected": len(rows), "children_completed": len(completed), "terminal_before_parse": True, "retry_or_replacement": False, "capture_start": desc(root / "capture-start.json") if (root / "capture-start.json").exists() else None, "ended_utc": now()}
        if error is not None:
            terminal["error"] = error
        try:
            write_once(root / "capture-terminal.json", terminal)
        finally:
            _restore_signal_handlers(previous_handlers)
    if status == "completed":
        write_once(root / "completed-packet.json", {"schema": RUNNER_SCHEMA + ".completed-packet", "status": "completed", "lane": lane, "children": [{"child_id": row["child_id"], "validation": desc(root / "children" / row["child_id"] / "validation.json")} for row in rows], "capture_terminal": desc(root / "capture-terminal.json")})
    print(json.dumps({"status": status, "lane": lane, "children": len(completed), "children_started": len(launched), "capture_root": str(root)}, sort_keys=True))
    return 0 if status == "completed" else 2


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--check-plan", action="store_true")
    group.add_argument("--capture", action="store_true")
    group.add_argument("--affinity-controls", action="store_true")
    parser.add_argument("--plan", type=Path, default=PLAN_PATH)
    parser.add_argument("--lane", choices=("native", "allocator"))
    parser.add_argument("--go", action="store_true")
    parser.add_argument("--packet", type=Path)
    parser.add_argument("--affinity-controls-output", type=Path)
    args = parser.parse_args(argv)
    if args.affinity_controls:
        result = run_affinity_controls()
        if args.affinity_controls_output is not None:
            write_once(args.affinity_controls_output.expanduser().resolve(), result)
        print(json.dumps(result, sort_keys=True))
        return 0
    plan_path = args.plan.expanduser().resolve()
    plan = read_json(plan_path)
    plan["_plan_path"] = str(plan_path)
    if args.capture:
        require(args.go and args.lane is not None, "capture requires --go and --lane")
        require(args.packet is not None, "capture requires --packet")
        plan, base = validate_plan(plan, capture=True)
        require(args.packet.parent.resolve().is_dir(), f"packet parent unavailable: {args.packet.parent}")
        return capture(plan, base, args.lane, args.packet)
    plan, base = validate_plan(plan)
    print(json.dumps({"status": "check-plan-pass", "native_children": 36, "allocator_children": 18, "reduction": base["reduction"]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except ScreenError as exc:
        print(f"screen_0964_v2: {exc}", file=sys.stderr)
        raise SystemExit(2)
