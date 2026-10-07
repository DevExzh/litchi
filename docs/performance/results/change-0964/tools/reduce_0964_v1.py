#!/usr/bin/env python3
"""Fail-closed reducer for the 0964 screen-v2 packet contract.

The protocol runner owns the raw plan shape and capture lifecycle.  This
adapter keeps the frozen 0964 report/oracle and schedule contract from
``reduce_0964_core_v1.py`` while translating the protocol plan's equivalent
top-level binary bindings and list-shaped block orders into the reducer's
canonical view.  It does not loosen report, source, binary, or terminal
custody checks, and it emits descriptive screening output only.
"""
from __future__ import annotations

import copy
import importlib.util
import json
import math
import os
from pathlib import Path
import sys
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0964").resolve()
V1 = ROOT / "tools" / "reduce_0964_core_v1.py"
PROTOCOL_SCHEMA = "litchi.performance.0964.raw-cell-vector-reserve-screening.v1"
CANONICAL_SCHEMA = "litchi.performance.0964.inline-number-screen-plan.v1"
CPU_MASK = [12, 13, 14, 15]
DEADLINE = 900
FIXED_ENV = {"PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C", "RUST_BACKTRACE": "0", "TZ": "UTC"}
CASES = ["medium-one-percent", "dense-sparse-one-percent", "medium-one-edit"]
NATIVE_BLOCKS = [["A", "B", "B", "A"], ["B", "A", "A", "B"], ["A", "B", "B", "A"]]
ALLOCATOR_BLOCKS = [["A", "B"], ["B", "A"], ["A", "B"]]


class ReductionError(ValueError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReductionError(message)


def load_module(path: Path):
    spec = importlib.util.spec_from_file_location("reduce_0964_core_v1_frozen", path)
    require(spec is not None and spec.loader is not None, f"cannot load reducer: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


v1 = load_module(V1)
base = v1.base


def digest(path: Path) -> str:
    import hashlib
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def descriptor(path: Path, executable: bool | None = None) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"regular file required: {path}")
    stat = path.stat()
    actual = {
        "path": str(path.resolve()),
        "bytes": stat.st_size,
        "mode": stat.st_mode & 0o777,
        "executable": bool(stat.st_mode & 0o111),
        "sha256": digest(path),
    }
    if executable is not None:
        require(actual["executable"] is executable, f"executable mismatch: {path}")
    return actual


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"JSON missing: {path}")
    require(path.stat().st_size <= 64 * 1024 * 1024, f"JSON too large: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ReductionError(f"invalid JSON {path}: {exc}") from exc


def checked_external(value: Any, label: str, executable: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict) and isinstance(value.get("path"), str), f"{label} descriptor missing")
    actual = descriptor(Path(value["path"]), executable)
    for key in ("path", "bytes", "mode", "executable", "sha256"):
        if key in value:
            require(value[key] == actual[key], f"{label}.{key} changed")
    return actual


def _alias(mapping: dict[str, Any], *keys: str) -> Any:
    for key in keys:
        if key in mapping:
            return mapping[key]
    return None


def _canonical_rows(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    return copy.deepcopy(plan["lanes"][lane]["children"])


def canonical_plan(raw: dict[str, Any]) -> tuple[dict[str, Any], str]:
    """Build the report-validator view without changing the bound plan bytes."""
    require(isinstance(raw, dict), "plan root is not an object")
    schema = raw.get("schema")
    if schema == CANONICAL_SCHEMA:
        return copy.deepcopy(raw), "canonical-v1"
    require(schema == PROTOCOL_SCHEMA, f"unsupported 0964 plan schema: {schema}")
    plan = copy.deepcopy(raw)
    comparison = plan.get("comparison")
    require(isinstance(comparison, dict), "comparison binding missing")
    native_orders = comparison.get("native_blocks")
    allocator_orders = comparison.get("allocator_blocks")
    require(native_orders == NATIVE_BLOCKS, "native block order changed")
    require(allocator_orders == ALLOCATOR_BLOCKS, "allocator block order changed")
    plan["comparison"] = dict(comparison)
    plan["comparison"]["blocks"] = [{"id": i, "order": list(order)} for i, order in enumerate(native_orders)]
    plan["comparison"]["allocator_blocks"] = [{"id": i, "order": list(order)} for i, order in enumerate(allocator_orders)]
    plan["comparison"].update({
        "block_count": 3,
        "pairs_per_case_native": 6,
        "pairs_per_case_allocator": 3,
        "screening_only": True,
        "full_followup_required": True,
    })
    contracts = plan.get("report_contract")
    require(isinstance(contracts, dict), "report contract missing")
    lanes = copy.deepcopy(plan.get("lanes"))
    require(isinstance(lanes, dict), "lanes missing")
    baseline = plan.get("baseline")
    candidate = plan.get("candidate")
    require(isinstance(baseline, dict) and isinstance(candidate, dict), "top-level binary bindings missing")
    for lane, warmups, samples in (("native", 10, 60), ("allocator", 2, 10)):
        spec = lanes.get(lane)
        require(isinstance(spec, dict), f"{lane} lane missing")
        contract = contracts.get(lane)
        require(isinstance(contract, dict), f"{lane} report contract missing")
        spec["report_schema"] = contract.get("schema")
        require(isinstance(spec["report_schema"], str), f"{lane} report schema missing")
        require(spec.get("warmup_iterations") == warmups and spec.get("measured_samples") == samples, f"{lane} sample schedule changed")
        spec["execution"] = {
            "cpu_mask": list(plan.get("cpu_mask", CPU_MASK)),
            "deadline_seconds": plan.get("deadline_seconds"),
            "max_parallel_children": 1,
            "retry_or_replacement": False,
            "terminal_before_parse": True,
        }
        # Prospective templates may carry the literal pending marker.  It is
        # deliberately copied here only to permit static schedule inspection;
        # validate_plan_bindings rejects it before any packet is opened.
        spec["baseline_binary"] = copy.deepcopy(baseline.get(f"{lane}_binary"))
        spec["candidate_binary"] = copy.deepcopy(candidate.get(f"{lane}_binary"))
        spec["children"] = _canonical_rows(plan, lane)
        lanes[lane] = spec
    plan["lanes"] = lanes
    plan["schema"] = CANONICAL_SCHEMA
    plan["gates"] = dict(plan.get("gates") or {})
    plan["gates"].update({"adoption_allowed": False, "full_followup_required": True})
    return plan, "protocol-v1"


def validate_plan_bindings(raw: dict[str, Any], canonical: dict[str, Any], plan_path: Path, helper_path: Path) -> dict[str, Any]:
    require(raw.get("status") == "authorized", "reduction requires an authorized plan")
    require(raw.get("measurements_started") is False, "plan measurements already started")
    require(raw.get("capture_authorization") == "root_GO_required", "root GO policy missing")
    require(raw.get("root_go", {}).get("status") == "GO", "root GO is not bound")
    require(raw.get("cpu_mask") == CPU_MASK and raw.get("deadline_seconds") == DEADLINE, "CPU/deadline binding changed")
    require(raw.get("environment") == FIXED_ENV, "fixed environment changed")
    execution = raw.get("execution")
    require(isinstance(execution, dict), "execution binding missing")
    for key in ("serial", "fresh_process_per_child", "no_retries", "stop_on_first_failure", "child_owns_output_directory"):
        require(execution.get(key) is True, f"execution.{key} changed")
    # The canonical runner names this policy terminal_before_parse; the old
    # compact draft called it terminal_before_report_parse.  Both refer to the
    # same terminal-first boundary, but an authorized canonical plan must bind
    # the canonical spelling.
    require(execution.get("terminal_before_parse") is True, "execution.terminal_before_parse changed")
    if raw.get("schema") != CANONICAL_SCHEMA:
        require(execution.get("terminal_before_report_parse") is True, "execution.terminal_before_report_parse changed")
    require(raw.get("lanes", {}).get("native", {}).get("children") == _canonical_rows(raw, "native"), "native schedule changed")
    require(raw.get("lanes", {}).get("allocator", {}).get("children") == _canonical_rows(raw, "allocator"), "allocator schedule changed")
    comparison = raw.get("comparison")
    require(isinstance(comparison, dict), "comparison binding missing")
    if raw.get("schema") == CANONICAL_SCHEMA:
        require([block.get("order") for block in comparison.get("blocks", [])] == NATIVE_BLOCKS, "native block order changed")
        require([block.get("order") for block in comparison.get("allocator_blocks", [])] == ALLOCATOR_BLOCKS, "allocator block order changed")
    # v5 uses the existing contract's lane-local executable bindings.  The
    # earlier compact protocol draft kept those at plan root; preserve a
    # read-only adapter for that historical shape, but require the canonical
    # v5 shape for an authorized reduction.
    if raw.get("schema") == CANONICAL_SCHEMA:
        for lane in ("native", "allocator"):
            lane_spec = raw.get("lanes", {}).get(lane)
            require(isinstance(lane_spec, dict), f"{lane} lane missing")
            for role in ("baseline_binary", "candidate_binary"):
                checked_external(lane_spec.get(role), f"{lane}.{role}", executable=True)
            require(isinstance(lane_spec.get("execution", {}).get("workload_cwd"), str), f"{lane}.workload_cwd missing")
            require(Path(lane_spec["execution"]["workload_cwd"]).is_dir(), f"{lane}.workload_cwd missing on disk")
    else:
        for role in ("baseline", "candidate"):
            binding = raw.get(role)
            require(isinstance(binding, dict), f"{role} binding missing")
            require(binding.get("status") not in (None, "pending_root_binding"), f"{role} binding is pending")
            for key in ("source_manifest", "native_binary", "allocator_binary"):
                checked_external(binding.get(key), f"{role}.{key}", key.endswith("binary"))
            require(isinstance(binding.get("source_root"), str) and Path(binding["source_root"]).is_dir(), f"{role}.source_root missing")
    source = raw.get("source")
    require(isinstance(source, dict), "source binding missing")
    # Final protocol plans may name these descriptors under source, while the
    # retained v2 draft carries the manifest under baseline/candidate.  The
    # fallback is explicit and still rehashes the actual manifest above.
    baseline_manifest_fallback = _alias(source, "baseline_manifest")
    candidate_manifest_fallback = _alias(source, "candidate_manifest")
    if raw.get("schema") == CANONICAL_SCHEMA:
        baseline_manifest_fallback = baseline_manifest_fallback or source.get("baseline_manifest")
        candidate_manifest_fallback = candidate_manifest_fallback or source.get("candidate_manifest")
    else:
        baseline_manifest_fallback = baseline_manifest_fallback or raw["baseline"].get("source_manifest")
        candidate_manifest_fallback = candidate_manifest_fallback or raw["candidate"].get("source_manifest")
    for label, value in (("baseline_manifest", baseline_manifest_fallback), ("candidate_manifest", candidate_manifest_fallback)):
        checked_external(value, f"source.{label}")
    # The runtime manifest is a separate custody input from the compact source
    # manifest.  If a final plan carries it, it is required to be a real
    # descriptor too; accepting only a status string would let the reducer
    # validate a different source closure than the capture runner used.
    for label in ("baseline_runtime_manifest", "candidate_runtime_manifest"):
        value = source.get(label)
        if value is not None:
            checked_external(value, f"source.{label}")
    for label in ("baseline_source_root", "baseline_runtime_source_root", "candidate_source_root", "candidate_runtime_source_root"):
        value = source.get(label)
        if value is not None:
            require(isinstance(value, str) and Path(value).is_dir() and not Path(value).is_symlink(), f"source.{label} missing or symlinked")
    for label, value in (("baseline_receipt", _alias(source, "baseline_receipt", "baseline_build_receipt")), ("candidate_freeze", _alias(source, "candidate_freeze", "candidate_build_receipt", "candidate_source_freeze"))):
        require(isinstance(value, dict) and value.get("status") != "pending_root_binding", f"source.{label} is pending")
        checked_external(value, f"source.{label}")
    helpers = raw.get("helpers")
    require(isinstance(helpers, dict), "helper bindings missing")
    reducer = _alias(helpers, "reducer")
    actual_helper = descriptor(helper_path)
    require(isinstance(reducer, dict) and all(reducer.get(k) == actual_helper[k] for k in ("path", "bytes", "sha256")), "reducer helper binding differs")
    for label in ("capture", "controls"):
        value = _alias(helpers, label, f"{label}_adapter")
        require(isinstance(value, dict), f"{label} helper binding missing")
        checked_external(value, f"helpers.{label}", executable=(True if label == "capture" else None))
    runner = raw.get("runner")
    require(isinstance(runner, dict), "runner binding missing")
    if raw.get("schema") == CANONICAL_SCHEMA:
        capture_helper = _alias(helpers, "capture")
        checked_runner = checked_external(capture_helper, "helpers.capture", executable=True)
        require(runner.get("schema") == "litchi.performance.0964.capture.v1", "runner schema changed")
        require(runner.get("deadline_seconds") == DEADLINE and runner.get("cpu_mask") == CPU_MASK, "runner custody values changed")
    else:
        checked_runner = checked_external(runner, "runner", executable=True)
        require(checked_runner["path"] == runner["path"], "runner path changed")
    host = raw.get("host")
    require(isinstance(host, dict) and host.get("status") not in (None, "pending_root_binding"), "host binding missing")
    host_receipt = _alias(host, "receipt", "host_receipt", "identity_receipt")
    require(isinstance(host_receipt, dict), "host receipt descriptor missing")
    checked_external(host_receipt, "host.receipt")
    root_go = raw.get("root_go")
    require(isinstance(root_go, dict) and root_go.get("status") == "GO", "root GO binding missing")
    root_receipt = _alias(root_go, "receipt", "go_receipt", "authorization_receipt")
    require(isinstance(root_receipt, dict), "root GO receipt descriptor missing")
    checked_external(root_receipt, "root_go.receipt")
    return {"plan": descriptor(plan_path), "runner": checked_runner, "reducer": actual_helper}


def _report_rows(raw_plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    rows = raw_plan["lanes"][lane].get("children")
    require(isinstance(rows, list), f"{lane} child schedule missing")
    if raw_plan.get("schema") == CANONICAL_SCHEMA:
        expected = base.expected_rows(raw_plan, lane)
        require(rows == expected, f"{lane} child schedule differs from frozen order")
        return copy.deepcopy(expected)
    expected = []
    blocks = NATIVE_BLOCKS if lane == "native" else ALLOCATOR_BLOCKS
    for case in raw_plan["cases"]:
        for block, order in enumerate(blocks):
            for position, role in enumerate(order):
                expected.append({
                    "child_id": f"{lane}__{case['case_id']}__b{block:02d}__p{position}",
                    "lane": lane, "case_id": case["case_id"], "role": role,
                    "binary_role": "baseline" if role == "A" else "candidate",
                    "block": block, "position": position, "shape": case["shape"],
                    "workload": case["workload"], "workers": 1,
                    "warmup_iterations": raw_plan["lanes"][lane]["warmup_iterations"],
                    "measured_samples": raw_plan["lanes"][lane]["measured_samples"],
                    "cpu_mask": CPU_MASK,
                })
    require(rows == expected, f"{lane} child schedule differs from frozen order")
    return expected


def _time(path: Path) -> dict[str, float | int]:
    values = {}
    for line in path.read_text(errors="replace").splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        token = value.strip().split()[0] if value.strip() else ""
        if key.strip() == "Maximum resident set size (kbytes)":
            values["max_rss_kib"] = int(token)
        elif key.strip() == "User time (seconds)":
            values["user_seconds"] = float(token)
        elif key.strip() == "System time (seconds)":
            values["system_seconds"] = float(token)
    require(type(values.get("max_rss_kib")) is int and values["max_rss_kib"] > 0, "positive max RSS missing")
    return values


def _schema(value: Any, *accepted: str) -> None:
    require(value in accepted, f"schema changed: {value}")


def validate_packet(packet: Path, plan_path: Path, raw_plan: dict[str, Any], canonical: dict[str, Any], lane: str, runner_schema: str) -> list[dict[str, Any]]:
    require(packet.is_dir() and not packet.is_symlink(), f"{lane} packet missing")
    rows = _report_rows(raw_plan, lane)
    start_path = packet / "capture-start.json"
    terminal_path = packet / "capture-terminal.json"
    start = read_json(start_path)
    runner_binding = _alias(raw_plan.get("helpers", {}), "capture")
    expected_runner = checked_external(runner_binding, f"{lane}.capture-runner", executable=True)
    _schema(start.get("schema"), runner_schema, runner_schema + ".capture-start")
    require(start.get("lane") == lane and start.get("serial") is True, f"{lane} capture start custody changed")
    require(start.get("no_retry") is True or start.get("retry_or_replacement") is False, f"{lane} capture retry policy changed")
    require(start.get("rows") == rows or start.get("children_expected") == len(rows), f"{lane} capture schedule missing")
    require(start.get("plan", {}).get("sha256") == descriptor(plan_path)["sha256"], f"{lane} capture plan binding changed")
    start_runner = checked_external(start.get("runner"), f"{lane}.capture-start.runner", executable=True)
    require(start_runner == expected_runner, f"{lane} capture runner binding changed")
    terminal = read_json(terminal_path)
    _schema(terminal.get("schema"), runner_schema, runner_schema + ".capture-terminal")
    require(terminal.get("status") in {"passed", "completed"}, f"{lane} capture terminal failed")
    expected_ids = [row["child_id"] for row in rows]
    if terminal.get("completed_children") is not None:
        require(terminal.get("completed_children") == expected_ids and terminal.get("expected_children") == expected_ids, f"{lane} children incomplete or reordered")
    else:
        require(terminal.get("children_started") == len(rows) and terminal.get("children_expected") == len(rows), f"{lane} children incomplete")
    require(terminal.get("no_retry") is True or terminal.get("retry_or_replacement") is False, f"{lane} retry policy changed")
    if terminal.get("capture_start") is not None:
        require(terminal["capture_start"].get("sha256") == descriptor(start_path)["sha256"], f"{lane} terminal start binding changed")
    completed = packet / "completed-packet.json"
    if completed.exists():
        completed_value = read_json(completed)
        require(completed_value.get("schema") == runner_schema + ".completed-packet" and completed_value.get("status") == "completed", f"{lane} completed packet changed")
        require([item.get("child_id") for item in completed_value.get("children", [])] == expected_ids, f"{lane} completed packet order changed")
        require(completed_value.get("capture_terminal", {}).get("sha256") == descriptor(terminal_path)["sha256"], f"{lane} completed terminal binding changed")
    records = []
    for row in rows:
        child = packet / "children" / row["child_id"]
        require(child.is_dir() and not child.is_symlink(), f"{row['child_id']} child missing")
        child_start = read_json(child / "start.json")
        require(child_start.get("schema") == runner_schema + ".child-start", f"{row['child_id']} start schema changed")
        v3_child = isinstance(child_start.get("child"), dict)
        require((child_start.get("child") if v3_child else child_start.get("row")) == row, f"{row['child_id']} start row changed")
        lane_spec = raw_plan["lanes"][lane]
        expected_environment = dict(lane_spec["execution"].get("environment", FIXED_ENV))
        if "approval_variable" in expected_environment:
            approval = expected_environment.pop("approval_variable")
            expected_environment[approval] = "1"
        require(child_start.get("deadline_seconds") == DEADLINE and child_start.get("environment") == expected_environment, f"{row['child_id']} start custody changed")
        expected_binary = checked_external(lane_spec["baseline_binary" if row["role"] == "A" else "candidate_binary"], f"{row['child_id']}.binary", executable=True)
        expected_cwd = lane_spec["execution"].get("workload_cwd")
        require(child_start.get("cwd") == expected_cwd, f"{row['child_id']} cwd changed")
        output = child / "output"
        time_path = child / ("time-v.txt" if v3_child else "time.txt")
        direct = ["/usr/bin/taskset", "--cpu-list", "12-15", expected_binary["path"], "--shape", row["shape"], "--workload", row["workload"], "--workers", "1", "--warmup", str(row["warmup_iterations"]), "--samples", str(row["measured_samples"]), "--output", str(output)]
        expected_argv = ["/usr/bin/taskset", "--cpu-list", "12-15", "/usr/bin/time", "-v", "-o", str(time_path), "--", *direct[3:]] if v3_child else ["/usr/bin/taskset", "--cpu-list", "12-15", "/usr/bin/time", "-v", "-o", str(time_path), expected_binary["path"], "--shape", row["shape"], "--workload", row["workload"], "--workers", "1", "--warmup", str(row["warmup_iterations"]), "--samples", str(row["measured_samples"]), "--output", str(output)]
        require(child_start.get("argv") == expected_argv, f"{row['child_id']} argv changed")
        if v3_child:
            require(child_start.get("workload_argv") == direct, f"{row['child_id']} workload argv changed")
        require(child_start.get("binary", {}).get("sha256") == expected_binary["sha256"], f"{row['child_id']} binary binding changed")
        if v3_child:
            launch = read_json(child / "launch.json")
            require(launch.get("schema") == runner_schema + ".child-launch" and launch.get("observed_affinity") == CPU_MASK, f"{row['child_id']} affinity changed")
        child_terminal = read_json(child / "terminal.json")
        require(child_terminal.get("schema") == runner_schema + ".child-terminal" and child_terminal.get("status") in {"pass", "completed"}, f"{row['child_id']} terminal status changed")
        run = child_terminal.get("run", {})
        require(run.get("returncode", child_terminal.get("return_code")) == 0 and run.get("timed_out", child_terminal.get("timed_out")) is False and run.get("reaped", True) is True and run.get("process_group_clear", child_terminal.get("process_group_clear")) is True, f"{row['child_id']} process did not terminate cleanly")
        require(child_terminal.get("terminal_before_report_parse") is True, f"{row['child_id']} terminal-first marker missing")
        report_path = output / "report.json"
        validation = read_json(child / "validation.json")
        require(validation.get("schema") == runner_schema + ".child-validation" and validation.get("status") == "passed" and (validation.get("child") if v3_child else validation.get("row")) == row, f"{row['child_id']} validation failed")
        require(validation.get("terminal", {}).get("sha256") == descriptor(child / "terminal.json")["sha256"], f"{row['child_id']} validation terminal binding changed")
        require(validation.get("report", {}).get("sha256") == descriptor(report_path)["sha256"], f"{row['child_id']} validation report binding changed")
        require(validation.get("time", {}).get("sha256") == descriptor(time_path)["sha256"], f"{row['child_id']} validation time binding changed")
        if v3_child:
            post = read_json(child / "post-terminal-artifacts.json")
            require(post.get("schema") == runner_schema + ".post-terminal-artifacts", f"{row['child_id']} post-terminal schema changed")
            for key, path in (("terminal", child / "terminal.json"), ("stdout", child / "stdout.log"), ("stderr", child / "stderr.log"), ("time", time_path), ("report", report_path), ("validation", child / "validation.json")):
                require(post.get(key, {}).get("sha256") == descriptor(path)["sha256"], f"{row['child_id']} post-terminal {key} binding changed")
        report = read_json(report_path)
        if lane == "native":
            v1.validate_native_screen_report(report, row, canonical)
            samples = report.get("case", {}).get("samples", [])
            require(all(type(sample.get("edit_commit_ns")) is int and sample["edit_commit_ns"] > 0 and type(sample.get("total_ns")) is int and sample["total_ns"] > 0 for sample in samples), f"{row['child_id']} useful primary samples are not positive")
        else:
            base.validate_allocator_report(report, row, canonical)
        records.append({"row": row, "report": report, "time": _time(time_path), "child": row["child_id"], "child_path": child})
    return records


def reduce_packets(plan_path: Path, native_packet: Path, allocator_packet: Path, output: Path, helper_path: Path) -> dict[str, Any]:
    raw = read_json(plan_path)
    canonical, mapping = canonical_plan(raw)
    require(raw.get("status") == "authorized", "authorized plan required")
    binding = validate_plan_bindings(raw, canonical, plan_path, helper_path)
    runner_schema = raw["runner"].get("schema")
    require(isinstance(runner_schema, str) and runner_schema, "runner schema missing")
    base.RUNNER_SCHEMA = runner_schema
    base.DEADLINE = DEADLINE
    base.CPU_MASK = CPU_MASK
    base.MAX_SINK_WRITE = 65_536
    base.validate_native_report = v1.validate_native_screen_report
    native_records = validate_packet(native_packet.resolve(), plan_path.resolve(), raw, canonical, "native", runner_schema)
    allocator_records = validate_packet(allocator_packet.resolve(), plan_path.resolve(), raw, canonical, "allocator", runner_schema)
    native_cases = {case["case_id"]: v1.screen_native_case(native_records, canonical, case["case_id"]) for case in raw["cases"]}
    allocator_cases = {case["case_id"]: v1.screen_allocator_case(allocator_records, canonical, case["case_id"]) for case in raw["cases"]}
    signal_by_case = {}
    for case_id, value in native_cases.items():
        primary = [value["metrics"].get(f"{phase}.{quantile}", {}).get("pair_count", 0) for phase in ("edit_commit_ns", "total_ns") for quantile in ("p50", "p95")]
        rss = value["metrics"].get("max_rss_kib", {})
        primary_defined = any(count == 6 for count in primary)
        rss_defined = rss.get("pair_count") == 6 and rss.get("undefined_count", 1) == 0
        signal_by_case[case_id] = {
            "positive_primary_samples_validated": True,
            "primary_ratio_defined": primary_defined,
            "rss_ratio_defined": rss_defined,
            "confidence_interval_defined": False,
            "useful_signal": False,
            "reason": "screening has no confidence interval; a descriptive result cannot be treated as a useful qualification signal",
        }
    all_defined = all(item["primary_ratio_defined"] and item["rss_ratio_defined"] for item in signal_by_case.values())
    screening_decision = "followup_required_for_confirmation" if all_defined else "reject_no_useful_signal"
    result = {
        "schema": "litchi.performance.0964.screen-reduction.v3-screen-adapter",
        "status": "screened",
        "scope": "offline descriptive 0964 screening; no adoption or production claim",
        "plan": descriptor(plan_path),
        "plan_schema": raw.get("schema"),
        "canonical_mapping": mapping,
        "runner": binding["runner"],
        "packets": {"native": descriptor(native_packet / "capture-terminal.json"), "allocator": descriptor(allocator_packet / "capture-terminal.json")},
        "schedule": {"native_children": len(native_records), "allocator_children": len(allocator_records), "native_blocks_per_case": 3, "native_samples": 60, "native_warmups": 10, "allocator_repeats_per_case": 3, "allocator_rows": 10, "allocator_warmups": 2},
        "screening_policy": {"adoption_allowed": False, "full_followup_required": all_defined, "safety_review_ratio": v1.SCREEN_SAFETY, "native_statistic": "median of six paired screening legs", "allocator_statistic": "median of three paired screening repeats", "confidence_interval_required_for_useful_signal": True},
        "screening_decision": {"decision": screening_decision, "useful_signal": False, "by_case": signal_by_case},
        "native": {"cases": native_cases},
        "allocator": {"cases": allocator_cases},
        "limitations": ["Three native blocks and 60 samples are descriptive screening only; no bootstrap or adoption decision is emitted.", "Allocator observer peaks and max RSS remain distinct metrics.", "The protocol-plan mapping is structural only; all bound file descriptors and report/oracle validators remain fail-closed."],
    }
    encoded = (json.dumps(result, indent=2, sort_keys=True) + "\n").encode()
    require(not output.exists() and not output.is_symlink(), f"output already exists: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    fd = os.open(output, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o664)
    with os.fdopen(fd, "wb") as stream:
        stream.write(encoded)
        stream.flush()
        os.fsync(stream.fileno())
    return result


def main(argv: list[str] | None = None) -> int:
    parser = __import__("argparse").ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--native-packet", type=Path, required=True)
    parser.add_argument("--allocator-packet", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--helper", type=Path, default=Path(__file__))
    args = parser.parse_args(argv)
    try:
        result = reduce_packets(args.plan.resolve(), args.native_packet.resolve(), args.allocator_packet.resolve(), args.output.resolve(), args.helper.resolve())
        print(json.dumps({"status": result["status"], "output": str(args.output.resolve())}, sort_keys=True))
        return 0
    except (ReductionError, OSError, ValueError, KeyError) as exc:
        print(f"reduce_0964_v1: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
