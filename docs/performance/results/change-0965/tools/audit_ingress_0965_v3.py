#!/usr/bin/env python3
"""Independently audit one completed 0965 ingress diagnostic packet.

This helper deliberately does not import the 0965 reducer.  It validates the
sealed plan and custody descriptors, recomputes all five phase quantiles and
tails for all 24 children, recomputes the twelve pair summaries and the v3
diagnostic labels, and compares those results with the completed reduction.
It is read-only with respect to the packet and writes one immutable receipt.
It must be run only after the root runner and reducer have completed.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import os
import statistics
from pathlib import Path
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0965").resolve()
PLAN_DEFAULT = ROOT / "protocol/plan-final-v2.json"
PACKET_DEFAULT = ROOT / "capture-ingress-diagnostic-v1"
CPU_MASK = [12, 13, 14, 15]
WARMUP = 20
SAMPLES = 256
THRESHOLD = 1.05
CASES = ("medium-one-percent", "dense-sparse-one-percent", "medium-one-edit")
PAIR_SCHEDULE = {
    "medium-one-percent": ("RR", "RC", "CC", "CR"),
    "dense-sparse-one-percent": ("RC", "CC", "CR", "RR"),
    "medium-one-edit": ("CC", "CR", "RR", "RC"),
}
PHASES = ("ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns")
Q_NAMES = ("p50", "p95", "p99")
TIME_KEYS = {
    "Maximum resident set size (kbytes)": "max_rss_kib",
    "User time (seconds)": "user_seconds",
    "System time (seconds)": "system_seconds",
    "Major (requiring I/O) page faults": "major_page_faults",
    "Minor (reclaiming a frame) page faults": "minor_page_faults",
    "Voluntary context switches": "voluntary_context_switches",
    "Involuntary context switches": "involuntary_context_switches",
}
EXPECTED_BINARY_HASHES = {
    "R": "5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3",
    "C": "d21863a6613800391a01914865fc4eaecdf54ad3061447be8b6be7d47d0dc910",
}
RUNNER_SCHEMA = "litchi.performance.0965.diagnostic-capture.v1"
PLAN_SCHEMA = "litchi.performance.0965.ingress-diagnostic-plan.v1"
REDUCER_SCHEMA = "litchi.performance.0965.ingress-diagnostic-reduction.v1"
SCOPE_SCHEMA = "litchi.performance.0965.report-scope-contract.v1"


class AuditError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def read_json(path: Path) -> Any:
    path = Path(path).expanduser()
    require(not path.is_symlink(), f"symlink JSON is not accepted: {path}")
    require(path.is_file(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise AuditError(f"invalid JSON {path}: {exc}") from exc


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as exc:
        raise AuditError(f"cannot hash {path}: {exc}") from exc
    return digest.hexdigest()


def regular(path: Path, label: str) -> Path:
    path = Path(path).expanduser()
    require(not path.is_symlink(), f"{label} is a symlink: {path}")
    require(path.is_file(), f"{label} is not a regular file: {path}")
    return path.resolve()


def descriptor(path: Path, executable: bool | None = None, label: str = "artifact") -> dict[str, Any]:
    path = regular(path, label)
    mode = path.stat().st_mode & 0o777
    is_executable = bool(mode & 0o111)
    if executable is not None:
        require(is_executable is executable, f"{label} executable bit changed")
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "mode": mode,
        "executable": is_executable,
        "sha256": sha256(path),
    }


def check_descriptor(value: Any, path: Path, label: str, executable: bool | None = None) -> dict[str, Any]:
    actual = descriptor(path, executable, label)
    require(isinstance(value, dict), f"{label} descriptor missing")
    for key in ("path", "bytes", "mode", "executable", "sha256"):
        require(value.get(key) == actual[key], f"{label}.{key} changed")
    return actual


def check_compact_descriptor(value: Any, path: Path, label: str) -> dict[str, Any]:
    """Check the compact path/size/hash binding used by root confirmation."""
    actual = descriptor(path, label=label)
    require(isinstance(value, dict), f"{label} compact descriptor missing")
    for key in ("path", "bytes", "sha256"):
        require(value.get(key) == actual[key], f"{label}.{key} changed")
    return actual


def check_dir(path: Path, label: str) -> Path:
    path = Path(path).expanduser()
    require(not path.is_symlink() and path.is_dir(), f"{label} directory missing or symlink: {path}")
    return path.resolve()


def iso_seconds(value: Any, label: str) -> float:
    require(isinstance(value, str), f"{label} timestamp missing")
    try:
        result = dt.datetime.fromisoformat(value).timestamp()
    except ValueError as exc:
        raise AuditError(f"{label} timestamp invalid: {value!r}") from exc
    require(math.isfinite(result), f"{label} timestamp is not finite")
    return result


def expected_rows(plan: dict[str, Any]) -> list[dict[str, Any]]:
    cases = {case["case_id"]: case for case in plan["cases"]}
    rows: list[dict[str, Any]] = []
    for case_id in CASES:
        for pair_index, pair_type in enumerate(PAIR_SCHEDULE[case_id]):
            pair_id = f"{case_id}__pair{pair_index:02d}__{pair_type}"
            for position, role in enumerate(pair_type):
                case = cases[case_id]
                rows.append({
                    "child_id": f"native__{pair_id}__p{position}",
                    "lane": "native",
                    "case_id": case_id,
                    "pair_id": pair_id,
                    "pair_type": pair_type,
                    "pair_index": pair_index,
                    "position": position,
                    "role": role,
                    "binary_role": "baseline" if role == "R" else "candidate",
                    "comparison_id": "retained-baseline-vs-retained-candidate",
                    "shape": case["shape"],
                    "workload": case["workload"],
                    "workers": 1,
                    "warmup_iterations": WARMUP,
                    "measured_samples": SAMPLES,
                    "cpu_mask": list(CPU_MASK),
                })
    return rows


def expected_pairs() -> list[dict[str, Any]]:
    result = []
    for case_id in CASES:
        for pair_index, pair_type in enumerate(PAIR_SCHEDULE[case_id]):
            pair_id = f"{case_id}__pair{pair_index:02d}__{pair_type}"
            result.append({
                "pair_id": pair_id,
                "case_id": case_id,
                "pair_index": pair_index,
                "pair_type": pair_type,
                "children": [f"native__{pair_id}__p0", f"native__{pair_id}__p1"],
                "roles": list(pair_type),
                "adjacent": True,
            })
    return result


def nearest_rank(values: list[int], quantile: float) -> int:
    ordered = sorted(values)
    require(ordered, "empty sample vector")
    rank = max(1, math.ceil(quantile * len(ordered)))
    return ordered[rank - 1]


def compact(value: float | int) -> float | int:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def child_stats(values: list[int]) -> dict[str, Any]:
    require(len(values) == SAMPLES, "unexpected sample count")
    ordered = sorted(((value, index) for index, value in enumerate(values)), key=lambda item: (-item[0], item[1]))
    top = ordered[:5]
    quantiles = {
        "p50": nearest_rank(values, 0.50),
        "p95": nearest_rank(values, 0.95),
        "p99": nearest_rank(values, 0.99),
    }
    tail = [
        {"sample": index, "value": value}
        for value, index in ordered
        if value >= quantiles["p99"]
    ]
    return {
        "quantiles_ns": quantiles,
        "max_ns": max(values),
        "mean_ns": compact(statistics.fmean(values)),
        "p99_sample_indices": [index for value, index in ordered if value == quantiles["p99"]],
        "top5_mean_ns": compact(statistics.fmean(value for value, _ in top)),
        "top5_sample_indices": [index for _, index in top],
        "top5_values_ns": [value for value, _ in top],
        "p99_and_above_tail": tail,
    }


def parse_time(path: Path, label: str) -> dict[str, int | float]:
    values: dict[str, int | float] = {}
    for line in regular(path, label).read_text(errors="replace").splitlines():
        if ":" not in line:
            continue
        key, raw = line.split(":", 1)
        name = TIME_KEYS.get(key.strip())
        if name is None:
            continue
        token = raw.strip().split()[0] if raw.strip() else ""
        try:
            value: int | float = float(token) if name in {"user_seconds", "system_seconds"} else int(token)
        except ValueError as exc:
            raise AuditError(f"{label} metric {name} is not numeric") from exc
        require(math.isfinite(float(value)) and value >= 0, f"{label} metric {name} invalid")
        values[name] = value
    require(set(values) >= set(TIME_KEYS.values()), f"{label} is missing /usr/bin/time metrics")
    require(type(values["max_rss_kib"]) is int and values["max_rss_kib"] > 0, f"{label} RSS invalid")
    return values


def check_host(path: Path, mapping: dict[str, list[str]], label: str) -> dict[str, Any]:
    value = read_json(path)
    require(isinstance(value.get("created_utc"), str) and isinstance(value.get("scope"), str), f"{label} metadata missing")
    for category, keys in mapping.items():
        require(isinstance(keys, list) and keys, f"{label} mapping for {category} missing")
        for key in keys:
            require(key in value, f"{label} missing mapped field {category}:{key}")
    require(isinstance(value.get("scheduling_settings"), dict), f"{label} scheduling settings malformed")
    return {
        "descriptor": descriptor(path, label=label),
        "created_utc": value["created_utc"],
        "scope": value["scope"],
        "keys": sorted(value),
    }


def check_plan(plan_path: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    plan_path = regular(plan_path, "plan")
    plan = read_json(plan_path)
    require(plan.get("schema") == PLAN_SCHEMA and plan.get("status") == "authorized", "final authorized plan required")
    require(plan.get("plan_id") == "xlsx-ingress-tail-diagnostic-0965", "plan id changed")
    require(plan.get("scope", {}).get("children") == 24 and plan.get("scope", {}).get("pairs") == 12, "plan child/pair scope changed")
    require(plan.get("scope", {}).get("warmup_iterations") == WARMUP and plan.get("scope", {}).get("measured_samples") == SAMPLES, "plan sample scope changed")
    require(plan.get("scope", {}).get("serial") is True and plan.get("scope", {}).get("retry_or_replacement") is False, "plan serial/retry policy changed")
    require(plan.get("analysis", {}).get("adoption_allowed") is False and plan.get("analysis", {}).get("screen_reopened") is False, "plan decision policy changed")
    require(plan.get("diagnostic_rule", {}).get("frozen_0964_screen_rule_relaxed") is False, "0964 screen rule was relaxed")
    require(plan.get("diagnostic_rule", {}).get("labels") == ["concordant", "nonreproduced", "mixed"], "diagnostic labels changed")
    require(plan.get("diagnostic_rule", {}).get("threshold") == THRESHOLD, "diagnostic threshold changed")

    plan_binding = descriptor(plan_path, label="plan")
    runner_binding = check_descriptor(plan["runner"]["descriptor"], Path(plan["runner"]["descriptor"]["path"]), "runner", True)
    reducer_binding = check_descriptor(plan["reducer"]["descriptor"], Path(plan["reducer"]["descriptor"]["path"]), "reducer", True)
    design_binding = check_descriptor(plan["design"]["descriptor"], Path(plan["design"]["descriptor"]["path"]), "design")
    require(plan["design"].get("status") == "sealed" and plan["design"].get("schema") == "prospective-design-v3", "sealed design binding changed")

    native_contract = plan.get("report_contract", {}).get("native", {})
    require(native_contract.get("schema") == "litchi.xlsx-managed-compression-benchmark.v1", "report schema changed")
    require(native_contract.get("warmup_iterations") == WARMUP and native_contract.get("measured_samples") == SAMPLES and native_contract.get("raw_sample_count") == SAMPLES, "report sample contract changed")
    require(native_contract.get("sample_phase_fields") == list(PHASES), "report phases changed")
    contract_binding = check_descriptor(plan["report_scope_contract"], Path(plan["report_scope_contract"]["path"]), "report scope contract")
    scope_contract = read_json(Path(contract_binding["path"]))
    require(scope_contract.get("schema") == SCOPE_SCHEMA and scope_contract.get("status") == "sealed", "scope contract status/schema changed")
    scope_values = scope_contract.get("shared_exact_values")
    require(isinstance(scope_values, dict), "scope contract values missing")
    require(native_contract.get("independent_postcapture_exact_scope_values") == scope_values, "plan scope values differ from contract")
    require(set(scope_values) == {"timing_scope", "phase_scope", "setup_scope", "verification_scope", "performance_claim"}, "scope contract field set changed")
    scope_roles = scope_contract.get("roles")
    require(isinstance(scope_roles, list) and len(scope_roles) == 2, "scope contract role bindings missing")
    require(all(role.get("exact_values") == scope_values for role in scope_roles), "scope role values differ")

    mapping = plan.get("capture", {}).get("host_snapshot_required", {}).get("field_mapping")
    required_fields = plan.get("capture", {}).get("host_snapshot_required", {}).get("fields")
    require(isinstance(mapping, dict) and set(mapping) == set(required_fields), "host field mapping changed")
    host_before_binding = plan["capture"]["host_snapshot_before"]
    host_before = check_host(Path(host_before_binding["path"]), mapping, "host-before")
    host_after_path = Path(plan["capture"]["host_snapshot_after"]["path"])
    require(plan["capture"]["host_snapshot_after"].get("status") == "required_after_terminal", "host-after policy changed")

    controls_binding = check_descriptor(plan["affinity_controls"]["descriptor"], Path(plan["affinity_controls"]["descriptor"]["path"]), "affinity controls")
    controls = read_json(Path(controls_binding["path"]))
    require(controls.get("schema") == "litchi.performance.0965.affinity-launch-controls.v1" and controls.get("status") == "pass", "affinity controls status/schema changed")
    require(controls.get("workload_started") is False and controls.get("cpu_mask") == CPU_MASK, "affinity controls workload/mask changed")
    require(controls.get("runner") == runner_binding, "affinity control runner binding changed")
    require(controls.get("positive", {}).get("observed_affinity") == CPU_MASK and controls["positive"].get("guard_passed") is True, "positive affinity control failed")
    require(controls.get("negative", {}).get("observed_affinity") != CPU_MASK and controls["negative"].get("guard_rejected") is True, "negative affinity control failed")

    binaries = {}
    for role in ("R", "C"):
        binding = plan["binaries"][role]
        binaries[role] = check_descriptor(binding, Path(binding["path"]), f"binary {role}", True)
        require(binaries[role]["sha256"] == EXPECTED_BINARY_HASHES[role], f"binary {role} hash changed")
    source_bindings = {}
    for key, value in plan.get("source", {}).items():
        if isinstance(value, dict) and "path" in value:
            source_bindings[key] = check_descriptor(value, Path(value["path"]), f"source.{key}")
    for key in ("baseline_source_root", "candidate_source_root"):
        check_dir(Path(plan["source"][key]), f"source.{key}")
    inherited = plan.get("inherited_0964", {})
    inherited_bindings = {
        key: check_descriptor(inherited[key], Path(inherited[key]["path"]), f"inherited.{key}")
        for key in ("plan", "raw_audit")
    }
    require(inherited.get("rejection_remains_binding") is True, "0964 rejection policy changed")

    rows = expected_rows(plan)
    require(plan.get("schedule", {}).get("children") == rows, "child schedule changed")
    require(plan.get("schedule", {}).get("pairs") == expected_pairs(), "pair schedule changed")
    require(plan.get("analysis", {}).get("packet_wide_pair_counts") == {"RR": 3, "RC": 3, "CC": 3, "CR": 3}, "packet pair counts changed")
    require(plan.get("analysis", {}).get("cross_binary_pairs_per_case") == {"RC": 1, "CR": 1}, "cross-order pair counts changed")

    root_go = plan.get("root_go", {})
    require(root_go.get("status") == "GO", "root GO status missing")
    go_binding = check_descriptor(root_go["descriptor"], Path(root_go["descriptor"]["path"]), "root GO")
    go = read_json(Path(go_binding["path"]))
    require(go.get("status") == "GO", "root GO receipt status changed")
    confirmation_path = ROOT / "protocol/root-final-confirmation-v1.json"
    confirmation = None
    if confirmation_path.exists():
        confirmation = read_json(confirmation_path)
        require(confirmation.get("status") == "GO", "root final confirmation status changed")
        check_compact_descriptor(confirmation.get("final_plan"), plan_path, "root final confirmation plan")
        check_compact_descriptor(confirmation.get("mandatory_scope_contract"), Path(contract_binding["path"]), "root final confirmation scope contract")
        check_compact_descriptor(confirmation.get("host_before"), Path(host_before_binding["path"]), "root final confirmation host")

    return plan, {
        "plan": plan_binding,
        "runner": runner_binding,
        "reducer": reducer_binding,
        "design": design_binding,
        "scope_contract": contract_binding,
        "host_before": host_before,
        "host_after_path": str(host_after_path.resolve()),
        "controls": controls_binding,
        "binaries": binaries,
        "source": source_bindings,
        "inherited": inherited_bindings,
        "root_go": go_binding,
        "root_confirmation": descriptor(confirmation_path, label="root confirmation") if confirmation is not None else None,
    }


def expected_environment(plan: dict[str, Any]) -> dict[str, str]:
    configured = plan["execution"]["environment"]
    result = {key: str(value) for key, value in configured.items() if key != "approval_variable"}
    result[str(configured["approval_variable"])] = "1"
    return result


def check_case_report(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any], scope_values: dict[str, str]) -> list[dict[str, Any]]:
    native = plan["report_contract"]["native"]
    require(report.get("schema") == native["schema"] and report.get("status") == "completed", f"{row['child_id']} report schema/status")
    for key, wanted in (("shape", row["shape"]), ("workload", row["workload"]), ("workers", 1), ("warmup_iterations", WARMUP), ("measured_samples", SAMPLES)):
        require(report.get(key) == wanted, f"{row['child_id']} report root {key}")
    for key, wanted in scope_values.items():
        require(report.get(key) == wanted, f"{row['child_id']} report scope {key}")
    case_cfg = next(case for case in plan["cases"] if case["case_id"] == row["case_id"])
    case = report.get("case")
    require(isinstance(case, dict), f"{row['child_id']} case missing")
    expected = {
        "id": f"{row['shape']}-{row['workload']}-workers1",
        "shape": row["shape"], "workload": row["workload"], "workers": 1,
        "update_count": case_cfg["update_count"], "source_bytes": case_cfg["source_bytes"], "source_sha256": case_cfg["source_sha256"],
        "output_bytes": case_cfg["output_bytes"], "output_sha256": case_cfg["output_sha256"], "semantic_sha256": case_cfg["semantic_sha256"],
        "expected_cpu_tasks": case_cfg["expected_cpu_tasks"], "admission_expected_worker_permits": 0,
        "admission_observed_worker_permits": 0, "admission_cpu_tasks_used": case_cfg["expected_cpu_tasks"],
        "preflight_actual_commit_inverse_verified": True, "context_limits": case_cfg["context_limits"],
        "sink_capacity_bytes": case_cfg["context_limits"]["output_bytes"], "sink_max_write_bytes": 65536,
        "warmup_iterations": WARMUP, "measured_samples": SAMPLES,
    }
    for key, wanted in expected.items():
        require(case.get(key) == wanted, f"{row['child_id']} case oracle {key}")
    samples = case.get("samples")
    require(isinstance(samples, list) and len(samples) == SAMPLES, f"{row['child_id']} samples")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("sample") == index, f"{row['child_id']} sample index {index}")
        values = []
        for phase in PHASES:
            value = sample.get(phase)
            require(type(value) is int and value > 0, f"{row['child_id']} {phase} value")
            values.append(value)
        require(sum(values[:4]) <= values[4], f"{row['child_id']} phase sum")
        for key, wanted in (("output_bytes", case_cfg["output_bytes"]), ("output_bytes_used", case_cfg["output_bytes"]), ("cpu_tasks_used", case_cfg["expected_cpu_tasks"]), ("output_sha256", case_cfg["output_sha256"]), ("semantic_sha256", case_cfg["semantic_sha256"])):
            require(sample.get(key) == wanted, f"{row['child_id']} sample oracle {key}")
        require(sample.get("fixture_inverse_restored_exact_source") is True, f"{row['child_id']} inverse oracle")
        require(type(sample.get("sink_write_calls")) is int and sample["sink_write_calls"] > 0, f"{row['child_id']} sink calls")
        require(type(sample.get("sink_largest_write")) is int and 0 < sample["sink_largest_write"] <= 65536, f"{row['child_id']} sink largest write")
        resources = sample.get("resources_after_drop")
        require(isinstance(resources, dict) and all(resources.get(key) == 0 for key in ("memory", "workers", "io_concurrency", "objects")), f"{row['child_id']} cleanup resources")
    return samples


def check_child(packet: Path, row: dict[str, Any], plan: dict[str, Any], plan_binding: dict[str, Any], scope_values: dict[str, str], index: int) -> dict[str, Any]:
    child = check_dir(packet / "children" / row["child_id"], row["child_id"])
    output = check_dir(child / "output", f"{row['child_id']} output")
    paths = {
        "start": child / "start.json",
        "launch": child / "launch.json",
        "terminal": child / "terminal.json",
        "validation": child / "validation.json",
        "report": output / "report.json",
        "time": child / "time-v.txt",
        "post": child / "post-terminal-artifacts.json",
    }
    values = {key: read_json(path) for key, path in paths.items() if key != "time"}
    start = values["start"]
    require(start.get("schema") == RUNNER_SCHEMA + ".child-start" and start.get("index") == index, f"{row['child_id']} start schema/index")
    require(start.get("child") == row and start.get("plan") == plan_binding, f"{row['child_id']} start binding")
    binary_key = "R" if row["role"] == "R" else "C"
    binary_path = Path(plan["binaries"][binary_key]["path"]).resolve()
    require(start.get("binary") == check_descriptor(plan["binaries"][binary_key], binary_path, f"{row['child_id']} binary", True), f"{row['child_id']} binary binding")
    expected_output = output
    direct = [
        "/usr/bin/taskset", "--cpu-list", "12-15", str(binary_path),
        "--shape", row["shape"], "--workload", row["workload"],
        "--workers", "1", "--warmup", str(WARMUP), "--samples", str(SAMPLES),
        "--output", str(expected_output),
    ]
    wrapped = [
        "/usr/bin/taskset", "--cpu-list", "12-15", "/usr/bin/time", "-v",
        "-o", str(child / "time-v.txt"), "--", *direct[3:],
    ]
    require(start.get("argv") == wrapped and start.get("workload_argv") == direct, f"{row['child_id']} argv custody")
    require(start.get("cwd") == plan["execution"]["workload_cwd"], f"{row['child_id']} cwd changed")
    require(start.get("environment") == expected_environment(plan), f"{row['child_id']} environment changed")
    require(start.get("deadline_seconds") == plan["execution"]["deadline_seconds"] and start.get("output_absent_before_launch") is True, f"{row['child_id']} launch contract")
    start_seconds = iso_seconds(start.get("started_utc"), f"{row['child_id']} start")

    launch = values["launch"]
    require(launch.get("schema") == RUNNER_SCHEMA + ".child-launch", f"{row['child_id']} launch schema")
    require(launch.get("observed_affinity") == CPU_MASK and launch.get("expected_affinity") == CPU_MASK and launch.get("affinity_verified_before_wait") is True, f"{row['child_id']} affinity")
    require(type(launch.get("pid")) is int and launch.get("pid") > 0, f"{row['child_id']} pid")

    terminal = values["terminal"]
    require(terminal.get("schema") == RUNNER_SCHEMA + ".child-terminal" and terminal.get("index") == index and terminal.get("child_id") == row["child_id"], f"{row['child_id']} terminal identity")
    require(terminal.get("status") == "completed" and terminal.get("return_code") == 0 and terminal.get("timed_out") is False and terminal.get("process_group_clear") is True, f"{row['child_id']} terminal status")
    require(terminal.get("terminal_before_report_parse") is True and terminal.get("pid") == launch["pid"], f"{row['child_id']} terminal order/pid")
    require(terminal.get("output", {}).get("path") == str(output) and terminal.get("output", {}).get("exists_after_terminal") is True, f"{row['child_id']} terminal output")
    terminal_seconds = iso_seconds(terminal.get("ended_utc"), f"{row['child_id']} terminal")
    require(terminal_seconds >= start_seconds, f"{row['child_id']} terminal precedes start")
    run = terminal.get("run", {})
    require(run.get("pid") == launch["pid"] and run.get("process_group_id") == launch["pid"], f"{row['child_id']} process identity")
    require(run.get("reaped") is True and run.get("process_group_clear") is True and run.get("affinity_verified") is True, f"{row['child_id']} cleanup/affinity")

    validation = values["validation"]
    require(validation.get("schema") == RUNNER_SCHEMA + ".child-validation" and validation.get("status") == "passed" and validation.get("index") == index and validation.get("child") == row, f"{row['child_id']} validation")
    require(validation.get("plan") == plan_binding, f"{row['child_id']} validation plan")
    require(validation.get("terminal") == check_descriptor(validation.get("terminal"), paths["terminal"], f"{row['child_id']} validation terminal"), f"{row['child_id']} validation terminal binding")
    require(validation.get("report") == check_descriptor(validation.get("report"), paths["report"], f"{row['child_id']} validation report"), f"{row['child_id']} validation report binding")
    require(validation.get("time") == check_descriptor(validation.get("time"), paths["time"], f"{row['child_id']} validation time"), f"{row['child_id']} validation time binding")
    time_metrics = parse_time(paths["time"], f"{row['child_id']} time")
    require(validation.get("time_metrics") == time_metrics, f"{row['child_id']} time metrics")
    samples = check_case_report(values["report"], row, plan, scope_values)

    post = values["post"]
    require(post.get("schema") == RUNNER_SCHEMA + ".post-terminal-artifacts", f"{row['child_id']} post-terminal schema")
    for key, path in (("terminal", paths["terminal"]), ("stdout", child / "stdout.log"), ("stderr", child / "stderr.log"), ("time", paths["time"]), ("report", paths["report"]), ("validation", paths["validation"])):
        require(post.get(key) == check_descriptor(post.get(key), path, f"{row['child_id']} post {key}"), f"{row['child_id']} post {key} binding")
    stats = {phase: child_stats([sample[phase] for sample in samples]) for phase in PHASES}
    return {
        "row": row,
        "child_path": str(child),
        "start": start,
        "terminal": terminal,
        "report": values["report"],
        "samples": samples,
        "stats": stats,
        "time_metrics": time_metrics,
        "start_seconds": start_seconds,
        "terminal_seconds": terminal_seconds,
        "report_descriptor": descriptor(paths["report"], label=f"{row['child_id']} report"),
    }


def finite_ratio(numerator: float, denominator: float) -> float:
    require(denominator > 0, "ratio denominator must be positive")
    value = numerator / denominator
    require(math.isfinite(value) and value >= 0, "ratio is not finite")
    return value


def pair_summary(first: dict[str, Any], second: dict[str, Any], pair: dict[str, Any]) -> dict[str, Any]:
    cross = pair["pair_type"] in {"RC", "CR"}
    phases: dict[str, Any] = {}
    for phase in PHASES:
        first_stats, second_stats = first["stats"][phase], second["stats"][phase]
        first_q, second_q = first_stats["quantiles_ns"], second_stats["quantiles_ns"]
        if cross:
            baseline = first if first["row"]["role"] == "R" else second
            candidate = second if first["row"]["role"] == "R" else first
            baseline_q = baseline["stats"][phase]["quantiles_ns"]
            candidate_q = candidate["stats"][phase]["quantiles_ns"]
            ratios = {name: finite_ratio(candidate_q[name], baseline_q[name]) for name in Q_NAMES}
            baseline_top = baseline["stats"][phase]["top5_mean_ns"]
            candidate_top = candidate["stats"][phase]["top5_mean_ns"]
            phases[phase] = {
                "baseline_child": baseline["row"]["child_id"],
                "candidate_child": candidate["row"]["child_id"],
                "baseline_quantiles_ns": baseline_q,
                "candidate_quantiles_ns": candidate_q,
                "candidate_over_baseline_ratio": ratios,
                "candidate_minus_baseline_ns": {name: candidate_q[name] - baseline_q[name] for name in Q_NAMES},
                "baseline_top5_mean_ns": baseline_top,
                "candidate_top5_mean_ns": candidate_top,
                "candidate_over_baseline_top5_ratio": finite_ratio(candidate_top, baseline_top),
                "candidate_p99_above_1_05": ratios["p99"] > THRESHOLD,
                "candidate_p95_above_1_05": ratios["p95"] > THRESHOLD,
                "candidate_top5_above_1_05": finite_ratio(candidate_top, baseline_top) > THRESHOLD,
                "baseline_p99_sample_indices": baseline["stats"][phase]["p99_sample_indices"],
                "candidate_p99_sample_indices": candidate["stats"][phase]["p99_sample_indices"],
                "baseline_top5_sample_indices": baseline["stats"][phase]["top5_sample_indices"],
                "candidate_top5_sample_indices": candidate["stats"][phase]["top5_sample_indices"],
            }
        else:
            directional = {name: finite_ratio(second_q[name], first_q[name]) for name in Q_NAMES}
            absolute = {name: max(directional[name], finite_ratio(first_q[name], second_q[name])) for name in Q_NAMES}
            first_top, second_top = first_stats["top5_mean_ns"], second_stats["top5_mean_ns"]
            top_directional = finite_ratio(second_top, first_top)
            top_absolute = max(top_directional, finite_ratio(first_top, second_top))
            phases[phase] = {
                "first_child": first["row"]["child_id"],
                "second_child": second["row"]["child_id"],
                "first_quantiles_ns": first_q,
                "second_quantiles_ns": second_q,
                "second_over_first_ratio": directional,
                "absolute_ratio": absolute,
                "first_top5_mean_ns": first_top,
                "second_top5_mean_ns": second_top,
                "second_over_first_top5_ratio": top_directional,
                "absolute_top5_ratio": top_absolute,
                "absolute_p99_above_1_05": absolute["p99"] > THRESHOLD,
                "absolute_p95_above_1_05": absolute["p95"] > THRESHOLD,
                "absolute_top5_above_1_05": top_absolute > THRESHOLD,
                "first_p99_sample_indices": first_stats["p99_sample_indices"],
                "second_p99_sample_indices": second_stats["p99_sample_indices"],
                "first_top5_sample_indices": first_stats["top5_sample_indices"],
                "second_top5_sample_indices": second_stats["top5_sample_indices"],
            }
    return {
        "pair_id": pair["pair_id"],
        "case_id": pair["case_id"],
        "pair_type": pair["pair_type"],
        "children": [first["row"]["child_id"], second["row"]["child_id"]],
        "phases": phases,
    }


def classify(case_id: str, pairs: dict[str, dict[str, Any]]) -> dict[str, Any]:
    rc = pairs["RC"]["phases"]["ingress_ns"]["candidate_over_baseline_ratio"]["p99"]
    cr = pairs["CR"]["phases"]["ingress_ns"]["candidate_over_baseline_ratio"]["p99"]
    rr = pairs["RR"]["phases"]["ingress_ns"]["absolute_ratio"]["p99"]
    cc = pairs["CC"]["phases"]["ingress_ns"]["absolute_ratio"]["p99"]
    rc_p95 = pairs["RC"]["phases"]["ingress_ns"]["candidate_over_baseline_ratio"]["p95"]
    cr_p95 = pairs["CR"]["phases"]["ingress_ns"]["candidate_over_baseline_ratio"]["p95"]
    rc_top5 = pairs["RC"]["phases"]["ingress_ns"]["candidate_over_baseline_top5_ratio"]
    cr_top5 = pairs["CR"]["phases"]["ingress_ns"]["candidate_over_baseline_top5_ratio"]
    flags = {
        "RC_p99_above_1_05": rc > THRESHOLD,
        "CR_p99_above_1_05": cr > THRESHOLD,
        "RR_absolute_p99_above_1_05": rr > THRESHOLD,
        "CC_absolute_p99_above_1_05": cc > THRESHOLD,
    }
    controls_stable = rr <= THRESHOLD and cc <= THRESHOLD
    if flags["RC_p99_above_1_05"] and flags["CR_p99_above_1_05"] and controls_stable:
        label = "concordant"
    elif not flags["RC_p99_above_1_05"] and not flags["CR_p99_above_1_05"] and controls_stable:
        label = "nonreproduced"
    else:
        label = "mixed"
    return {
        "case_id": case_id,
        "label": label,
        "threshold": THRESHOLD,
        "cross_order_p99_ratios": {"RC": rc, "CR": cr},
        "cross_order_p95_ratios": {"RC": rc_p95, "CR": cr_p95},
        "cross_order_top5_ratios": {"RC": rc_top5, "CR": cr_top5},
        "cross_order_spread": max(rc, cr) / min(rc, cr),
        "same_binary_absolute_p99_ratios": {"RR": rr, "CC": cc},
        "flags": flags,
        "diagnostic_only": True,
        "source_build_host_causal_attribution": False,
    }


def same_value(left: Any, right: Any) -> bool:
    if isinstance(left, bool) or isinstance(right, bool):
        return left == right
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return math.isclose(float(left), float(right), rel_tol=1e-13, abs_tol=1e-12)
    if isinstance(left, dict) and isinstance(right, dict):
        return set(left) == set(right) and all(same_value(left[key], right[key]) for key in left)
    if isinstance(left, list) and isinstance(right, list):
        return len(left) == len(right) and all(same_value(a, b) for a, b in zip(left, right))
    return left == right


def compare_reducer(own: dict[str, Any], reduced: dict[str, Any], bindings: dict[str, Any], packet: Path) -> list[dict[str, Any]]:
    mismatches: list[dict[str, Any]] = []
    require(reduced.get("schema") == REDUCER_SCHEMA and reduced.get("status") == "descriptive_read_only", "reduction schema/status")
    decision = reduced.get("decision", {})
    require(decision.get("adoption_allowed") is False and decision.get("screen_reopened") is False and decision.get("causal_claim") is False and decision.get("0964_rejection_preserved") is True, "reduction decision policy")
    require(reduced.get("bindings", {}).get("plan") == bindings["plan"], "reduction plan binding")
    require(reduced.get("bindings", {}).get("runner") == bindings["runner"], "reduction runner binding")
    require(reduced.get("bindings", {}).get("reducer") == bindings["reducer"], "reduction reducer binding")
    require(reduced.get("bindings", {}).get("completed_packet") == descriptor(packet / "completed-packet.json", label="completed packet"), "reduction completed packet binding")
    child_by_id = {item["child_id"]: item for item in reduced.get("child_summaries", [])}
    require(len(child_by_id) == 24, "reduction child summary count")
    for child in own["children"]:
        child_id = child["row"]["child_id"]
        reduced_child = child_by_id.get(child_id)
        require(reduced_child is not None, f"reduction missing child {child_id}")
        for phase in PHASES:
            own_stats = child["stats"][phase]
            reduced_stats = reduced_child.get("stats", {}).get(phase, {})
            for key in ("quantiles_ns", "max_ns", "mean_ns", "p99_sample_indices", "top5_mean_ns", "top5_sample_indices"):
                if not same_value(own_stats[key], reduced_stats.get(key)):
                    mismatches.append({"kind": "child_stat", "child_id": child_id, "phase": phase, "field": key, "independent": own_stats[key], "reduced": reduced_stats.get(key)})
    reduced_pairs = reduced.get("pairs", {})
    for pair_id, own_pair in own["pairs"].items():
        reduced_pair = reduced_pairs.get(pair_id)
        require(reduced_pair is not None, f"reduction missing pair {pair_id}")
        for phase in PHASES:
            if own_pair["pair_type"] in {"RC", "CR"}:
                fields = (
                    ("baseline_child", "baseline_child"), ("candidate_child", "candidate_child"),
                    ("baseline_quantiles_ns", "baseline_quantiles_ns"), ("candidate_quantiles_ns", "candidate_quantiles_ns"),
                    ("candidate_over_baseline_ratio", "candidate_over_baseline_ratio"), ("candidate_minus_baseline_ns", "candidate_minus_baseline_ns"),
                    ("baseline_top5_mean_ns", "baseline_top5_mean_ns"), ("candidate_top5_mean_ns", "candidate_top5_mean_ns"),
                    ("candidate_over_baseline_top5_ratio", "candidate_over_baseline_top5_ratio"),
                    ("candidate_p99_above_1_05", "candidate_p99_above_1_05"), ("candidate_p95_above_1_05", "candidate_p95_above_1_05"),
                    ("candidate_top5_above_1_05", "candidate_top5_above_1_05"),
                    ("baseline_p99_sample_indices", "baseline_p99_sample_indices"), ("candidate_p99_sample_indices", "candidate_p99_sample_indices"),
                    ("baseline_top5_sample_indices", "baseline_top5_sample_indices"), ("candidate_top5_sample_indices", "candidate_top5_sample_indices"),
                )
            else:
                fields = (
                    ("first_child", "first_child"), ("second_child", "second_child"),
                    ("first_quantiles_ns", "first_quantiles_ns"), ("second_quantiles_ns", "second_quantiles_ns"),
                    ("second_over_first_ratio", "second_over_first_ratio"), ("absolute_ratio", "absolute_ratio"),
                    ("first_top5_mean_ns", "first_top5_mean_ns"), ("second_top5_mean_ns", "second_top5_mean_ns"),
                    ("second_over_first_top5_ratio", "second_over_first_top5_ratio"), ("absolute_top5_ratio", "absolute_top5_ratio"),
                    ("absolute_p99_above_1_05", "absolute_p99_above_1_05"), ("absolute_p95_above_1_05", "absolute_p95_above_1_05"),
                    ("absolute_top5_above_1_05", "absolute_top5_above_1_05"),
                    ("first_p99_sample_indices", "first_p99_sample_indices"), ("second_p99_sample_indices", "second_p99_sample_indices"),
                    ("first_top5_sample_indices", "first_top5_sample_indices"), ("second_top5_sample_indices", "second_top5_sample_indices"),
                )
            for own_field, reduced_field in fields:
                if not same_value(own_pair["phases"][phase].get(own_field), reduced_pair.get("phases", {}).get(phase, {}).get(reduced_field)):
                    mismatches.append({"kind": "pair", "pair_id": pair_id, "phase": phase, "field": own_field, "independent": own_pair["phases"][phase].get(own_field), "reduced": reduced_pair.get("phases", {}).get(phase, {}).get(reduced_field)})
    reduced_cases = reduced.get("case_results", {})
    for case_id, own_case in own["case_results"].items():
        reduced_classification = reduced_cases.get(case_id, {}).get("classification")
        require(isinstance(reduced_classification, dict), f"reduction classification missing {case_id}")
        if not same_value(own_case, reduced_classification):
            mismatches.append({"kind": "classification", "case_id": case_id, "independent": own_case, "reduced": reduced_classification})
    return mismatches


def audit(plan_path: Path, packet_path: Path, reduced_path: Path) -> dict[str, Any]:
    plan, bindings = check_plan(plan_path)
    packet = check_dir(packet_path, "capture packet")
    require(packet != ROOT / "capture-native-epoch2-v1", "historical packet is not allowed")
    scope_values = json.loads(Path(plan["report_scope_contract"]["path"]).read_text())["shared_exact_values"]
    capture_start_path = packet / "capture-start.json"
    capture_terminal_path = packet / "capture-terminal.json"
    completed_path = packet / "completed-packet.json"
    capture_start = read_json(capture_start_path)
    require(capture_start.get("schema") == RUNNER_SCHEMA + ".capture-start" and capture_start.get("status") == "started" and capture_start.get("lane") == "native", "capture start")
    require(capture_start.get("plan") == bindings["plan"] and capture_start.get("runner") == bindings["runner"], "capture start bindings")
    require(capture_start.get("serial") is True and capture_start.get("children_expected") == 24 and capture_start.get("warmup_iterations") == WARMUP and capture_start.get("measured_samples") == SAMPLES and capture_start.get("retry_or_replacement") is False, "capture start schedule")
    require(capture_start.get("host_snapshot_before") == plan["capture"]["host_snapshot_before"], "capture start host-before binding")
    capture_terminal = read_json(capture_terminal_path)
    require(capture_terminal.get("schema") == RUNNER_SCHEMA + ".capture-terminal" and capture_terminal.get("status") == "completed", "capture terminal")
    require(capture_terminal.get("children_started") == 24 and capture_terminal.get("children_expected") == 24 and capture_terminal.get("children_completed") == 24 and capture_terminal.get("retry_or_replacement") is False and capture_terminal.get("terminal_before_parse") is True, "capture terminal counts/policy")
    require(capture_terminal.get("capture_start") == descriptor(capture_start_path, label="capture start"), "capture terminal/start binding")
    completed = read_json(completed_path)
    require(completed.get("schema") == RUNNER_SCHEMA + ".completed-packet" and completed.get("status") == "completed", "completed packet")
    require(completed.get("plan") == bindings["plan"] and completed.get("runner") == bindings["runner"], "completed packet bindings")
    require(completed.get("capture_terminal") == descriptor(capture_terminal_path, label="capture terminal"), "completed terminal binding")

    host_after_path = Path(plan["capture"]["host_snapshot_after"]["path"])
    host_after = check_host(host_after_path, plan["capture"]["host_snapshot_required"]["field_mapping"], "host-after")
    terminal_time = iso_seconds(capture_terminal.get("ended_utc"), "capture terminal")
    require(iso_seconds(host_after["created_utc"], "host-after") >= terminal_time, "host-after precedes capture terminal")

    rows = expected_rows(plan)
    require(len(completed.get("children", [])) == 24 and len({item.get("child_id") for item in completed["children"]}) == 24, "completed child count")
    children = []
    for index, row in enumerate(rows):
        children.append(check_child(packet, row, plan, bindings["plan"], scope_values, index))
    completed_by_id = {item["child_id"]: item for item in completed["children"]}
    for child in children:
        child_id = child["row"]["child_id"]
        require(completed_by_id[child_id].get("validation") == descriptor(Path(child["child_path"]) / "validation.json", label=f"{child_id} validation"), f"{child_id} completed validation binding")
    for first, second in zip(children, children[1:]):
        require(second["start_seconds"] >= first["terminal_seconds"], "children are not serial terminal-first")

    records_by_id = {item["row"]["child_id"]: item for item in children}
    pair_results: dict[str, dict[str, Any]] = {}
    case_results: dict[str, dict[str, Any]] = {}
    for case_id in CASES:
        case_pairs: dict[str, dict[str, Any]] = {}
        for pair_type in PAIR_SCHEDULE[case_id]:
            pair = next(item for item in plan["schedule"]["pairs"] if item["case_id"] == case_id and item["pair_type"] == pair_type)
            first, second = records_by_id[pair["children"][0]], records_by_id[pair["children"][1]]
            result = pair_summary(first, second, pair)
            pair_results[pair["pair_id"]] = result
            case_pairs[pair_type] = result
        case_results[case_id] = classify(case_id, case_pairs)

    reduced = read_json(reduced_path)
    mismatches = compare_reducer({"children": children, "pairs": pair_results, "case_results": case_results}, reduced, bindings, packet)
    require(not mismatches, f"independent/reducer mismatch count {len(mismatches)}")
    reduced_binding = descriptor(reduced_path, label="reduction")
    return {
        "schema": "litchi.performance.0965.independent-raw-audit.v1",
        "status": "passed",
        "created_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="microseconds"),
        "reviewer": "/root/staging_review",
        "workload_started_by_auditor": False,
        "build_started_by_auditor": False,
        "bindings": bindings | {
            "packet": {
                "capture_start": descriptor(capture_start_path, label="capture start"),
                "capture_terminal": descriptor(capture_terminal_path, label="capture terminal"),
                "completed_packet": descriptor(completed_path, label="completed packet"),
            },
            "host_after": host_after["descriptor"],
            "reduction": reduced_binding,
        },
        "scope_contract": {
            "fields": sorted(scope_values),
            "exact_values_checked": 24,
            "host_field_mapping": plan["capture"]["host_snapshot_required"]["field_mapping"],
        },
        "custody": {
            "binary_hashes": EXPECTED_BINARY_HASHES,
            "source_descriptor_count": len(bindings["source"]),
            "historical_sample_vectors_read": False,
            "allocator_reports_read": False,
        },
        "children": [{
            "child_id": child["row"]["child_id"],
            "case_id": child["row"]["case_id"],
            "pair_id": child["row"]["pair_id"],
            "pair_type": child["row"]["pair_type"],
            "position": child["row"]["position"],
            "role": child["row"]["role"],
            "binary_role": child["row"]["binary_role"],
            "start_utc": child["start"].get("started_utc"),
            "stats": child["stats"],
            "report": child["report_descriptor"],
            "time_metrics": child["time_metrics"],
        } for child in sorted(children, key=lambda item: item["start_seconds"])],
        "pairs": pair_results,
        "case_results": case_results,
        "comparison": {
            "reducer": reduced_binding,
            "mismatch_count": 0,
            "quantiles_recomputed": 24 * len(PHASES),
            "pair_count_recomputed": 12,
            "classification_rule": "design-v3; RC/CR candidate-over-baseline p99; RR/CC reciprocal p99 controls; threshold=1.05",
        },
        "decision": {
            "adoption_allowed": False,
            "screen_reopened": False,
            "causal_claim": False,
            "0964_rejection_preserved": True,
            "labels_are_packet_scoped_diagnostics": True,
        },
        "interpretation_limits": [
            "Sample-index alignment is descriptive within separate fresh processes, not paired experimental units.",
            "Labels do not identify a source, build, host, kernel, or allocator cause.",
            "The frozen 0964 rejection and 1.05 safety rule remain binding.",
            "Any later custody or scope failure invalidates the packet; no replacements or adaptive samples are allowed.",
        ],
    }


def write_once(path: Path, value: Any) -> None:
    path = Path(path).expanduser()
    require(not path.exists() and not path.is_symlink(), f"output already exists: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write((json.dumps(value, indent=2, sort_keys=True) + "\n").encode())
        stream.flush()
        os.fsync(stream.fileno())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, default=PLAN_DEFAULT)
    parser.add_argument("--packet", type=Path, default=PACKET_DEFAULT)
    parser.add_argument("--reduced", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = audit(args.plan.resolve(), args.packet.resolve(), args.reduced.resolve())
    except AuditError as exc:
        failure = {
            "schema": "litchi.performance.0965.independent-raw-audit.v1",
            "status": "failed",
            "created_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="microseconds"),
            "reviewer": "/root/staging_review",
            "error": str(exc),
            "workload_started_by_auditor": False,
            "build_started_by_auditor": False,
        }
        try:
            write_once(args.output, failure)
        except AuditError:
            pass
        print(json.dumps(failure, sort_keys=True))
        return 1
    write_once(args.output, result)
    print(json.dumps({"status": result["status"], "output": str(args.output.resolve()), "mismatch_count": 0}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
