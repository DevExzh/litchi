#!/usr/bin/env python3
"""Independent offline audit for a completed 0945 XLSX capture packet.

This file intentionally does not import the capture or reduction helpers.  It
recomputes schedule custody, report identities, timing and allocation ratios
from retained packet files.  It never starts a workload, opens a benchmark
binary for execution, or writes inside the packet.  A moved packet can be
checked with ``--moved-packet``; all packet-owned descriptors are resolved
relative to the packet supplied for that audit.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import random
import statistics
import sys
from typing import Any

SCHEMA = "litchi.paired-audit-capture-0945.v1"
PLAN_SCHEMA = SCHEMA
ROLES = ("baseline", "candidate")
CASES = (
    ("primary-percent-medium", "primary", "xlsx_source_backed_cell_values_one_percent_edit_save", "medium"),
    ("primary-percent-dense-sparse", "primary", "xlsx_source_backed_cell_values_one_percent_edit_save", "dense-sparse"),
    ("control-xlsx-one-edit-medium", "control", "xlsx_source_backed_cell_values_one_edit_save", "medium"),
    ("control-xlsx-managed-one-edit-medium", "managed-control", "xlsx_source_backed_managed_cell_values_one_edit_save", "medium"),
)
CASE_IDS = tuple(row[0] for row in CASES)
CASE_MAP = {row[0]: row for row in CASES}
PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns", "reopen_ns")
ALLOCATOR_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
SAFETY_FIELDS = ("allocation_calls", "region_peak_live_bytes", "live_bytes_after", "max_rss_kib")
HEX = set("0123456789abcdef")
MAX_JSON_BYTES = 64 * 1024 * 1024
DROP_IDENTITY_KEYS = set(PHASES) | {
    "elapsed_ns", "operation_metrics", "plan_allocation_metrics",
    "staging_allocation_metrics", "commit_allocation_metrics",
    "commit_core_allocation_metrics", "publication_allocation_metrics",
}


class AuditError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"),
                      ensure_ascii=False, allow_nan=False).encode("utf-8")


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise AuditError(f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result

    require(path.is_file() and not path.is_symlink(), f"missing regular JSON file: {path}")
    require(path.stat().st_size <= MAX_JSON_BYTES,
            f"JSON file exceeds {MAX_JSON_BYTES} bytes: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"),
                          object_pairs_hook=reject_duplicates)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AuditError(f"cannot parse {path}: {error}") from error


def packet_relative(packet: Path, path: Path, label: str) -> str:
    try:
        return path.resolve().relative_to(packet.resolve()).as_posix()
    except ValueError as error:
        raise AuditError(f"{label} escaped packet: {path}") from error


def file_descriptor(packet: Path, path: Path, label: str) -> dict[str, Any]:
    require(not path.is_symlink() and path.is_file(), f"{label} is not a regular file: {path}")
    stat = path.stat()
    return {
        "path": packet_relative(packet, path, label),
        "bytes": stat.st_size,
        "mode": stat.st_mode & 0o777,
        "executable": bool(stat.st_mode & 0o111),
        "sha256": sha256_file(path),
    }


def descriptor(packet: Path, value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict) and isinstance(value.get("path"), str),
            f"{label} is not a descriptor")
    raw = Path(value["path"])
    require(not raw.is_absolute() and raw.as_posix() == value["path"] and ".." not in raw.parts,
            f"{label} path is not packet-relative")
    path = (packet / raw).resolve()
    require(path.is_relative_to(packet.resolve()), f"{label} escaped packet")
    actual = file_descriptor(packet, path, label)
    for key in ("bytes", "mode", "executable", "sha256"):
        if key in value:
            require(value[key] == actual[key], f"{label}.{key} differs from retained file")
    return actual


def add_error(result: dict[str, Any], scope: str, message: str) -> None:
    result.setdefault("errors", []).append({"scope": scope, "message": message})


def add_failure(result: dict[str, Any], scope: str, detail: Any) -> None:
    result.setdefault("attempted_failures", []).append({"scope": scope, "detail": detail})


def midpoint(values: list[int]) -> int:
    ordered = sorted(values)
    left = ordered[(len(ordered) - 1) // 2]
    right = ordered[len(ordered) // 2]
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def nearest_rank(values: list[int], percentile: int) -> int:
    ordered = sorted(values)
    index = min((percentile * len(ordered) + 99) // 100 - 1, len(ordered) - 1)
    return ordered[index]


def ci_critical(df: int) -> float:
    values = (12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228,
              2.201, 2.179, 2.160, 2.145, 2.131, 2.120, 2.110, 2.101, 2.093, 2.086,
              2.080, 2.074, 2.069, 2.064, 2.060, 2.056, 2.052, 2.048, 2.045, 2.042)
    if df <= 0:
        return 0.0
    if df <= len(values):
        return values[df - 1]
    z = 1.959963984540054
    d = float(df)
    return (z + (z**3 + z) / (4 * d)
            + (5 * z**5 + 16 * z**3 + 3 * z) / (96 * d**2)
            + (3 * z**7 + 19 * z**5 + 17 * z**3 - 15 * z) / (384 * d**3))


def timing_summary(elapsed: Any, samples: int, label: str) -> dict[str, Any]:
    require(isinstance(elapsed, dict) and elapsed.get("unit") == "ns",
            f"{label} elapsed unit")
    values = elapsed.get("samples")
    order = elapsed.get("sample_order")
    require(isinstance(values, list) and len(values) == samples
            and all(type(item) is int and item > 0 for item in values),
            f"{label} elapsed samples")
    require(values == sorted(values), f"{label} elapsed samples are not sorted")
    require(isinstance(order, list) and sorted(order) == list(range(samples)),
            f"{label} elapsed sample_order")
    ordered = sorted(values)
    expected = {
        "min": ordered[0], "p50": midpoint(values), "p95": nearest_rank(values, 95),
        "p99": nearest_rank(values, 99), "max": ordered[-1], "mean": statistics.mean(values),
    }
    for key, expected_value in expected.items():
        actual = elapsed.get(key)
        require(actual == expected_value if isinstance(expected_value, int)
                else isinstance(actual, (int, float))
                and math.isclose(float(actual), float(expected_value), rel_tol=1e-12, abs_tol=1e-9),
                f"{label} elapsed.{key} differs from raw samples")
    deviation = statistics.stdev(values) if samples > 1 else 0.0
    require(isinstance(elapsed.get("standard_deviation"), (int, float))
            and math.isclose(float(elapsed["standard_deviation"]), deviation,
                             rel_tol=1e-12, abs_tol=1e-9),
            f"{label} elapsed standard deviation differs")
    interval = elapsed.get("confidence_interval_95")
    require(isinstance(interval, dict)
            and interval.get("method") == "two-sided Student's t interval for the mean",
            f"{label} confidence interval method differs")
    margin = ci_critical(samples - 1) * deviation / math.sqrt(samples) if samples > 1 else 0.0
    require(isinstance(interval.get("lower"), (int, float))
            and isinstance(interval.get("upper"), (int, float))
            and math.isclose(float(interval["lower"]), max(0.0, expected["mean"] - margin),
                             rel_tol=1e-12, abs_tol=1e-9)
            and math.isclose(float(interval["upper"]), expected["mean"] + margin,
                             rel_tol=1e-12, abs_tol=1e-9),
            f"{label} confidence interval differs from raw samples")
    return {**expected, "samples": values, "sample_order": order}


def ratio(candidate: float, baseline: float, label: str) -> dict[str, Any]:
    require(candidate >= 0 and baseline >= 0, f"{label} ratio input is negative")
    if baseline == 0:
        if candidate == 0:
            return {"ratio": 1.0, "status": "degenerate_equal_zero"}
        return {"ratio": None, "status": "non_comparable_zero_baseline"}
    return {"ratio": candidate / baseline, "status": "comparable"}


def bootstrap(values: list[float], plan: dict[str, Any]) -> dict[str, Any]:
    gate = plan["gate"]["bootstrap"]
    require(len(values) == 6, "native bootstrap requires six paired blocks")
    reps = gate["replicates"]
    rng = random.Random(gate["seed"])
    draws = []
    for _ in range(reps):
        draws.append(float(statistics.median(values[rng.randrange(len(values))] for _ in values)))
    draws.sort()
    lower = draws[math.ceil(gate["lower_percentile"] / 100 * reps) - 1]
    upper = draws[math.ceil(gate["upper_percentile"] / 100 * reps) - 1]
    return {
        "replicates": reps, "seed": gate["seed"], "statistic": gate["statistic"],
        "estimate": float(statistics.median(values)), "lower_95": lower, "upper_95": upper,
        "lower_percentile": gate["lower_percentile"], "upper_percentile": gate["upper_percentile"],
    }


def normalize_identity(value: Any, key: str = "") -> Any:
    if (key in DROP_IDENTITY_KEYS or key.endswith("_ns")
            or "allocation" in key):
        return None
    if isinstance(value, dict):
        output = {}
        for child_key, child in value.items():
            normalized = normalize_identity(child, child_key)
            if normalized is not None:
                output[child_key] = normalized
        return output
    if isinstance(value, list):
        values = [normalize_identity(item, key) for item in value]
        if values and all(item == values[0] for item in values):
            return values[0]
        return values
    return value


def semantic_identity(result: dict[str, Any]) -> tuple[dict[str, Any], str]:
    identity = {
        "corpus": result.get("corpus"),
        "sink": result.get("sink"),
        "source": normalize_identity(result.get("source")),
        "output_sha256": result.get("output_sha256"),
    }
    return identity, sha256_bytes(canonical(identity))


def expected_cases(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    rows = plan.get("cases")
    require(isinstance(rows, list) and len(rows) == len(CASES), "plan case count changed")
    expected = {}
    for (case_id, kind, case, shape), row in zip(CASES, rows):
        require(isinstance(row, dict) and row.get("id") == case_id
                and row.get("kind") == kind and row.get("case") == case
                and row.get("xlsx_cell_crud_shape") == shape,
                f"plan case binding changed for {case_id}")
        expected[case_id] = row
    return expected


def validate_plan(plan: dict[str, Any]) -> dict[str, Any]:
    require(plan.get("schema") == PLAN_SCHEMA and plan.get("plan_id") == "0945"
            and plan.get("protocol_revision") == 1,
            "plan identity changed")
    require(plan.get("status") in {"frozen", "capture-ready", "ready"},
            f"plan is not capture-ready: {plan.get('status')!r}")
    require(plan.get("native", {}).get("blocks") == 6
            and plan["native"].get("children") == 48
            and plan["native"].get("samples") == 100
            and plan["native"].get("warmup") == 10,
            "native schedule/count changed")
    require(plan.get("allocator", {}).get("repeats") == 2
            and plan["allocator"].get("children") == 16
            and plan["allocator"].get("samples") == 3
            and plan["allocator"].get("warmup") == 2,
            "allocator schedule/count changed")
    require(plan.get("qualification", {}).get("children") == 16
            and plan["qualification"].get("samples") == 2
            and plan["qualification"].get("warmup") == 0,
            "qualification schedule/count changed")
    require(isinstance(plan.get("gate"), dict), "plan gates are missing")
    expected_cases(plan)
    return plan


def jobs(stage: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if stage == "qualification":
        for kind in ("native", "allocator"):
            for role in ROLES:
                for case_id in CASE_IDS:
                    rows.append({"stage": stage, "index": 0, "role": role,
                                 "kind": kind, "case_id": case_id})
    elif stage == "native":
        for block in range(6):
            roles = ROLES if block % 2 == 0 else tuple(reversed(ROLES))
            case_ids = CASE_IDS if block % 2 == 0 else tuple(reversed(CASE_IDS))
            for case_id in case_ids:
                for role in roles:
                    rows.append({"stage": stage, "index": block, "role": role,
                                 "kind": "native", "case_id": case_id})
    elif stage == "allocator":
        for repeat in range(2):
            roles = ROLES if repeat % 2 == 0 else tuple(reversed(ROLES))
            case_ids = CASE_IDS if repeat % 2 == 0 else tuple(reversed(CASE_IDS))
            for case_id in case_ids:
                for role in roles:
                    rows.append({"stage": stage, "index": repeat, "role": role,
                                 "kind": "allocator", "case_id": case_id})
    else:
        raise AuditError(f"unknown stage {stage}")
    return rows


def child_name(job: dict[str, Any]) -> str:
    prefix = {"qualification": "q", "native": "b", "allocator": "r"}[job["stage"]]
    suffix = f"-{job['kind']}" if job["stage"] == "qualification" else ""
    return f"{prefix}{job['index']:02d}-{job['role']}{suffix}-{job['case_id']}"


def stage_counts(plan: dict[str, Any], stage: str) -> tuple[int, int]:
    if stage == "qualification":
        return plan["qualification"]["samples"], plan["qualification"]["warmup"]
    if stage == "native":
        return plan["native"]["samples"], plan["native"]["warmup"]
    return plan["allocator"]["samples"], plan["allocator"]["warmup"]


def role_binary(packet: Path, plan: dict[str, Any], role: str, kind: str) -> dict[str, Any]:
    roles = plan.get("roles")
    require(isinstance(roles, dict) and isinstance(roles.get(role), dict),
            f"missing plan role {role}")
    value = roles[role].get(kind)
    return descriptor(packet, value, f"plan {role}.{kind} binary")


def check_argv(argv: Any, packet: Path, plan: dict[str, Any], job: dict[str, Any],
               folder: Path, samples: int, warmup: int) -> None:
    require(isinstance(argv, list) and all(isinstance(item, str) for item in argv),
            f"{child_name(job)} argv is malformed")
    require(argv[:3] == ["/usr/bin/taskset", "-c", str(plan["cpu"])],
            f"{child_name(job)} CPU binding changed")
    for flag, expected in (("--warmup", str(warmup)), ("--samples", str(samples)),
                           ("--case", CASE_MAP[job["case_id"]][2]),
                           ("--xlsx-cell-crud-shape", CASE_MAP[job["case_id"]][3])):
        require(flag in argv and argv[argv.index(flag) + 1] == expected,
                f"{child_name(job)} {flag} binding changed")
    require("--json" in argv, f"{child_name(job)} report argument missing")
    report_arg = argv[argv.index("--json") + 1]
    require(isinstance(report_arg, str)
            and report_arg.endswith(f"/{job['stage']}/{folder.name}/report.json"),
            f"{child_name(job)} report path is not bound to its child")
    resource = plan.get("resource", {})
    if isinstance(resource, dict) and resource.get("enabled", True):
        require("-o" in argv, f"{child_name(job)} resource output argument missing")
        resource_arg = argv[argv.index("-o") + 1]
        require(isinstance(resource_arg, str)
                and resource_arg.endswith(f"/{job['stage']}/{folder.name}/resource.json"),
                f"{child_name(job)} resource path is not bound to its child")


def check_allocator_sample(item: Any, label: str) -> None:
    require(isinstance(item, dict) and item.get("status") == "measured"
            and item.get("scope") == "operation_global_system_allocator",
            f"{label} allocation status/scope")
    values = {}
    for field in ALLOCATOR_FIELDS:
        require(type(item.get(field)) is int and item[field] >= 0,
                f"{label}.{field} is not a nonnegative integer")
        values[field] = item[field]
    require(values["failed_allocation_calls"] == 0
            and values["allocation_calls"] >= values["reallocation_calls"],
            f"{label} allocation call conservation")
    require(values["live_bytes_before"] + values["allocated_bytes"]
            == values["live_bytes_after"] + values["deallocated_bytes"],
            f"{label} allocation byte conservation")
    require(values["peak_live_bytes_before"] >= values["live_bytes_before"]
            and values["peak_live_bytes_after"] >= values["live_bytes_after"]
            and values["peak_live_bytes_after"] >= values["peak_live_bytes_before"]
            and values["region_peak_live_bytes"] >= values["live_bytes_before"]
            and values["region_peak_live_bytes"] >= values["live_bytes_after"]
            and values["region_peak_live_bytes"] <= values["peak_live_bytes_after"],
            f"{label} allocation peak/live boundaries")


def validate_report(report: dict[str, Any], plan: dict[str, Any], job: dict[str, Any],
                    binary: dict[str, Any], resource: dict[str, Any] | None) -> dict[str, Any]:
    label = child_name(job)
    samples, warmup = stage_counts(plan, job["stage"])
    require(report.get("schema_version") == 1, f"{label} report schema")
    require(report.get("binary_identity", {}).get("binary_sha256") == binary["sha256"],
            f"{label} report binary identity")
    tool = report.get("tool", {})
    require(tool.get("profile") == "release", f"{label} report profile")
    allocator = job["kind"] == "allocator"
    if allocator:
        require(tool.get("instrumentation") == "system_allocator_operation_scoped"
                and tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{label} allocator instrumentation")
    else:
        require(tool.get("instrumentation") == "none", f"{label} native instrumentation")
    environment = report.get("environment", {})
    require(environment.get("cpu_affinity") == str(plan["cpu"])
            and environment.get("git_revision") is None
            and environment.get("git_worktree_dirty") is None
            and environment.get("rustc_version") is None,
            f"{label} report runtime metadata")
    configuration = report.get("configuration", {})
    case = CASE_MAP[job["case_id"]]
    require(configuration.get("cases") == [case[2]]
            and configuration.get("samples_per_case") == samples
            and configuration.get("warmup_iterations_per_case") == warmup
            and configuration.get("xlsx_cell_crud_shapes") == [case[3]],
            f"{label} report configuration")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1
            and results[0].get("case") == case[2], f"{label} result case")
    result = results[0]
    timing = timing_summary(result.get("elapsed_ns"), samples, label)
    source = result.get("source")
    require(isinstance(source, dict) and isinstance(source.get("xlsx_cell_values"), dict),
            f"{label} XLSX source summary")
    summary = source["xlsx_cell_values"]
    phase_vectors = {}
    for phase in PHASES:
        values = summary.get(phase)
        require(isinstance(values, list) and len(values) == samples
                and all(type(item) is int and item >= 0 for item in values),
                f"{label} {phase} vector")
        phase_vectors[phase] = values
    phase_sums = [sum(phase_vectors[phase][index] for phase in PHASES
                      if phase != "reopen_ns")
                  for index in range(samples)]
    require([phase_sums[index] for index in timing["sample_order"]]
            == timing["samples"], f"{label} phase vectors do not align with elapsed samples")
    require(isinstance(result.get("output_sha256"), str)
            and len(result["output_sha256"]) == 64
            and set(result["output_sha256"]) <= HEX,
            f"{label} output digest")
    require(isinstance(result.get("sink"), dict)
            and type(result["sink"].get("accepted_bytes")) is int
            and result["sink"]["accepted_bytes"] > 0,
            f"{label} sink evidence")
    allocator_metrics: list[dict[str, Any]] = []
    if allocator:
        values = summary.get("publication_allocation_metrics")
        require(isinstance(values, list) and len(values) == samples,
                f"{label} publication allocation vector")
        for index, item in enumerate(values):
            check_allocator_sample(item, f"{label} publication[{index}]")
            allocator_metrics.append(item)
    elif summary.get("publication_allocation_metrics"):
        for index, item in enumerate(summary["publication_allocation_metrics"]):
            require(isinstance(item, dict) and item.get("status") == "unavailable",
                    f"{label} native allocation marker {index}")
    identity, identity_sha = semantic_identity(result)
    publication = sorted(phase_vectors["publication_ns"])
    return {
        "workflow_p50": timing["p50"],
        "workflow_p95": timing["p95"],
        "workflow_p99": timing["p99"],
        "publication_p50": midpoint(publication),
        "publication_p95": nearest_rank(publication, 95),
        "publication_p99": nearest_rank(publication, 99),
        "publication_samples": publication,
        "identity": identity,
        "identity_sha256": identity_sha,
        "allocator_metrics": allocator_metrics,
        "max_rss_kib": resource.get("max_rss_kib") if resource else None,
    }


def check_record_descriptor(packet: Path, value: Any, label: str,
                            expected: dict[str, Any] | None = None) -> dict[str, Any]:
    actual = descriptor(packet, value, label)
    if expected is not None:
        require(actual == expected, f"{label} differs from expected descriptor")
    return actual


def artifact_list(packet: Path, folder: Path, *, terminal: bool) -> list[dict[str, Any]]:
    excluded = {"start.json", "command.json", "receipt.json"}
    if terminal:
        excluded |= {"terminal.json", "validation-failure.json"}
    return [file_descriptor(packet, path, f"artifact {path}")
            for path in sorted(folder.iterdir())
            if path.name not in excluded]


def check_record_artifacts(packet: Path, folder: Path, record: dict[str, Any],
                           label: str, *, terminal: bool) -> None:
    recorded = record.get("artifacts")
    require(isinstance(recorded, list), f"{label} artifact descriptors missing")
    expected = artifact_list(packet, folder, terminal=terminal)
    require(recorded == expected, f"{label} artifact digest list differs")


def source_binding_check(packet: Path, value: Any, label: str) -> None:
    if not isinstance(value, dict):
        return
    for key in ("baseline_manifest", "candidate_manifest"):
        if key in value:
            descriptor(packet, value[key], f"{label}.{key}")


def audit_child(packet: Path, plan: dict[str, Any], job: dict[str, Any],
                result: dict[str, Any], plan_desc: dict[str, Any]) -> dict[str, Any]:
    name = child_name(job)
    folder = packet / job["stage"] / name
    row: dict[str, Any] = {"name": name, **job, "healthy": False, "status": "missing"}
    if not folder.is_dir():
        add_error(result, name, "expected child folder is missing")
        return row
    required = {}
    for filename in ("start.json", "terminal.json", "command.json", "receipt.json"):
        path = folder / filename
        try:
            required[filename] = read_json(path)
        except Exception as error:
            add_error(result, f"{name}/{filename}", str(error))
    if len(required) != 4:
        add_failure(result, name, {"status": "unreadable_terminal_records"})
        return row
    start, terminal = required["start.json"], required["terminal.json"]
    command, receipt = required["command.json"], required["receipt.json"]
    validation_failure = None
    if (folder / "validation-failure.json").is_file():
        try:
            validation_failure = read_json(folder / "validation-failure.json")
        except Exception as error:
            add_error(result, f"{name}/validation-failure.json", str(error))
            validation_failure = {"parse_error": str(error)}
    if validation_failure is not None:
        add_failure(result, name, validation_failure)
    samples, warmup = stage_counts(plan, job["stage"])
    binary = None
    try:
        binary = role_binary(packet, plan, job["role"], job["kind"])
    except Exception as error:
        add_error(result, name, f"binary binding: {error}")
    expected_env = plan.get("environment", {}).get("fixed")
    try:
        require(isinstance(expected_env, dict), "plan runtime environment missing")
        for label, record in (("start", start), ("terminal", terminal),
                              ("command", command), ("receipt", receipt)):
            require(record.get("schema") == SCHEMA, f"{label} schema")
            require(record.get("stage") == job["stage"]
                    and record.get("index") == job["index"]
                    and record.get("role") == job["role"]
                    and record.get("kind") == job["kind"]
                    and record.get("case_id") == job["case_id"], f"{label} schedule identity")
            require(record.get("plan") == plan_desc, f"{label} plan descriptor")
            if binary is not None:
                require(record.get("binary") == binary, f"{label} binary descriptor")
            require(record.get("environment") == expected_env
                    and record.get("environment_sha256") == sha256_bytes(canonical(expected_env)),
                    f"{label} environment binding")
            require(isinstance(record.get("cwd"), str), f"{label} cwd missing")
            if label == "start":
                check_record_descriptor(packet, record.get("helper"), f"{name}/{label}.helper")
            else:
                require(record.get("helper") == start.get("helper"), f"{label} helper descriptor drift")
            check_argv(record.get("argv"), packet, plan, job, folder, samples, warmup)
        require(start.get("process_group") is True and start.get("start_new_session") is True
                and start.get("deadline_seconds") == plan["resource"]["deadline_seconds"],
                "start process policy")
        for record in (terminal, command, receipt):
            require(record.get("argv") == start.get("argv"), "argv drift across terminal records")
        source_start = start.get("source")
        require(isinstance(source_start, dict), "source_start missing")
        source_binding_check(packet, source_start, f"{name}.source_start")
        require(command.get("source") == source_start
                and receipt.get("source_start") == source_start
                and terminal.get("source") == source_start,
                "source start custody drift")
        require(command.get("source_end") == receipt.get("source_end")
                and receipt.get("source_error") is None and command.get("source_error") is None,
                "source end custody drift")
        # The start record is written before stdout/stderr and workload
        # artifacts exist and therefore intentionally has no artifact list.
        # Terminal carries the pre-parse list; command/receipt carry the
        # final list after validation-failure.json, when present.
        for record in (terminal, command, receipt):
            check_record_artifacts(packet, folder, record, f"{name}/{record.get('stage') or 'record'}",
                                   terminal=(record is terminal))
        report_path = folder / "report.json"
        report = read_json(report_path)
        resource = read_json(folder / "resource.json") if plan.get("resource", {}).get("enabled", True) else None
        if resource is not None:
            require(type(resource.get("max_rss_kib")) is int and resource["max_rss_kib"] >= 0,
                    "resource max_rss_kib missing")
        report_values = validate_report(report, plan, job, binary or {}, resource)
        require(terminal.get("returncode") == 0 and terminal.get("timed_out") is False
                and terminal.get("launch_error") is None
                and terminal.get("artifact_error") is None
                and terminal.get("process_group_clear") is True
                and terminal.get("terminal_before_parse") is True,
                "terminal policy failed")
        require(command.get("exit_code") == 0 and receipt.get("exit_code") == 0
                and receipt.get("status") == "pass" and not receipt.get("errors"),
                "child receipt failed")
        require(command.get("terminal_receipt") == file_descriptor(packet, folder / "terminal.json", "terminal")
                and receipt.get("terminal") == file_descriptor(packet, folder / "terminal.json", "terminal"),
                "terminal descriptor custody failed")
        row.update(report_values)
        row["healthy"] = True
        row["status"] = "pass"
    except Exception as error:
        row["status"] = "failed"
        add_failure(result, name, {"status": "audit_rejected", "error": str(error)})
    digests = {}
    for filename in ("stdout.txt", "stderr.txt", "report.json", "resource.json",
                     "terminal.json", "command.json", "receipt.json", "validation-failure.json"):
        path = folder / filename
        if path.is_file() and not path.is_symlink():
            digests[filename] = file_descriptor(packet, path, f"{name}/{filename}")
    row["artifacts"] = digests
    if row["status"] == "failed" and not any(item.get("scope") == name for item in result.get("attempted_failures", [])):
        add_failure(result, name, {"status": "failed", "terminal": digests.get("terminal.json")})
    return row


def audit_stage(packet: Path, plan: dict[str, Any], stage: str, result: dict[str, Any],
                plan_desc: dict[str, Any]) -> list[dict[str, Any]]:
    expected = jobs(stage)
    stage_root = packet / stage
    info: dict[str, Any] = {"expected_children": len(expected), "expected_order": [child_name(job) for job in expected]}
    result.setdefault("stages", {})[stage] = info
    capture_path = stage_root / "capture.json"
    capture = None
    if capture_path.is_file():
        try:
            capture = read_json(capture_path)
            info["capture"] = file_descriptor(packet, capture_path, f"{stage}/capture.json")
            require(capture.get("schema") == SCHEMA and capture.get("stage") == stage,
                    "capture manifest schema/stage changed")
            require(capture.get("status") == "pass" and capture.get("children") == len(expected),
                    "capture manifest is not a complete successful cohort")
            require(capture.get("plan") == plan_desc, "capture manifest plan descriptor changed")
            require(capture.get("jobs") == [
                 {"name": child_name(job), "stage": stage, "index": job["index"],
                 "role": job["role"], "kind": job["kind"], "case_id": job["case_id"]}
                 for job in expected], "capture manifest schedule/order changed")
        except Exception as error:
            add_error(result, f"{stage}/capture.json", str(error))
            add_failure(result, f"{stage}/capture.json", {
                "status": capture.get("status") if isinstance(capture, dict) else "unreadable",
                "children": capture.get("children") if isinstance(capture, dict) else None,
            })
    else:
        add_error(result, f"{stage}/capture.json", "capture manifest is missing; attempted children are still audited")
    actual_dirs = {path.name for path in stage_root.iterdir() if path.is_dir()} if stage_root.is_dir() else set()
    expected_names = {child_name(job) for job in expected}
    for extra in sorted(actual_dirs - expected_names):
        add_error(result, f"{stage}/{extra}", "unexpected child directory")
        add_failure(result, f"{stage}/{extra}", {"status": "unexpected_attempt"})
    rows = []
    for job in expected:
        rows.append(audit_child(packet, plan, job, result, plan_desc))
    if capture is not None and isinstance(capture.get("receipts"), list):
        retained = capture["receipts"]
        require(len(retained) == len(expected),
                f"{stage}/capture.json receipt count differs from schedule")
        for index, job in enumerate(expected):
            path = stage_root / child_name(job) / "receipt.json"
            if path.is_file():
                try:
                    require(retained[index] == read_json(path),
                            f"capture receipt differs for {child_name(job)}")
                except Exception as error:
                    add_error(result, f"{stage}/capture.json/{child_name(job)}", str(error))
    info["observed_children"] = sum(row["status"] != "missing" for row in rows)
    info["passed_children"] = sum(row["healthy"] for row in rows)
    info["failed_or_missing_children"] = len(rows) - info["passed_children"]
    return rows


def native_summary(rows: list[dict[str, Any]], plan: dict[str, Any], result: dict[str, Any]) -> dict[str, Any]:
    pairs = {}
    case_summaries = {}
    for case_id in CASE_IDS:
        case_pairs = []
        for block in range(6):
            found = [row for row in rows if row.get("healthy") and row["case_id"] == case_id
                     and row["index"] == block]
            if len(found) != 2 or {row["role"] for row in found} != set(ROLES):
                add_error(result, f"native/{case_id}/block-{block}", "complete baseline/candidate pair missing")
                continue
            base = next(row for row in found if row["role"] == "baseline")
            cand = next(row for row in found if row["role"] == "candidate")
            workflow = ratio(cand["workflow_p50"], base["workflow_p50"], f"{case_id}/workflow/{block}")
            publication = ratio(cand["publication_p50"], base["publication_p50"], f"{case_id}/publication/{block}")
            case_pairs.append({
                "block": block,
                "baseline_workflow_p50": base["workflow_p50"], "candidate_workflow_p50": cand["workflow_p50"],
                "workflow": workflow,
                "baseline_publication_p50": base["publication_p50"], "candidate_publication_p50": cand["publication_p50"],
                "publication": publication,
                "baseline_workflow_p95": base["workflow_p95"], "candidate_workflow_p95": cand["workflow_p95"],
                "baseline_workflow_p99": base["workflow_p99"], "candidate_workflow_p99": cand["workflow_p99"],
                "baseline_publication_p95": base["publication_p95"], "candidate_publication_p95": cand["publication_p95"],
                "baseline_publication_p99": base["publication_p99"], "candidate_publication_p99": cand["publication_p99"],
            })
        pairs[case_id] = case_pairs
        if len(case_pairs) != 6:
            case_summaries[case_id] = {"status": "incomplete", "pairs": case_pairs}
            continue
        metrics = {}
        for key in ("workflow", "publication"):
            values = [item[key]["ratio"] for item in case_pairs]
            if any(value is None for value in values):
                metrics[key] = {"status": "non_comparable", "pairs": case_pairs}
                continue
            boot = bootstrap(values, plan)
            gate = plan["gate"]["latency"]
            boot["accepted"] = (boot["estimate"] <= gate["paired_median_ratio_max"]
                                and boot["upper_95"] < gate["bootstrap_upper_95_lt"])
            metrics[key] = boot
        case_summaries[case_id] = {"status": "complete", "pairs": case_pairs, "metrics": metrics}
    complete = all(item.get("status") == "complete" for item in case_summaries.values())
    passed = complete and all(
        case_summaries[case_id]["metrics"][metric].get("accepted") is True
        for case_id in CASE_IDS for metric in ("workflow", "publication")
    )
    return {"status": "pass" if passed else ("fail" if complete else "incomplete"),
            "all_case_latency_gate_pass": passed if complete else None,
            "cases": case_summaries, "paired_case_order": list(CASE_IDS)}


def allocator_metric_median(row: dict[str, Any], field: str) -> int | None:
    if field == "max_rss_kib":
        return row.get("max_rss_kib")
    values = row.get("allocator_metrics", [])
    if not values:
        return None
    raw = [item.get(field) for item in values]
    if not all(type(item) is int and item >= 0 for item in raw):
        return None
    return sorted(raw)[len(raw) // 2]


def allocator_summary(rows: list[dict[str, Any]], plan: dict[str, Any], result: dict[str, Any]) -> dict[str, Any]:
    gate = plan["gate"]["allocation"]
    case_summaries = {}
    for case_id in CASE_IDS:
        pairs = []
        for repeat in range(2):
            found = [row for row in rows if row.get("healthy") and row["case_id"] == case_id
                     and row["index"] == repeat]
            if len(found) != 2 or {row["role"] for row in found} != set(ROLES):
                add_error(result, f"allocator/{case_id}/repeat-{repeat}", "complete baseline/candidate pair missing")
                continue
            base = next(row for row in found if row["role"] == "baseline")
            cand = next(row for row in found if row["role"] == "candidate")
            metrics = {}
            for field in (*ALLOCATOR_FIELDS, "max_rss_kib"):
                base_value = allocator_metric_median(base, field)
                cand_value = allocator_metric_median(cand, field)
                if base_value is None or cand_value is None:
                    metrics[field] = {"status": "unavailable"}
                else:
                    pair_ratio = ratio(cand_value, base_value, f"{case_id}/{field}/{repeat}")
                    metrics[field] = {"baseline_median": base_value,
                                      "candidate_median": cand_value,
                                      **pair_ratio}
            allocated = metrics.get("allocated_bytes", {})
            pairs.append({"repeat": repeat, "metrics": metrics,
                          "publication_allocated_bytes_savings":
                          (1.0 - allocated["ratio"] if allocated.get("ratio") is not None else None)})
        if len(pairs) != 2:
            case_summaries[case_id] = {"status": "incomplete", "pairs": pairs}
            continue
        primary = case_id in gate["primary_case_ids"]
        allocated_ratios = [item["metrics"].get("allocated_bytes", {}).get("ratio") for item in pairs]
        safety = {}
        for field in gate["safety_metrics"]:
            values = [item["metrics"].get(field, {}).get("ratio") for item in pairs]
            accepted = all(value is not None and value <= gate["safety_paired_median_ratio_max"]
                           for value in values)
            safety[field] = {"pair_ratios": values,
                             "paired_median": statistics.median(values) if all(value is not None for value in values) else None,
                             "accepted": accepted}
        primary_pass = (not primary or all(value is not None and value <= gate["primary_every_pair_ratio_max"]
                                           for value in allocated_ratios))
        case_summaries[case_id] = {
            "status": "complete", "pairs": pairs,
            "publication_allocated_bytes_pair_ratios": allocated_ratios,
            "publication_allocated_bytes_savings": [item["publication_allocated_bytes_savings"] for item in pairs],
            "primary_every_pair_gate_pass": primary_pass if primary else None,
            "safety": safety,
        }
    complete = all(item.get("status") == "complete" for item in case_summaries.values())
    primary_pass = complete and all(case_summaries[case]["primary_every_pair_gate_pass"] is True
                                    for case in gate["primary_case_ids"])
    safety_pass = complete and all(summary["safety"][field]["accepted"]
                                   for summary in case_summaries.values()
                                   for field in gate["safety_metrics"])
    return {"status": "pass" if primary_pass and safety_pass else ("fail" if complete else "incomplete"),
            "primary_every_pair_gate_pass": primary_pass if complete else None,
            "all_row_safety_gate_pass": safety_pass if complete else None,
            "cases": case_summaries}


def semantic_consistency(rows: list[dict[str, Any]], result: dict[str, Any]) -> dict[str, Any]:
    by_case: dict[str, dict[str, Any]] = {}
    checked = 0
    mismatch = False
    for row in rows:
        if not row.get("healthy"):
            continue
        checked += 1
        prior = by_case.get(row["case_id"])
        if prior is None:
            by_case[row["case_id"]] = {"identity_sha256": row["identity_sha256"],
                                        "first_child": row["name"]}
        elif prior["identity_sha256"] != row["identity_sha256"]:
            mismatch = True
            add_error(result, f"semantic/{row['case_id']}",
                      f"identity drift {prior['first_child']} vs {row['name']}")
    return {"status": "pass" if not mismatch else "fail",
            "checked_reports": checked, "case_identities": by_case}


def reducer_outputs(packet: Path, explicit: list[str]) -> list[tuple[Path, dict[str, Any]]]:
    names = list(explicit)
    if not names:
        names = ["qualification-analysis.json", "analysis.json", "allocator-analysis.json"]
    output = []
    seen = set()
    for value in names:
        path = Path(value)
        if not path.is_absolute():
            path = packet / path
        path = path.resolve()
        if path in seen or not path.is_file():
            continue
        seen.add(path)
        output.append((path, read_json(path)))
    return output


def compare_reducer(packet: Path, outputs: list[str], audit: dict[str, Any]) -> list[dict[str, Any]]:
    comparisons = []
    qualification_rows = audit.get("rows", {}).get("qualification", [])
    qualification_expected = (len(qualification_rows) == len(jobs("qualification"))
                              and all(row.get("healthy") for row in qualification_rows))
    expected = {
        "qualification": qualification_expected,
        "native": audit.get("native", {}).get("all_case_latency_gate_pass"),
        "allocator": (audit.get("allocator", {}).get("primary_every_pair_gate_pass") is True
                       and audit.get("allocator", {}).get("all_row_safety_gate_pass") is True),
        "all": (audit.get("native", {}).get("all_case_latency_gate_pass") is True
                and audit.get("allocator", {}).get("primary_every_pair_gate_pass") is True
                and audit.get("allocator", {}).get("all_row_safety_gate_pass") is True),
    }
    for path, raw in reducer_outputs(packet, outputs):
        try:
            path_label = packet_relative(packet, path, "reducer output")
        except AuditError:
            path_label = str(path)
        if not isinstance(raw, dict):
            add_error(audit, path_label, "reducer output is not an object")
            continue
        stage = raw.get("stage")
        if stage not in expected:
            add_error(audit, path_label, f"unknown reducer stage {stage!r}")
            continue
        independent = expected[stage]
        reported = raw.get("status") == "pass"
        item = {"path": path_label, "sha256": sha256_file(path), "stage": stage,
                "reported_status": raw.get("status"), "independent_gate_pass": independent,
                "status_agrees": (reported == independent) if independent is not None else None}
        reduction = raw.get("reduction")
        if stage == "native" and isinstance(reduction, dict):
            item["reported_gate"] = reduction.get("all_case_latency_gate_pass")
        if stage == "allocator" and isinstance(reduction, dict):
            item["reported_gate"] = {
                "primary": reduction.get("primary_every_pair_gate_pass"),
                "safety": reduction.get("all_row_safety_gate_pass"),
            }
        if stage == "all":
            item["reported_gate"] = {
                "native": raw.get("reduction", {}).get("all_case_latency_gate_pass"),
                "primary": raw.get("allocation_reduction", {}).get("primary_every_pair_gate_pass"),
                "safety": raw.get("allocation_reduction", {}).get("all_row_safety_gate_pass"),
            }
        if item["status_agrees"] is False:
            add_error(audit, path_label, "reducer status disagrees with independent gate result")
        comparisons.append(item)
    return comparisons


def replay_projection(audit: dict[str, Any]) -> dict[str, Any]:
    stages = {}
    for stage, info in audit.get("stages", {}).items():
        children = []
        for row in audit.get("rows", {}).get(stage, []):
            children.append({"name": row.get("name"), "status": row.get("status"),
                             "healthy": row.get("healthy"), "artifacts": row.get("artifacts"),
                             "identity_sha256": row.get("identity_sha256"),
                             "workflow_p50": row.get("workflow_p50"),
                             "publication_p50": row.get("publication_p50")})
        stages[stage] = {"expected_children": info.get("expected_children"), "children": children}
    return {"plan": audit.get("plan"), "stages": stages,
            "native": audit.get("native"), "allocator": audit.get("allocator"),
            "semantic": audit.get("semantic"), "reducer": audit.get("reducer")}


def audit_packet(packet: Path, explicit_plan: str | None, reducer_paths: list[str]) -> dict[str, Any]:
    packet = packet.resolve()
    result: dict[str, Any] = {
        "schema": "litchi.performance.0945.independent-post-capture-audit.v1",
        "packet": str(packet), "status": "fail", "errors": [], "attempted_failures": [],
        "stages": {}, "rows": {}, "reducer": [],
    }
    plan_path = Path(explicit_plan) if explicit_plan else packet / "plan.json"
    if not plan_path.is_absolute():
        plan_path = packet / plan_path
    try:
        plan = read_json(plan_path.resolve())
        validate_plan(plan)
        plan_desc = file_descriptor(packet, plan_path.resolve(), "plan")
        result["plan"] = plan_desc
        for stage in ("qualification", "native", "allocator"):
            result["rows"][stage] = audit_stage(packet, plan, stage, result, plan_desc)
        all_rows = [row for rows in result["rows"].values() for row in rows]
        result["semantic"] = semantic_consistency(all_rows, result)
        result["native"] = native_summary(result["rows"]["native"], plan, result)
        result["allocator"] = allocator_summary(result["rows"]["allocator"], plan, result)
        result["reducer"] = compare_reducer(packet, reducer_paths, result)
        result["counts"] = {stage: {"expected": len(jobs(stage)),
                                     "observed": result["stages"][stage].get("observed_children", 0),
                                     "passed": result["stages"][stage].get("passed_children", 0)}
                             for stage in ("qualification", "native", "allocator")}
        complete = (not result["errors"] and not result["attempted_failures"]
                    and result["native"].get("all_case_latency_gate_pass") is True
                    and result["allocator"].get("primary_every_pair_gate_pass") is True
                    and result["allocator"].get("all_row_safety_gate_pass") is True)
        result["status"] = "pass" if complete else "fail"
    except Exception as error:
        add_error(result, "packet", str(error))
    result["replay_projection_sha256"] = sha256_bytes(canonical(replay_projection(result)))
    return result


def remap_packet_argument(value: str, source_packet: Path, target_packet: Path) -> str:
    """Map an absolute argument inside the original packet to a moved packet."""
    path = Path(value)
    if not path.is_absolute():
        return value
    try:
        return path.resolve().relative_to(source_packet.resolve()).as_posix()
    except ValueError:
        # External reducer outputs remain valid external inputs for both
        # audits; packet-owned absolute paths must be remapped above.
        return value


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", required=True, type=Path)
    parser.add_argument("--moved-packet", type=Path)
    parser.add_argument("--plan", type=str,
                        help="packet-relative plan path; defaults to plan.json")
    parser.add_argument("--reducer-output", action="append", default=[],
                        help="packet-relative reducer output; may be repeated")
    parser.add_argument("--output", required=True, type=Path,
                        help="external audit JSON path; must not already exist")
    args = parser.parse_args()
    primary = audit_packet(args.packet, args.plan, args.reducer_output)
    output: dict[str, Any] = {"schema": "litchi.performance.0945.independent-post-capture-audit.v1",
                              "status": primary["status"], "original": primary,
                              "workloads_started": False, "measurements_started": False}
    if args.moved_packet is not None:
        moved_plan = (remap_packet_argument(args.plan, args.packet, args.moved_packet)
                      if args.plan is not None else None)
        moved_reducers = [remap_packet_argument(value, args.packet, args.moved_packet)
                          for value in args.reducer_output]
        moved = audit_packet(args.moved_packet, moved_plan, moved_reducers)
        output["moved"] = moved
        left = primary.get("replay_projection_sha256")
        right = moved.get("replay_projection_sha256")
        output["moved_replay"] = {"requested": True, "equal": left == right,
                                   "original_projection_sha256": left,
                                   "moved_projection_sha256": right}
        if left != right:
            output["status"] = "fail"
            output.setdefault("errors", []).append("moved packet replay projection differs")
    else:
        output["moved_replay"] = {"requested": False}
    output_path = args.output.resolve()
    require(not output_path.exists(), f"refusing to overwrite audit output {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    encoded = json.dumps(output, indent=2, sort_keys=True, ensure_ascii=False, allow_nan=False) + "\n"
    with output_path.open("x", encoding="utf-8") as handle:
        handle.write(encoded)
        handle.flush()
    print(json.dumps({"status": output["status"], "output": str(output_path),
                      "bytes": len(encoded.encode("utf-8"))}, sort_keys=True))
    return 0 if output["status"] == "pass" else 2


if __name__ == "__main__":
    raise SystemExit(main())
