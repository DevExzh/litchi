"""Offline custody, correctness, and paired reduction for the 0948 XLSX packet.

This module never launches a workload.  It replays retained JSON, stdout,
stderr, resource, source, binary, and receipt metadata.  The qualification
action is intentionally independent from native reduction: it must pass before
the native population is admitted.  The standard XLSX primary uses the
individual validators retained in change-0705; the one-edit controls use the
case-specific checks below so their one-update corpus is never presented as a
one-percent corpus.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import random
import statistics
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent.parent
REPO = Path("/home/zhuhe/code/litchi")
PLAN_RELATIVE = "plan-v2.json"
PLAN_PATH = PACKET / PLAN_RELATIVE
RUNNER_RELATIVE = "tools/capture_0948_v1.py"
RUNNER = PACKET / RUNNER_RELATIVE
SCHEMA = "litchi.paired-audit-capture-0948.v1"
ROLES = ("baseline", "candidate")
BUILD_REVISIONS = {"baseline": "v1", "candidate": "v3"}
BUILD_BINDING_IDS = {"baseline": "baseline1", "candidate": "candidate2"}
PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")
ALL_PHASES = PHASES + ("reopen_ns",)
ALLOCATOR_FIELDS = {
    "plan_allocation_metrics", "staging_allocation_metrics",
    "commit_allocation_metrics", "commit_core_allocation_metrics",
    "publication_allocation_metrics",
}
ALLOCATOR_SCOPE = "operation_global_system_allocator"
ALLOCATOR_INSTRUMENTATION = "system_allocator_operation_scoped"
ALLOCATOR_COUNTER_REVISION = "serialized_region_peak_v3"
ALLOCATOR_SAMPLE_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
HEX = set("0123456789abcdef")
CASES = (
    ("primary-percent-medium", "primary", "xlsx_source_backed_cell_values_one_percent_edit_save", "xlsx_cell_crud_shape", "medium"),
    ("primary-percent-dense-sparse", "primary", "xlsx_source_backed_cell_values_one_percent_edit_save", "xlsx_cell_crud_shape", "dense-sparse"),
    ("control-xlsx-one-edit-medium", "control", "xlsx_source_backed_cell_values_one_edit_save", "xlsx_cell_crud_shape", "medium"),
    ("control-xlsx-managed-one-edit-medium", "managed-control", "xlsx_source_backed_managed_cell_values_one_edit_save", "xlsx_cell_crud_shape", "medium"),
)
CASE_BY_ID = {row[0]: row for row in CASES}
CASE_ORDER = tuple(row[0] for row in CASES)
MAX_JSON_BYTES = 64 * 1024 * 1024


class ReviewError(RuntimeError):
    """A retained packet failed an independent review check."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReviewError(message)


def read_json(path: Path) -> Any:
    def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ReviewError(f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result
    try:
        require(path.is_file() and not path.is_symlink(),
                f"JSON artifact is not a regular file: {path}")
        require(path.stat().st_size <= MAX_JSON_BYTES,
                f"JSON artifact exceeds {MAX_JSON_BYTES} bytes: {path}")
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_duplicates)
    except (OSError, json.JSONDecodeError, ReviewError) as error:
        raise ReviewError(f"cannot read JSON {path}: {error}") from error

def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode()


def digest(path: Path) -> dict[str, Any]:
    path = Path(path)
    require(not path.is_symlink(), f"retained artifact is a symlink: {path}")
    require(path.is_file(), f"retained artifact is missing: {path}")
    path = path.resolve()
    try:
        stored_path = path.relative_to(PACKET.resolve()).as_posix()
    except ValueError as error:
        raise ReviewError(f"retained artifact escaped packet: {path}") from error
    data = path.read_bytes()
    mode = path.stat().st_mode & 0o777
    return {"path": stored_path, "bytes": len(data), "mode": mode,
            "executable": bool(mode & 0o111),
            "sha256": hashlib.sha256(data).hexdigest()}


def path_value(value: Any, label: str, base: Path = PACKET, *,
               allow_absolute: bool = False) -> Path:
    if isinstance(value, (str, Path)):
        raw = Path(value)
    else:
        require(isinstance(value, dict) and isinstance(value.get("path"), str),
                f"{label} is not a path descriptor")
        raw = Path(value["path"])
    if raw.is_absolute():
        require(allow_absolute, f"{label} path must be packet-relative")
        return raw
    normalized = raw.as_posix() == str(value) if isinstance(value, (str, Path)) else True
    require(normalized and ".." not in raw.parts,
            f"{label} path is not normalized")
    return (base / raw).absolute()


def descriptor(value: Any, label: str) -> dict[str, Any]:
    path = path_value(value, label)
    actual = digest(path)
    if isinstance(value, dict):
        for key in ("bytes", "sha256"):
            if key in value:
                require(actual[key] == value[key], f"{label}.{key} does not match {path}")
    return actual


def plan_descriptor() -> dict[str, Any]:
    """Return the packet-relative descriptor for this versioned plan."""
    return descriptor({"path": PLAN_RELATIVE}, "plan")


def runner_descriptor() -> dict[str, Any]:
    """Return the packet-relative descriptor for the versioned capture helper."""
    return descriptor({"path": RUNNER_RELATIVE}, "capture helper")


def load_plan() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict) and plan.get("schema") == SCHEMA,
            "plan schema is not frozen 0948")
    require(plan.get("status") in {"frozen", "capture-ready", "ready"},
            f"plan is not capture-ready: {plan.get('status')!r}")
    require(plan.get("plan_id") == "0948" and plan.get("protocol_revision") == 1
            and plan.get("seed") == 948009, "plan identity changed")
    require(plan.get("build_revisions") == BUILD_REVISIONS
            and plan.get("build_binding_ids") == BUILD_BINDING_IDS,
            "plan build revision bindings changed")
    require(plan.get("qualification") == {
        "children": 16, "kinds": ["native", "allocator"],
        "roles": ["baseline", "candidate"], "samples": 2, "warmup": 0,
    }, "qualification protocol changed")
    require(plan.get("native") == {
        "blocks": 6, "children": 48, "samples": 100, "warmup": 10,
        "schedule": "six alternating blocks; even baseline then candidate and forward case order; odd candidate then baseline and reverse case order",
    }, "native protocol changed")
    require(plan.get("allocator") == {
        "repeats": 3, "children": 24, "samples": 5, "warmup": 2,
        "schedule": "three alternating repeats; even baseline then candidate and forward case order; odd candidate then baseline and reverse case order",
    }, "allocator protocol changed")
    require(plan.get("environment", {}).get("fixed") == {
        "PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C", "TZ": "UTC",
        "RUST_BACKTRACE": "0",
    }, "fixed environment changed")
    gate = plan.get("gate")
    require(isinstance(gate, dict), "plan.gate missing")
    require(gate.get("bootstrap") == {
        "confidence_level": 0.95, "lower_percentile": 2.5, "replicates": 10000,
        "seed": 948009, "statistic": "paired_median_ratio", "upper_percentile": 97.5,
    }, "bootstrap gate changed")
    require(gate.get("latency") == {
        "case_ids": list(CASE_ORDER),
        "metrics": ["whole_workflow_p50", "whole_workflow_p95",
                     "publication_p50", "publication_p95"],
        "per_block": True,
        "per_block_ratio_max": 1.05,
        "primary_case_ids": ["primary-percent-medium", "primary-percent-dense-sparse"],
        "primary_publication_p50_ratio_max": 0.95,
        "primary_publication_bootstrap_upper_95_lt": 1.0,
    }, "latency gate changed")
    require(gate.get("allocation") == {
        "case_ids": list(CASE_ORDER),
        "metrics": ["allocated_bytes", "allocation_calls", "region_peak_live_bytes",
                     "live_bytes_after", "max_rss_kib"],
        "paired_median_ratio_max": 1.05,
        "per_pair_ratio_max": 1.05,
        "per_pair": True,
        "zero_denominator": "equal-zero-only",
        "missing_metric": "fail-closed",
        "repeats": 3,
    }, "allocation gate changed")
    require(gate.get("review") == {
        "p50_spread_gt": 1.2, "regression_ratio_gt": 1.05,
        "tail_metrics": ["p50", "p95", "p99"],
        "allocation_bootstrap": False,
    }, "review gate changed")
    return plan

def case_rows(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    rows = plan.get("cases")
    require(isinstance(rows, list) and len(rows) == len(CASES),
            "plan must contain exactly four cases")
    seen: set[str] = set()
    output: dict[str, dict[str, Any]] = {}
    for row in rows:
        require(isinstance(row, dict) and isinstance(row.get("id"), str),
                "malformed plan case")
        case_id = row["id"]
        require(case_id in CASE_BY_ID and case_id not in seen,
                f"unexpected or duplicate case id {case_id}")
        expected = CASE_BY_ID[case_id]
        shape_key = expected[3]
        require((case_id, row.get("kind"), row.get("case"), shape_key,
                 row.get(shape_key)) == expected, f"case definition changed for {case_id}")
        seen.add(case_id)
        output[case_id] = row
    require(tuple(output) == CASE_ORDER, "plan case order changed")
    return output

def environment(plan: dict[str, Any]) -> dict[str, str]:
    fixed = plan.get("environment", {}).get("fixed")
    expected = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C",
                "TZ": "UTC", "RUST_BACKTRACE": "0"}
    require(fixed == expected, "fixed child environment changed")
    return dict(expected)

def source_specs(plan: dict[str, Any]) -> tuple[Any, Any, set[str], Path, dict[str, Path]]:
    source = plan.get("source")
    require(isinstance(source, dict), "plan.source missing")
    baseline, candidate, names = source.get("baseline_manifest"), source.get("candidate_manifest"), source.get("candidate_files")
    require(baseline is not None and candidate is not None and isinstance(names, list) and names,
            "source bindings are incomplete")
    candidate_names: set[str] = set()
    for name in names:
        require(isinstance(name, str) and name and not Path(name).is_absolute()
                and Path(name).as_posix() == name and ".." not in Path(name).parts,
                f"invalid candidate source path {name!r}")
        candidate_names.add(name)
    root = path_value(source.get("root_worktree", str(REPO)), "source.root_worktree",
                      allow_absolute=True)
    roots = {}
    for role, key in (("baseline", "baseline_root"), ("candidate", "candidate_root")):
        value = source.get(key)
        require(isinstance(value, str), f"source.{key} is missing")
        roots[role] = path_value(value, f"source.{key}")
    return baseline, candidate, candidate_names, root, roots

def manifest(value: Any, label: str) -> tuple[Path, dict[str, dict[str, Any]]]:
    path = path_value(value, label)
    descriptor(value, label)
    raw = read_json(path)
    files = raw.get("files") if isinstance(raw, dict) else raw
    result: dict[str, dict[str, Any]] = {}
    if isinstance(files, list):
        for item in files:
            require(isinstance(item, dict) and isinstance(item.get("path"), str),
                    f"{label} list entry is malformed")
            name = item["path"]
            require(name not in result, f"{label} duplicate path {name}")
            result[name] = item
    elif isinstance(files, dict):
        result = files
    else:
        raise ReviewError(f"{label} files map is empty")
    require(result, f"{label} files map is empty")
    for name, item in result.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute()
                and Path(name).as_posix() == name and ".." not in Path(name).parts,
                f"{label} has unsafe path {name!r}")
        require(isinstance(item, dict) and isinstance(item.get("bytes"), int)
                and isinstance(item.get("sha256"), str) and len(item["sha256"]) == 64,
                f"{label} entry {name} is malformed")
    return path, result


def source_tree_names(root: Path, label: str) -> set[str]:
    """Return every regular source path and reject symlinks or special files."""
    require(root.is_dir() and not root.is_symlink(), f"{label} source root is not a directory")
    names: set[str] = set()
    for path in root.rglob("*"):
        require(not path.is_symlink(), f"{label} source tree contains symlink {path}")
        if path.is_dir():
            continue
        require(path.is_file(), f"{label} source tree contains non-regular file {path}")
        relative = path.relative_to(root).as_posix()
        require(relative and Path(relative).as_posix() == relative
                and ".." not in Path(relative).parts,
                f"{label} source path is not normalized: {relative}")
        require(relative not in names, f"{label} source path is duplicated: {relative}")
        names.add(relative)
    return names

def source_check(plan: dict[str, Any]) -> dict[str, Any]:
    baseline_spec, candidate_spec, candidate_names, root, roots = source_specs(plan)
    require(root.is_dir(), f"source worktree missing: {root}")
    baseline_path, baseline = manifest(baseline_spec, "baseline source manifest")
    candidate_path, candidate = manifest(candidate_spec, "candidate source manifest")
    expected_commit = plan.get("source", {}).get("source_commit")
    require(isinstance(expected_commit, str) and expected_commit,
            "source binding lacks source_commit")
    baseline_raw = read_json(baseline_path)
    candidate_raw = read_json(candidate_path)
    require(baseline_raw.get("source_commit") == expected_commit,
            "baseline source commit differs from frozen binding")
    require(candidate_raw.get("base_commit") == expected_commit,
            "candidate source base commit differs from frozen binding")
    expected_count = plan.get("source", {}).get("file_count")
    require(isinstance(expected_count, int) and len(candidate) == expected_count,
            "source census count changed")
    require(set(baseline) == set(candidate), "source file sets differ")
    changed = {name for name in baseline if
               (baseline[name].get("bytes"), baseline[name].get("sha256")) !=
               (candidate[name].get("bytes"), candidate[name].get("sha256"))}
    require(changed == candidate_names, f"source delta differs: {sorted(changed)}")
    checked = 0
    for role, expected in (("baseline", baseline), ("candidate", candidate)):
        current = roots[role]
        require(current.is_dir(), f"{role} source root missing: {current}")
        observed = source_tree_names(current, role)
        require(observed == set(expected),
                f"{role} source census differs: missing={sorted(set(expected) - observed)[:3]} "
                f"extra={sorted(observed - set(expected))[:3]}")
        for name, item in expected.items():
            actual = digest(current / name)
            require(actual["bytes"] == item["bytes"] and actual["sha256"] == item["sha256"],
                    f"{role} source differs at {name}")
            checked += 1
    return {"baseline_manifest": descriptor(baseline_spec, "baseline source manifest"),
            "candidate_manifest": descriptor(candidate_spec, "candidate source manifest"),
            "candidate_files": sorted(candidate_names), "file_count": len(candidate),
            "hashed_entries": checked, "root_worktree": ".",
            "role_roots": {role: plan["source"][f"{role}_root"] for role in ROLES},
            "source_commit": expected_commit}

def role_descriptors(plan: dict[str, Any], role: str) -> dict[str, Any]:
    roles = plan.get("roles", plan.get("binaries"))
    require(isinstance(roles, dict) and isinstance(roles.get(role), dict),
            f"missing {role} role descriptors")
    output = {}
    for kind in ("native", "allocator"):
        value = roles[role].get(kind)
        require(isinstance(value, dict), f"missing {role}.{kind} descriptor")
        require(value.get("build_revision") == BUILD_REVISIONS[role]
                and value.get("build_binding_id") == BUILD_BINDING_IDS[role],
                f"{role}.{kind} build revision binding changed")
        output[kind] = value
    return output


def role_input_bindings(raw: dict[str, Any], plan: dict[str, Any], role: str,
                        label: str) -> None:
    before, after = raw.get("inputs_before"), raw.get("inputs_after")
    if before is None and after is None:
        require(isinstance(raw.get("source"), dict),
                f"{label} lacks nested source provenance")
        return
    require(before is not None and after is not None,
            f"{label} has only one side of its input census")
    require(isinstance(before, list) and before == after and before,
            f"{label} input census is missing or changed")
    baseline_spec, candidate_spec, _, _, roots = source_specs(plan)
    _, expected = manifest(baseline_spec if role == "baseline" else candidate_spec,
                           f"{role} source manifest")
    indexed: dict[str, list[dict[str, Any]]] = {}
    for item in after:
        require(isinstance(item, dict) and isinstance(item.get("path"), str),
                f"{label} input descriptor malformed")
        indexed.setdefault(str(Path(item["path"]).resolve()), []).append(item)
    root = roots[role].resolve()
    for name, wanted in expected.items():
        matches = indexed.get(str((root / name).resolve()), [])
        require(len(matches) == 1, f"{label} input census does not bind {name}")
        require(matches[0].get("bytes") == wanted["bytes"] and
                matches[0].get("sha256") == wanted["sha256"],
                f"{label} input differs at {name}")

def _reuse_validator() -> Any:
    path = PACKET / "tools" / "reuse_0948_v2.py"
    require(path.is_file() and not path.is_symlink(),
            f"missing shared provenance validator {path}")
    spec = importlib.util.spec_from_file_location("litchi_0948_reuse_v2_reducer", path)
    require(spec is not None and spec.loader is not None,
            "cannot load shared provenance validator")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def build_check(value: Any, label: str, role: str, plan: dict[str, Any], *,
                kind: str, binary: dict[str, Any] | None = None) -> dict[str, Any]:
    actual = descriptor(value, label)
    raw = read_json(path_value(value, label))
    require(isinstance(raw, dict), f"{label} is not an object")
    validator = _reuse_validator()
    if raw.get("schema") in {"litchi.performance.0948.baseline-reuse.v1",
                              "litchi.performance.0948.baseline-reuse.v2"}:
        try:
            validator.validate_baseline_reuse_0948(raw, plan, kind, binary=binary)
        except Exception as error:
            raise ReviewError(f"{label} 0948 baseline reuse validation failed: {error}") from error
        return actual
    if isinstance(raw.get("schema"), str) and raw["schema"].startswith("litchi.performance.0948."):
        try:
            validator.validate_build_terminal_0948(raw, role, kind, plan, binary=binary)
        except Exception as error:
            raise ReviewError(f"{label} 0948 build validation failed: {error}") from error
        return actual
    if raw.get("schema") == "litchi.performance.0945.native-baseline-reuse-v2":
        try:
            validator.validate_receipt(path_value(value, label), binary=binary, plan=plan)
        except Exception as error:
            raise ReviewError(f"{label} portable reuse validation failed: {error}") from error
        return actual
    require(raw.get("schema") != "litchi.performance.0945.native-baseline-reuse-v1",
            f"{label} uses the superseded nonportable reuse receipt")
    if "exit_code" in raw:
        require(raw["exit_code"] == 0, f"{label} exit_code is not zero")
    elif "returncode" in raw:
        require(raw["returncode"] == 0, f"{label} returncode is not zero")
    else:
        raise ReviewError(f"{label} has no terminal exit status")
    if "timed_out" in raw:
        require(raw.get("timed_out") is False, f"{label} timed out")
    if "deadline_exceeded" in raw:
        require(raw.get("deadline_exceeded") is False, f"{label} exceeded its deadline")
    if "terminal_before_parse" in raw:
        require(raw["terminal_before_parse"] is True, f"{label} terminal was not durable")
    if "process_group_clear" in raw:
        require(raw["process_group_clear"] is True, f"{label} process group not clear")
    if isinstance(raw.get("cleanup"), dict) and "process_group_clear" in raw["cleanup"]:
        require(raw["cleanup"]["process_group_clear"] is True,
                f"{label} build process group not clear")
    try:
            validator.validate_build_terminal(raw, role, kind, plan, binary=binary)
    except Exception as error:
        raise ReviewError(f"{label} nested source/build validation failed: {error}") from error
    role_input_bindings(raw, plan, role, label)
    return actual

def binaries_check(plan: dict[str, Any]) -> dict[str, Any]:
    output: dict[str, Any] = {}
    for role in ROLES:
        output[role] = {}
        for kind, value in role_descriptors(plan, role).items():
            actual = descriptor(value, f"{role} {kind} binary")
            binary_path = path_value(value["path"], f"{role} {kind} binary path")
            require(bool(actual.get("executable")) or bool(binary_path.stat().st_mode & 0o111),
                    f"{role} {kind} binary is not executable")
            receipt = value.get("build_receipt")
            require(receipt is not None, f"{role} {kind} build receipt is missing")
            output[role][kind] = {"binary": actual,
                                  "build_receipt": build_check(receipt, f"{role} {kind} build receipt", role, plan,
                                                                kind=kind, binary=actual)}
    return output


def counts(plan: dict[str, Any], stage: str) -> tuple[int, int]:
    if stage == "qualification":
        require(plan.get("qualification") == {
            "children": 16, "kinds": ["native", "allocator"],
            "roles": ["baseline", "candidate"], "samples": 2, "warmup": 0,
        }, "qualification counts changed")
        return 2, 0
    if stage == "native":
        require(plan.get("native", {}).get("blocks") == 6
                and plan["native"].get("samples") == 100
                and plan["native"].get("warmup") == 10, "native counts changed")
        return 100, 10
    require(plan.get("allocator", {}).get("repeats") == 3
            and plan["allocator"].get("samples") == 5
            and plan["allocator"].get("warmup") == 2, "allocator counts changed")
    return 5, 2

def jobs(stage: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if stage == "qualification":
        for kind in ("native", "allocator"):
            for role in ROLES:
                for case_id in CASE_ORDER:
                    rows.append({"stage": stage, "index": 0, "role": role,
                                 "kind": kind, "case_id": case_id})
    elif stage == "native":
        for block in range(6):
            roles = ROLES if block % 2 == 0 else tuple(reversed(ROLES))
            order = CASE_ORDER if block % 2 == 0 else tuple(reversed(CASE_ORDER))
            for case_id in order:
                for role in roles:
                    rows.append({"stage": stage, "index": block, "role": role,
                                 "kind": "native", "case_id": case_id})
    elif stage == "allocator":
        for repeat in range(3):
            roles = ROLES if repeat % 2 == 0 else tuple(reversed(ROLES))
            order = CASE_ORDER if repeat % 2 == 0 else tuple(reversed(CASE_ORDER))
            for case_id in order:
                for role in roles:
                    rows.append({"stage": stage, "index": repeat, "role": role,
                                 "kind": "allocator", "case_id": case_id})
    else:
        raise ReviewError(f"unknown stage {stage}")
    return rows

def child_name(job: dict[str, Any]) -> str:
    prefix = {"qualification": "q", "native": "b", "allocator": "r"}[job["stage"]]
    kind = f"-{job['kind']}" if job["stage"] == "qualification" else ""
    return f"{prefix}{job['index']:02d}-{job['role']}{kind}-{job['case_id']}"

def command_for(plan: dict[str, Any], job: dict[str, Any], folder: Path) -> list[str]:
    samples, warmup = counts(plan, job["stage"])
    kind = job["kind"]
    binary = path_value(role_descriptors(plan, job["role"])[kind]["path"], "binary")
    command = ["/usr/bin/taskset", "-c", str(plan["cpu"])]
    resource = plan.get("resource", {})
    require(isinstance(resource, dict) and resource.get("enabled") is True,
            "resource configuration is not frozen")
    command += [resource["time_path"], "-f", resource["format"], "-o", str(folder / "resource.json")]
    row = CASE_BY_ID[job["case_id"]]
    command += [str(binary), "--warmup", str(warmup), "--samples", str(samples),
                "--case", row[2], "--xlsx-cell-crud-shape", row[4],
                "--json", str(folder / "report.json")]
    return command

def resource_enabled(plan: dict[str, Any]) -> bool:
    value = plan.get("resource", {})
    return isinstance(value, dict) and value.get("enabled", True)


def artifact_check(folder: Path, recorded: Any, *, exclude: set[str]) -> None:
    require(isinstance(recorded, list), f"{folder} artifact descriptors are missing")
    for item in recorded:
        require(isinstance(item, dict) and isinstance(item.get("path"), str),
                f"{folder} artifact descriptor is malformed")
        raw = Path(item["path"])
        require(not raw.is_absolute() and raw.as_posix() == item["path"]
                and ".." not in raw.parts,
                f"{folder} artifact descriptor is not packet-relative")
    actual: list[dict[str, Any]] = []
    for path in sorted(folder.iterdir()):
        if path.name in exclude:
            continue
        require(path.is_file() and not path.is_symlink(),
                f"unexpected non-file artifact {path}")
        actual.append(digest(path))
    require(recorded == actual, f"artifact descriptor drift in {folder}")

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


def midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def nearest_rank(values: list[int], percentile: int) -> int:
    index = min((percentile * len(values) + 99) // 100 - 1, len(values) - 1)
    return sorted(values)[index]


def timing_check(elapsed: Any, samples: int, label: str) -> dict[str, Any]:
    require(isinstance(elapsed, dict) and elapsed.get("unit") == "ns", f"{label} elapsed unit")
    values = elapsed.get("samples")
    order = elapsed.get("sample_order")
    require(isinstance(values, list) and len(values) == samples and
            all(isinstance(v, int) and not isinstance(v, bool) and v > 0 for v in values),
            f"{label} elapsed samples")
    require(isinstance(order, list) and len(order) == samples and sorted(order) == list(range(samples)),
            f"{label} sample_order")
    require(values == sorted(values), f"{label} elapsed samples are not sorted")
    ordered = sorted(values)
    expected: dict[str, Any] = {
        "min": ordered[0], "p50": midpoint(ordered[(samples - 1) // 2], ordered[samples // 2]),
        "p95": nearest_rank(values, 95), "p99": nearest_rank(values, 99), "max": ordered[-1],
        "mean": statistics.mean(values),
    }
    for key, value in expected.items():
        reported = elapsed.get(key)
        require(reported == value if isinstance(value, int)
                else isinstance(reported, (int, float)) and math.isclose(float(reported), float(value), rel_tol=1e-12, abs_tol=1e-9),
                f"{label} elapsed.{key} differs from recomputation")
    deviation = statistics.stdev(values) if samples > 1 else 0.0
    reported_sd = elapsed.get("standard_deviation")
    require(isinstance(reported_sd, (int, float))
            and math.isclose(float(reported_sd), deviation, rel_tol=1e-12, abs_tol=1e-9),
            f"{label} standard deviation differs")
    interval = elapsed.get("confidence_interval_95")
    require(isinstance(interval, dict)
            and interval.get("method") == "two-sided Student's t interval for the mean",
            f"{label} confidence interval method changed")
    margin = ci_critical(samples - 1) * deviation / math.sqrt(samples) if samples > 1 else 0.0
    lower, upper = interval.get("lower"), interval.get("upper")
    require(isinstance(lower, (int, float)) and isinstance(upper, (int, float))
            and math.isclose(float(lower), max(0.0, float(expected["mean"]) - margin), rel_tol=1e-12, abs_tol=1e-9)
            and math.isclose(float(upper), float(expected["mean"]) + margin, rel_tol=1e-12, abs_tol=1e-9)
            and float(lower) <= float(upper), f"{label} confidence interval differs")
    return {**expected, "standard_deviation": deviation,
            "confidence_interval_95": {"lower": float(lower), "upper": float(upper)}}


def hex64(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def vector(value: Any, count: int, label: str, *, nonnegative: bool = True) -> list[Any]:
    require(isinstance(value, list) and len(value) == count, f"{label} vector length")
    if nonnegative:
        require(all(isinstance(item, int) and not isinstance(item, bool) and item >= 0 for item in value),
                f"{label} contains an invalid integer")
    return value


def sink_check(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label} sink is missing")
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        require(isinstance(value.get(key), int) and value[key] >= 0, f"{label} sink.{key}")
    require(value["accepted_bytes"] > 0 and value["largest_write"] <= 65536,
            f"{label} sink bounds")
    buckets = value.get("write_size_buckets")
    if buckets is not None:
        names = {"bytes_0", "bytes_1_to_512", "bytes_513_to_4096",
                 "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536"}
        require(isinstance(buckets, dict) and set(buckets) == names
                and all(isinstance(item, int) and item >= 0 for item in buckets.values())
                and sum(buckets.values()) == value["write_calls"]
                and buckets["bytes_over_65536"] == 0,
                f"{label} sink buckets")


def corpus_common(corpus: Any, shape: str, package: str, generator: str, label: str) -> dict[str, Any]:
    require(isinstance(corpus, dict), f"{label} corpus is missing")
    require(corpus.get("shape") == shape and corpus.get("package_format") == package
            and corpus.get("generator") == generator and corpus.get("compression") == "deflate",
            f"{label} corpus identity")
    for key in ("entry_count", "archive_member_count", "entry_bytes",
                "uncompressed_payload_bytes", "archive_bytes", "target_payload_bytes"):
        require(isinstance(corpus.get(key), int) and corpus[key] > 0, f"{label} corpus.{key}")
    for key in ("archive_sha256", "target_payload_sha256"):
        hex64(corpus.get(key), f"{label} corpus.{key}")
    return corpus


def xlsx_corpus(corpus: Any, shape: str, label: str) -> dict[str, Any]:
    corpus = corpus_common(corpus, shape, "XLSX/OPC/ZIP",
                           "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1", label)
    expected = {"medium": (9216, 48, 93), "dense-sparse": (17792, 128, 178)}[shape]
    require(corpus.get("name") == f"xlsx-cell-values-{shape}"
            and corpus.get("payload_kind") == "deterministic-multi-sheet-scalar-grid-with-media"
            and corpus.get("target_entry") == "Sheet1!A1"
            and corpus.get("target_payload_bytes") == 1
            and corpus.get("entry_count") == expected[0], f"{label} XLSX corpus fields")
    xlsx = corpus.get("xlsx")
    require(isinstance(xlsx, dict) and xlsx.get("sheet_count") == 4
            and xlsx.get("rows_per_sheet") == expected[1]
            and xlsx.get("columns_per_sheet") == expected[1]
            and xlsx.get("one_percent_update_count") == expected[2],
            f"{label} XLSX shape fields")
    members = xlsx.get("source_members")
    require(isinstance(members, dict) and members.get("workbook") == "xl/workbook.xml"
            and members.get("styles") == "xl/styles.xml"
            and members.get("shared_strings") is None
            and members.get("worksheets") == [f"xl/worksheets/sheet{i}.xml" for i in range(1, 5)],
            f"{label} XLSX member manifest")
    return corpus


def docx_corpus(corpus: Any, label: str) -> dict[str, Any]:
    corpus = corpus_common(corpus, "medium", "DOCX/OPC/ZIP", "litchi-docx-semantic-v1", label)
    require(corpus.get("name") == "docx-semantic-medium"
            and corpus.get("payload_kind") == "deterministic-semantic-text"
            and corpus.get("entry_count") == 200 and corpus.get("archive_member_count") == 12
            and corpus.get("entry_bytes") == 50 and corpus.get("uncompressed_payload_bytes") == 10000
            and corpus.get("target_entry") == "paragraph:0"
            and corpus.get("target_payload_bytes") == 50,
            f"{label} DOCX corpus fields")
    require(corpus.get("xlsx") is None, f"{label} DOCX corpus carries XLSX metadata")
    return corpus


def allocator_sample(value: Any, label: str, *, measured: bool) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} allocator sample is not an object")
    if not measured:
        require(value.get("status") == "unavailable" and value.get("scope") == ALLOCATOR_SCOPE
                and set(value) <= {"status", "scope"},
                f"{label} unexpectedly contains allocator measurements")
        return {"status": "unavailable", "scope": ALLOCATOR_SCOPE}
    require(value.get("status") == "measured" and value.get("scope") == ALLOCATOR_SCOPE,
            f"{label} allocator measurement scope")
    parsed: dict[str, Any] = {"status": "measured", "scope": ALLOCATOR_SCOPE}
    for field in ALLOCATOR_SAMPLE_FIELDS:
        item = value.get(field)
        require(type(item) is int and item >= 0, f"{label}.{field} is not nonnegative")
        parsed[field] = item
    require(parsed["failed_allocation_calls"] == 0
            and parsed["allocation_calls"] >= parsed["reallocation_calls"],
            f"{label} allocation failure/call conservation")
    require(parsed["live_bytes_before"] + parsed["allocated_bytes"] ==
            parsed["live_bytes_after"] + parsed["deallocated_bytes"],
            f"{label} allocator byte conservation")
    require(parsed["peak_live_bytes_before"] >= parsed["live_bytes_before"]
            and parsed["peak_live_bytes_after"] >= parsed["live_bytes_after"]
            and parsed["peak_live_bytes_after"] >= parsed["peak_live_bytes_before"]
            and parsed["region_peak_live_bytes"] >= parsed["live_bytes_before"]
            and parsed["region_peak_live_bytes"] >= parsed["live_bytes_after"]
            and parsed["region_peak_live_bytes"] <= parsed["peak_live_bytes_after"],
            f"{label} allocator peak/live boundaries")
    return parsed

def operation_metrics_check(value: Any, samples: int, label: str, *, allocator: bool) -> Any:
    if value is None:
        require(not allocator, f"{label} allocator operation_metrics are missing")
        return None
    require(isinstance(value, dict), f"{label} operation_metrics is not an object")
    if allocator:
        require(value.get("sample_count") == samples
                and isinstance(value.get("sample_indices"), list)
                and sorted(value["sample_indices"]) == list(range(samples))
                and value.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
                f"{label} allocator operation_metrics alignment")
    def visit(item: Any, path: str) -> Any:
        if isinstance(item, dict):
            return {key: visit(child, f"{path}.{key}") for key, child in item.items()}
        if isinstance(item, list) and path.endswith(".values"):
            require(len(item) == samples, f"{label} operation metric vector length at {path}")
            require(all((isinstance(child, (int, float)) and not isinstance(child, bool)
                         and math.isfinite(float(child)) and child >= 0)
                        or isinstance(child, str) for child in item),
                    f"{label} operation metric vector values at {path}")
            return item
        return item
    return visit(value, label)


def budget_snapshot(value: Any, label: str, *, managed: bool, post: bool) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an object")
    fields = {
        "input_bytes_used", "input_bytes_limit", "output_bytes_used", "output_bytes_limit",
        "work_used", "work_limit", "objects_used", "objects_limit",
        "catalog_reserved_objects", "cache_reserved_objects",
    }
    require(set(value) == fields, f"{label} fields changed")
    if not managed:
        for key in ("input_bytes_limit", "output_bytes_limit", "work_limit", "objects_limit"):
            require(value[key] is None, f"{label}.{key} is non-null for unmanaged evidence")
        for key in ("input_bytes_used", "output_bytes_used", "work_used", "objects_used"):
            require(value[key] == 0, f"{label}.{key} is nonzero for unmanaged evidence")
        if post:
            require(value["catalog_reserved_objects"] is None
                    and value["cache_reserved_objects"] is None,
                    f"{label} unmanaged reservation state changed")
        else:
            require(value["catalog_reserved_objects"] == 0
                    and value["cache_reserved_objects"] == 0,
                    f"{label} unmanaged reservation state changed")
    else:
        for key in ("input_bytes_used", "output_bytes_used", "work_used", "objects_used"):
            require(type(value[key]) is int and value[key] >= 0, f"{label}.{key} invalid")
        for key in ("input_bytes_limit", "output_bytes_limit", "work_limit", "objects_limit"):
            require(type(value[key]) is int and value[key] > 0, f"{label}.{key} limit invalid")
        if post:
            require(value["catalog_reserved_objects"] is None
                    and value["cache_reserved_objects"] is None,
                    f"{label} managed reservations were not released")
        else:
            require(type(value["catalog_reserved_objects"]) is int
                    and value["catalog_reserved_objects"] >= 0
                    and type(value["cache_reserved_objects"]) is int
                    and value["cache_reserved_objects"] >= 0,
                    f"{label} managed reservations invalid")
    return value


def source_identity(source: Any, samples: int, label: str, *, one_edit: bool,
                    allocator: bool, worksheet_count: int, managed: bool = False) -> dict[str, Any]:
    require(isinstance(source, dict), f"{label} source summary is missing")
    summary = source.get("xlsx_cell_values")
    require(isinstance(summary, dict), f"{label} XLSX source summary is missing")
    expected = ("managed-source-backed", "managed-budget", True) if managed else \
               ("source-backed", "unmanaged-control", False)
    require((summary.get("implementation"), summary.get("cache_mode"),
             summary.get("cache_budget_managed")) == expected,
            f"{label} source ownership mode changed")
    timing_scope = summary.get("timing_scope")
    require(isinstance(timing_scope, str)
            and all(word in timing_scope for word in ("open", "commit", "publication", "reopen")),
            f"{label} timing scope")
    require(type(worksheet_count) is int and worksheet_count > 0,
            f"{label} worksheet count is invalid")
    if one_edit:
        require(summary.get("update_count") == 1
                and summary.get("selected_worksheet_count") == 1,
                f"{label} one-edit selection counts")
    else:
        require(type(summary.get("update_count")) is int and summary["update_count"] > 1
                and summary.get("selected_worksheet_count") == worksheet_count,
                f"{label} percent selection counts")
    if managed:
        for key in ("payload_memory_limit", "publication_planning_memory_headroom",
                    "cache_budget_memory_limit"):
            require(type(summary.get(key)) is int and summary[key] > 0,
                    f"{label} managed {key} is missing")
        require(summary["publication_planning_memory_headroom"] == 64 * 1024
                and summary["cache_budget_memory_limit"] ==
                summary["payload_memory_limit"] + summary["publication_planning_memory_headroom"],
                f"{label} managed memory-limit accounting changed")
    else:
        for key in ("payload_memory_limit", "publication_planning_memory_headroom",
                    "cache_budget_memory_limit"):
            require(summary.get(key) is None, f"{label} unmanaged {key} is non-null")

    phases: dict[str, list[int]] = {}
    allocator_metrics: dict[str, list[dict[str, Any]]] = {}
    constants: dict[str, Any] = {}
    for key, value in summary.items():
        if key in ALLOCATOR_FIELDS:
            values = vector(value, samples, f"{label} source.xlsx_cell_values.{key}", nonnegative=False)
            allocator_metrics[key] = [allocator_sample(item, f"{label}.{key}[{index}]",
                                                        measured=allocator)
                                      for index, item in enumerate(values)]
            continue
        if isinstance(value, list):
            vector(value, samples, f"{label} source.xlsx_cell_values.{key}", nonnegative=False)
            if key in ALL_PHASES:
                require(all(type(item) is int and item >= 0 for item in value),
                        f"{label} phase vector")
                phases[key] = value
            else:
                require(all(item == value[0] for item in value),
                        f"{label} source identity varies: {key}")
                constants[key] = value[0]
        else:
            constants[key] = value
    require(set(phases) == set(ALL_PHASES), f"{label} incomplete phase vectors")
    for key in ("output_sha256", "semantic_sha256", "untouched_member_sha256"):
        hex64(constants.get(key), f"{label} {key}")
    require(constants.get("payload_materializations") == constants.get("cache_successful_loads")
            and constants.get("cache_cold_loads") == constants.get("cache_successful_loads")
            and constants.get("cache_retained_entries") == constants.get("cache_successful_loads")
            and constants.get("cache_retained_bytes", 0) > 0,
            f"{label} source cache counters")
    for key in ("cache_hits", "cache_waiter_joins", "cache_failed_loads", "cache_evictions",
                "cache_bypasses", "cache_oversized_bypasses", "cache_allocation_bypasses",
                "cache_in_flight_loads", "cache_budget_reservation_failures"):
        require(constants.get(key) == 0, f"{label} source invariant {key}")
    if managed:
        require(constants.get("cache_budget_memory_used", 0) >= 0
                and constants.get("cache_budget_reserved_bytes", 0) >= 0,
                f"{label} managed cache budget counters")
        for key in ("budget_used_after_package_drop", "budget_used_after_handles_drop",
                    "budget_objects_used_after_handles_drop"):
            vector(summary.get(key), samples, f"{label}.{key}")
    else:
        for key in ("cache_budget_memory_used", "cache_budget_reserved_bytes",
                    "budget_used_after_package_drop", "budget_used_after_handles_drop",
                    "budget_objects_used_after_handles_drop"):
            values = constants.get(key)
            if isinstance(values, list):
                require(all(value == 0 for value in values), f"{label} unmanaged {key} is nonzero")
            else:
                require(values == 0, f"{label} unmanaged {key} is nonzero")
    if one_edit:
        expected_unselected = worksheet_count - summary["selected_worksheet_count"]
        require(expected_unselected > 0
                and constants.get("unselected_worksheet_read_calls") == expected_unselected
                and constants.get("unselected_worksheet_read_bytes", 0) > 0,
                f"{label} untouched worksheet reads")
    else:
        require(constants.get("unselected_worksheet_read_calls") == 0
                and constants.get("unselected_worksheet_read_bytes") == 0,
                f"{label} all-selected worksheet reads")
    require(constants.get("source_read_calls", 0) > 0
            and constants.get("source_read_bytes", 0) > 0
            and constants.get("workbook_read_calls") == 1
            and constants.get("selected_worksheet_read_calls", 0) >= summary["selected_worksheet_count"]
            and constants.get("selected_worksheet_read_bytes", 0) > 0
            and constants.get("workbook_read_bytes", 0) > 0,
            f"{label} source read counters")
    for generic, specific in (("read_calls", "source_read_calls"),
                              ("read_bytes", "source_read_bytes"),
                              ("ordinary_payload_materializations", "payload_materializations")):
        generic_values = vector(source.get(generic), samples, f"{label} source.{generic}",
                                nonnegative=False)
        specific_values = vector(summary.get(specific), samples,
                                 f"{label} source.xlsx_cell_values.{specific}",
                                 nonnegative=False)
        require(generic_values == specific_values, f"{label} generic {generic}")
    for key, value in source.items():
        if isinstance(value, list):
            vector(value, samples, f"{label} source.{key}", nonnegative=False)
            require(all(item == value[0] for item in value),
                    f"{label} generic source identity varies: {key}")
    refusal = constants.get("output_budget_refusal")
    fields = {"successful_output_ceiling", "first_output_request_bytes", "one_under_output_limit",
              "accepted_output_bytes", "source_read_calls", "source_read_bytes", "output_bytes_used",
              "typed_output_resource_refusal", "zero_output_verified", "source_identity_preserved"}
    require(isinstance(refusal, dict) and set(refusal) == fields,
            f"{label} output-budget refusal evidence changed")
    if managed:
        require(refusal["successful_output_ceiling"] > 0
                and refusal["first_output_request_bytes"] == refusal["one_under_output_limit"] + 1
                and refusal["first_output_request_bytes"] > 0
                and refusal["accepted_output_bytes"] == 0
                and refusal["output_bytes_used"] == 0
                and refusal["source_read_calls"] > 0
                and refusal["source_read_bytes"] > 0
                and refusal["typed_output_resource_refusal"] is True
                and refusal["zero_output_verified"] is True
                and refusal["source_identity_preserved"] is True,
                f"{label} managed output refusal gate")
        pre = [budget_snapshot(item, f"{label} pre_publication_budget[{i}]", managed=True, post=False)
               for i, item in enumerate(summary.get("pre_publication_budget", []))]
        post = [budget_snapshot(item, f"{label} post_publication_budget[{i}]", managed=True, post=True)
                for i, item in enumerate(summary.get("post_publication_budget", []))]
        require(len(pre) == samples and len(post) == samples, f"{label} managed budget vector length")
        require(all(item["input_bytes_used"] > 0 and item["work_used"] > 0
                    and item["objects_used"] > 0 and item["output_bytes_used"] == 0 for item in pre),
                f"{label} managed pre-budget usage")
        require(all(item["output_bytes_used"] > 0 for item in post),
                f"{label} managed post-budget output usage")
        require(all(value > 0 for value in summary["budget_used_after_package_drop"])
                and all(value == 0 for value in summary["budget_used_after_handles_drop"])
                and all(value == 0 for value in summary["budget_objects_used_after_handles_drop"]),
                f"{label} managed reservation release")
    else:
        budget_snapshot(constants.get("pre_publication_budget"), f"{label} pre_publication_budget", managed=False, post=False)
        budget_snapshot(constants.get("post_publication_budget"), f"{label} post_publication_budget", managed=False, post=True)
        require(all(refusal[field] == 0 for field in fields - {
            "typed_output_resource_refusal", "zero_output_verified", "source_identity_preserved"})
                and all(refusal[field] is False for field in
                        ("typed_output_resource_refusal", "zero_output_verified", "source_identity_preserved")),
                f"{label} unmanaged output refusal evidence")
    return {"summary": summary, "constants": constants, "phases": phases,
            "allocator_metrics": allocator_metrics,
            "managed": managed,
            "generic": {key: source[key] for key in source if key != "xlsx_cell_values"}}

def metric_summary(values: list[int], label: str) -> dict[str, int]:
    require(values and all(type(value) is int and value > 0 for value in values),
            f"{label} metric vector is invalid")
    ordered = sorted(values)
    return {
        "p50": midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2]),
        "p95": nearest_rank(ordered, 95),
        "p99": nearest_rank(ordered, 99),
        "min": ordered[0],
        "max": ordered[-1],
    }


def phase_alignment(result: dict[str, Any], phases: dict[str, list[int]], label: str) -> None:
    if not phases:
        return
    phase_sums = [sum(phases[phase][index] for phase in PHASES)
                  for index in range(len(phases[PHASES[0]]))]
    order = result["elapsed_ns"]["sample_order"]
    require([phase_sums[index] for index in order] == result["elapsed_ns"]["samples"]
            and sorted(range(len(phase_sums)), key=lambda index: (phase_sums[index], index)) == order,
            f"{label} phase vectors do not align with elapsed samples")


def primary_validate(result: dict[str, Any], shape: str, samples: int, validator: Any,
                     label: str, *, allocator: bool) -> dict[str, Any]:
    corpus = validator.validate_corpus(result.get("corpus", {}), shape)
    sink_check(result.get("sink"), label)
    validator.validate_sink(result.get("sink"), label)
    if samples > 30:
        validator.check_reported_statistics(result.get("elapsed_ns", {}), samples)
    source = result.get("source")
    if allocator:
        source_details = source_identity(source, samples, label, one_edit=False,
                                         allocator=True, worksheet_count=corpus["xlsx"]["sheet_count"],
                                         managed=False)
        phases, constants = source_details["phases"], source_details["constants"]
        require(constants["output_sha256"] == result.get("output_sha256"),
                f"{label} output digest")
        phase_alignment(result, phases, label)
    else:
        phases, constants = validator.validate_source(result, corpus, shape, samples, compatibility=False)
        require(result.get("output_sha256") == constants["output_sha256"], f"{label} output digest")
        source_details = source_identity(source, samples, label, one_edit=False,
                                         allocator=False, worksheet_count=corpus["xlsx"]["sheet_count"],
                                         managed=False)
    return {"corpus": corpus, "sink": result["sink"], "source": source,
            "phase_vectors": phases, "source_constants": constants,
            "allocator_metrics": source_details["allocator_metrics"], "managed": False}

def xlsx_control_validate(result: dict[str, Any], shape: str, samples: int, label: str,
                          *, allocator: bool, managed: bool) -> dict[str, Any]:
    corpus = xlsx_corpus(result.get("corpus"), shape, label)
    sink_check(result.get("sink"), label)
    source_details = source_identity(result.get("source"), samples, label, one_edit=True,
                                     allocator=allocator, worksheet_count=corpus["xlsx"]["sheet_count"],
                                     managed=managed)
    require(isinstance(result.get("output_sha256"), str), f"{label} output digest missing")
    hex64(result["output_sha256"], f"{label} output digest")
    require(source_details["constants"]["output_sha256"] == result["output_sha256"],
            f"{label} source/output digest mismatch")
    phase_alignment(result, source_details["phases"], label)
    if managed:
        post = source_details["summary"]["post_publication_budget"]
        require(all(item["output_bytes_used"] == result["sink"]["accepted_bytes"] for item in post),
                f"{label} managed post-budget output differs from sink")
    return {"corpus": corpus, "sink": result["sink"], "source": result["source"],
            "phase_vectors": source_details["phases"],
            "source_constants": source_details["constants"],
            "allocator_metrics": source_details["allocator_metrics"], "managed": managed}

def docx_control_validate(result: dict[str, Any], samples: int, label: str,
                          *, allocator: bool) -> dict[str, Any]:
    # DOCX is deliberately outside 0948.  Keep the historical helper as an
    # explicit rejection so an accidentally reintroduced non-XLSX row cannot
    # pass with only a sink/null-source check.
    raise ReviewError(f"{label}: DOCX controls are excluded from 0948")


def semantic_identity(result: dict[str, Any], validated: dict[str, Any]) -> dict[str, Any]:
    source = validated.get("source")
    identity: dict[str, Any] = {"corpus": validated["corpus"], "sink": validated["sink"]}
    if result.get("output_sha256") is not None:
        identity["output_sha256"] = result["output_sha256"]
    if isinstance(source, dict):
        # Keep source counters, cache/resource constants, and digests.  Timing
        # vectors and allocator measurements are deliberately excluded.
        def strip(value: Any, key: str = "") -> Any:
            if key.endswith("_ns") or "allocation" in key:
                return None
            if isinstance(value, dict):
                return {k: stripped for k, item in value.items()
                        if (stripped := strip(item, k)) is not None}
            if isinstance(value, list):
                items = [strip(item, key) for item in value]
                # Source evidence vectors are required to be constant by the
                # case validators.  Compare their semantic value rather than
                # their population-specific vector length (qualification=2,
                # native=100, allocator=5).
                if items and all(item == items[0] for item in items):
                    return items[0]
                return items
            return value
        identity["source"] = strip(source)
    return identity


def report_validate(plan: dict[str, Any], job: dict[str, Any], binary: dict[str, Any],
                    validator: Any, report: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    label = f"{job['stage']}/{child_name(job)}"
    samples, warmup = counts(plan, job["stage"])
    require(report.get("schema_version") == 1, f"{label} report schema")
    require(report.get("binary_identity", {}).get("binary_sha256") == binary["sha256"],
            f"{label} report binary identity")
    tool = report.get("tool", {})
    require(tool.get("profile") == "release", f"{label} report profile")
    allocator = job["kind"] == "allocator"
    if allocator:
        require(tool.get("instrumentation") == ALLOCATOR_INSTRUMENTATION
                and tool.get("allocator_counter_revision") == ALLOCATOR_COUNTER_REVISION,
                f"{label} allocator instrumentation")
    else:
        require(tool.get("instrumentation") == "none", f"{label} native instrumentation")
    env = report.get("environment", {})
    require(env.get("cpu_affinity") == str(plan["cpu"])
            and env.get("git_revision") is None and env.get("git_worktree_dirty") is None
            and env.get("rustc_version") is None, f"{label} runtime metadata")
    config = report.get("configuration", {})
    row = CASE_BY_ID[job["case_id"]]
    require(config.get("cases") == [row[2]]
            and config.get("samples_per_case") == samples
            and config.get("warmup_iterations_per_case") == warmup,
            f"{label} report configuration")
    require(config.get("xlsx_cell_crud_shapes") == [row[4]],
            f"{label} XLSX shape configuration")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1 and results[0].get("case") == row[2],
            f"{label} result case")
    result = results[0]
    timing = timing_check(result.get("elapsed_ns"), samples, label)
    operation_metrics = operation_metrics_check(result.get("operation_metrics"), samples, label,
                                                allocator=allocator)
    if row[1] == "primary":
        validated = primary_validate(result, row[4], samples, validator, label, allocator=allocator)
    else:
        validated = xlsx_control_validate(
            result, row[4], samples, label, allocator=allocator,
            managed=(row[1] == "managed-control"))
    publication = validated["phase_vectors"].get("publication_ns")
    require(publication is not None, f"{label} publication phase vector is missing")
    publication_stats = metric_summary(publication, f"{label} publication")
    identity = semantic_identity(result, validated)
    return {"label": label, "stage": job["stage"], "index": job["index"],
            "role": job["role"], "kind": job["kind"], "case_id": job["case_id"],
            "samples": samples, "warmup": warmup, "p50": timing["p50"],
            "p95": timing["p95"], "p99": timing["p99"], "mean": timing["mean"],
            "whole_workflow": timing, "whole_workflow_p50": timing["p50"],
            "publication": publication_stats, "publication_p50": publication_stats["p50"],
            "publication_p95": publication_stats["p95"], "publication_p99": publication_stats["p99"],
            "identity": identity, "operation_metrics": operation_metrics,
            "operation_metrics_status": "available" if operation_metrics is not None else "not_applicable",
            "allocator_metrics": validated.get("allocator_metrics", {}),
            "allocator_metrics_status": "measured" if validated.get("allocator_metrics") else "not_applicable",
            "resource": None}, result

def leaf_review(plan: dict[str, Any], job: dict[str, Any], binaries: dict[str, Any],
                validator: Any) -> dict[str, Any]:
    folder = PACKET / job["stage"] / child_name(job)
    require(folder.is_dir() and not folder.is_symlink(),
            f"child folder must be a real directory {folder}")
    start = read_json(folder / "start.json")
    terminal = read_json(folder / "terminal.json")
    command = read_json(folder / "command.json")
    receipt = read_json(folder / "receipt.json")
    require(not (folder / "validation-failure.json").exists(), f"retained child failure {folder}")
    binary = binaries[job["role"]][job["kind"]]["binary"]
    expected_command = command_for(plan, job, folder)
    env = environment(plan)
    env_hash = hashlib.sha256(canonical(env)).hexdigest()
    workload = path_value(plan["workload_cwd"], "workload_cwd", allow_absolute=True)
    require(workload.is_dir() and not workload.is_symlink(),
            f"workload cwd must be a real directory {workload}")
    expected_cwd = str(workload)
    for record, name in ((start, "start"), (terminal, "terminal"),
                         (command, "command"), (receipt, "receipt")):
        require(record.get("schema") == SCHEMA, f"{folder} {name} schema")
        require(record.get("stage") == job["stage"] and record.get("index") == job["index"]
                and record.get("role") == job["role"] and record.get("kind") == job["kind"]
                and record.get("case_id") == job["case_id"], f"{folder} {name} schedule binding")
        require(record.get("argv") == expected_command and record.get("cwd") == expected_cwd
                and record.get("environment") == env
                and record.get("environment_sha256") == env_hash
                and record.get("binary") == binary,
                f"{folder} {name} invocation custody")
        require(record.get("plan") == plan_descriptor(),
                f"{folder} {name} plan custody")
        require(record.get("helper") == runner_descriptor(),
                f"{folder} {name} capture helper custody")
    require(start.get("process_group") is True and start.get("start_new_session") is True
            and start.get("deadline_seconds") == 900,
            f"{folder} start process policy")
    require(terminal.get("returncode") == 0 and terminal.get("timed_out") is False
            and terminal.get("launch_error") is None
            and terminal.get("artifact_error") is None
            and terminal.get("process_group_clear") is True
            and terminal.get("terminal_before_parse") is True,
            f"{folder} terminal policy failed")
    require(command.get("exit_code") == 0 and receipt.get("exit_code") == 0
            and receipt.get("status") == "pass" and not receipt.get("errors"),
            f"{folder} child receipt failed")
    require(command.get("terminal_receipt") == digest(folder / "terminal.json")
            and receipt.get("terminal") == digest(folder / "terminal.json"),
            f"{folder} terminal descriptor custody")
    require(all((folder / name).is_file() and not (folder / name).is_symlink()
                for name in ("stdout.txt", "stderr.txt", "report.json")),
            f"{folder} retained output missing or symlinked")
    if resource_enabled(plan):
        resource = read_json(folder / "resource.json")
        require(isinstance(resource, dict) and type(resource.get("max_rss_kib")) is int
                and resource["max_rss_kib"] >= 0, f"{folder} RSS receipt")
    artifact_check(folder, command.get("artifacts"), exclude={"start.json", "command.json", "receipt.json"})
    artifact_check(folder, receipt.get("artifacts"), exclude={"start.json", "command.json", "receipt.json"})
    terminal_artifacts = [digest(path) for path in sorted(folder.iterdir())
                          if path.name not in {"start.json", "terminal.json", "command.json", "receipt.json"}
                          and path.name != "validation-failure.json"]
    require(terminal.get("artifacts") == terminal_artifacts,
            f"{folder} terminal artifact set changed")
    require(start.get("source") == command.get("source") == receipt.get("source_start")
            and terminal.get("source") == start.get("source")
            and command.get("source_end") == receipt.get("source_end")
            and receipt.get("source_error") is None and command.get("source_error") is None,
            f"{folder} source custody receipt")
    reviewed, _result = report_validate(plan, job, binary, validator, read_json(folder / "report.json"))
    reviewed["folder"] = folder.relative_to(PACKET.resolve()).as_posix()
    reviewed["report"] = digest(folder / "report.json")
    reviewed["command"] = digest(folder / "command.json")
    reviewed["terminal"] = digest(folder / "terminal.json")
    reviewed["receipt"] = digest(folder / "receipt.json")
    reviewed["resource"] = read_json(folder / "resource.json") if resource_enabled(plan) else None
    reviewed["max_rss_kib"] = reviewed["resource"]["max_rss_kib"] if reviewed["resource"] else None
    return reviewed

def capture_review(plan: dict[str, Any], stage: str, binaries: dict[str, Any],
                   validator: Any) -> list[dict[str, Any]]:
    stage_path = PACKET / stage
    require(stage_path.is_dir() and not stage_path.is_symlink(),
            f"{stage} capture stage must be a real directory")
    capture = read_json(stage_path / "capture.json")
    expected = jobs(stage)
    samples, warmup = counts(plan, stage)
    require(capture.get("schema") == SCHEMA and capture.get("status") == "pass"
            and capture.get("stage") == stage and capture.get("children") == len(expected)
            and capture.get("samples_per_child") == samples
            and capture.get("warmup_per_child") == warmup,
            f"{stage} capture manifest counts")
    require(capture.get("plan") == plan_descriptor()
            and capture.get("helper") == runner_descriptor()
            and capture.get("environment") == environment(plan)
            and capture.get("environment_sha256") == hashlib.sha256(canonical(environment(plan))).hexdigest()
            and capture.get("binary_bindings") == binaries
            and capture.get("source_start") == capture.get("source_end")
            and capture.get("source_start") == source_check(plan),
            f"{stage} capture custody")
    expected_jobs = [{"name": child_name(job), "stage": stage, "index": job["index"],
                      "role": job["role"], "kind": job["kind"], "case_id": job["case_id"]}
                     for job in expected]
    require(capture.get("jobs") == expected_jobs, f"{stage} schedule changed")
    expected_names = {item["name"] for item in expected_jobs}
    actual_names: set[str] = set()
    for path in stage_path.iterdir():
        if path.name == "capture.json":
            require(path.is_file() and not path.is_symlink(),
                    f"{stage} capture manifest is not a regular file")
            continue
        require(path.is_dir() and not path.is_symlink(),
                f"{stage} contains a non-directory or symlink child {path}")
        actual_names.add(path.name)
    require(actual_names == expected_names, f"{stage} has extra or missing child folders")
    rows = [leaf_review(plan, job, binaries, validator) for job in expected]
    retained = capture.get("receipts")
    require(isinstance(retained, list) and len(retained) == len(expected),
            f"{stage} capture receipts are incomplete")
    for index, job in enumerate(expected):
        receipt = read_json(PACKET / stage / child_name(job) / "receipt.json")
        require(retained[index] == receipt,
                f"{stage} receipt differs for {child_name(job)}")
    return rows

def load_validator(plan: dict[str, Any]) -> Any:
    value = plan.get("semantic_validator") or plan.get("prior_semantic_validator")
    require(value is not None, "plan.semantic_validator descriptor is missing")
    actual = descriptor(value, "semantic validator")
    spec = importlib.util.spec_from_file_location(
        "litchi_0948_semantic_validator", path_value(value, "semantic validator"))
    require(spec is not None and spec.loader is not None, "cannot load semantic validator")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def compare_identities(rows: list[dict[str, Any]]) -> None:
    grouped: dict[str, dict[str, Any]] = {}
    for row in rows:
        key = row["case_id"]
        previous = grouped.setdefault(key, row["identity"])
        require(previous == row["identity"], f"semantic identity drift for {key}: {row['label']}")


def bootstrap(values: list[float], replicates: int, seed: int) -> dict[str, Any]:
    require(values, "paired bootstrap has no values")
    require(replicates == 10000 and seed == 948009, "bootstrap parameters changed")
    rng = random.Random(seed)
    draws: list[float] = []
    for _ in range(replicates):
        sample = [values[rng.randrange(len(values))] for _ in values]
        draws.append(float(statistics.median(sample)))
    draws.sort()
    lower_index = math.ceil(0.025 * replicates) - 1
    upper_index = math.ceil(0.975 * replicates) - 1
    return {"replicates": replicates, "seed": seed, "statistic": "paired_median_ratio",
            "confidence_level": 0.95, "lower_percentile": 2.5,
            "upper_percentile": 97.5, "estimate": float(statistics.median(values)),
            "lower_95": draws[lower_index], "upper_95": draws[upper_index]}

def ratio(candidate: float, baseline: float, label: str) -> float:
    require(candidate >= 0 and baseline >= 0, f"{label} ratio input is negative")
    if baseline == 0:
        # Equal measured zeros are a finite, explicitly degenerate comparison.
        # A nonzero candidate over a zero baseline has no meaningful ratio and
        # fails closed before any JSON result can be written.
        require(candidate == 0,
                f"{label} has a nonzero candidate over a zero baseline (non-comparable)")
        return 1.0
    return candidate / baseline


def ratio_status(candidate: float, baseline: float, label: str) -> str:
    """Return the frozen comparability state for a nonnegative pair."""
    require(candidate >= 0 and baseline >= 0, f"{label} ratio input is negative")
    if baseline == 0:
        require(candidate == 0,
                f"{label} has a nonzero candidate over a zero baseline (non-comparable)")
        return "degenerate_equal_zero"
    return "comparable"


def native_reduce(rows: list[dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    native = [row for row in rows if row["stage"] == "native"]
    require(len(native) == 48, "native child count is not 48")
    gate = plan["gate"]
    latency_gate = gate["latency"]
    review = gate["review"]
    pairs: dict[str, list[dict[str, Any]]] = {case_id: [] for case_id in CASE_ORDER}
    for block in range(6):
        for case_id in CASE_ORDER:
            pair = [row for row in native if row["index"] == block and row["case_id"] == case_id]
            require(len(pair) == 2 and {row["role"] for row in pair} == set(ROLES),
                    f"native pair missing for block {block}/{case_id}")
            base = next(row for row in pair if row["role"] == "baseline")
            cand = next(row for row in pair if row["role"] == "candidate")
            pair_result = {
                "block": block,
                "baseline_whole_workflow_p50": base["whole_workflow_p50"],
                "candidate_whole_workflow_p50": cand["whole_workflow_p50"],
                "whole_workflow_ratio": ratio(cand["whole_workflow_p50"],
                                                base["whole_workflow_p50"], case_id),
                "whole_workflow_comparability": ratio_status(cand["whole_workflow_p50"],
                                                               base["whole_workflow_p50"], case_id),
                "baseline_publication_p50": base["publication_p50"],
                "candidate_publication_p50": cand["publication_p50"],
                "publication_ratio": ratio(cand["publication_p50"],
                                             base["publication_p50"], case_id),
                "publication_comparability": ratio_status(cand["publication_p50"],
                                                            base["publication_p50"], case_id),
            }
            # Keep every per-pair tail ratio.  Aggregate tail medians below are
            # descriptive and cannot hide an individual >1.05 regression.
            for metric, base_key, cand_key in (
                ("whole_workflow_p95", "p95", "p95"),
                ("whole_workflow_p99", "p99", "p99"),
                ("publication_p95", "publication_p95", "publication_p95"),
                ("publication_p99", "publication_p99", "publication_p99"),
            ):
                pair_result[f"baseline_{metric}"] = base[base_key]
                pair_result[f"candidate_{metric}"] = cand[cand_key]
                pair_result[f"{metric}_ratio"] = ratio(cand[cand_key], base[base_key],
                                                         f"{case_id}/{metric}/{block}")
                pair_result[f"{metric}_comparability"] = ratio_status(
                    cand[cand_key], base[base_key], f"{case_id}/{metric}/{block}")
            pairs[case_id].append(pair_result)
    summaries: dict[str, Any] = {}
    review_flags: list[dict[str, Any]] = []
    for case_id, values in pairs.items():
        metrics = {}
        for metric, key in (("whole_workflow_p50", "whole_workflow_ratio"),
                            ("publication_p50", "publication_ratio")):
            ratios = [item[key] for item in values]
            boot = bootstrap(ratios, gate["bootstrap"]["replicates"], gate["bootstrap"]["seed"])
            boot["pairs"] = values
            if metric == "publication_p50" and case_id in latency_gate["primary_case_ids"]:
                estimate_limit = latency_gate["primary_publication_p50_ratio_max"]
                upper_limit = latency_gate["primary_publication_bootstrap_upper_95_lt"]
                boot["accepted"] = (boot["estimate"] <= estimate_limit
                                     and boot["upper_95"] < upper_limit)
            else:
                estimate_limit = None
                upper_limit = None
                boot["accepted"] = True
            boot["estimate_limit"] = estimate_limit
            boot["upper_95_limit"] = upper_limit
            metrics[metric] = boot
            for item in values:
                if item[key] > review["regression_ratio_gt"]:
                    review_flags.append({"case_id": case_id, "metric": metric,
                                         "kind": "individual_pair", "block": item["block"],
                                         "ratio": item[key], "threshold": review["regression_ratio_gt"]})
        for metric, base_key, cand_key in (("whole_workflow_p95", "p95", "p95"),
                                            ("whole_workflow_p99", "p99", "p99"),
                                            ("publication_p95", "publication_p95", "publication_p95"),
                                            ("publication_p99", "publication_p99", "publication_p99")):
            pair_key = f"{metric}_ratio"
            for item in values:
                if item[pair_key] > review["regression_ratio_gt"]:
                    review_flags.append({"case_id": case_id, "metric": metric,
                                         "kind": "individual_pair", "block": item["block"],
                                         "ratio": item[pair_key],
                                         "threshold": review["regression_ratio_gt"]})
            base_values = [row[base_key] for row in native
                           if row["case_id"] == case_id and row["role"] == "baseline"]
            cand_values = [row[cand_key] for row in native
                           if row["case_id"] == case_id and row["role"] == "candidate"]
            tail_ratio = ratio(statistics.median(cand_values), statistics.median(base_values),
                               f"{case_id}/{metric}")
            if tail_ratio > review["regression_ratio_gt"]:
                review_flags.append({"case_id": case_id, "metric": metric,
                                     "kind": "tail", "ratio": tail_ratio,
                                         "threshold": review["regression_ratio_gt"]})
        for item in values:
            for metric in ("whole_workflow_p95", "publication_p95"):
                if item[f"{metric}_ratio"] > latency_gate["per_block_ratio_max"]:
                    review_flags.append({"case_id": case_id, "metric": metric,
                                         "kind": "per_block_safety_gate", "block": item["block"],
                                         "ratio": item[f"{metric}_ratio"],
                                         "threshold": latency_gate["per_block_ratio_max"]})
            if item["whole_workflow_ratio"] > latency_gate["per_block_ratio_max"]:
                review_flags.append({"case_id": case_id, "metric": "whole_workflow_p50",
                                     "kind": "per_block_safety_gate", "block": item["block"],
                                     "ratio": item["whole_workflow_ratio"],
                                     "threshold": latency_gate["per_block_ratio_max"]})
            if item["publication_ratio"] > latency_gate["per_block_ratio_max"]:
                review_flags.append({"case_id": case_id, "metric": "publication_p50",
                                     "kind": "per_block_safety_gate", "block": item["block"],
                                     "ratio": item["publication_ratio"],
                                     "threshold": latency_gate["per_block_ratio_max"]})
        for metric in ("whole_workflow_p50", "publication_p50"):
            for role in ROLES:
                values_for_spread = [row[metric] for row in native
                                     if row["case_id"] == case_id and row["role"] == role]
                spread = max(values_for_spread) / min(values_for_spread)
                if spread > review["p50_spread_gt"]:
                    review_flags.append({"case_id": case_id, "role": role,
                                         "metric": metric + "_spread", "ratio": spread,
                                         "threshold": review["p50_spread_gt"]})
        summaries[case_id] = metrics
    return {"cases": summaries, "review_flags": review_flags,
            "all_case_latency_gate_pass": all(
                summaries[case_id][metric]["accepted"]
                for case_id in CASE_ORDER
                for metric in ("whole_workflow_p50", "publication_p50"))
            and all(item["whole_workflow_ratio"] <= latency_gate["per_block_ratio_max"]
                    and item["publication_ratio"] <= latency_gate["per_block_ratio_max"]
                    for values in pairs.values() for item in values)
            and all(item[f"{metric}_ratio"] <= latency_gate["per_block_ratio_max"]
                    for values in pairs.values() for item in values
                    for metric in ("whole_workflow_p95", "publication_p95")),
            "scope": "primary publication p50 improvement, all-row p50 bootstrap safety, and per-block p50/p95 latency safety"}


def child_allocator_metric(row: dict[str, Any], field: str, expected_samples: int) -> list[int]:
    if field == "max_rss_kib":
        value = row.get("max_rss_kib")
        require(type(value) is int and value >= 0, f"{row['label']} RSS is invalid")
        return [value]
    values = row.get("allocator_metrics", {}).get("publication_allocation_metrics")
    require(isinstance(values, list) and values and all(item.get("status") == "measured" for item in values),
            f"{row['label']} publication allocator vector is missing")
    require(len(values) == expected_samples,
            f"{row['label']} publication allocator vector has {len(values)} samples, expected {expected_samples}")
    output = []
    for item in values:
        value = item.get(field)
        require(type(value) is int and value >= 0, f"{row['label']} allocator {field} is invalid")
        output.append(value)
    return output


def allocator_reduce(rows: list[dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    allocator = [row for row in rows if row["stage"] == "allocator"]
    require(len(allocator) == 24, "allocator child count is not 24")
    gate = plan["gate"]["allocation"]
    expected_samples = plan["allocator"]["samples"]
    by_case: dict[str, list[dict[str, Any]]] = {case_id: [] for case_id in CASE_ORDER}
    review_flags: list[dict[str, Any]] = []
    for repeat in range(3):
        for case_id in CASE_ORDER:
            pair = [row for row in allocator if row["index"] == repeat and row["case_id"] == case_id]
            require(len(pair) == 2 and {row["role"] for row in pair} == set(ROLES),
                    f"allocator pair missing for repeat {repeat}/{case_id}")
            base = next(row for row in pair if row["role"] == "baseline")
            cand = next(row for row in pair if row["role"] == "candidate")
            metrics: dict[str, Any] = {}
            # Preserve every measured allocator counter in the reduction.  The
            # frozen safety gate intentionally consumes only a disjoint subset;
            # deallocation/reallocation and absolute live/high-water gauges
            # remain available for audit and replay.
            for field in [*ALLOCATOR_SAMPLE_FIELDS, "max_rss_kib"]:
                base_values = child_allocator_metric(base, field, expected_samples)
                cand_values = child_allocator_metric(cand, field, expected_samples)
                base_sorted = sorted(base_values)
                cand_sorted = sorted(cand_values)
                base_mid = base_sorted[len(base_sorted) // 2]
                cand_mid = cand_sorted[len(cand_sorted) // 2]
                label = f"{case_id}/{field}/{repeat}"
                value = {"baseline_median": base_mid, "candidate_median": cand_mid,
                         "ratio": ratio(cand_mid, base_mid, label),
                         "comparability": ratio_status(cand_mid, base_mid, label)}
                metrics[field] = value
                review_limit = plan["gate"]["review"]["regression_ratio_gt"]
                if value["ratio"] > review_limit:
                    review_flags.append({"case_id": case_id, "repeat": repeat, "metric": field,
                                         "ratio": value["ratio"], "threshold": review_limit})
                if len(base_values) > 1:
                    low = min(base_values)
                    if low > 0 and max(base_values) / low > plan["gate"]["review"]["p50_spread_gt"]:
                        review_flags.append({"case_id": case_id, "repeat": repeat,
                                             "metric": field + "_baseline_sample_spread",
                                             "ratio": max(base_values) / low,
                                             "threshold": plan["gate"]["review"]["p50_spread_gt"]})
            by_case[case_id].append({"repeat": repeat, "metrics": metrics})
    summaries: dict[str, Any] = {}
    for case_id, pairs in by_case.items():
        publication_pairs = [item["metrics"]["allocated_bytes"]["ratio"] for item in pairs]
        per_pair_ok: dict[str, bool] = {}
        safety: dict[str, Any] = {}
        for field in gate["metrics"]:
            values = [item["metrics"][field]["ratio"] for item in pairs]
            safety[field] = {"pair_ratios": values, "paired_median": statistics.median(values),
                             "pair_gate_pass": all(value <= gate["per_pair_ratio_max"] for value in values),
                             "accepted": statistics.median(values) <= gate["paired_median_ratio_max"]
                             and all(value <= gate["per_pair_ratio_max"] for value in values)}
            per_pair_ok[field] = safety[field]["accepted"]
        summaries[case_id] = {
            "pairs": pairs,
            "publication_allocated_bytes_pair_ratios": publication_pairs,
            "per_pair_gate_pass": per_pair_ok,
            "safety": safety,
        }
    return {"cases": summaries, "review_flags": review_flags,
            "all_pair_gate_pass": all(
                summaries[case_id]["per_pair_gate_pass"][field]
                for case_id in CASE_ORDER for field in gate["metrics"]),
            "all_row_safety_gate_pass": all(
                summary["safety"][field]["accepted"]
                for summary in summaries.values() for field in gate["metrics"]),
            "scope": "all five allocator/RSS metrics safety-gated per repeat and by paired median; no allocator bootstrap"}

def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as handle:
            encoded = json.dumps(value, indent=2, sort_keys=True,
                                 ensure_ascii=False, allow_nan=False) + "\n"
            handle.write(encoded)
            handle.flush()
            os.fsync(handle.fileno())
        directory = os.open(path.parent, os.O_DIRECTORY)
        try:
            os.fsync(directory)
        finally:
            os.close(directory)
    except FileExistsError as error:
        raise ReviewError(f"refusing to overwrite {path}") from error


def analysis(plan: dict[str, Any], stage: str, output: Path) -> dict[str, Any]:
    case_rows(plan)
    source = source_check(plan)
    binaries = binaries_check(plan)
    validator = load_validator(plan)
    if stage == "qualification":
        rows = capture_review(plan, "qualification", binaries, validator)
        compare_identities(rows)
        result = {"schema": "litchi.paired-audit-reduction-0948.v1", "status": "pass",
                  "stage": "qualification", "source": source,
                  "rows": rows, "qualification_admitted": True,
                  "scope": "semantic/statistical qualification only; no formal timing reduction"}
    elif stage == "native":
        qualification = capture_review(plan, "qualification", binaries, validator)
        compare_identities(qualification)
        require((PACKET / "qualification-analysis.json").is_file(),
                "native reduction requires qualification-analysis.json")
        native = capture_review(plan, "native", binaries, validator)
        compare_identities(qualification + native)
        reduction = native_reduce(native, plan)
        result = {"schema": "litchi.paired-audit-reduction-0948.v1",
                  "status": "pass" if reduction["all_case_latency_gate_pass"] else "fail",
                  "stage": "native", "source": source, "qualification": qualification,
                  "native": native, "reduction": reduction,
                  "scope": "native whole-workflow and publication latency safety for all four rows"}
    elif stage == "allocator":
        qualification = capture_review(plan, "qualification", binaries, validator)
        compare_identities(qualification)
        allocator = capture_review(plan, "allocator", binaries, validator)
        compare_identities(qualification + allocator)
        require(len(allocator) == 24, "allocator child count is not 24")
        reduction = allocator_reduce(allocator, plan)
        result = {"schema": "litchi.paired-audit-reduction-0948.v1",
                  "status": "pass" if reduction["all_pair_gate_pass"]
                  and reduction["all_row_safety_gate_pass"] else "fail",
                  "stage": "allocator", "source": source, "allocator": allocator,
                  "reduction": reduction,
                  "scope": "allocator hard and safety gates for all four rows"}
    else:
        qualification = capture_review(plan, "qualification", binaries, validator)
        compare_identities(qualification)
        require((PACKET / "qualification-analysis.json").is_file(),
                "full reduction requires qualification-analysis.json")
        native = capture_review(plan, "native", binaries, validator)
        allocator = capture_review(plan, "allocator", binaries, validator)
        compare_identities(qualification + native + allocator)
        reduction = native_reduce(native, plan)
        allocation_reduction = allocator_reduce(allocator, plan)
        result = {"schema": "litchi.paired-audit-reduction-0948.v1",
                  "status": "pass" if reduction["all_case_latency_gate_pass"]
                  and allocation_reduction["all_pair_gate_pass"]
                  and allocation_reduction["all_row_safety_gate_pass"] else "fail",
                  "stage": "all", "source": source, "qualification": qualification,
                  "native": native, "allocator": allocator,
                  "reduction": reduction,
                  "allocation_reduction": allocation_reduction,
                  "scope": "qualification plus native latency and allocator gates for all four rows"}
    write_once(output, result)
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=["qualification", "native", "allocator", "all"])
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    plan = load_plan()
    default = {"qualification": PACKET / "qualification-analysis.json",
               "native": PACKET / "analysis.json", "allocator": PACKET / "allocator-analysis.json",
               "all": PACKET / "analysis.json"}[args.action]
    result = analysis(plan, args.action, args.output or default)
    print(json.dumps({"status": result["status"], "stage": result["stage"]}, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except (ReviewError, AssertionError, OSError, ValueError) as error:
        print(f"0948 reduction failed: {error}", file=sys.stderr)
        raise SystemExit(1)
