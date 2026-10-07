"""Source-bound capture driver for the frozen 0945 XLSX sharing experiment.

The driver deliberately contains no analysis or acceptance decision.  It runs
only the frozen child schedule described by ``plan-v2.json`` and records enough
custody information for a later, independent reducer.  A child is never
retried: stdout, stderr, the JSON report, resource observation, and a receipt
are retained even when the child exits unsuccessfully.

The root coordinator writes ``plan-v2.json`` after both roles have been built.
This module is kept packet-local so a later experiment cannot silently reuse
the 0945 schedule or source/binary bindings.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import signal
import subprocess
import sys
import time
from typing import Any


PACKET = Path(__file__).resolve().parent.parent
REPO = Path("/home/zhuhe/code/litchi")
PLAN_RELATIVE = "plan-v2.json"
PLAN_PATH = PACKET / PLAN_RELATIVE
SCHEMA = "litchi.paired-audit-capture-0945.v1"
ROLES = ("baseline", "candidate")
ROLE_KINDS = ("native", "allocator")
BUILD_REVISIONS = {"baseline": "v1", "candidate": "v2"}
BUILD_BINDING_IDS = {"baseline": "baseline1", "candidate": "candidate2"}
SAFE_NAME = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")
MAX_JSON_BYTES = 64 * 1024 * 1024


class CaptureError(RuntimeError):
    """A frozen-capture precondition or child-capture failure."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CaptureError(message)


def read_json(path: Path) -> Any:
    def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise CaptureError(f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result
    try:
        require(path.is_file() and not path.is_symlink(), f"JSON artifact is not a regular file: {path}")
        require(path.stat().st_size <= MAX_JSON_BYTES,
                f"JSON artifact exceeds {MAX_JSON_BYTES} bytes: {path}")
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_duplicates)
    except (OSError, json.JSONDecodeError, CaptureError) as error:
        raise CaptureError(f"cannot read JSON {path}: {error}") from error

def canonical_json(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode()


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as source:
            for chunk in iter(lambda: source.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as error:
        raise CaptureError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest()


def resolve_path(value: str | Path, base: Path = PACKET) -> Path:
    path = Path(value)
    return path if path.is_absolute() else (base / path).absolute()


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def process_group_clear(pgid: int | None) -> bool:
    if pgid is None:
        return True
    for _ in range(100):
        try:
            os.killpg(pgid, 0)
        except ProcessLookupError:
            return True
        except PermissionError:
            return False
        time.sleep(0.05)
    return False


def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite retained artifact {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False,
                          allow_nan=False) + "\n")
    try:
        with path.open("x", encoding="utf-8") as output:
            output.write(encoded)
            output.flush()
            os.fsync(output.fileno())
        directory = os.open(path.parent, os.O_DIRECTORY)
        try:
            os.fsync(directory)
        finally:
            os.close(directory)
    except FileExistsError as error:
        raise CaptureError(f"refusing to overwrite retained artifact {path}") from error

def file_descriptor(path: Path) -> dict[str, Any]:
    path = Path(path)
    require(not path.is_symlink(), f"artifact must not be a symlink {path}")
    path = path.resolve()
    require(path.exists(), f"missing artifact {path}")
    require(path.is_file(), f"artifact is not a regular file {path}")
    try:
        packet_path = path.relative_to(PACKET.resolve()).as_posix()
    except ValueError as error:
        raise CaptureError(f"artifact descriptor escaped packet: {path}") from error
    stat = path.stat()
    return {
        "path": packet_path,
        "bytes": stat.st_size,
        "mode": stat.st_mode & 0o777,
        "executable": bool(stat.st_mode & 0o111),
        "sha256": sha256_file(path),
    }


def descriptor_path(value: Any, label: str) -> Path:
    if isinstance(value, (str, Path)):
        raw = Path(value)
        require(not raw.is_absolute() and raw.as_posix() == str(value)
                and ".." not in raw.parts,
                f"{label} descriptor path must be packet-relative")
        return resolve_path(raw)
    require(isinstance(value, dict), f"{label} must be a path or descriptor")
    require(isinstance(value.get("path"), str), f"{label}.path missing")
    raw = Path(value["path"])
    require(not raw.is_absolute() and raw.as_posix() == value["path"]
            and ".." not in raw.parts,
            f"{label} descriptor path must be packet-relative")
    return resolve_path(raw)


def verify_descriptor(value: Any, label: str) -> dict[str, Any]:
    """Verify a plan-owned descriptor without trusting its recorded metadata."""

    path = descriptor_path(value, label)
    actual = file_descriptor(path)
    if isinstance(value, dict):
        for key in ("bytes", "sha256"):
            if key in value:
                require(actual[key] == value[key], f"{label}.{key} does not match {path}")
        if "mode" in value:
            require(actual["mode"] == value["mode"], f"{label}.mode does not match {path}")
        if "executable" in value:
            require(actual["executable"] == value["executable"],
                    f"{label}.executable does not match {path}")
    return actual


def plan_descriptor() -> dict[str, Any]:
    """Return the packet-relative descriptor for this versioned plan."""
    return verify_descriptor({"path": PLAN_RELATIVE}, "plan")


def load_plan() -> dict[str, Any]:
    require(PLAN_PATH.is_file(), f"missing frozen plan {PLAN_PATH}")
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan must be an object")
    require(plan.get("schema") == SCHEMA, f"unexpected plan schema {plan.get('schema')!r}")
    require(plan.get("plan_id") == "0945" and plan.get("protocol_revision") == 1,
            "plan identity or protocol revision changed")
    require(plan.get("build_revisions") == BUILD_REVISIONS
            and plan.get("build_binding_ids") == BUILD_BINDING_IDS,
            "plan build revision bindings changed")
    require(plan.get("status") in {"frozen", "capture-ready", "ready"},
            f"plan is not capture-ready: {plan.get('status')!r}")
    require(plan.get("seed") == 945009, "plan seed changed")
    require(plan.get("cases") and len(plan["cases"]) == 4, "plan must contain four rows")
    qualification = plan.get("qualification")
    require(qualification == {"children": 16, "kinds": ["native", "allocator"],
                              "roles": ["baseline", "candidate"], "samples": 2, "warmup": 0},
            "qualification protocol changed")
    require(plan.get("native") == {
        "blocks": 6, "children": 48, "samples": 100, "warmup": 10,
        "schedule": "six alternating blocks; even baseline then candidate and forward case order; odd candidate then baseline and reverse case order",
    }, "native protocol changed")
    require(plan.get("allocator") == {
        "repeats": 2, "children": 16, "samples": 3, "warmup": 2,
        "schedule": "two alternating repeats; even baseline then candidate and forward case order; odd candidate then baseline and reverse case order",
    }, "allocator protocol changed")
    env = plan.get("environment")
    require(isinstance(env, dict) and env.get("fixed") == {
        "PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C", "TZ": "UTC",
        "RUST_BACKTRACE": "0",
    }, "fixed runtime environment changed")
    resource = plan.get("resource")
    require(resource == {
        "deadline_seconds": 900,
        "enabled": True,
        "format": '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
        "process_group": True,
        "time_path": "/usr/bin/time",
    }, "resource/process-group protocol changed")
    gate = plan.get("gate")
    require(isinstance(gate, dict), "plan.gate is missing")
    require(gate.get("bootstrap") == {
        "confidence_level": 0.95, "lower_percentile": 2.5, "replicates": 10000,
        "seed": 945009, "statistic": "paired_median_ratio", "upper_percentile": 97.5,
    }, "latency bootstrap gate changed")
    require(gate.get("latency") == {
        "case_ids": ["primary-percent-medium", "primary-percent-dense-sparse",
                      "control-xlsx-one-edit-medium", "control-xlsx-managed-one-edit-medium"],
        "bootstrap_upper_95_lt": 1.05,
        "metrics": ["whole_workflow_p50", "publication_p50"],
        "paired_median_ratio_max": 1.05,
    }, "latency safety gate changed")
    require(gate.get("allocation") == {
        "individual_review_ratio_gt": 1.05,
        "primary_bootstrap": False,
        "primary_case_ids": ["primary-percent-medium", "primary-percent-dense-sparse"],
        "primary_every_pair_ratio_max": 0.95,
        "primary_pair_count": 2,
        "publication_metric": "allocated_bytes",
        "safety_metrics": ["allocation_calls", "region_peak_live_bytes",
                            "live_bytes_after", "max_rss_kib"],
        "safety_paired_median_ratio_max": 1.05,
    }, "allocation gate changed")
    require(gate.get("review") == {
        "p50_spread_gt": 1.2, "regression_ratio_gt": 1.05,
        "tail_metrics": ["p50", "p95", "p99"],
    }, "review thresholds changed")
    return plan

def environment(plan: dict[str, Any]) -> dict[str, str]:
    value = plan.get("environment")
    require(isinstance(value, dict) and value.get("fixed") is not None,
            "plan.environment.fixed is missing")
    fixed = value["fixed"]
    expected = {"PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C",
                "TZ": "UTC", "RUST_BACKTRACE": "0"}
    require(fixed == expected, "runtime environment is not the exact frozen environment")
    return dict(expected)

def source_specs(plan: dict[str, Any]) -> tuple[Any, Any, set[str], Path, dict[str, Path]]:
    source = plan.get("source")
    require(isinstance(source, dict), "plan.source must be an object")
    baseline = source.get("baseline_manifest")
    candidate = source.get("candidate_manifest")
    names = source.get("candidate_files")
    require(baseline is not None and candidate is not None,
            "plan.source requires bound baseline and candidate manifests")
    require(isinstance(names, list) and names,
            "plan.source.candidate_files must be a nonempty frozen list")
    candidate_files: set[str] = set()
    for name in names:
        require(isinstance(name, str) and name and not Path(name).is_absolute(),
                f"invalid candidate source path {name!r}")
        path = Path(name)
        require(".." not in path.parts and path.as_posix() == name,
                f"candidate source path is not normalized {name!r}")
        candidate_files.add(name)
    root_value = source.get("root_worktree", str(REPO))
    require(isinstance(root_value, str), "source.root_worktree must be a path")
    root = resolve_path(root_value, PACKET)
    roots: dict[str, Path] = {}
    for role, key in (("baseline", "baseline_root"), ("candidate", "candidate_root")):
        value = source.get(key)
        require(isinstance(value, str), f"source.{key} must be a path")
        roots[role] = resolve_path(value, PACKET)
    return baseline, candidate, candidate_files, root, roots

def manifest_files(value: Any, label: str) -> tuple[Path, dict[str, dict[str, Any]]]:
    path = descriptor_path(value, label)
    verify_descriptor(value, label)
    raw = read_json(path)
    files = raw.get("files") if isinstance(raw, dict) else raw
    normalized: dict[str, dict[str, Any]] = {}
    if isinstance(files, list):
        entries = files
        for item in entries:
            require(isinstance(item, dict) and isinstance(item.get("path"), str),
                    f"{label} has a malformed list entry")
            name = item["path"]
            require(name not in normalized, f"{label} contains duplicate path {name}")
            normalized[name] = item
    elif isinstance(files, dict):
        for name, item in files.items():
            require(isinstance(item, dict), f"{label} entry {name} is malformed")
            normalized[name] = item
    else:
        raise CaptureError(f"{label} has no files map")
    require(normalized, f"{label} has no files")
    for name, item in normalized.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute(),
                f"{label} has invalid path {name!r}")
        path_name = Path(name)
        require(".." not in path_name.parts and path_name.as_posix() == name,
                f"{label} has non-normalized path {name!r}")
        require(isinstance(item.get("sha256"), str) and len(item["sha256"]) == 64,
                f"{label} entry {name} lacks sha256")
        require(isinstance(item.get("bytes"), int) and item["bytes"] >= 0,
                f"{label} entry {name} lacks bytes")
    return path, normalized

def source_check(plan: dict[str, Any], full: bool = True) -> dict[str, Any]:
    """Check both immutable role source roots against their bound manifests."""
    baseline_spec, candidate_spec, candidate_names, root, roots = source_specs(plan)
    require(root.is_dir(), f"source root does not exist: {root}")
    _, baseline = manifest_files(baseline_spec, "baseline source manifest")
    _, candidate = manifest_files(candidate_spec, "candidate source manifest")
    require(set(baseline) == set(candidate), "baseline/candidate source file sets differ")
    changed = {name for name in baseline
               if (baseline[name].get("sha256"), baseline[name].get("bytes"))
               != (candidate[name].get("sha256"), candidate[name].get("bytes"))}
    require(changed == candidate_names,
            f"source delta is not exactly candidate_files: {sorted(changed)}")
    checked: dict[str, int] = {}
    for role, expected in (("baseline", baseline), ("candidate", candidate)):
        current_root = roots[role]
        require(current_root.is_dir(), f"{role} source root does not exist: {current_root}")
        names = sorted(expected) if full else sorted(candidate_names)
        for name in names:
            actual = file_descriptor(current_root / name)
            wanted = expected[name]
            require(actual["bytes"] == wanted.get("bytes") and actual["sha256"] == wanted["sha256"],
                    f"{role} source differs at {name}")
        checked[role] = len(names)
    return {
        "baseline_manifest": verify_descriptor(baseline_spec, "baseline source manifest"),
        "candidate_manifest": verify_descriptor(candidate_spec, "candidate source manifest"),
        "candidate_files": sorted(candidate_names),
        "checked_files": checked,
        "full": full,
        "root_worktree": str(root),
        "role_roots": {role: str(path) for role, path in roots.items()},
    }

def role_descriptors(plan: dict[str, Any], role: str) -> dict[str, Any]:
    require(role in ROLES, f"unknown role {role}")
    roles = plan.get("roles", plan.get("binaries"))
    require(isinstance(roles, dict) and role in roles and isinstance(roles[role], dict),
            f"missing role descriptor {role}")
    result: dict[str, Any] = {}
    for kind in ROLE_KINDS:
        value = roles[role].get(kind)
        require(isinstance(value, dict), f"missing {role} {kind} descriptor")
        require(value.get("build_revision") == BUILD_REVISIONS[role]
                and value.get("build_binding_id") == BUILD_BINDING_IDS[role],
                f"{role} {kind} build revision binding changed")
        result[kind] = value
    return result


def role_input_bindings(raw: dict[str, Any], plan: dict[str, Any], role: str,
                        label: str) -> None:
    """Bind a build receipt's retained source census to its role manifest."""
    before = raw.get("inputs_before")
    after = raw.get("inputs_after")
    if before is None or after is None:
        require(before is None and after is None,
                f"{label} has only one side of its input census")
        # The current 0945 builders intentionally publish the immutable
        # census under nested source.manifest/source_root fields.  Those
        # fields are checked by reuse_v2.validate_build_terminal before this
        # helper is reached; an absent source object is never accepted.
        require(isinstance(raw.get("source"), dict),
                f"{label} lacks nested source provenance")
        return
    require(isinstance(before, list) and before and before == after,
            f"{label} inputs census is missing or changed")
    baseline_spec, candidate_spec, _, _, roots = source_specs(plan)
    manifest_spec = baseline_spec if role == "baseline" else candidate_spec
    _, expected_manifest = manifest_files(manifest_spec, f"{role} source manifest")
    expected_root = roots[role].resolve()
    indexed: dict[str, list[dict[str, Any]]] = {}
    for item in after:
        require(isinstance(item, dict) and isinstance(item.get("path"), str),
                f"{label} has malformed input descriptor")
        indexed.setdefault(str(Path(item["path"]).resolve()), []).append(item)
    for name, expected in expected_manifest.items():
        matches = indexed.get(str((expected_root / name).resolve()), [])
        require(len(matches) == 1, f"{label} does not bind exactly one input for {name}")
        item = matches[0]
        require(item.get("bytes") == expected.get("bytes") and
                item.get("sha256") == expected.get("sha256"),
                f"{label} {role} input differs from manifest: {name}")

def _reuse_validator() -> Any:
    path = PACKET / "tools" / "reuse_v2.py"
    require(path.is_file() and not path.is_symlink(),
            f"missing shared provenance validator {path}")
    spec = importlib.util.spec_from_file_location("litchi_0945_reuse_v2_capture", path)
    require(spec is not None and spec.loader is not None,
            "cannot load shared provenance validator")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def check_build_receipt(value: Any, label: str, role: str, plan: dict[str, Any],
                        *, kind: str,
                        binary: dict[str, Any] | None = None) -> dict[str, Any]:
    actual = verify_descriptor(value, label)
    raw = read_json(descriptor_path(value, label))
    require(isinstance(raw, dict), f"{label} is not an object")
    validator = _reuse_validator()
    if raw.get("schema") == "litchi.performance.0945.native-baseline-reuse-v2":
        try:
            validator.validate_receipt(Path(actual["path"]), binary=binary, plan=plan)
        except Exception as error:
            raise CaptureError(f"{label} portable reuse validation failed: {error}") from error
        return actual
    require(raw.get("schema") != "litchi.performance.0945.native-baseline-reuse-v1",
            f"{label} uses the superseded nonportable reuse receipt")
    if "exit_code" in raw:
        require(raw.get("exit_code") == 0, f"{label} does not record exit_code=0")
    elif "returncode" in raw:
        require(raw.get("returncode") == 0, f"{label} does not record returncode=0")
    else:
        raise CaptureError(f"{label} has no successful exit status")
    if "timed_out" in raw:
        require(raw.get("timed_out") is False, f"{label} timed out")
    if "deadline_exceeded" in raw:
        require(raw.get("deadline_exceeded") is False, f"{label} exceeded its deadline")
    if "terminal_before_parse" in raw:
        require(raw["terminal_before_parse"] is True, f"{label} terminal was not durable before parse")
    if "process_group_clear" in raw:
        require(raw["process_group_clear"] is True, f"{label} process group was not cleared")
    if isinstance(raw.get("cleanup"), dict) and "process_group_clear" in raw["cleanup"]:
        require(raw["cleanup"]["process_group_clear"] is True,
                f"{label} build process group was not cleared")
    try:
        validator.validate_build_terminal(raw, role, kind, plan, binary=binary)
    except Exception as error:
        raise CaptureError(f"{label} nested source/build validation failed: {error}") from error
    role_input_bindings(raw, plan, role, label)
    return actual

def check_binary(plan: dict[str, Any], role: str, kind: str) -> dict[str, Any]:
    descriptor = role_descriptors(plan, role)[kind]
    actual = verify_descriptor(descriptor, f"{role} {kind} binary")
    require(actual["executable"], f"{role} {kind} binary is not executable")
    build = descriptor.get("build_receipt")
    require(build is not None, f"{role} {kind} binary lacks build_receipt")
    build_actual = check_build_receipt(build, f"{role} {kind} build receipt", role, plan,
                                       kind=kind, binary=actual)
    return {"binary": actual, "build_receipt": build_actual}


def check_binaries(plan: dict[str, Any]) -> dict[str, Any]:
    return {role: {kind: check_binary(plan, role, kind) for kind in ROLE_KINDS}
            for role in ROLES}


def case_rows(plan: dict[str, Any]) -> list[dict[str, Any]]:
    expected = [
        ("primary-percent-medium", "primary", "xlsx_source_backed_cell_values_one_percent_edit_save", "medium"),
        ("primary-percent-dense-sparse", "primary", "xlsx_source_backed_cell_values_one_percent_edit_save", "dense-sparse"),
        ("control-xlsx-one-edit-medium", "control", "xlsx_source_backed_cell_values_one_edit_save", "medium"),
        ("control-xlsx-managed-one-edit-medium", "managed-control", "xlsx_source_backed_managed_cell_values_one_edit_save", "medium"),
    ]
    rows = plan.get("cases")
    require(isinstance(rows, list) and len(rows) == 4, "plan must define exactly four XLSX cases")
    by_id: dict[str, dict[str, Any]] = {}
    for row in rows:
        require(isinstance(row, dict) and isinstance(row.get("id"), str)
                and SAFE_NAME.fullmatch(row["id"]), "malformed case row")
        require(row["id"] not in by_id, f"duplicate case id {row['id']}")
        by_id[row["id"]] = row
    ordered: list[dict[str, Any]] = []
    for case_id, kind, case, shape in expected:
        row = by_id.get(case_id)
        require(row is not None and row.get("kind") == kind and row.get("case") == case
                and row.get("xlsx_cell_crud_shape") == shape,
                f"case definition changed for {case_id}")
        ordered.append(row)
    require([row["id"] for row in rows] == [row["id"] for row in ordered],
            "plan case order changed")
    return ordered

def native_settings(plan: dict[str, Any]) -> tuple[int, int, int, int]:
    native = plan.get("native")
    qualification = plan.get("qualification")
    require(isinstance(native, dict) and isinstance(qualification, dict),
            "native and qualification settings are required")
    require((native.get("blocks"), native.get("samples"), native.get("warmup"),
             qualification.get("samples"), qualification.get("warmup")) == (6, 100, 10, 2, 0),
            "native/qualification counts changed")
    return 6, 100, 10, 2

def allocator_settings(plan: dict[str, Any]) -> tuple[int, int, int]:
    value = plan.get("allocator")
    require(isinstance(value, dict), "allocator settings are required")
    require((value.get("repeats"), value.get("samples"), value.get("warmup")) == (2, 3, 2),
            "allocator counts changed")
    return 2, 3, 2

def cpu(plan: dict[str, Any]) -> int:
    value = plan.get("cpu", plan.get("native", {}).get("cpu"))
    require(isinstance(value, int) and value >= 0, "plan.cpu must be a nonnegative integer")
    return value


def workload_cwd(plan: dict[str, Any]) -> Path:
    value = plan.get("workload_cwd")
    require(isinstance(value, str), "workload_cwd must be a path")
    path = resolve_path(value, PACKET)
    require(path.is_dir(), f"workload cwd does not exist {path}")
    return path

def command_for(plan: dict[str, Any], role: str, kind: str, row: dict[str, Any],
                samples: int, warmup: int, folder: Path) -> list[str]:
    descriptor = role_descriptors(plan, role)[kind]
    binary = resolve_path(descriptor["path"])
    command: list[str] = ["/usr/bin/taskset", "-c", str(cpu(plan))]
    resource = plan.get("resource", {})
    if isinstance(resource, dict) and resource.get("enabled", True):
        time_path = resource.get("time_path", "/usr/bin/time")
        require(isinstance(time_path, str) and time_path, "resource.time_path must be a string")
        resource_format = resource.get(
            "format",
            '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
        )
        require(isinstance(resource_format, str), "resource.format must be a string")
        command += [time_path, "-f", resource_format, "-o", str(folder / "resource.json")]
    command += [str(binary), "--warmup", str(warmup), "--samples", str(samples),
                "--case", row["case"]]
    if row["case"].startswith("xlsx_"):
        shape = row.get("xlsx_cell_crud_shape", row.get("shape"))
        require(isinstance(shape, str), f"{row['id']} lacks xlsx_cell_crud_shape")
        command += ["--xlsx-cell-crud-shape", shape]
    elif row["case"].startswith("docx_"):
        shape = row.get("semantic_shape", row.get("shape"))
        require(isinstance(shape, str), f"{row['id']} lacks semantic_shape")
        command += ["--semantic-shape", shape]
    else:
        raise CaptureError(f"unsupported 0905 case {row['case']}")
    command += ["--json", str(folder / "report.json")]
    return command


def expected_jobs(plan: dict[str, Any], stage: str) -> list[dict[str, Any]]:
    cases = case_rows(plan)
    jobs: list[dict[str, Any]] = []
    if stage == "qualification":
        for kind in ROLE_KINDS:
            for role in ROLES:
                for row in cases:
                    jobs.append({"stage": stage, "index": 0, "role": role, "kind": kind,
                                 "case": row["id"], "row": row})
    elif stage == "native":
        blocks, _, _, _ = native_settings(plan)
        for block in range(blocks):
            roles = ROLES if block % 2 == 0 else tuple(reversed(ROLES))
            block_cases = cases if block % 2 == 0 else list(reversed(cases))
            for row in block_cases:
                for role in roles:
                    jobs.append({"stage": stage, "index": block, "role": role, "kind": "native",
                                 "case": row["id"], "row": row})
    elif stage == "allocator":
        repeats, _, _ = allocator_settings(plan)
        for repeat in range(repeats):
            roles = ROLES if repeat % 2 == 0 else tuple(reversed(ROLES))
            cases_order = cases if repeat % 2 == 0 else list(reversed(cases))
            for row in cases_order:
                for role in roles:
                    jobs.append({"stage": stage, "index": repeat, "role": role, "kind": "allocator",
                                 "case": row["id"], "row": row})
    else:
        raise CaptureError(f"unknown capture stage {stage}")
    return jobs

def child_name(job: dict[str, Any]) -> str:
    prefix = {"qualification": "q", "native": "b", "allocator": "r"}[job["stage"]]
    kind = f"-{job['kind']}" if job["stage"] == "qualification" else ""
    return f"{prefix}{job['index']:02d}-{job['role']}{kind}-{job['case']}"

def stage_counts(plan: dict[str, Any], stage: str) -> tuple[int, int, int]:
    jobs = expected_jobs(plan, stage)
    if stage == "qualification":
        samples, warmup = 2, 0
    elif stage == "native":
        _, samples, warmup, _ = native_settings(plan)
    else:
        _, samples, warmup = allocator_settings(plan)
    return len(jobs), samples, warmup


def resource_enabled(plan: dict[str, Any]) -> bool:
    value = plan.get("resource", {})
    return not isinstance(value, dict) or value.get("enabled", True)


def artifact_descriptors(folder: Path) -> list[dict[str, Any]]:
    allowed = {"stdout.txt", "stderr.txt", "report.json", "resource.json",
               "terminal.json", "validation-failure.json"}
    artifacts: list[dict[str, Any]] = []
    for path in sorted(folder.iterdir()):
        if path.name in {"start.json", "command.json", "receipt.json"}:
            continue
        require(path.name in allowed, f"unexpected capture artifact {path}")
        require(not path.is_symlink() and path.is_file(),
                f"capture artifact is not a regular file {path}")
        artifacts.append(file_descriptor(path))
    return artifacts

def run_child(plan: dict[str, Any], job: dict[str, Any], binaries: dict[str, Any],
              source_start: dict[str, Any], env: dict[str, str], cwd: Path) -> dict[str, Any]:
    stage = job["stage"]
    name = child_name(job)
    folder = PACKET / stage / name
    require(not folder.exists(), f"capture output already exists; refusing retry {folder}")
    folder.mkdir(parents=True)
    role = job["role"]
    kind = job["kind"]
    descriptor = role_descriptors(plan, role)[kind]
    binary = binaries[role][kind]["binary"]
    require(binary == verify_descriptor(descriptor, f"{role} {kind} binary"),
            f"{role} {kind} binary changed before {name}")
    if stage == "qualification":
        samples, warmup = 2, 0
    elif stage == "native":
        _, samples, warmup, _ = native_settings(plan)
    else:
        _, samples, warmup = allocator_settings(plan)
    command = command_for(plan, role, kind, job["row"], samples, warmup, folder)
    started = utc_now()
    plan_desc = plan_descriptor()
    runner_desc = file_descriptor(Path(__file__))
    start_record = {
        "schema": SCHEMA,
        "stage": stage,
        "index": job["index"],
        "role": role,
        "kind": kind,
        "case_id": job["case"],
        "case": job["row"],
        "argv": command,
        "cwd": str(cwd),
        "environment": env,
        "environment_sha256": sha256_bytes(canonical_json(env)),
        "plan": plan_desc,
        "helper": runner_desc,
        "binary": binary,
        "source": source_start,
        "started_utc": started,
        "process_group": True,
        "start_new_session": True,
        "deadline_seconds": plan["resource"]["deadline_seconds"],
        "terminal_receipt": "terminal.json",
    }
    write_once(folder / "start.json", start_record)

    exit_code: int | None = None
    exception: str | None = None
    pid: int | None = None
    pgid: int | None = None
    timed_out = False
    tick = time.monotonic()
    try:
        with (folder / "stdout.txt").open("x", encoding="utf-8") as stdout, \
             (folder / "stderr.txt").open("x", encoding="utf-8") as stderr:
            try:
                child = subprocess.Popen(command, cwd=cwd, env=env, stdout=stdout, stderr=stderr,
                                         start_new_session=True)
                pid = child.pid
                pgid = os.getpgid(pid)
                deadline = time.monotonic() + float(plan["resource"]["deadline_seconds"])
                while True:
                    exit_code = child.poll()
                    if exit_code is not None:
                        break
                    if time.monotonic() >= deadline:
                        timed_out = True
                        try:
                            os.killpg(pgid, signal.SIGTERM)
                        except ProcessLookupError:
                            pass
                        try:
                            child.wait(timeout=10)
                        except subprocess.TimeoutExpired:
                            try:
                                os.killpg(pgid, signal.SIGKILL)
                            except ProcessLookupError:
                                pass
                            child.wait()
                        exit_code = child.returncode
                        break
                    time.sleep(0.2)
            except (OSError, subprocess.SubprocessError) as error:
                exception = repr(error)
    except (OSError, subprocess.SubprocessError) as error:
        exception = repr(error)
    elapsed = time.monotonic() - tick
    group_clear = process_group_clear(pgid)

    source_end: dict[str, Any] | None = None
    source_error: str | None = None
    try:
        source_end = source_check(plan, full=False)
    except CaptureError as error:
        source_error = str(error)

    artifact_error: str | None = None
    try:
        artifacts_before_terminal = artifact_descriptors(folder)
    except CaptureError as error:
        artifact_error = str(error)
        artifacts_before_terminal = []
    terminal = {
        "schema": SCHEMA,
        "stage": stage,
        "index": job["index"],
        "role": role,
        "kind": kind,
        "case_id": job["case"],
        "case": job["row"],
        "argv": command,
        "cwd": str(cwd),
        "environment": env,
        "environment_sha256": sha256_bytes(canonical_json(env)),
        "plan": plan_desc,
        "helper": runner_desc,
        "binary": binary,
        "source": source_start,
        "source_end": source_end,
        "source_error": source_error,
        "started_utc": started,
        "terminal_utc": utc_now(),
        "pid": pid,
        "process_group_id": pgid,
        "returncode": exit_code,
        "timed_out": timed_out,
        "launch_error": exception,
        "artifact_error": artifact_error,
        "process_group_clear": group_clear,
        "terminal_before_parse": True,
        "deadline_seconds": plan["resource"]["deadline_seconds"],
        "artifacts": artifacts_before_terminal,
    }
    write_once(folder / "terminal.json", terminal)

    failures: list[str] = []
    if exception:
        failures.append(f"subprocess exception: {exception}")
    if exit_code != 0:
        failures.append(f"child exit code: {exit_code}")
    if timed_out:
        failures.append("child exceeded frozen deadline")
    if not group_clear:
        failures.append("child process group was not cleared")
    if source_error:
        failures.append(f"source changed during child: {source_error}")
    if artifact_error:
        failures.append(f"artifact custody failed: {artifact_error}")
    if exit_code == 0 and not (folder / "report.json").is_file():
        failures.append("child exited zero without report.json")
    if resource_enabled(plan) and exit_code == 0 and not (folder / "resource.json").is_file():
        failures.append("child exited zero without resource.json")
    if failures:
        write_once(folder / "validation-failure.json", {
            "schema": SCHEMA,
            "status": "fail",
            "errors": failures,
            "terminal": file_descriptor(folder / "terminal.json"),
            "retained_stdout": (folder / "stdout.txt").is_file(),
            "retained_stderr": (folder / "stderr.txt").is_file(),
        })

    finished = utc_now()
    final_artifacts = artifact_descriptors(folder)
    command_record = {
        **start_record,
        "terminal": terminal,
        "terminal_receipt": file_descriptor(folder / "terminal.json"),
        "finished_utc": finished,
        "elapsed_seconds": elapsed,
        "exit_code": exit_code,
        "exception": exception,
        "source_end": source_end,
        "source_error": source_error,
        "artifacts": final_artifacts,
    }
    write_once(folder / "command.json", command_record)
    receipt = {
        "schema": SCHEMA,
        "status": "pass" if not failures else "fail",
        "stage": stage,
        "index": job["index"],
        "role": role,
        "kind": kind,
        "case_id": job["case"],
        "argv": command,
        "cwd": str(cwd),
        "environment": env,
        "environment_sha256": sha256_bytes(canonical_json(env)),
        "plan": plan_desc,
        "helper": runner_desc,
        "binary": binary,
        "source_start": source_start,
        "source_end": source_end,
        "source_error": source_error,
        "terminal": file_descriptor(folder / "terminal.json"),
        "started_utc": started,
        "finished_utc": finished,
        "elapsed_seconds": elapsed,
        "exit_code": exit_code,
        "exception": exception,
        "errors": failures,
        "artifacts": final_artifacts,
    }
    write_once(folder / "receipt.json", receipt)
    require(not failures, f"capture child failed ({name}); retained terminal and stderr")
    return receipt

def stage_capture(plan: dict[str, Any], stage: str) -> None:
    if stage in {"native", "allocator"}:
        qualification_capture = PACKET / "qualification" / "capture.json"
        require(qualification_capture.is_file(), "formal capture requires completed qualification")
        q = read_json(qualification_capture)
        require(q.get("status") == "pass" and q.get("children") == 16,
                "qualification is not sealed with all 16 children")
    source_start = source_check(plan, full=True)
    binaries = check_binaries(plan)
    env = environment(plan)
    cwd = workload_cwd(plan)
    jobs = expected_jobs(plan, stage)
    expected_children, samples, warmup = stage_counts(plan, stage)
    require(len(jobs) == expected_children, "internal schedule count mismatch")
    receipts = []
    for job in jobs:
        receipts.append(run_child(plan, job, binaries, source_start, env, cwd))
    source_end = source_check(plan, full=True)
    write_once(PACKET / stage / "capture.json", {
        "schema": SCHEMA,
        "status": "pass",
        "stage": stage,
        "children": len(receipts),
        "samples_per_child": samples,
        "warmup_per_child": warmup,
        "total_samples": len(receipts) * samples,
        "jobs": [{"name": child_name(job), "stage": job["stage"], "index": job["index"],
                  "role": job["role"], "kind": job["kind"], "case_id": job["case"]} for job in jobs],
        "plan": plan_descriptor(),
        "helper": file_descriptor(Path(__file__)),
        "environment": env,
        "environment_sha256": sha256_bytes(canonical_json(env)),
        "source_start": source_start,
        "source_end": source_end,
        "binary_bindings": binaries,
        "receipts": receipts,
    })
    print(f"{stage}: {len(receipts)} children / {len(receipts) * samples} samples passed", flush=True)

def check_only(plan: dict[str, Any]) -> None:
    rows = case_rows(plan)
    native_settings(plan)
    allocator_settings(plan)
    source = source_check(plan, full=True)
    binaries = check_binaries(plan)
    env = environment(plan)
    cwd = workload_cwd(plan)
    expected = {stage: len(expected_jobs(plan, stage))
                for stage in ("qualification", "native", "allocator")}
    print(json.dumps({"status": "pass", "schema": SCHEMA, "cases": [row["id"] for row in rows],
                     "expected_children": expected, "cwd": str(cwd), "environment": env,
                     "source": source, "binaries": binaries}, indent=2, sort_keys=True))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=["check", "qualification", "native", "allocator", "all"])
    args = parser.parse_args()
    plan = load_plan()
    if args.action == "check":
        check_only(plan)
        return
    if args.action == "all":
        stage_capture(plan, "qualification")
        stage_capture(plan, "native")
        stage_capture(plan, "allocator")
        return
    stage_capture(plan, args.action)


if __name__ == "__main__":
    try:
        main()
    except CaptureError as error:
        print(f"0945 capture failed: {error}", file=sys.stderr)
        raise SystemExit(1)
