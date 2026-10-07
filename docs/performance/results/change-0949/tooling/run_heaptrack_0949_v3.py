#!/usr/bin/env python3
"""Inert-until-GO serial heaptrack capture driver for the 0949 diagnostic.

This driver has no reducer and makes no performance or adoption decision.  It
validates the frozen 0948 allocator binaries/source closures and, only when
invoked with both ``--capture`` and ``--go``, runs the four fixed heaptrack
children in order.  It writes each child's start receipt before launch and its
terminal receipt after reaping/clearing the process group, before any later
analysis is allowed to inspect reports or profiles.  A failed child is retained
and the schedule continues without retrying it.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import signal
import subprocess
import sys
import time
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
PLAN_PATH = ROOT / "heaptrack-plan-v3.json"
SCHEMA = "litchi.performance.0949.heaptrack-allocation-diagnostic.v1"
CHILD_SCHEMA = "litchi.performance.0949.heaptrack-child-terminal.v1"
RUNNER_SCHEMA = "litchi.performance.0949.heaptrack-runner-terminal.v1"
MAX_JSON_BYTES = 64 * 1024 * 1024
FIXED_ENVIRONMENT = {
    "PATH": "/usr/bin:/bin",
    "LANG": "C",
    "LC_ALL": "C",
    "RUST_BACKTRACE": "0",
    "TZ": "UTC",
}


class CaptureError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CaptureError(message)


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat(timespec="microseconds")


def canonical_json(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"),
                      ensure_ascii=False, allow_nan=False).encode("utf-8")


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"not a regular JSON file: {path}")
    require(path.stat().st_size <= MAX_JSON_BYTES,
            f"JSON file exceeds limit ({MAX_JSON_BYTES}): {path}")

    def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result

    try:
        return json.loads(path.read_text(encoding="utf-8"),
                          object_pairs_hook=reject_duplicates)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise CaptureError(f"cannot read JSON {path}: {error}") from error


def write_once(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(),
            f"refusing to overwrite retained artifact: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False,
                          allow_nan=False) + "\n").encode("utf-8")
    try:
        with path.open("xb") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
        directory = os.open(path.parent, os.O_DIRECTORY)
        try:
            os.fsync(directory)
        finally:
            os.close(directory)
    except FileExistsError as error:
        raise CaptureError(f"refusing to overwrite retained artifact: {path}") from error


def file_descriptor(path: Path, *, require_executable: bool | None = None) -> dict[str, Any]:
    path = path.expanduser()
    require(path.is_absolute(), f"descriptor path must be absolute: {path}")
    require(not path.is_symlink(), f"descriptor path must not be a symlink: {path}")
    require(path.is_file(), f"descriptor path is not a regular file: {path}")
    stat = path.stat()
    executable = bool(stat.st_mode & 0o111)
    if require_executable is not None:
        require(executable is require_executable,
                f"executable flag mismatch for {path}: {executable}")
    return {
        "path": str(path),
        "bytes": stat.st_size,
        "mode": stat.st_mode & 0o777,
        "executable": executable,
        "sha256": sha256_file(path),
    }


def verify_descriptor(value: Any, label: str, *, executable: bool | None = None) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} must be a descriptor object")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and Path(raw_path).is_absolute(),
            f"{label}.path must be absolute")
    actual = file_descriptor(Path(raw_path), require_executable=executable)
    for key in ("bytes", "sha256", "mode", "executable"):
        if key in value:
            require(actual[key] == value[key],
                    f"{label}.{key} does not match {actual['path']}")
    return actual


def descriptor_for_plan(path: Path) -> dict[str, Any]:
    return file_descriptor(path)


def source_entries(manifest_path: Path, label: str) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    raw = read_json(manifest_path)
    require(isinstance(raw, dict) and isinstance(raw.get("files"), list),
            f"{label} does not contain a files list")
    entries: dict[str, dict[str, Any]] = {}
    for item in raw["files"]:
        require(isinstance(item, dict) and isinstance(item.get("path"), str),
                f"{label} contains malformed file entry")
        name = item["path"]
        path = Path(name)
        require(name and not path.is_absolute() and path.as_posix() == name
                and ".." not in path.parts,
                f"{label} contains non-normalized path {name!r}")
        require(name not in entries, f"{label} contains duplicate path {name}")
        require(isinstance(item.get("bytes"), int) and item["bytes"] >= 0,
                f"{label} entry lacks byte count: {name}")
        require(isinstance(item.get("sha256"), str) and len(item["sha256"]) == 64,
                f"{label} entry lacks SHA-256: {name}")
        entries[name] = item
    require(entries, f"{label} is empty")
    return entries, raw


def source_tree_names(root: Path, label: str) -> set[str]:
    require(root.is_dir() and not root.is_symlink(), f"{label} root is not a directory: {root}")
    names: set[str] = set()
    for path in root.rglob("*"):
        require(not path.is_symlink(), f"{label} contains symlink: {path}")
        if path.is_dir():
            continue
        require(path.is_file(), f"{label} contains non-regular file: {path}")
        name = path.relative_to(root).as_posix()
        require(name and Path(name).as_posix() == name and ".." not in Path(name).parts,
                f"{label} path is not normalized: {name}")
        require(name not in names, f"{label} path is duplicated: {name}")
        names.add(name)
    return names


def verify_source_binding(role: str, binding: dict[str, Any]) -> dict[str, Any]:
    root_raw = binding.get("root")
    require(isinstance(root_raw, str) and Path(root_raw).is_absolute(),
            f"{role} source root must be absolute")
    root = Path(root_raw)
    manifest_desc = verify_descriptor(binding.get("manifest"),
                                      f"{role} source manifest")
    manifest_entries, manifest = source_entries(Path(manifest_desc["path"]),
                                                f"{role} source manifest")
    if "source_commit" in binding:
        require(manifest.get("source_commit") == binding["source_commit"],
                f"{role} source commit differs from manifest")
    if "file_count" in binding:
        require(len(manifest_entries) == binding["file_count"],
                f"{role} source file count differs from manifest")
    observed_names = source_tree_names(root, role)
    require(observed_names == set(manifest_entries),
            f"{role} source census differs: missing="
            f"{sorted(set(manifest_entries) - observed_names)[:3]} extra="
            f"{sorted(observed_names - set(manifest_entries))[:3]}")
    for name, wanted in manifest_entries.items():
        actual = file_descriptor(root / name)
        require(actual["bytes"] == wanted["bytes"] and actual["sha256"] == wanted["sha256"],
                f"{role} source differs at {name}")
    return {
        "role": role,
        "root": str(root),
        "manifest": manifest_desc,
        "file_count": len(manifest_entries),
        "checked_files": len(manifest_entries),
        "source_commit": manifest.get("source_commit"),
    }


def verify_build_receipt(role: str, binding: dict[str, Any], binary: dict[str, Any]) -> dict[str, Any]:
    receipt = verify_descriptor(binding.get("build_receipt"), f"{role} build receipt")
    raw = read_json(Path(receipt["path"]))
    require(isinstance(raw, dict), f"{role} build receipt is not an object")
    require(raw.get("status") in {"pass", "terminal", "preserved-before-v6-reuse-bridge"},
            f"{role} build receipt status is not successful: {raw.get('status')!r}")
    if "returncode" in raw:
        require(raw["returncode"] == 0, f"{role} build receipt returncode is not zero")
    if "exit_code" in raw:
        require(raw["exit_code"] == 0, f"{role} build receipt exit_code is not zero")
    if "timed_out" in raw:
        require(raw["timed_out"] is False, f"{role} build receipt timed out")
    if isinstance(raw.get("cleanup"), dict):
        require(raw["cleanup"].get("process_group_clear", True) is True,
                f"{role} build process group was not cleared")
    if "terminal_before_binary_parse" in raw:
        require(raw["terminal_before_binary_parse"] is True,
                f"{role} build terminal was not durable before binary parse")
    if role == "candidate":
        require(str(raw.get("schema", "")).startswith("litchi.performance.0948.candidate-"),
                f"candidate build receipt schema is not 0948 candidate build")
    else:
        require(raw.get("schema") == "litchi.performance.0948.baseline-reuse.v2",
                f"baseline receipt schema changed: {raw.get('schema')!r}")
    return receipt


def verify_binary_binding(role: str, binding: dict[str, Any]) -> dict[str, Any]:
    binary = verify_descriptor(binding.get("binary"), f"{role} allocator binary", executable=True)
    build = verify_build_receipt(role, binding, binary)
    require(binary["bytes"] == binding["binary"].get("bytes"),
            f"{role} binary descriptor missing exact size")
    require(binary["sha256"] == binding["binary"].get("sha256"),
            f"{role} binary descriptor missing exact SHA-256")
    return {"binary": binary, "build_receipt": build}


def verify_semantic_context(plan: dict[str, Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for index, value in enumerate(plan.get("semantic_context", [])):
        result.append(verify_descriptor(value, f"semantic context[{index}]"))
    require(result, "semantic_context must bind at least one 0948 identity")
    return result


def verify_tools(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    tools = plan.get("tools")
    require(isinstance(tools, dict), "plan.tools is missing")
    result: dict[str, dict[str, Any]] = {}
    for name in ("heaptrack", "heaptrack_print", "taskset", "time"):
        value = tools.get(name)
        actual = verify_descriptor(value, f"tool {name}", executable=True)
        require(value.get("version") == plan["tools_versions"].get(name),
                f"tool {name} version binding changed")
        result[name] = actual
    return result


def jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    matrix = plan.get("matrix")
    require(isinstance(matrix, dict), "plan.matrix is missing")
    expected = [
        ("h00-baseline-primary-percent-medium", "baseline", "medium"),
        ("h01-candidate-primary-percent-medium", "candidate", "medium"),
        ("h02-baseline-primary-percent-dense-sparse", "baseline", "dense-sparse"),
        ("h03-candidate-primary-percent-dense-sparse", "candidate", "dense-sparse"),
    ]
    configured = matrix.get("jobs")
    require(isinstance(configured, list) and len(configured) == len(expected),
            "matrix.jobs must contain exactly four rows")
    result: list[dict[str, Any]] = []
    for item, (name, role, shape) in zip(configured, expected):
        require(isinstance(item, dict), "matrix job must be an object")
        require(item.get("name") == name and item.get("role") == role
                and item.get("shape") == shape,
                f"matrix schedule changed for {name}")
        require(item.get("case") == "xlsx_source_backed_cell_values_one_percent_edit_save",
                f"matrix case changed for {name}")
        result.append(dict(item))
    return result


def load_and_validate_plan(*, verify_sources: bool = True) -> tuple[dict[str, Any], dict[str, Any]]:
    require(PLAN_PATH.is_file(), f"missing plan: {PLAN_PATH}")
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict) and plan.get("schema") == SCHEMA,
            f"unexpected plan schema: {plan.get('schema') if isinstance(plan, dict) else None!r}")
    require(plan.get("plan_id") == "0949" and plan.get("protocol_revision") == 1
            and plan.get("plan_revision") == "v3",
            "plan identity/protocol revision changed")
    require(plan.get("status") in {"prospective_no_capture", "capture_ready"},
            f"plan is not a prospective capture plan: {plan.get('status')!r}")
    require(plan.get("capture_authorization") == "root_GO_required",
            "plan does not require explicit root GO")
    matrix = plan.get("matrix")
    require(isinstance(matrix, dict), "plan.matrix missing")
    require(matrix.get("children") == 4 and matrix.get("serial") is True
            and matrix.get("warmup") == 0 and matrix.get("samples") == 1
            and matrix.get("cpu") == 12 and matrix.get("deadline_seconds") == 600
            and matrix.get("no_retry") is True, "frozen diagnostic matrix changed")
    require(matrix.get("timezone") == "UTC" and matrix.get("locale") == "C",
            "frozen diagnostic locale changed")
    expected_env = plan.get("environment", {}).get("fixed")
    require(expected_env == FIXED_ENVIRONMENT, "fixed runtime environment changed")
    configured_tools = verify_tools(plan)
    runner_desc = verify_descriptor(plan.get("runner"), "runner helper")
    semantic = verify_semantic_context(plan)
    binary_results = {
        role: verify_binary_binding(role, plan["binaries"][role])
        for role in ("baseline", "candidate")
    }
    source_results: dict[str, Any] = {}
    for role in ("baseline", "candidate"):
        source_results[role] = verify_source_binding(role, plan["sources"][role]) \
            if verify_sources else {"role": role, "verified": False}
    job_list = jobs(plan)
    capture_root = plan.get("capture", {}).get("root")
    require(isinstance(capture_root, str) and Path(capture_root).is_absolute(),
            "capture.root must be absolute")
    validation = {
        "schema": SCHEMA,
        "plan": descriptor_for_plan(PLAN_PATH),
        "runner": runner_desc,
        "tools": configured_tools,
        "binaries": binary_results,
        "sources": source_results,
        "semantic_context": semantic,
        "jobs": [item["name"] for item in job_list],
        "capture_root": capture_root,
        "workload_not_started": True,
    }
    return plan, validation


def group_is_clear(pgid: int | None) -> bool:
    if pgid is None:
        return True
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return True
    except PermissionError:
        return False
    return False


def wait_group_clear(pgid: int | None, seconds: float) -> bool:
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if group_is_clear(pgid):
            return True
        time.sleep(0.05)
    return group_is_clear(pgid)


def child_artifact_paths(folder: Path) -> dict[str, str]:
    # heaptrack may append its compression suffix (for example .zst) to the
    # requested -o name.  Start receipts bind the requested prefix; terminal
    # receipts bind the observed concrete file after reaping.
    return {
        "stdout": "stdout.log",
        "stderr": "stderr.log",
        "time": "time.txt",
        "report": "report.json",
        "profile_requested": "heaptrack.data.gz",
        "profile_prefix": "heaptrack.data.gz",
        "start": "start.json",
        "terminal": "terminal.json",
    }


def observed_profile_candidates(folder: Path) -> list[str]:
    prefix = "heaptrack.data.gz"
    candidates: list[str] = []
    for path in sorted(folder.iterdir(), key=lambda item: item.name):
        if path.name in (prefix, prefix + ".zst") and path.is_file() and not path.is_symlink():
            candidates.append(path.name)
    return candidates


def child_command(job: dict[str, Any], plan: dict[str, Any], folder: Path) -> list[str]:
    role = job["role"]
    binary = plan["binaries"][role]["binary"]["path"]
    cpu = str(plan["matrix"]["cpu"])
    return [
        plan["tools"]["time"]["path"], "-v", "-o", str(folder / "time.txt"),
        plan["tools"]["taskset"]["path"], "--cpu-list", cpu,
        plan["tools"]["heaptrack"]["path"], "--record-only", "-o",
        str(folder / "heaptrack.data.gz"), binary,
        "--warmup", str(plan["matrix"]["warmup"]),
        "--samples", str(plan["matrix"]["samples"]),
        "--case", job["case"],
        "--xlsx-cell-crud-shape", job["shape"],
        "--json", str(folder / "report.json"),
    ]


def child_terminal(plan: dict[str, Any], validation: dict[str, Any], job: dict[str, Any],
                   folder: Path, command: list[str], started_utc: str,
                   monotonic_start: float, pid: int | None, pgid: int | None,
                   returncode: int | None, timed_out: bool, launch_error: str | None,
                   group_clear: bool, profile_candidates: list[str]) -> dict[str, Any]:
    source_desc = validation["sources"][job["role"]]
    binary_desc = validation["binaries"][job["role"]]["binary"]
    profile_match = "exactly_one" if len(profile_candidates) == 1 else (
        "missing" if not profile_candidates else "multiple")
    return {
        "schema": CHILD_SCHEMA,
        "status": "pass" if returncode == 0 and not timed_out and not launch_error
        and group_clear and len(profile_candidates) == 1 else "fail",
        "plan": validation["plan"],
        "runner": validation["runner"],
        "id": job["name"],
        "role": job["role"],
        "case": job["case"],
        "xlsx_cell_crud_shape": job["shape"],
        "sample_index": 0,
        "warmup": 0,
        "argv": command,
        "cwd": source_desc["root"],
        "environment": FIXED_ENVIRONMENT,
        "environment_sha256": sha256_bytes(canonical_json(FIXED_ENVIRONMENT)),
        "started_utc": started_utc,
        "terminal_utc": utc_now(),
        "wall_seconds": time.monotonic() - monotonic_start,
        "deadline_seconds": plan["matrix"]["deadline_seconds"],
        "pid": pid,
        "process_group_id": pgid,
        "start_new_session": True,
        "returncode": returncode,
        "timed_out": timed_out,
        "launch_error": launch_error,
        "process_group_clear": group_clear,
        "binary": binary_desc,
        "source": source_desc,
        "artifacts": {
            **child_artifact_paths(folder),
            "profile_candidates": profile_candidates,
            "profile_match": profile_match,
            "profile_actual_name": profile_candidates[0] if len(profile_candidates) == 1 else None,
            "profile_actual_path": str(folder / profile_candidates[0]) if len(profile_candidates) == 1 else None,
        },
        "terminal_before_any_report_or_profile_parse": True,
        "analysis_started": False,
        "raw_profile_interpretation": "deferred; whole-process heaptrack interval only",
    }


def run_child(plan: dict[str, Any], validation: dict[str, Any], job: dict[str, Any],
              capture_root: Path) -> dict[str, Any]:
    folder = capture_root / "children" / job["name"]
    require(not folder.exists() and not folder.is_symlink(),
            f"refusing retained child folder (no retry): {folder}")
    folder.mkdir(parents=True)
    command = child_command(job, plan, folder)
    started_utc = utc_now()
    monotonic_start = time.monotonic()
    start_receipt = {
        "schema": CHILD_SCHEMA,
        "status": "started",
        "plan": validation["plan"],
        "runner": validation["runner"],
        "id": job["name"],
        "role": job["role"],
        "case": job["case"],
        "xlsx_cell_crud_shape": job["shape"],
        "sample_index": 0,
        "warmup": 0,
        "argv": command,
        "cwd": validation["sources"][job["role"]]["root"],
        "environment": FIXED_ENVIRONMENT,
        "environment_sha256": sha256_bytes(canonical_json(FIXED_ENVIRONMENT)),
        "started_utc": started_utc,
        "deadline_seconds": plan["matrix"]["deadline_seconds"],
        "start_new_session": True,
        "process_group": True,
        "binary": validation["binaries"][job["role"]]["binary"],
        "source": validation["sources"][job["role"]],
        "artifacts": child_artifact_paths(folder),
        "no_retry": True,
        "terminal_receipt": "terminal.json",
    }
    write_once(folder / "start.json", start_receipt)
    pid: int | None = None
    pgid: int | None = None
    returncode: int | None = None
    timed_out = False
    launch_error: str | None = None
    try:
        with (folder / "stdout.log").open("xb") as stdout, (folder / "stderr.log").open("xb") as stderr:
            try:
                process = subprocess.Popen(
                    command,
                    cwd=validation["sources"][job["role"]]["root"],
                    env=dict(FIXED_ENVIRONMENT),
                    stdin=subprocess.DEVNULL,
                    stdout=stdout,
                    stderr=stderr,
                    start_new_session=True,
                )
                pid = process.pid
                pgid = pid
                try:
                    returncode = process.wait(timeout=plan["matrix"]["deadline_seconds"])
                except subprocess.TimeoutExpired:
                    timed_out = True
                    try:
                        os.killpg(pgid, signal.SIGTERM)
                    except ProcessLookupError:
                        pass
                    try:
                        returncode = process.wait(timeout=15)
                    except subprocess.TimeoutExpired:
                        try:
                            os.killpg(pgid, signal.SIGKILL)
                        except ProcessLookupError:
                            pass
                        returncode = process.wait(timeout=15)
            except Exception as error:  # receipt records launch/IO failures; no retry
                launch_error = f"{type(error).__name__}: {error}"
    except Exception as error:
        launch_error = launch_error or f"{type(error).__name__}: {error}"
    group_clear = wait_group_clear(pgid, 30.0)
    if not group_clear and pgid is not None:
        try:
            os.killpg(pgid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        group_clear = wait_group_clear(pgid, 5.0)
    # Listing the concrete profile name is custody metadata, not a profile
    # parse.  heaptrack's terminal receipt is written only after this listing.
    profile_candidates = observed_profile_candidates(folder)
    terminal = child_terminal(
        plan, validation, job, folder, command, started_utc, monotonic_start,
        pid, pgid, returncode, timed_out, launch_error, group_clear,
        profile_candidates,
    )
    # This is the first operation that may inspect neither report.json nor
    # heaptrack.data.gz; both are deliberately left for an independent reducer.
    write_once(folder / "terminal.json", terminal)
    return terminal


def run_capture(plan: dict[str, Any], validation: dict[str, Any]) -> int:
    capture_root = Path(plan["capture"]["root"])
    require(not capture_root.exists() and not capture_root.is_symlink(),
            f"refusing existing capture root (no retry): {capture_root}")
    capture_root.mkdir(parents=True)
    (capture_root / "children").mkdir()
    started = utc_now()
    monotonic_start = time.monotonic()
    job_list = jobs(plan)
    runner_start = {
        "schema": RUNNER_SCHEMA,
        "status": "started",
        "plan": validation["plan"],
        "runner": validation["runner"],
        "capture_root": str(capture_root),
        "started_utc": started,
        "matrix": plan["matrix"],
        "children": [item["name"] for item in job_list],
        "serial": True,
        "no_retry": True,
        "terminal_before_any_report_or_profile_parse": True,
        "analysis_started": False,
    }
    write_once(capture_root / "runner-start.json", runner_start)
    terminals: list[dict[str, Any]] = []
    runner_error: str | None = None
    try:
        for job in job_list:
            terminals.append(run_child(plan, validation, job, capture_root))
    except BaseException as error:
        runner_error = f"{type(error).__name__}: {error}"
    status = "pass" if runner_error is None and len(terminals) == len(job_list) \
        and all(item["status"] == "pass" for item in terminals) else "fail"
    outer = {
        "schema": RUNNER_SCHEMA,
        "status": "terminal",
        "capture_status": status,
        "plan": validation["plan"],
        "runner": validation["runner"],
        "capture_root": str(capture_root),
        "started_utc": started,
        "terminal_utc": utc_now(),
        "wall_seconds": time.monotonic() - monotonic_start,
        "children": [item["id"] for item in terminals],
        "expected_children": [item["name"] for item in job_list],
        "child_terminals_written": len(terminals),
        "runner_error": runner_error,
        "failures": [item["id"] for item in terminals if item["status"] != "pass"],
        "terminal_before_any_report_or_profile_parse": True,
        "analysis_started": False,
        "no_retry": True,
    }
    write_once(capture_root / "runner-terminal.json", outer)
    print(json.dumps({"capture_root": str(capture_root), "status": status,
                      "children": len(terminals), "runner_error": runner_error},
                     sort_keys=True))
    return 0 if status == "pass" else 1


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--validate-only", action="store_true",
                      help="validate all immutable bindings; launch nothing")
    mode.add_argument("--capture", action="store_true",
                      help="run the four-child schedule; requires --go")
    parser.add_argument("--go", action="store_true",
                        help="explicit root coordination authorization for capture")
    args = parser.parse_args(argv)
    try:
        plan, validation = load_and_validate_plan(verify_sources=True)
        if args.validate_only:
            require(not args.go, "--go is only valid with --capture")
            print(json.dumps({"status": "validated", "workload_started": False,
                              "validation": validation}, sort_keys=True))
            return 0
        require(args.go, "capture is inert until explicit --go")
        return run_capture(plan, validation)
    except (CaptureError, OSError) as error:
        print(f"0949 capture refused: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
