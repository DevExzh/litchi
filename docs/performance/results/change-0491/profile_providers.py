#!/usr/bin/env python3
"""Capture bounded whole-child provider diagnostics for the retained build.

This helper is intentionally separate from the provider matrix.  It runs one
``perf stat`` child for each selected provider arm and one ``strace -c`` child
for the file arm.  The Rust report remains the source of truth for the
30-sample/3-warmup operation; profiler counters and ``/usr/bin/time`` are
whole-child observations that include setup, corpus access, report writing,
and diagnostic work.  No counter is attributed to an individual operation and
the output makes no speedup claim.

The coordinator must invoke this script through ``gate.py``.  The gate owns
the shared CPU lock; this helper pins each benchmark child to CPU 2 and never
acquires the lock itself.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import math
import os
from pathlib import Path
import re
import signal
import shutil
import statistics
import subprocess
import sys
from typing import Any, Iterable, Mapping

from support import ENV, ENV_KEYS, REPO, ROOT, TEMP, meta, now, read, sha, write

import provider_matrix as matrix


SCHEMA = "docx-provider-whole-child-profile-v1"
TERMINAL_SCHEMA = "docx-provider-whole-child-profile-terminal-v1"
RESULT_SCHEMA = "docx-provider-whole-child-profile-result-v1"
VERSION = 1
CPU = 2
CPU_LOCK = "/home/zhuhe/.cache/litchi-goal-0484/cpu.lock"
SAMPLES = 30
WARMUPS = 3
DEFAULT_TIMEOUT_SECONDS = 1_800
TERM_GRACE_SECONDS = 10
ATTEMPT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")

PROFILE_ROOT = ROOT / "profile-providers"
PROFILE_GATE_PATH = ROOT / "validation" / "profiles2.json"
GATE_SCHEMA = "docx-stream-append-gate-v1"
ARM_NAMES = (
    "file",
    "range-64-0us",
    "range-65536-1000us-104857600bps-minimum-service",
)
RUNS = tuple(("perf", arm) for arm in ARM_NAMES) + (("strace", "file"),)
PERF_EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "cache-references",
    "cache-misses",
    "page-faults",
    "duration_time",
)
STRACE_SYSCALLS = ("read", "pread64", "statx", "write", "fsync", "fdatasync")
HELPER_FILES = (
    "profile_providers.py",
    "provider_matrix.py",
    "support.py",
    "retain_build.py",
    "gate.py",
)
WHOLE_CHILD_SCOPE = (
    "the /usr/bin/time/perf-or-strace child includes provider setup, corpus/file "
    "access, Package/document extraction, report serialization, and diagnostic "
    "work; operation rows retain the provider lifecycle timer scope separately"
)
NO_ATTRIBUTION_CLAIM = (
    "whole-child profiler counters and resource values are descriptive; they are "
    "not attributed to an individual operation and do not establish a speedup"
)
CALLER_GATE = {
    "required": True,
    "driver": "gate.py",
    "lock": CPU_LOCK,
    "ownership": "outer caller holds the shared lock for the complete command",
}

PROFILE_ARTIFACT_NAMES = (
    "stdout.txt",
    "stderr.txt",
    "resource.txt",
    "report.json",
    "replay-cleanup.json",
)
STARTED_KEYS = frozenset(
    {
        "schema",
        "version",
        "status",
        "attempt",
        "label",
        "tool",
        "tool_executable",
        "arm",
        "run",
        "samples",
        "warmups",
        "scope",
        "operation_attribution_claim",
        "performance_claim",
        "caller_gate",
        "cpu",
        "protocol",
        "build",
        "binary",
        "source",
        "source_revision",
        "helpers",
        "machine",
        "argv",
        "cwd",
        "environment",
        "started_utc",
        "timeout_seconds",
        "tmpdir",
    }
)
TERMINAL_KEYS = STARTED_KEYS | frozenset(
    {
        "exit_code",
        "timed_out",
        "termination",
        "finished_utc",
        "artifacts",
        "missing_artifacts",
        "cleanup",
        "started_artifact",
        "profiler",
        "report_summary",
        "resource",
    }
)
RECORD_KEYS = frozenset(
    {
        "label",
        "tool",
        "arm",
        "status",
        "terminal",
        "report",
        "resource",
        "profiler",
        "tool_executable",
        "scope",
        "operation_attribution_claim",
    }
)
RESULT_KEYS = frozenset(
    {
        "schema",
        "version",
        "status",
        "attempt",
        "runs",
        "samples",
        "warmups",
        "scope",
        "operation_attribution_claim",
        "performance_claim",
        "caller_gate",
        "cpu",
        "protocol",
        "build",
        "source_revision",
        "helpers",
        "records",
        "finished_utc",
    }
)
CLEANUP_KEYS = frozenset(
    {"schema", "status", "run_root", "tmpdir", "removed", "remaining"}
)


class ProfileError(RuntimeError):
    """A fail-closed profile custody or parsing error."""


def fail(message: str) -> None:
    raise ProfileError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _attempt(value: str) -> str:
    require(ATTEMPT_RE.fullmatch(value) is not None, "attempt must be a path-safe token")
    return value


def _finite(value: Any, path: str = "json") -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            require(isinstance(key, str), f"{path}: JSON key is not a string")
            _finite(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{path}[{index}]")


def _read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing regular JSON file: {path}")
    try:
        value = read(path)
    except (OSError, ValueError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    _finite(value, str(path))
    return value


def _artifact(path: Path, *, executable: bool = False) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing regular artifact: {path}")
    if executable:
        require(os.access(path, os.X_OK), f"{path}: executable bit is absent")
    value = meta(path)
    require(type(value["bytes"]) is int and value["bytes"] >= 0, f"{path}: invalid byte count")
    require(SHA256_RE.fullmatch(value["sha256"]) is not None, f"{path}: invalid SHA-256")
    return {"path": str(path), **value, **({"executable": True} if executable else {})}


def _tool_executable(tool: str) -> dict[str, Any]:
    """Resolve and authenticate the executable used by a profiler child."""

    require(tool in {"perf", "strace"}, f"unknown profiler tool: {tool}")
    located = shutil.which(tool)
    require(isinstance(located, str) and located, f"cannot resolve profiler executable: {tool}")
    try:
        path = Path(located).resolve(strict=True)
    except OSError as error:
        fail(f"cannot resolve profiler executable {tool}: {error}")
    return _artifact(path, executable=True)


def _expected_environment(tmp_root: Path) -> dict[str, str]:
    environment = {key: ENV[key] for key in ENV_KEYS}
    environment["TMPDIR"] = str(tmp_root)
    return environment


def _profile_artifacts(directory: Path, tool: str) -> dict[str, dict[str, Any]]:
    profiler_name = "perf.txt" if tool == "perf" else "strace.txt"
    names = (*PROFILE_ARTIFACT_NAMES[:4], profiler_name, PROFILE_ARTIFACT_NAMES[4])
    return {name: _artifact(directory / name) for name in names}


def _timestamp(value: Any, path: str) -> _datetime.datetime:
    require(isinstance(value, str) and value, f"{path}: timestamp is missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{path}: invalid timestamp: {error}")
    require(parsed.tzinfo is not None, f"{path}: timestamp lacks timezone")
    return parsed


def _helper_hashes() -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for name in HELPER_FILES:
        path = ROOT / name
        result[name] = _artifact(path)
    return result


def _load_build() -> tuple[Path, dict[str, Any]]:
    path = ROOT / "build-normal.json"
    value = _read_json(path)
    try:
        build = matrix._binary_from_build(value, "normal", path)
    except (matrix.ProviderMatrixError, KeyError, TypeError, ValueError) as error:
        fail(f"{path}: retained build validation failed: {error}")
    require(build["binary"]["executable"] is True, f"{path}: retained binary is not executable")
    return path, build


def _load_protocol() -> tuple[dict[str, Any], str]:
    try:
        protocol, protocol_hash = matrix._load_protocol()
    except (matrix.ProviderMatrixError, KeyError, TypeError, ValueError) as error:
        fail(f"provider protocol validation failed: {error}")
    require(protocol.get("machine") is not None, "provider protocol machine binding is missing")
    return protocol, protocol_hash


def _source_revision() -> str:
    try:
        revision = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=REPO, text=True, stderr=subprocess.PIPE
        ).strip()
    except (OSError, subprocess.SubprocessError) as error:
        fail(f"cannot determine source revision: {error}")
    require(REVISION_RE.fullmatch(revision) is not None, "source revision is malformed")
    return revision


def _spec(arm_name: str) -> dict[str, Any]:
    arm = matrix.ARM_BY_NAME.get(arm_name)
    require(isinstance(arm, dict), f"unknown provider arm: {arm_name}")
    return {
        "kind": "profile",
        "repeat": 1,
        "role": "normal",
        "arm": arm_name,
        "provider": arm["provider"],
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "label": f"normal-{arm_name}",
    }


def _benchmark_argv(spec: Mapping[str, Any], binary: Mapping[str, Any], report: Path,
                    resource: Path, source_revision: str, profiler: Path, tool: str,
                    tool_executable: str | Path | None = None) -> list[str]:
    base = matrix._command(dict(spec), dict(binary), report, resource, source_revision)
    executable = str(tool_executable) if tool_executable is not None else _tool_executable(tool)["path"]
    if tool == "perf":
        return [
            executable, "stat", "-x,", "--no-big-num", "-o", str(profiler),
            "-e", ",".join(PERF_EVENTS), "--", *base,
        ]
    require(tool == "strace", f"unknown profiler tool: {tool}")
    return [
        executable, "-f", "-c", "-e", "trace=" + ",".join(STRACE_SYSCALLS),
        "-o", str(profiler), "--", *base,
    ]


def _kill_group(process: subprocess.Popen[bytes]) -> str | None:
    if process.poll() is not None:
        return None
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return None
    try:
        process.communicate(timeout=TERM_GRACE_SECONDS)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        process.communicate()
        return "SIGKILL"
    return "SIGTERM"


def _cleanup_private(run_root: Path, tmp_root: Path) -> dict[str, Any]:
    require(tmp_root.parent == run_root, "private TMPDIR escaped its run root")
    remaining: list[str] = []
    removed: list[str] = []
    if tmp_root.exists():
        require(tmp_root.is_dir() and not tmp_root.is_symlink(), f"private TMPDIR is not a directory: {tmp_root}")
        children = list(tmp_root.iterdir())
        if children:
            remaining.extend(str(child) for child in children)
        else:
            tmp_root.rmdir()
            removed.append(str(tmp_root))
    else:
        remaining.append(str(tmp_root))
    if not remaining and run_root.exists():
        require(run_root.is_dir() and not run_root.is_symlink(), f"private run root is not a directory: {run_root}")
        if list(run_root.iterdir()):
            remaining.extend(str(child) for child in run_root.iterdir())
        else:
            run_root.rmdir()
            removed.append(str(run_root))
    return {
        "schema": "docx-provider-profile-private-cleanup-v1",
        "status": "pass" if not remaining else "failed",
        "run_root": str(run_root),
        "tmpdir": str(tmp_root),
        "removed": removed,
        "remaining": remaining,
    }


def _validate_cleanup(cleanup: Any, run_root: Path, tmp_root: Path, label: str) -> dict[str, Any]:
    require(isinstance(cleanup, dict), f"{label}: cleanup receipt is not an object")
    require(set(cleanup) == CLEANUP_KEYS, f"{label}: cleanup receipt fields differ")
    require(cleanup["schema"] == "docx-provider-profile-private-cleanup-v1",
            f"{label}: cleanup schema differs")
    require(cleanup["status"] == "pass", f"{label}: private cleanup did not pass")
    require(cleanup["run_root"] == str(run_root) and cleanup["tmpdir"] == str(tmp_root),
            f"{label}: private cleanup path differs")
    require(isinstance(cleanup["removed"], list) and
            all(isinstance(item, str) and item for item in cleanup["removed"]),
            f"{label}: cleanup removed inventory is malformed")
    require(cleanup["remaining"] == [], f"{label}: private scratch remains")
    require(str(tmp_root) in cleanup["removed"] and str(run_root) in cleanup["removed"],
            f"{label}: cleanup did not remove both private roots")
    return dict(cleanup)


def _validate_profile_gate(build: Mapping[str, Any]) -> dict[str, Any]:
    """Validate the outer gate after profile capture has produced its terminal."""

    path = PROFILE_GATE_PATH
    require(path.is_file() and not path.is_symlink(), f"missing profile gate receipt: {path}")
    try:
        receipt = read(path)
    except (OSError, ValueError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid gate receipt: {error}")
    _finite(receipt, str(path))
    require(isinstance(receipt, dict), f"{path}: gate receipt is not an object")
    required = {
        "argv", "artifacts", "attempt", "common_sha256", "cwd", "driver_sha256",
        "environment", "exit_code", "finished_utc", "label", "schema",
        "source_after", "source_before", "source_unchanged", "started_utc",
    }
    require(set(receipt) == required, f"{path}: gate receipt fields differ")
    require(receipt["schema"] == GATE_SCHEMA and receipt["label"] == "profiles2",
            f"{path}: profile gate identity differs")
    require(receipt["attempt"] is None and receipt["exit_code"] == 0 and
            receipt["source_unchanged"] is True,
            f"{path}: profile gate did not pass with unchanged source")
    expected_argv = ["python3", "-B", str(ROOT / "profile_providers.py"), "profiles2"]
    require(receipt["argv"] == expected_argv,
            f"{path}: profile gate argv differs from the canonical post-capture command")
    require(receipt["cwd"] == str(REPO), f"{path}: profile gate cwd differs")
    require(receipt["environment"] == {key: ENV[key] for key in ENV_KEYS},
            f"{path}: profile gate environment differs")
    require(receipt["driver_sha256"] == sha(ROOT / "gate.py"),
            f"{path}: gate driver binding differs")
    require(receipt["common_sha256"] == sha(ROOT / "support.py"),
            f"{path}: gate support binding differs")
    started_at = _timestamp(receipt["started_utc"], f"{path}.started_utc")
    finished_at = _timestamp(receipt["finished_utc"], f"{path}.finished_utc")
    require(finished_at > started_at, f"{path}: gate chronology is invalid")
    try:
        source_before = matrix._source_binding(receipt["source_before"], path, "source_before")
        source_after = matrix._source_binding(receipt["source_after"], path, "source_after")
    except (matrix.ProviderMatrixError, KeyError, TypeError, ValueError) as error:
        fail(f"{path}: gate source binding is invalid: {error}")
    require(source_before == source_after == build["source"],
            f"{path}: profile gate source differs from retained build")
    expected_artifacts = {"profiles2.stdout", "profiles2.stderr"}
    require(set(receipt["artifacts"]) == expected_artifacts,
            f"{path}: profile gate artifact inventory differs")
    artifacts: dict[str, dict[str, Any]] = {}
    for name in sorted(expected_artifacts):
        value = receipt["artifacts"][name]
        require(isinstance(value, dict) and set(value) == {"bytes", "sha256"},
                f"{path}: malformed profile gate artifact {name}")
        artifact_path = path.with_suffix("." + name.rsplit(".", 1)[1])
        actual = _artifact(artifact_path)
        require(value == {"bytes": actual["bytes"], "sha256": actual["sha256"]},
                f"{path}: profile gate artifact {name} changed")
        artifacts[name] = dict(value)
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
        "schema": GATE_SCHEMA,
        "label": "profiles2",
        "attempt": None,
        "source": source_after,
        "argv": list(receipt["argv"]),
        "artifacts": artifacts,
    }


def _numeric(token: str) -> int | float | None:
    token = token.strip()
    if not token or token.startswith("<"):
        return None
    try:
        number = float(token)
    except ValueError:
        return None
    return int(number) if number.is_integer() else number


def _parse_perf(path: Path) -> dict[str, Any]:
    """Parse one raw ``perf stat -x,`` file with unsupported events explicit."""

    require(path.is_file(), f"missing perf output: {path}")
    raw = path.read_bytes()
    require(raw, f"empty perf output: {path}")
    events: dict[str, dict[str, Any]] = {
        event: {"status": "missing", "value": None} for event in PERF_EVENTS
    }
    for line_number, line in enumerate(raw.decode("utf-8", errors="replace").splitlines(), start=1):
        text = line.strip()
        if not text or text.startswith("#"):
            continue
        fields = [field.strip() for field in text.split(",")]
        event = next((field for field in fields if field in PERF_EVENTS), None)
        if event is None:
            continue
        value_token = fields[0] if fields else ""
        lowered = value_token.lower()
        if lowered in {"<not supported>", "<not counted>"}:
            status = "unsupported" if lowered == "<not supported>" else "not_counted"
            value = None
        else:
            value = _numeric(value_token)
            if value is None:
                fail(f"{path}:{line_number}: perf event {event} has an invalid value")
            require(value >= 0, f"{path}:{line_number}: perf event {event} is negative")
            status = "observed"
        require(events[event]["status"] == "missing", f"{path}: duplicate perf event {event}")
        events[event] = {"status": status, "value": value, "csv_fields": fields}
    require(all(row["status"] != "missing" for row in events.values()), f"{path}: missing perf events")
    return {
        "schema": "docx-provider-profile-perf-v1",
        "events": events,
        "raw_bytes": len(raw),
        "raw_sha256": hashlib.sha256(raw).hexdigest(),
    }


def _parse_strace(path: Path) -> dict[str, Any]:
    """Parse selected syscall call/error counts, preserving missing rows."""

    require(path.is_file(), f"missing strace output: {path}")
    raw = path.read_bytes()
    require(raw, f"empty strace output: {path}")
    decoded = raw.decode("utf-8", errors="replace")
    require("syscall" in decoded and "calls" in decoded, f"{path}: missing strace header")
    require(any(line.split() and line.split()[-1] == "total" for line in decoded.splitlines()), f"{path}: missing strace total")
    selected: dict[str, dict[str, Any]] = {
        syscall: {"status": "missing", "calls": 0, "errors": 0}
        for syscall in STRACE_SYSCALLS
    }
    for line_number, line in enumerate(raw.decode("utf-8", errors="replace").splitlines(), start=1):
        fields = line.split()
        if not fields or fields[-1] not in selected:
            continue
        syscall = fields[-1]
        if len(fields) >= 6:
            calls_token, errors_token = fields[-3], fields[-2]
        elif len(fields) == 5:
            calls_token, errors_token = fields[-2], "0"
        else:
            fail(f"{path}:{line_number}: malformed strace row for {syscall}")
        try:
            calls = int(calls_token)
            errors = int(errors_token)
        except ValueError:
            fail(f"{path}:{line_number}: malformed counts for {syscall}")
        require(calls >= 0 and errors >= 0, f"{path}:{line_number}: negative syscall count")
        require(errors <= calls, f"{path}:{line_number}: errors exceed calls")
        require(selected[syscall]["status"] == "missing", f"{path}: duplicate syscall {syscall}")
        selected[syscall] = {"status": "observed", "calls": calls, "errors": errors}
    require(any(row["status"] == "observed" for row in selected.values()), f"{path}: no syscall observations")
    return {
        "schema": "docx-provider-profile-strace-v1",
        "syscalls": selected,
        "selected": list(STRACE_SYSCALLS),
        "raw_bytes": len(raw),
        "raw_sha256": hashlib.sha256(raw).hexdigest(),
    }


def _percentiles(values: Iterable[int | float]) -> dict[str, Any]:
    ordered = sorted(values)
    require(ordered, "cannot summarize an empty vector")
    p50 = statistics.median(ordered)
    p95 = ordered[max(0, math.ceil(len(ordered) * 0.95) - 1)]
    p99 = ordered[max(0, math.ceil(len(ordered) * 0.99) - 1)]
    return {
        "n": len(ordered),
        "min": ordered[0],
        "max": ordered[-1],
        "mean": statistics.fmean(ordered),
        "p50": p50,
        "p95": p95,
        "p99": p99,
    }


def _report_summary(report: Mapping[str, Any]) -> dict[str, Any]:
    rows = report.get("rows")
    require(isinstance(rows, list) and rows, "report rows are missing")
    latency = [row["latency_ns"] for row in rows]
    wrapper = [row["reads"]["wrapper"] for row in rows]
    return {
        "samples": len(rows),
        "latency_ns": _percentiles(latency),
        "logical_read_calls": _percentiles(item["logical_calls"] for item in wrapper),
        "logical_read_requested_bytes": _percentiles(item["requested_bytes"] for item in wrapper),
        "logical_read_returned_bytes": _percentiles(item["returned_bytes"] for item in wrapper),
        "logical_read_short_reads": _percentiles(item["short_reads"] for item in wrapper),
    }


def _profile_one(attempt: str, tool: str, arm_name: str, build: Mapping[str, Any],
                 build_path: Path, protocol_hash: str, source_revision: str,
                 helper_hashes: Mapping[str, Any], machine: Any,
                 timeout_seconds: int) -> dict[str, Any]:
    spec = _spec(arm_name)
    label = f"{tool}-{arm_name}"
    tool_executable = _tool_executable(tool)
    attempt_root = PROFILE_ROOT / attempt
    directory = attempt_root / label
    require(not directory.exists(), f"refusing to replace profile output: {directory}")
    directory.mkdir(parents=True, exist_ok=False)
    run_root = TEMP / "provider-profiles" / attempt / label
    require(not run_root.exists(), f"refusing to replace private profile scratch: {run_root}")
    run_root.mkdir(parents=True, exist_ok=False)
    tmp_root = run_root / "tmp"
    tmp_root.mkdir(exist_ok=False)
    report = directory / "report.json"
    resource = directory / "resource.txt"
    stdout = directory / "stdout.txt"
    stderr = directory / "stderr.txt"
    profiler = directory / ("perf.txt" if tool == "perf" else "strace.txt")
    binary = dict(build["binary"])
    argv = _benchmark_argv(
        spec, binary, report, resource, source_revision, profiler, tool,
        tool_executable["path"],
    )
    environment = _expected_environment(tmp_root)
    started = {
        "schema": SCHEMA,
        "version": VERSION,
        "status": "running",
        "attempt": attempt,
        "label": label,
        "tool": tool,
        "tool_executable": tool_executable,
        "arm": dict(matrix.ARM_BY_NAME[arm_name]),
        "run": dict(spec),
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "scope": WHOLE_CHILD_SCOPE,
        "operation_attribution_claim": False,
        "performance_claim": NO_ATTRIBUTION_CLAIM,
        "caller_gate": dict(CALLER_GATE),
        "cpu": CPU,
        "protocol": {"path": "protocol.json", "sha256": protocol_hash},
        "build": {"path": str(build_path), "sha256": sha(build_path), "source": build["source"]},
        "binary": binary,
        "source": build["source"],
        "source_revision": source_revision,
        "helpers": dict(helper_hashes),
        "machine": machine,
        "argv": argv,
        "cwd": str(REPO),
        "environment": environment,
        "started_utc": now(),
        "timeout_seconds": timeout_seconds,
        "tmpdir": str(tmp_root),
    }
    write(directory / "started.json", started)
    stdout.touch(mode=0o664, exist_ok=False)
    stderr.touch(mode=0o664, exist_ok=False)
    process: subprocess.Popen[bytes] | None = None
    timed_out = False
    termination: str | None = None
    launch_error: str | None = None
    exit_code: int | None = None
    validation_error: str | None = None
    profiler_data: dict[str, Any] | None = None
    report_data: dict[str, Any] | None = None
    resource_data: dict[str, Any] | None = None
    run_environment = dict(os.environ)
    run_environment.update(ENV)
    run_environment["TMPDIR"] = str(tmp_root)
    try:
        with stdout.open("wb") as out, stderr.open("wb") as err:
            process = subprocess.Popen(
                argv,
                cwd=REPO,
                env=run_environment,
                stdin=subprocess.DEVNULL,
                stdout=out,
                stderr=err,
                start_new_session=True,
            )
            try:
                process.communicate(timeout=timeout_seconds)
            except subprocess.TimeoutExpired:
                timed_out = True
                termination = _kill_group(process)
            exit_code = process.returncode
    except (OSError, subprocess.SubprocessError) as error:
        launch_error = f"{type(error).__name__}: {error}"
        if process is not None and process.poll() is None:
            termination = _kill_group(process)
            exit_code = process.returncode
    if exit_code == 0 and not timed_out and launch_error is None:
        try:
            report_data = matrix.validate_report(
                report, role="normal", arm_name=arm_name, binary=binary,
                samples=SAMPLES, warmups=WARMUPS, source_revision=source_revision,
            )
            resource_data = matrix._resource_values(resource)
            profiler_data = _parse_perf(profiler) if tool == "perf" else _parse_strace(profiler)
        except (matrix.ProviderMatrixError, ProfileError, OSError, ValueError) as error:
            validation_error = str(error)
    cleanup: dict[str, Any]
    try:
        cleanup = _cleanup_private(run_root, tmp_root)
    except (ProfileError, OSError) as error:
        cleanup = {
            "schema": "docx-provider-profile-private-cleanup-v1",
            "status": "failed",
            "run_root": str(run_root),
            "tmpdir": str(tmp_root),
            "removed": [],
            "remaining": [str(error)],
        }
    write(directory / "replay-cleanup.json", cleanup)
    artifacts: dict[str, dict[str, Any]] = {}
    for path in (stdout, stderr, resource, report, profiler, directory / "replay-cleanup.json"):
        if path.is_file() and not path.is_symlink():
            artifacts[path.name] = _artifact(path)
    missing = [
        path.name for path in (stdout, stderr, resource, report, profiler, directory / "replay-cleanup.json")
        if not path.is_file()
    ]
    passed = (
        exit_code == 0 and not timed_out and termination is None and launch_error is None
        and validation_error is None and not missing and cleanup.get("status") == "pass"
    )
    terminal = dict(started)
    terminal.update({
        "schema": TERMINAL_SCHEMA,
        "status": "pass" if passed else "failed",
        "exit_code": exit_code,
        "timed_out": timed_out,
        "termination": termination,
        "finished_utc": now(),
        "artifacts": artifacts,
        "missing_artifacts": missing,
        "cleanup": cleanup,
        "tmpdir": str(tmp_root),
        "started_artifact": _artifact(directory / "started.json"),
        "profiler": profiler_data,
        "report_summary": _report_summary(report_data) if report_data is not None else None,
        "resource": resource_data,
    })
    if launch_error is not None:
        terminal["launch_error"] = launch_error
    if validation_error is not None:
        terminal["validation_error"] = validation_error
    # The artifact inventory intentionally excludes terminal.json itself; its
    # own hash is bound by the attempt result after it is written.
    write(directory / "terminal.json", terminal)
    result = {
        "label": label,
        "tool": tool,
        "arm": arm_name,
        "status": terminal["status"],
        "terminal": _artifact(directory / "terminal.json"),
        "report": artifacts.get("report.json"),
        "resource": artifacts.get("resource.txt"),
        "profiler": artifacts.get(profiler.name),
        "tool_executable": tool_executable,
        "scope": WHOLE_CHILD_SCOPE,
        "operation_attribution_claim": False,
    }
    return result


def _write_result(attempt: str, records: list[dict[str, Any]], build_path: Path,
                  build: Mapping[str, Any], protocol_hash: str,
                  source_revision: str, helpers: Mapping[str, Any]) -> Path:
    path = PROFILE_ROOT / attempt / "result.json"
    value = {
        "schema": RESULT_SCHEMA,
        "version": VERSION,
        "status": "pass" if all(row["status"] == "pass" for row in records) else "incomplete",
        "attempt": attempt,
        "runs": [f"{tool}:{arm}" for tool, arm in RUNS],
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "scope": WHOLE_CHILD_SCOPE,
        "operation_attribution_claim": False,
        "performance_claim": NO_ATTRIBUTION_CLAIM,
        "caller_gate": dict(CALLER_GATE),
        "cpu": CPU,
        "protocol": {"path": "protocol.json", "sha256": protocol_hash},
        "build": {"path": str(build_path), "sha256": sha(build_path), "source": build["source"]},
        "source_revision": source_revision,
        "helpers": dict(helpers),
        "records": records,
        "finished_utc": now(),
    }
    write(path, value)
    return path


def run(attempt: str, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> int:
    attempt = _attempt(attempt)
    require(timeout_seconds > 0, "timeout must be positive")
    attempt_root = PROFILE_ROOT / attempt
    require(not attempt_root.exists(), f"refusing to replace profile attempt: {attempt_root}")
    attempt_root.mkdir(parents=True, exist_ok=False)
    build_path, build = _load_build()
    protocol, protocol_hash = _load_protocol()
    source_revision = _source_revision()
    require(source_revision == build["git_revision"], "profile revision differs from build")
    require(protocol["source"]["sha256"] == build["source"]["sha256"], "profile protocol source differs from build")
    helpers = _helper_hashes()
    records: list[dict[str, Any]] = []
    for tool, arm_name in RUNS:
        try:
            records.append(_profile_one(
                attempt, tool, arm_name, build, build_path, protocol_hash,
                source_revision, helpers, protocol["machine"], timeout_seconds,
            ))
        except (ProfileError, OSError, ValueError, KeyError) as error:
            records.append({"label": f"{tool}-{arm_name}", "tool": tool, "arm": arm_name,
                            "status": "failed", "error": str(error)})
    result_path = _write_result(attempt, records, build_path, build, protocol_hash, source_revision, helpers)
    print(result_path)
    return 0 if all(row["status"] == "pass" for row in records) else 1


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("attempt")
    parser.add_argument("--timeout-seconds", type=int, default=DEFAULT_TIMEOUT_SECONDS)
    parser.add_argument("--verify", action="store_true")
    return parser


def verify(attempt: str) -> dict[str, Any]:
    """Revalidate every retained profile artifact and its post-capture gate."""

    attempt = _attempt(attempt)
    directory = PROFILE_ROOT / attempt
    result = _read_json(directory / "result.json")
    require(isinstance(result, dict) and set(result) == RESULT_KEYS,
            "profile result schema or fields differ")
    build_path, build = _load_build()
    protocol, protocol_hash = _load_protocol()
    helpers = _helper_hashes()
    expected_build = {
        "path": str(build_path),
        "sha256": sha(build_path),
        "source": build["source"],
    }
    expected_protocol = {"path": "protocol.json", "sha256": protocol_hash}
    expected_environment = {key: ENV[key] for key in ENV_KEYS}
    require(result["schema"] == RESULT_SCHEMA and result["version"] == VERSION,
            "profile result schema or version differs")
    require(result["status"] == "pass" and result["attempt"] == attempt,
            "incomplete profile attempt or attempt identity differs")
    require(result["runs"] == [f"{tool}:{arm}" for tool, arm in RUNS],
            "profile run inventory differs")
    require(result["samples"] == SAMPLES and result["warmups"] == WARMUPS,
            "profile sample contract differs")
    require(result["scope"] == WHOLE_CHILD_SCOPE and
            result["operation_attribution_claim"] is False and
            result["performance_claim"] == NO_ATTRIBUTION_CLAIM,
            "profile scope or attribution claim differs")
    require(result["caller_gate"] == CALLER_GATE and result["cpu"] == CPU,
            "profile caller gate or CPU binding differs")
    require(result["protocol"] == expected_protocol and result["build"] == expected_build,
            "profile protocol or build binding differs")
    require(result["source_revision"] == build["git_revision"],
            "profile build revision changed")
    require(result["helpers"] == helpers, "profile helpers changed")
    require(isinstance(result["records"], list) and len(result["records"]) == len(RUNS),
            "profile result record inventory differs")
    previous_end: _datetime.datetime | None = None
    for record, (tool, arm) in zip(result["records"], RUNS, strict=True):
        label = f"{tool}-{arm}"
        require(isinstance(record, dict) and set(record) == RECORD_KEYS,
                f"{label}: profile result record fields differ")
        child = directory / label
        require(child.is_dir() and not child.is_symlink(), f"{label}: profile directory is missing")
        terminal_path = child / "terminal.json"
        started_path = child / "started.json"
        profiler_name = "perf.txt" if tool == "perf" else "strace.txt"
        require(
            {item.name for item in child.iterdir()}
            == set(PROFILE_ARTIFACT_NAMES) | {profiler_name, "started.json", "terminal.json"},
            f"{label}: profile artifact inventory differs",
        )
        started = _read_json(started_path)
        terminal = _read_json(terminal_path)
        tool_executable = _tool_executable(tool)
        spec = _spec(arm)
        run_root = TEMP / "provider-profiles" / attempt / label
        tmp_root = run_root / "tmp"
        expected_started = {
            "schema": SCHEMA,
            "version": VERSION,
            "status": "running",
            "attempt": attempt,
            "label": label,
            "tool": tool,
            "tool_executable": tool_executable,
            "arm": dict(matrix.ARM_BY_NAME[arm]),
            "run": dict(spec),
            "samples": SAMPLES,
            "warmups": WARMUPS,
            "scope": WHOLE_CHILD_SCOPE,
            "operation_attribution_claim": False,
            "performance_claim": NO_ATTRIBUTION_CLAIM,
            "caller_gate": dict(CALLER_GATE),
            "cpu": CPU,
            "protocol": expected_protocol,
            "build": expected_build,
            "binary": build["binary"],
            "source": build["source"],
            "source_revision": build["git_revision"],
            "helpers": helpers,
            "machine": protocol["machine"],
            "argv": _benchmark_argv(
                spec, build["binary"], child / "report.json", child / "resource.txt",
                build["git_revision"], child / ("perf.txt" if tool == "perf" else "strace.txt"),
                tool, tool_executable["path"],
            ),
            "cwd": str(REPO),
            "environment": _expected_environment(tmp_root),
            "started_utc": started.get("started_utc"),
            "timeout_seconds": started.get("timeout_seconds"),
            "tmpdir": str(tmp_root),
        }
        require(isinstance(started, dict) and set(started) == STARTED_KEYS,
                f"{label}: started header fields differ")
        require(started == expected_started, f"{label}: started header binding differs")
        _timestamp(started["started_utc"], f"{label}.started_utc")
        require(type(started["timeout_seconds"]) is int and started["timeout_seconds"] > 0,
                f"{label}: timeout binding is malformed")
        require(isinstance(terminal, dict) and set(terminal) == TERMINAL_KEYS,
                f"{label}: terminal header fields differ")
        for key in STARTED_KEYS - {"schema", "status"}:
            require(terminal[key] == started[key], f"{label}: terminal {key} binding differs")
        require(terminal["schema"] == TERMINAL_SCHEMA and terminal["status"] == "pass",
                f"{label}: terminal status or schema differs")
        require(terminal["exit_code"] == 0 and terminal["timed_out"] is False and
                terminal["termination"] is None and terminal["missing_artifacts"] == [],
                f"{label}: profile process failed")
        require(terminal["started_artifact"] == _artifact(started_path),
                f"{label}: started artifact binding differs")
        artifacts = _profile_artifacts(child, tool)
        require(terminal["artifacts"] == artifacts, f"{label}: artifact inventory differs")
        cleanup = _validate_cleanup(
            terminal["cleanup"], run_root, tmp_root, f"{label}: cleanup"
        )
        require(cleanup == _read_json(child / "replay-cleanup.json"),
                f"{label}: cleanup receipt differs")
        profiler = child / ("perf.txt" if tool == "perf" else "strace.txt")
        report = matrix.validate_report(
            child / "report.json", role="normal", arm_name=arm, binary=build["binary"],
            samples=SAMPLES, warmups=WARMUPS, source_revision=build["git_revision"],
        )
        require(terminal["report_summary"] == _report_summary(report),
                f"{label}: profile report summary differs")
        require(terminal["resource"] == matrix._resource_values(child / "resource.txt"),
                f"{label}: profile resource differs")
        parsed = _parse_perf(profiler) if tool == "perf" else _parse_strace(profiler)
        require(terminal["profiler"] == parsed, f"{label}: profile raw counters differ")
        expected_record = {
            "label": label,
            "tool": tool,
            "arm": arm,
            "status": "pass",
            "terminal": _artifact(terminal_path),
            "report": artifacts["report.json"],
            "resource": artifacts["resource.txt"],
            "profiler": artifacts[profiler.name],
            "tool_executable": tool_executable,
            "scope": WHOLE_CHILD_SCOPE,
            "operation_attribution_claim": False,
        }
        require(record == expected_record, f"{label}: profile record binding differs")
        start = _timestamp(started["started_utc"], label)
        end = _timestamp(terminal["finished_utc"], label)
        require(end > start and (previous_end is None or start >= previous_end),
                f"{label}: profile chronology differs")
        previous_end = end
    result_end = _timestamp(result["finished_utc"], "profile result")
    require(previous_end is not None and result_end >= previous_end,
            "profile result finished before terminal capture")
    return _validate_profile_gate(build) if attempt == "profiles2" else {}


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.verify:
            verify(args.attempt)
            print("provider profiles verified")
            return 0
        return run(args.attempt, args.timeout_seconds)
    except (ProfileError, OSError, ValueError, KeyError, subprocess.SubprocessError) as error:
        print(f"profile_providers.py: FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
