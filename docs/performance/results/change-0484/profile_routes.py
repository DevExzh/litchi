#!/usr/bin/env python3
"""Collect external resource profiles for frozen DOCX replay routes.

This is an opt-in diagnostic driver.  It never builds a binary or freezes a
protocol and it does not report a latency, throughput, speedup, or allocation
claim.  The coordinator supplies the normal route executable and invokes this
command under the existing gate CPU lock.  Each selected profiler starts a
fresh child for one route/case pair, retaining the route report, raw streams,
resource record, and profiler artifacts in a separate ``route-profiles``
tree.

The file-store route receives a fresh caller-owned replay directory, the
frozen ceiling, and the frozen sync policy for every child.  That directory is
removed only after a successful report validation and an exact empty-directory
check.  Failed runs retain it for diagnosis.  Missing tools and profiler
permission failures are recorded as unavailable, with no missing counter
converted to zero.
"""

from __future__ import annotations

import argparse
import csv
import math
import os
from pathlib import Path
import shutil
import signal
import subprocess
import time
from typing import Any, Sequence

import measure_routes as routes
from common import ENV, ENV_KEYS, REPO, ROOT, TEMP, meta, now, read, write


SCHEMA = "docx-replayable-tail-append-route-profile-v1"
RUN_SCHEMA = "docx-replayable-tail-append-route-profile-run-v1"
PROFILE_CASES = ("s131072-a64-short-c64", "s64-a16384-short-c64")
PROFILE_TOOLS = ("perf-stat", "perf-record", "strace", "heaptrack")
PERF_EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "cache-references",
    "cache-misses",
    "L1-dcache-loads",
    "L1-dcache-load-misses",
    "LLC-loads",
    "LLC-load-misses",
    "page-faults",
)
PROFILE_SAMPLES = 1
PROFILE_WARMUPS = 1
DEFAULT_TIMEOUT_SECONDS = 900.0
TOOL_ALIASES = {
    "perf_stat": "perf-stat",
    "perf_record": "perf-record",
}
TOOL_EXECUTABLES = {
    "perf-stat": "perf",
    "perf-record": "perf",
    "strace": "strace",
    "heaptrack": "heaptrack",
}


class ProfileError(RuntimeError):
    """A profile identity, command, or retained artifact failed closed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ProfileError(message)


def _display_path(path: Path) -> str:
    path = path.resolve()
    try:
        return path.relative_to(ROOT).as_posix()
    except ValueError:
        try:
            return path.relative_to(REPO).as_posix()
        except ValueError:
            return str(path)


def _driver_binding() -> dict[str, Any]:
    details = meta(Path(__file__))
    return {"path": _display_path(Path(__file__)), **details}


def _artifact(path: Path) -> dict[str, Any]:
    value: dict[str, Any] = {"path": _display_path(path), "present": False}
    if path.is_file() and not path.is_symlink():
        value.update(meta(path))
        value["present"] = True
    return value


def _artifacts(directory: Path) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    if not directory.exists():
        return result
    for path in sorted(directory.rglob("*")):
        if path.is_file() and not path.is_symlink():
            relative = path.relative_to(directory).as_posix()
            if relative in {"started.json", "receipt.json"}:
                continue
            result[relative] = _artifact(path)
    return result


def _touch_empty(path: Path) -> None:
    with path.open("xb"):
        pass


def _normalize_selection(
    values: Sequence[str] | None,
    allowed: Sequence[str],
    defaults: Sequence[str],
    *,
    label: str,
) -> tuple[str, ...]:
    selected = list(defaults if values is None else values)
    normalized: list[str] = []
    for value in selected:
        value = TOOL_ALIASES.get(value, value)
        if value == "all":
            normalized.extend(allowed)
            continue
        if value not in allowed:
            raise ProfileError(f"unknown {label}: {value}")
        normalized.append(value)
    result = tuple(dict.fromkeys(normalized))
    require(result, f"at least one {label} is required")
    return result


def _tool_info(tool: str) -> dict[str, Any]:
    executable = TOOL_EXECUTABLES[tool]
    path = shutil.which(executable)
    if path is None:
        return {
            "name": tool,
            "executable": executable,
            "path": None,
            "status": "unavailable",
            "reason": f"{executable} is not on PATH",
        }
    return {
        "name": tool,
        "executable": executable,
        "path": str(Path(path).resolve()),
        "status": "available",
    }


def _heaptrack_print_info() -> dict[str, Any]:
    path = shutil.which("heaptrack_print")
    if path is None:
        return {
            "name": "heaptrack_print",
            "path": None,
            "status": "unavailable",
            "reason": "heaptrack_print is not on PATH",
        }
    return {
        "name": "heaptrack_print",
        "path": str(Path(path).resolve()),
        "status": "available",
    }


def _classify_failure(
    tool: str,
    returncode: int | None,
    stderr: str,
    *,
    timed_out: bool = False,
) -> str | None:
    """Classify a failed profiler without turning absent counters into zeroes."""

    if timed_out:
        return "timeout"
    if returncode is None:
        return "launch_error"
    if returncode == 0:
        return None
    lowered = stderr.lower()
    if any(
        marker in lowered
        for marker in (
            "permission denied",
            "operation not permitted",
            "no permission to enable",
            "perf_event_paranoid",
            "access denied",
        )
    ):
        return "permission_denied"
    if any(
        marker in lowered
        for marker in (
            "not supported",
            "unsupported",
            "event syntax error",
            "failed to open",
        )
    ) and tool.startswith("perf"):
        return "unsupported"
    return "failed"


def _parse_perf_stat(path: Path) -> dict[str, dict[str, Any]]:
    """Parse only counters explicitly emitted by perf; missing values stay null."""

    parsed: dict[str, dict[str, Any]] = {
        event: {"value": None, "available": False}
        for event in PERF_EVENTS
    }
    if not path.is_file():
        return parsed
    with path.open("r", encoding="utf-8", errors="replace", newline="") as stream:
        for row in csv.reader(stream):
            if len(row) < 3:
                continue
            event = row[2].strip()
            if event not in parsed:
                continue
            raw = ",".join(row).strip()
            value_text = row[0].strip().replace(" ", "")
            value: int | float | None
            if not value_text or value_text.startswith("<"):
                value = None
            else:
                try:
                    numeric = float(value_text)
                    value = int(numeric) if numeric.is_integer() else numeric
                except ValueError:
                    value = None
            parsed[event] = {
                "value": value,
                "available": value is not None,
                "raw": raw,
            }
    return parsed


def _command_for(
    tool: str,
    tool_info: dict[str, Any],
    route_argv: Sequence[str],
    directory: Path,
) -> list[str]:
    require(tool_info.get("status") == "available", f"{tool}: profiler is unavailable")
    profiler = str(tool_info["path"])
    if tool == "perf-stat":
        return [
            profiler,
            "stat",
            "-x,",
            "--no-big-num",
            "-o",
            str(directory / "perf-stat.txt"),
            "-e",
            ",".join(PERF_EVENTS),
            "--",
            *route_argv,
        ]
    if tool == "perf-record":
        return [
            profiler,
            "record",
            "-F",
            "99",
            "--call-graph",
            "dwarf",
            "-o",
            str(directory / "perf.data"),
            "--",
            *route_argv,
        ]
    if tool == "strace":
        return [
            profiler,
            "-f",
            "-qq",
            "-ttt",
            "-T",
            "-yy",
            "-e",
            "trace=read,write,pread64,pwrite64,readv,writev,lseek,openat,close,fsync,fdatasync,unlink,unlinkat",
            "-o",
            str(directory / "strace.log"),
            "--",
            *route_argv,
        ]
    if tool == "heaptrack":
        # _route_argv deliberately starts with the existing /usr/bin/time and
        # taskset prefix.  Keep those outer wrappers for the resource record
        # and CPU binding, but inject heaptrack immediately before the actual
        # benchmark executable so allocation attribution belongs to the route
        # process rather than to the time/taskset wrappers.
        require(len(route_argv) >= 8, "route argv is too short for the time/taskset prefix")
        outer = list(route_argv[:7])
        target = list(route_argv[7:])
        return [
            *outer,
            profiler,
            "--record-only",
            "-o",
            str(directory / "heaptrack-profile"),
            *target,
        ]
    raise ProfileError(f"unknown profiler: {tool}")


def _run_command(
    command: Sequence[str],
    *,
    stdout_path: Path,
    stderr_path: Path,
    timeout_seconds: float,
) -> dict[str, Any]:
    with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
        process: subprocess.Popen[bytes] | None = None
        try:
            process = subprocess.Popen(
                list(command),
                cwd=REPO,
                env=ENV,
                stdout=stdout,
                stderr=stderr,
                start_new_session=True,
            )
            try:
                return {
                    "returncode": process.wait(timeout=timeout_seconds),
                    "timed_out": False,
                    "route_started": True,
                }
            except subprocess.TimeoutExpired:
                # The profiler may have spawned the route, /usr/bin/time, and
                # taskset children.  Terminate this fresh process group so a
                # timed-out child cannot continue writing into a later profile.
                termination: dict[str, Any] = {
                    "pid": process.pid,
                    "start_new_session": True,
                    "signals": [],
                    "term_grace_seconds": 0.25,
                }
                try:
                    os.killpg(process.pid, signal.SIGTERM)
                    termination["signals"].append("SIGTERM")
                except ProcessLookupError:
                    pass
                # Do not reap the leader yet.  Its PID remains reserved while
                # it is a child of this process, so the process-group ID cannot
                # be reused during the bounded grace period.  This observes
                # descendants even when the profiler parent exits immediately
                # after SIGTERM.
                term_deadline = time.monotonic() + 0.25
                group_alive = True
                while time.monotonic() < term_deadline:
                    try:
                        os.killpg(process.pid, 0)
                    except ProcessLookupError:
                        group_alive = False
                        break
                    except PermissionError:
                        group_alive = True
                        break
                    time.sleep(0.01)
                termination["group_alive_before_sigkill"] = group_alive
                if group_alive:
                    try:
                        os.killpg(process.pid, signal.SIGKILL)
                        termination["signals"].append("SIGKILL")
                    except ProcessLookupError:
                        pass
                    kill_deadline = time.monotonic() + 1.0
                    while time.monotonic() < kill_deadline:
                        try:
                            os.killpg(process.pid, 0)
                        except ProcessLookupError:
                            group_alive = False
                            break
                        except PermissionError:
                            group_alive = True
                            break
                        time.sleep(0.01)
                termination["group_alive_after_sigkill"] = group_alive
                # Reap the profiler parent only after the final group signal.
                process.wait()
                return {
                    "returncode": process.returncode,
                    "timed_out": True,
                    "route_started": True,
                    "process_group_termination": termination,
                }
        except subprocess.TimeoutExpired:
            # Kept for unusual Popen implementations that raise at creation
            # time; the normal timeout path above handles process groups.
            return {"returncode": None, "timed_out": True, "route_started": False}
        except OSError as error:
            return {
                "returncode": None,
                "timed_out": False,
                "route_started": False,
                "launch_error_type": type(error).__name__,
                "launch_error": f"{type(error).__name__}: {error}",
            }


def _load_build_binding(
    binary_path: Path,
    attempt: str,
    protocol_sha256: str,
    build_receipt: Path | None,
) -> dict[str, Any]:
    path = (build_receipt or (ROOT / "route-attempts" / attempt / "build-normal.json")).resolve()
    require(path.is_file() and not path.is_symlink(), f"normal route build receipt is missing: {path}")
    value = read(path)
    require(isinstance(value, dict), f"{path}: build receipt is not an object")
    require(value.get("schema") == routes.base.BUILD_SCHEMA, f"{path}: build schema differs")
    require(value.get("version") == 1, f"{path}: build version differs")
    require(value.get("attempt") == attempt, f"{path}: build attempt differs")
    require(value.get("role") == "normal", f"{path}: profiling requires the normal build")
    protocol = value.get("protocol")
    require(
        isinstance(protocol, dict)
        and protocol.get("path") == routes.ROUTE_PROTOCOL_FILE
        and protocol.get("sha256") == protocol_sha256,
        f"{path}: build is not bound to the frozen route protocol",
    )
    require(value.get("source_unchanged") is True, f"{path}: source custody failed")
    require(value.get("source_before") == value.get("source_after"), f"{path}: source identities differ")
    binary = value.get("binary")
    require(isinstance(binary, dict), f"{path}: copied binary identity is missing")
    expected_path = Path(str(binary.get("path", ""))).resolve()
    require(expected_path == binary_path.resolve(), f"{path}: supplied binary differs from build binary")
    actual_binary = routes.base._binary_metadata(binary_path, f"{path}.binary")
    require(actual_binary == binary, f"{path}: copied binary metadata changed")
    return {
        "path": _display_path(path),
        "sha256": routes.base._json_hash(path),
        "attempt": attempt,
        "role": "normal",
        "protocol": dict(protocol),
        "source_before": value["source_before"],
        "source_after": value["source_after"],
        "binary": dict(binary),
    }


def _replay_binding(
    spec: routes.RouteSpec,
    replay_dir: Path | None,
) -> dict[str, Any]:
    return {
        "route": spec.name,
        "directory": str(replay_dir.resolve()) if replay_dir is not None else None,
        "max_bytes": spec.replay_max_bytes,
        "sync": spec.replay_sync,
        "cleanup": "successful_empty_directory_only",
    }


def _cleanup_file_replay(replay_dir: Path, run_scratch: Path) -> tuple[bool, str | None]:
    if not replay_dir.is_dir() or replay_dir.is_symlink():
        return False, f"file route replay directory is missing or not regular: {replay_dir}"
    try:
        if any(replay_dir.iterdir()):
            return False, "file route replay directory retained files after successful child"
        replay_dir.rmdir()
        run_scratch.rmdir()
    except OSError as error:
        return False, f"file route scratch cleanup failed: {type(error).__name__}: {error}"
    return True, None


def _heaptrack_capture_paths(directory: Path) -> list[Path]:
    return [
        path
        for path in sorted(directory.glob("heaptrack-profile*"))
        if path.is_file() and not path.is_symlink()
    ]


def _nonempty_regular(path: Path) -> bool:
    return path.is_file() and not path.is_symlink() and path.stat().st_size > 0


def _missing_profile_artifacts(
    tool: str,
    directory: Path,
    report: Path,
    resource: Path,
) -> list[str]:
    """Return required empty/missing artifacts before a profile can be ok."""

    required = [report, resource]
    if tool == "perf-stat":
        required.append(directory / "perf-stat.txt")
    elif tool == "perf-record":
        required.append(directory / "perf.data")
    elif tool == "strace":
        required.append(directory / "strace.log")
    elif tool == "heaptrack":
        captures = _heaptrack_capture_paths(directory)
        if not any(_nonempty_regular(path) for path in captures):
            return [
                *(_display_path(path) for path in required if not _nonempty_regular(path)),
                f"{_display_path(directory)}/heaptrack-profile*",
            ]
    return [_display_path(path) for path in required if not _nonempty_regular(path)]


def _run_profile(
    *,
    attempt: str,
    binary: dict[str, Any],
    protocol: dict[str, Any],
    protocol_sha256: str,
    build: dict[str, Any],
    tool: str,
    tool_info: dict[str, Any],
    case_label: str,
    route_name: str,
    destination: Path,
    scratch: Path,
    timeout_seconds: float,
) -> dict[str, Any]:
    spec = routes.ROUTE_BY_NAME[route_name]
    case = dict(routes.ROUTE_CASE_BY_LABEL[case_label])
    label = f"{tool}-{route_name}-{case_label}"
    directory = destination / label
    directory.mkdir()
    stdout_path = directory / "stdout.txt"
    stderr_path = directory / "stderr.txt"
    report = directory / "report.json"
    resource = directory / "resource.txt"
    run_scratch: Path | None = None
    replay_dir: Path | None = None
    if route_name == "file_store" and tool_info.get("status") == "available":
        run_scratch = scratch / label
        replay_dir = run_scratch / "replay"
        replay_dir.mkdir(parents=True)
    route_argv = routes._route_argv(
        binary,
        case,
        spec,
        samples=PROFILE_SAMPLES,
        warmups=PROFILE_WARMUPS,
        report=report,
        resource=resource,
        replay_dir=replay_dir,
    ) if tool_info.get("status") == "available" else None
    command = (
        _command_for(tool, tool_info, route_argv, directory)
        if route_argv is not None
        else None
    )
    started = {
        "schema": RUN_SCHEMA,
        "version": 1,
        "status": "running",
        "attempt": attempt,
        "label": label,
        "driver": _driver_binding(),
        "profile_driver_sha256": _driver_binding()["sha256"],
        "profile_only": True,
        "performance_claim": "none",
        "tool": dict(tool_info),
        "route": {
            "name": route_name,
            "cli_provider": spec.cli_provider,
            "report_provider": spec.report_provider,
            "replay_max_bytes": spec.replay_max_bytes,
            "replay_sync": spec.replay_sync,
        },
        "case": case,
        "samples": PROFILE_SAMPLES,
        "warmups": PROFILE_WARMUPS,
        "binary": dict(binary),
        "protocol": {"path": routes.ROUTE_PROTOCOL_FILE, "sha256": protocol_sha256},
        "build": dict(build),
        "argv": route_argv,
        "command": command,
        "cwd": str(REPO),
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "replay": _replay_binding(spec, replay_dir),
        "started_utc": now(),
    }
    write(directory / "started.json", started)

    status = "unavailable"
    failure_class: str | None = None
    error: str | None = None
    exit_code: int | None = None
    timed_out = False
    command_result: dict[str, Any] = {"status": "not_run"}
    validation: dict[str, Any] | None = None
    missing_profile_artifacts: list[str] = []
    profiler: dict[str, Any] = {"name": tool, "scope": "whole child process; profiler overhead included"}
    if tool_info.get("status") != "available":
        _touch_empty(stdout_path)
        _touch_empty(stderr_path)
        profiler.update({"status": "unavailable", "reason": tool_info.get("reason")})
    else:
        command_result = _run_command(
            command or [],
            stdout_path=stdout_path,
            stderr_path=stderr_path,
            timeout_seconds=timeout_seconds,
        )
        exit_code = command_result.get("returncode")
        timed_out = bool(command_result.get("timed_out"))
        stderr_text = ""
        if stderr_path.is_file():
            stderr_text = stderr_path.read_text(encoding="utf-8", errors="replace")
        failure_class = _classify_failure(tool, exit_code, stderr_text, timed_out=timed_out)
        launch_unavailable = (
            command_result.get("route_started") is False
            and command_result.get("launch_error_type") in {"PermissionError", "FileNotFoundError"}
        )
        if launch_unavailable:
            failure_class = "profiler_unavailable"
            status = "unavailable"
            error = str(command_result.get("launch_error") or "profiler could not be launched")
        elif command_result.get("launch_error"):
            error = str(command_result["launch_error"])
        elif failure_class is not None:
            error = f"profiler exited {exit_code}: {failure_class}"
        if launch_unavailable or failure_class in {"permission_denied", "unsupported"}:
            status = "unavailable"
        elif failure_class is not None:
            status = "failed"
        else:
            status = "ok"

        if status == "ok":
            missing_profile_artifacts = _missing_profile_artifacts(
                tool,
                directory,
                report,
                resource,
            )
            if missing_profile_artifacts:
                status = "failed"
                failure_class = "missing_profile_artifact"
                error = "required profile artifacts are missing or empty: " + ", ".join(
                    missing_profile_artifacts
                )
            else:
                try:
                    validation = routes.check_route_report(
                        report,
                        "normal",
                        case,
                        spec,
                        samples=PROFILE_SAMPLES,
                        warmups=PROFILE_WARMUPS,
                        binary=binary,
                        argv=route_argv or [],
                        replay_dir=replay_dir,
                    )
                except Exception as caught:
                    status = "failed"
                    failure_class = "report_validation"
                    error = f"{type(caught).__name__}: {caught}"

        if tool == "perf-stat":
            stat_path = directory / "perf-stat.txt"
            profiler["counters"] = _parse_perf_stat(stat_path)
            profiler["counter_status"] = (
                "observed" if any(item["available"] for item in profiler["counters"].values()) else "unavailable"
            )
            if status == "ok" and (not stat_path.is_file() or stat_path.stat().st_size == 0):
                status = "failed"
                failure_class = "missing_profile_artifact"
                error = "perf stat returned success without a nonempty perf-stat.txt"
        elif tool == "perf-record":
            perf_data = directory / "perf.data"
            profiler["sampling_artifacts"] = [_artifact(perf_data)]
            profiler["counter_status"] = "not_measured"
            if status == "ok" and (not perf_data.is_file() or perf_data.stat().st_size == 0):
                status = "failed"
                failure_class = "missing_profile_artifact"
                error = "perf record returned success without a nonempty perf.data"
        elif tool == "strace":
            strace_log = directory / "strace.log"
            profiler["syscall_trace"] = _artifact(strace_log)
            profiler["counter_status"] = "not_measured"
            if status == "ok" and (not strace_log.is_file() or strace_log.stat().st_size == 0):
                status = "failed"
                failure_class = "missing_profile_artifact"
                error = "strace returned success without a nonempty strace.log"
        elif tool == "heaptrack":
            capture_paths = _heaptrack_capture_paths(directory)
            profiler["captures"] = [_artifact(path) for path in capture_paths]
            nonempty_captures = [path for path in capture_paths if path.stat().st_size > 0]
            print_info = _heaptrack_print_info()
            profiler["heaptrack_print"] = dict(print_info)
            profiler["counter_status"] = "unavailable"
            profiler["summary_status"] = "unavailable"
            if nonempty_captures and print_info.get("status") == "available":
                printed = directory / "heaptrack-print.txt"
                print_stderr = directory / "heaptrack-print-stderr.txt"
                print_command = [
                    str(print_info["path"]),
                    "-f",
                    str(nonempty_captures[-1]),
                    "--merge-backtraces=0",
                    "-H",
                    str(directory / "heaptrack-histogram.tsv"),
                ]
                print_result = _run_command(
                    print_command,
                    stdout_path=printed,
                    stderr_path=print_stderr,
                    timeout_seconds=timeout_seconds,
                )
                print_error = None
                print_artifacts_ok = (
                    print_result.get("returncode") == 0
                    and not print_result.get("timed_out")
                    and not print_result.get("launch_error")
                    and _nonempty_regular(printed)
                    and _nonempty_regular(directory / "heaptrack-histogram.tsv")
                )
                if print_result.get("returncode") != 0 or print_result.get("timed_out"):
                    print_error = "heaptrack_print failed"
                elif print_result.get("launch_error"):
                    print_error = str(print_result["launch_error"])
                elif not print_artifacts_ok:
                    print_error = (
                        "heaptrack_print returned success without nonempty "
                        "heaptrack-print.txt and heaptrack-histogram.tsv"
                    )
                profiler["heaptrack_print_run"] = {
                    "command": print_command,
                    **print_result,
                    "error": print_error,
                }
                if print_error is None:
                    profiler["counter_status"] = "observed"
                    profiler["summary_status"] = "observed"
            elif not nonempty_captures:
                profiler["capture_status"] = "missing"
            if status == "ok" and not nonempty_captures:
                status = "failed"
                failure_class = "missing_profile_artifact"
                error = "heaptrack returned success without a nonempty capture"

    cleanup_error: str | None = None
    scratch_cleaned = False
    if status in {"ok", "unavailable"} and route_name == "file_store" and replay_dir is not None and run_scratch is not None:
        if validation is not None or status == "unavailable":
            scratch_cleaned, cleanup_error = _cleanup_file_replay(replay_dir, run_scratch)
            if not scratch_cleaned:
                if status == "ok":
                    status = "failed"
                    failure_class = "replay_cleanup"
                    error = cleanup_error
    elif status in {"failed", "unavailable"} and run_scratch is not None:
        scratch_cleaned = False

    receipt = dict(
        started,
        status=status,
        passed=status == "ok",
        exit_code=exit_code,
        timed_out=timed_out,
        process=command_result,
        finished_utc=now(),
        profiler=profiler,
        validation_status="ok" if validation is not None else "not_run",
        scratch_cleaned=scratch_cleaned,
        artifacts=_artifacts(directory),
    )
    if cleanup_error is not None:
        receipt["scratch_cleanup_error"] = cleanup_error
    if missing_profile_artifacts:
        receipt["missing_profile_artifacts"] = missing_profile_artifacts
    if failure_class is not None:
        receipt["failure_class"] = failure_class
    if error is not None:
        receipt["error"] = error
    write(directory / "receipt.json", receipt)
    print(f"{label}: {status}", flush=True)
    return {
        "label": label,
        "status": status,
        "passed": status == "ok",
        "receipt": _artifact(directory / "receipt.json"),
    }


def _remove_empty_scratch(
    scratch: Path,
    *,
    allow_retained: bool = False,
) -> tuple[bool, str | None]:
    if not scratch.exists():
        return True, None
    if scratch.is_symlink() or not scratch.is_dir():
        return False, f"profile scratch root is not a regular directory: {scratch}"
    try:
        if any(scratch.iterdir()):
            if allow_retained:
                return False, None
            return False, "profile scratch retains failed replay artifacts"
        scratch.rmdir()
    except OSError as error:
        return False, f"profile scratch cleanup failed: {type(error).__name__}: {error}"
    return True, None


def run(
    binary_path: Path,
    attempt: str,
    *,
    build_attempt: str | None = None,
    selected_tools: Sequence[str] | None = None,
    selected_cases: Sequence[str] | None = None,
    selected_routes: Sequence[str] | None = None,
    build_receipt: Path | None = None,
    scratch_root: Path | None = None,
    timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS,
) -> int:
    """Run selected route profiles for one already-built normal binary."""

    attempt = routes.base._attempt(attempt)
    build_attempt = routes.base._attempt(build_attempt or attempt)
    require(math.isfinite(timeout_seconds) and timeout_seconds > 0, "timeout must be finite and positive")
    tools = _normalize_selection(selected_tools, PROFILE_TOOLS, PROFILE_TOOLS, label="profiler")
    cases = _normalize_selection(selected_cases, tuple(routes.ROUTE_CASE_BY_LABEL), PROFILE_CASES, label="case")
    route_names = _normalize_selection(selected_routes, routes.ROUTE_NAMES, routes.ROUTE_NAMES, label="route")
    require(binary_path.is_file() and not binary_path.is_symlink(), f"normal binary is missing: {binary_path}")
    binary_path = binary_path.resolve()

    destination = ROOT / "route-profiles" / attempt
    destination.mkdir(parents=True, exist_ok=False)
    started_path = destination / "started.json"
    scratch = (scratch_root or (TEMP / "route-profiles" / attempt)).resolve()
    failures: list[dict[str, str]] = []
    receipts: list[dict[str, Any]] = []
    protocol: dict[str, Any] | None = None
    protocol_sha256: str | None = None
    build: dict[str, Any] | None = None
    binary: dict[str, Any] | None = None
    tool_info = {tool: _tool_info(tool) for tool in tools}
    if "heaptrack" in tools:
        tool_info["heaptrack_print"] = _heaptrack_print_info()

    try:
        protocol, protocol_sha256 = routes.load_protocol()
        binary = routes.base._binary_metadata(binary_path, "profile normal binary")
        build = _load_build_binding(binary_path, build_attempt, protocol_sha256, build_receipt)
        scratch.parent.mkdir(parents=True, exist_ok=True)
        scratch.mkdir(exist_ok=False)
    except Exception as error:
        failures.append({"label": "setup", "error": f"{type(error).__name__}: {error}"})

    bindings: dict[str, Any] = {
        "schema": SCHEMA,
        "version": 1,
        "attempt": attempt,
        "build_attempt": build_attempt,
        "driver": _driver_binding(),
        "profile_driver_sha256": _driver_binding()["sha256"],
        "profile_only": True,
        "performance_claim": "none",
        "tools": tools,
        "tool_info": tool_info,
        "routes": route_names,
        "cases": cases,
        "samples": PROFILE_SAMPLES,
        "warmups": PROFILE_WARMUPS,
        "timeout_seconds": timeout_seconds,
        "binary": binary,
        "protocol": (
            {"path": routes.ROUTE_PROTOCOL_FILE, "sha256": protocol_sha256}
            if protocol_sha256 is not None
            else None
        ),
        "build": build,
        "scratch": str(scratch),
        "output": _display_path(destination),
        "machine": routes._machine_binding(required=False),
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "started_utc": now(),
    }
    write(started_path, bindings)

    if not failures and protocol is not None and protocol_sha256 is not None and build is not None and binary is not None:
        for tool in tools:
            for route_name in route_names:
                for case_label in cases:
                    try:
                        receipts.append(
                            _run_profile(
                                attempt=attempt,
                                binary=binary,
                                protocol=protocol,
                                protocol_sha256=protocol_sha256,
                                build=build,
                                tool=tool,
                                tool_info=tool_info[tool],
                                case_label=case_label,
                                route_name=route_name,
                                destination=destination,
                                scratch=scratch,
                                timeout_seconds=timeout_seconds,
                            )
                        )
                    except Exception as error:
                        label = f"{tool}-{route_name}-{case_label}"
                        failures.append({"label": label, "error": f"{type(error).__name__}: {error}"})

    if binary is not None:
        try:
            unchanged = routes.base._binary_metadata(binary_path, "profile normal binary after") == binary
        except Exception as error:
            unchanged = False
            failures.append({"label": "binary", "error": f"{type(error).__name__}: {error}"})
    else:
        unchanged = False

    scratch_removed = False
    if not failures:
        allow_retained = bool(receipts) and not any(
            item["status"] == "failed" for item in receipts
        )
        scratch_removed, cleanup_error = _remove_empty_scratch(
            scratch,
            allow_retained=allow_retained,
        )
        if cleanup_error is not None:
            failures.append({"label": "scratch", "error": cleanup_error})

    statuses = [item["status"] for item in receipts]
    if failures or any(status == "failed" for status in statuses):
        status = "failed"
    elif any(status in {"unavailable", "partial"} for status in statuses):
        status = "unavailable"
    else:
        status = "ok"
    result = dict(
        bindings,
        status=status,
        passed=status == "ok" and not failures and unchanged,
        binary_unchanged=unchanged,
        scratch_removed=scratch_removed,
        failures=failures,
        receipts=receipts,
        finished_utc=now(),
        limitations=[
            "External profiler observations cover the whole child process and include profiler overhead.",
            "No elapsed-time, throughput, speedup, cold-cache, or allocation claim is made by this helper.",
            "Missing profiler counters remain unavailable/null; they are never replaced with zero.",
            "Timeout cleanup covers the profiler's private process group; daemonized or escaped-session descendants are outside this helper's custody.",
        ],
    )
    write(destination / "result.json", result)
    # Missing tools are an explicit fallback and do not make the command fail;
    # actual child/validation/cleanup failures still return nonzero.
    return int(status == "failed" or bool(failures) or not unchanged)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--attempt", required=True)
    parser.add_argument(
        "--tool",
        action="append",
        choices=("all", *PROFILE_TOOLS, *TOOL_ALIASES),
        help="repeatable profiler selection; default is all four",
    )
    parser.add_argument(
        "--route",
        action="append",
        choices=routes.ROUTE_NAMES,
        help="repeatable route selection; default is all routes",
    )
    parser.add_argument(
        "--case",
        action="append",
        choices=tuple(routes.ROUTE_CASE_BY_LABEL),
        help="repeatable route-case selection; default is source/authored-heavy pair",
    )
    parser.add_argument("--build-receipt", type=Path)
    parser.add_argument(
        "--build-attempt",
        help="attempt containing build-normal.json; defaults to the profile output attempt",
    )
    parser.add_argument("--scratch-root", type=Path)
    parser.add_argument("--timeout-seconds", type=float, default=DEFAULT_TIMEOUT_SECONDS)
    return parser


def main() -> None:
    args = _parser().parse_args()
    raise SystemExit(
        run(
            args.binary,
            args.attempt,
            build_attempt=args.build_attempt,
            selected_tools=args.tool,
            selected_cases=args.case,
            selected_routes=args.route,
            build_receipt=args.build_receipt,
            scratch_root=args.scratch_root,
            timeout_seconds=args.timeout_seconds,
        )
    )


if __name__ == "__main__":
    main()
