#!/usr/bin/env python3
"""Capture bounded syscall metadata for the deterministic owned/file inputs.

This opt-in diagnostic driver runs the frozen normal route binary for exactly
the source-heavy and authored-heavy deterministic cases, once with the owned
input and once with the prepared file input.  The primary capture is a
strace -f -c summary over metadata and positional-I/O syscalls.  An
optional --raw-authored mode adds one raw trace for each authored-heavy
input so a reviewer can attribute source paths; those raw traces are hashed
before being gzip-compressed.

The helper does not build, freeze, or compare performance.  The coordinator
supplies the normal formal binary and runs this helper under the external gate
and CPU lock.  It reuses the frozen route argv and axis report oracle, and it
uses the already-bound profile command runner for process-group cleanup.
"""

from __future__ import annotations

import argparse
import gzip
import math
from pathlib import Path
import shutil
from typing import Any, Sequence

import measure_routes as routes
import profile_routes as profile_driver
from common import ENV, ENV_KEYS, REPO, ROOT, meta, now, write


SCHEMA = "docx-replayable-tail-append-input-metadata-profile-v1"
RUN_SCHEMA = "docx-replayable-tail-append-input-metadata-profile-run-v1"
PROFILE_CASES = ("s131072-a64-short-c64", "s64-a16384-short-c64")
INPUT_MODES = ("owned", "file")
ROLE = "normal"
ROUTE = "deterministic"
SAMPLES = 1
WARMUPS = 1
DEFAULT_TIMEOUT_SECONDS = 300.0
TRACE_SYSCALLS = (
    "fstat",
    "fstat64",
    "newfstatat",
    "statx",
    "pread64",
    "read",
    "write",
    "open",
    "openat",
    "close",
    "lseek",
    "sync",
    "fsync",
    "fdatasync",
)
TRACE_FILTER = "trace=" + ",".join(TRACE_SYSCALLS)


class InputMetadataError(RuntimeError):
    """The input metadata profile could not be bound or retained."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise InputMetadataError(message)


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
    return {"path": _display_path(Path(__file__)), **meta(Path(__file__))}


def _profile_driver_binding() -> dict[str, Any]:
    """Bind the reused process runner to the exact frozen helper content."""

    return profile_driver._driver_binding()


def _artifact(path: Path) -> dict[str, Any]:
    value: dict[str, Any] = {"path": _display_path(path), "present": False}
    if path.is_file() and not path.is_symlink():
        value.update(meta(path))
        value["present"] = True
    return value


def _nonempty_regular(path: Path) -> bool:
    return path.is_file() and not path.is_symlink() and path.stat().st_size > 0


def _normalize_selection(
    values: Sequence[str] | None,
    allowed: Sequence[str],
    defaults: Sequence[str],
    *,
    label: str,
) -> tuple[str, ...]:
    selected = list(defaults if values is None else values)
    result: list[str] = []
    for value in selected:
        if value == "all":
            result.extend(allowed)
            continue
        require(value in allowed, f"unknown {label}: {value}")
        result.append(value)
    normalized = tuple(dict.fromkeys(result))
    require(normalized, f"at least one {label} is required")
    return normalized


def _arm_for(workload: str, input_mode: str) -> dict[str, Any]:
    require(workload in routes.CASE_BY_LABEL, f"unknown workload: {workload}")
    require(input_mode in INPUT_MODES, f"unknown input mode: {input_mode}")
    if input_mode == "file":
        arm = dict(routes._axis_arm("input", "file", workload))
    else:
        # _axis_arm's baseline compression arm carries the frozen owned-input
        # defaults.  Re-label it so the receipt describes the input comparison
        # explicitly while _axis_argv/_check_axis_report see the exact fields.
        arm = dict(routes._axis_arm("compression", "current", workload))
        arm.update(axis="input", value="owned")
    arm["label"] = f"input-{input_mode}-{workload}"
    return arm


def _input_metadata(arm: dict[str, Any]) -> dict[str, Any]:
    if arm["input_mode"] == "file":
        value = dict(routes._axis_input_metadata(arm) or {})
        value.update(
            mode="file",
            backing="file",
            source_contract=arm["source_contract"],
        )
        return value
    return {
        "mode": "owned",
        "backing": "owned",
        "path": None,
        "absolute_path": None,
        "bytes": None,
        "sha256": None,
        "source_contract": arm["source_contract"],
        "identity": "owned_input_is_created_and_fingerprinted_by_the_route_binary",
    }


def _strace_command(
    tool_info: dict[str, Any],
    route_argv: Sequence[str],
    output: Path,
    *,
    summary: bool,
) -> list[str]:
    require(tool_info.get("status") == "available", "strace is unavailable")
    command = [str(tool_info["path"]), "-f"]
    if summary:
        command.append("-c")
    else:
        command.extend(("-qq", "-ttt", "-T", "-yy"))
    command.extend(("-e", TRACE_FILTER, "-o", str(output), "--", *route_argv))
    return command


def _missing_artifacts(paths: Sequence[Path]) -> list[str]:
    return [_display_path(path) for path in paths if not _nonempty_regular(path)]


def _gzip_after_hash(raw_path: Path) -> dict[str, Any]:
    """Record the raw hash first, then retain a deterministic gzip artifact."""

    raw_binding = meta(raw_path)
    compressed = raw_path.with_name(raw_path.name + ".gz")
    require(not compressed.exists(), f"refusing to replace compressed trace: {compressed}")
    with raw_path.open("rb") as source, compressed.open("xb") as destination:
        with gzip.GzipFile(
            filename=raw_path.name,
            mode="wb",
            fileobj=destination,
            mtime=0,
        ) as stream:
            shutil.copyfileobj(source, stream)
    return {
        "raw_before_compression": {"path": _display_path(raw_path), **raw_binding},
        "compressed": _artifact(compressed),
        "compression": "gzip_after_raw_sha256_record",
    }


def _classify_process(
    command_result: dict[str, Any],
    stderr: str,
) -> tuple[str, str | None, str | None]:
    returncode = command_result.get("returncode")
    timed_out = bool(command_result.get("timed_out"))
    failure = profile_driver._classify_failure("strace", returncode, stderr, timed_out=timed_out)
    launch_unavailable = (
        command_result.get("route_started") is False
        and command_result.get("launch_error_type") in {"PermissionError", "FileNotFoundError"}
    )
    if launch_unavailable:
        return (
            "unavailable",
            "profiler_unavailable",
            str(command_result.get("launch_error") or "strace could not be launched"),
        )
    if failure == "permission_denied":
        return "unavailable", failure, f"strace exited {returncode}: {failure}"
    if failure is not None:
        if command_result.get("launch_error"):
            return "failed", failure, str(command_result["launch_error"])
        return "failed", failure, f"strace exited {returncode}: {failure}"
    return "ok", None, None


def _run_trace(
    *,
    destination: Path,
    binary: dict[str, Any],
    build: dict[str, Any],
    protocol_sha256: str,
    tool_info: dict[str, Any],
    attempt: str,
    build_attempt: str,
    case_label: str,
    arm: dict[str, Any],
    input_metadata: dict[str, Any],
    timeout_seconds: float,
    raw: bool,
) -> dict[str, Any]:
    case = routes._axis_case(arm)
    input_label = arm["input_mode"]
    kind = "raw-metadata" if raw else "summary"
    label = f"strace-{kind}-deterministic-{input_label}-{case_label}"
    directory = destination / label
    directory.mkdir()
    report = directory / "report.json"
    resource = directory / "resource.txt"
    stdout = directory / "stdout.txt"
    stderr = directory / "stderr.txt"
    profile = directory / ("strace-raw.log" if raw else "strace-summary.txt")
    route_argv = routes._axis_argv(
        binary,
        case,
        arm,
        samples=SAMPLES,
        warmups=WARMUPS,
        report=report,
        resource=resource,
    )
    command = (
        _strace_command(tool_info, route_argv, profile, summary=not raw)
        if tool_info.get("status") == "available"
        else None
    )
    started = {
        "schema": RUN_SCHEMA,
        "version": 1,
        "status": "running",
        "attempt": attempt,
        "build_attempt": build_attempt,
        "label": label,
        "profile_kind": kind,
        "driver": _driver_binding(),
        "profile_driver": _profile_driver_binding(),
        "profile_driver_sha256": _profile_driver_binding()["sha256"],
        "route_driver_sha256": routes._script_hashes()["measure_routes.py"],
        "profile_only": True,
        "performance_claim": "none",
        "role": ROLE,
        "route": {
            "name": ROUTE,
            "provider": arm["provider"],
            "cli_provider": arm["cli_provider"],
            "report_provider": arm["report_provider"],
        },
        "case": case,
        "axis": dict(arm),
        "input": dict(input_metadata),
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "binary": dict(binary),
        "build": dict(build),
        "protocol": {"path": routes.ROUTE_PROTOCOL_FILE, "sha256": protocol_sha256},
        "argv": route_argv,
        "command": command,
        "cwd": str(REPO),
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "execution": {
            "gate_required": True,
            "cpu_lock_required": True,
            "cpu": routes.base.CPU,
            "timeout_seconds": timeout_seconds,
        },
        "started_utc": now(),
    }
    write(directory / "started.json", started)

    status = "unavailable"
    failure_class: str | None = None
    error: str | None = None
    exit_code: int | None = None
    timed_out = False
    command_result: dict[str, Any] = {"status": "not_run"}
    validation_status = "not_run"
    missing: list[str] = []
    raw_binding: dict[str, Any] | None = None
    if tool_info.get("status") != "available":
        profile_driver._touch_empty(stdout)
        profile_driver._touch_empty(stderr)
        status = "unavailable"
        failure_class = "profiler_unavailable"
        error = str(tool_info.get("reason") or "strace is unavailable")
    else:
        command_result = profile_driver._run_command(
            command or [],
            stdout_path=stdout,
            stderr_path=stderr,
            timeout_seconds=timeout_seconds,
        )
        exit_code = command_result.get("returncode")
        timed_out = bool(command_result.get("timed_out"))
        stderr_text = stderr.read_text(encoding="utf-8", errors="replace") if stderr.is_file() else ""
        status, failure_class, error = _classify_process(command_result, stderr_text)
        if raw and _nonempty_regular(profile):
            try:
                raw_binding = _gzip_after_hash(profile)
            except Exception as caught:
                status = "failed"
                failure_class = "raw_compression"
                error = f"{type(caught).__name__}: {caught}"
        if status == "ok":
            missing = _missing_artifacts((report, resource, profile))
            if missing:
                status = "failed"
                failure_class = "missing_profile_artifact"
                error = "required input metadata artifacts are missing or empty: " + ", ".join(missing)
            else:
                try:
                    routes._check_axis_report(
                        report,
                        ROLE,
                        arm,
                        samples=SAMPLES,
                        warmups=WARMUPS,
                        binary=binary,
                        argv=route_argv,
                        input_metadata=input_metadata if arm["input_mode"] == "file" else None,
                    )
                    validation_status = "ok"
                except Exception as caught:
                    status = "failed"
                    failure_class = "report_validation"
                    error = f"{type(caught).__name__}: {caught}"

    receipt = dict(
        started,
        status=status,
        passed=status == "ok",
        exit_code=exit_code,
        timed_out=timed_out,
        process=command_result,
        validation_status=validation_status,
        finished_utc=now(),
        profile_artifact=_artifact(profile),
        artifacts=profile_driver._artifacts(directory),
    )
    if raw_binding is not None:
        receipt["raw_trace_binding"] = raw_binding
    if missing:
        receipt["missing_artifacts"] = missing
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


def run(
    binary_path: Path,
    attempt: str,
    build_attempt: str,
    *,
    selected_cases: Sequence[str] | None = None,
    selected_inputs: Sequence[str] | None = None,
    raw_authored: bool = False,
    build_receipt: Path | None = None,
    timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS,
) -> int:
    """Run the bounded metadata inventory for one retained normal binary."""

    attempt = routes.base._attempt(attempt)
    build_attempt = routes.base._attempt(build_attempt)
    require(math.isfinite(timeout_seconds) and timeout_seconds > 0, "timeout must be finite and positive")
    cases = _normalize_selection(selected_cases, PROFILE_CASES, PROFILE_CASES, label="case")
    inputs = _normalize_selection(selected_inputs, INPUT_MODES, INPUT_MODES, label="input mode")
    require(binary_path.is_file() and not binary_path.is_symlink(), f"normal binary is missing: {binary_path}")
    binary_path = binary_path.resolve()

    destination = ROOT / "input-metadata-profiles" / attempt
    destination.mkdir(parents=True, exist_ok=False)
    failures: list[dict[str, str]] = []
    receipts: list[dict[str, Any]] = []
    protocol: dict[str, Any] | None = None
    protocol_sha256: str | None = None
    build: dict[str, Any] | None = None
    binary: dict[str, Any] | None = None
    tool_info = profile_driver._tool_info("strace")
    validators: dict[str, str] | None = None
    try:
        protocol, protocol_sha256 = routes.load_protocol()
        validators = routes._script_hashes()
        binary_value = routes.base._binary_metadata(binary_path, "input metadata normal binary")
        build = profile_driver._load_build_binding(
            binary_path,
            build_attempt,
            protocol_sha256,
            build_receipt,
        )
        binary = binary_value
    except Exception as caught:
        failures.append({"label": "setup", "error": f"{type(caught).__name__}: {caught}"})

    bindings: dict[str, Any] = {
        "schema": SCHEMA,
        "version": 1,
        "attempt": attempt,
        "build_attempt": build_attempt,
        "driver": _driver_binding(),
        "profile_driver": _profile_driver_binding(),
        "profile_driver_sha256": _profile_driver_binding()["sha256"],
        "route_driver_sha256": validators.get("measure_routes.py") if validators else None,
        "validators": validators,
        "profile_only": True,
        "performance_claim": "none",
        "role": ROLE,
        "route": ROUTE,
        "cases": cases,
        "input_modes": inputs,
        "raw_authored": raw_authored,
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "timeout_seconds": timeout_seconds,
        "tool": tool_info,
        "protocol": (
            {"path": routes.ROUTE_PROTOCOL_FILE, "sha256": protocol_sha256}
            if protocol_sha256 is not None
            else None
        ),
        "build": build,
        "binary": binary,
        "execution": {
            "gate_required": True,
            "cpu_lock_required": True,
            "cpu": routes.base.CPU,
            "cwd": str(REPO),
            "environment": {key: ENV[key] for key in ENV_KEYS},
        },
        "output": _display_path(destination),
        "started_utc": now(),
    }
    write(destination / "started.json", bindings)

    if not failures and protocol_sha256 is not None and build is not None and binary is not None:
        for case_label in cases:
            for input_mode in inputs:
                arm = _arm_for(case_label, input_mode)
                input_metadata = _input_metadata(arm)
                try:
                    primary = _run_trace(
                        destination=destination,
                        binary=binary,
                        build=build,
                        protocol_sha256=protocol_sha256,
                        tool_info=tool_info,
                        attempt=attempt,
                        build_attempt=build_attempt,
                        case_label=case_label,
                        arm=arm,
                        input_metadata=input_metadata,
                        timeout_seconds=timeout_seconds,
                        raw=False,
                    )
                    receipts.append(primary)
                    if raw_authored and case_label == "s64-a16384-short-c64" and primary["status"] == "ok":
                        receipts.append(
                            _run_trace(
                                destination=destination,
                                binary=binary,
                                build=build,
                                protocol_sha256=protocol_sha256,
                                tool_info=tool_info,
                                attempt=attempt,
                                build_attempt=build_attempt,
                                case_label=case_label,
                                arm=arm,
                                input_metadata=input_metadata,
                                timeout_seconds=timeout_seconds,
                                raw=True,
                            )
                        )
                except Exception as caught:
                    failures.append({
                        "label": f"{input_mode}-{case_label}",
                        "error": f"{type(caught).__name__}: {caught}",
                    })

    unchanged = False
    if binary is not None:
        try:
            unchanged = routes.base._binary_metadata(binary_path, "input metadata normal binary after") == binary
        except Exception as caught:
            failures.append({"label": "binary", "error": f"{type(caught).__name__}: {caught}"})

    statuses = [item["status"] for item in receipts]
    if failures or any(status == "failed" for status in statuses):
        status = "failed"
    elif any(status == "unavailable" for status in statuses):
        status = "unavailable"
    else:
        status = "ok"
    result = dict(
        bindings,
        status=status,
        passed=status == "ok" and not failures and unchanged,
        binary_unchanged=unchanged,
        failures=failures,
        receipts=receipts,
        finished_utc=now(),
        limitations=[
            "Syscall summaries are diagnostic observations for the selected one-sample executions.",
            "No elapsed-time, throughput, speedup, optimization, or causal claim is made.",
            "The summary filter includes metadata, positional input, output, and sync syscalls; absent or unavailable profiler evidence is retained as unavailable.",
            "Raw trace cleanup is not used to alter the prepared input fixture; failures retain their complete output directories.",
        ],
    )
    write(destination / "result.json", result)
    return int(status == "failed" or bool(failures) or not unchanged)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--attempt", required=True)
    parser.add_argument("--build-attempt", required=True)
    parser.add_argument("--build-receipt", type=Path)
    parser.add_argument(
        "--case",
        action="append",
        choices=PROFILE_CASES,
        help="repeatable case selection; default is both heavy cases",
    )
    parser.add_argument(
        "--input-mode",
        action="append",
        choices=INPUT_MODES,
        help="repeatable input selection; default is owned and file",
    )
    parser.add_argument(
        "--raw-authored",
        action="store_true",
        help="add raw metadata-only traces for the authored-heavy owned/file inputs",
    )
    parser.add_argument("--timeout-seconds", type=float, default=DEFAULT_TIMEOUT_SECONDS)
    return parser


def main() -> None:
    args = _parser().parse_args()
    raise SystemExit(
        run(
            args.binary,
            args.attempt,
            args.build_attempt,
            selected_cases=args.case,
            selected_inputs=args.input_mode,
            raw_authored=args.raw_authored,
            build_receipt=args.build_receipt,
            timeout_seconds=args.timeout_seconds,
        )
    )


if __name__ == "__main__":
    main()
