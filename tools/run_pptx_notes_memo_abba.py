#!/usr/bin/env python3
"""Run the PPTX notes-memo probe in four fresh processes and validate ABBA.

The local probe schema is intentionally not accepted by the default ABBA
summary tool.  This runner provides the missing process boundary: it admits
exactly ``A1, B1, B2, A2`` commands, gives each leg a fresh report path,
captures a bounded process identity and output receipt, and invokes the
dedicated schema validator only after all four child processes finish.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import selectors
import signal
import stat
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Mapping

from tools import pptx_notes_memo_abba as validator


MANIFEST_SCHEMA = "pptx_notes_memo_abba_process_manifest_v1"
RECEIPT_SCHEMA = "pptx_notes_memo_abba_process_receipt_v1"
LEG_ORDER = validator.LEG_ORDER
MANIFEST_KEYS = frozenset({"schema", "mode", "output_dir", "timeout_seconds", "legs"})
LEG_KEYS = frozenset({"label", "command", "report"})
RECEIPT_KEYS = frozenset(
    {"schema", "mode", "leg_order", "process_count", "output_dir", "processes", "summary"}
)
PROCESS_KEYS = frozenset(
    {
        "label",
        "pid",
        "start_time_ticks",
        "command",
        "command_executable",
        "process_executable",
        "report_path",
        "report_bytes",
        "report_sha256",
        "stdout_bytes",
        "stdout_sha256",
        "stderr_bytes",
        "stderr_sha256",
        "source_revision",
        "binary_identity",
    }
)
MAX_REPORT_BYTES = 8 * 1024 * 1024
MAX_LOG_BYTES = 1 * 1024 * 1024
MIN_TIMEOUT_SECONDS = 1.0
MAX_TIMEOUT_SECONDS = 600.0


class RunnerError(ValueError):
    """The external process protocol failed closed."""


def _exact_keys(value: Mapping[str, Any], expected: frozenset[str], path: str) -> None:
    actual = set(value)
    missing = sorted(expected - actual)
    unknown = sorted(actual - expected)
    if missing or unknown:
        details = []
        if missing:
            details.append(f"missing={missing}")
        if unknown:
            details.append(f"unknown={unknown}")
        raise RunnerError(f"{path} has an invalid key set ({'; '.join(details)})")


def _mapping(value: Any, path: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise RunnerError(f"{path} must be an object")
    return value


def _string(value: Any, path: str) -> str:
    if not isinstance(value, str) or not value:
        raise RunnerError(f"{path} must be a non-empty string")
    return value


def _absolute_path(value: Any, path: str) -> Path:
    text = _string(value, path)
    result = Path(text)
    if not result.is_absolute():
        raise RunnerError(f"{path} must be absolute")
    return result


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _read_start_time_ticks(pid: int) -> int:
    if os.name != "posix":
        raise RunnerError("fresh process identity requires a POSIX /proc process clock")
    try:
        contents = Path(f"/proc/{pid}/stat").read_text(encoding="ascii")
    except (OSError, UnicodeError) as error:
        raise RunnerError(f"cannot read the fresh child process identity: {error}") from error
    closing = contents.rfind(")")
    if closing < 0:
        raise RunnerError("fresh child process stat is malformed")
    fields = contents[closing + 2 :].split()
    # The suffix starts at stat field 3 (state); starttime is field 22.
    if len(fields) <= 19:
        raise RunnerError("fresh child process stat lacks start time")
    try:
        start_time = int(fields[19])
    except ValueError as error:
        raise RunnerError("fresh child process start time is not an integer") from error
    if start_time <= 0:
        raise RunnerError("fresh child process start time is not positive")
    return start_time


def _read_process_executable(pid: int) -> str:
    if os.name != "posix":
        raise RunnerError("fresh process identity requires a POSIX executable identity")
    try:
        executable = os.readlink(f"/proc/{pid}/exe")
    except OSError as error:
        raise RunnerError(f"cannot read the fresh child executable identity: {error}") from error
    if not executable.startswith("/"):
        raise RunnerError("fresh child executable identity is not absolute")
    return executable


def _resolve_command_executable(command: list[str], path: str) -> str:
    if not command:
        raise RunnerError(f"{path} must not be empty")
    executable = Path(command[0])
    if not executable.is_absolute():
        raise RunnerError(f"{path}[0] must be an absolute executable path")
    try:
        metadata = executable.stat()
    except OSError as error:
        raise RunnerError(f"{path}[0] cannot be stat'ed: {error}") from error
    if not stat.S_ISREG(metadata.st_mode) or not os.access(executable, os.X_OK):
        raise RunnerError(f"{path}[0] is not an executable regular file")
    return str(executable.resolve())


def _load_manifest(path: Path) -> tuple[dict[str, Any], Path, float, list[dict[str, Any]]]:
    try:
        document = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"cannot read process manifest {path}: {error}") from error
    manifest = _mapping(document, str(path))
    _exact_keys(manifest, MANIFEST_KEYS, str(path))
    if manifest.get("schema") != MANIFEST_SCHEMA:
        raise RunnerError(f"{path}.schema is not {MANIFEST_SCHEMA}")
    mode = manifest.get("mode")
    if mode not in ("native", "allocator"):
        raise RunnerError(f"{path}.mode must be native or allocator")
    output_dir = _absolute_path(manifest.get("output_dir"), f"{path}.output_dir")
    timeout = manifest.get("timeout_seconds")
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)) or not math.isfinite(timeout):
        raise RunnerError(f"{path}.timeout_seconds must be finite")
    if not MIN_TIMEOUT_SECONDS <= timeout <= MAX_TIMEOUT_SECONDS:
        raise RunnerError(f"{path}.timeout_seconds is outside the finite protocol bound")
    legs_value = manifest.get("legs")
    if not isinstance(legs_value, list) or len(legs_value) != len(LEG_ORDER):
        raise RunnerError(f"{path}.legs must contain exactly four process legs")
    legs: list[dict[str, Any]] = []
    seen_reports: set[Path] = set()
    for index, leg_value in enumerate(legs_value):
        leg = _mapping(leg_value, f"{path}.legs[{index}]")
        _exact_keys(leg, LEG_KEYS, f"{path}.legs[{index}]")
        label = leg.get("label")
        if label != LEG_ORDER[index]:
            raise RunnerError(f"{path}.legs[{index}].label must be {LEG_ORDER[index]}")
        command_value = leg.get("command")
        if not isinstance(command_value, list) or not command_value or not all(
            isinstance(argument, str) and argument for argument in command_value
        ):
            raise RunnerError(f"{path}.legs[{index}].command must be a non-empty string array")
        command = list(command_value)
        command_executable = _resolve_command_executable(command, f"{path}.legs[{index}].command")
        report = _absolute_path(leg.get("report"), f"{path}.legs[{index}].report")
        if report.parent != output_dir:
            raise RunnerError(f"{path}.legs[{index}].report must be directly under output_dir")
        if report in seen_reports:
            raise RunnerError("process report paths must be unique")
        seen_reports.add(report)
        legs.append(
            {
                "label": label,
                "command": command,
                "command_executable": command_executable,
                "report": report,
            }
        )
    return {"mode": mode, "output_dir": output_dir}, output_dir, float(timeout), legs


def _terminate_process(process: subprocess.Popen[bytes]) -> None:
    if process.poll() is not None:
        return
    try:
        if os.name == "posix" and process.pid > 0:
            os.killpg(process.pid, signal.SIGTERM)
        else:
            process.terminate()
        process.wait(timeout=1.0)
    except (OSError, subprocess.TimeoutExpired):
        try:
            if os.name == "posix" and process.pid > 0:
                os.killpg(process.pid, signal.SIGKILL)
            else:
                process.kill()
        except OSError:
            pass
        try:
            process.wait(timeout=1.0)
        except subprocess.TimeoutExpired:
            pass


def _run_child(leg: Mapping[str, Any], timeout: float, seen_processes: set[tuple[int, int]]) -> dict[str, Any]:
    command = leg["command"]
    try:
        process = subprocess.Popen(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            start_new_session=os.name == "posix",
        )
    except OSError as error:
        raise RunnerError(f"{leg['label']} could not start: {error}") from error
    try:
        start_time = _read_start_time_ticks(process.pid)
        process_executable = _read_process_executable(process.pid)
    except RunnerError:
        _terminate_process(process)
        raise
    identity = (process.pid, start_time)
    if identity in seen_processes:
        _terminate_process(process)
        raise RunnerError(f"{leg['label']} reused a prior process identity")
    seen_processes.add(identity)
    selector = selectors.DefaultSelector()
    streams: dict[int, bytearray] = {}
    stdout_fd: int | None = None
    stderr_fd: int | None = None
    for stream in (process.stdout, process.stderr):
        if stream is None:
            raise RunnerError(f"{leg['label']} did not expose bounded child output")
        selector.register(stream, selectors.EVENT_READ)
        streams[stream.fileno()] = bytearray()
    if process.stdout is not None:
        stdout_fd = process.stdout.fileno()
    if process.stderr is not None:
        stderr_fd = process.stderr.fileno()
    deadline = time.monotonic() + timeout
    try:
        while selector.get_map():
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise RunnerError(f"{leg['label']} exceeded the process timeout")
            for key, _events in selector.select(min(remaining, 0.1)):
                chunk = os.read(key.fd, 64 * 1024)
                if not chunk:
                    selector.unregister(key.fileobj)
                    continue
                buffer = streams[key.fd]
                buffer.extend(chunk)
                if len(buffer) > MAX_LOG_BYTES:
                    raise RunnerError(f"{leg['label']} exceeded the bounded child-output limit")
        return_code = process.wait(timeout=max(0.0, deadline - time.monotonic()))
    except (OSError, subprocess.TimeoutExpired, RunnerError):
        _terminate_process(process)
        raise
    finally:
        selector.close()
        stdout = bytes(streams.get(stdout_fd, ()))
        stderr = bytes(streams.get(stderr_fd, ()))
        if process.stdout is not None:
            process.stdout.close()
        if process.stderr is not None:
            process.stderr.close()
    if return_code != 0:
        raise RunnerError(f"{leg['label']} exited with status {return_code}")
    return {
        "label": leg["label"],
        "pid": process.pid,
        "start_time_ticks": start_time,
        "command": list(command),
        "command_executable": leg["command_executable"],
        "process_executable": process_executable,
        "stdout_bytes": len(stdout),
        "stdout_sha256": _sha256(stdout),
        "stderr_bytes": len(stderr),
        "stderr_sha256": _sha256(stderr),
    }


def _load_report(path: Path, label: str, mode: str) -> tuple[Mapping[str, Any], dict[str, Any], bytes]:
    try:
        metadata = path.lstat()
    except OSError as error:
        raise RunnerError(f"{label} report is missing: {error}") from error
    if stat.S_ISLNK(metadata.st_mode) or not stat.S_ISREG(metadata.st_mode):
        raise RunnerError(f"{label} report is not a regular non-symlink file")
    if metadata.st_size <= 0 or metadata.st_size > MAX_REPORT_BYTES:
        raise RunnerError(f"{label} report size is outside the finite protocol bound")
    try:
        data = path.read_bytes()
        report = json.loads(data.decode("utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"{label} report is not valid UTF-8 JSON: {error}") from error
    report = _mapping(report, f"{label} report")
    try:
        validated = validator.validate_report(report, label, mode)
    except validator.ValidationError as error:
        raise RunnerError(f"{label} report failed schema validation: {error}") from error
    return report, validated, data


def run_manifest(manifest_path: Path, receipt_path: Path) -> dict[str, Any]:
    _absolute_path(str(manifest_path), "manifest")
    receipt_path = _absolute_path(str(receipt_path), "receipt")
    if receipt_path.exists():
        raise RunnerError("receipt path already exists")
    if not receipt_path.parent.is_dir():
        raise RunnerError("receipt parent must already be a directory")
    manifest, output_dir, timeout, legs = _load_manifest(manifest_path)
    if output_dir.exists():
        raise RunnerError("output_dir must not exist before the fresh process matrix")
    output_dir.mkdir()
    seen_processes: set[tuple[int, int]] = set()
    process_receipts: list[dict[str, Any]] = []
    reports: dict[str, Mapping[str, Any]] = {}
    try:
        for leg in legs:
            process_receipt = _run_child(leg, timeout, seen_processes)
            report, validated, data = _load_report(leg["report"], leg["label"], manifest["mode"])
            process_receipt.update(
                {
                    "report_path": str(leg["report"]),
                    "report_bytes": len(data),
                    "report_sha256": _sha256(data),
                    "source_revision": validated["revision"],
                    "binary_identity": validated["binary"],
                }
            )
            process_receipts.append(process_receipt)
            reports[leg["label"]] = report
        expected_paths = {leg["report"] for leg in legs}
        actual_paths = set(output_dir.iterdir())
        if actual_paths != expected_paths:
            raise RunnerError("fresh output directory contains an unexpected file or missing report")
        report_hashes = [receipt["report_sha256"] for receipt in process_receipts]
        if len(set(report_hashes)) != len(report_hashes):
            raise RunnerError("fresh ABBA legs reused identical report content")
        try:
            summary = validator.validate_abba(reports, manifest["mode"])
        except validator.ValidationError as error:
            raise RunnerError(f"four-process ABBA matrix failed validation: {error}") from error
        if summary["rounds"] < validator.MIN_ROUNDS or summary["samples_per_leg"] < validator.REQUIRED_SAMPLE_COUNT:
            raise RunnerError("validated process matrix is shorter than the finite ABBA minimum")
        receipt = {
            "schema": RECEIPT_SCHEMA,
            "mode": manifest["mode"],
            "leg_order": list(LEG_ORDER),
            "process_count": len(process_receipts),
            "output_dir": str(output_dir),
            "processes": process_receipts,
            "summary": summary,
        }
        _exact_keys(receipt, RECEIPT_KEYS, "receipt")
        for index, process_receipt in enumerate(process_receipts):
            _exact_keys(process_receipt, PROCESS_KEYS, f"receipt.processes[{index}]")
        try:
            with receipt_path.open("x", encoding="utf-8") as handle:
                json.dump(receipt, handle, indent=2, sort_keys=True)
                handle.write("\n")
        except OSError as error:
            raise RunnerError(f"cannot write receipt: {error}") from error
        return receipt
    except Exception:
        # The fresh output is intentionally left intact for failure diagnosis;
        # the caller must not mistake a failed matrix for an accepted receipt.
        raise


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", required=True, type=Path)
    parser.add_argument("--receipt", required=True, type=Path)
    arguments = parser.parse_args(argv)
    try:
        run_manifest(arguments.manifest, arguments.receipt)
    except (OSError, RunnerError) as error:
        print(f"pptx notes memo process ABBA failed: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
