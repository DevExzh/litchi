#!/usr/bin/env python3
"""Run the PPTX notes-memo probe in repeated fresh processes and validate ABBA.

The local probe schema is intentionally not accepted by the default ABBA
summary tool.  This runner owns the process boundary: it admits an explicit
matrix of at least four fresh ``A1, B1, B2, A2`` rounds, gives every leg a
fresh report path, binds each command to an expected source and executable
identity, captures bounded process identities, and invokes the dedicated
schema validator only after every process finishes.
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


MANIFEST_SCHEMA = "pptx_notes_memo_abba_process_manifest_v2"
RECEIPT_SCHEMA = "pptx_notes_memo_abba_process_receipt_v2"
LEG_ORDER = validator.LEG_ORDER
MANIFEST_KEYS = frozenset({"schema", "mode", "output_dir", "timeout_seconds", "rounds"})
ROUND_KEYS = frozenset({"index", "legs"})
LEG_KEYS = frozenset({"label", "command", "report", "source_revision", "target_binary"})
TARGET_BINARY_KEYS = frozenset({"path", "sha256", "bytes", "mode_bits"})
RECEIPT_KEYS = frozenset(
    {
        "schema",
        "mode",
        "manifest_path",
        "manifest_bytes",
        "manifest_sha256",
        "leg_order",
        "process_round_count",
        "process_count",
        "output_dir",
        "processes",
        "round_summaries",
        "matrix",
    }
)
PROCESS_KEYS = frozenset(
    {
        "round",
        "label",
        "pid",
        "start_time_ticks",
        "command",
        "command_executable",
        "target_binary",
        "initial_process_executable",
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
ROUND_SUMMARY_KEYS = frozenset({"index", "summary"})
MATRIX_KEYS = frozenset(
    {
        "process_rounds",
        "legs_per_round",
        "processes",
        "sample_rounds_per_report",
        "samples_per_leg",
        "fixture",
    }
)
MAX_REPORT_BYTES = 8 * 1024 * 1024
MAX_LOG_BYTES = 1 * 1024 * 1024
MAX_TARGET_BINARY_BYTES = 512 * 1024 * 1024
MIN_TIMEOUT_SECONDS = 1.0
MAX_TIMEOUT_SECONDS = 600.0
MIN_PROCESS_ROUNDS = 4
MAX_PROCESS_ROUNDS = 64
TARGET_EXECUTION_WAIT_SECONDS = 2.0


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
        metadata = executable.lstat()
    except OSError as error:
        raise RunnerError(f"{path}[0] cannot be stat'ed: {error}") from error
    if stat.S_ISLNK(metadata.st_mode) or not stat.S_ISREG(metadata.st_mode):
        raise RunnerError(f"{path}[0] is not an executable regular file")
    if not metadata.st_mode & (stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH):
        raise RunnerError(f"{path}[0] has no execute permission bits")
    return str(executable.resolve())


def _file_identity(path: Path, label: str) -> dict[str, Any]:
    """Read a bounded, non-symlink executable identity and detect a read race."""

    try:
        before = path.lstat()
    except OSError as error:
        raise RunnerError(f"{label} cannot be stat'ed: {error}") from error
    if stat.S_ISLNK(before.st_mode) or not stat.S_ISREG(before.st_mode):
        raise RunnerError(f"{label} must be a regular non-symlink file")
    if before.st_size <= 0 or before.st_size > MAX_TARGET_BINARY_BYTES:
        raise RunnerError(f"{label} size is outside the finite protocol bound")
    mode_bits = stat.S_IMODE(before.st_mode)
    if not mode_bits & (stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH):
        raise RunnerError(f"{label} has no execute permission bits")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as handle:
            while chunk := handle.read(1024 * 1024):
                digest.update(chunk)
        after = path.lstat()
    except OSError as error:
        raise RunnerError(f"{label} cannot be read: {error}") from error
    if (
        before.st_dev,
        before.st_ino,
        before.st_size,
        before.st_mtime_ns,
        stat.S_IMODE(before.st_mode),
    ) != (
        after.st_dev,
        after.st_ino,
        after.st_size,
        after.st_mtime_ns,
        stat.S_IMODE(after.st_mode),
    ):
        raise RunnerError(f"{label} changed while its identity was being read")
    return {
        "path": str(path.resolve()),
        "sha256": digest.hexdigest(),
        "bytes": before.st_size,
        "mode_bits": mode_bits,
    }


def _validate_target_binary(value: Any, path: str) -> tuple[Path, dict[str, Any]]:
    target = _mapping(value, path)
    _exact_keys(target, TARGET_BINARY_KEYS, path)
    target_path = _absolute_path(target.get("path"), f"{path}.path")
    if str(target_path.resolve()) != str(target_path):
        raise RunnerError(f"{path}.path must be a canonical absolute path")
    digest = _string(target.get("sha256"), f"{path}.sha256")
    if validator.SHA256.fullmatch(digest) is None:
        raise RunnerError(f"{path}.sha256 has the wrong identity format")
    size = target.get("bytes")
    if isinstance(size, bool) or not isinstance(size, int) or not 0 < size <= MAX_TARGET_BINARY_BYTES:
        raise RunnerError(f"{path}.bytes is outside the finite protocol bound")
    mode_bits = target.get("mode_bits")
    if isinstance(mode_bits, bool) or not isinstance(mode_bits, int) or not 0 <= mode_bits <= 0o7777:
        raise RunnerError(f"{path}.mode_bits is outside the finite permission bound")
    if not mode_bits & 0o111:
        raise RunnerError(f"{path}.mode_bits has no execute permission bits")
    expected = {"path": str(target_path), "sha256": digest, "bytes": size, "mode_bits": mode_bits}
    actual = _file_identity(target_path, path)
    if actual != expected:
        raise RunnerError(f"{path} does not match the current target executable identity")
    return target_path, expected


def _load_manifest(path: Path) -> tuple[dict[str, Any], Path, float, list[dict[str, Any]], bytes]:
    try:
        raw = path.read_bytes()
        document = validator.load_strict_json(raw.decode("utf-8"), str(path))
    except (OSError, UnicodeError, validator.ValidationError) as error:
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
    rounds_value = manifest.get("rounds")
    if not isinstance(rounds_value, list) or not MIN_PROCESS_ROUNDS <= len(rounds_value) <= MAX_PROCESS_ROUNDS:
        raise RunnerError(
            f"{path}.rounds must contain between {MIN_PROCESS_ROUNDS} and {MAX_PROCESS_ROUNDS} explicit process rounds"
        )
    rounds: list[dict[str, Any]] = []
    seen_reports: set[Path] = set()
    expected_by_label: dict[str, tuple[str, dict[str, Any], str]] = {}
    for round_position, round_value in enumerate(rounds_value, start=1):
        round_path = f"{path}.rounds[{round_position - 1}]"
        round_mapping = _mapping(round_value, round_path)
        _exact_keys(round_mapping, ROUND_KEYS, round_path)
        index = round_mapping.get("index")
        if isinstance(index, bool) or not isinstance(index, int) or index != round_position:
            raise RunnerError(f"{round_path}.index must be the one-based manifest round index")
        legs_value = round_mapping.get("legs")
        if not isinstance(legs_value, list) or len(legs_value) != len(LEG_ORDER):
            raise RunnerError(f"{round_path}.legs must contain exactly four process legs")
        legs: list[dict[str, Any]] = []
        for leg_position, leg_value in enumerate(legs_value):
            leg_path = f"{round_path}.legs[{leg_position}]"
            leg = _mapping(leg_value, leg_path)
            _exact_keys(leg, LEG_KEYS, leg_path)
            label = leg.get("label")
            if label != LEG_ORDER[leg_position]:
                raise RunnerError(f"{leg_path}.label must be {LEG_ORDER[leg_position]}")
            command_value = leg.get("command")
            if not isinstance(command_value, list) or not command_value or not all(
                isinstance(argument, str) and argument for argument in command_value
            ):
                raise RunnerError(f"{leg_path}.command must be a non-empty string array")
            command = list(command_value)
            command_executable = _resolve_command_executable(command, f"{leg_path}.command")
            source_revision = _string(leg.get("source_revision"), f"{leg_path}.source_revision")
            if validator.REVISION.fullmatch(source_revision) is None:
                raise RunnerError(f"{leg_path}.source_revision has the wrong identity format")
            target_path, target_binary = _validate_target_binary(leg.get("target_binary"), f"{leg_path}.target_binary")
            report = _absolute_path(leg.get("report"), f"{leg_path}.report")
            if report.parent.resolve() != output_dir.resolve():
                raise RunnerError(f"{leg_path}.report must be directly under output_dir")
            if report in seen_reports:
                raise RunnerError("process report paths must be unique across all rounds")
            seen_reports.add(report)
            prior = expected_by_label.get(label)
            binding = (source_revision, target_binary, command_executable)
            if prior is not None and prior != binding:
                raise RunnerError(f"{leg_path} changes the source/target/command binding for {label}")
            expected_by_label[label] = binding
            legs.append(
                {
                    "label": label,
                    "command": command,
                    "command_executable": command_executable,
                    "source_revision": source_revision,
                    "target_path": target_path,
                    "target_binary": target_binary,
                    "report": report,
                }
            )
        rounds.append({"index": index, "legs": legs})
    return {"mode": mode, "manifest_path": path.resolve()}, output_dir, float(timeout), rounds, raw


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


def _observe_target_executable(process: subprocess.Popen[bytes], target_path: Path, label: str) -> tuple[str, str]:
    """Observe target identity in the same PID, permitting direct execution or exec wrappers."""

    initial = _read_process_executable(process.pid)
    expected = str(target_path)
    if initial == expected:
        return initial, initial
    deadline = time.monotonic() + TARGET_EXECUTION_WAIT_SECONDS
    while process.poll() is None and time.monotonic() < deadline:
        try:
            observed = _read_process_executable(process.pid)
        except RunnerError:
            time.sleep(0.01)
            continue
        if observed == expected:
            return initial, observed
        time.sleep(0.01)
    raise RunnerError(f"{label} did not exec the expected target binary in the same PID")


def _run_child(
    round_index: int,
    leg: Mapping[str, Any],
    timeout: float,
    seen_processes: set[tuple[int, int]],
) -> dict[str, Any]:
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
        raise RunnerError(f"{leg['label']} round {round_index} could not start: {error}") from error
    try:
        start_time = _read_start_time_ticks(process.pid)
        initial_executable, process_executable = _observe_target_executable(
            process, leg["target_path"], f"{leg['label']} round {round_index}"
        )
    except RunnerError:
        _terminate_process(process)
        raise
    identity = (process.pid, start_time)
    if identity in seen_processes:
        _terminate_process(process)
        raise RunnerError(f"{leg['label']} round {round_index} reused a prior process identity")
    seen_processes.add(identity)
    selector = selectors.DefaultSelector()
    streams: dict[int, bytearray] = {}
    stdout_fd: int | None = None
    stderr_fd: int | None = None
    try:
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
        while selector.get_map():
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise RunnerError(f"{leg['label']} round {round_index} exceeded the process timeout")
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
        raise RunnerError(f"{leg['label']} round {round_index} exited with status {return_code}")
    actual_target = _file_identity(leg["target_path"], f"{leg['label']} round {round_index} target binary")
    if actual_target != leg["target_binary"]:
        raise RunnerError(f"{leg['label']} round {round_index} target binary changed after execution")
    return {
        "round": round_index,
        "label": leg["label"],
        "pid": process.pid,
        "start_time_ticks": start_time,
        "command": list(command),
        "command_executable": leg["command_executable"],
        "target_binary": dict(leg["target_binary"]),
        "initial_process_executable": initial_executable,
        "process_executable": process_executable,
        "stdout_bytes": len(stdout),
        "stdout_sha256": _sha256(stdout),
        "stderr_bytes": len(stderr),
        "stderr_sha256": _sha256(stderr),
    }


def _load_report(
    path: Path,
    label: str,
    mode: str,
    expected_source_revision: str,
    expected_binary: Mapping[str, Any],
) -> tuple[Mapping[str, Any], dict[str, Any], bytes]:
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
        report = validator.load_strict_json(data.decode("utf-8"), str(path))
    except (OSError, UnicodeError, validator.ValidationError) as error:
        raise RunnerError(f"{label} report is not strict UTF-8 JSON: {error}") from error
    report = _mapping(report, f"{label} report")
    try:
        validated = validator.validate_report(report, label, mode)
    except validator.ValidationError as error:
        raise RunnerError(f"{label} report failed schema validation: {error}") from error
    if validated["revision"] != expected_source_revision:
        raise RunnerError(f"{label} report source revision does not match its manifest leg")
    expected_report_binary = {
        "path": expected_binary["path"],
        "binary_sha256": expected_binary["sha256"],
        "binary_bytes": expected_binary["bytes"],
        "mode_bits": expected_binary["mode_bits"],
        "executable": True,
        "profile": "release",
    }
    if validated["binary"] != expected_report_binary:
        raise RunnerError(f"{label} report binary identity does not match its manifest leg")
    return report, validated, data


def _summary_fingerprint(summary: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "fixture": summary["fixture"],
        "warmup": summary["warmup"],
        "samples_per_leg": summary["samples_per_leg"],
        "source_revisions": summary["source_revisions"],
        "binary_identities": summary["binary_identities"],
        "rows": [
            {
                key: row[key]
                for key in (
                    "case",
                    "output_identity",
                    "source_inventory",
                    "fidelity_evidence",
                    "fidelity_gate",
                    "allocation_metric_gate",
                )
            }
            for row in summary["rows"]
        ],
    }


def run_manifest(manifest_path: Path, receipt_path: Path) -> dict[str, Any]:
    manifest_path = _absolute_path(str(manifest_path), "manifest").resolve()
    receipt_path = _absolute_path(str(receipt_path), "receipt")
    if receipt_path.exists():
        raise RunnerError("receipt path already exists")
    if not receipt_path.parent.is_dir():
        raise RunnerError("receipt parent must already be a directory")
    manifest, output_dir, timeout, rounds, manifest_bytes = _load_manifest(manifest_path)
    if output_dir.exists():
        raise RunnerError("output_dir must not exist before the fresh process matrix")
    output_dir.mkdir()
    seen_processes: set[tuple[int, int]] = set()
    process_receipts: list[dict[str, Any]] = []
    round_summaries: list[dict[str, Any]] = []
    baseline_fingerprint: dict[str, Any] | None = None
    try:
        for round_value in rounds:
            round_index = round_value["index"]
            reports: dict[str, Mapping[str, Any]] = {}
            expected_paths: set[Path] = set()
            for leg in round_value["legs"]:
                expected_paths.add(leg["report"])
                process_receipt = _run_child(round_index, leg, timeout, seen_processes)
                report, validated, data = _load_report(
                    leg["report"],
                    f"round {round_index} {leg['label']}",
                    manifest["mode"],
                    leg["source_revision"],
                    leg["target_binary"],
                )
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
            actual_round_paths = set(output_dir.iterdir())
            if not expected_paths <= actual_round_paths:
                raise RunnerError(f"round {round_index} did not produce every expected report")
            try:
                summary = validator.validate_abba(reports, manifest["mode"])
            except validator.ValidationError as error:
                raise RunnerError(f"process ABBA round {round_index} failed validation: {error}") from error
            if summary["rounds"] < validator.MIN_ROUNDS or summary["samples_per_leg"] < validator.REQUIRED_SAMPLE_COUNT:
                raise RunnerError(f"process ABBA round {round_index} has too few in-process samples")
            fingerprint = _summary_fingerprint(summary)
            if baseline_fingerprint is None:
                baseline_fingerprint = fingerprint
            elif fingerprint != baseline_fingerprint:
                raise RunnerError("ABBA process rounds disagree on fixture, binary, source, or fidelity identity")
            round_summaries.append({"index": round_index, "summary": summary})
        expected_paths = {
            leg["report"]
            for round_value in rounds
            for leg in round_value["legs"]
        }
        actual_paths = set(output_dir.iterdir())
        if actual_paths != expected_paths:
            raise RunnerError("fresh output directory contains an unexpected file or missing report")
        report_hashes = [receipt["report_sha256"] for receipt in process_receipts]
        if len(set(report_hashes)) != len(report_hashes):
            raise RunnerError("fresh ABBA legs reused identical report content")
        if len(process_receipts) != len(rounds) * len(LEG_ORDER):
            raise RunnerError("process receipt count does not equal explicit round cardinality")
        first_summary = round_summaries[0]["summary"]
        matrix = {
            "process_rounds": len(rounds),
            "legs_per_round": len(LEG_ORDER),
            "processes": len(process_receipts),
            "sample_rounds_per_report": first_summary["rounds"],
            "samples_per_leg": first_summary["samples_per_leg"],
            "fixture": first_summary["fixture"],
        }
        receipt = {
            "schema": RECEIPT_SCHEMA,
            "mode": manifest["mode"],
            "manifest_path": str(manifest_path),
            "manifest_bytes": len(manifest_bytes),
            "manifest_sha256": _sha256(manifest_bytes),
            "leg_order": list(LEG_ORDER),
            "process_round_count": len(rounds),
            "process_count": len(process_receipts),
            "output_dir": str(output_dir),
            "processes": process_receipts,
            "round_summaries": round_summaries,
            "matrix": matrix,
        }
        _exact_keys(receipt, RECEIPT_KEYS, "receipt")
        _exact_keys(matrix, MATRIX_KEYS, "receipt.matrix")
        for index, process_receipt in enumerate(process_receipts):
            _exact_keys(process_receipt, PROCESS_KEYS, f"receipt.processes[{index}]")
        for index, round_summary in enumerate(round_summaries):
            _exact_keys(round_summary, ROUND_SUMMARY_KEYS, f"receipt.round_summaries[{index}]")
        try:
            with receipt_path.open("x", encoding="utf-8") as handle:
                json.dump(receipt, handle, indent=2, sort_keys=True, allow_nan=False)
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
