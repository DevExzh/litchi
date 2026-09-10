#!/usr/bin/env python3
"""Recover call stacks from the retained 0494 ``perf.data`` artifact.

The original profile run already completed the workload and produced a valid
owned-source ``perf.data`` file.  Its first ``perf script`` export failed only
because the event did not contain a CPU attribute while the command requested
the ``cpu`` output field.  This helper authenticates the original profile,
build, gate, protocol, report, and raw data receipts, then reprocesses that
existing file with a field list that omits ``cpu``.  It never launches the DOCX
workload or mutates the original profiling directory.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import json
import os
from pathlib import Path
import signal
import subprocess
import sys
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))
from support import ENV, REPO, ROOT, meta, sha, snapshot  # noqa: E402
import measure as canonical_measure  # noqa: E402
import profile as profiler  # noqa: E402


SCHEMA = "docx-edit-provider-profile-recovery-v1"
VERSION = 1
PROCESS_STARTED_SCHEMA = "docx-edit-provider-profile-recovery-process-start-v1"
PROCESS_TERMINAL_SCHEMA = "docx-edit-provider-profile-recovery-process-terminal-v1"
PROFILE_SCHEMA = "docx-edit-provider-profile-v1"
PROFILE_ATTEMPT = "profile-r1"
RECOVERY_LABEL = "owned-perf-script-no-cpu"
SCRIPT_FIELDS = "comm,pid,tid,time,period,event,ip,sym,dso"


class RecoveryError(RuntimeError):
    """A fail-closed retained-profile recovery error."""


def fail(message: str) -> None:
    raise RecoveryError(message)


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")


def _write_new(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except FileExistsError:
        fail(f"refusing to replace immutable recovery artifact: {path}")


def _file_meta(path: Path) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink():
        fail(f"{path}: expected a regular non-symlink file")
    return {"path": str(path), **meta(path)}


def _same_meta(actual: dict[str, Any], expected: dict[str, Any], label: str) -> None:
    for key in ("path", "bytes", "sha256"):
        if actual.get(key) != expected.get(key):
            fail(f"{label}: {key} differs")


def _same_digest(actual: dict[str, Any], expected: dict[str, Any], label: str) -> None:
    for key in ("bytes", "sha256"):
        if actual.get(key) != expected.get(key):
            fail(f"{label}: {key} differs")


def _require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def script_command(data: Path, cpu: int) -> list[str]:
    """Build the recovery command; ``cpu`` is intentionally absent."""

    return [
        "/usr/bin/taskset", "-c", str(cpu), "/usr/bin/perf", "script",
        "--header", "--demangle", "-F", SCRIPT_FIELDS, "-i", str(data),
    ]


def _profile_gate(path: Path, archived_profile: dict[str, Any]) -> dict[str, Any]:
    gate = _json(path)
    _require(isinstance(gate, dict), f"{path}: gate receipt is not an object")
    _require(gate.get("schema") == "docx-edit-provider-gate-v1",
             f"{path}: gate schema differs")
    _require(gate.get("driver_sha256") == sha(ROOT / "gate.py"),
             f"{path}: gate helper hash differs")
    _require(gate.get("common_sha256") == sha(ROOT / "support.py"),
             f"{path}: support helper hash differs")
    argv = gate.get("argv")
    _require(isinstance(argv, list) and all(isinstance(item, str) for item in argv),
             f"{path}: original profile command is missing")
    profile_argv = []
    for item in argv:
        candidate = Path(item)
        if not candidate.is_absolute():
            candidate = REPO / candidate
        if candidate.resolve() == (ROOT / "profile.py").resolve():
            profile_argv.append(item)
    _require(profile_argv, f"{path}: original profile helper is absent from gate argv")
    _require(gate.get("source_unchanged") is True,
             f"{path}: original profile gate changed source")
    _require(gate.get("source_before") == gate.get("source_after"),
             f"{path}: original profile gate source manifests differ")
    _require(isinstance(gate.get("exit_code"), int),
             f"{path}: original profile gate exit code is missing")
    return {
        "path": str(path),
        **meta(path),
        "exit_code": gate["exit_code"],
        "source_unchanged": gate["source_unchanged"],
        "driver_sha256": gate["driver_sha256"],
        "common_sha256": gate["common_sha256"],
        "profile_helper_sha256": archived_profile["sha256"],
        "argv": argv,
    }


def _helper_archive(path: Path) -> dict[str, Any]:
    custody_path = path / "helper-test-custody.json"
    custody = _json(custody_path)
    _require(isinstance(custody, dict), f"{custody_path}: custody is not an object")
    helpers = custody.get("helpers")
    _require(isinstance(helpers, dict), f"{custody_path}: helper map is missing")
    for name in ("profile.py", "test_profile.py"):
        expected = helpers.get(name)
        _require(isinstance(expected, dict), f"{custody_path}: {name} binding is missing")
        archived = _file_meta(path / name)
        _same_meta(archived, {"path": str(path / name), **expected},
                   f"{custody_path}: archived {name}")
        current = _file_meta(ROOT / name)
        _same_meta(current, {"path": str(ROOT / name), **expected},
                   f"current {name}: frozen helper differs from the original profile run")
    return {
        "custody": {"path": str(custody_path), **meta(custody_path)},
        "profile": _file_meta(path / "profile.py"),
        "test_profile": _file_meta(path / "test_profile.py"),
        "label": custody.get("label"),
    }


def _build_protocol(build_path: Path, summary: dict[str, Any]) -> tuple[
        dict[str, Any], dict[str, Any], dict[str, Any]]:
    try:
        builds = canonical_measure.load_builds(build_path.parent)
        protocol, protocol_hash = canonical_measure._load_protocol(builds)
    except (canonical_measure.ProviderMatrixError, OSError, ValueError) as error:
        fail(f"canonical build/protocol custody failed: {error}")
    normal = builds["normal"]
    _require(Path(normal["path"]).resolve() == build_path,
             "recovery build path is not the retained normal receipt")
    build_record = summary.get("build_record")
    _require(isinstance(build_record, dict), "profile summary build receipt is missing")
    _require(build_record.get("path") == str(build_path)
             and build_record.get("sha256") == normal["receipt_sha256"],
             "profile summary build receipt differs from canonical normal build")
    custody = summary.get("custody")
    _require(isinstance(custody, dict), "profile summary custody is missing")
    _require(custody.get("source") == normal["source"],
             "profile summary source binding differs from canonical build")
    protocol_binding = custody.get("protocol")
    _require(isinstance(protocol_binding, dict), "profile summary protocol binding is missing")
    _require(protocol_binding.get("path") == str(ROOT / "protocol.json")
             and protocol_binding.get("sha256") == protocol_hash,
             "profile summary protocol binding differs from frozen protocol")
    build_bindings = custody.get("builds")
    _require(isinstance(build_bindings, dict), "profile summary build bindings are missing")
    for role in canonical_measure.ROLES:
        binding = build_bindings.get(role)
        _require(isinstance(binding, dict), f"profile summary {role} binding is missing")
        _require(binding.get("receipt_sha256") == builds[role]["receipt_sha256"]
                 and binding.get("gate") == builds[role]["gate"]
                 and binding.get("source") == builds[role]["source"],
                 f"profile summary {role} binding differs from canonical receipt")
    _require(summary.get("source_revision") == normal["git_revision"],
             "profile summary source revision differs from canonical normal build")
    binary = summary.get("binary")
    _require(isinstance(binary, dict), "profile summary binary binding is missing")
    _require(binary.get("path") == normal["binary"]["path"]
             and binary.get("bytes") == normal["binary"]["bytes"]
             and binary.get("sha256") == normal["binary"]["sha256"],
             "profile summary binary differs from canonical normal build")
    return builds, protocol, {
        "protocol": {"path": str(ROOT / "protocol.json"), "sha256": protocol_hash,
                     "schema": protocol["schema"], "version": protocol["version"],
                     "change": protocol["change"], "case": protocol["case"]},
        "builds": {
            role: {"path": builds[role]["path"],
                   "receipt_sha256": builds[role]["receipt_sha256"],
                   "binary": builds[role]["binary"],
                   "source": builds[role]["source"],
                   "gate": builds[role]["gate"],
                   "git_revision": builds[role]["git_revision"]}
            for role in canonical_measure.ROLES
        },
        "source": normal["source"],
        "git_revision": normal["git_revision"],
    }


def _record_inputs(profile_dir: Path, summary: dict[str, Any], builds: dict[str, Any]) -> dict[str, Any]:
    record_dir = profile_dir / "owned" / "record"
    data_path = record_dir / "perf.data"
    report_path = record_dir / "perf-record-report.json"
    terminal_path = record_dir / "perf-record.terminal.json"
    failed_script_terminal_path = record_dir / "perf-script.terminal.json"
    data_meta = _file_meta(data_path)
    report_meta = _file_meta(report_path)
    terminal_meta = _file_meta(terminal_path)
    failed_script_terminal_meta = _file_meta(failed_script_terminal_path)
    terminal = _json(terminal_path)
    _require(isinstance(terminal, dict), f"{terminal_path}: terminal is not an object")
    _require(terminal.get("schema") == "docx-edit-provider-profile-process-terminal-v1"
             and terminal.get("role") == "owned-perf-record"
             and terminal.get("status") == "pass",
             f"{terminal_path}: retained perf-record terminal is not passing")
    process = terminal.get("process")
    _require(isinstance(process, dict) and process.get("exit_code") == 0
             and process.get("timed_out") is False and process.get("termination") is None,
             f"{terminal_path}: retained perf-record process is not successful")
    _require(terminal.get("source_unchanged") is True
             and terminal.get("source_before") == terminal.get("source_after"),
             f"{terminal_path}: retained perf-record source changed")
    terminal_artifacts = terminal.get("artifacts")
    _require(isinstance(terminal_artifacts, dict), f"{terminal_path}: artifact inventory is missing")
    _same_meta(terminal_artifacts.get("perf.data", {}), data_meta,
               f"{terminal_path}: perf.data artifact")
    _same_meta(terminal_artifacts.get("perf-record-report.json", {}), report_meta,
               f"{terminal_path}: report artifact")

    script_terminal = _json(failed_script_terminal_path)
    _require(isinstance(script_terminal, dict),
             f"{failed_script_terminal_path}: terminal is not an object")
    _require(script_terminal.get("schema") == "docx-edit-provider-profile-process-terminal-v1"
             and script_terminal.get("role") == "owned-perf-script"
             and script_terminal.get("status") == "failed",
             f"{failed_script_terminal_path}: original script failure is not retained")
    script_process = script_terminal.get("process")
    _require(isinstance(script_process, dict) and script_process.get("exit_code") != 0,
             f"{failed_script_terminal_path}: original script did not record its failure")

    owned = summary.get("owned_record")
    _require(isinstance(owned, dict), "profile summary owned record is missing")
    record = owned.get("record")
    _require(isinstance(record, dict), "profile summary perf-record details are missing")
    _require(record.get("samples") == 100 and record.get("warmup") == 3
             and record.get("frequency_hz") == 199
             and record.get("event") == "cycles:u"
             and record.get("callgraph") == "dwarf",
             "profile summary perf-record controls differ")
    _same_meta(record.get("data", {}), data_meta, "profile summary perf.data")
    _same_digest(record.get("report", {}), report_meta, "profile summary perf report")
    _require(record.get("process") == process,
             "profile summary perf-record process differs from terminal")
    _require(owned.get("status") == "failed" and "script" in owned,
             "profile summary does not retain the original script failure")

    report = canonical_measure.validate_report(
        report_path,
        role="normal",
        arm_name="owned",
        samples=record["samples"],
        warmups=record["warmup"],
        source_revision=builds["normal"]["git_revision"],
        binary_sha256=builds["normal"]["binary"]["sha256"],
        binary_bytes=builds["normal"]["binary"]["bytes"],
    )
    _require(report["provider"]["name"] == "owned", "retained report is not the owned arm")
    return {
        "perf_data": data_meta,
        "report": report_meta,
        "report_identity": {
            "schema": report["schema"], "version": report["version"],
            "case_name": report["case_name"], "rows": len(report["rows"]),
            "source_revision": report["source_revision"],
            "binary_sha256": report["binary_sha256"],
            "binary_bytes": report["binary_bytes"],
        },
        "record_terminal": terminal_meta,
        "record_terminal_status": terminal["status"],
        "record_process": process,
        "failed_script_terminal": failed_script_terminal_meta,
        "failed_script_status": script_terminal["status"],
        "failed_script_reason": script_terminal.get("reason"),
        "samples": record["samples"],
        "warmup": record["warmup"],
        "frequency_hz": record["frequency_hz"],
    }


def _run_script(argv: list[str], stdout_path: Path, stderr_path: Path,
                source_before: dict[str, Any], timeout: int) -> dict[str, Any]:
    start = _now()
    process: subprocess.Popen[bytes] | None = None
    timed_out = False
    termination: str | None = None
    launch_error: str | None = None
    try:
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            process = subprocess.Popen(
                argv,
                cwd=REPO,
                env=dict(ENV, PYTHONDONTWRITEBYTECODE="1"),
                stdin=subprocess.DEVNULL,
                stdout=stdout,
                stderr=stderr,
                start_new_session=True,
            )
            try:
                process.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                timed_out = True
                termination = "SIGTERM"
                try:
                    os.killpg(process.pid, signal.SIGTERM)
                except ProcessLookupError:
                    pass
                try:
                    process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    termination = "SIGKILL"
                    try:
                        os.killpg(process.pid, signal.SIGKILL)
                    except ProcessLookupError:
                        pass
                    process.wait()
    except OSError as error:
        launch_error = f"{type(error).__name__}: {error}"
    return {
        "argv": argv,
        "started_utc": start,
        "finished_utc": _now(),
        "pid": process.pid if process is not None else None,
        "process_group_id": process.pid if process is not None else None,
        "new_session": True,
        "exit_code": process.returncode if process is not None else None,
        "timed_out": timed_out,
        "termination": termination,
        "launch_error": launch_error,
    }


def run(args: argparse.Namespace) -> Path:
    if args.cpu < 0 or args.timeout < 1:
        fail("cpu must be non-negative and timeout must be positive")
    profile_dir = args.profile_dir.resolve()
    output = args.output_dir.resolve()
    if output.exists():
        fail(f"refusing to reuse recovery output directory: {output}")
    if not shutil_which("/usr/bin/taskset"):
        fail("/usr/bin/taskset is unavailable")
    if not shutil_which("/usr/bin/perf"):
        fail("/usr/bin/perf is unavailable")
    summary_path = profile_dir / "profile-summary.json"
    summary_meta = _file_meta(summary_path)
    summary = _json(summary_path)
    _require(isinstance(summary, dict), f"{summary_path}: profile summary is not an object")
    _require(summary.get("schema") == PROFILE_SCHEMA and summary.get("version") == VERSION,
             f"{summary_path}: profile summary schema differs")
    build_path = args.build_record.resolve()
    builds, protocol, custody = _build_protocol(build_path, summary)
    archive = _helper_archive(args.helper_archive.resolve())
    original_gate = _profile_gate(args.gate_receipt.resolve(), archive["profile"])
    record_inputs = _record_inputs(profile_dir, summary, builds)
    recovery_helper = _file_meta(Path(__file__).resolve())
    source_before = snapshot()
    canonical_source_before = canonical_measure._normalized_snapshot()
    source_matches_retained = canonical_source_before == custody["source"]
    output.mkdir(parents=True, exist_ok=False)
    data_path = profile_dir / "owned" / "record" / "perf.data"
    command = script_command(data_path, args.cpu)
    stdout_path = output / "perf-script.txt"
    stderr_path = output / "perf-script.stderr"
    started_path = output / "perf-script.started.json"
    terminal_path = output / "perf-script.terminal.json"
    _write_new(started_path, {
        "schema": PROCESS_STARTED_SCHEMA,
        "role": "owned-perf-script-recovery",
        "argv": command,
        "cwd": str(REPO),
        "environment": {key: ENV.get(key) for key in
                         ("RUSTUP_TOOLCHAIN", "CARGO_TARGET_DIR", "TMPDIR", "LC_ALL")},
        "new_session": True,
        "input": record_inputs["perf_data"],
        "recovery_helper": recovery_helper,
        "source_before": source_before,
        "started_utc": _now(),
        "scope": "reprocess retained perf.data only; no DOCX workload is launched",
    })
    process = _run_script(command, stdout_path, stderr_path, source_before, args.timeout)
    source_after = snapshot()
    canonical_source_after = canonical_measure._normalized_snapshot()
    source_unchanged_during_recovery = source_after == source_before
    canonical_source_unchanged_during_recovery = canonical_source_after == canonical_source_before
    status = "pass"
    reason: str | None = None
    parsed: dict[str, Any] | None = None
    stack_summary: dict[str, Any] | None = None
    folded_path = output / "perf-folded.txt"
    stack_summary_path = output / "perf-stack-summary.json"
    if process.get("launch_error"):
        status, reason = "unavailable", "perf script could not be launched"
    elif process.get("timed_out"):
        status, reason = "failed", "perf script timed out"
    elif process.get("exit_code") != 0:
        status, reason = "failed", "perf script exited nonzero"
    elif not stdout_path.is_file():
        status, reason = "failed", "perf script produced no export"
    else:
        parsed = profiler.parse_perf_script_text(stdout_path.read_text(encoding="utf-8", errors="replace"))
        if parsed["sample_count"] < 1:
            status, reason = "failed", "perf script exported no symbolized samples"
        else:
            stack_summary = profiler.summarize_perf_stacks(parsed)
            stack_summary["report_identity"] = {
                **record_inputs["report_identity"],
                "provider": "owned",
                "samples": record_inputs["samples"],
                "warmup": record_inputs["warmup"],
                "timing_scope": canonical_measure.TIMING_SCOPE,
                "source_archive_sha256": canonical_measure.CORPUS["archive_sha256"],
                "source_archive_bytes": canonical_measure.CORPUS["archive_bytes"],
            }
            folded_path.open("x", encoding="utf-8").writelines(
                f"{item['stack']} {item['period']}\n"
                for item in stack_summary["folded_stacks"]
            )
            _write_new(stack_summary_path, stack_summary)
    terminal = {
        "schema": PROCESS_TERMINAL_SCHEMA,
        "role": "owned-perf-script-recovery",
        "status": status,
        "reason": reason,
        "process": process,
        "source_before": source_before,
        "source_after": source_after,
        "source_unchanged": source_before == source_after,
        "finished_utc": _now(),
        "artifacts": {
            path.name: _file_meta(path)
            for path in (stdout_path, stderr_path)
            if path.is_file()
        },
        "input": record_inputs["perf_data"],
    }
    _write_new(terminal_path, terminal)
    recovery_summary = {
        "schema": SCHEMA,
        "version": VERSION,
        "status": status,
        "reason": reason,
        "created_utc": _now(),
        "mode": "retained-perf-data-script-without-cpu-field",
        "source_attempt": {
            "profile_summary": summary_meta,
            "profile_summary_status": summary.get("owned_record", {}).get("status"),
            "record": record_inputs,
        },
        "custody": {
            "recovery_helper": recovery_helper,
            "helper_archive": archive,
            "original_gate": original_gate,
            "build_protocol": custody,
            "source_before": source_before,
            "source_after": source_after,
            "canonical_source_before": canonical_source_before,
            "canonical_source_after": canonical_source_after,
            "retained_source": custody["source"],
            "source_matches_retained": source_matches_retained,
            "source_unchanged_during_recovery": source_unchanged_during_recovery,
            "canonical_source_unchanged_during_recovery": canonical_source_unchanged_during_recovery,
        },
        "command": command,
        "process": process,
        "artifacts": {
            name: _file_meta(path)
            for name, path in (
                ("started", started_path), ("terminal", terminal_path),
                ("stdout", stdout_path), ("stderr", stderr_path),
                ("folded", folded_path), ("stack_summary", stack_summary_path),
            )
            if path.is_file()
        },
        "stack_summary": None if stack_summary is None else {
            "path": str(stack_summary_path),
            "bytes": stack_summary_path.stat().st_size,
            "sha256": sha(stack_summary_path),
            "sample_count": stack_summary["sample_count"],
            "total_period": stack_summary["total_period"],
        },
        "interpretation": (
            "This recovery reprocessed the retained perf.data file after the original "
            "perf script failed because its requested CPU field was absent. It adds no "
            "workload samples; periods remain statistical call-stack weights. The current "
            "checkout source is recorded separately from the source bound to the original "
            "workload, because recovery launches no workload. Source changes observed during "
            "postprocessing are retained as custody observations and are not attributed to "
            "this read-only helper."
        ),
    }
    summary_output = output / "recovery-summary.json"
    _write_new(summary_output, recovery_summary)
    if status != "pass":
        fail(reason or "perf script recovery failed")
    return summary_output


def shutil_which(path: str) -> str | None:
    """Resolve an absolute tool path without importing the shell environment."""

    candidate = Path(path)
    return str(candidate) if candidate.is_file() and os.access(candidate, os.X_OK) else None


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--profile-dir", type=Path, default=ROOT / "profiling-r1")
    parser.add_argument("--build-record", type=Path, default=ROOT / "build-normal.json")
    parser.add_argument("--gate-receipt", type=Path, default=ROOT / "validation" / "profile-r1.json")
    parser.add_argument("--helper-archive", type=Path,
                        default=ROOT / "profiling-r1-helper-sources")
    parser.add_argument("--cpu", type=int, default=2)
    parser.add_argument("--timeout", type=int, default=600)
    parser.add_argument("--output-dir", type=Path, required=True)
    return parser


def main(argv: list[str] | None = None) -> int:
    try:
        args = _parser().parse_args(argv)
        path = run(args)
    except (RecoveryError, OSError, ValueError, subprocess.SubprocessError,
            canonical_measure.ProviderMatrixError) as error:
        print(f"recover_profile.py: FAIL: {error}", file=sys.stderr)
        return 1
    print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
