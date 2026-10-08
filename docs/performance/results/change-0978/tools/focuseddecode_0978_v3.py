#!/usr/bin/env python3
"""Offline decoder for one 0978 instruction-profile child.

The runner owns the native command and the immutable ``perf.data`` handoff.
This module only runs perf's offline readers, the sealed variable-layout raw
scanner, and the 0978 lifecycle ledger.  It never installs a runner patch,
starts a workload, or changes a retained helper.

The public entry point is ``decode_child(plan_path, child_dir)``.  The packet
keeps the old ``perf.corrected`` summary shape while putting the disjoint
sample ledger at ``attribution.ledger``.  Full raw, lifecycle, and corrected
sample identities are kept in sibling sidecars so reducers can reparse them.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import os
import re
import struct
import subprocess
import sys
from collections import Counter
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Mapping, Sequence


SCHEMA = "litchi.performance.0978.child_packet.v1"
RAW_SCHEMA = "litchi.performance.0978.raw_stack.v1"
LEDGER_SCHEMA = "litchi.performance.0978.lifecycle_ledger.v1"
ABI_SCHEMA = "litchi.performance.0978.abi_scan.v1"
ATTRIBUTION_SCHEMA = "litchi.performance.0978.attribution_samples.v1"
DECODER_SCHEMA = "litchi.performance.0978.decoder.v1"

PERF_PATH = Path("/usr/bin/perf")
DEFAULT_PERIOD = 1_000_000
DEFAULT_CPU = 31
DEFAULT_STACK_CAP = 32_768
DEFAULT_MARKER = "doc_retention_probe_0942_cfb_profile::profiled_measured_case"
DEFAULT_BUILD_ID = "e5c457384eeebd057932c444efda77618f137399"
FULL_LIFETIME_EVLIST_ATTRIBUTES = ("comm", "comm_exec", "mmap_data", "mmap2", "build_id")
TERMINAL_LOST_RE = re.compile(r"^#\s*Total Lost Samples:\s*(\d+)\s*$", re.MULTILINE)
LOST_EVENT_RE = re.compile(r"(?:^|\s)(?:lost|dropped)[^0-9]*(\d+)", re.IGNORECASE)
HEADER_RE = re.compile(
    r"^\s*(?:(?P<comm>\S+)\s+)?(?P<pid>\d+)(?:/(?P<tid>\d+))?"
    r"\s+\[(?P<cpu>\d+)\]\s+(?P<time>\d+(?:\.\d+)?):\s+"
    r"(?P<period>\d+)\s+(?P<event>[^\n]+):\s*$"
)
FRAME_RE = re.compile(r"^\s*(?P<ip>[0-9a-fA-F]+)\s+(?P<symbol>.*?)\s+\((?P<dso>.*)\)\s*$")
MMAP_RE = re.compile(
    r"PERF_RECORD_MMAP2\s+(?P<pid>\d+)/(?:\s*)?(?P<tid>\d+):\s+"
    r"\[(?P<start>0x[0-9a-fA-F]+)\((?P<length>0x[0-9a-fA-F]+)\)\s+"
    r"@\s+(?P<offset>(?:0x)?[0-9a-fA-F]+)\s+(?P<meta>[^\]]*)\]:\s+"
    r"(?P<perm>\S+)\s+(?P<path>.*)$"
)


class DecodeFailure(RuntimeError):
    """Typed fail-closed decoder error."""

    def __init__(self, code: str, message: str, details: Mapping[str, Any] | None = None):
        super().__init__(f"{code}: {message}")
        self.code = code
        self.message = message
        self.details = dict(details or {})

    def as_dict(self) -> dict[str, Any]:
        return {"status": "failed", "code": self.code, "message": self.message, "details": self.details}


def _fail(code: str, message: str, **details: Any) -> None:
    raise DecodeFailure(code, message, details)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    path = Path(path).absolute()
    if not path.is_file():
        _fail("input_missing", "bound file is missing", path=str(path))
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        _fail("invalid_json", "cannot load JSON", path=str(path), error=str(exc))
    if not isinstance(value, dict):
        _fail("invalid_json", "JSON root is not an object", path=str(path))
    return value


def write_exclusive(path: Path, value: Any) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as handle:
            json.dump(value, handle, indent=2, sort_keys=True, allow_nan=False)
            handle.write("\n")
    except FileExistsError:
        _fail("output_exists", "refusing to overwrite decoder output", path=str(path))


def load_module(name: str, path: Path) -> Any:
    path = Path(path).absolute()
    if not path.is_file():
        _fail("helper_missing", "bound helper is missing", path=str(path))
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        _fail("helper_import", "cannot create module spec", path=str(path))
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as exc:  # noqa: BLE001
        _fail("helper_import", "bound helper import failed", path=str(path), error=str(exc))
    return module


def _find_binding(plan: Mapping[str, Any], names: Sequence[str]) -> Mapping[str, Any] | None:
    containers: list[Mapping[str, Any]] = [plan]
    for key in ("dependencies", "helpers", "decoder", "marker", "perf", "binary"):
        value = plan.get(key)
        if isinstance(value, Mapping):
            containers.append(value)
    for container in containers:
        for name in names:
            value = container.get(name)
            if isinstance(value, Mapping) and isinstance(value.get("path"), str):
                return value
    return None


def _check_binding(row: Mapping[str, Any] | None, label: str, *, required: bool = True) -> Path | None:
    if row is None:
        if required:
            _fail("plan_binding", f"missing {label} binding")
        return None
    raw_path = row.get("path")
    if not isinstance(raw_path, str):
        _fail("plan_binding", f"{label} path is malformed")
    path = Path(raw_path).absolute()
    actual = binding(path)
    wanted = row.get("sha256")
    if wanted is not None and actual["sha256"] != wanted:
        _fail("plan_binding", f"{label} SHA-256 differs", actual=actual["sha256"], expected=wanted)
    if row.get("bytes") is not None and actual["bytes"] != row["bytes"]:
        _fail("plan_binding", f"{label} byte count differs", actual=actual["bytes"], expected=row["bytes"])
    return path


def _plan_data(plan_path: Path | Mapping[str, Any]) -> tuple[dict[str, Any], Path | None]:
    if isinstance(plan_path, Mapping):
        return dict(plan_path), Path(str(plan_path["plan_path"])).absolute() if isinstance(plan_path.get("plan_path"), str) else None
    path = Path(plan_path).absolute()
    return load_json(path), path


def _validate_plan(plan: Mapping[str, Any]) -> dict[str, Any]:
    schema = str(plan.get("schema", ""))
    if schema and not schema.startswith("litchi.performance.0978.plan"):
        _fail("plan_schema", "plan schema is not 0978", schema=schema)
    binary_row = plan.get("binary") if isinstance(plan.get("binary"), Mapping) else _find_binding(plan, ("binary", "target_binary"))
    binary = _check_binding(binary_row, "binary")
    assert binary is not None
    perf_row = plan.get("perf") if isinstance(plan.get("perf"), Mapping) else _find_binding(plan, ("perf",))
    perf_path = Path(str(perf_row.get("path", PERF_PATH))).absolute() if isinstance(perf_row, Mapping) else PERF_PATH
    if not perf_path.is_file():
        _fail("perf_binding", "perf executable is missing", path=str(perf_path))
    if isinstance(perf_row, Mapping) and perf_row.get("sha256") and sha256(perf_path) != perf_row["sha256"]:
        _fail("perf_binding", "perf executable SHA-256 differs", path=str(perf_path))
    perf_contract = plan.get("perf_contract") if isinstance(plan.get("perf_contract"), Mapping) else {}
    perf = plan.get("perf") if isinstance(plan.get("perf"), Mapping) else {}
    period = int(perf.get("period", perf_contract.get("period", DEFAULT_PERIOD)))
    cpu = int(perf.get("cpu", (plan.get("batch") or {}).get("cpu", DEFAULT_CPU)))
    stack_cap = int(perf.get("stack_cap", perf_contract.get("stack_cap", DEFAULT_STACK_CAP)))
    if period <= 0 or cpu < 0 or stack_cap <= 0:
        _fail("perf_contract", "period, CPU, or stack cap is invalid", period=period, cpu=cpu, stack_cap=stack_cap)
    marker_row = plan.get("marker") if isinstance(plan.get("marker"), Mapping) else {}
    marker = str(marker_row.get("symbol", DEFAULT_MARKER))
    marker_dso = str(marker_row.get("dso_path", binary))
    if Path(marker_dso).absolute() != binary.absolute():
        _fail("marker_binding", "marker DSO differs from bound binary", marker_dso=marker_dso, binary=str(binary))
    build_id = str(marker_row.get("build_id", plan.get("binary_build_id", DEFAULT_BUILD_ID))).lower()
    if not re.fullmatch(r"[0-9a-f]{8,64}", build_id):
        _fail("marker_binding", "target BuildID is malformed", build_id=build_id)
    focused_row = _find_binding(plan, ("focused_decode", "focused_decoder", "decoder_python"))
    focused_path = _check_binding(focused_row, "focused decoder", required=True)
    assert focused_path is not None and isinstance(focused_row, Mapping)
    if focused_path.absolute() != Path(__file__).absolute():
        _fail("plan_binding", "focused decoder binding is not the executing decoder file", bound=str(focused_path), executing=str(Path(__file__).absolute()))
    # The Python decoder is part of the formal helper closure.  Requiring both
    # fields here prevents a plan from naming a mutable path while omitting
    # the exact bytes that were reviewed.
    actual_focused = binding(focused_path)
    if focused_row.get("bytes") != actual_focused["bytes"] or focused_row.get("sha256") != actual_focused["sha256"]:
        _fail("plan_binding", "focused decoder binding lacks exact bytes/SHA", expected={"bytes": focused_row.get("bytes"), "sha256": focused_row.get("sha256")}, actual=actual_focused)
    decoder_so = _check_binding(_find_binding(plan, ("decoder_so", "shim_so", "shared_object", "so")), "decoder shared object", required=True)
    decoder_source = _check_binding(_find_binding(plan, ("decoder_source", "shim_source", "source")), "decoder source", required=True)
    decoder_map = _check_binding(_find_binding(plan, ("decoder_map", "shim_map", "version_script", "map")), "decoder map", required=True)
    scanner_path = _check_binding(_find_binding(plan, ("scanner",)), "raw scanner", required=True)
    lifecycle_path = _check_binding(_find_binding(plan, ("lifecycle",)), "lifecycle helper", required=True)
    return {
        "binary": binary,
        "binary_binding": dict(binary_row) if isinstance(binary_row, Mapping) else binding(binary),
        "perf": perf_path,
        "period": period,
        "cpu": cpu,
        "stack_cap": stack_cap,
        "marker": marker,
        "marker_dso": marker_dso,
        "build_id": build_id,
        "decoder_so": decoder_so,
        "decoder_source": decoder_source,
        "decoder_map": decoder_map,
        "focused_decode": focused_path,
        "analysis_only_replay": bool(plan.get("analysis_only_replay", False)),
        "scanner": scanner_path,
        "lifecycle": lifecycle_path,
    }


def _path(child_dir: Path, suffix: str) -> Path:
    return Path(child_dir) / f"{Path(child_dir).name}.{suffix}"


def _input(child_dir: Path, suffix: str, *, required: bool = True) -> Path | None:
    path = _path(child_dir, suffix)
    if required and not path.is_file():
        _fail("missing_capture_artifact", "required child artifact is missing", path=str(path))
    return path if path.is_file() else None


def _write_or_bind_existing(path: Path, value: Any | None = None) -> Path:
    if value is not None:
        write_exclusive(path, value)
    elif not path.is_file():
        _fail("output_missing", "decoder output is missing", path=str(path))
    return path


def _command_environment(env: Mapping[str, str]) -> dict[str, str]:
    return {str(key): str(value) for key, value in sorted(env.items())}


def _command_receipt(
    path: Path,
    *,
    kind: str,
    argv: Sequence[str],
    env: Mapping[str, str],
    output: Path,
    stderr: Path,
    executed: bool,
    reused_existing: bool,
    returncode: int,
    source: str,
) -> dict[str, Any]:
    value = {
        "schema": "litchi.performance.0978.offline_command.v1",
        "status": "complete",
        "kind": kind,
        "argv": list(argv),
        "environment": _command_environment(env),
        "output": binding(output),
        "stderr": binding(stderr),
        "executed": bool(executed),
        "reused_existing": bool(reused_existing),
        "returncode": int(returncode),
        "native_workload_started": False,
        "source": source,
    }
    write_exclusive(path, value)
    return value


def _validate_command_receipt(
    path: Path,
    *,
    kind: str,
    argv: Sequence[str],
    env: Mapping[str, str],
    output: Path,
    stderr: Path,
) -> dict[str, Any]:
    value = load_json(path)
    if value.get("schema") != "litchi.performance.0978.offline_command.v1" or value.get("kind") != kind:
        _fail("offline_command_receipt", "offline command receipt schema/kind differs", path=str(path), kind=kind)
    if value.get("argv") != list(argv) or value.get("environment") != _command_environment(env):
        _fail("offline_command_receipt", "offline command argv/environment differs", path=str(path), kind=kind)
    for label, expected in (("output", output), ("stderr", stderr)):
        row = value.get(label)
        if not isinstance(row, Mapping) or row.get("path") != str(expected.absolute()) or row.get("bytes") != expected.stat().st_size or row.get("sha256") != sha256(expected):
            _fail("offline_command_receipt", f"offline command {label} binding differs", path=str(path), kind=kind)
    if value.get("returncode") != 0 or value.get("native_workload_started") is not False:
        _fail("offline_command_receipt", "offline command receipt is not a successful offline command", path=str(path), kind=kind)
    return value


def _reuse_or_fail(
    cfg: Mapping[str, Any],
    receipt: Path,
    *,
    kind: str,
    argv: Sequence[str],
    env: Mapping[str, str],
    output: Path,
    stderr: Path,
) -> dict[str, Any]:
    if cfg.get("analysis_only_replay") is not True:
        _fail("offline_command_receipt", "strict decoder plan cannot consume pre-existing offline output", kind=kind, output=str(output))
    if receipt.is_file():
        value = _validate_command_receipt(receipt, kind=kind, argv=argv, env=env, output=output, stderr=stderr)
        return value
    return _command_receipt(receipt, kind=kind, argv=argv, env=env, output=output, stderr=stderr, executed=False, reused_existing=True, returncode=0, source="retained_offline_input")


def _run_to_file(argv: Sequence[str], env: Mapping[str, str], output: Path, stderr: Path, receipt: Path, *, kind: str) -> dict[str, Any]:
    if output.exists() or stderr.exists():
        _fail("output_exists", "refusing to overwrite offline command output", output=str(output), stderr=str(stderr))
    try:
        with output.open("xb") as out_handle, stderr.open("xb") as err_handle:
            result = subprocess.run(list(argv), stdin=subprocess.DEVNULL, stdout=out_handle, stderr=err_handle,
                                    env=dict(env), check=False, timeout=900)
    except (OSError, subprocess.TimeoutExpired) as exc:
        _fail("offline_command", "offline perf command failed to execute", argv=list(argv), error=str(exc))
    if result.returncode != 0:
        tail = stderr.read_text(encoding="utf-8", errors="replace")[-2000:]
        _fail("offline_command", "offline perf command returned nonzero", argv=list(argv), returncode=result.returncode, stderr_tail=tail)
    return _command_receipt(receipt, kind=kind, argv=argv, env=env, output=output, stderr=stderr, executed=True, reused_existing=False, returncode=result.returncode, source="decoder_offline_command")


def _base_env() -> dict[str, str]:
    # Offline perf readers must not inherit arbitrary session variables.  The
    # corrected shim environment adds only its explicit allowlisted keys.
    return {"LC_ALL": "C", "LANG": "C", "PATH": "/usr/bin:/bin", "PYTHONDONTWRITEBYTECODE": "1", "TZ": "UTC"}


def _load_launch(child_dir: Path, binary: Path) -> dict[str, Any]:
    launch = load_json(_input(child_dir, "launch.json"))
    try:
        pid = int(launch["exec_pid"])
    except (KeyError, TypeError, ValueError):
        _fail("launch_contract", "launch receipt lacks integer exec_pid")
    if pid <= 0 or int(launch.get("pid", pid)) != pid:
        _fail("launch_contract", "launch PID identity differs", launch=launch)
    if launch.get("packet_id", child_dir.name) != child_dir.name:
        _fail("launch_contract", "launch packet ID differs from child directory")
    if launch.get("execve") is False:
        _fail("launch_contract", "launch receipt records execve failure")
    if launch.get("exec_path") and Path(str(launch["exec_path"])).absolute() != binary.absolute():
        _fail("launch_contract", "launch executable differs from plan binary", launch_path=launch.get("exec_path"), binary=str(binary))
    return launch


def _load_terminal(child_dir: Path) -> dict[str, Any]:
    terminal = load_json(_input(child_dir, "terminal.json"))
    if terminal.get("returncode") not in (None, 0) or terminal.get("timed_out") is True or terminal.get("launch_error") is True:
        _fail("terminal_contract", "native terminal receipt failed", returncode=terminal.get("returncode"), timed_out=terminal.get("timed_out"), launch_error=terminal.get("launch_error"))
    if terminal.get("native_ld_preload") not in (None, "absent"):
        _fail("terminal_contract", "native terminal receipt contains LD_PRELOAD", value=terminal.get("native_ld_preload"))
    return terminal


def _terminal_report(child_dir: Path, cfg: Mapping[str, Any]) -> tuple[Path, dict[str, Any], int, dict[str, Any]]:
    perf_path = Path(cfg["perf"])
    data = _input(child_dir, "perf.data")
    report_path = _path(child_dir, "terminal-report.txt")
    stderr_path = _path(child_dir, "terminal-report.stderr")
    receipt_path = _path(child_dir, "terminal-report.command.json")
    argv = [str(perf_path), "report", "--stdio", "--percent-limit", "100", "-i", str(data)]
    env = _base_env()
    if report_path.exists() != stderr_path.exists():
        _fail("offline_command", "terminal report has only one output stream", report=str(report_path), stderr=str(stderr_path))
    if not report_path.is_file():
        receipt = _run_to_file(argv, env, report_path, stderr_path, receipt_path, kind="terminal_report")
    else:
        receipt = _reuse_or_fail(cfg, receipt_path, kind="terminal_report", argv=argv, env=env, output=report_path, stderr=stderr_path)
    text = report_path.read_text(encoding="utf-8", errors="strict")
    values = [int(value) for value in TERMINAL_LOST_RE.findall(text)]
    if len(values) != 1:
        _fail("terminal_loss", "terminal perf report must have exactly one Total Lost Samples line", count=len(values), path=str(report_path))
    summary = {"schema": "litchi.performance.0978.terminal_summary.v1", "report": binding(report_path), "lost_samples": values[0], "command": argv, "command_receipt": binding(receipt_path), "returncode": 0}
    write_exclusive(_path(child_dir, "terminal-summary.json"), summary)
    # The terminal report owns only PERF_RECORD_LOST_SAMPLES.  The final loss
    # sidecar is written after the raw scanner has independently counted
    # PERF_RECORD_LOST records; emitting a placeholder null here would make
    # custody consumers unable to distinguish "not scanned" from zero loss.
    return report_path, summary, values[0], receipt


def _perf_header(child_dir: Path, cfg: Mapping[str, Any]) -> tuple[Path, Path, dict[str, Any]]:
    """Retain perf report header stdout/stderr for each fresh child."""
    data = _input(child_dir, "perf.data")
    output = _path(child_dir, "perf-header.txt")
    stderr = _path(child_dir, "perf-header.stderr")
    receipt_path = _path(child_dir, "perf-header.command.json")
    argv = [str(cfg["perf"]), "report", "--header-only", "--stdio", "-i", str(data)]
    env = _base_env()
    if output.exists() != stderr.exists():
        _fail("offline_command", "perf header has only one output stream", output=str(output), stderr=str(stderr))
    if not output.is_file():
        receipt = _run_to_file(argv, env, output, stderr, receipt_path, kind="perf_header")
    else:
        receipt = _reuse_or_fail(cfg, receipt_path, kind="perf_header", argv=argv, env=env, output=output, stderr=stderr)
    return output, stderr, receipt


def _require_full_lifetime_evlist(contract: Mapping[str, Any]) -> None:
    """Require the event metadata needed to reconstruct process lifetime maps."""
    line = contract.get("line")
    if not isinstance(line, str) or not line:
        _fail("evlist_lifetime_contract", "perf evlist contract has no authoritative event line")
    required: dict[str, str] = {}
    for label in FULL_LIFETIME_EVLIST_ATTRIBUTES:
        values = re.findall(rf"(?:^|,\s*){re.escape(label)}:\s*([^,]*)", line)
        if len(values) != 1:
            _fail("evlist_lifetime_contract", "required lifetime event attribute is missing or duplicated", attribute=label, occurrences=len(values))
        value = values[0].strip()
        if value != "1":
            _fail("evlist_lifetime_contract", "required lifetime event attribute is not enabled", attribute=label, value=value)
        required[label] = value
    return None


def _evlist_and_dump(child_dir: Path, cfg: Mapping[str, Any]) -> tuple[dict[str, Any], Path]:
    data = _input(child_dir, "perf.data")
    evlist = _path(child_dir, "perf-evlist-v.txt")
    evlist_stderr = _path(child_dir, "perf-evlist-v.stderr")
    evlist_receipt = _path(child_dir, "perf-evlist-v.command.json")
    evlist_argv = [str(cfg["perf"]), "evlist", "-v", "-i", str(data)]
    env = _base_env()
    if evlist.exists() != evlist_stderr.exists():
        _fail("offline_command", "perf evlist has only one output stream", output=str(evlist), stderr=str(evlist_stderr))
    if not evlist.is_file():
        _run_to_file(evlist_argv, env, evlist, evlist_stderr, evlist_receipt, kind="perf_evlist")
    else:
        _reuse_or_fail(cfg, evlist_receipt, kind="perf_evlist", argv=evlist_argv, env=env, output=evlist, stderr=evlist_stderr)
    scanner = load_module("litchi_raw_stack_0978_decoder", cfg["scanner"])
    try:
        contract = scanner.parse_evlist_contract(evlist.read_text(encoding="utf-8", errors="strict"), expected_period=cfg["period"], expected_stack_cap=cfg["stack_cap"])
    except Exception as exc:  # noqa: BLE001
        _fail("evlist_contract", "bound perf evlist contract failed", error=str(exc))
    _require_full_lifetime_evlist(contract)
    dump = _path(child_dir, "perf-mmap-dump.txt")
    dump_stderr = _path(child_dir, "perf-mmap-dump.stderr")
    dump_receipt = _path(child_dir, "perf-mmap-dump.command.json")
    dump_argv = [str(cfg["perf"]), "script", "-D", "-i", str(data)]
    if dump.exists() != dump_stderr.exists():
        _fail("offline_command", "perf mmap dump has only one output stream", output=str(dump), stderr=str(dump_stderr))
    if not dump.is_file():
        _run_to_file(dump_argv, env, dump, dump_stderr, dump_receipt, kind="perf_mmap_dump")
    else:
        _reuse_or_fail(cfg, dump_receipt, kind="perf_mmap_dump", argv=dump_argv, env=env, output=dump, stderr=dump_stderr)
    return contract, dump


def _read_report(child_dir: Path) -> dict[str, Any]:
    report = load_json(_input(child_dir, "report.json"))
    expected = {"schema": "litchi.performance.0917.doc_retention.v1", "case": "docfloat", "edits": 1, "retention": "default_8mib", "warmups": 5, "samples_requested": 80}
    for key, value in expected.items():
        if report.get(key) != value:
            _fail("report_contract", "measurement report field differs", field=key, actual=report.get(key), expected=value)
    samples = report.get("samples")
    elapsed = report.get("elapsed_ns")
    if not isinstance(samples, list) or len(samples) != 80 or not isinstance(elapsed, list) or len(elapsed) != 80:
        _fail("report_contract", "measurement report does not contain 80 timing samples")
    for index, row in enumerate(samples):
        if not isinstance(row, Mapping) or row.get("index") != index or type(row.get("elapsed_ns")) is not int or row["elapsed_ns"] <= 0:
            _fail("report_oracle", "measurement report timing row is malformed", index=index)
        oracle = row.get("oracle")
        if not isinstance(oracle, Mapping) or oracle.get("oracle_ok") is not True:
            _fail("report_oracle", "measurement report oracle failed", index=index)
    if any(type(value) is not int or value <= 0 for value in elapsed):
        _fail("report_contract", "elapsed_ns vector is malformed")
    return report


def _parse_perf_time(value: str) -> int:
    try:
        parsed = Decimal(value) * Decimal(1_000_000_000)
    except (InvalidOperation, ValueError) as exc:
        _fail("corrected_time", "corrected perf timestamp is malformed", value=value, error=str(exc))
    if parsed != parsed.to_integral_value():
        _fail("corrected_time", "corrected perf timestamp is not nanosecond integral", value=value)
    return int(parsed)


def _parse_corrected(path: Path, *, expected_pid: int, expected_cpu: int, expected_period: int, marker: str, marker_dso: str) -> tuple[list[dict[str, Any]], int]:
    rows: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    lost_events = 0

    def finish() -> None:
        nonlocal current
        if current is not None:
            rows.append(current)
            current = None

    try:
        stream = path.open("r", encoding="utf-8", errors="strict")
    except OSError as exc:
        _fail("corrected_input", "cannot open corrected perf script", path=str(path), error=str(exc))
    with stream:
        for line_number, raw in enumerate(stream, 1):
            line = raw.rstrip("\n")
            if not line.strip():
                finish()
                continue
            header = HEADER_RE.match(line)
            if header is not None:
                finish()
                if header.group("event").strip() != "instructions:u":
                    _fail("corrected_event", "corrected stream event differs", line=line_number, event=header.group("event"))
                pid = int(header.group("pid"))
                tid = int(header.group("tid") or pid)
                cpu = int(header.group("cpu"))
                period = int(header.group("period"))
                if pid != expected_pid or tid != expected_pid or cpu != expected_cpu or period != expected_period:
                    _fail("corrected_identity", "corrected sample identity differs from launch/event contract", line=line_number, pid=pid, tid=tid, cpu=cpu, period=period)
                current = {"index": len(rows) + 1, "time_ns": _parse_perf_time(header.group("time")), "pid": pid, "tid": tid, "cpu": cpu, "period": period, "frames": [], "header_line": line_number}
                continue
            if current is not None and line[:1].isspace():
                frame_match = FRAME_RE.match(line)
                if frame_match is None:
                    _fail("corrected_frame", "corrected frame is malformed", line=line_number, text=line[:240])
                symbol = frame_match.group("symbol").strip()
                base, _, offset = symbol.partition("+")
                dso = frame_match.group("dso").strip()
                resolved = base.strip() not in {"", "[unknown]", "unknown"} and dso not in {"", "[unknown]"}
                current["frames"].append({"ip": int(frame_match.group("ip"), 16), "ip_hex": frame_match.group("ip").lower(), "symbol": symbol, "symbol_base": base.strip(), "offset": offset or None, "dso": dso, "resolved": resolved})
                continue
            low = line.strip().lower()
            if "lost" in low or "dropped" in low:
                match = LOST_EVENT_RE.search(low)
                if match is None:
                    _fail("corrected_loss", "lost/dropped line lacks an integer count", line=line_number, text=line[:240])
                lost_events += int(match.group(1))
                continue
            _fail("corrected_line", "unexpected corrected perf-script line", line=line_number, text=line[:240])
    finish()
    if not rows:
        _fail("corrected_samples_missing", "corrected perf script has no samples", path=str(path))
    for row in rows:
        for frame in row["frames"]:
            if frame["symbol_base"] == marker and frame["dso"] != marker_dso:
                _fail("marker_binding", "marker symbol resolved in unexpected DSO", index=row["index"], dso=frame["dso"], expected=marker_dso)
    return rows, lost_events


def _pair_samples(corrected: Sequence[Mapping[str, Any]], raw: Sequence[Mapping[str, Any]]) -> None:
    if len(corrected) != len(raw):
        _fail("sample_conservation", "corrected and raw sample counts differ", corrected=len(corrected), raw=len(raw))
    for index, (left, right) in enumerate(zip(corrected, raw), 1):
        for key in ("pid", "tid", "cpu", "period"):
            if left.get(key) != right.get(key):
                _fail("sample_identity", "corrected and raw sample identities differ", index=index, field=key, corrected=left.get(key), raw=right.get(key))
        if left.get("time_ns", 0) // 1000 != int(right.get("time_ns", -1)) // 1000:
            _fail("sample_identity", "corrected and raw timestamps differ at microsecond precision", index=index, corrected=left.get("time_ns"), raw=right.get("time_ns"))
        if int(left.get("index", -1)) != index or int(right.get("index", -1)) != index:
            _fail("sample_identity", "sample indices are not ordered and contiguous", index=index)


def _load_raw(scanner: Any, dump: Path, data: Path, binary: Path, contract: Mapping[str, Any], *, expected_pid: int, expected_cpu: int, period: int, stack_cap: int) -> dict[str, Any]:
    try:
        result = scanner.scan_dump(dump, perf_data=data, binary=binary, expected_pid=expected_pid, expected_cpu=expected_cpu, expected_period=period, stack_cap=stack_cap, event_contract=contract, retain_samples=True)
    except Exception as exc:  # noqa: BLE001
        _fail("raw_stack", "variable-layout scanner failed", error=str(exc))
    if result.get("status") != "complete" or result.get("sample_conservation") is not True:
        _fail("raw_stack", "raw scanner did not prove sample conservation", summary={key: result.get(key) for key in ("status", "sample_count", "sample_line_count", "sample_conservation")})
    if result.get("period_valid") is not True or result.get("pid_tid_cpu_valid") is not True or result.get("time_valid") is not True or result.get("bounds_valid") is not True or result.get("violations") != []:
        _fail("raw_stack", "raw scanner identity or bounds gate failed", summary={key: result.get(key) for key in ("period_valid", "pid_tid_cpu_valid", "time_valid", "bounds_valid", "violations")})
    rows = result.get("sample_ledger")
    if not isinstance(rows, list) or len(rows) != result.get("sample_count"):
        _fail("raw_stack", "raw scanner did not retain its full sample ledger")
    return result


def _elf_loads(path: Path) -> tuple[list[dict[str, int]], int]:
    data = path.read_bytes()
    if len(data) < 64 or data[:4] != b"\x7fELF" or data[4] != 2 or data[5] != 1:
        _fail("elf_layout", "target is not a little-endian ELF64", path=str(path))
    phoff = struct.unpack_from("<Q", data, 32)[0]
    phentsize, phnum = struct.unpack_from("<HH", data, 54)
    if phentsize < 56 or phoff + phentsize * phnum > len(data):
        _fail("elf_layout", "program header table is out of bounds")
    loads: list[dict[str, int]] = []
    for index in range(phnum):
        off = phoff + index * phentsize
        p_type, p_flags = struct.unpack_from("<II", data, off)
        if p_type != 1:
            continue
        p_offset, p_vaddr, p_filesz, p_memsz, p_align = struct.unpack_from("<QQQQQ", data, off + 8)
        page_vaddr = p_vaddr & ~0xFFF
        page_offset = p_offset & ~0xFFF
        loads.append({"index": index, "offset_page": page_offset, "vaddr_page": page_vaddr, "filesz": p_filesz, "memsz": p_memsz, "flags": p_flags, "align": p_align})
    if not loads:
        _fail("elf_layout", "target has no PT_LOAD segments")
    return loads, 0x1000


def _derive_layout(binary: Path, dump: Mapping[str, Any], *, pid: int, build_id: str, mmap_dump: Path) -> dict[str, Any]:
    records = dump.get("records")
    if not isinstance(records, list):
        _fail("layout_contract", "lifecycle records are missing for layout derivation")
    maps = [dict(row) for row in records if row.get("kind") == "PERF_RECORD_MMAP2" and row.get("pid") == pid and row.get("tid") == pid and row.get("path") == str(binary) and str(row.get("build_id") or "").lower() == build_id.lower()]
    if not maps:
        _fail("layout_contract", "lifecycle stream has no exact target MMAP2 records")
    loads, page_size = _elf_loads(binary)
    selected: list[dict[str, Any]] = []
    for load in loads:
        # A PT_LOAD is covered when [map.start,map.start+length) contains the
        # biased segment and the map has the matching file offset.
        valid: list[dict[str, Any]] = []
        for row in maps:
            bias = int(row["start"]) - load["vaddr_page"]
            seg_start = bias + load["vaddr_page"]
            seg_end = bias + load["vaddr_page"] + max(int(load["memsz"]), 1)
            if int(row.get("map_pgoff", -1)) == load["offset_page"] and int(row["start"]) <= seg_start < int(row["start"]) + int(row["length"]) and seg_end <= int(row["start"]) + int(row["length"]):
                valid.append(row)
        if not valid:
            _fail("layout_contract", "no target map covers PT_LOAD", offset=hex(load["offset_page"]))
        biases = {int(row["start"]) - load["vaddr_page"] for row in valid}
        if len(biases) != 1:
            _fail("layout_contract", "target maps imply inconsistent load bias", offset=hex(load["offset_page"]))
        selected.append({"load": load, "mmap": max(valid, key=lambda row: int(row.get("record_sequence", 0)))})
    biases = {int(item["mmap"]["start"]) - int(item["load"]["vaddr_page"]) for item in selected}
    if len(biases) != 1:
        _fail("layout_contract", "PT_LOAD mappings do not imply one load bias")
    corrected_base = next(iter(biases))
    executable = [item for item in selected if int(item["load"]["flags"]) & 1]
    if len(executable) != 1:
        _fail("layout_contract", "expected one executable PT_LOAD", count=len(executable))
    exec_map = executable[0]["mmap"]
    perf_expected_start = int(exec_map["start"]) - int(exec_map["map_pgoff"])
    if perf_expected_start - corrected_base != page_size:
        _fail("layout_contract", "perf/corrected base gap differs from one page", corrected_base=hex(corrected_base), perf_expected_start=hex(perf_expected_start))
    return {
        "schema": "litchi.performance.0978.target_layout.v1",
        "binary": binding(binary),
        "perf_mmap_dump": binding(mmap_dump),
        "corrected_base": hex(corrected_base),
        "perf_expected_start": hex(perf_expected_start),
        "page_size": page_size,
        "target_mmap2": [item["mmap"] for item in selected],
        "executable_mapping": exec_map,
        "formula": "load bias = selected target MMAP2 start - PT_LOAD page-down(vaddr); perf start = executable map start - map file offset",
    }


def _decoder_env(cfg: Mapping[str, Any], layout: Mapping[str, Any], log_path: Path) -> dict[str, str]:
    env = _base_env()
    shim = cfg.get("decoder_so")
    if shim is None:
        _fail("shim_binding", "corrected decode requires decoder shared object binding")
    env.update({
        "LD_PRELOAD": str(shim),
        "LITCHI_0973_SHIM_MODE": "corrected",
        "LITCHI_0973_SHIM_LOG": str(log_path),
        "LITCHI_0973_TARGET_PATH": str(cfg["binary"]),
        "LITCHI_0973_TARGET_NAME": cfg["binary"].name,
        "LITCHI_0973_CORRECTED_BASE": str(layout["corrected_base"]),
        "LITCHI_0973_PERF_EXPECTED_START": str(layout["perf_expected_start"]),
        "LITCHI_0976_TARGET_BUILD_ID": str(cfg["build_id"]),
        "LITCHI_0976_TARGET_SHA256": sha256(cfg["binary"]),
        "LITCHI_0976_EXPECTED_DELTA": hex(0x1000),
    })
    return env


def _ensure_corrected(child_dir: Path, cfg: Mapping[str, Any], layout: Mapping[str, Any]) -> tuple[Path, Path, Path, dict[str, Any] | None]:
    script = _path(child_dir, "corrected.perf-script-no-inline.txt")
    stderr = _path(child_dir, "corrected.perf.stderr")
    log = _path(child_dir, "corrected.dwfl.jsonl")
    data = _input(child_dir, "perf.data")
    if script is None or stderr is None:
        _fail("corrected_input", "corrected output paths could not be constructed")
    command_path = _path(child_dir, "corrected.command.json")
    argv = [str(cfg["perf"]), "script", "--show-lost-events", "--no-inline", "-F", "cpu,pid,tid,time,event,period,ip,sym,symoff,dso", "-i", str(data)]
    command_env = _decoder_env(cfg, layout, log)
    if script.exists() != stderr.exists():
        _fail("offline_command", "corrected stream has only one output stream", script=str(script), stderr=str(stderr))
    if not script.is_file():
        _run_to_file(argv, command_env, script, stderr, command_path, kind="corrected")
    else:
        _reuse_or_fail(cfg, command_path, kind="corrected", argv=argv, env=command_env, output=script, stderr=stderr)
    command: dict[str, Any] = load_json(command_path)
    if log is None or not log.is_file():
        _fail("shim_log", "corrected decoder did not produce bound DWFL log", path=str(log) if log else None)
    return script, stderr, log, command


def _physical_owner(cfg: Mapping[str, Any], lifecycle_module: Any, records: Sequence[Mapping[str, Any]], row: Mapping[str, Any], corrected: Mapping[str, Any], *, comm_sequence: int) -> dict[str, Any] | None:
    frames = corrected.get("frames")
    if not isinstance(frames, list) or not frames:
        return None
    leaf = frames[0]
    if leaf.get("resolved") is not True or not isinstance(leaf.get("dso"), str):
        return None
    ip = leaf.get("ip")
    if type(ip) is not int:
        return None
    sequence = row.get("record_sequence")
    if type(sequence) is not int:
        return None
    maps = [event for event in records if event.get("kind") == "PERF_RECORD_MMAP2" and event.get("pid") == row.get("pid") and event.get("tid") == row.get("tid")]
    try:
        # The target COMM_EXEC cutoff applies only after the sample itself is
        # in the target exec sequence.  Pre-exec launcher samples may resolve
        # against mappings established before that COMM event.
        map_cutoff = comm_sequence if sequence >= comm_sequence else None
        active = lifecycle_module._interval_owner(maps, ip, record_sequence=sequence, comm_sequence=map_cutoff)
    except Exception:  # noqa: BLE001
        active = None
    if not isinstance(active, Mapping):
        return None
    dso = str(leaf.get("dso"))
    build = str(active.get("build_id") or "").lower()
    if dso != str(active.get("path")) or not re.fullmatch(r"[0-9a-f]{8,64}", build):
        return None
    symbol = str(leaf.get("symbol_base") or leaf.get("symbol") or "")
    if not symbol:
        return None
    return {"physical_leaf": {"ip": ip, "dso_path": dso, "build_id": build, "symbol": symbol, "map_record_sequence": int(active["record_sequence"])}, "target_source_eligible": dso == str(cfg["binary"]) and build == str(cfg["build_id"]).lower()}


def _make_attributions(cfg: Mapping[str, Any], lifecycle_module: Any, lifecycle: Mapping[str, Any], corrected: Sequence[Mapping[str, Any]]) -> tuple[dict[int, dict[str, Any]], list[dict[str, Any]]]:
    records = lifecycle.get("records")
    sample_records = lifecycle.get("sample_records")
    if not isinstance(records, list) or not isinstance(sample_records, list) or len(sample_records) != len(corrected):
        _fail("sample_identity", "lifecycle sample records do not match corrected samples")
    comms = [event for event in records if event.get("kind") == "PERF_RECORD_COMM" and event.get("pid") == corrected[0].get("pid") and event.get("tid") == corrected[0].get("pid") and event.get("exec") is True and str(event.get("comm")) == cfg["binary"].name[:15]]
    if len(comms) != 1:
        # Do not guess a COMM identity from frame names; build_ledger reports a
        # precise failure if the lifecycle stream is ambiguous.
        comm_sequence = -1
    else:
        comm_sequence = int(comms[0]["record_sequence"])
    attributions: dict[int, dict[str, Any]] = {}
    corrected_with_raw: list[dict[str, Any]] = []
    for index, (row, descriptor) in enumerate(zip(corrected, sample_records), 1):
        out = dict(row)
        out["lifecycle_record_sequence"] = descriptor.get("record_sequence")
        out["raw_identity"] = {"index": index, "record_sequence": descriptor.get("record_sequence"), "record_offset": descriptor.get("record_offset"), "time_ns": descriptor.get("time_ns"), "pid": descriptor.get("pid"), "tid": descriptor.get("tid"), "cpu": descriptor.get("cpu"), "period": descriptor.get("period"), "raw_bytes": descriptor.get("record_bytes")}
        corrected_with_raw.append(out)
        # A sample before the target COMM_EXEC is still allowed to have a
        # defensible physical leaf: the process may already have libc/Python
        # mappings from its launcher.  Ask the lifecycle map validator about
        # that leaf with its pre-exec cutoff rules instead of dropping the
        # attribution.  The lifecycle ledger keeps this row in ``startup``
        # (and therefore out of T/target-source totals), while retaining the
        # exact DSO/build-id/map-sequence evidence for full conservation.
        if comm_sequence < 0:
            continue
        attr = _physical_owner(cfg, lifecycle_module, records, descriptor, row, comm_sequence=comm_sequence)
        marker_frames = [frame for frame in row.get("frames", []) if frame.get("symbol_base") == cfg["marker"] and frame.get("dso") == cfg["marker_dso"]]
        if marker_frames:
            marker_frame = marker_frames[0]
            if attr is None:
                # The marker can be known while frame zero is unresolved.  A
                # typed physical_unknown claim preserves M and U_marker
                # without inventing a physical owner.
                attr = {"physical_unknown": True}
            marker_maps = [event for event in records if event.get("kind") == "PERF_RECORD_MMAP2" and event.get("pid") == descriptor.get("pid") and event.get("tid") == descriptor.get("tid")]
            marker_ip = marker_frame.get("ip")
            marker_map = None
            if type(marker_ip) is int and comm_sequence >= 0:
                marker_map = lifecycle_module._interval_owner(marker_maps, marker_ip, record_sequence=int(descriptor.get("record_sequence", -1)), comm_sequence=comm_sequence if int(descriptor.get("record_sequence", -1)) >= comm_sequence else None)
            marker_info = {"present": True, "resolved": bool(marker_frame.get("resolved")), "ip": marker_ip, "dso_path": cfg["marker_dso"], "build_id": cfg["build_id"], "symbol": cfg["marker"]}
            if isinstance(marker_map, Mapping):
                marker_info["map_record_sequence"] = int(marker_map["record_sequence"])
            attr["marker_frame"] = marker_info
        if attr is not None:
            attributions[index] = attr
    return attributions, corrected_with_raw


def _summary(cfg: Mapping[str, Any], raw: Mapping[str, Any], ledger: Mapping[str, Any], corrected: Sequence[Mapping[str, Any]], terminal_loss: int, corrected_script: Path, corrected_log: Path, shim_log: Path, lost_events: int, *, lifecycle_rows: Sequence[Mapping[str, Any]] | None = None) -> dict[str, Any]:
    total_period = sum(int(row["period"]) for row in corrected)
    marker_rows = [row for row in corrected if any(frame.get("symbol_base") == cfg["marker"] and frame.get("dso") == cfg["marker_dso"] for frame in row.get("frames", []))]
    marker_period = sum(int(row["period"]) for row in marker_rows)
    unresolved = sum(1 for row in corrected for frame in row.get("frames", []) if frame.get("resolved") is not True)
    empty = sum(1 for row in corrected if not row.get("frames"))
    known = ledger.get("target_known_owner_period", {})
    physical_known = ledger.get("marker_known_owner_period", {})
    # U_marker follows the lifecycle policy, which marks kernel-mode marker
    # samples unknown even when the physical instruction frame is mapped to a
    # DSO.  Counting only an unresolved frame-0 would under-report this
    # disjoint bucket (and disagree with the independent reducer).  The
    # lifecycle rows are the authority for this phase/mode classification.
    if lifecycle_rows is not None:
        marker_unknown_samples = sum(1 for row in lifecycle_rows if row.get("marker") is True and row.get("unknown") is True)
    else:
        marker_unknown_samples = sum(1 for row in corrected if any(frame.get("symbol_base") == cfg["marker"] and frame.get("dso") == cfg["marker_dso"] for frame in row.get("frames", [])) and (not row.get("frames") or row["frames"][0].get("resolved") is not True))
    return {
        "schema": "litchi.performance.0978.corrected_summary.v1",
        "decoder_version": "focuseddecode-0978-v1",
        "sample_count": len(corrected), "total_period": total_period,
        "marker_sample_count": len(marker_rows), "marker_period": marker_period,
        "marker_period_share": marker_period / total_period if total_period else None,
        "empty_callchain_count": empty, "unresolved_frames": unresolved,
        "explicit_stack_truncated_count": 0, "lost_events": lost_events,
        "raw_lost_sum": int(raw.get("raw_lost_sum", 0)), "terminal_lost_samples": terminal_loss,
        "physical_leaf_owner_period": dict(sorted((str(k), int(v)) for k, v in ledger.get("marker_known_owner_period", {}).items())),
        "marker_known_owner_period": dict(sorted((str(k), int(v)) for k, v in ledger.get("marker_known_owner_period", {}).items())),
        "target_known_owner_period": dict(sorted((str(k), int(v)) for k, v in ledger.get("target_known_owner_period", {}).items())),
        "target_source_eligible_period": dict(sorted((str(k), int(v)) for k, v in ledger.get("target_source_eligible_period", {}).items())),
        "physical_leaf_unknown_period": int(ledger.get("U_marker", 0)), "physical_leaf_unknown_samples": marker_unknown_samples,
        "physical_leaf_conservation": {"passed": sum(int(v) for v in physical_known.values()) + int(ledger.get("U_marker", 0)) == int(ledger.get("M", 0)), "known_period": sum(int(v) for v in physical_known.values()), "unknown_period": int(ledger.get("U_marker", 0)), "marker_period": int(ledger.get("M", 0))},
        "corrected_script_sha256": sha256(corrected_script), "corrected_log_sha256": sha256(corrected_log), "shim_log_sha256": sha256(shim_log),
        "summary_scope": "physical no-inline stream; full sample/frame rows are in attribution sidecar",
    }


def _abi_summary(raw: Mapping[str, Any]) -> dict[str, Any]:
    rows = raw.get("sample_ledger")
    if not isinstance(rows, list):
        _fail("abi_scan", "raw scanner sample ledger is missing")
    total = sum(int(row["period"]) for row in rows)
    empty_rows = [row for row in rows if int(row.get("callchain_nr", 0)) == 0 and not bool(row.get("regs_available", False)) and int(row.get("stack_size", 0)) == 0]
    payload_present = total - sum(int(row["period"]) for row in empty_rows)
    # This sidecar is a layout/presence ledger.  It has no right to call a
    # non-empty stack "unwound"; physical resolution is proved separately by
    # the corrected stream and lifecycle attribution.
    return {"schema": ABI_SCHEMA, "status": "complete", "sample_count": len(rows), "total_period": total, "valid_record_count": len(rows), "empty_valid_record_count": len(empty_rows), "payload_present_period": payload_present, "resolved_period": 0, "unknown_period": total, "malformed_count": 0, "duplicate_sample_count": 0, "missing_sample_count": 0, "period_conservation": {"passed": total == 0 + total, "resolved_plus_unknown": total}, "scope": "raw variable-layout presence; no unwindability claim; valid empty payloads are unknown"}


def _artifacts(child_dir: Path, paths: Mapping[str, Path]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for name, path in paths.items():
        if path.is_file():
            result[name] = binding(path)
    return result


def decode_child(plan_path: Path | Mapping[str, Any], child_dir: Path) -> dict[str, Any]:
    """Decode one child from an immutable plan and capture directory."""
    child_dir = Path(child_dir).absolute()
    if not child_dir.is_dir():
        _fail("child_input", "child directory is missing", child=str(child_dir))
    plan, plan_file = _plan_data(plan_path)
    cfg = _validate_plan(plan)
    launch = _load_launch(child_dir, cfg["binary"])
    terminal = _load_terminal(child_dir)
    command = load_json(_input(child_dir, "command.json"))
    report = _read_report(child_dir)
    perf_data = _input(child_dir, "perf.data")
    if command.get("packet_id", child_dir.name) != child_dir.name:
        _fail("command_contract", "command packet ID differs from child")
    perf_header, perf_header_stderr, perf_header_command = _perf_header(child_dir, cfg)
    terminal_report, terminal_summary, terminal_loss, terminal_report_command = _terminal_report(child_dir, cfg)
    contract, mmap_dump = _evlist_and_dump(child_dir, cfg)
    lifecycle_module = load_module("litchi_lifecycle_ledger_0978_decoder", cfg["lifecycle"])
    try:
        lifecycle = lifecycle_module.parse_lifecycle_dump(mmap_dump, expected_pid=int(launch["exec_pid"]), expected_cpu=cfg["cpu"], require_raw_metadata=True)
    except Exception as exc:  # noqa: BLE001
        _fail("lifecycle", "lifecycle parser rejected retained perf script -D stream", error=str(exc))
    if lifecycle.get("status") != "complete" or lifecycle.get("raw_metadata_crosscheck") is not True or lifecycle.get("unique_record_offsets") is not True:
        _fail("lifecycle", "lifecycle parser did not prove framed metadata custody", summary={key: lifecycle.get(key) for key in ("status", "raw_event_count", "raw_metadata_count", "raw_metadata_crosscheck", "unique_record_offsets")})
    layout = _derive_layout(cfg["binary"], lifecycle, pid=int(launch["exec_pid"]), build_id=cfg["build_id"], mmap_dump=mmap_dump)
    layout_path = _path(child_dir, "target-layout.json")
    if not layout_path.is_file():
        write_exclusive(layout_path, layout)
    scanner = load_module("litchi_raw_stack_0978_decoder", cfg["scanner"])
    raw = _load_raw(scanner, mmap_dump, perf_data, cfg["binary"], contract, expected_pid=int(launch["exec_pid"]), expected_cpu=cfg["cpu"], period=cfg["period"], stack_cap=cfg["stack_cap"])
    if int(raw.get("raw_lost_sum", 0)) != 0 or terminal_loss != 0:
        _fail("loss_gate", "raw or terminal loss is nonzero", raw_lost_sum=raw.get("raw_lost_sum"), terminal_lost_samples=terminal_loss)
    write_exclusive(_path(child_dir, "loss.json"), {"schema": "litchi.performance.0978.loss.v1", "raw_lost_sum": int(raw.get("raw_lost_sum", 0)), "raw_lost_record_count": int(raw.get("raw_lost_record_count", 0)), "terminal_lost_samples": terminal_loss, "raw_scope": "PERF_RECORD_LOST fields; separate from PERF_RECORD_LOST_SAMPLES"})
    raw_path = _path(child_dir, "raw-stack.json")
    write_exclusive(raw_path, raw)
    corrected_path, corrected_log, shim_log, corrected_command = _ensure_corrected(child_dir, cfg, layout)
    corrected, corrected_lost_events = _parse_corrected(corrected_path, expected_pid=int(launch["exec_pid"]), expected_cpu=cfg["cpu"], expected_period=cfg["period"], marker=cfg["marker"], marker_dso=cfg["marker_dso"])
    raw_rows = raw.get("sample_ledger")
    if not isinstance(raw_rows, list):
        _fail("sample_conservation", "raw sample ledger is absent")
    corrected_identity = [{"index": row["index"], "time_ns": row["time_ns"], "pid": row["pid"], "tid": row["tid"], "cpu": row["cpu"], "period": row["period"], "frames": row["frames"], "marker": any(frame.get("symbol_base") == cfg["marker"] and frame.get("dso") == cfg["marker_dso"] for frame in row["frames"])} for row in corrected]
    _pair_samples(corrected_identity, raw_rows)
    attributions, corrected_with_raw = _make_attributions(cfg, lifecycle_module, lifecycle, corrected_identity)
    target_comm = cfg["binary"].name[:15]
    try:
        ledger_packet = lifecycle_module.build_ledger(lifecycle, raw_rows, expected_pid=int(launch["exec_pid"]), expected_cpu=cfg["cpu"], expected_period=cfg["period"], target_path=str(cfg["binary"]), target_build_id=cfg["build_id"], target_comm=target_comm, terminal_lost_samples=terminal_loss, marker=cfg["marker"], marker_dso=cfg["marker_dso"], attributions=attributions)
    except Exception as exc:  # noqa: BLE001
        _fail("lifecycle_ledger", "lifecycle ledger rejected paired attribution", error=str(exc))
    lifecycle_rows: list[dict[str, Any]] = []
    inclusive_period: Counter[str] = Counter()
    for sample, corrected_row in zip(ledger_packet.get("samples", []), corrected_with_raw):
        raw_identity = sample.get("raw_identity") if isinstance(sample.get("raw_identity"), Mapping) else {}
        row = dict(sample)
        for key in ("time_ns", "pid", "tid", "cpu", "period"):
            row[key] = raw_identity.get(key)
        owners = sorted({f"{frame.get('dso')}::{frame.get('symbol_base')}" for frame in corrected_row.get("frames", []) if frame.get("resolved") is True and frame.get("dso") not in {None, "", "[unknown]"} and frame.get("symbol_base") not in {None, "", "[unknown]", "unknown"}})
        row["inclusive_owners"] = owners if row.get("marker") else []
        if row.get("marker"):
            for owner in owners:
                inclusive_period[owner] += int(row["period"])
        lifecycle_rows.append(row)
    lifecycle_sidecar = {"schema": LEDGER_SCHEMA, "status": "complete", "stream": lifecycle, "target": ledger_packet.get("target"), "lifecycle": ledger_packet.get("lifecycle"), "ledger": ledger_packet.get("ledger"), "gates": ledger_packet.get("gates"), "samples": lifecycle_rows}
    lifecycle_path = _path(child_dir, "lifecycle-ledger.json")
    write_exclusive(lifecycle_path, lifecycle_sidecar)
    attribution_sidecar = {"schema": ATTRIBUTION_SCHEMA, "status": "complete", "sample_count": len(corrected_with_raw), "corrected_samples": corrected_with_raw, "raw_samples": raw_rows, "lifecycle_samples": lifecycle_rows, "attributions": attributions, "identity": {"ordered": True, "join": "index + pid + tid + cpu + period + time_ns//1000", "sample_count": len(corrected_with_raw)}}
    attribution_path = _path(child_dir, "attribution-samples.json")
    write_exclusive(attribution_path, attribution_sidecar)
    abi = _abi_summary(raw)
    abi_path = _path(child_dir, "abi-scan.json")
    write_exclusive(abi_path, abi)
    corrected_summary = _summary(cfg, raw, ledger_packet["ledger"], corrected_identity, terminal_loss, corrected_path, corrected_log, shim_log, corrected_lost_events, lifecycle_rows=lifecycle_rows)
    corrected_summary["physical_leaf_owner_period"] = ledger_packet["ledger"].get("marker_known_owner_period", {})
    corrected_summary["marker_known_owner_period"] = ledger_packet["ledger"].get("marker_known_owner_period", {})
    corrected_summary["inclusive_owner_period"] = dict(sorted(inclusive_period.items()))
    corrected_summary["target_known_owner_period"] = ledger_packet["ledger"].get("target_known_owner_period", {})
    corrected_summary["target_source_eligible_period"] = ledger_packet["ledger"].get("target_source_eligible_period", {})
    corrected_summary["lifecycle_unknown_period"] = ledger_packet["ledger"].get("U_target")
    gate_reasons: list[str] = []
    if corrected_summary["marker_sample_count"] < int((plan.get("child_gate") or {}).get("min_marker_samples", 100)):
        gate_reasons.append("marker_sample_count_lt_minimum")
    if corrected_summary["marker_period"] <= 0:
        gate_reasons.append("marker_period_nonpositive")
    if corrected_summary["physical_leaf_conservation"].get("passed") is not True:
        gate_reasons.append("physical_leaf_conservation_failed")
    if not all(bool(value) for value in ledger_packet["ledger"].get("conservation", {}).values()):
        gate_reasons.append("ledger_conservation_failed")
    if gate_reasons:
        _fail("child_gate", "decoder quality gate failed", reasons=gate_reasons, summary={key: corrected_summary.get(key) for key in ("sample_count", "marker_sample_count", "marker_period", "empty_callchain_count")})
    artifacts = _artifacts(child_dir, {"perf_data": perf_data, "command": _path(child_dir, "command.json"), "launch": _path(child_dir, "launch.json"), "terminal": _path(child_dir, "terminal.json"), "report": _path(child_dir, "report.json"), "perf_mmap_dump": mmap_dump, "raw_task": mmap_dump, "perf_mmap_dump_stderr": _path(child_dir, "perf-mmap-dump.stderr"), "perf_mmap_dump_command": _path(child_dir, "perf-mmap-dump.command.json"), "perf_evlist": _path(child_dir, "perf-evlist-v.txt"), "perf_evlist_stderr": _path(child_dir, "perf-evlist-v.stderr"), "perf_evlist_command": _path(child_dir, "perf-evlist-v.command.json"), "perf_header": perf_header, "perf_header_stderr": perf_header_stderr, "perf_header_command": _path(child_dir, "perf-header.command.json"), "corrected_script": corrected_path, "corrected_log": corrected_log, "shim_log": shim_log, "corrected_command": _path(child_dir, "corrected.command.json"), "raw_stack": raw_path, "lifecycle_ledger": lifecycle_path, "attribution_samples": attribution_path, "abi_scan": abi_path, "terminal_report": terminal_report, "terminal_report_stderr": _path(child_dir, "terminal-report.stderr"), "terminal_report_command": _path(child_dir, "terminal-report.command.json"), "terminal_summary": _path(child_dir, "terminal-summary.json"), "loss": _path(child_dir, "loss.json"), "target_layout": layout_path})
    packet = {
        "schema": SCHEMA, "status": "complete", "packet_id": child_dir.name,
        "child_number": int(re.search(r"(\d+)$", child_dir.name).group(1)) if re.search(r"(\d+)$", child_dir.name) else None,
        "created_utc": dt.datetime.now(dt.timezone.utc).isoformat(), "plan": {"path": str(plan_file) if plan_file else None, "schema": plan.get("schema")},
        "launch": launch, "terminal": terminal, "report": report,
        "timing": {"elapsed_ns": list(report["elapsed_ns"]), "warmups": report["warmups"], "measured_samples": report["samples_requested"]},
        "layout": layout, "perf_contract": {"evlist": contract, "terminal": terminal_summary},
        "raw_stack": {key: value for key, value in raw.items() if key != "sample_ledger"},
        # Keep the packet field as the direct lifecycle summary.  The full
        # framed event stream is retained in lifecycle_ledger; a nested
        # ``stream`` alias here is ambiguous to the runner packet gate.
        "lifecycle": {**ledger_packet["lifecycle"]},
        "attribution": {**ledger_packet["ledger"], "physical_leaf_owner_period": ledger_packet["ledger"].get("marker_known_owner_period", {}), "inclusive_owner_period": dict(sorted(inclusive_period.items())), "physical_leaf_unknown_period": corrected_summary["physical_leaf_unknown_period"], "physical_leaf_unknown_samples": corrected_summary["physical_leaf_unknown_samples"]},
        "perf": {"corrected": corrected_summary}, "abi_scan": abi, "artifacts": artifacts,
        "decoder": {"schema": DECODER_SCHEMA, "native_workload_started": False, "offline_only": True, "focused_decode": binding(cfg["focused_decode"]), "corrected_command": corrected_command, "raw_scanner": binding(cfg["scanner"]), "lifecycle_helper": binding(cfg["lifecycle"]), "command_receipts": {name: binding(path) for name, path in (("perf_header", _path(child_dir, "perf-header.command.json")), ("terminal_report", _path(child_dir, "terminal-report.command.json")), ("perf_evlist", _path(child_dir, "perf-evlist-v.command.json")), ("perf_mmap_dump", _path(child_dir, "perf-mmap-dump.command.json")), ("corrected", _path(child_dir, "corrected.command.json"))) }},
        "summary": {"sample_count": corrected_summary["sample_count"], "marker_sample_count": corrected_summary["marker_sample_count"], "marker_period": corrected_summary["marker_period"], "marker_period_share": corrected_summary["marker_period_share"], "empty_callchain_count": corrected_summary["empty_callchain_count"], "raw_lost_sum": raw.get("raw_lost_sum", 0), "terminal_lost_samples": terminal_loss, "scope": "0978 corrected physical stream plus disjoint lifecycle ledger"},
        "gate": {"passed": True, "reasons": [], "raw_lost_sum": int(raw.get("raw_lost_sum", 0)), "terminal_lost_samples": terminal_loss, "passed_structural_conservation": True, "valid_empty_is_unknown": True, "minimum_marker_samples": int((plan.get("child_gate") or {}).get("min_marker_samples", 100)), "unresolved_frames": corrected_summary["unresolved_frames"], "empty_callchain_count": corrected_summary["empty_callchain_count"]},
    }
    write_exclusive(_path(child_dir, "decoder.json"), {"schema": DECODER_SCHEMA, "status": "complete", "packet_id": child_dir.name, "packet_schema": SCHEMA, "artifacts": artifacts, "commands": {"perf_header": binding(_path(child_dir, "perf-header.command.json")), "terminal_report": binding(_path(child_dir, "terminal-report.command.json")), "perf_evlist": binding(_path(child_dir, "perf-evlist-v.command.json")), "perf_mmap_dump": binding(_path(child_dir, "perf-mmap-dump.command.json")), "corrected": binding(_path(child_dir, "corrected.command.json"))}})
    return packet


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--child", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        packet = decode_child(args.plan, args.child)
        output = _path(args.child.absolute(), "packet.json")
        write_exclusive(output, packet)
        print(json.dumps({"status": packet["status"], "packet": str(output), "packet_sha256": sha256(output)}, indent=2, sort_keys=True))
        return 0
    except DecodeFailure as exc:
        print(json.dumps(exc.as_dict(), indent=2, sort_keys=True), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
