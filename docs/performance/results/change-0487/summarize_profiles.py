#!/usr/bin/env python3
"""Validate and summarize the bounded 0487 diagnostic profile children.

This is a read-only analyzer for ``profile.py`` output.  It never starts a
profiler, a benchmark, Cargo, or Git.  The profile driver deliberately keeps
one child per phase, input, workload, and tool; this helper validates those
receipts, replays the sealed 0484 report validator in memory, and writes a
content-bound JSON/Markdown summary.  A missing or unsupported PMU counter is
retained as unavailable instead of being represented as zero.  Syscall rows
from the after phase are compared with the retained 0485 after diagnostics;
the 0484 metadata profile is not used as a syscall baseline.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
from pathlib import Path
import re
import shutil
import sys
from typing import Any, Iterable, Mapping

from support import REPO, ROOT, environment, meta, now, read, sha, write


# The 0484 tree is a sealed input.  It is used only for its route construction
# and report validator; this module never calls its capture or mutation paths.
SEALED_ROOT = ROOT.parent / "change-0484"
RETAINED_ROOT = ROOT.parent / "change-0485"
if str(SEALED_ROOT) not in sys.path:
    sys.path.insert(0, str(SEALED_ROOT))
import measure_routes as routes  # noqa: E402
import profile_input_metadata as inputs  # noqa: E402


SCHEMA = "docx-opc-splice-audit-profile-summary-v1"
VERSION = 1
PROFILE_DRIVER = ROOT / "profile.py"
RETAINED_PROFILE_DRIVER = RETAINED_ROOT / "profile.py"
SUMMARIZER = Path(__file__).resolve()
BEFORE_BUILD = RETAINED_ROOT / "build-normal.json"
AFTER_BUILD = ROOT / "build-normal.json"
RETAINED_PROFILES = RETAINED_ROOT / "profiles" / "profiles1"
PHASES = ("before", "after")
INPUT_MODES = ("owned", "file")
TOOLS = ("perf", "strace")
WORKLOADS = tuple(inputs.PROFILE_CASES)
PERF_EVENTS = ("cycles", "instructions", "branches", "branch-misses", "page-faults")
TRACE_SYSCALLS = tuple(inputs.TRACE_SYSCALLS)
TRACE_FILTER = inputs.TRACE_FILTER
SAMPLES = 1
WARMUPS = 1
RSS_REVIEW_PERCENT = 5.0
PERF_REVIEW_PERCENT = 5.0
WHOLE_CHILD_SCOPE = "whole_child_including_setup_and_oracle_and_profiler_overhead"
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
ATTEMPT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
EXPECTED_BUILD_ATTEMPTS = {"before": "after1", "after": "after1"}


class SummaryError(RuntimeError):
    """A retained profile or its binding failed closed."""


def fail(message: str) -> None:
    raise SummaryError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _finite(value: Any, path: str = "value") -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            _finite(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{path}[{index}]")


def _read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing regular JSON file: {path}")
    value = read(path)
    _finite(value, str(path))
    return value


def _dict(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: expected an object")
    return value


def _list(value: Any, label: str) -> list[Any]:
    require(isinstance(value, list), f"{label}: expected a list")
    return value


def _string(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label}: expected a non-empty string")
    return value


def _sha(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None, f"{label}: invalid SHA-256")
    return value


def _integer(value: Any, label: str, *, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{label}: expected integer >= {minimum}")
    return value


def _artifact(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing regular file: {path}")
    details = meta(path)
    require(type(details["bytes"]) is int and details["bytes"] >= 0, f"{label}: invalid byte count")
    _sha(details["sha256"], f"{label}.sha256")
    return {"path": str(path), **details}


def _nonempty(path: Path) -> bool:
    return path.is_file() and not path.is_symlink() and path.stat().st_size > 0


def _percent_change(before: int | float | None, after: int | float | None) -> dict[str, Any]:
    """Return a safe pair change; a zero baseline never becomes infinity."""

    result: dict[str, Any] = {
        "before": before,
        "after": after,
        "absolute_change": None,
        "percent_change": None,
        "zero_baseline": False,
    }
    if before is None or after is None:
        return result
    result["absolute_change"] = after - before
    if before == 0:
        result["zero_baseline"] = True
        return result
    result["percent_change"] = (after - before) * 100.0 / before
    return result


def _review(change: Mapping[str, Any], threshold: float) -> bool | None:
    value = change.get("percent_change")
    if value is None:
        return None
    return abs(float(value)) >= threshold


def _build_command() -> list[str]:
    return [
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        "tools/perf-baseline/Cargo.toml",
        "--bin",
        "docx_replayable_tail_append",
    ]


def _resolve(root: Path, value: Any, label: str) -> Path:
    path = Path(_string(value, f"{label}.path"))
    return path.resolve() if path.is_absolute() else (root / path).resolve()


def _source_binding(value: Any, root: Path, label: str) -> dict[str, Any]:
    source = _dict(value, label)
    source_path = _resolve(root, source.get("path"), label)
    source_sha = _sha(source.get("sha256"), f"{label}.sha256")
    source_files = _integer(source.get("files"), f"{label}.files", minimum=1)
    actual = _artifact(source_path, label)
    require(actual["sha256"] == source_sha, f"{label}: source manifest hash changed")
    manifest = _read(source_path)
    require(isinstance(manifest, dict), f"{label}: source manifest is not an object")
    require(len(manifest) == source_files, f"{label}: source manifest file count differs")
    for name, digest in manifest.items():
        require(isinstance(name, str) and name, f"{label}: source manifest name is invalid")
        _sha(digest, f"{label}.{name}")
    return {
        "path": str(source_path),
        "relative_path": str(source.get("path")),
        "files": source_files,
        "sha256": source_sha,
        "map": {str(name): str(digest) for name, digest in manifest.items()},
    }


def _sealed_protocol_binding() -> dict[str, Any]:
    """Bind the unchanged 0484 route protocol used by every profile."""

    path = SEALED_ROOT / "route-protocol.json"
    actual = _artifact(path, "sealed 0484 route protocol")
    value = _read(path)
    require(value.get("schema") == routes.ROUTE_SCHEMA, f"{path}: sealed route protocol schema differs")
    require(value.get("version") == routes.ROUTE_PROTOCOL_VERSION, f"{path}: sealed route protocol version differs")
    return {"path": str(path), "sha256": actual["sha256"], "schema": value["schema"], "version": value["version"]}


def _load_build(phase: str) -> dict[str, Any]:
    require(phase in PHASES, f"unknown build phase: {phase}")
    root = RETAINED_ROOT if phase == "before" else ROOT
    path = BEFORE_BUILD if phase == "before" else AFTER_BUILD
    record = _read(path)
    require(record.get("schema") == "docx-replayable-tail-append-build-v1", f"{path}: build schema differs")
    require(record.get("version") == 1, f"{path}: build version differs")
    expected_attempt = EXPECTED_BUILD_ATTEMPTS[phase]
    require(record.get("attempt") == expected_attempt, f"{path}: build attempt differs")
    require(record.get("role") == "normal", f"{path}: profile summary requires normal build")
    require(record.get("command") == _build_command(), f"{path}: build command differs")
    require(record.get("source_unchanged") is True, f"{path}: source custody failed")
    source_before = _source_binding(record.get("source_before"), root, f"{path}.source_before")
    source_after = _source_binding(record.get("source_after"), root, f"{path}.source_after")
    require(source_before["map"] == source_after["map"], f"{path}: source changed during build")
    gate = _dict(record.get("gate"), f"{path}.gate")
    gate_path = _resolve(root, gate.get("path"), f"{path}.gate")
    gate_sha = _sha(gate.get("sha256"), f"{path}.gate.sha256")
    gate_actual = _artifact(gate_path, f"{path}.gate")
    require(gate_actual["sha256"] == gate_sha, f"{path}: gate receipt hash changed")
    gate_value = _read(gate_path)
    require(gate_value.get("schema") == "docx-stream-append-gate-v1", f"{gate_path}: gate schema differs")
    require(gate_value.get("exit_code") == 0, f"{gate_path}: gate command failed")
    require(gate_value.get("source_unchanged") is True, f"{gate_path}: gate source custody failed")
    require(gate_value.get("argv") == _build_command(), f"{gate_path}: gate command differs")
    require(gate_value.get("cwd") == str(REPO), f"{gate_path}: gate cwd differs")
    require(gate_value.get("environment") == environment(), f"{gate_path}: gate environment differs")
    expected_gate = root / "gate.py"
    expected_common = root / "support.py"
    require(gate_value.get("driver_sha256") == sha(expected_gate), f"{gate_path}: gate driver differs")
    require(gate_value.get("common_sha256") == sha(expected_common), f"{gate_path}: common helper differs")
    require(gate_value.get("source_before") == record.get("source_before"), f"{gate_path}: source before differs")
    require(gate_value.get("source_after") == record.get("source_after"), f"{gate_path}: source after differs")
    require(record.get("environment") == environment(), f"{path}: build environment differs")
    binary = _dict(record.get("binary"), f"{path}.binary")
    binary_path = Path(_string(binary.get("path"), f"{path}.binary.path"))
    require(binary_path.is_absolute(), f"{path}.binary.path: expected absolute copied binary path")
    actual_binary = _artifact(binary_path, f"{path}.binary")
    require(actual_binary["bytes"] == binary.get("bytes"), f"{path}: binary byte count changed")
    require(actual_binary["sha256"] == binary.get("sha256"), f"{path}: binary hash changed")
    require(binary.get("executable") is True and os.access(binary_path, os.X_OK), f"{path}: binary is not executable")
    protocol_binding = _sealed_protocol_binding() if phase == "before" else None
    return {
        "phase": phase,
        "path": str(path),
        "record_sha256": sha(path),
        "attempt": expected_attempt,
        "role": "normal",
        "binary": dict(binary),
        "source_before": {key: value for key, value in source_before.items() if key != "map"},
        "source_after": {key: value for key, value in source_after.items() if key != "map"},
        "source_map": source_after["map"],
        "gate": {"path": str(gate_path), "sha256": gate_sha},
        "protocol": protocol_binding,
    }


def _profile_build_binding(path: Path, build: Mapping[str, Any], label: str) -> None:
    expected = {"path": str(path), **meta(path)}
    actual = _dict(build, label)
    require(actual == expected, f"{label}: profile build receipt binding differs")


def _expected_labels() -> list[str]:
    labels: list[str] = []
    for phase in PHASES:
        for workload in WORKLOADS:
            for input_mode in INPUT_MODES:
                tools = ("perf",) if phase == "before" else ("perf", "strace")
                labels.extend(f"{phase}-{tool}-{input_mode}-{workload}" for tool in tools)
    return labels


def _expected_source(workload: str, input_mode: str) -> dict[str, Any]:
    arm = inputs._arm_for(workload, input_mode)
    value = inputs._input_metadata(arm)
    require(isinstance(value, dict), f"{workload}/{input_mode}: input metadata is missing")
    return value


def _expected_arm(workload: str, input_mode: str) -> dict[str, Any]:
    return dict(inputs._arm_for(workload, input_mode))


def _expected_argv(binary: Mapping[str, Any], arm: Mapping[str, Any], directory: Path) -> list[str]:
    case = routes._axis_case(dict(arm))
    return routes._axis_argv(
        dict(binary),
        case,
        dict(arm),
        samples=SAMPLES,
        warmups=WARMUPS,
        report=directory / "report.json",
        resource=directory / "resource.txt",
    )


def _expected_command(tool: str, argv: list[str], directory: Path) -> list[str]:
    profile = str(directory / "profile.txt")
    if tool == "perf":
        return [
            "perf",
            "stat",
            "-x,",
            "-o",
            profile,
            "-e",
            ",".join(PERF_EVENTS),
            "--",
            *argv,
        ]
    require(tool == "strace", f"unknown profile tool: {tool}")
    return ["strace", "-f", "-c", "-e", TRACE_FILTER, "-o", profile, "--", *argv]


def _validate_started(
    started: Mapping[str, Any],
    directory: Path,
    phase: str,
    tool: str,
    workload: str,
    input_mode: str,
    build: Mapping[str, Any],
) -> tuple[list[str], dict[str, Any], dict[str, Any]]:
    label = f"{phase}-{tool}-{input_mode}-{workload}"
    require(started.get("label") == label, f"{directory}/started.json: label differs")
    require(started.get("samples") == SAMPLES and started.get("warmups") == WARMUPS, f"{directory}: sample count differs")
    require(started.get("scope") == "Whole diagnostic child including setup and oracle work; profiler overhead included.", f"{directory}: scope differs")
    expected_arm = _expected_arm(workload, input_mode)
    expected_source = _expected_source(workload, input_mode)
    require(started.get("arm") == expected_arm, f"{directory}: axis arm differs")
    require(started.get("source") == expected_source, f"{directory}: input source binding differs")
    _profile_build_binding(Path(str(build["path"])), started.get("build"), f"{directory}.started.build")
    require(started.get("binary") == build["binary"], f"{directory}: binary binding differs")
    require(started.get("driver") == meta(PROFILE_DRIVER), f"{directory}: profile driver hash differs")
    require(started.get("validators") == routes._script_hashes(), f"{directory}: sealed validator hashes differ")
    argv = _expected_argv(build["binary"], expected_arm, directory)
    require(started.get("argv") == argv, f"{directory}: route argv differs")
    command = _expected_command(tool, argv, directory)
    require(started.get("command") == command, f"{directory}: profiler command differs")
    return argv, expected_arm, expected_source


def _validate_artifact_inventory(receipt: Mapping[str, Any], directory: Path) -> dict[str, dict[str, Any]]:
    actual: dict[str, dict[str, Any]] = {}
    for path in sorted(directory.iterdir()):
        if path.is_symlink() or not path.is_file():
            fail(f"{directory}: unexpected non-regular or directory artifact: {path.name}")
        if path.name == "receipt.json":
            continue
        actual[path.name] = {"bytes": path.stat().st_size, "sha256": sha(path)}
    retained = _dict(receipt.get("artifacts"), f"{directory}/receipt.json.artifacts")
    require(retained == actual, f"{directory}: retained artifact inventory or hash differs")
    for name in ("started.json", "stdout.txt", "stderr.txt"):
        require(name in actual, f"{directory}: required artifact missing: {name}")
    return actual


def _validate_report(
    report_path: Path,
    arm: Mapping[str, Any],
    binary: Mapping[str, Any],
    argv: list[str],
    input_metadata: Mapping[str, Any],
) -> dict[str, Any]:
    try:
        return routes._check_axis_report(
            report_path,
            "normal",
            dict(arm),
            samples=SAMPLES,
            warmups=WARMUPS,
            binary=dict(binary),
            argv=argv,
            input_metadata=dict(input_metadata) if input_metadata["mode"] == "file" else None,
        )
    except Exception as error:
        raise SummaryError(f"{report_path}: sealed 0484 report validator rejected report: {error}") from error


def _report_identity(report: Mapping[str, Any], label: str) -> dict[str, Any]:
    cases = _list(report.get("cases"), f"{label}.cases")
    require(len(cases) == 1, f"{label}: expected one report case")
    case = _dict(cases[0], f"{label}.cases[0]")
    source = _dict(case.get("source"), f"{label}.source")
    authored = _dict(case.get("authored"), f"{label}.authored")
    oracle = _dict(case.get("oracle"), f"{label}.oracle")
    semantic = _dict(oracle.get("candidate_semantic"), f"{label}.oracle.candidate_semantic")
    source_keys = (
        "archive_bytes",
        "archive_sha256",
        "main_xml_bytes",
        "main_xml_sha256",
        "member_count",
        "members",
    )
    for key in source_keys:
        require(key in source, f"{label}.source.{key}: missing")
    authored_keys = (
        "authored_count",
        "chunk_mode",
        "text_mode",
        "max_chunk_bytes",
        "replay_window_bytes",
        "max_encoded_paragraph_bytes",
        "xml_entity_reference_count",
        "text_bytes",
        "encoded_xml_bytes",
        "event_count",
        "text_chunk_count",
        "expected_event_sha256",
        "expected_encoded_sha256",
    )
    for key in authored_keys:
        require(key in authored, f"{label}.authored.{key}: missing")
    oracle_keys = (
        "candidate_archive_bytes",
        "candidate_archive_sha256",
        "candidate_main_xml_bytes",
        "candidate_main_xml_sha256",
    )
    for key in oracle_keys:
        require(key in oracle, f"{label}.oracle.{key}: missing")
    flag_keys = (
        "candidate_xml_exact",
        "candidate_semantic_exact",
        "untouched_member_metadata_exact",
        "untouched_raw_members_preserved",
        "physical_order_exact",
        "opaque_member_exact",
        "source_unchanged",
        "inverse_exact",
    )
    flags = {key: oracle.get(key) for key in flag_keys}
    require(all(value is True for value in flags.values()), f"{label}: content oracle did not pass")
    return {
        "source": {key: source[key] for key in source_keys},
        "authored": {key: authored[key] for key in authored_keys},
        "candidate": {
            "archive_bytes": oracle["candidate_archive_bytes"],
            "archive_sha256": oracle["candidate_archive_sha256"],
            "main_xml_bytes": oracle["candidate_main_xml_bytes"],
            "main_xml_sha256": oracle["candidate_main_xml_sha256"],
            "semantic": semantic,
            "oracle_flags": flags,
        },
    }


def _parse_resource(path: Path) -> dict[str, Any]:
    prefix = "Maximum resident set size (kbytes):"
    require(path.is_file() and not path.is_symlink(), f"{path}: resource receipt missing")
    values: list[int] = []
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip()
        if line.startswith(prefix):
            raw = line[len(prefix) :].strip()
            require(raw.isdigit(), f"{path}: malformed maximum RSS")
            values.append(int(raw) * 1024)
    require(len(values) == 1, f"{path}: expected exactly one GNU time maximum RSS value")
    return {
        "status": "observed",
        "n": 1,
        "bytes": values[0],
        "scope": WHOLE_CHILD_SCOPE,
        "samples": SAMPLES,
        "warmups": WARMUPS,
    }


def _unavailable(reason: str) -> dict[str, Any]:
    return {"status": "unavailable", "value": None, "reason": reason}


def _perf_value(raw: str) -> tuple[int | float | None, str]:
    cleaned = raw.strip()
    lowered = cleaned.lower()
    if not cleaned:
        return None, "missing"
    if "not supported" in lowered or "unsupported" in lowered:
        return None, "unsupported"
    if "not counted" in lowered or "notcounted" in lowered:
        return None, "not_counted"
    if cleaned.startswith("<"):
        return None, "unavailable"
    try:
        numeric = float(cleaned.replace(",", ""))
    except ValueError:
        return None, "invalid"
    if not math.isfinite(numeric):
        return None, "invalid"
    return (int(numeric) if numeric.is_integer() else numeric), "measured"


def _perf_event_index(row: list[str], event: str) -> int | None:
    for index, value in enumerate(row):
        token = value.strip()
        if token == event or token.split(":", 1)[0] == event:
            return index
    return None


def _parse_perf(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{path}: perf artifact missing")
    counters: dict[str, dict[str, Any]] = {}
    unknown: list[str] = []
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    for line in lines:
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        try:
            row = next(csv.reader([line]))
        except csv.Error:
            unknown.append(line)
            continue
        event = next((name for name in PERF_EVENTS if _perf_event_index(row, name) is not None), None)
        if event is None:
            unknown.append(line)
            continue
        index = _perf_event_index(row, event)
        require(index is not None and index > 0, f"{path}: malformed perf row for {event}")
        require(event not in counters, f"{path}: duplicate perf event {event}")
        raw_value = ",".join(row[:index]).strip().strip(",")
        value, status = _perf_value(raw_value)
        counters[event] = {
            "value": value,
            "status": status,
            "raw_value": raw_value,
            "event": row[index].strip(),
            "unit": row[1].strip() if len(row) > 1 else None,
            "line": line,
        }
    for event in PERF_EVENTS:
        if event not in counters:
            counters[event] = {"value": None, "status": "missing", "event": event}
    measured = [event for event in PERF_EVENTS if counters[event]["status"] == "measured"]
    unavailable = [
        event
        for event in PERF_EVENTS
        if counters[event]["status"] in {"unsupported", "not_counted", "unavailable", "missing"}
    ]
    if measured:
        overall = "observed"
    elif any(counters[event]["status"] == "unsupported" for event in PERF_EVENTS):
        overall = "unsupported"
    elif counters and path.stat().st_size > 0:
        overall = "unavailable"
    else:
        overall = "unavailable"
    return {
        "status": overall,
        "n": 1,
        "scope": WHOLE_CHILD_SCOPE,
        "events": counters,
        "measured_events": measured,
        "unavailable_events": unavailable,
        "unknown_lines": unknown,
        "requested_events": list(PERF_EVENTS),
    }


def _parse_count(value: str) -> int | None:
    if not re.fullmatch(r"[0-9][0-9,]*", value):
        return None
    return int(value.replace(",", ""))


def _parse_strace(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{path}: strace artifact missing")
    rows: dict[str, dict[str, Any]] = {}
    total: dict[str, Any] | None = None
    malformed: list[str] = []
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("% time") or stripped.startswith("-"):
            continue
        tokens = stripped.split()
        if not tokens:
            continue
        syscall = tokens[-1]
        if syscall == "total":
            if len(tokens) >= 5:
                calls = _parse_count(tokens[-2])
                if calls is not None:
                    total = {"calls": calls, "line": line}
            continue
        # strace omits the errors column when the filtered syscall has no
        # errors, despite retaining it in the heading.  Accept both the
        # five-token (implicit zero errors) and six-token forms.
        if len(tokens) < 5:
            malformed.append(line)
            continue
        numeric_tokens = tokens[:-1]
        calls = _parse_count(numeric_tokens[-1] if len(numeric_tokens) == 4 else numeric_tokens[-2])
        errors = 0 if len(numeric_tokens) == 4 else _parse_count(numeric_tokens[-1])
        try:
            percent = float(numeric_tokens[0])
            seconds = float(numeric_tokens[1])
            usecs = float(numeric_tokens[2])
        except ValueError:
            calls = None
            errors = None
            percent = seconds = usecs = float("nan")
        if calls is None or errors is None or not all(math.isfinite(value) for value in (percent, seconds, usecs)):
            malformed.append(line)
            continue
        require(syscall not in rows, f"{path}: duplicate strace syscall row {syscall}")
        rows[syscall] = {
            "calls": calls,
            "errors": errors,
            "percent_time": percent,
            "seconds": seconds,
            "usecs_per_call": usecs,
            "line": line,
        }
    require(rows, f"{path}: strace summary has no syscall rows")
    selected: dict[str, Any] = {}
    for syscall in TRACE_SYSCALLS:
        if syscall in rows:
            selected[syscall] = {"status": "observed", **rows[syscall]}
        else:
            # The filter is explicit.  Absence is retained as absent evidence;
            # it is never silently changed into a measured zero.
            selected[syscall] = {
                "status": "absent",
                "calls": None,
                "errors": None,
                "reason": "strace did not emit a row for this filtered syscall",
            }
    return {
        "status": "observed",
        "n": 1,
        "scope": WHOLE_CHILD_SCOPE,
        "filter": TRACE_FILTER,
        "rows": rows,
        "selected": selected,
        "total": total,
        "malformed_lines": malformed,
    }


def _tool_status(receipt: Mapping[str, Any], tool: str, directory: Path) -> tuple[str, str | None]:
    status = receipt.get("status")
    require(status in {"pass", "failed_or_unavailable"}, f"{directory}/receipt.json: unknown status {status!r}")
    process = _dict(receipt.get("process"), f"{directory}/receipt.json.process")
    if status == "pass":
        require(process.get("returncode") == 0 and process.get("timed_out") is False, f"{directory}: passing receipt has failed process")
        require(receipt.get("error") is None, f"{directory}: passing receipt has an error")
        return "pass", None
    error = receipt.get("error")
    require(isinstance(error, str) and error, f"{directory}: unavailable receipt lacks an error")
    return "failed_or_unavailable", error


def _row_from_directory(
    directory: Path,
    phase: str,
    tool: str,
    input_mode: str,
    workload: str,
    build: Mapping[str, Any],
) -> dict[str, Any]:
    started_path = directory / "started.json"
    receipt_path = directory / "receipt.json"
    started = _dict(_read(started_path), str(started_path))
    receipt = _dict(_read(receipt_path), str(receipt_path))
    argv, arm, source = _validate_started(started, directory, phase, tool, workload, input_mode, build)
    artifacts = _validate_artifact_inventory(receipt, directory)
    expected_receipt = {"path": str(receipt_path), **meta(receipt_path)}
    status, error = _tool_status(receipt, tool, directory)
    report_path = directory / "report.json"
    resource_path = directory / "resource.txt"
    profile_path = directory / "profile.txt"
    report: dict[str, Any] | None = None
    identity: dict[str, Any] | None = None
    report_status = "unavailable"
    report_error: str | None = None
    if _nonempty(report_path):
        try:
            report = _validate_report(report_path, arm, build["binary"], argv, source)
            identity = _report_identity(report, str(report_path))
            report_status = "pass"
        except SummaryError:
            raise
        except Exception as caught:
            if status == "pass":
                raise
            report_error = f"{type(caught).__name__}: {caught}"
            report_status = "failed"
    elif status == "pass":
        fail(f"{directory}: passing profile has no nonempty report.json")
    else:
        report_error = "profiler did not retain a report.json"

    resource: dict[str, Any] | None = None
    if _nonempty(resource_path):
        try:
            resource = _parse_resource(resource_path)
        except SummaryError:
            if status == "pass":
                raise
            report_error = report_error or f"invalid resource receipt: {resource_path}"
    elif status == "pass":
        fail(f"{directory}: passing profile has no nonempty resource.txt")

    profiler_data: dict[str, Any] | None = None
    if _nonempty(profile_path):
        profiler_data = _parse_perf(profile_path) if tool == "perf" else _parse_strace(profile_path)
    elif status == "pass":
        fail(f"{directory}: passing profile has no nonempty profile.txt")
    else:
        profiler_data = _unavailable("profiler artifact was not retained")

    process = _dict(receipt.get("process"), f"{receipt_path}.process")
    require(receipt.get("label") == started.get("label"), f"{receipt_path}: receipt label differs")
    require(receipt.get("argv") == started.get("argv"), f"{receipt_path}: receipt argv differs")
    require(receipt.get("binary") == started.get("binary"), f"{receipt_path}: receipt binary differs")
    require(receipt.get("driver") == started.get("driver"), f"{receipt_path}: receipt driver differs")
    return {
        "label": started["label"],
        "phase": phase,
        "tool": tool,
        "input_mode": input_mode,
        "workload": workload,
        "status": status,
        "error": error,
        "process": {
            "returncode": process.get("returncode"),
            "timed_out": process.get("timed_out"),
            "route_started": process.get("route_started"),
        },
        "command": list(started["command"]),
        "argv": list(argv),
        "receipt": expected_receipt,
        "artifacts": artifacts,
        "build": {
            "path": str(build["path"]),
            "sha256": build["record_sha256"],
            "binary": build["binary"],
        },
        "source": source,
        "arm": arm,
        "scope": {
            "samples": SAMPLES,
            "warmups": WARMUPS,
            "observation_count": 1,
            "description": WHOLE_CHILD_SCOPE,
            "logical_source_read_substitution": False,
        },
        "report": {
            "status": report_status,
            "path": str(report_path),
            "artifact": {"bytes": report_path.stat().st_size, "sha256": sha(report_path)} if report_path.is_file() else None,
            "identity": identity,
            "error": report_error,
        },
        "resource": resource or {"status": "unavailable", "reason": "resource.txt was not retained"},
        "profiler": profiler_data,
        "validator": {
            "driver": {"path": str(PROFILE_DRIVER), **meta(PROFILE_DRIVER)},
            "sealed_route_validators": routes._script_hashes(),
            "report": report_status,
            "artifacts": "receipt.artifacts matched every retained regular file",
        },
    }


def _load_profile_rows(attempt: str, builds: Mapping[str, Mapping[str, Any]]) -> list[dict[str, Any]]:
    root = ROOT / "profiles" / attempt
    require(root.is_dir() and not root.is_symlink(), f"profile output directory is missing: {root}")
    result_path = root / "result.json"
    result = _dict(_read(result_path), str(result_path))
    expected_labels = _expected_labels()
    records = _list(result.get("records"), f"{result_path}.records")
    require([_dict(item, "result.record").get("label") for item in records] == expected_labels, f"{result_path}: record inventory/order differs")
    actual_dirs = sorted(path.name for path in root.iterdir() if path.is_dir() and not path.is_symlink())
    require(sorted(expected_labels) == actual_dirs, f"{root}: profile child inventory differs")
    root_files = sorted(path.name for path in root.iterdir() if path.is_file())
    require(root_files == ["result.json"], f"{root}: unexpected root artifact(s): {root_files}")
    expected_complete = True
    rows: list[dict[str, Any]] = []
    record_by_label = {_dict(item, "result.record")["label"]: _dict(item, "result.record") for item in records}
    for phase in PHASES:
        for workload in WORKLOADS:
            for input_mode in INPUT_MODES:
                tools_for_phase = ("perf",) if phase == "before" else ("perf", "strace")
                for tool in tools_for_phase:
                    label = f"{phase}-{tool}-{input_mode}-{workload}"
                    record = record_by_label[label]
                    receipt_path = root / label / "receipt.json"
                    binding = _dict(record.get("receipt"), f"{result_path}.{label}.receipt")
                    require(binding.get("path") == str(receipt_path), f"{label}: result receipt path differs")
                    require(binding.get("bytes") == meta(receipt_path)["bytes"], f"{label}: result receipt byte count differs")
                    require(binding.get("sha256") == meta(receipt_path)["sha256"], f"{label}: result receipt hash differs")
                    build = builds[phase]
                    row = _row_from_directory(root / label, phase, tool, input_mode, workload, build)
                    require(record.get("status") == row["status"], f"{label}: result and receipt statuses differ")
                    expected_complete = expected_complete and row["status"] == "pass"
                    rows.append(row)
    expected_result_status = "pass" if expected_complete else "incomplete"
    require(result.get("status") == expected_result_status, f"{result_path}: overall status differs")
    return rows


def _identity_pair(before: Mapping[str, Any], after: Mapping[str, Any], label: str) -> dict[str, Any]:
    before_identity = before.get("report", {}).get("identity")
    after_identity = after.get("report", {}).get("identity")
    if not isinstance(before_identity, dict) or not isinstance(after_identity, dict):
        return {"label": label, "status": "unavailable", "reason": "one report identity is unavailable"}
    source_equal = before_identity["source"] == after_identity["source"]
    authored_equal = before_identity["authored"] == after_identity["authored"]
    candidate_main_equal = (
        before_identity["candidate"]["main_xml_bytes"] == after_identity["candidate"]["main_xml_bytes"]
        and before_identity["candidate"]["main_xml_sha256"] == after_identity["candidate"]["main_xml_sha256"]
    )
    semantic_equal = before_identity["candidate"]["semantic"] == after_identity["candidate"]["semantic"]
    oracle_equal = before_identity["candidate"]["oracle_flags"] == after_identity["candidate"]["oracle_flags"]
    archive_equal = (
        before_identity["candidate"]["archive_bytes"] == after_identity["candidate"]["archive_bytes"]
        and before_identity["candidate"]["archive_sha256"] == after_identity["candidate"]["archive_sha256"]
    )
    primary = source_equal and authored_equal and candidate_main_equal and semantic_equal and oracle_equal
    return {
        "label": label,
        "status": "pass" if primary else "identity_mismatch",
        "source_equal": source_equal,
        "authored_proof_equal": authored_equal,
        "candidate_main_xml_equal": candidate_main_equal,
        "candidate_semantic_equal": semantic_equal,
        "unchanged_member_oracles_equal": oracle_equal,
        "candidate_archive_equal": archive_equal,
        "candidate_archive_difference_is_recorded": not archive_equal,
        "before": before_identity,
        "after": after_identity,
    }


def _rows_by_key(rows: Iterable[Mapping[str, Any]], tool: str) -> dict[tuple[str, str, str], Mapping[str, Any]]:
    return {
        (row["phase"], row["workload"], row["input_mode"]): row
        for row in rows
        if row["tool"] == tool
    }


def _perf_comparisons(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    by_key = _rows_by_key(rows, "perf")
    comparisons: list[dict[str, Any]] = []
    rss_rows: list[dict[str, Any]] = []
    for workload in WORKLOADS:
        for input_mode in INPUT_MODES:
            key_before = ("before", workload, input_mode)
            key_after = ("after", workload, input_mode)
            before = by_key[key_before]
            after = by_key[key_after]
            identity = _identity_pair(before, after, f"perf-{input_mode}-{workload}")
            for event in PERF_EVENTS:
                before_counter = ((before.get("profiler") or {}).get("events") or {}).get(event, {})
                after_counter = ((after.get("profiler") or {}).get("events") or {}).get(event, {})
                before_value = before_counter.get("value") if before_counter.get("status") == "measured" else None
                after_value = after_counter.get("value") if after_counter.get("status") == "measured" else None
                change = _percent_change(before_value, after_value)
                comparisons.append({
                    "workload": workload,
                    "input_mode": input_mode,
                    "event": event,
                    "status": "observed" if before_value is not None and after_value is not None else "unavailable",
                    "before_status": before_counter.get("status"),
                    "after_status": after_counter.get("status"),
                    "change": change,
                    "review_threshold_percent": PERF_REVIEW_PERCENT,
                    "review_flag": _review(change, PERF_REVIEW_PERCENT),
                    "scope": WHOLE_CHILD_SCOPE,
                })
            before_rss = before.get("resource", {}).get("bytes") if before.get("resource", {}).get("status") == "observed" else None
            after_rss = after.get("resource", {}).get("bytes") if after.get("resource", {}).get("status") == "observed" else None
            change = _percent_change(before_rss, after_rss)
            rss_rows.append({
                "workload": workload,
                "input_mode": input_mode,
                "identity_status": identity.get("status"),
                "change": change,
                "review_threshold_percent": RSS_REVIEW_PERCENT,
                "review_flag": _review(change, RSS_REVIEW_PERCENT),
                "scope": WHOLE_CHILD_SCOPE,
                "n_before": 1 if before_rss is not None else 0,
                "n_after": 1 if after_rss is not None else 0,
            })
    return comparisons, rss_rows


def _baseline_path(workload: str, input_mode: str) -> Path:
    """Return the retained 0485 after-strace child for one matched arm."""

    return RETAINED_PROFILES / f"after-strace-{input_mode}-{workload}"


def _load_baseline(workload: str, input_mode: str, build: Mapping[str, Any]) -> dict[str, Any]:
    """Validate one immutable 0485 after diagnostic as a syscall baseline."""

    directory = _baseline_path(workload, input_mode)
    if not directory.is_dir():
        return {"status": "unavailable", "reason": f"retained 0485 after profile is missing: {directory}"}
    receipt_path = directory / "receipt.json"
    started_path = directory / "started.json"
    receipt = _dict(_read(receipt_path), str(receipt_path))
    started = _dict(_read(started_path), str(started_path))
    require(receipt.get("status") == "pass", f"{receipt_path}: retained 0485 after profile did not pass")
    process = _dict(receipt.get("process"), f"{receipt_path}.process")
    require(process.get("returncode") == 0 and process.get("timed_out") is False, f"{receipt_path}: retained 0485 process failed")
    actual = _validate_artifact_inventory(receipt, directory)
    for name in ("report.json", "resource.txt", "profile.txt"):
        require(name in actual and actual[name]["bytes"] > 0, f"{directory}: retained 0485 artifact is missing or empty: {name}")
    require(started.get("label") == f"after-strace-{input_mode}-{workload}", f"{started_path}: retained label differs")
    require(started.get("samples") == SAMPLES and started.get("warmups") == WARMUPS, f"{started_path}: retained sample counts differ")
    require(started.get("scope") == "Whole diagnostic child including setup and oracle work; profiler overhead included.", f"{started_path}: retained scope differs")
    arm = _expected_arm(workload, input_mode)
    source = _expected_source(workload, input_mode)
    require(started.get("arm") == arm, f"{started_path}: retained 0485 arm differs")
    require(started.get("source") == source, f"{started_path}: retained 0485 source binding differs")
    require(build["phase"] == "before", "retained 0485 after profile must bind the before build")
    binary = _dict(started.get("binary"), f"{started_path}.binary")
    require(binary == build["binary"], f"{started_path}: retained binary differs from 0485 build")
    _profile_build_binding(BEFORE_BUILD, started.get("build"), f"{started_path}.build")
    require(started.get("driver") == meta(RETAINED_PROFILE_DRIVER), f"{started_path}: retained profile driver hash differs")
    require(started.get("validators") == routes._script_hashes(), f"{started_path}: sealed route validator hashes differ")
    argv = _expected_argv(binary, arm, directory)
    require(started.get("argv") == argv, f"{started_path}: retained route argv differs")
    command = _list(started.get("command"), f"{started_path}.command")
    require(command == _expected_command("strace", argv, directory), f"{started_path}: retained strace command differs")
    report = _validate_report(directory / "report.json", arm, binary, argv, source)
    identity = _report_identity(report, str(directory / "report.json"))
    profile = _parse_strace(directory / "profile.txt")
    resource = _parse_resource(directory / "resource.txt")
    require(receipt.get("label") == started.get("label"), f"{receipt_path}: retained label differs")
    require(receipt.get("argv") == started.get("argv"), f"{receipt_path}: retained argv differs")
    require(receipt.get("binary") == started.get("binary"), f"{receipt_path}: retained binary differs")
    require(receipt.get("driver") == started.get("driver"), f"{receipt_path}: retained driver differs")
    return {
        "status": "pass",
        "directory": str(directory),
        "receipt": {"path": str(receipt_path), **meta(receipt_path)},
        "started": started,
        "source": source,
        "identity": identity,
        "profile": profile,
        "resource": resource,
        "trace_executable": str(command[0]),
    }


def _trace_path(value: str) -> str | None:
    if os.path.isabs(value):
        path = Path(value)
    else:
        resolved = shutil.which(value)
        if resolved is None:
            return None
        path = Path(resolved)
    return str(path.resolve()) if path.exists() else None


def _strace_comparisons(rows: list[dict[str, Any]], build: Mapping[str, Any]) -> list[dict[str, Any]]:
    after_rows = _rows_by_key(rows, "strace")
    comparisons: list[dict[str, Any]] = []
    for workload in WORKLOADS:
        for input_mode in INPUT_MODES:
            key = ("after", workload, input_mode)
            after = after_rows[key]
            baseline = _load_baseline(workload, input_mode, build)
            base_profile = baseline.get("profile") if baseline.get("status") == "pass" else None
            candidate_profile = after.get("profiler") if isinstance(after.get("profiler"), dict) else None
            item: dict[str, Any] = {
                "workload": workload,
                "input_mode": input_mode,
                "baseline": {
                    "status": baseline.get("status"),
                    "directory": baseline.get("directory"),
                    "receipt": baseline.get("receipt"),
                },
                "candidate": {"label": after["label"], "status": after["status"]},
                "scope": WHOLE_CHILD_SCOPE,
            }
            if base_profile is None or candidate_profile is None or candidate_profile.get("status") != "observed":
                item["status"] = "unavailable"
                item["reason"] = baseline.get("reason") or "after strace summary is unavailable"
                comparisons.append(item)
                continue
            after_command = after.get("command", [])
            baseline_command = baseline.get("started", {}).get("command", [])
            candidate_executable = _trace_path(str(after_command[0])) if after_command else None
            baseline_executable = _trace_path(str(baseline_command[0])) if baseline_command else None
            path_equal = candidate_executable is not None and candidate_executable == baseline_executable
            candidate_filter = after_command[after_command.index("-e") + 1] if "-e" in after_command else None
            baseline_filter = baseline_command[baseline_command.index("-e") + 1] if "-e" in baseline_command else None
            filter_equal = candidate_filter == TRACE_FILTER and baseline_filter == TRACE_FILTER
            base_selected = base_profile.get("selected", {})
            candidate_selected = candidate_profile.get("selected", {})
            counts: list[dict[str, Any]] = []
            for syscall in TRACE_SYSCALLS:
                base_value = base_selected.get(syscall, {}).get("calls")
                after_value = candidate_selected.get(syscall, {}).get("calls")
                change = _percent_change(base_value, after_value)
                counts.append({
                    "syscall": syscall,
                    "baseline_status": base_selected.get(syscall, {}).get("status"),
                    "after_status": candidate_selected.get(syscall, {}).get("status"),
                    "change": change,
                })
            identity = _identity_pair(
                {"report": {"identity": baseline.get("identity")}},
                {"report": {"identity": after.get("report", {}).get("identity")}},
                f"strace-{input_mode}-{workload}",
            )
            input_exact = after.get("source") == baseline.get("source")
            item.update({
                "status": "pass" if identity.get("status") == "pass" and input_exact and path_equal and filter_equal else "identity_mismatch",
                "trace_filter": TRACE_FILTER,
                "trace_executable": {
                    "candidate_raw": after_command[0] if after_command else None,
                    "baseline_raw": baseline_command[0] if baseline_command else None,
                    "candidate_resolved": candidate_executable,
                    "baseline_resolved": baseline_executable,
                    "equal": path_equal,
                },
                "trace_filter_equal": filter_equal,
                "input_binding_equal": input_exact,
                "identity": identity,
                "counts": counts,
                "total": {
                    "baseline": base_profile.get("total"),
                    "after": candidate_profile.get("total"),
                },
                "exact_path_policy": "workload/input arm, prepared fixture identity, report source identity, and trace filter must match; binary paths may differ by phase",
            })
            comparisons.append(item)
    return comparisons


def _source_comparison(builds: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    before = builds["before"]["source_map"]
    after = builds["after"]["source_map"]
    changed: list[dict[str, Any]] = []
    for name in sorted(set(before) | set(after)):
        if before.get(name) != after.get(name):
            changed.append({"path": name, "before": before.get(name), "after": after.get(name)})
    return {
        "equal": not changed,
        "changed_files": changed,
        "before_manifest": builds["before"]["source_before"],
        "after_manifest": builds["after"]["source_after"],
        "review_required": True,
    }


def _markdown(summary: Mapping[str, Any]) -> str:
    lines = [
        "# 0487 profile summary",
        "",
        f"Status: **{summary['status']}**; profile validation: **{summary['validation_status']}**.",
        "",
        "The 12 rows below are one whole diagnostic child each, with one sample and one warmup. The child includes setup and oracle work; profiler overhead is included. Missing PMU values remain unavailable. No elapsed-time or causal performance claim is made.",
        "",
        f"Profile attempt: `{summary['attempt']}`. Before build: `{summary['builds']['before']['record_sha256']}`. After build: `{summary['builds']['after']['record_sha256']}`.",
        "",
        "## Child rows",
        "",
        "| phase | tool | input | workload | receipt | report | RSS bytes | profile status |",
        "|---|---|---|---|---|---|---:|---|",
    ]
    for row in summary["children"]:
        rss = row["resource"].get("bytes", "—") if row["resource"].get("status") == "observed" else "—"
        report_status = row["report"].get("status", "unavailable")
        profile_status = (row.get("profiler") or {}).get("status", "unavailable")
        lines.append(f"| {row['phase']} | {row['tool']} | {row['input_mode']} | {row['workload']} | {row['status']} | {report_status} | {rss} | {profile_status} |")
    lines.extend([
        "",
        "## Perf counters",
        "",
        "Each row is a matched before/after whole-child observation. A review flag marks an absolute change of at least 5%; it is a review signal, not a performance claim.",
        "",
        "| input | workload | event | before | before status | after | after status | percent change | review |",
        "|---|---|---|---:|---|---:|---|---:|---|",
    ])
    for item in summary["perf_comparisons"]:
        change = item["change"]
        lines.append(f"| {item['input_mode']} | {item['workload']} | {item['event']} | {change.get('before', '—')} | {item.get('before_status', '—')} | {change.get('after', '—')} | {item.get('after_status', '—')} | {change.get('percent_change', '—')} | {item.get('review_flag', '—')} |")
    lines.extend([
        "",
        "## RSS review",
        "",
        "GNU time contributes one maximum RSS observation per child; zero-baseline percentage changes are left undefined.",
        "",
        "| input | workload | before | after | percent change | review |",
        "|---|---|---:|---:|---:|---|",
    ])
    for item in summary["rss_comparisons"]:
        change = item["change"]
        lines.append(f"| {item['input_mode']} | {item['workload']} | {change.get('before', '—')} | {change.get('after', '—')} | {change.get('percent_change', '—')} | {item.get('review_flag', '—')} |")
    lines.extend([
        "",
        "## After strace counts",
        "",
        "The after strace rows are compared with retained 0485 after diagnostics only when the arm, prepared fixture identity, source identity, report identity, and trace filter bind exactly. Candidate archive framing differences are retained separately.",
        "",
        "| input | workload | status | statx before/after | pread64 before/after |",
        "|---|---|---|---:|---:|",
    ])
    for item in summary["strace_comparisons"]:
        selected = {row["syscall"]: row for row in item.get("counts", [])}
        def pair(name: str) -> str:
            change = selected.get(name, {}).get("change", {})
            return f"{change.get('before', '—')} / {change.get('after', '—')}"
        lines.append(f"| {item['input_mode']} | {item['workload']} | {item['status']} | {pair('statx')} | {pair('pread64')} |")
    lines.extend([
        "",
        "## Content identity",
        "",
        "| pair | status | source | authored proof | candidate main XML | semantic identity | unchanged oracles | candidate archive |",
        "|---|---|---|---|---|---|---|---|",
    ])
    for item in summary["identity_comparisons"]:
        lines.append(
            f"| {item['label']} | {item['status']} | {item.get('source_equal', '—')} | {item.get('authored_proof_equal', '—')} | {item.get('candidate_main_xml_equal', '—')} | {item.get('candidate_semantic_equal', '—')} | {item.get('unchanged_member_oracles_equal', '—')} | {item.get('candidate_archive_equal', '—')} |"
        )
    lines.extend([
        "",
        "## Custody",
        "",
        f"Source manifests differ in {len(summary['source_comparison']['changed_files'])} path(s); the full path/hash diff is retained in JSON for review.",
        "",
        "The summarizer binds the profile driver, sealed 0484 route validator hashes, the retained 0485 before build, the new 0487 after build, build gate receipts, binary hashes, report artifacts, resource artifacts, and profiler artifacts. It performs no capture or build work.",
        "",
    ])
    return "\n".join(lines)


def summarize(attempt: str, output_json: Path | None = None, output_md: Path | None = None) -> dict[str, Any]:
    require(ATTEMPT_RE.fullmatch(attempt) is not None, "attempt must be a path-safe token")
    output_json = (output_json or (ROOT / "profiles" / f"{attempt}-summary.json")).resolve()
    output_md = (output_md or (ROOT / "profiles" / f"{attempt}-summary.md")).resolve()
    require(not output_json.exists() and not output_md.exists(), "refusing to replace an existing profile summary")
    builds = {phase: _load_build(phase) for phase in PHASES}
    children = _load_profile_rows(attempt, builds)
    perf_comparisons, rss_comparisons = _perf_comparisons(children)
    identity_comparisons = []
    by_perf = _rows_by_key(children, "perf")
    for workload in WORKLOADS:
        for input_mode in INPUT_MODES:
            identity_comparisons.append(_identity_pair(
                by_perf[("before", workload, input_mode)],
                by_perf[("after", workload, input_mode)],
                f"perf-{input_mode}-{workload}",
            ))
    strace_comparisons = _strace_comparisons(children, builds["before"])
    summary: dict[str, Any] = {
        "schema": SCHEMA,
        "version": VERSION,
        "attempt": attempt,
        "status": "complete",
        "validation_status": "pass",
        "availability": {
            "unsupported_or_unavailable_pmu": [
                row["label"]
                for row in children
                if row["tool"] == "perf" and (row.get("profiler") or {}).get("status") != "observed"
            ],
            "failed_or_unavailable_children": [row["label"] for row in children if row["status"] != "pass"],
        },
        "attempt_output": str((ROOT / "profiles" / attempt).resolve()),
        "inventory": {
            "expected_children": 12,
            "observed_children": len(children),
            "before_perf": 4,
            "after_perf": 4,
            "after_strace": 4,
            "samples_per_child": SAMPLES,
            "warmups_per_child": WARMUPS,
            "scope": WHOLE_CHILD_SCOPE,
        },
        "builds": builds,
        "helpers": {
            "summarizer": {"path": str(SUMMARIZER), **meta(SUMMARIZER)},
            "profile_driver": {"path": str(PROFILE_DRIVER), **meta(PROFILE_DRIVER)},
            "sealed_route_validators": routes._script_hashes(),
        },
        "source_comparison": _source_comparison(builds),
        "children": children,
        "perf_comparisons": perf_comparisons,
        "rss_comparisons": rss_comparisons,
        "strace_comparisons": strace_comparisons,
        "identity_comparisons": identity_comparisons,
        "limitations": [
            "Each profiler row is one whole-child observation with one warmup; no percentile or statistical claim is made.",
            "Perf PMU values that are unsupported, not counted, missing, or unavailable remain null with an explicit status.",
            "GNU time RSS is n=1 per child and percentage change is undefined for a zero baseline.",
            "The summary does not substitute a logical source read for the measured route input; source archive and main XML identities are retained.",
            "Candidate ZIP archive framing may differ even when source, candidate main XML, semantic identity, and unchanged-member oracles match; archive differences remain visible.",
        ],
        "generated_utc": now(),
        "outputs": {"json": str(output_json), "markdown": str(output_md)},
    }
    _finite(summary)
    output_json.parent.mkdir(parents=True, exist_ok=True)
    output_md.parent.mkdir(parents=True, exist_ok=True)
    write(output_json, summary)
    try:
        with output_md.open("x", encoding="utf-8", newline="\n") as stream:
            stream.write(_markdown(summary))
    except FileExistsError as error:
        raise SummaryError(f"refusing to replace existing profile Markdown summary: {output_md}") from error
    print(f"wrote {output_json}")
    print(f"wrote {output_md}")
    return summary


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("attempt", help="profile output attempt under change-0487/profiles")
    parser.add_argument("--output-json", type=Path)
    parser.add_argument("--output-md", type=Path)
    args = parser.parse_args(argv)
    try:
        summarize(args.attempt, args.output_json, args.output_md)
    except (SummaryError, OSError, ValueError, KeyError, json.JSONDecodeError) as error:
        print(f"summarize_profiles.py: FAIL: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
