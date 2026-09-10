#!/usr/bin/env python3
"""Capture bounded syscall evidence for the DOCX tail-append publication arms.

The helper deliberately sits outside the formal timing matrix.  It launches
the retained after executable as a complete child under ``strace`` and keeps
the benchmark's report, the unmodified trace, and terminal custody together.
The binary owns all source, preflight, candidate, inverse, and destination
oracles.  This driver only checks that those oracles passed and classifies
observed syscall paths; it does not attribute a syscall to a timed operation
or make a performance claim.

The outer coordinator must hold the shared CPU lock.  This file never takes
that lock.  Each child is pinned to CPU 2, receives a fresh private TMPDIR
and replay directory, and runs in its own process group so timeout cleanup is
bounded and auditable.
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
import subprocess
import sys
from typing import Any, Iterable, Mapping


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TEMP = Path("/home/zhuhe/.cache/litchi-goal-0497")
CPU = 2
CPU_LOCK = "/home/zhuhe/.cache/litchi-goal-0484/cpu.lock"
VERSION = 1

SCHEMA = "docx-tail-append-publication-profile-v1"
TERMINAL_SCHEMA = "docx-tail-append-publication-profile-terminal-v1"
RESULT_SCHEMA = "docx-tail-append-publication-profile-result-v1"
CLEANUP_SCHEMA = "docx-tail-append-publication-profile-cleanup-v1"
REPORT_SCHEMA = "docx-replayable-tail-append-v1"
PUBLICATION_SCHEMA = "docx-replayable-tail-append-publication-v1"
COUNTING_TIMING_SCOPE = "source_admission_prepare_sequential_sink_publication_drop"
ATOMIC_TIMING_SCOPE = (
    "source_admission_prepare_atomic_write_data_sync_rename_parent_directory_sync_publication_drop"
)
COUNTING_VERIFICATION_SCOPE = "timed_production_artifact_proof_plus_untimed_candidate_oracle"
ATOMIC_VERIFICATION_SCOPE = "timed_production_artifact_proof_plus_post_timer_path_oracle"
INVERSE_ORACLE_SCOPE = (
    "untimed_fixture_publication_inverse_exact; timed_atomic_publication_inverse_not_reexecuted"
)

DEFAULT_SAMPLES = 1
DEFAULT_WARMUPS = 1
DEFAULT_TIMEOUT_SECONDS = 900
TERM_GRACE_SECONDS = 10
DEFAULT_SOURCE_COUNT = 8_192
DEFAULT_AUTHORED_COUNT = 256
DEFAULT_CHUNK = "window"
DEFAULT_TEXT = "near"
DEFAULT_PROVIDER = "file-store"
DEFAULT_REPLAY_SYNC = "data"
DEFAULT_REPLAY_MAX_BYTES = 64 * 1024 * 1024

MODES = ("counting", "atomic")
MODE_ROUTE = {"counting": "counting_sink", "atomic": "atomic_path"}
SYSCALLS = (
    "open",
    "openat",
    "openat2",
    "creat",
    "read",
    "pread64",
    "preadv",
    "readv",
    "write",
    "pwrite64",
    "pwritev",
    "writev",
    "close",
    "fsync",
    "fdatasync",
    "rename",
    "renameat",
    "renameat2",
    "unlink",
    "unlinkat",
    "stat",
    "statx",
    "fstat",
    "fstatat",
    "newfstatat",
    "mkdir",
    "mkdirat",
)
SYNC_CALLS = ("fsync", "fdatasync")
RENAME_CALLS = ("rename", "renameat", "renameat2")
WRITE_CALLS = ("write", "pwrite64", "pwritev", "writev")
READ_CALLS = ("read", "pread64", "preadv", "readv")
ORACLE_FLAGS = (
    "candidate_xml_exact",
    "candidate_semantic_exact",
    "untouched_member_metadata_exact",
    "untouched_raw_members_preserved",
    "physical_order_exact",
    "opaque_member_exact",
    "source_unchanged",
    "inverse_exact",
)
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
TOKEN_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")

SCOPE = (
    "whole strace child including benchmark setup, corpus construction, "
    "preflight, warmups, measured samples, post-timer destination/oracle "
    "checks, cleanup, and report serialization"
)
NO_CLAIM = (
    "diagnostic syscall paths and counts only; no operation attribution, "
    "latency comparison, throughput comparison, or optimization claim"
)
CALLER_GATE = {
    "required": True,
    "driver": "root-owned change-0497 gate.py plus flock",
    "lock": CPU_LOCK,
    "ownership": "outer caller runs flock -x for the complete profiler command",
}

TRACE_PATH_RE = re.compile(r"<([^>\n]+)>")
TRACE_QUOTED_PATH_RE = re.compile(r'"((?:\\.|[^"\\])*)"')
TRACE_SYSCALL_RE = re.compile(r"(?<![A-Za-z0-9_])([a-z][a-z0-9_]*)\(")


class ProfileError(RuntimeError):
    """A fail-closed profile or custody invariant failed."""


def fail(message: str) -> None:
    raise ProfileError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


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
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    _finite(value, str(path))
    return value


def _sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def _meta(path: Path) -> dict[str, Any]:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": _sha(path)}


def _artifact(path: Path, *, executable: bool = False) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing regular artifact: {path}")
    if executable:
        require(os.access(path, os.X_OK), f"artifact is not executable: {path}")
    value: dict[str, Any] = _meta(path)
    if executable:
        value["executable"] = True
    return value


def _write_exclusive(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except FileExistsError as error:
        raise ProfileError(f"refusing to replace retained artifact: {path}") from error


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _timestamp(value: Any, label: str) -> _datetime.datetime:
    require(isinstance(value, str) and value, f"{label}: timestamp missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: invalid timestamp: {error}")
    require(parsed.tzinfo is not None, f"{label}: timestamp has no timezone")
    return parsed


def _safe_token(value: str, label: str = "attempt") -> str:
    require(bool(TOKEN_RE.fullmatch(value)), f"{label} must be a path-safe token")
    return value


def _tool_executable(name: str) -> dict[str, Any]:
    require(name == "strace", f"unsupported profiler tool: {name}")
    located = shutil.which(name)
    require(isinstance(located, str) and located, "strace is unavailable")
    try:
        path = Path(located).resolve(strict=True)
    except OSError as error:
        fail(f"cannot resolve strace: {error}")
    return _artifact(path, executable=True)


def _current_revision() -> str:
    try:
        revision = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=REPO, text=True, stderr=subprocess.PIPE
        ).strip()
    except (OSError, subprocess.SubprocessError) as error:
        fail(f"cannot determine source revision: {error}")
    require(REVISION_RE.fullmatch(revision) is not None, "source revision is malformed")
    return revision


def _walk_mappings(value: Any) -> Iterable[Mapping[str, Any]]:
    if isinstance(value, Mapping):
        yield value
        for child in value.values():
            yield from _walk_mappings(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk_mappings(child)


def _build_entry(value: Any, role: str) -> Mapping[str, Any]:
    candidates: list[Mapping[str, Any]] = []
    for mapping in _walk_mappings(value):
        binary = mapping.get("binary")
        if isinstance(binary, Mapping) and isinstance(binary.get("path"), str):
            candidates.append(mapping)
    require(candidates, "build receipt has no binary binding")
    for mapping in candidates:
        binary = mapping.get("binary")
        binary_path = binary.get("path", "") if isinstance(binary, Mapping) else ""
        if (
            role in str(mapping.get("role", ""))
            or role in str(mapping.get("label", ""))
            or f"/{role}/" in str(binary_path)
        ):
            return mapping
    return candidates[0]


def _build_binding(path: Path, role: str) -> tuple[dict[str, Any], dict[str, Any]]:
    raw = _read_json(path)
    require(isinstance(raw, Mapping), f"{path}: build receipt is not an object")
    entry = _build_entry(raw, role)
    binary_value = entry.get("binary")
    require(isinstance(binary_value, Mapping), f"{path}: binary binding is malformed")
    binary_path = Path(str(binary_value.get("path"))).resolve()
    binary = _artifact(binary_path, executable=True)
    expected_sha = binary_value.get("sha256")
    if isinstance(expected_sha, str):
        require(binary["sha256"] == expected_sha, f"{path}: retained binary hash changed")
    expected_bytes = binary_value.get("bytes")
    if isinstance(expected_bytes, int):
        require(binary["bytes"] == expected_bytes, f"{path}: retained binary size changed")
    revision = entry.get("git_revision") or entry.get("source_revision")
    if not isinstance(revision, str):
        revision = raw.get("git_revision") if isinstance(raw, Mapping) else None
    require(isinstance(revision, str) and REVISION_RE.fullmatch(revision),
            f"{path}: build revision is missing or malformed")
    source = entry.get("source") or entry.get("source_manifest") or raw.get("source")
    binding = {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": _sha(path),
        "role": role,
        "binary": binary,
        "git_revision": revision,
        "source": source,
    }
    return binding, binary


def _path_text(path: Path) -> str:
    return str(path.resolve())


def _benchmark_argv(
    *,
    binary: Path,
    report: Path,
    replay_dir: Path,
    mode: str,
    samples: int,
    warmups: int,
    source_count: int,
    authored_count: int,
    chunk: str,
    text: str,
    provider: str,
    replay_sync: str,
    replay_max_bytes: int,
    publication_flag: str,
    publication_path_flag: str | None,
    destination: Path,
) -> list[str]:
    values = [
        "/usr/bin/taskset",
        "-c",
        str(CPU),
        str(binary),
        "--source-counts",
        str(source_count),
        "--authored-counts",
        str(authored_count),
        "--chunks",
        chunk,
        "--text",
        text,
        "--samples",
        str(samples),
        "--warmups",
        str(warmups),
        "--sink-write",
        "4096",
        "--authored-provider",
        provider,
        "--replay-max-bytes",
        str(replay_max_bytes),
        "--replay-sync",
        replay_sync,
        "--replay-dir",
        str(replay_dir),
        publication_flag,
        MODE_ROUTE[mode],
        "--json",
        str(report),
    ]
    if publication_path_flag is not None:
        values.extend([publication_path_flag, str(destination)])
    return values


def _trace_path_tokens(line: str) -> list[str]:
    paths: set[str] = set()
    for match in TRACE_PATH_RE.finditer(line):
        token = match.group(1)
        if token and not token.startswith("unfinished") and not token.startswith("detached"):
            paths.add(token)
    for match in TRACE_QUOTED_PATH_RE.finditer(line):
        token = match.group(1)
        try:
            decoded = json.loads('"' + token + '"')
        except (TypeError, ValueError, json.JSONDecodeError):
            decoded = token
        if decoded.startswith("/") or decoded in {".", ".."}:
            paths.add(decoded)
    return sorted(paths)


def _path_scope(path: str, *, tmpdir: Path, replay_dir: Path, report: Path) -> str:
    resolved_tmp = _path_text(tmpdir)
    resolved_replay = _path_text(replay_dir)
    resolved_report = _path_text(report)
    if path == resolved_report:
        return "report"
    if path == resolved_replay or path.startswith(resolved_replay + os.sep):
        return "authored_file_store"
    if path == resolved_tmp or path.startswith(resolved_tmp + os.sep):
        return "output_atomic"
    return "other"


def _parse_trace(
    path: Path,
    *,
    mode: str,
    tmpdir: Path,
    replay_dir: Path,
    report: Path,
    expected_atomic_paths: Iterable[Mapping[str, Any]] | None = None,
    warmups: int = 0,
    samples: int | None = None,
) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing raw strace: {path}")
    raw = path.read_bytes()
    require(raw, f"empty raw strace: {path}")
    text = raw.decode("utf-8", errors="replace")
    counts = {name: {"calls": 0, "errors": 0} for name in SYSCALLS}
    scoped = {
        scope: {name: {"calls": 0, "errors": 0} for name in SYSCALLS}
        for scope in ("authored_file_store", "output_atomic", "report", "other")
    }
    events: list[dict[str, Any]] = []
    lines = text.splitlines()
    for number, line in enumerate(lines, start=1):
        match = TRACE_SYSCALL_RE.search(line)
        if match is None:
            continue
        name = match.group(1)
        if name not in counts:
            continue
        failed = bool(re.search(r"\s=\s-1(?:\s|$)", line))
        counts[name]["calls"] += 1
        if failed:
            counts[name]["errors"] += 1
        paths = _trace_path_tokens(line)
        scopes = {_path_scope(item, tmpdir=tmpdir, replay_dir=replay_dir, report=report) for item in paths}
        if not scopes:
            scopes = {"other"}
        for scope in scopes:
            scoped[scope][name]["calls"] += 1
            if failed:
                scoped[scope][name]["errors"] += 1
        if name in (*SYNC_CALLS, *RENAME_CALLS, *READ_CALLS, *WRITE_CALLS):
            events.append({
                "line": number,
                "syscall": name,
                "failed": failed,
                "paths": paths,
                "scopes": sorted(scopes),
            })

    def totals(names: Iterable[str], scope: str | None = None) -> dict[str, int]:
        table = counts if scope is None else scoped[scope]
        return {
            "calls": sum(table[name]["calls"] for name in names),
            "errors": sum(table[name]["errors"] for name in names),
        }

    path_order: list[str] = []
    seen_paths: set[str] = set()
    for line in lines:
        for item in _trace_path_tokens(line):
            if item not in seen_paths:
                seen_paths.add(item)
                path_order.append(item)
    paths = sorted(seen_paths)
    output_paths = [item for item in paths if _path_scope(item, tmpdir=tmpdir, replay_dir=replay_dir, report=report) == "output_atomic"]
    replay_paths = [item for item in paths if _path_scope(item, tmpdir=tmpdir, replay_dir=replay_dir, report=report) == "authored_file_store"]
    destination_paths_ordered = [
        item for item in path_order
        if _path_scope(item, tmpdir=tmpdir, replay_dir=replay_dir, report=report)
        == "output_atomic" and item.endswith("/published.docx")
    ]
    destination_paths = sorted(set(destination_paths_ordered))
    sibling_paths = sorted({item for item in output_paths if "/.litchi-" in item and item.endswith(".tmp")})
    # The report records the generated private parent for each destination.
    # Derive the trace-side parent set from the exact destination paths rather
    # than accepting every directory below TMPDIR (which would also include
    # TMPDIR itself and unrelated setup paths).
    parent_paths = sorted({str(Path(item).parent) for item in destination_paths})

    expected_paths = list(expected_atomic_paths or [])
    require(warmups >= 0, f"{path}: warmup count is negative")
    if mode == "atomic":
        require(expected_paths,
                f"{path}: atomic trace requires report destination path bindings")
        if samples is not None:
            require(samples > 0, f"{path}: sample count is not positive")
            require(len(expected_paths) == samples,
                    f"{path}: report destination count differs from samples")
        reported_destinations: list[str] = []
        reported_parents: list[str] = []
        for index, binding in enumerate(expected_paths):
            label = f"{path}: report atomic path binding {index}"
            require(isinstance(binding, Mapping), f"{label} is not an object")
            destination = binding.get("destination")
            private_parent = binding.get("private_parent")
            require(isinstance(destination, str) and Path(destination).is_absolute(),
                    f"{label} destination is not absolute")
            require(isinstance(private_parent, str) and Path(private_parent).is_absolute(),
                    f"{label} private parent is not absolute")
            require(destination.endswith("/published.docx"),
                    f"{label} destination name differs")
            require(private_parent == str(Path(destination).parent),
                    f"{label} private parent is not destination's parent")
            reported_destinations.append(destination)
            reported_parents.append(private_parent)
        require(len(set(reported_destinations)) == len(reported_destinations),
                f"{path}: report repeats an atomic destination path")
        expected_destination_count = warmups + len(reported_destinations)
        require(len(destination_paths_ordered) == expected_destination_count,
                f"{path}: raw trace destination count differs from warmups plus samples")
        warmup_destinations = destination_paths_ordered[:warmups]
        trace_report_destinations = destination_paths_ordered[warmups:]
        require(trace_report_destinations == reported_destinations,
                f"{path}: reported destinations do not match final raw trace sample sequence")
        require(sorted(zip(reported_destinations, reported_parents)) == sorted(
            (destination, str(Path(destination).parent))
            for destination in trace_report_destinations
        ), f"{path}: report destination/parent pairs do not match raw trace paths")
        trace_report_parents = sorted({str(Path(destination).parent)
                                      for destination in trace_report_destinations})
        require(sorted(set(reported_parents)) == trace_report_parents,
                f"{path}: report private parents do not exactly match final raw trace paths")
    else:
        require(not expected_paths,
                f"{path}: counting trace unexpectedly has atomic report path bindings")
        warmup_destinations = []
        trace_report_destinations = []

    def event_count(names: Iterable[str], paths: Iterable[str]) -> int:
        wanted = set(paths)
        return sum(
            1
            for event in events
            if event["syscall"] in names and wanted.intersection(event["paths"])
        )

    replacement_paths = (*sibling_paths, *destination_paths, *parent_paths)

    atomic_lifecycle = {
        "sibling_temporary": {
            "paths": sibling_paths,
            "write_calls": event_count(WRITE_CALLS, sibling_paths),
            "sync_calls": event_count(SYNC_CALLS, sibling_paths),
        },
        "destination": {
            "paths": destination_paths,
            "write_calls": event_count(WRITE_CALLS, destination_paths),
            "sync_calls": event_count(SYNC_CALLS, destination_paths),
        },
        "replacement": {
            "rename_calls": event_count(RENAME_CALLS, replacement_paths),
        },
        "parent_directory": {
            "paths": parent_paths,
            "sync_calls": event_count(SYNC_CALLS, parent_paths),
        },
    }
    if mode == "atomic":
        def successful_event_lines(
            names: Iterable[str], wanted: Iterable[str]
        ) -> list[int]:
            wanted_paths = set(wanted)
            return sorted(
                int(event["line"])
                for event in events
                if not event["failed"]
                and event["syscall"] in names
                and wanted_paths.intersection(event["paths"])
            )

        event_order: list[dict[str, Any]] = []
        for destination in destination_paths_ordered:
            parent = str(Path(destination).parent)
            siblings = [
                item for item in sibling_paths if str(Path(item).parent) == parent
            ]
            require(siblings, f"{path}: no sibling temporary for {destination}")
            # renameat(2) commonly shows only the private parent descriptor
            # with -yy when both names are relative.  Accept that exact parent
            # path as the binding, while also accepting absolute old/new names.
            rename_lines = successful_event_lines(
                RENAME_CALLS, (*siblings, destination, parent)
            )
            parent_sync_lines = successful_event_lines(SYNC_CALLS, (parent,))
            require(rename_lines,
                    f"{path}: no successful replacement rename for {destination}")
            require(parent_sync_lines,
                    f"{path}: no successful parent-directory sync for {destination}")
            first_rename = min(rename_lines)
            last_rename = max(rename_lines)
            sibling_orders: list[dict[str, Any]] = []
            for sibling in siblings:
                sibling_write_lines = successful_event_lines(WRITE_CALLS, (sibling,))
                sibling_sync_lines = successful_event_lines(SYNC_CALLS, (sibling,))
                require(sibling_write_lines,
                        f"{path}: no successful sibling write for {destination} ({sibling})")
                require(sibling_sync_lines,
                        f"{path}: no successful sibling sync for {destination} ({sibling})")
                final_sibling_sync = max(sibling_sync_lines)
                require(max(sibling_write_lines) < final_sibling_sync < first_rename,
                        f"{path}: sibling write/fsync order invalid for {destination} ({sibling})")
                sibling_orders.append({
                    "path": sibling,
                    "write_lines": sibling_write_lines,
                    "sync_lines": sibling_sync_lines,
                    "final_sync_line": final_sibling_sync,
                    "ordered": True,
                })
            sibling_write_lines = successful_event_lines(WRITE_CALLS, siblings)
            sibling_sync_lines = successful_event_lines(SYNC_CALLS, siblings)
            sibling_write_before = [line for line in sibling_write_lines if line < first_rename]
            sibling_sync_before = [line for line in sibling_sync_lines if line < first_rename]
            parent_sync_after = [line for line in parent_sync_lines if line > last_rename]
            require(parent_sync_after,
                    f"{path}: parent-directory fsync must follow rename for {destination}")
            event_order.append({
                "destination": destination,
                "private_parent": parent,
                "sibling_temporary": siblings,
                "sibling_order": sibling_orders,
                "sibling_write_lines": sibling_write_lines,
                "sibling_sync_lines": sibling_sync_lines,
                "sibling_write_before_rename_lines": sibling_write_before,
                "sibling_sync_before_rename_lines": sibling_sync_before,
                "rename_lines": rename_lines,
                "parent_sync_lines": parent_sync_lines,
                "parent_sync_after_rename_lines": parent_sync_after,
                "ordered": True,
            })
        atomic_lifecycle["event_order"] = event_order
    else:
        atomic_lifecycle["event_order"] = []
    summary = {
        "schema": "docx-tail-append-strace-v1",
        "mode": mode,
        "warmups": warmups,
        "reported_samples": len(trace_report_destinations),
        "selected_syscalls": list(SYSCALLS),
        "line_count": len(lines),
        "raw_bytes": len(raw),
        "raw_sha256": hashlib.sha256(raw).hexdigest(),
        "syscalls": counts,
        "scoped_syscalls": scoped,
        "read_calls": totals(READ_CALLS),
        "write_calls": totals(WRITE_CALLS),
        "sync_calls": totals(SYNC_CALLS),
        "rename_calls": totals(RENAME_CALLS),
        "paths": {
            "report": _path_text(report),
            "replay_root": _path_text(replay_dir),
            "tmp_root": _path_text(tmpdir),
            "all": paths,
            "authored_file_store": replay_paths,
            "output_atomic": output_paths,
            "destination": destination_paths,
            "destination_chronological": destination_paths_ordered,
            "warmup_trace_only_destinations": warmup_destinations,
            "reported_destinations": trace_report_destinations,
            "sibling_temporary": sibling_paths,
            "output_parent_directories": parent_paths,
        },
        "atomic_lifecycle": atomic_lifecycle,
        "events": events,
        "phases": {
            "authored_file_store": {
                "write": totals(WRITE_CALLS, "authored_file_store"),
                "read": totals(READ_CALLS, "authored_file_store"),
                "sync": totals(SYNC_CALLS, "authored_file_store"),
                "rename": totals(RENAME_CALLS, "authored_file_store"),
            },
            "output_atomic": {
                "write": totals(WRITE_CALLS, "output_atomic"),
                "read": totals(READ_CALLS, "output_atomic"),
                "sync": totals(SYNC_CALLS, "output_atomic"),
                "rename": totals(RENAME_CALLS, "output_atomic"),
            },
            "report": {
                "write": totals(WRITE_CALLS, "report"),
                "read": totals(READ_CALLS, "report"),
                "sync": totals(SYNC_CALLS, "report"),
                "rename": totals(RENAME_CALLS, "report"),
            },
        },
    }
    require(any(row["calls"] for row in counts.values()), f"{path}: no selected syscall rows")
    require(replay_paths, f"{path}: authored replay path was not observed")
    require(summary["phases"]["authored_file_store"]["sync"]["calls"] > 0,
            f"{path}: --replay-sync data produced no authored file-store sync")
    if mode == "atomic":
        require(destination_paths, f"{path}: atomic destination path was not observed")
        require(sibling_paths, f"{path}: sibling temporary path was not observed")
        require(summary["phases"]["output_atomic"]["write"]["calls"] > 0,
                f"{path}: atomic output produced no sibling writes")
        lifecycle = summary["atomic_lifecycle"]
        require(lifecycle["sibling_temporary"]["write_calls"] > 0,
                f"{path}: sibling temporary write was not observed")
        require(lifecycle["sibling_temporary"]["sync_calls"] > 0,
                f"{path}: sibling temporary fsync was not observed")
        require(lifecycle["replacement"]["rename_calls"] > 0,
                f"{path}: destination replacement was not observed")
        require(lifecycle["parent_directory"]["sync_calls"] > 0,
                f"{path}: parent-directory sync was not observed")
    else:
        require(not destination_paths and not sibling_paths,
                f"{path}: counting route unexpectedly left atomic destination paths")
    return summary


def _oracle_flags(value: Any, label: str) -> dict[str, bool]:
    require(isinstance(value, Mapping), f"{label}: oracle is not an object")
    result: dict[str, bool] = {}
    for name in ORACLE_FLAGS:
        require(value.get(name) is True, f"{label}.{name}: oracle did not pass")
        result[name] = True
    return result


def _sha_field(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value), f"{label}: invalid SHA-256")
    return value


def _validate_report(path: Path, mode: str, *, samples: int, warmups: int) -> dict[str, Any]:
    report = _read_json(path)
    require(isinstance(report, Mapping), f"{path}: report is not an object")
    require(report.get("schema") == REPORT_SCHEMA and report.get("version") == 1,
            f"{path}: report schema/version differs")
    config = report.get("config")
    require(isinstance(config, Mapping), f"{path}: config is missing")
    require(config.get("samples") == samples and config.get("warmups") == warmups,
            f"{path}: samples/warmups differ from profile contract")
    require(config.get("provider") == "file_store",
            f"{path}: profile must use the file_store authored provider")
    require(config.get("replay_sync") == "data",
            f"{path}: profile must use data-synchronized authored replay")
    expected_route = MODE_ROUTE[mode]
    expected_sink = (
        "non_seek_counting_short_write_no_archive_retention_production_artifact_proof"
        if mode == "counting"
        else "production_atomic_path_output_proof_post_timer_destination_oracle"
    )
    require(config.get("sink") == expected_sink, f"{path}: publication sink contract differs")
    require(config.get("fixture_dir") is None, f"{path}: fixture export is not a profile arm")
    publication_config = config.get("publication")
    if isinstance(publication_config, str):
        require(publication_config == expected_route, f"{path}: publication route differs")
    elif isinstance(publication_config, Mapping):
        require(publication_config.get("route") == expected_route,
                f"{path}: publication route differs")
    else:
        fail(f"{path}: publication route is missing")
    cases = report.get("cases")
    require(isinstance(cases, list) and cases, f"{path}: report cases are missing")
    case_summaries: list[dict[str, Any]] = []
    atomic_paths: list[dict[str, str]] = []
    for index, case in enumerate(cases):
        label = f"{path}.cases[{index}]"
        require(isinstance(case, Mapping), f"{label}: case is not an object")
        source = case.get("source")
        require(isinstance(source, Mapping) and source.get("unchanged_oracle") is True,
                f"{label}: source oracle did not pass")
        require(case.get("provider") == "file_store",
                f"{label}: case provider differs from file_store profile arm")
        _sha_field(source.get("archive_sha256"), f"{label}.source.archive_sha256")
        _oracle_flags(case.get("oracle"), f"{label}.oracle")
        proof = case.get("proof")
        require(isinstance(proof, Mapping), f"{label}: proof is missing")
        for field in ("source_sha256", "candidate_sha256"):
            _sha_field(proof.get(field), f"{label}.proof.{field}")
        rows = case.get("samples")
        require(isinstance(rows, list) and len(rows) == samples,
                f"{label}: sample count differs")
        for sample_index, sample in enumerate(rows):
            sample_label = f"{label}.samples[{sample_index}]"
            require(isinstance(sample, Mapping), f"{sample_label}: sample is not an object")
            elapsed = sample.get("elapsed_ns")
            require(isinstance(elapsed, int) and elapsed >= 0, f"{sample_label}: elapsed is invalid")
            replay = sample.get("replay")
            require(isinstance(replay, Mapping) and replay.get("route") == "file_store",
                    f"{sample_label}: file replay observation is missing")
            require(isinstance(replay.get("file_sync_calls"), int)
                    and replay.get("file_sync_calls") > 0,
                    f"{sample_label}: authored file-store sync observation is missing")
            require(isinstance(replay.get("file_write_calls"), int)
                    and replay.get("file_write_calls") > 0
                    and replay.get("file_cleanup_verified") is True,
                    f"{sample_label}: authored file-store cleanup observation failed")
            publication = sample.get("publication")
            require(isinstance(publication, Mapping), f"{sample_label}: publication proof is missing")
            publication_fields = {
                "schema", "route", "timing_scope", "timed_candidate_artifact_bytes",
                "timed_candidate_artifact_sha256", "timed_candidate_matches_oracle",
                "verification_scope",
            }
            actual_publication_fields = set(publication)
            if mode == "atomic":
                publication_fields.add("atomic")
            require(actual_publication_fields == publication_fields
                    or (mode == "counting" and actual_publication_fields == publication_fields | {"atomic"}),
                    f"{sample_label}: publication proof fields differ")
            require(publication.get("schema") == PUBLICATION_SCHEMA,
                    f"{sample_label}: publication schema differs")
            require(publication.get("route") == expected_route,
                    f"{sample_label}: publication route differs")
            expected_timing_scope = COUNTING_TIMING_SCOPE if mode == "counting" else ATOMIC_TIMING_SCOPE
            expected_verification_scope = (
                COUNTING_VERIFICATION_SCOPE
                if mode == "counting"
                else ATOMIC_VERIFICATION_SCOPE
            )
            require(publication.get("timing_scope") == expected_timing_scope,
                    f"{sample_label}: publication timing scope differs")
            require(publication.get("verification_scope") == expected_verification_scope,
                    f"{sample_label}: publication verification scope differs")
            require(publication.get("timed_candidate_matches_oracle") is True,
                    f"{sample_label}: timed artifact oracle did not pass")
            candidate_bytes = case.get("oracle", {}).get("candidate_archive_bytes")
            require(isinstance(candidate_bytes, int) and candidate_bytes > 0,
                    f"{sample_label}: candidate archive length is missing")
            require(publication.get("timed_candidate_artifact_bytes") == candidate_bytes,
                    f"{sample_label}: timed artifact length differs")
            _sha_field(publication.get("timed_candidate_artifact_sha256"),
                       f"{sample_label}.publication.timed_candidate_artifact_sha256")
            candidate_sha = _sha_field(
                case.get("oracle", {}).get("candidate_archive_sha256"),
                f"{label}.oracle.candidate_archive_sha256",
            )
            require(publication.get("timed_candidate_artifact_sha256") == candidate_sha,
                    f"{sample_label}: timed artifact digest differs")
            sink = sample.get("sink")
            require(isinstance(sink, Mapping), f"{sample_label}: sink observation is missing")
            if mode == "counting":
                require(sink.get("accepted_bytes") == candidate_bytes
                        and isinstance(sink.get("write_calls"), int)
                        and isinstance(sink.get("largest_write"), int)
                        and isinstance(sink.get("histogram"), Mapping)
                        and sink.get("sha256") is None,
                        f"{sample_label}: counting sink observation is incomplete")
                require(publication.get("atomic") is None,
                        f"{sample_label}: counting route has an atomic record")
            else:
                require(all(sink.get(field) is None for field in (
                    "accepted_bytes", "write_calls", "largest_write", "histogram", "sha256"
                )), f"{sample_label}: atomic route invented sink counters")
            if mode == "atomic":
                atomic = publication.get("atomic")
                require(isinstance(atomic, Mapping), f"{sample_label}: atomic proof is missing")
                require(set(atomic) == {
                    "destination_path", "private_parent_path", "before", "after",
                    "post_timer_archive_bytes", "post_timer_archive_sha256",
                    "output_bytes_exact", "output_sha256_exact", "inverse_oracle_scope",
                    "post_timer_oracle", "cleanup",
                }, f"{sample_label}: atomic proof fields differ")
                for field in ("destination_path", "private_parent_path"):
                    recorded_path = atomic.get(field)
                    require(isinstance(recorded_path, str) and Path(recorded_path).is_absolute(),
                            f"{sample_label}.{field}: path is not absolute")
                    require(not Path(recorded_path).exists(),
                            f"{sample_label}.{field}: cleaned path still exists")
                destination_path = str(atomic["destination_path"])
                private_parent_path = str(atomic["private_parent_path"])
                require(destination_path.endswith("/published.docx"),
                        f"{sample_label}: atomic destination name differs")
                require(private_parent_path == str(Path(destination_path).parent),
                        f"{sample_label}: atomic private parent is not destination's parent")
                atomic_paths.append({
                    "destination": destination_path,
                    "private_parent": private_parent_path,
                })
                require(atomic.get("before") == {
                    "exists": False, "regular_file": False, "bytes": None
                }, f"{sample_label}: atomic destination was not absent before publish")
                after = atomic.get("after")
                require(isinstance(after, Mapping)
                        and after.get("exists") is True
                        and after.get("regular_file") is True
                        and after.get("bytes") == candidate_bytes,
                        f"{sample_label}: atomic post-timer destination state differs")
                require(atomic.get("output_bytes_exact") is True and
                        atomic.get("output_sha256_exact") is True,
                        f"{sample_label}: final artifact oracle did not pass")
                require(atomic.get("post_timer_archive_bytes") == candidate_bytes,
                        f"{sample_label}: final artifact length differs")
                _sha_field(atomic.get("post_timer_archive_sha256"),
                           f"{sample_label}.publication.atomic.post_timer_archive_sha256")
                require(atomic.get("post_timer_archive_sha256") == candidate_sha,
                        f"{sample_label}: final artifact digest differs")
                require(atomic.get("inverse_oracle_scope") == INVERSE_ORACLE_SCOPE,
                        f"{sample_label}: inverse oracle scope differs")
                _oracle_flags(atomic.get("post_timer_oracle"),
                              f"{sample_label}.publication.atomic.post_timer_oracle")
                cleanup = atomic.get("cleanup")
                require(isinstance(cleanup, Mapping) and set(cleanup) == {
                    "destination_removed", "parent_removed"
                } and
                        cleanup.get("destination_removed") is True and
                        cleanup.get("parent_removed") is True,
                        f"{sample_label}: atomic destination cleanup did not pass")
        case_summaries.append({
            "source_count": case.get("source_count"),
            "authored_count": case.get("authored_count"),
            "samples": len(rows),
            "candidate_archive_sha256": case["oracle"].get("candidate_archive_sha256"),
            "oracle_flags": {name: True for name in ORACLE_FLAGS},
        })
    return {
        "schema": "docx-tail-append-profile-report-summary-v1",
        "cases": len(cases),
        "samples": sum(len(case.get("samples", [])) for case in cases),
        "mode": mode,
        "route": expected_route,
        "case_summaries": case_summaries,
        "atomic_paths": atomic_paths,
    }


def _remove_empty_tree(path: Path, remaining: list[str], removed: list[str]) -> None:
    if not path.exists():
        return
    require(path.is_dir() and not path.is_symlink(), f"private path is not a directory: {path}")
    for child in sorted(path.iterdir()):
        if child.is_dir() and not child.is_symlink():
            _remove_empty_tree(child, remaining, removed)
        else:
            remaining.append(str(child))
    if not any(path.iterdir()):
        path.rmdir()
        removed.append(str(path))


def _cleanup_private(run_root: Path, tmpdir: Path, replay_dir: Path) -> dict[str, Any]:
    require(tmpdir.parent == run_root and replay_dir.parent == run_root,
            "private profile directories escaped run root")
    remaining: list[str] = []
    removed: list[str] = []
    for path in (tmpdir, replay_dir):
        _remove_empty_tree(path, remaining, removed)
    if run_root.exists():
        require(run_root.is_dir() and not run_root.is_symlink(),
                f"private run root is not a directory: {run_root}")
        for child in sorted(run_root.iterdir()):
            remaining.append(str(child))
        if not any(run_root.iterdir()):
            run_root.rmdir()
            removed.append(str(run_root))
    return {
        "schema": CLEANUP_SCHEMA,
        "status": "pass" if not remaining else "failed",
        "run_root": str(run_root),
        "tmpdir": str(tmpdir),
        "replay_dir": str(replay_dir),
        "removed": removed,
        "remaining": remaining,
    }


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


def _run_one(
    *,
    attempt: str,
    mode: str,
    binary: Path,
    build: Mapping[str, Any],
    build_path: Path,
    samples: int,
    warmups: int,
    timeout_seconds: int,
    source_count: int,
    authored_count: int,
    chunk: str,
    text: str,
    provider: str,
    replay_sync: str,
    replay_max_bytes: int,
    publication_flag: str,
    publication_path_flag: str | None,
) -> dict[str, Any]:
    label = f"strace-{mode}"
    destination = ROOT / "profiles" / attempt / label
    require(not destination.exists(), f"refusing to replace profile directory: {destination}")
    destination.mkdir(parents=True, exist_ok=False)
    run_root = TEMP / "publication-profiles" / attempt / label
    require(not run_root.exists(), f"refusing to replace private run root: {run_root}")
    run_root.mkdir(parents=True, exist_ok=False)
    tmpdir = run_root / "tmp"
    replay_dir = run_root / "replay"
    tmpdir.mkdir()
    replay_dir.mkdir()
    report = destination / "report.json"
    trace = destination / "strace.raw"
    stdout = destination / "stdout.txt"
    stderr = destination / "stderr.txt"
    cleanup_path = destination / "replay-cleanup.json"
    terminal_path = destination / "terminal.json"
    started_path = destination / "started.json"
    publication_destination = (
        tmpdir / "published.docx"
        if publication_path_flag is not None
        else tmpdir / "litchi-docx-replay-atomic-<pid>-<serial>" / "published.docx"
    )
    benchmark = _benchmark_argv(
        binary=binary,
        report=report,
        replay_dir=replay_dir,
        mode=mode,
        samples=samples,
        warmups=warmups,
        source_count=source_count,
        authored_count=authored_count,
        chunk=chunk,
        text=text,
        provider=provider,
        replay_sync=replay_sync,
        replay_max_bytes=replay_max_bytes,
        publication_flag=publication_flag,
        publication_path_flag=publication_path_flag,
        destination=publication_destination,
    )
    profiler = _tool_executable("strace")
    profiler_argv = [
        profiler["path"],
        "-f",
        "-yy",
        "-ttt",
        "-T",
        "-s",
        "512",
        "-e",
        "trace=" + ",".join(SYSCALLS),
        "-o",
        str(trace),
        "--",
        *benchmark,
    ]
    source_revision_before = _current_revision()
    run_environment = dict(os.environ)
    run_environment.update({
        "TMPDIR": str(tmpdir),
        "LC_ALL": "C",
        "PYTHONDONTWRITEBYTECODE": "1",
    })
    started = {
        "schema": SCHEMA,
        "version": VERSION,
        "status": "running",
        "attempt": attempt,
        "label": label,
        "mode": mode,
        "tool": "strace",
        "tool_executable": profiler,
        "samples": samples,
        "warmups": warmups,
        "scope": SCOPE,
        "operation_attribution_claim": False,
        "performance_claim": NO_CLAIM,
        "caller_gate": dict(CALLER_GATE),
        "cpu": CPU,
        "build": {"path": str(build_path), "bytes": build_path.stat().st_size,
                   "sha256": _sha(build_path), "binding": dict(build)},
        "binary": _artifact(binary, executable=True),
        "source_revision": build["git_revision"],
        "source_revision_before": source_revision_before,
        "helper": _artifact(Path(__file__)),
        "run": {
            "source_count": source_count,
            "authored_count": authored_count,
            "chunk": chunk,
            "text": text,
            "provider": provider,
            "replay_sync": replay_sync,
            "replay_max_bytes": replay_max_bytes,
            "publication_flag": publication_flag,
            "publication_path_flag": publication_path_flag,
        },
        "benchmark_argv": benchmark,
        "argv": profiler_argv,
        "cwd": str(REPO),
        "environment": {"TMPDIR": str(tmpdir), "LC_ALL": "C", "PYTHONDONTWRITEBYTECODE": "1"},
        "started_utc": _now(),
        "timeout_seconds": timeout_seconds,
        "tmpdir": str(tmpdir),
        "replay_dir": str(replay_dir),
        "destination_contract": {
            "destination_descriptor": str(publication_destination),
            "sibling_temporary_suffix": ".tmp",
            "sibling_temporary_prefix": ".litchi-",
            "directory_sync_parent": "the generated atomic destination parent directory",
            "exact_paths_source": "strace -yy raw trace and parsed paths.paths",
        },
    }
    _write_exclusive(started_path, started)
    stdout.touch(mode=0o664, exist_ok=False)
    stderr.touch(mode=0o664, exist_ok=False)
    process: subprocess.Popen[bytes] | None = None
    exit_code: int | None = None
    timed_out = False
    termination: str | None = None
    launch_error: str | None = None
    validation_error: str | None = None
    report_summary: dict[str, Any] | None = None
    trace_summary: dict[str, Any] | None = None
    with stdout.open("wb") as out, stderr.open("wb") as err:
        try:
            process = subprocess.Popen(
                profiler_argv,
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
            report_summary = _validate_report(report, mode, samples=samples, warmups=warmups)
            trace_summary = _parse_trace(
                trace, mode=mode, tmpdir=tmpdir, replay_dir=replay_dir, report=report,
                expected_atomic_paths=report_summary["atomic_paths"],
                warmups=warmups,
                samples=samples,
            )
        except (ProfileError, OSError, ValueError, KeyError) as error:
            validation_error = str(error)
    try:
        cleanup = _cleanup_private(run_root, tmpdir, replay_dir)
    except (ProfileError, OSError) as error:
        cleanup = {
            "schema": CLEANUP_SCHEMA,
            "status": "failed",
            "run_root": str(run_root),
            "tmpdir": str(tmpdir),
            "replay_dir": str(replay_dir),
            "removed": [],
            "remaining": [str(error)],
        }
    _write_exclusive(cleanup_path, cleanup)
    source_revision_after = None
    try:
        source_revision_after = _current_revision()
        if source_revision_after != source_revision_before:
            validation_error = validation_error or "source revision changed during profile child"
    except ProfileError as error:
        validation_error = validation_error or str(error)
    artifacts = {
        path.name: _artifact(path)
        for path in (stdout, stderr, report, trace, cleanup_path)
        if path.is_file() and not path.is_symlink()
    }
    expected_artifacts = {stdout.name, stderr.name, report.name, trace.name, cleanup_path.name}
    missing_artifacts = sorted(expected_artifacts - set(artifacts))
    passed = (
        exit_code == 0 and not timed_out and termination is None and launch_error is None
        and validation_error is None and not missing_artifacts and cleanup.get("status") == "pass"
    )
    terminal = dict(started)
    terminal.update({
        "schema": TERMINAL_SCHEMA,
        "status": "pass" if passed else "failed",
        "exit_code": exit_code,
        "timed_out": timed_out,
        "termination": termination,
        "finished_utc": _now(),
        "source_revision_after": source_revision_after,
        "artifacts": artifacts,
        "missing_artifacts": missing_artifacts,
        "cleanup": cleanup,
        "report_summary": report_summary,
        "profiler": trace_summary,
        "started_artifact": _artifact(started_path),
    })
    if launch_error is not None:
        terminal["launch_error"] = launch_error
    if validation_error is not None:
        terminal["validation_error"] = validation_error
    _write_exclusive(terminal_path, terminal)
    record = {
        "label": label,
        "mode": mode,
        "status": terminal["status"],
        "terminal": _artifact(terminal_path),
        "report": artifacts.get(report.name),
        "trace": artifacts.get(trace.name),
        "scope": SCOPE,
        "operation_attribution_claim": False,
    }
    if not passed:
        record["error"] = validation_error or launch_error or "profile child failed"
    return record


def _load_result(path: Path) -> Mapping[str, Any]:
    value = _read_json(path)
    require(isinstance(value, Mapping), f"{path}: result is not an object")
    require(value.get("schema") == RESULT_SCHEMA and value.get("version") == VERSION,
            f"{path}: result schema/version differs")
    return value


def _capture(args: argparse.Namespace) -> int:
    attempt = _safe_token(args.attempt)
    require(args.samples > 0 and args.warmups > 0, "samples and warmups must be positive")
    require(args.timeout_seconds > 0, "timeout must be positive")
    require(args.source_count > 0 and args.authored_count > 0, "counts must be positive")
    require(args.replay_max_bytes > 0, "replay maximum must be positive")
    modes = tuple(args.modes.split(","))
    require(modes and all(mode in MODES for mode in modes), "modes must be counting and/or atomic")
    require(len(set(modes)) == len(modes), "modes must not repeat")
    attempt_root = ROOT / "profiles" / attempt
    require(not attempt_root.exists(), f"refusing to replace profile attempt: {attempt_root}")
    build_path = Path(args.build).resolve() if args.build else None
    if build_path is None:
        for candidate in (ROOT / "build-after.json", ROOT / "build-after-after1.json"):
            if candidate.is_file():
                build_path = candidate
                break
    require(build_path is not None, "--build is required when no default after receipt exists")
    build, binary_binding = _build_binding(build_path, args.role)
    binary = Path(args.binary).resolve() if args.binary else Path(binary_binding["path"])
    actual_binary = _artifact(binary, executable=True)
    require(actual_binary == binary_binding,
            "--binary does not match the retained binary bound by the build receipt")
    records: list[dict[str, Any]] = []
    for mode in modes:
        try:
            records.append(_run_one(
                attempt=attempt,
                mode=mode,
                binary=binary,
                build=build,
                build_path=build_path,
                samples=args.samples,
                warmups=args.warmups,
                timeout_seconds=args.timeout_seconds,
                source_count=args.source_count,
                authored_count=args.authored_count,
                chunk=args.chunk,
                text=args.text,
                provider=args.provider,
                replay_sync=args.replay_sync,
                replay_max_bytes=args.replay_max_bytes,
                publication_flag=args.publication_flag,
                publication_path_flag=args.publication_path_flag,
            ))
        except (ProfileError, OSError, ValueError, KeyError) as error:
            records.append({
                "label": f"strace-{mode}",
                "mode": mode,
                "status": "failed",
                "error": str(error),
                "scope": SCOPE,
                "operation_attribution_claim": False,
            })
    result = {
        "schema": RESULT_SCHEMA,
        "version": VERSION,
        "status": "pass" if all(row.get("status") == "pass" for row in records) else "incomplete",
        "attempt": attempt,
        "modes": list(modes),
        "samples": args.samples,
        "warmups": args.warmups,
        "scope": SCOPE,
        "operation_attribution_claim": False,
        "performance_claim": NO_CLAIM,
        "caller_gate": dict(CALLER_GATE),
        "cpu": CPU,
        "build": {"path": str(build_path), "bytes": build_path.stat().st_size,
                   "sha256": _sha(build_path), "binding": build},
        "binary": actual_binary,
        "source_revision": build["git_revision"],
        "helper": _artifact(Path(__file__)),
        "records": records,
        "finished_utc": _now(),
    }
    _write_exclusive(attempt_root / "result.json", result)
    print(attempt_root / "result.json")
    return 0 if result["status"] == "pass" else 1


def _verify(attempt: str) -> int:
    attempt = _safe_token(attempt)
    attempt_root = ROOT / "profiles" / attempt
    result = _load_result(attempt_root / "result.json")
    require(result.get("status") == "pass" and result.get("attempt") == attempt,
            "profile result is incomplete or bound to another attempt")
    require(result.get("scope") == SCOPE and result.get("operation_attribution_claim") is False
            and result.get("performance_claim") == NO_CLAIM,
            "profile result scope or claim differs")
    require(result.get("caller_gate") == CALLER_GATE and result.get("cpu") == CPU,
            "profile caller gate or CPU binding differs")
    helper = _artifact(Path(__file__))
    require(result.get("helper") == helper, "profile helper changed")
    records = result.get("records")
    modes = result.get("modes")
    require(isinstance(records, list) and isinstance(modes, list) and
            len(records) == len(modes), "profile record inventory differs")
    previous_end: _datetime.datetime | None = None
    for record, mode in zip(records, modes, strict=True):
        require(isinstance(record, Mapping) and record.get("status") == "pass",
                f"{mode}: retained record is not passing")
        label = f"strace-{mode}"
        require(record.get("label") == label and record.get("mode") == mode,
                f"{label}: record identity differs")
        directory = attempt_root / label
        require(directory.is_dir() and not directory.is_symlink(), f"{label}: directory missing")
        expected_names = {
            "started.json", "terminal.json", "stdout.txt", "stderr.txt", "report.json",
            "strace.raw", "replay-cleanup.json",
        }
        require({child.name for child in directory.iterdir()} == expected_names,
                f"{label}: artifact inventory differs")
        started = _read_json(directory / "started.json")
        terminal = _read_json(directory / "terminal.json")
        require(isinstance(started, Mapping) and started.get("schema") == SCHEMA,
                f"{label}: started schema differs")
        require(isinstance(terminal, Mapping) and terminal.get("schema") == TERMINAL_SCHEMA,
                f"{label}: terminal schema differs")
        require(started.get("attempt") == attempt and started.get("label") == label
                and started.get("mode") == mode,
                f"{label}: started identity differs")
        require(started.get("binary") == result.get("binary")
                and started.get("source_revision") == result.get("source_revision")
                and started.get("build") == result.get("build")
                and started.get("helper") == result.get("helper"),
                f"{label}: build/source/helper binding differs")
        started_binary = started.get("binary")
        require(isinstance(started_binary, Mapping), f"{label}: binary binding is malformed")
        binary_path = Path(str(started_binary.get("path")))
        require(_artifact(binary_path, executable=True) == dict(started_binary),
                f"{label}: retained binary changed")
        build_binding = started.get("build")
        require(isinstance(build_binding, Mapping), f"{label}: build binding is malformed")
        build_path = Path(str(build_binding.get("path")))
        require(_meta(build_path) == {
            "path": str(build_path),
            "bytes": build_binding.get("bytes"),
            "sha256": build_binding.get("sha256"),
        }, f"{label}: build receipt changed")
        tool_executable = _tool_executable("strace")
        require(started.get("tool_executable") == tool_executable,
                f"{label}: strace executable changed")
        benchmark_argv = started.get("benchmark_argv")
        require(isinstance(benchmark_argv, list) and benchmark_argv[0:3] ==
                ["/usr/bin/taskset", "-c", str(CPU)],
                f"{label}: benchmark CPU binding differs")
        expected_profiler_argv = [
            tool_executable["path"], "-f", "-yy", "-ttt", "-T", "-s", "512", "-e",
            "trace=" + ",".join(SYSCALLS), "-o", str(directory / "strace.raw"),
            "--", *benchmark_argv,
        ]
        require(started.get("argv") == expected_profiler_argv,
                f"{label}: profiler argv binding differs")
        for key in ("tool", "samples", "warmups", "scope", "operation_attribution_claim",
                    "performance_claim", "caller_gate", "cpu"):
            require(terminal.get(key) == started.get(key),
                    f"{label}: terminal {key} binding differs")
        require(started.get("status") == "running" and terminal.get("status") == "pass",
                f"{label}: terminal status differs")
        require(terminal.get("exit_code") == 0 and terminal.get("timed_out") is False
                and terminal.get("termination") is None and terminal.get("missing_artifacts") == [],
                f"{label}: child did not exit cleanly")
        cleanup = terminal.get("cleanup")
        require(isinstance(cleanup, Mapping) and cleanup.get("schema") == CLEANUP_SCHEMA
                and cleanup.get("status") == "pass" and cleanup.get("remaining") == [],
                f"{label}: private cleanup did not pass")
        require(cleanup == _read_json(directory / "replay-cleanup.json"),
                f"{label}: cleanup receipt changed")
        tmpdir = Path(str(started["tmpdir"]))
        replay_dir = Path(str(started["replay_dir"]))
        run_root = tmpdir.parent
        require(not run_root.exists() and not tmpdir.exists() and not replay_dir.exists(),
                f"{label}: private scratch remains on disk")
        report_summary = _validate_report(
            directory / "report.json", mode,
            samples=int(started["samples"]), warmups=int(started["warmups"]),
        )
        trace_summary = _parse_trace(
            directory / "strace.raw", mode=mode, tmpdir=tmpdir,
            replay_dir=replay_dir, report=directory / "report.json",
            expected_atomic_paths=report_summary["atomic_paths"],
            warmups=int(started["warmups"]),
            samples=int(started["samples"]),
        )
        require(terminal.get("report_summary") == report_summary,
                f"{label}: report summary changed")
        require(terminal.get("profiler") == trace_summary,
                f"{label}: parsed trace summary changed")
        require(terminal.get("artifacts") == {
            name: _artifact(directory / name)
            for name in ("stdout.txt", "stderr.txt", "report.json", "strace.raw", "replay-cleanup.json")
        }, f"{label}: terminal artifact inventory changed")
        require(terminal.get("started_artifact") == _artifact(directory / "started.json"),
                f"{label}: started artifact binding changed")
        require(record.get("terminal") == _artifact(directory / "terminal.json"),
                f"{label}: terminal binding changed")
        require(record.get("report") == _artifact(directory / "report.json") and
                record.get("trace") == _artifact(directory / "strace.raw"),
                f"{label}: result artifact binding changed")
        start = _timestamp(started.get("started_utc"), f"{label}.started_utc")
        end = _timestamp(terminal.get("finished_utc"), f"{label}.finished_utc")
        require(end > start and (previous_end is None or start >= previous_end),
                f"{label}: chronology is invalid")
        previous_end = end
    result_end = _timestamp(result.get("finished_utc"), "profile result")
    require(previous_end is not None and result_end >= previous_end,
            "profile result finished before terminal capture")
    print(f"verified {attempt}")
    return 0


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    capture = subparsers.add_parser("capture")
    capture.add_argument("attempt")
    capture.add_argument("--build")
    capture.add_argument("--binary")
    capture.add_argument("--role", choices=("normal", "allocator"), default="normal")
    capture.add_argument("--modes", default=",".join(MODES))
    capture.add_argument("--samples", type=int, default=DEFAULT_SAMPLES)
    capture.add_argument("--warmups", type=int, default=DEFAULT_WARMUPS)
    capture.add_argument("--timeout-seconds", type=int, default=DEFAULT_TIMEOUT_SECONDS)
    capture.add_argument("--source-count", type=int, default=DEFAULT_SOURCE_COUNT)
    capture.add_argument("--authored-count", type=int, default=DEFAULT_AUTHORED_COUNT)
    capture.add_argument("--chunk", default=DEFAULT_CHUNK)
    capture.add_argument("--text", default=DEFAULT_TEXT)
    capture.add_argument("--provider", default=DEFAULT_PROVIDER)
    capture.add_argument("--replay-sync", default=DEFAULT_REPLAY_SYNC)
    capture.add_argument("--replay-max-bytes", type=int, default=DEFAULT_REPLAY_MAX_BYTES)
    capture.add_argument("--publication-flag", default="--publication")
    capture.add_argument("--publication-path-flag")
    verify = subparsers.add_parser("verify")
    verify.add_argument("attempt")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "verify":
            return _verify(args.attempt)
        return _capture(args)
    except (ProfileError, OSError, ValueError, KeyError, subprocess.SubprocessError) as error:
        print(f"profile.py: FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
