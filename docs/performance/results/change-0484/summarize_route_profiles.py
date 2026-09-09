#!/usr/bin/env python3
"""Validate and interpret the external profiles for one frozen route attempt.

The profile driver deliberately records whole-child observations.  This helper
keeps that scope, validates the identity chain before interpreting an artifact,
and leaves unavailable counters unavailable.  It does not launch a profiler or
make a timing, throughput, speedup, or allocation claim.

The normal invocation is::

    python3 summarize_route_profiles.py \
      --profiles-dir route-profiles/profiles1

The output is a JSON diagnostic summary beside the profile directory.  A
``perf-record`` profile also gets an explicit ``perf script`` command in the
summary and in a command-manifest file.  The command is intentionally emitted
for the coordinator to run separately; this program never invokes perf,
strace, or heaptrack.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any


DOCROOT = Path(__file__).resolve().parent
REPO = DOCROOT.parents[3]
PROFILE_RUN_SCHEMA = "docx-replayable-tail-append-route-profile-run-v1"
PROFILE_SCHEMA = "docx-replayable-tail-append-route-profile-v1"
SUMMARY_SCHEMA = "docx-replayable-tail-append-route-profile-summary-v1"
PERF_SCRIPT_SCHEMA = "docx-replayable-tail-append-perf-script-v1"

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
CACHE_EVENTS = {
    "cache-references",
    "cache-misses",
    "L1-dcache-loads",
    "L1-dcache-load-misses",
    "LLC-loads",
    "LLC-load-misses",
}
PAYLOAD_SYSCALLS = {
    "read",
    "pread64",
    "readv",
    "write",
    "pwrite64",
    "writev",
}
TRACE_SYSCALLS = (
    "read",
    "pread64",
    "readv",
    "write",
    "pwrite64",
    "writev",
    "fsync",
    "fdatasync",
    "unlink",
    "unlinkat",
)
FD_SYSCALLS = ("openat", "close")
SIZE_BUCKETS = (
    "bytes_0",
    "bytes_1_to_512",
    "bytes_513_to_4096",
    "bytes_4097_to_16384",
    "bytes_16385_to_65536",
    "bytes_over_65536",
)
PROFILE_TOOLS = ("perf-stat", "perf-record", "strace", "heaptrack")

# These are deliberately narrow names.  A broad ``replay`` pattern would
# match the benchmark module name in every frame and would fabricate ancestry.
STACK_MARKERS = {
    "lifecycle": (
        # Demangled Rust names use ``::run_case``; a raw symbol can carry a
        # length prefix such as ``8run_case``.  Matching the function suffix
        # keeps both forms observable without matching the module name.
        r"run_case",
        r"run_iteration",
        r"run_from_args",
    ),
    "auditor": (
        r"audit_splice",
        r"verify_source",
        r"verify_candidate",
        r"verify_store_preflight",
        r"verify_reader",
        r"scan_document",
        r"scan_main_part",
        r"semantic_record",
        r"source_main_xml",
        r"build_fixture",
    ),
    "replay": (
        r"read_with_accounting",
        r"write_replacing",
        r"stream_(?:part|splice)",
        r"run_replay",
        r"replay_(?:read|open|finish|append)",
        r"Replay(?:Reader|Store|Route|Source)",
    ),
}


class Issues:
    """Collect validation errors without hiding the rest of the artifacts."""

    def __init__(self) -> None:
        self.errors: list[str] = []
        self.warnings: list[str] = []

    def error(self, message: str) -> None:
        self.errors.append(message)

    def warning(self, message: str) -> None:
        self.warnings.append(message)


def _json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as stream:
        return json.load(stream)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _metadata(path: Path, cache: dict[str, dict[str, Any]]) -> dict[str, Any]:
    resolved = path.resolve()
    key = str(resolved)
    if key not in cache:
        cache[key] = {
            "bytes": resolved.stat().st_size,
            "sha256": _sha256(resolved),
        }
    return dict(cache[key])


def _record_path(value: Any, *, base: Path = DOCROOT) -> Path:
    path = Path(str(value))
    return path if path.is_absolute() else base / path


def _display(path: Path) -> str:
    resolved = path.resolve()
    for base in (DOCROOT.resolve(), REPO.resolve()):
        try:
            return resolved.relative_to(base).as_posix()
        except ValueError:
            pass
    return str(resolved)


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")


def _integer(value: str) -> int | None:
    text = value.strip().replace(" ", "").replace(",", "")
    if not text or text.startswith("<"):
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return int(number) if number.is_integer() else None


def _number(value: str) -> int | float | None:
    text = value.strip().replace(" ", "").replace(",", "")
    if not text or text.startswith("<"):
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return int(number) if number.is_integer() else number


def _path_matches(value: Any, expected: Path, *, base: Path = DOCROOT) -> bool:
    try:
        return _record_path(value, base=base).resolve() == expected.resolve()
    except OSError:
        return False


def _artifact_summary(
    profile_dir: Path,
    name: str,
    record: Any,
    cache: dict[str, dict[str, Any]],
    issues: Issues,
    label: str,
    required: bool,
) -> dict[str, Any]:
    if not isinstance(record, dict):
        issues.error(f"{label}: artifact {name} metadata is not an object")
        return {"present": False, "verified": False, "reason": "metadata_not_object"}
    path_value = record.get("path")
    path = _record_path(path_value) if path_value is not None else profile_dir / name
    present = bool(record.get("present"))
    actual_present = path.is_file() and not path.is_symlink()
    verified = True
    reason: str | None = None
    if present != actual_present:
        verified = False
        reason = "present_flag_differs"
    if required and not actual_present:
        verified = False
        reason = reason or "required_artifact_missing"
    actual: dict[str, Any] | None = None
    if actual_present:
        actual = _metadata(path, cache)
        for key in ("bytes", "sha256"):
            if key in record and record[key] != actual[key]:
                verified = False
                reason = reason or f"{key}_differs"
    if not verified:
        issues.error(f"{label}: artifact {name}: {reason}")
    return {
        "path": _display(path),
        "present": actual_present,
        "verified": verified,
        "bytes": actual["bytes"] if actual else None,
        "sha256": actual["sha256"] if actual else None,
        "required": required,
        "reason": reason,
    }


def _check_metadata(
    expected: Any,
    actual_path: Path,
    label: str,
    cache: dict[str, dict[str, Any]],
    issues: Issues,
) -> dict[str, Any] | None:
    if not isinstance(expected, dict):
        issues.error(f"{label}: metadata is not an object")
        return None
    if not actual_path.is_file() or actual_path.is_symlink():
        issues.error(f"{label}: file is missing: {actual_path}")
        return None
    actual = _metadata(actual_path, cache)
    for key in ("bytes", "sha256"):
        # Build bindings carry the JSON receipt hash but intentionally do not
        # duplicate the receipt file size.  Compare only fields the producer
        # recorded while still requiring the path to exist.
        if key in expected and expected.get(key) != actual[key]:
            issues.error(f"{label}: {key} does not match")
    return actual


def _empty_histogram() -> dict[str, int]:
    return {bucket: 0 for bucket in SIZE_BUCKETS}


def _bucket(size: int) -> str:
    if size <= 0:
        return "bytes_0"
    if size <= 512:
        return "bytes_1_to_512"
    if size <= 4096:
        return "bytes_513_to_4096"
    if size <= 16384:
        return "bytes_4097_to_16384"
    if size <= 65536:
        return "bytes_16385_to_65536"
    return "bytes_over_65536"


def _perf_stat(path: Path, artifact: dict[str, Any] | None) -> dict[str, Any]:
    counters: dict[str, dict[str, Any]] = {
        event: {
            "value": None,
            "available": False,
            "raw": None,
            "perf_runtime_field": None,
            "running_percent": None,
            "multiplexed": None,
        }
        for event in PERF_EVENTS
    }
    if not path.is_file():
        return {
            "status": "unavailable",
            "scope": "whole child process; profiler overhead included",
            "artifact": artifact,
            "counters": counters,
            "multiplexing": {"status": "unavailable", "events": []},
        }
    with path.open("r", encoding="utf-8", errors="replace", newline="") as stream:
        for row in csv.reader(stream):
            if len(row) < 3:
                continue
            event = row[2].strip()
            if event not in counters:
                continue
            value = _number(row[0])
            runtime = _number(row[3]) if len(row) > 3 else None
            running = None
            if len(row) > 4:
                try:
                    running = float(row[4].strip())
                except ValueError:
                    running = None
            counters[event] = {
                "value": value,
                "available": value is not None,
                "raw": ",".join(row).strip(),
                "perf_runtime_field": runtime,
                "running_percent": running,
                "multiplexed": running is not None and running < 99.999,
            }
            if event in CACHE_EVENTS and value == 0:
                counters[event]["zero_note"] = (
                    "raw zero retained; this event is uninformative for a cache "
                    "claim and is not used in a miss ratio"
                )
    observed = [
        event
        for event, counter in counters.items()
        if counter["available"]
    ]
    running_values = [
        counter["running_percent"]
        for counter in counters.values()
        if counter["available"] and counter["running_percent"] is not None
    ]
    return {
        "status": "observed" if observed else "unavailable",
        "scope": "whole child process; profiler overhead included",
        "artifact": artifact,
        "events": list(PERF_EVENTS),
        "counters": counters,
        "multiplexing": {
            "status": "observed" if running_values else "unavailable",
            "events": observed,
            "any_multiplexed": any(
                counters[event]["multiplexed"] is True for event in observed
            ),
            "running_percent_min": min(running_values) if running_values else None,
            "running_percent_max": max(running_values) if running_values else None,
            "interpretation": (
                "perf's emitted value and runtime fields are retained; a running "
                "percentage below 100 indicates multiplexing or partial scheduling"
            ),
        },
        "cache_interpretation": {
            "status": "unsupported_for_cache_claims",
            "zero_events": [
                event
                for event in CACHE_EVENTS
                if counters[event]["available"] and counters[event]["value"] == 0
            ],
            "miss_ratio": "not_computed",
            "note": (
                "Raw zero cache counters and unsupported LLC events are retained; "
                "no zero-denominator cache ratio or zero-miss claim is made."
            ),
        },
    }


_STRACE_LINE = re.compile(
    r"^\s*\d+\s+\S+\s+(?P<syscall>[A-Za-z0-9_]+)\((?P<body>.*)$"
)
_STRACE_RETURN = re.compile(r"\s=\s(-?\d+)(?:\s|$)")
_ANGLE_PATH = re.compile(r"<([^>\n]+)>")
_QUOTED_PATH = re.compile(r'"((?:\\.|[^"\\])*)"')


def _strace_path(line: str) -> str | None:
    angle = [item for item in _ANGLE_PATH.findall(line) if "/" in item]
    if angle:
        # For openat, the final annotation is the opened path; for read and
        # close there is only one.  The cwd annotation is therefore ignored.
        return angle[-1]
    quoted = [item for item in _QUOTED_PATH.findall(line) if "/" in item]
    if quoted:
        return quoted[0].replace(r'\"', '"')
    return None


def _new_syscall_counts() -> dict[str, dict[str, Any]]:
    return {
        syscall: {
            "calls": 0,
            "successful_calls": 0,
            "errors": 0,
            "return_bytes": 0,
            "size_histogram": _empty_histogram(),
        }
        for syscall in TRACE_SYSCALLS
    }


def _strace_scope() -> dict[str, Any]:
    return {
        "syscalls": _new_syscall_counts(),
        "fd_lifecycle": {syscall: 0 for syscall in FD_SYSCALLS},
        "lines": 0,
        "matched_lines": 0,
        "unparsed_returns": 0,
    }


def _strace(path: Path, receipt: dict[str, Any], profile_dir: Path) -> dict[str, Any]:
    if not path.is_file():
        return {
            "status": "unavailable",
            "scope": "whole child process; strace overhead included",
            "reason": "strace artifact is missing",
        }
    replay = receipt.get("replay", {}).get("directory")
    replay_text = str(replay) if replay else None
    known_paths = {str(profile_dir.resolve())}
    artifacts = receipt.get("artifacts", {})
    if isinstance(artifacts, dict):
        for item in artifacts.values():
            if isinstance(item, dict) and item.get("path"):
                known_paths.add(str(_record_path(item["path"]).resolve()))
    scopes = {
        "whole_child": _strace_scope(),
        "replay_path": _strace_scope() if replay_text else None,
        "setup_report_io": _strace_scope(),
        "other_known_path": _strace_scope(),
        "unknown_path": _strace_scope(),
    }
    with path.open("r", encoding="utf-8", errors="replace") as stream:
        for line in stream:
            scopes["whole_child"]["lines"] += 1
            match = _STRACE_LINE.match(line)
            if not match:
                continue
            syscall = match.group("syscall")
            path_hint = _strace_path(line)
            if replay_text and replay_text in line:
                category = "replay_path"
            elif path_hint and any(
                path_hint == known or path_hint.startswith(known + "/")
                for known in known_paths
            ):
                category = "setup_report_io"
            elif path_hint:
                category = "other_known_path"
            else:
                category = "unknown_path"
            selected = [scopes["whole_child"]]
            if category != "whole_child" and scopes[category] is not None:
                selected.append(scopes[category])
            for scope in selected:
                scope["lines"] += 0
                if syscall in FD_SYSCALLS:
                    scope["fd_lifecycle"][syscall] += 1
                if syscall not in TRACE_SYSCALLS:
                    continue
                scope["matched_lines"] += 1
                entry = scope["syscalls"][syscall]
                entry["calls"] += 1
                returned_match = _STRACE_RETURN.search(line)
                if returned_match is None:
                    scope["unparsed_returns"] += 1
                    continue
                returned = int(returned_match.group(1))
                if returned < 0:
                    entry["errors"] += 1
                    continue
                entry["successful_calls"] += 1
                if syscall in PAYLOAD_SYSCALLS:
                    entry["return_bytes"] += returned
                    entry["size_histogram"][_bucket(returned)] += 1
    for scope in scopes.values():
        if scope is None:
            continue
        # ``lines`` is only meaningful for the whole-child trace.  Category
        # line counts are represented by matched syscall counts instead.
        scope.pop("lines", None)
    result = {
        "status": "observed",
        "scope": "whole child process; strace overhead included",
        "path_attribution": {
            "replay_directory": replay_text,
            "replay_status": "observed" if replay_text else "not_applicable",
            "known_setup_report_paths": sorted(known_paths),
            "classification": (
                "replay_path is an exact path-string match; setup_report_io is a "
                "profile/report/resource path match; other calls remain separate"
            ),
        },
        "scopes": scopes,
    }
    return result


def _scaled_bytes(value: str) -> int | None:
    match = re.fullmatch(r"\s*([0-9]+(?:\.[0-9]+)?)\s*([KMGTP]?)B?\s*", value)
    if not match:
        return None
    number = float(match.group(1))
    multiplier = {"": 1, "K": 1024, "M": 1024**2, "G": 1024**3, "T": 1024**4, "P": 1024**5}
    return int(number * multiplier[match.group(2)])


def _heaptrack(path: Path, artifact: dict[str, Any] | None) -> dict[str, Any]:
    if not path.is_file():
        return {
            "status": "unavailable",
            "scope": "whole child process; heaptrack overhead included",
            "artifact": artifact,
            "reason": "heaptrack_print artifact is missing",
        }
    lines = _read_text(path).splitlines()
    rows: list[dict[str, Any]] = []
    section: str | None = None
    headline_re = re.compile(
        r"^\s*([0-9.]+[KMGTP]?)\s+(peak memory consumed|temporary allocations)"
        r"(?: over ([0-9,]+) calls)?"
    )
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped in {
            "PEAK MEMORY CONSUMERS",
            "MOST CALLS TO ALLOCATION FUNCTIONS",
            "MOST TEMPORARY ALLOCATIONS",
        }:
            section = stripped
            continue
        if section != "PEAK MEMORY CONSUMERS":
            continue
        match = headline_re.match(line)
        if not match or "peak memory consumed" not in match.group(2):
            continue
        symbols: list[str] = []
        for candidate in lines[index + 1 : index + 35]:
            candidate = candidate.strip()
            if candidate.startswith("at ") or candidate.startswith("in "):
                continue
            if candidate and not candidate.startswith("_"):
                continue
            if candidate:
                symbols.append(candidate)
        frame = symbols[0] if symbols else None
        application_frame = next(
            (
                symbol
                for symbol in symbols
                if ("litchi_docx" in symbol or "litchi_perf_baseline" in symbol)
                and "alloc" not in symbol
            ),
            None,
        )
        rows.append(
            {
                "headline": stripped,
                "peak_display": match.group(1),
                "peak_bytes_base_1024": _scaled_bytes(match.group(1)),
                "calls": int(match.group(3).replace(",", "")) if match.group(3) else None,
                "first_symbol": frame,
                "first_application_symbol": application_frame,
                "symbols": symbols[:8],
            }
        )
    tail_patterns = {
        "allocation_calls": re.compile(r"calls to allocation functions:\s*([0-9,]+)"),
        "temporary_allocations": re.compile(r"temporary memory allocations:\s*([0-9,]+)"),
        "peak_heap_display": re.compile(r"peak heap memory consumption:\s*(.+)$"),
        "peak_rss_display": re.compile(r"peak RSS \(including heaptrack overhead\):\s*(.+)$"),
        "total_leaked_display": re.compile(r"total memory leaked:\s*(.+)$"),
    }
    totals: dict[str, Any] = {}
    for line in lines:
        for key, pattern in tail_patterns.items():
            match = pattern.search(line)
            if match:
                value = match.group(1).strip()
                totals[key] = value
                if key in {"peak_heap_display", "peak_rss_display", "total_leaked_display"}:
                    totals[key.replace("_display", "_bytes_base_1024")] = _scaled_bytes(value)
                else:
                    totals[key] = int(value.replace(",", ""))
    return {
        "status": "observed",
        "scope": "whole child process; heaptrack overhead included",
        "artifact": artifact,
        "top_only": False,
        "top_rows": rows[:10],
        "top_rows_note": (
            "The rows are a presentation subset of the complete heaptrack print; "
            "the artifact hash and tail totals bind the whole-child report."
        ),
        "totals": totals,
        "print_lines": len(lines),
    }


_PERF_HEADER = re.compile(r"^\s*[^#\s].*:.*$")


def _perf_script(path: Path, expected_binary: dict[str, Any]) -> dict[str, Any]:
    if not path.is_file() or path.stat().st_size == 0:
        return {
            "status": "pending",
            "scope": "whole child process; perf-record sampling overhead included",
            "reason": "no perf script export has been retained",
            "whole_child_samples": None,
            "marker_status": "unavailable",
        }
    text = _read_text(path)
    lines = text.splitlines()
    samples: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    command_match = False
    expected_path = str(expected_binary.get("path", ""))
    expected_name = Path(expected_path).name
    for line in lines:
        if expected_path and expected_path in line:
            command_match = True
        if line.startswith("#"):
            continue
        if not line.strip():
            if current is not None and current["frames"]:
                samples.append(current)
            current = None
            continue
        if not line[0].isspace() and _PERF_HEADER.match(line):
            if current is not None and current["frames"]:
                samples.append(current)
            current = {"header": line.strip(), "frames": []}
            continue
        if current is not None:
            frame = line.strip()
            if frame:
                current["frames"].append(frame)
    if current is not None and current["frames"]:
        samples.append(current)
    if not samples:
        return {
            "status": "unavailable",
            "scope": "whole child process; perf-record sampling overhead included",
            "reason": "perf script contained no parseable stack samples",
            "whole_child_samples": 0,
            "marker_status": "unavailable",
            "binary_command_match": command_match or (expected_name in text if expected_name else False),
        }
    marker_counts: dict[str, int] = {}
    marker_symbols: dict[str, list[str]] = {}
    membership: dict[str, set[int]] = {}
    for group, patterns in STACK_MARKERS.items():
        compiled = [re.compile(pattern) for pattern in patterns]
        members: set[int] = set()
        symbols: set[str] = set()
        for index, sample in enumerate(samples):
            matched = [
                frame
                for frame in sample["frames"]
                if any(pattern.search(frame) for pattern in compiled)
            ]
            if matched:
                members.add(index)
                symbols.update(matched)
        membership[group] = members
        marker_counts[group] = len(members)
        marker_symbols[group] = sorted(symbols)[:20]
    overlaps: dict[str, int] = {}
    groups = tuple(STACK_MARKERS)
    for left_index, left in enumerate(groups):
        for right in groups[left_index + 1 :]:
            overlaps[f"{left}+{right}"] = len(membership[left] & membership[right])
    overlaps["all_three"] = len(set.intersection(*(membership[group] for group in groups)))
    return {
        "status": "observed",
        "scope": "whole child process; perf-record sampling overhead included",
        "whole_child_samples": len(samples),
        "marker_status": "observed" if any(marker_counts.values()) else "unavailable",
        "markers": {
            group: {
                "inclusive_samples": marker_counts[group],
                "share_of_whole_child": marker_counts[group] / len(samples),
                "actual_matching_frames": marker_symbols[group],
                "status": "observed" if marker_counts[group] else "unavailable",
            }
            for group in groups
        },
        "overlap_samples": overlaps,
        "overlap_note": "Group shares are inclusive and overlap; they are not additive.",
        "binary_command_match": command_match or (expected_name in text if expected_name else False),
        "binary_command_match_note": (
            "A perf script text header is only a name/path check; the sidecar "
            "binding retains the source perf.data and build binary hashes."
        ),
    }


def _route_axis(protocol: dict[str, Any], route: str) -> dict[str, Any] | None:
    for item in protocol.get("route_axis", []):
        if isinstance(item, dict) and item.get("route") == route:
            return item
    return None


def _case_axis(protocol: dict[str, Any], label: str) -> dict[str, Any] | None:
    for item in protocol.get("cases", []):
        if isinstance(item, dict) and item.get("label") == label:
            return item
    return None


def _compare(value: Any, expected: Any, label: str, issues: Issues) -> None:
    if value != expected:
        issues.error(f"{label}: expected {expected!r}, found {value!r}")


def _validate_protocol_report(
    protocol: dict[str, Any],
    receipt: dict[str, Any],
    report: dict[str, Any],
    label: str,
    issues: Issues,
) -> None:
    route = receipt.get("route", {})
    case = receipt.get("case", {})
    route_name = route.get("name")
    case_label = case.get("label")
    axis = _route_axis(protocol, str(route_name))
    case_axis = _case_axis(protocol, str(case_label))
    if axis is None:
        issues.error(f"{label}: route {route_name!r} is absent from protocol route_axis")
    if case_axis is None:
        issues.error(f"{label}: case {case_label!r} is absent from protocol cases")
    config = report.get("config") if isinstance(report, dict) else None
    if not isinstance(config, dict):
        issues.error(f"{label}: report config is missing")
        return
    if axis is not None:
        for key in ("cli_provider", "report_provider", "replay_max_bytes", "replay_sync"):
            _compare(route.get(key), axis.get(key), f"{label}: route.{key}", issues)
        _compare(config.get("expected_authored_opens"), axis.get("expected_authored_opens"), f"{label}: report expected_authored_opens", issues)
        _compare(config.get("expected_replay_opens"), axis.get("expected_replay_opens"), f"{label}: report expected_replay_opens", issues)
        for key in ("provider", "authored_provider", "replay_max_bytes", "replay_sync", "compression", "input_mode", "input_storage_kind", "sink_write_bytes"):
            expected_key = key
            expected = axis.get(expected_key)
            if expected is None and key == "provider":
                expected = route_name
            if expected is None and key == "authored_provider":
                expected = axis.get("report_provider")
            if expected is not None:
                _compare(config.get(key), expected, f"{label}: report config.{key}", issues)
    if case_axis is not None:
        for key in ("source_count", "authored_count", "text_mode", "input_mode", "input_backing", "compression", "sink_write_bytes", "source_contract"):
            if key in case_axis:
                if key in {"source_count", "authored_count", "text_mode", "input_mode", "input_backing", "compression", "source_contract"}:
                    actual = case.get(key)
                else:
                    actual = case.get(key, config.get(key))
                _compare(actual, case_axis.get(key), f"{label}: case.{key}", issues)
    _compare(report.get("schema"), "docx-replayable-tail-append-v1", f"{label}: report schema", issues)
    _compare(report.get("version"), 1, f"{label}: report version", issues)


def _required_artifacts(tool: str) -> set[str]:
    required = {"report.json", "resource.txt", "stdout.txt", "stderr.txt"}
    if tool == "perf-stat":
        required.add("perf-stat.txt")
    elif tool == "perf-record":
        required.add("perf.data")
    elif tool == "strace":
        required.add("strace.log")
    elif tool == "heaptrack":
        # The capture suffix is version-dependent; validation below accepts
        # any nonempty heaptrack-profile* artifact.
        required.update(
            {
                "heaptrack-print.txt",
                "heaptrack-histogram.tsv",
                "heaptrack-print-stderr.txt",
            }
        )
    return required


def _validate_common_binding(
    label: str,
    receipt: dict[str, Any],
    started: dict[str, Any],
    protocol_path: Path,
    protocol_sha: str,
    driver_path: Path,
    driver_sha: str,
    build_path: Path,
    build_sha: str,
    binary: dict[str, Any],
    cache: dict[str, dict[str, Any]],
    issues: Issues,
) -> None:
    if receipt.get("schema") != PROFILE_RUN_SCHEMA:
        issues.error(f"{label}: profile receipt schema differs")
    if receipt.get("label") != label:
        issues.error(f"{label}: receipt label differs")
    if receipt.get("profile_driver_sha256") != driver_sha:
        issues.error(f"{label}: profile helper hash differs")
    for binding_name, value in (("protocol", receipt.get("protocol")), ("protocol", started.get("protocol"))):
        if not isinstance(value, dict) or value.get("path") != "route-protocol.json" or value.get("sha256") != protocol_sha:
            issues.error(f"{label}: {binding_name} binding differs from route-protocol.json")
    driver = receipt.get("driver")
    if not isinstance(driver, dict) or driver.get("sha256") != driver_sha:
        issues.error(f"{label}: helper driver binding is missing or stale")
    elif not _path_matches(driver.get("path"), driver_path, base=DOCROOT):
        issues.error(f"{label}: helper driver path differs")
    build = receipt.get("build")
    if not isinstance(build, dict):
        issues.error(f"{label}: build binding is missing")
    else:
        if build.get("sha256") != build_sha:
            issues.error(f"{label}: build receipt hash differs")
        if not _path_matches(build.get("path"), build_path, base=DOCROOT):
            issues.error(f"{label}: build receipt path differs")
        if build.get("protocol", {}).get("sha256") != protocol_sha:
            issues.error(f"{label}: build protocol hash differs")
        if build.get("role") != "normal" or build.get("source_before") != build.get("source_after"):
            issues.error(f"{label}: normal build/source custody binding is invalid")
    if receipt.get("binary") != binary:
        issues.error(f"{label}: binary identity differs from profile-run binding")
    binary_path = _record_path(binary.get("path"), base=REPO)
    if not binary_path.is_file() or binary_path.is_symlink():
        issues.error(f"{label}: bound binary is missing: {binary_path}")
    else:
        actual_binary = _metadata(binary_path, cache)
        for key in ("bytes", "sha256"):
            if binary.get(key) != actual_binary[key]:
                issues.error(f"{label}: bound binary {key} differs")
    replay = receipt.get("replay")
    if isinstance(replay, dict) and replay.get("route") == "file_store":
        if not replay.get("directory"):
            issues.error(f"{label}: file-store profile has no replay directory binding")
        if receipt.get("status") == "ok" and receipt.get("scratch_cleaned") is not True:
            issues.error(f"{label}: successful file-store profile did not clean replay scratch")


def _perf_export(
    row: dict[str, Any],
    receipt: dict[str, Any],
    perf_script_root: Path,
    cache: dict[str, dict[str, Any]],
    issues: Issues,
) -> dict[str, Any] | None:
    if row["tool"] != "perf-record":
        return None
    artifacts = row["artifacts"]
    perf_data = row["profile_dir"] / "perf.data"
    artifact = artifacts.get("perf.data")
    if not perf_data.is_file():
        return {
            "status": "unavailable",
            "reason": "perf.data artifact is missing",
            "command": None,
        }
    perf_meta = _metadata(perf_data, cache)
    export_dir = perf_script_root / row["label"]
    export_path = export_dir / "perf-script.txt"
    binding_path = export_dir / "binding.json"
    build = receipt.get("build", {})
    command = [
        "perf",
        "script",
        "--header",
        "-i",
        str(perf_data.resolve()),
        "--demangle",
    ]
    export: dict[str, Any] = {
        "schema": PERF_SCRIPT_SCHEMA,
        "status": "pending",
        "profile_label": row["label"],
        "command": command,
        "working_directory": str(REPO),
        "input_perf_data": {"path": _display(perf_data), **perf_meta},
        "input_receipt_artifact": artifact,
        "binary": receipt.get("binary"),
        "build": {
            "path": build.get("path"),
            "sha256": build.get("sha256"),
        },
        "protocol": receipt.get("protocol"),
        "output": {
            "path": _display(export_path),
            "binding_path": _display(binding_path),
        },
    }
    if export_path.is_file() and not export_path.is_symlink() and export_path.stat().st_size > 0:
        export["script_artifact"] = {"path": _display(export_path), **_metadata(export_path, cache)}
        if binding_path.is_file() and not binding_path.is_symlink():
            try:
                binding = _json(binding_path)
            except (OSError, ValueError) as error:
                issues.error(f"{row['label']}: perf script binding is not JSON: {error}")
                binding = None
            if isinstance(binding, dict):
                expected = {
                    "schema": PERF_SCRIPT_SCHEMA,
                    "profile_label": row["label"],
                    "perf_data_sha256": perf_meta["sha256"],
                    "binary_sha256": receipt.get("binary", {}).get("sha256"),
                    "protocol_sha256": receipt.get("protocol", {}).get("sha256"),
                    "script_sha256": export["script_artifact"]["sha256"],
                }
                for key, value in expected.items():
                    if binding.get(key) != value:
                        issues.error(f"{row['label']}: perf script binding {key} differs")
                export["binding"] = binding
        else:
            # The binding is generated by this helper on the next invocation;
            # no profile directory is modified and the source hashes are fixed
            # by the receipt/perf.data chain above.
            binding = {
                "schema": PERF_SCRIPT_SCHEMA,
                "profile_label": row["label"],
                "perf_data_sha256": perf_meta["sha256"],
                "binary_sha256": receipt.get("binary", {}).get("sha256"),
                "protocol_sha256": receipt.get("protocol", {}).get("sha256"),
                "script_sha256": _metadata(export_path, cache)["sha256"],
            }
            export_dir.mkdir(parents=True, exist_ok=True)
            binding_path.write_text(json.dumps(binding, indent=2, sort_keys=True) + "\n", encoding="utf-8")
            export["binding"] = binding
        export["status"] = "observed"
        export["analysis"] = _perf_script(export_path, receipt.get("binary", {}))
    return export


def _load_expected_labels(started: dict[str, Any]) -> list[str]:
    labels: list[str] = []
    for tool in started.get("tools", []):
        for route in started.get("routes", []):
            for case in started.get("cases", []):
                labels.append(f"{tool}-{route}-{case}")
    return labels


def _validate_profile_root(
    profiles_dir: Path,
    result: dict[str, Any],
    started: dict[str, Any],
    protocol: dict[str, Any],
    protocol_path: Path,
    protocol_sha: str,
    issues: Issues,
) -> None:
    if result.get("schema") != PROFILE_SCHEMA:
        issues.error("profile result schema differs")
    if result.get("attempt") != profiles_dir.name:
        issues.error("profile result attempt does not match directory")
    if result.get("status") != "ok" or result.get("passed") is not True:
        issues.error("profile result is not a passed terminal result")
    if result.get("binary_unchanged") is not True:
        issues.error("profile result binary_unchanged is false")
    if result.get("scratch_removed") is not True:
        issues.error("profile result scratch_removed is false")
    if result.get("failures"):
        issues.error("profile result contains failures")
    if started.get("schema") != PROFILE_SCHEMA:
        issues.error("profile started schema differs")
    if started.get("protocol", {}).get("sha256") != protocol_sha:
        issues.error("profile started protocol hash differs")
    if started.get("profile_driver_sha256") != result.get("profile_driver_sha256"):
        issues.error("profile result and started helper hashes differ")
    if started.get("protocol", {}).get("path") != protocol_path.name:
        issues.error("profile started protocol path differs")
    expected = _load_expected_labels(started)
    receipt_labels = [item.get("label") for item in result.get("receipts", []) if isinstance(item, dict)]
    if sorted(receipt_labels) != sorted(expected):
        issues.error(
            f"profile result receipt inventory differs: expected {len(expected)}, found {len(receipt_labels)}"
        )


def summarize(
    profiles_dir: Path,
    protocol_path: Path,
    output_path: Path,
    perf_script_root: Path,
    *,
    allow_incomplete: bool = False,
) -> tuple[dict[str, Any], int]:
    issues = Issues()
    cache: dict[str, dict[str, Any]] = {}
    profiles_dir = profiles_dir.resolve()
    protocol_path = protocol_path.resolve()
    output_path = output_path.resolve()
    perf_script_root = perf_script_root.resolve()
    result_path = profiles_dir / "result.json"
    started_path = profiles_dir / "started.json"
    try:
        result = _json(result_path)
    except (OSError, ValueError) as error:
        result = {}
        issues.error(f"cannot read profile result: {error}")
    try:
        started = _json(started_path)
    except (OSError, ValueError) as error:
        started = {}
        issues.error(f"cannot read profile started binding: {error}")
    try:
        protocol = _json(protocol_path)
    except (OSError, ValueError) as error:
        protocol = {}
        issues.error(f"cannot read route protocol: {error}")
    if protocol_path.is_file():
        protocol_sha = _sha256(protocol_path)
    else:
        protocol_sha = ""
    _validate_profile_root(profiles_dir, result, started, protocol, protocol_path, protocol_sha, issues)

    driver_ref = started.get("driver", {}) if isinstance(started, dict) else {}
    driver_path = _record_path(driver_ref.get("path", "profile_routes.py"))
    driver_meta = _check_metadata(driver_ref, driver_path, "profile helper", cache, issues)
    driver_sha = driver_meta["sha256"] if driver_meta else ""
    build_ref = started.get("build", {}) if isinstance(started, dict) else {}
    build_path = _record_path(build_ref.get("path", ""))
    build_meta = _check_metadata(build_ref, build_path, "normal build receipt", cache, issues)
    build_sha = build_meta["sha256"] if build_meta else ""
    binary = started.get("binary") if isinstance(started, dict) else None
    if not isinstance(binary, dict):
        binary = result.get("binary") if isinstance(result, dict) else {}
    if not isinstance(binary, dict):
        binary = {}

    records: list[dict[str, Any]] = []
    expected_labels = _load_expected_labels(started)
    for label in expected_labels:
        profile_dir = profiles_dir / label
        receipt_path = profile_dir / "receipt.json"
        row_issues_before = len(issues.errors)
        try:
            receipt = _json(receipt_path)
        except (OSError, ValueError) as error:
            if allow_incomplete:
                issues.warning(f"{label}: terminal receipt is not available: {error}")
                records.append({"label": label, "status": "pending", "binding_status": "unavailable"})
                continue
            issues.error(f"{label}: terminal receipt is not available: {error}")
            records.append({"label": label, "status": "missing", "binding_status": "invalid"})
            continue
        if not isinstance(receipt, dict):
            issues.error(f"{label}: receipt is not an object")
            continue
        status = receipt.get("status")
        if status not in {"ok", "unavailable", "failed"}:
            issues.error(f"{label}: unknown terminal status {status!r}")
        if status == "failed":
            issues.error(f"{label}: profiler receipt failed")
        if status == "unavailable":
            issues.warning(f"{label}: profiler observation is unavailable")
        route_name = receipt.get("route", {}).get("name")
        case = receipt.get("case", {})
        tool = receipt.get("tool", {}).get("name")
        _validate_common_binding(
            label,
            receipt,
            started,
            protocol_path,
            protocol_sha,
            driver_path,
            driver_sha,
            build_path,
            build_sha,
            binary,
            cache,
            issues,
        )
        artifact_records: dict[str, dict[str, Any]] = {}
        artifact_map = receipt.get("artifacts", {})
        if not isinstance(artifact_map, dict):
            issues.error(f"{label}: receipt artifacts are not an object")
            artifact_map = {}
        required = _required_artifacts(str(tool)) if status == "ok" else set()
        for name, value in sorted(artifact_map.items()):
            artifact_records[name] = _artifact_summary(
                profile_dir,
                name,
                value,
                cache,
                issues,
                label,
                name in required,
            )
        for name in sorted(required - artifact_records.keys()):
            issues.error(f"{label}: required artifact is absent from receipt: {name}")
        if tool == "heaptrack" and status == "ok":
            captures = [
                item
                for name, item in artifact_records.items()
                if name.startswith("heaptrack-profile") and item.get("present")
            ]
            if not captures:
                issues.error(f"{label}: successful heaptrack receipt has no capture artifact")
        report: dict[str, Any] = {}
        report_path = profile_dir / "report.json"
        if report_path.is_file():
            try:
                report = _json(report_path)
                _validate_protocol_report(protocol, receipt, report, label, issues)
            except (OSError, ValueError) as error:
                issues.error(f"{label}: route report cannot be read: {error}")
        elif status == "ok":
            issues.error(f"{label}: route report is missing")
        row: dict[str, Any] = {
            "label": label,
            "status": status,
            "tool": tool,
            "route": route_name,
            "case": case.get("label"),
            "source_count": case.get("source_count"),
            "authored_count": case.get("authored_count"),
            "text_mode": case.get("text_mode"),
            "chunk_mode": case.get("chunk_mode"),
            "profile_dir": profile_dir,
            "binding_status": "ok" if len(issues.errors) == row_issues_before else "invalid",
            "artifacts": artifact_records,
            "report": {
                "path": _display(report_path),
                "schema": report.get("schema"),
                "version": report.get("version"),
            },
            "perf_stat": None,
            "strace": None,
            "heaptrack": None,
            "perf_script": None,
        }
        if tool == "perf-stat":
            row["perf_stat"] = _perf_stat(profile_dir / "perf-stat.txt", artifact_records.get("perf-stat.txt"))
        elif tool == "strace":
            row["strace"] = _strace(profile_dir / "strace.log", receipt, profile_dir)
        elif tool == "heaptrack":
            row["heaptrack"] = _heaptrack(profile_dir / "heaptrack-print.txt", artifact_records.get("heaptrack-print.txt"))
        records.append(row)

    # Convert internal Path objects only at the output boundary.
    export_records: list[dict[str, Any]] = []
    for row in records:
        if row.get("status") != "pending":
            export = _perf_export(row, _json(row["profile_dir"] / "receipt.json"), perf_script_root, cache, issues)
            row["perf_script"] = export
        row["profile_dir"] = _display(row["profile_dir"])
        export_records.append(row)
    command_records = [
        row["perf_script"]
        for row in export_records
        if row.get("perf_script") is not None
    ]
    summary = {
        "schema": SUMMARY_SCHEMA,
        "version": 1,
        "status": "ok" if not issues.errors else ("incomplete" if allow_incomplete else "invalid"),
        "performance_claim": "none",
        "scope": "whole child process; profiler overhead included",
        "profile_root": _display(profiles_dir),
        "profile_result": {
            "path": _display(result_path),
            "sha256": _sha256(result_path) if result_path.is_file() else None,
        },
        "bindings": {
            "protocol": {"path": _display(protocol_path), "sha256": protocol_sha},
            "profile_driver": {"path": _display(driver_path), "sha256": driver_sha},
            "build": {"path": _display(build_path), "sha256": build_sha},
            "binary": binary,
            "machine": started.get("machine"),
        },
        "records": export_records,
        "perf_script_commands": {
            "root": _display(perf_script_root),
            "records": command_records,
            "pending_count": sum(1 for item in command_records if item.get("status") == "pending"),
        },
        "validation": {
            "errors": issues.errors,
            "warnings": issues.warnings,
            "terminal_records": len(export_records),
            "expected_records": len(expected_labels),
        },
        "limitations": [
            "External observations include profiler and process-launch overhead and cover the whole child.",
            "perf-stat values retain perf's raw event rows and running percentages; unavailable events remain null.",
            "strace path categories use exact path annotations from -yy where present; unknown paths are not guessed.",
            "heaptrack top rows are a presentation subset; whole-child totals and the complete print artifact remain bound by hash.",
            "perf script marker shares are inclusive stack ancestry and overlap; they are not additive or causal attribution.",
            "No optimization, speedup, allocation, or cold-cache claim is authorized by this diagnostic summary.",
        ],
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    manifest = {
        "schema": PERF_SCRIPT_SCHEMA,
        "version": 1,
        "performance_claim": "none",
        "instructions": "Run each pending command, then rerun this summarizer to hash and bind the export.",
        "records": command_records,
    }
    manifest_path = output_path.with_name(output_path.stem + "-perf-script-commands.json")
    manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return summary, 0 if not issues.errors else 1


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--profiles-dir",
        type=Path,
        default=DOCROOT / "route-profiles" / "profiles1",
        help="terminal profile directory (default: route-profiles/profiles1)",
    )
    parser.add_argument("--protocol", type=Path, default=DOCROOT / "route-protocol.json")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--perf-script-root", type=Path)
    parser.add_argument(
        "--allow-incomplete",
        action="store_true",
        help="retain a partial summary while profiles are still arriving",
    )
    return parser


def main() -> int:
    args = _parser().parse_args()
    profiles_dir = args.profiles_dir
    output = args.output or profiles_dir.parent / f"{profiles_dir.name}-summary.json"
    perf_script_root = args.perf_script_root or profiles_dir.parent / f"{profiles_dir.name}-perf-script"
    summary, status = summarize(
        profiles_dir,
        args.protocol,
        output,
        perf_script_root,
        allow_incomplete=args.allow_incomplete,
    )
    print(
        f"{summary['status']}: {summary['validation']['terminal_records']} records; "
        f"{len(summary['validation']['errors'])} errors; {len(summary['validation']['warnings'])} warnings; "
        f"output={output}",
        file=sys.stderr if status else sys.stdout,
    )
    return status


if __name__ == "__main__":
    raise SystemExit(main())
