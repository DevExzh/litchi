#!/usr/bin/env python3
"""Bounded, offline attribution for interpreted Heaptrack v3 traces.

Heaptrack's interpreted stream is a small line protocol.  The ``a`` records
describe an allocation size and a backtrace, while ``+`` and ``-`` records
refer to those descriptors.  ``heaptrack_print`` is useful for a human
summary, but it does not retain the descriptor-to-stack relationship needed
for an operation owner audit.  This reader keeps that relationship and emits
one row for every allocation descriptor referenced by an event.

The only operation owner understood by this packet is the exact Rust symbol
``opc_preservation_probe::profile_publish``.  A compiler hash suffix is
accepted only in the canonical form ``::h`` followed by sixteen lower-case
hexadecimal characters.  A matching allocation is counted once even if an
unusual trace contains the owner more than once; the duplicate is retained as
a diagnostic and the innermost (leaf-most) owner frame controls phase
classification.

All frames are retained in leaf-to-root order, including inline frames.  The
    phase buckets are deliberately conservative: they use concrete symbols in
    the actual full stack around the innermost owner.  A stack with no
recognized phase marker is put in ``unknown`` rather than guessed from the
workload case.  The reader makes no claim about peak live memory, physical
copies, allocator identity, or elapsed time.

The parser accepts the interpreted format emitted by Heaptrack 1.5.x
(``v 10500 3``).  It can consume decoded bytes through :func:`parse`, or a
retained ``.zst`` trace through :func:`parse_zst`; neither entry point starts
the profiled program.
"""

from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import dataclass
import json
from pathlib import Path
import re
import subprocess
import sys
from typing import Any, Iterable, NoReturn


SCHEMA = "litchi.heaptrack.0843.interpreted-v3.v1"
FORMAT = "interpreted_heaptrack_v3"
HEAPTRACK_VERSION = 0x10500
FILE_VERSION = 3
PROFILE_OWNER = "opc_preservation_probe::profile_publish"
PHASES = ("preservation_index", "topology_preparation", "plan", "emission", "unknown")

# These limits are intentionally explicit.  They protect both the zstd pipe
# and the in-memory descriptor/trace tables from a corrupt or hostile input.
MAX_TRACE_BYTES = 512 * 1024 * 1024
MAX_TRACE_RECORDS = 20_000_000
MAX_TRACE_DEPTH = 1_000_000
MAX_STACK_FRAMES = 2_000_000

_HEX = re.compile(rb"^[0-9a-f]+$")
_RUST_HASH = re.compile(r"::h[0-9a-f]{16}")


class HeapReaderError(ValueError):
    """Raised when a trace is malformed or exceeds an input bound."""


# A short alias makes the failure type convenient for packet-local callers.
ReaderError = HeapReaderError


def _fail(message: str) -> NoReturn:
    raise HeapReaderError(message)


def _require(condition: bool, message: str) -> None:
    if not condition:
        _fail(message)


def _hex(token: bytes, label: str, record: int) -> int:
    if not token or len(token) > 16 or _HEX.fullmatch(token) is None:
        _fail(f"invalid {label} hexadecimal token at record {record}: {token!r}")
    return int(token, 16)


def _tokens(line: bytes, mode: bytes, record: int) -> list[bytes]:
    """Parse a numeric record with Heaptrack's single-space separators."""

    if len(line) < 2 or line[:1] != mode or line[1:2] != b" ":
        _fail(f"malformed {mode.decode()} record at record {record}")
    rest = line[2:]
    if not rest:
        return []
    tokens = rest.split(b" ")
    if any(not token for token in tokens) or any(b"\t" in token for token in tokens):
        _fail(f"malformed spacing in {mode.decode()} record at record {record}")
    return tokens


def _one_hex(line: bytes, mode: bytes, label: str, record: int) -> int:
    tokens = _tokens(line, mode, record)
    if len(tokens) != 1:
        _fail(f"{mode.decode()} record needs one field at record {record}")
    return _hex(tokens[0], label, record)


def _two_hex(
    line: bytes,
    mode: bytes,
    first: str,
    second: str,
    record: int,
) -> tuple[int, int]:
    tokens = _tokens(line, mode, record)
    if len(tokens) != 2:
        _fail(f"{mode.decode()} record needs two fields at record {record}")
    return _hex(tokens[0], first, record), _hex(tokens[1], second, record)


def _sized_string(line: bytes, record: int) -> str:
    if len(line) < 3 or line[:2] != b"s ":
        _fail(f"malformed sized string record at record {record}")
    rest = line[2:]
    separator = rest.find(b" ")
    if separator <= 0:
        _fail(f"sized string lacks a length separator at record {record}")
    size = _hex(rest[:separator], "string byte length", record)
    payload = rest[separator + 1 :]
    if len(payload) != size:
        _fail(
            f"string byte length {size} differs from payload {len(payload)} "
            f"at record {record}"
        )
    try:
        return payload.decode("utf-8", errors="strict")
    except UnicodeDecodeError as error:
        _fail(f"string is not UTF-8 at record {record}: {error}")


def _iter_bytes_lines(data: bytes, max_bytes: int) -> Iterable[bytes]:
    _require(len(data) <= max_bytes, f"decoded trace exceeds {max_bytes} bytes")
    offset = 0
    length = len(data)
    while offset < length:
        end = data.find(b"\n", offset)
        if end < 0:
            line = data[offset:]
            offset = length
        else:
            line = data[offset:end]
            offset = end + 1
        if b"\r" in line:
            _fail("carriage returns/CRLF are not supported in the trace")
        yield line


def _zstd_lines(path: Path, max_bytes: int) -> Iterable[bytes]:
    path = Path(path)
    _require(path.suffix == ".zst", f"trace is not a .zst file: {path}")
    _require(path.is_file() and not path.is_symlink(), f"missing trace: {path}")
    try:
        process = subprocess.Popen(
            ["zstd", "--decompress", "--stdout", "--", str(path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except OSError as error:
        _fail(f"cannot start zstd for {path}: {error}")
    assert process.stdout is not None and process.stderr is not None
    total = 0
    try:
        for raw in process.stdout:
            total += len(raw)
            if total > max_bytes:
                _fail(f"decoded trace exceeds {max_bytes} bytes: {path}")
            line = raw[:-1] if raw.endswith(b"\n") else raw
            if b"\r" in line:
                _fail("carriage returns/CRLF are not supported in the trace")
            yield line
    except BaseException:
        # Parsing can fail while the decompressor is still writing.  Reap it
        # on every error, including GeneratorExit, so a malformed retained
        # trace cannot leave a child process behind.
        if process.poll() is None:
            try:
                process.kill()
            except OSError:
                pass
        process.stdout.close()
        process.stderr.close()
        process.wait()
        raise
    process.stdout.close()
    error = process.stderr.read().decode("utf-8", errors="replace")
    process.stderr.close()
    status = process.wait()
    _require(status == 0, f"zstd decode failed for {path}: {error.strip()}")


@dataclass(frozen=True)
class _Frame:
    module: str | None
    function: str | None
    file: str | None
    line: int | None
    inlined: bool

    def as_dict(self, index: int) -> dict[str, Any]:
        return {
            "index": index,
            "module": self.module,
            "function": self.function,
            "file": self.file,
            "line": self.line,
            "inlined": self.inlined,
        }


@dataclass(frozen=True)
class _InstructionPointer:
    address: int
    frames: tuple[_Frame, ...]


@dataclass(frozen=True)
class _Trace:
    instruction_id: int
    parent_id: int


@dataclass(frozen=True)
class _Allocation:
    size: int
    trace_id: int


def _metric() -> dict[str, int]:
    return {
        "allocation_events": 0,
        "requested_bytes": 0,
        "deallocation_events": 0,
        "deallocated_bytes": 0,
    }


def _add_metric(
    metric: dict[str, int],
    size: int,
    allocation_events: int,
    deallocation_events: int,
) -> None:
    metric["allocation_events"] += allocation_events
    metric["requested_bytes"] += size * allocation_events
    metric["deallocation_events"] += deallocation_events
    metric["deallocated_bytes"] += size * deallocation_events


def _owner_kind(function: str | None, owner: str) -> str | None:
    if function == owner:
        return "exact"
    if function is not None and function.startswith(owner):
        suffix = function[len(owner) :]
        if _RUST_HASH.fullmatch(suffix):
            return "rust_hash"
    return None


def _phase_matches(function: str | None) -> tuple[str, ...]:
    """Return concrete phase markers present in one actual symbol.

    The predicates intentionally require source-level words.  They do not
    infer a phase from a case label, allocation size, or the absence of a
    different marker.
    """

    if not function:
        return ()
    lowered = function.lower()
    matches: list[str] = []
    if (
        "preservation_index" in lowered
        or "preservationindex::new" in lowered
        or "preservationindex::parse" in lowered
        or "preservationindex::parser" in lowered
        or (
            "preservationindex" in lowered
            and re.search(
                r"preservationindex(?:\$lt\$[^:]*\$gt\$)?::(?:new|parse|parser)",
                lowered,
            )
            is not None
        )
    ):
        matches.append("preservation_index")
    # This is the concrete source-backed topology boundary.  A surrounding
    # ``write_topology_to_stream`` frame is deliberately not enough evidence
    # for either this bucket or emission.
    if (
        "prepare_topology_state" in lowered
        or "::prepare_topology" in lowered
        or "topology_preparation" in lowered
    ):
        matches.append("topology_preparation")
    # These are the concrete low-level preparation/validation symbols.  A
    # generic action-plan name or write_changed_overlays_with_appended_inner
    # does not enter this bucket without one of these deeper frames.
    if (
        "::prepare_with" in lowered
        or "::validate_output_layout" in lowered
    ):
        matches.append("plan")
    # Keep emission tied to concrete prepared-output symbols.  In particular,
    # broad ``write``/``emit`` substrings would misclassify the topology
    # wrapper and are therefore intentionally absent.
    if "::write_prepared" in lowered or "::generated_entry" in lowered:
        matches.append("emission")
    return tuple(matches)


def _parse_lines(
    lines: Iterable[bytes],
    *,
    owner: str,
    max_records: int,
    max_depth: int,
    max_stack_frames: int,
) -> dict[str, Any]:
    _require(isinstance(owner, str) and owner, "owner must be a non-empty string")
    _require(max_records > 0, "max_records must be positive")
    _require(max_depth > 0, "max_depth must be positive")
    _require(max_stack_frames > 0, "max_stack_frames must be positive")

    strings: dict[int, str] = {}
    instruction_pointers: dict[int, _InstructionPointer] = {}
    traces: dict[int, _Trace] = {}
    allocations: list[_Allocation] = []
    plus: Counter[int] = Counter()
    minus: Counter[int] = Counter()
    records: dict[str, int] = {
        "lines": 0,
        "comments": 0,
        "version": 0,
        "strings": 0,
        "instruction_pointers": 0,
        "traces": 0,
        "allocation_descriptors": 0,
        "allocation_events": 0,
        "deallocation_events": 0,
        "timestamps": 0,
        "rss_events": 0,
        "suppression_records": 0,
        "debuggee_records": 0,
        "attached_markers": 0,
    }
    header: tuple[int, int] | None = None
    timestamp: int | None = None

    for record_number, raw in enumerate(lines, 1):
        records["lines"] += 1
        if record_number > max_records:
            _fail(f"trace has more than {max_records} records")
        if not isinstance(raw, bytes):
            _fail(f"trace line {record_number} is not bytes")
        line = raw[:-1] if raw.endswith(b"\n") else raw
        if b"\r" in line:
            _fail(f"carriage return at trace record {record_number}")
        if not line or line.startswith(b"#"):
            records["comments"] += 1
            continue
        mode = line[:1]
        if header is None and mode != b"v":
            _fail(f"version header must precede records at record {record_number}")

        if mode == b"v":
            if header is not None:
                _fail(f"duplicate version header at record {record_number}")
            version, file_version = _two_hex(
                line, b"v", "Heaptrack version", "file version", record_number
            )
            if version != HEAPTRACK_VERSION or file_version != FILE_VERSION:
                _fail(
                    f"unsupported Heaptrack version at record {record_number}: "
                    f"{version:x} {file_version:x}"
                )
            header = (version, file_version)
            records["version"] += 1

        elif mode == b"s":
            strings[len(strings) + 1] = _sized_string(line, record_number)
            records["strings"] += 1

        elif mode == b"i":
            tokens = _tokens(line, b"i", record_number)
            # address + module, followed by either a function ID or one
            # function/file/line tuple and zero or more inline triples.
            if len(tokens) < 2 or (len(tokens) > 3 and (len(tokens) - 2) % 3):
                _fail(f"malformed instruction frames at record {record_number}")
            address = _hex(tokens[0], "instruction address", record_number)
            module_id = _hex(tokens[1], "module string ID", record_number)
            module = strings.get(module_id) if module_id else None
            if module_id and module is None:
                _fail(f"unknown module string {module_id} at record {record_number}")
            frames: list[_Frame] = []
            if len(tokens) == 2:
                frames.append(_Frame(module, None, None, None, False))
            elif len(tokens) == 3:
                function_id = _hex(tokens[2], "function string ID", record_number)
                function = strings.get(function_id) if function_id else None
                if function_id and function is None:
                    _fail(f"unknown function string {function_id} at record {record_number}")
                frames.append(_Frame(module, function, None, None, False))
            else:
                for offset in range(2, len(tokens), 3):
                    function_id = _hex(tokens[offset], "function string ID", record_number)
                    file_id = _hex(tokens[offset + 1], "file string ID", record_number)
                    line_number = _hex(tokens[offset + 2], "source line", record_number)
                    function = strings.get(function_id) if function_id else None
                    file = strings.get(file_id) if file_id else None
                    if function_id and function is None:
                        _fail(f"unknown function string {function_id} at record {record_number}")
                    if file_id and file is None:
                        _fail(f"unknown file string {file_id} at record {record_number}")
                    frames.append(
                        _Frame(module, function, file, line_number, offset > 2)
                    )
            instruction_id = len(instruction_pointers) + 1
            instruction_pointers[instruction_id] = _InstructionPointer(address, tuple(frames))
            records["instruction_pointers"] += 1

        elif mode == b"t":
            instruction_id, parent_id = _two_hex(
                line,
                b"t",
                "instruction pointer ID",
                "trace parent ID",
                record_number,
            )
            if instruction_id and instruction_id not in instruction_pointers:
                _fail(f"unknown instruction pointer {instruction_id} at record {record_number}")
            trace_id = len(traces) + 1
            if parent_id >= trace_id:
                _fail(f"forward or cyclic trace parent {parent_id} at record {record_number}")
            traces[trace_id] = _Trace(instruction_id, parent_id)
            records["traces"] += 1

        elif mode == b"a":
            size, trace_id = _two_hex(
                line, b"a", "allocation size", "trace ID", record_number
            )
            if trace_id and trace_id not in traces:
                _fail(f"unknown allocation trace {trace_id} at record {record_number}")
            allocations.append(_Allocation(size, trace_id))
            records["allocation_descriptors"] += 1

        elif mode in (b"+", b"-"):
            label = "allocation descriptor ID" if mode == b"+" else "deallocation descriptor ID"
            index = _one_hex(line, mode, label, record_number)
            if index >= len(allocations):
                _fail(
                    f"unknown {mode.decode()} allocation descriptor {index} "
                    f"at record {record_number}"
                )
            target = plus if mode == b"+" else minus
            target[index] += 1
            records["allocation_events" if mode == b"+" else "deallocation_events"] += 1

        elif mode == b"c":
            value = _one_hex(line, b"c", "timestamp", record_number)
            if timestamp is not None and value < timestamp:
                _fail(f"timestamp moved backwards at record {record_number}")
            timestamp = value
            records["timestamps"] += 1

        elif mode == b"R":
            _one_hex(line, b"R", "RSS value", record_number)
            records["rss_events"] += 1

        elif mode == b"X":
            if records["debuggee_records"]:
                _fail(f"duplicate debuggee record at record {record_number}")
            if len(line) < 3 or line[1:2] != b" " or not line[2:]:
                _fail(f"malformed debuggee record at record {record_number}")
            records["debuggee_records"] += 1

        elif mode == b"I":
            _two_hex(line, b"I", "page size", "page count", record_number)

        elif mode == b"S":
            if len(line) < 3 or line[1:2] != b" " or not line[2:]:
                _fail(f"malformed suppression record at record {record_number}")
            records["suppression_records"] += 1

        elif mode == b"A":
            if line != b"A":
                _fail(f"malformed attached marker at record {record_number}")
            records["attached_markers"] += 1

        else:
            _fail(f"unknown Heaptrack record mode at record {record_number}: {mode!r}")

    if header is None:
        _fail("trace has no v 10500 3 header")

    trace_cache: dict[int, tuple[_Frame, ...]] = {0: ()}

    def frames_for_trace(trace_id: int) -> tuple[_Frame, ...]:
        cached = trace_cache.get(trace_id)
        if cached is not None:
            return cached
        current = trace_id
        seen: set[int] = set()
        pieces: list[tuple[_Frame, ...]] = []
        depth = 0
        while current:
            if current in seen:
                _fail(f"backtrace cycle at trace ID {current}")
            seen.add(current)
            depth += 1
            if depth > max_depth:
                _fail(f"backtrace exceeds {max_depth} trace nodes")
            node = traces.get(current)
            if node is None:
                _fail(f"backtrace references unknown trace ID {current}")
            instruction = instruction_pointers.get(node.instruction_id)
            if instruction is None and node.instruction_id:
                _fail(f"backtrace references unknown instruction ID {node.instruction_id}")
            if instruction is not None:
                pieces.append(instruction.frames)
            current = node.parent_id
        frames = tuple(frame for piece in pieces for frame in piece)
        if len(frames) > max_stack_frames:
            _fail(f"backtrace has more than {max_stack_frames} frames")
        trace_cache[trace_id] = frames
        return frames

    whole = _metric()
    owner_metric = _metric()
    rows: list[dict[str, Any]] = []
    groups: dict[str, dict[str, Any]] = {
        phase: {**_metric(), "allocation_info_indexes": [], "row_count": 0}
        for phase in PHASES
    }
    unknown_frames: dict[str, dict[str, int]] = {
        reason: {"frame_records": 0, "allocation_events": 0, "requested_bytes": 0}
        for reason in ("empty_trace", "missing_module", "missing_function", "missing_file")
    }
    diagnostics: dict[str, Any] = {
        "duplicate_owner_frames": {"allocation_events": 0, "requested_bytes": 0},
        "owner_hash_matches": {"allocation_events": 0, "requested_bytes": 0},
        "unmatched_allocation_events": 0,
        "unmatched_requested_bytes": 0,
        "trace_cache_entries": 1,
        "deallocation_events_kept_separate": records["deallocation_events"],
    }

    # Only descriptors referenced by + or - records become rows.  This keeps
    # unused allocator metadata out of the report while preserving every event
    # that contributes to either process total.
    referenced = sorted(set(plus) | set(minus))
    for row_index, allocation_index in enumerate(referenced):
        allocation = allocations[allocation_index]
        allocation_events = plus.get(allocation_index, 0)
        deallocation_events = minus.get(allocation_index, 0)
        frames = frames_for_trace(allocation.trace_id)
        owner_matches: list[tuple[int, str]] = []
        for frame_index, frame in enumerate(frames):
            kind = _owner_kind(frame.function, owner)
            if kind is not None:
                owner_matches.append((frame_index, kind))
        matched = bool(owner_matches)
        innermost_owner = owner_matches[0][0] if owner_matches else None
        phase_evidence: list[tuple[int, str, str]] = []
        if innermost_owner is not None:
            for frame_index, frame in enumerate(frames):
                for phase in _phase_matches(frame.function):
                    phase_evidence.append((frame_index, phase, frame.function or ""))
        phase = phase_evidence[0][1] if phase_evidence else ("unknown" if matched else "unmatched")
        stack = [frame.as_dict(index) for index, frame in enumerate(frames)]
        row: dict[str, Any] = {
            "row_index": row_index,
            "allocation_info_index": allocation_index,
            "size": allocation.size,
            "trace_id": allocation.trace_id,
            "allocation_events": allocation_events,
            "requested_bytes": allocation.size * allocation_events,
            "deallocation_events": deallocation_events,
            "deallocated_bytes": allocation.size * deallocation_events,
            "frames_leaf_to_root": stack,
            "owner_match": matched,
            "owner_match_count": len(owner_matches),
            "owner_match_kinds": [kind for _index, kind in owner_matches],
            "owner_frame_indexes": [index for index, _kind in owner_matches],
            "innermost_owner_frame_index": innermost_owner,
            "phase": phase,
            "phase_evidence": [
                {"frame_index": index, "phase": evidence_phase, "function": function}
                for index, evidence_phase, function in phase_evidence
            ],
        }
        rows.append(row)
        _add_metric(whole, allocation.size, allocation_events, deallocation_events)
        if matched:
            _add_metric(owner_metric, allocation.size, allocation_events, deallocation_events)
            group = groups[phase]
            _add_metric(group, allocation.size, allocation_events, deallocation_events)
            group["allocation_info_indexes"].append(allocation_index)
            group["row_count"] += 1
            if len(owner_matches) > 1:
                diagnostics["duplicate_owner_frames"]["allocation_events"] += allocation_events
                diagnostics["duplicate_owner_frames"]["requested_bytes"] += (
                    allocation.size * allocation_events
                )
            if any(kind == "rust_hash" for _index, kind in owner_matches):
                diagnostics["owner_hash_matches"]["allocation_events"] += allocation_events
                diagnostics["owner_hash_matches"]["requested_bytes"] += (
                    allocation.size * allocation_events
                )
        else:
            diagnostics["unmatched_allocation_events"] += allocation_events
            diagnostics["unmatched_requested_bytes"] += allocation.size * allocation_events

        reasons: set[str] = set()
        if not frames:
            reasons.add("empty_trace")
        for frame in frames:
            if frame.module is None:
                reasons.add("missing_module")
            if frame.function is None:
                reasons.add("missing_function")
            if frame.file is None:
                reasons.add("missing_file")
        for reason in reasons:
            info = unknown_frames[reason]
            info["allocation_events"] += allocation_events
            info["requested_bytes"] += allocation.size * allocation_events
        for frame in frames:
            if frame.module is None:
                unknown_frames["missing_module"]["frame_records"] += 1
            if frame.function is None:
                unknown_frames["missing_function"]["frame_records"] += 1
            if frame.file is None:
                unknown_frames["missing_file"]["frame_records"] += 1

    diagnostics["trace_cache_entries"] = len(trace_cache)
    row_whole = _metric()
    row_owner = _metric()
    for row in rows:
        _add_metric(
            row_whole,
            row["size"],
            row["allocation_events"],
            row["deallocation_events"],
        )
        if row["owner_match"]:
            _add_metric(
                row_owner,
                row["size"],
                row["allocation_events"],
                row["deallocation_events"],
            )
    _require(row_whole == whole, "allocation rows do not conserve whole-process totals")
    _require(row_owner == owner_metric, "allocation rows do not conserve owner totals")
    group_indexes: set[int] = set()
    for phase in PHASES:
        group = groups[phase]
        group_indexes.update(group["allocation_info_indexes"])
    _require(
        sum(groups[phase]["allocation_events"] for phase in PHASES)
        == owner_metric["allocation_events"],
        "phase groups do not conserve owner allocation events",
    )
    _require(
        sum(groups[phase]["requested_bytes"] for phase in PHASES)
        == owner_metric["requested_bytes"],
        "phase groups do not conserve owner requested bytes",
    )
    _require(
        sum(groups[phase]["deallocation_events"] for phase in PHASES)
        == owner_metric["deallocation_events"],
        "phase groups do not conserve owner deallocation events",
    )
    _require(
        sum(groups[phase]["deallocated_bytes"] for phase in PHASES)
        == owner_metric["deallocated_bytes"],
        "phase groups do not conserve owner deallocated bytes",
    )
    _require(
        len(group_indexes)
        == sum(groups[phase]["row_count"] for phase in PHASES)
        == sum(1 for row in rows if row["owner_match"]),
        "phase groups overlap or omit owner rows",
    )

    return {
        "schema": SCHEMA,
        "format": FORMAT,
        "header": {"heaptrack_version": HEAPTRACK_VERSION, "file_version": FILE_VERSION},
        "owner": {
            "symbol": owner,
            "match": "exact or ::h plus sixteen lower-case hexadecimal hash digits",
            "module_constraint": "none; function symbol is matched exactly",
        },
        "records": records,
        "trace_tables": {
            "instruction_pointers": len(instruction_pointers),
            "traces": len(traces),
            "allocation_descriptors": len(allocations),
            "event_referenced_descriptors": len(referenced),
        },
        "whole_process": whole,
        "owner_attribution": {
            **owner_metric,
            "rows": [row for row in rows if row["owner_match"]],
            "scope": "+ allocation events whose full leaf-to-root ancestry contains the owner",
            "deallocation_scope": (
                "- events referring to those matched allocation descriptors; these are "
                "lifetime frees and are not asserted to occur inside the timed call"
            ),
        },
        "allocation_rows": rows,
        "phase_groups": {
            "scope": "matched owner rows only",
            "classification": (
                "first concrete phase marker in the full leaf-to-root stack for "
                "a row with an innermost owner frame"
            ),
            "symbol_name_only": True,
            "not_measured_phases": True,
            "disjoint": True,
            "groups": groups,
        },
        "conservation": {
            "whole_process_rows": row_whole,
            "owner_rows": row_owner,
            "phase_groups_equal_owner": True,
            "allocation_events": {
                "whole_process": whole["allocation_events"],
                "matched_owner": owner_metric["allocation_events"],
                "unmatched": diagnostics["unmatched_allocation_events"],
            },
            "requested_bytes": {
                "whole_process": whole["requested_bytes"],
                "matched_owner": owner_metric["requested_bytes"],
                "unmatched": diagnostics["unmatched_requested_bytes"],
            },
        },
        "diagnostics": {
            **diagnostics,
            "unknown_frames": unknown_frames,
            "peak_live_bytes": "not measured",
            "physical_copy_bytes": "not measured",
            "elapsed_time": "not measured",
        },
    }


def parse(
    data: bytes | bytearray | memoryview,
    owner: str = PROFILE_OWNER,
    *,
    max_bytes: int = MAX_TRACE_BYTES,
    max_records: int = MAX_TRACE_RECORDS,
    max_depth: int = MAX_TRACE_DEPTH,
    max_stack_frames: int = MAX_STACK_FRAMES,
) -> dict[str, Any]:
    """Parse decoded interpreted Heaptrack bytes and return attribution rows."""

    _require(isinstance(data, (bytes, bytearray, memoryview)), "parse expects bytes")
    payload = bytes(data)
    _require(max_bytes > 0, "max_bytes must be positive")
    return _parse_lines(
        _iter_bytes_lines(payload, max_bytes),
        owner=owner,
        max_records=max_records,
        max_depth=max_depth,
        max_stack_frames=max_stack_frames,
    )


def parse_zst(
    path: str | Path,
    owner: str = PROFILE_OWNER,
    *,
    max_bytes: int = MAX_TRACE_BYTES,
    max_records: int = MAX_TRACE_RECORDS,
    max_depth: int = MAX_TRACE_DEPTH,
    max_stack_frames: int = MAX_STACK_FRAMES,
) -> dict[str, Any]:
    """Decode one retained ``.zst`` trace through zstd and parse it."""

    _require(max_bytes > 0, "max_bytes must be positive")
    return _parse_lines(
        _zstd_lines(Path(path), max_bytes),
        owner=owner,
        max_records=max_records,
        max_depth=max_depth,
        max_stack_frames=max_stack_frames,
    )


def parse_heap_trace(
    value: bytes | bytearray | memoryview | str | Path,
    owner: str = PROFILE_OWNER,
    **kwargs: Any,
) -> dict[str, Any]:
    """Compatibility entry point accepting decoded bytes or a ``.zst`` path."""

    if isinstance(value, (bytes, bytearray, memoryview)):
        return parse(value, owner, **kwargs)
    return parse_zst(value, owner, **kwargs)


def _main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("trace", type=Path, help="retained interpreted Heaptrack .zst trace")
    parser.add_argument("--owner", default=PROFILE_OWNER)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    report = parse_zst(args.trace, args.owner)
    encoded = json.dumps(report, indent=2, sort_keys=True) + "\n"
    if args.output is None:
        sys.stdout.write(encoded)
    else:
        args.output.write_text(encoded, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main(sys.argv[1:]))
