#!/usr/bin/env python3
"""Strict attribution reader for interpreted Heaptrack 1.5 format 3.

The reader consumes the uncompressed, interpreted text stream produced by
Heaptrack.  The format is deliberately parsed here instead of through
``heaptrack_print`` so allocation ``+`` records remain aligned with their
requested sizes and trace IDs.  String IDs are one based; instruction-pointer
and trace records are also one based, while allocation descriptors referenced
by ``+`` and ``-`` are zero based.  A trace is followed from its leaf node to
its parent nodes.  Direct and inline frames on one instruction pointer are
kept in that order.

Only the exact ``binary_path`` module string and ``owner`` function string are
eligible for owner attribution.  A Rust demangled hash suffix is accepted only
when it is exactly ``::h`` followed by sixteen lower-case hexadecimal digits;
this avoids broad substring or malformed-symbol matches.  A matching stack
must contain exactly one owner frame; duplicate owner frames remain in the
whole totals and an explicit ambiguous bucket instead of being silently
double-counted.

Supported record modes are the modes accepted by KDE's interpreted reader:
``v``, ``s``, ``i``, ``t``, ``a``, ``+``, ``-``, ``#``/empty, ``c``, ``R``,
``X``, ``A``, ``I``, and ``S``.  Raw ``x``/``m`` records and every unknown
mode are rejected.  This module intentionally makes no peak, lifetime, or
allocator-equivalence claim.

Format references (official KDE sources):

* https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/src/util/linereader.h
* https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/src/util/linewriter.h
* https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/src/analyze/accumulatedtracedata.cpp
* https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/src/interpret/heaptrack_interpret.cpp
* https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/CMakeLists.txt
"""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Iterable, NoReturn


SCHEMA = "litchi.heaptrack.0830.interpreted-v3.v1"
HEAPTRACK_VERSION = 0x10500
FILE_VERSION = 3
_HEX = re.compile(rb"^[0-9a-f]+$")
_RUST_HASH = re.compile(r"::h[0-9a-f]{16}")


class HeaptrackError(ValueError):
    """Raised when the stream cannot support checked attribution."""


def _fail(message: str) -> NoReturn:
    raise HeaptrackError(message)


def _hex(token: bytes, label: str) -> int:
    """Read the lower-case hexadecimal integer used by KDE's LineReader."""

    if not token or _HEX.fullmatch(token) is None or len(token) > 16:
        _fail(f"invalid {label} hexadecimal token: {token!r}")
    return int(token, 16)


def _tokens(line: bytes, mode: bytes) -> list[bytes]:
    """Parse a numeric record after its mandatory mode-space separator."""

    if len(line) < 2 or line[:1] != mode or line[1:2] != b" ":
        _fail(f"malformed {mode.decode()} record")
    rest = line[2:]
    if not rest:
        return []
    tokens = rest.split(b" ")
    if any(not token for token in tokens) or any(b"\t" in token for token in tokens):
        _fail(f"malformed spacing in {mode.decode()} record")
    return tokens


def _one_hex(line: bytes, mode: bytes, label: str) -> int:
    tokens = _tokens(line, mode)
    if len(tokens) != 1:
        _fail(f"{mode.decode()} record needs one field")
    return _hex(tokens[0], label)


def _two_hex(line: bytes, mode: bytes, first: str, second: str) -> tuple[int, int]:
    tokens = _tokens(line, mode)
    if len(tokens) != 2:
        _fail(f"{mode.decode()} record needs two fields")
    return _hex(tokens[0], first), _hex(tokens[1], second)


@dataclass(frozen=True)
class _Frame:
    module: str | None
    function: str | None
    file: str | None
    line: int | None
    inlined: bool

    def key(self) -> tuple[object, ...]:
        return (self.module, self.function, self.file, self.line, self.inlined)

    def as_dict(self) -> dict[str, object]:
        return {
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
    ip_id: int
    parent_id: int


@dataclass(frozen=True)
class _Allocation:
    size: int
    trace_id: int


def _string(strings: list[str], index: int, label: str) -> str | None:
    if index == 0:
        return None
    if index > len(strings):
        _fail(f"{label} string ID {index} is not defined")
    return strings[index - 1]


def _iter_lines(data: bytes) -> Iterable[bytes]:
    """Yield getline-compatible records while rejecting CRLF ambiguity."""

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
        if line.endswith(b"\r"):
            _fail("CRLF or carriage return is not a supported stream encoding")
        yield line


def _sized_string(line: bytes) -> str:
    if len(line) < 3 or line[:2] != b"s ":
        _fail("malformed sized string record")
    rest = line[2:]
    separator = rest.find(b" ")
    if separator <= 0:
        _fail("sized string lacks a hexadecimal length separator")
    size = _hex(rest[:separator], "string byte length")
    payload = rest[separator + 1 :]
    if len(payload) != size:
        _fail(
            f"sized string length {size} does not match {len(payload)} payload bytes"
        )
    try:
        return payload.decode("utf-8", errors="strict")
    except UnicodeDecodeError as error:
        _fail(f"sized string is not valid UTF-8: {error}")


def _owner_kind(function: str | None, owner: str) -> str | None:
    if function == owner:
        return "exact"
    if function is not None and function.startswith(owner) and _RUST_HASH.fullmatch(function[len(owner) :]):
        return "rust_hash"
    return None


def _metric() -> dict[str, int]:
    return {"calls": 0, "requested_bytes": 0, "owner_calls": 0, "owner_requested_bytes": 0}


def _add_metric(metric: dict[str, int], size: int, owner: bool) -> None:
    metric["calls"] += 1
    metric["requested_bytes"] += size
    if owner:
        metric["owner_calls"] += 1
        metric["owner_requested_bytes"] += size


def parse(data: bytes, owner: str, binary_path: str) -> dict[str, object]:
    """Parse *data* and return deterministic whole/owner allocation metrics.

    ``whole`` counts every ``+`` event and its descriptor's requested size.
    ``owner`` counts an event once when any direct or inline frame in its
    complete leaf-to-root trace has exactly one exact module/function match.
    Duplicate matches are retained in the ambiguous diagnostic bucket.  The
    ``full_leaf`` rows partition events by trace ID; ``leaf`` rows aggregate
    events that have a first frame.  An allocation with trace ID zero has no
    leaf row but remains in the whole and full-trace totals.  ``caller_trace``
    rows are inclusive frame projections and are explicitly marked
    non-additive.
    """

    if not isinstance(data, bytes):
        _fail("parse expects bytes")
    if not isinstance(owner, str) or not owner:
        _fail("owner must be a non-empty string")
    if not isinstance(binary_path, str) or not binary_path:
        _fail("binary_path must be a non-empty string")

    strings: list[str] = []
    ips: list[_InstructionPointer] = []
    traces: list[_Trace] = []
    allocations: list[_Allocation] = []
    events: list[tuple[int, int]] = []
    records = {
        "lines": 0,
        "comments": 0,
        "strings": 0,
        "instruction_pointers": 0,
        "traces": 0,
        "allocation_descriptors": 0,
        "allocation_events": 0,
        "deallocation_events": 0,
        "timestamps": 0,
        "rss_events": 0,
        "suppression_records": 0,
    }
    header: tuple[int, int] | None = None
    debuggee_seen = False
    timestamp: int | None = None
    saw_record = False

    for line_number, line in enumerate(_iter_lines(data), 1):
        records["lines"] += 1
        mode = line[:1]
        if mode in (b"", b"#"):
            records["comments"] += 1
            continue
        if header is None and mode != b"v":
            _fail(f"line {line_number}: version header must precede records")
        saw_record = True

        if mode == b"v":
            if header is not None:
                _fail(f"line {line_number}: duplicate version header")
            version, file_version = _two_hex(line, b"v", "Heaptrack version", "file version")
            if version != HEAPTRACK_VERSION or file_version != FILE_VERSION:
                _fail(
                    f"line {line_number}: unsupported header v {version:x} {file_version:x}"
                )
            header = (version, file_version)

        elif mode == b"s":
            strings.append(_sized_string(line))
            records["strings"] += 1

        elif mode == b"i":
            tokens = _tokens(line, b"i")
            if len(tokens) < 2 or (len(tokens) > 3 and (len(tokens) - 2) % 3):
                _fail(f"line {line_number}: i record needs 2, 3, or 5+3n fields")
            address = _hex(tokens[0], "instruction pointer")
            module_id = _hex(tokens[1], "module string ID")
            module = _string(strings, module_id, "module")
            frames: list[_Frame] = []
            if len(tokens) == 2:
                # KDE's default InstructionPointer has no symbol/file IDs,
                # but it is still a real unknown leaf in the trace.
                frames.append(_Frame(module, None, None, None, False))
            elif len(tokens) == 3:
                function = _string(strings, _hex(tokens[2], "function string ID"), "function")
                frames.append(_Frame(module, function, None, None, False))
            else:
                for offset in range(2, len(tokens), 3):
                    function = _string(strings, _hex(tokens[offset], "function string ID"), "function")
                    file = _string(strings, _hex(tokens[offset + 1], "file string ID"), "file")
                    line_number_value = _hex(tokens[offset + 2], "source line")
                    frames.append(_Frame(module, function, file, line_number_value, offset > 2))
            ips.append(_InstructionPointer(address, tuple(frames)))
            records["instruction_pointers"] += 1

        elif mode == b"t":
            ip_id, parent_id = _two_hex(line, b"t", "instruction pointer ID", "parent trace ID")
            if not 1 <= ip_id <= len(ips):
                _fail(f"line {line_number}: trace references missing instruction pointer {ip_id}")
            if parent_id > len(traces):
                _fail(f"line {line_number}: trace references forward/missing parent {parent_id}")
            trace = _Trace(ip_id, parent_id)
            traces.append(trace)
            if trace.parent_id == len(traces):
                _fail(f"line {line_number}: trace has a self cycle")
            records["traces"] += 1

        elif mode == b"a":
            size, trace_id = _two_hex(line, b"a", "allocation size", "trace ID")
            if trace_id > len(traces):
                _fail(f"line {line_number}: allocation references missing trace {trace_id}")
            allocations.append(_Allocation(size, trace_id))
            records["allocation_descriptors"] += 1

        elif mode == b"+":
            allocation_id = _one_hex(line, b"+", "allocation descriptor ID")
            if allocation_id >= len(allocations):
                _fail(f"line {line_number}: + references missing allocation descriptor {allocation_id}")
            events.append((allocation_id, line_number))
            records["allocation_events"] += 1

        elif mode == b"-":
            allocation_id = _one_hex(line, b"-", "deallocation descriptor ID")
            if allocation_id >= len(allocations):
                _fail(f"line {line_number}: - references missing allocation descriptor {allocation_id}")
            records["deallocation_events"] += 1

        elif mode == b"c":
            value = _one_hex(line, b"c", "timestamp")
            if timestamp is not None and value < timestamp:
                _fail(f"line {line_number}: timestamp moved backwards")
            timestamp = value
            records["timestamps"] += 1

        elif mode == b"R":
            _one_hex(line, b"R", "RSS value")
            records["rss_events"] += 1

        elif mode == b"X":
            if debuggee_seen:
                _fail(f"line {line_number}: duplicate debuggee record")
            if len(line) < 2 or line[1:2] != b" ":
                _fail(f"line {line_number}: malformed debuggee record")
            debuggee_seen = True

        elif mode == b"I":
            _two_hex(line, b"I", "page size", "page count")

        elif mode == b"S":
            if len(line) < 2 or line[1:2] != b" ":
                _fail(f"line {line_number}: malformed suppression record")
            records["suppression_records"] += 1

        elif mode == b"A":
            if line != b"A":
                _fail(f"line {line_number}: attached marker has unexpected fields")

        else:
            _fail(f"line {line_number}: unsupported record mode {mode!r}")

    if header is None or not saw_record:
        _fail("stream has no v 10500 3 header")

    # Resolve traces iteratively.  The parent index is checked during parsing,
    # but retaining a cycle guard makes the invariant explicit and avoids
    # recursion depth failures on a very deep valid profile.
    trace_cache: dict[int, tuple[_Frame, ...]] = {0: ()}

    def trace_frames(trace_id: int) -> tuple[_Frame, ...]:
        cached = trace_cache.get(trace_id)
        if cached is not None:
            return cached
        current = trace_id
        seen: set[int] = set()
        pieces: list[tuple[_Frame, ...]] = []
        while current:
            if current in seen:
                _fail(f"trace cycle at trace ID {current}")
            seen.add(current)
            if not 1 <= current <= len(traces):
                _fail(f"trace references missing trace ID {current}")
            node = traces[current - 1]
            if not 1 <= node.ip_id <= len(ips):
                _fail(f"trace references missing instruction pointer {node.ip_id}")
            pieces.append(ips[node.ip_id - 1].frames)
            current = node.parent_id
        result: tuple[_Frame, ...] = tuple(frame for piece in pieces for frame in piece)
        trace_cache[trace_id] = result
        return result

    whole = _metric()
    owner_metric = _metric()
    size_buckets: dict[int, dict[str, int]] = {}
    leaf_buckets: dict[tuple[object, ...], dict[str, int]] = {}
    full_leaf_buckets: dict[int, dict[str, int]] = {}
    caller_buckets: dict[tuple[object, ...], dict[str, int]] = {}
    unknown_frame_records = {"missing_module": 0, "missing_function": 0, "missing_file": 0, "empty_trace": 0}
    unknown_frame_calls = {key: 0 for key in unknown_frame_records}
    unknown_frame_bytes = {key: 0 for key in unknown_frame_records}
    duplicate_owner_calls = 0
    duplicate_owner_bytes = 0
    ambiguous_owner_calls = 0
    ambiguous_owner_bytes = 0
    wrong_dso_calls = 0
    wrong_dso_bytes = 0
    hash_owner_calls = 0
    hash_owner_bytes = 0

    for allocation_id, _line_number in events:
        allocation = allocations[allocation_id]
        frames = trace_frames(allocation.trace_id)
        exact_matches = [
            _owner_kind(frame.function, owner)
            for frame in frames
            if frame.module == binary_path
        ]
        exact_matches = [kind for kind in exact_matches if kind is not None]
        wrong_matches = [
            _owner_kind(frame.function, owner)
            for frame in frames
            if frame.module != binary_path
        ]
        wrong_matches = [kind for kind in wrong_matches if kind is not None]
        is_owner = len(exact_matches) == 1
        if len(exact_matches) > 1:
            duplicate_owner_calls += 1
            duplicate_owner_bytes += allocation.size
            ambiguous_owner_calls += 1
            ambiguous_owner_bytes += allocation.size
        if wrong_matches:
            wrong_dso_calls += 1
            wrong_dso_bytes += allocation.size
        if "rust_hash" in exact_matches:
            hash_owner_calls += 1
            hash_owner_bytes += allocation.size

        _add_metric(whole, allocation.size, is_owner)
        if is_owner:
            _add_metric(owner_metric, allocation.size, True)
        size_metric = size_buckets.setdefault(allocation.size, _metric())
        _add_metric(size_metric, allocation.size, is_owner)
        trace_metric = full_leaf_buckets.setdefault(allocation.trace_id, _metric())
        _add_metric(trace_metric, allocation.size, is_owner)

        if frames:
            leaf_metric = leaf_buckets.setdefault(frames[0].key(), _metric())
            _add_metric(leaf_metric, allocation.size, is_owner)
            seen_keys: set[tuple[object, ...]] = set()
            for frame in frames:
                key = frame.key()
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                caller_metric = caller_buckets.setdefault(key, _metric())
                _add_metric(caller_metric, allocation.size, is_owner)

        reasons: set[str] = set()
        if not frames:
            reasons.add("empty_trace")
        for frame in frames:
            if frame.module is None:
                reasons.add("missing_module")
                unknown_frame_records["missing_module"] += 1
            if frame.function is None:
                reasons.add("missing_function")
                unknown_frame_records["missing_function"] += 1
            if frame.file is None:
                reasons.add("missing_file")
                unknown_frame_records["missing_file"] += 1
        for reason in reasons:
            unknown_frame_calls[reason] += 1
            unknown_frame_bytes[reason] += allocation.size

    for key in unknown_frame_calls:
        unknown_frame_records[key] = {
            "frame_records": unknown_frame_records[key],
            "allocation_calls": unknown_frame_calls[key],
            "requested_bytes": unknown_frame_bytes[key],
        }

    def frame_rows(buckets: dict[tuple[object, ...], dict[str, int]]) -> list[dict[str, object]]:
        rows: list[dict[str, object]] = []
        for key, metric in sorted(
            buckets.items(),
            key=lambda item: tuple("" if value is None else str(value) for value in item[0]),
        ):
            row = {
                "module": key[0],
                "function": key[1],
                "file": key[2],
                "line": key[3],
                "inlined": key[4],
            }
            row.update(metric)
            rows.append(row)
        return rows

    full_leaf_rows: list[dict[str, object]] = []
    for trace_id, metric in sorted(full_leaf_buckets.items()):
        row: dict[str, object] = {
            "trace_id": trace_id,
            "frames_leaf_to_root": [frame.as_dict() for frame in trace_frames(trace_id)],
        }
        row.update(metric)
        full_leaf_rows.append(row)

    size_rows = {
        str(size): {"size": size, **metric}
        for size, metric in sorted(size_buckets.items())
    }
    result = {
        "schema": SCHEMA,
        "format": "interpreted_heaptrack_v3",
        "header": {"heaptrack_version": HEAPTRACK_VERSION, "file_version": FILE_VERSION},
        "owner": owner,
        "binary_path": binary_path,
        "records": records,
        "whole": whole,
        "owner_attribution": owner_metric,
        "by_allocation_size": size_rows,
        "leaf": {
            "inclusive": False,
            "non_additive": False,
            "rows": frame_rows(leaf_buckets),
        },
        "full_leaf": {
            "inclusive": False,
            "non_additive": False,
            "rows": full_leaf_rows,
        },
        "caller_trace": {
            "inclusive": True,
            "non_additive": True,
            "rows": frame_rows(caller_buckets),
        },
        "diagnostics": {
            "unknown_frames": unknown_frame_records,
            "duplicate_owner_frames": {
                "allocation_calls": duplicate_owner_calls,
                "requested_bytes": duplicate_owner_bytes,
            },
            "ambiguous_owner": {
                "allocation_calls": ambiguous_owner_calls,
                "requested_bytes": ambiguous_owner_bytes,
            },
            "owner_symbol_in_wrong_dso": {
                "allocation_calls": wrong_dso_calls,
                "requested_bytes": wrong_dso_bytes,
            },
            "rust_hash_owner_matches": {
                "allocation_calls": hash_owner_calls,
                "requested_bytes": hash_owner_bytes,
            },
            "trace_cache_entries": len(trace_cache),
            "deallocation_events_ignored_for_attribution": records["deallocation_events"],
            "peak_and_lifetime_metrics": "not measured",
        },
    }
    # Keep the two primary projections easy to bind from shell/JSON drivers;
    # the nested objects above remain the canonical rows used by this reader.
    result.update(
        whole_calls=whole["calls"],
        whole_requested_bytes=whole["requested_bytes"],
        owner_calls=owner_metric["calls"],
        owner_requested_bytes=owner_metric["requested_bytes"],
    )
    return result
