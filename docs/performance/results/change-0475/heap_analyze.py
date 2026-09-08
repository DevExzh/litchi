#!/usr/bin/env python3
"""Parse interpreted Heaptrack v3 traces without inventing phase evidence.

Heaptrack's interpreted stream is an event log.  ``a`` records intern an
allocation's requested size and trace id, ``+`` and ``-`` mark allocation
lifetimes, and ``i``/``t``/``s`` resolve the trace.  This reader keeps those
relationships intact while it aggregates calls, requested bytes, and the
maximum live byte total in event order.

The parser is intentionally independent of ``heaptrack_print`` and uses only
the Python standard library.  It accepts a decoded plain stream or a gzip
stream.  Heaptrack's zstd input is expected to be exported to gzip text by the
capture driver; silently treating an unsupported format as text would make a
phase attribution unsafe.
"""

from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import dataclass, field
import gzip
import hashlib
import json
from pathlib import Path
import re
import sys
from typing import BinaryIO, Iterator, NoReturn, Sequence


SCHEMA = "litchi-0475-heap-attribution-v1"
SUPPORTED_FILE_VERSION = 3
HEX_RE = re.compile(rb"^[0-9a-fA-F]+$")


class AnalysisError(ValueError):
    """Raised when a trace cannot support checked attribution."""


def fail(message: str) -> NoReturn:
    raise AnalysisError(message)


@dataclass(frozen=True)
class Frame:
    """One instruction pointer's direct or inlined source frame."""

    function: str | None
    file: str | None
    line: int | None
    module: str | None
    inlined: bool = False
    complete: bool = True

    def display(self) -> str:
        name = self.function or "[unknown]"
        if self.module:
            return f"{name} in {self.module}"
        return name


@dataclass(frozen=True)
class InstructionPointer:
    module_id: int
    frames: tuple[Frame, ...]


@dataclass(frozen=True)
class TraceNode:
    ip_id: int
    parent_id: int


@dataclass(frozen=True)
class AllocationInfo:
    size: int
    trace_id: int


@dataclass
class Aggregate:
    """Counters for one phase/category/stack projection."""

    calls: int = 0
    deallocations: int = 0
    requested_bytes: int = 0
    deallocated_bytes: int = 0
    live_bytes: int = 0
    peak_live_bytes: int = 0
    peak_event: int | None = None
    outstanding_allocations: int = 0

    def allocate(self, size: int, event: int) -> None:
        self.calls = checked_add(self.calls, 1, "allocation calls")
        self.requested_bytes = checked_add(
            self.requested_bytes, size, "requested bytes"
        )
        self.live_bytes = checked_add(self.live_bytes, size, "live bytes")
        self.outstanding_allocations = checked_add(
            self.outstanding_allocations, 1, "outstanding allocations"
        )
        if self.live_bytes > self.peak_live_bytes:
            self.peak_live_bytes = self.live_bytes
            self.peak_event = event

    def deallocate(self, size: int) -> None:
        if self.live_bytes < size or self.outstanding_allocations <= 0:
            fail("deallocation exceeds its aligned allocation lifetime")
        self.deallocations = checked_add(self.deallocations, 1, "deallocation calls")
        self.deallocated_bytes = checked_add(
            self.deallocated_bytes, size, "deallocated bytes"
        )
        self.live_bytes -= size
        self.outstanding_allocations -= 1

    def as_dict(self) -> dict[str, int | None]:
        return {
            "allocation_calls": self.calls,
            "deallocation_calls": self.deallocations,
            "requested_bytes": self.requested_bytes,
            "deallocated_bytes": self.deallocated_bytes,
            "live_bytes_at_end": self.live_bytes,
            "peak_live_bytes": self.peak_live_bytes,
            "peak_event_ordinal": self.peak_event,
            "outstanding_allocations": self.outstanding_allocations,
        }


@dataclass(frozen=True)
class StackKey:
    phase: str
    category: str
    trace_id: int


@dataclass
class ParseStats:
    lines: int = 0
    comments: int = 0
    strings: int = 0
    instruction_pointers: int = 0
    traces: int = 0
    allocation_descriptors: int = 0
    allocation_events: int = 0
    deallocation_events: int = 0
    timestamps: int = 0
    rss_events: int = 0
    unknown_events: int = 0
    max_timestamp: int | None = None
    peak_rss: int | None = None


@dataclass
class TraceResult:
    lane: str
    path: str
    compressed_sha256: str
    compressed_bytes: int
    file_version: int
    heaptrack_version: int
    command: str | None
    stats: ParseStats
    phase: dict[str, Aggregate]
    category: dict[str, Aggregate]
    stacks: dict[StackKey, Aggregate]
    stack_names: dict[int, tuple[str, ...]]
    unresolved_trace_ids: set[int]
    unclassified_phase_events: int
    unknown_scope_writer_helper_events: int
    all_peak_live_bytes: int
    all_peak_event: int | None
    final_live_bytes: int
    final_outstanding_allocations: int
    limitations: list[str]
    unmeasured_phases: dict[str, str] = field(default_factory=dict)
    event_scope: str = "all_interpreted_plus_minus_records"
    complete_event_scan: bool = True
    fallback: dict[str, object] | None = None

    def as_dict(self) -> dict[str, object]:
        categories = {
            key: value.as_dict() for key, value in sorted(self.category.items())
        }
        phases = {key: value.as_dict() for key, value in sorted(self.phase.items())}
        stack_rows = []
        for key, value in sorted(
            self.stacks.items(), key=lambda item: (-item[1].requested_bytes, item[0].trace_id)
        ):
            stack_rows.append(
                {
                    "phase": key.phase,
                    "category": key.category,
                    "trace_id": key.trace_id,
                    "frames_leaf_to_root": list(self.stack_names.get(key.trace_id, ())),
                    **value.as_dict(),
                }
            )
        result: dict[str, object] = {
            "lane": self.lane,
            "trace": {
                "path": self.path,
                "compressed_sha256": self.compressed_sha256,
                "compressed_bytes": self.compressed_bytes,
                "format": "interpreted_heaptrack_v3",
                "heaptrack_version_hex": f"{self.heaptrack_version:x}",
                "file_version": self.file_version,
            },
            "command": self.command,
            "records": {
                "lines": self.stats.lines,
                "comments": self.stats.comments,
                "strings": self.stats.strings,
                "instruction_pointers": self.stats.instruction_pointers,
                "traces": self.stats.traces,
                "allocation_descriptors": self.stats.allocation_descriptors,
                "allocation_events": self.stats.allocation_events,
                "deallocation_events": self.stats.deallocation_events,
                "timestamps": self.stats.timestamps,
                "rss_events": self.stats.rss_events,
                "max_timestamp": self.stats.max_timestamp,
                "peak_rss_bytes": self.stats.peak_rss,
                "event_scope": self.event_scope,
                "complete_event_scan": self.complete_event_scan,
            },
            "timeline": {
                "basis": (
                    "interpreted Heaptrack plus/minus event order"
                    if self.complete_event_scan
                    else "filtered interpreted Heaptrack plus/minus event order for exact phase-matching descriptors"
                ),
                "peak_live_bytes": self.all_peak_live_bytes,
                "peak_event_ordinal": self.all_peak_event,
                "live_bytes_at_end": self.final_live_bytes,
                "outstanding_allocations_at_end": self.final_outstanding_allocations,
            },
            "phase_attribution": phases,
            "phase_scope": {
                key: {"status": "unavailable", "reason": reason}
                for key, reason in sorted(self.unmeasured_phases.items())
            },
            "category_attribution": categories,
            "stacks": stack_rows,
            "scope": {
                "primary_phase_allocation_events": (
                    self.stats.allocation_events - self.unclassified_phase_events
                    if self.complete_event_scan
                    else self.stats.allocation_events
                ),
                "unknown_scope_allocation_events": (
                    self.unclassified_phase_events if self.complete_event_scan else None
                ),
                "unknown_scope_writer_helper_records": self.unknown_scope_writer_helper_events,
                "whole_process": {
                    "status": "measured" if self.complete_event_scan else "unavailable",
                    "allocation_calls": self.stats.allocation_events
                    if self.complete_event_scan
                    else None,
                    "requested_bytes": sum(
                        aggregate.requested_bytes for aggregate in self.stacks.values()
                    )
                    if self.complete_event_scan
                    else None,
                    "reason": None
                    if self.complete_event_scan
                    else "fast scoped scan intentionally skipped non-matching plus/minus records",
                },
                "unresolved_trace_ids": sorted(self.unresolved_trace_ids),
                "classification": "phase requires exact pptx_streaming_create::build_corpus or ::run ancestry (Rust v0 mangled forms accepted); category uses the nearest visible direct/inlined frame, leaf first; totals are exact only for matched trace metadata",
            },
            "limitations": list(self.limitations),
        }
        if self.fallback is not None:
            result["heaptrack_print_fallback"] = self.fallback
        return result


def checked_add(left: int, right: int, label: str) -> int:
    if left < 0 or right < 0:
        fail(f"{label} cannot be negative")
    value = left + right
    if value > (1 << 64) - 1:
        fail(f"{label} overflows u64")
    return value


def parse_hex(token: bytes, label: str) -> int:
    if not token or HEX_RE.fullmatch(token) is None:
        fail(f"{label} is not a lowercase hexadecimal integer")
    try:
        return int(token, 16)
    except ValueError as error:  # pragma: no cover - regex guards this
        raise AnalysisError(f"{label} is not hexadecimal") from error


def parse_fields(line: bytes, expected: int, mode: bytes) -> list[bytes]:
    body = line[2:].rstrip(b"\n")
    fields = body.split(b" ")
    if any(field == b"" for field in fields):
        fail(f"{mode.decode()} record contains repeated spaces")
    if len(fields) != expected:
        fail(
            f"{mode.decode()} record has {len(fields)} fields; expected {expected}"
        )
    return fields


def decode_string(value: bytes, label: str) -> str:
    try:
        return value.decode("utf-8")
    except UnicodeDecodeError as error:
        raise AnalysisError(f"{label} is not UTF-8") from error


def parse_sized_string(body: bytes, label: str) -> str:
    separator = body.find(b" ")
    if separator <= 0:
        fail(f"{label} has no size separator")
    size = parse_hex(body[:separator], f"{label} size")
    payload = body[separator + 1 :]
    if len(payload) != size:
        fail(f"{label} declares {size} bytes but contains {len(payload)}")
    return decode_string(payload, label)


def open_trace(path: Path) -> BinaryIO:
    if path.suffix.lower() == ".zst":
        fail(
            f"{path} is zstd-compressed; export the decoded interpreted stream to .txt.gz"
        )
    try:
        return gzip.open(path, "rb") if path.suffix.lower() == ".gz" else path.open("rb")
    except OSError as error:
        raise AnalysisError(f"cannot open {path}: {error}") from error


def file_binding(path: Path) -> tuple[str, int]:
    digest = hashlib.sha256()
    size = 0
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                size = checked_add(size, len(block), "compressed trace bytes")
    except OSError as error:
        raise AnalysisError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest(), size


def _string(strings: Sequence[str], index: int) -> str | None:
    return strings[index - 1] if 0 < index <= len(strings) else None


def _frame_name(frame: Frame) -> str:
    return frame.function or "[unknown]"


def trace_frames(
    trace_id: int,
    traces: Sequence[TraceNode],
    instruction_pointers: Sequence[InstructionPointer],
    strings: Sequence[str],
) -> tuple[tuple[str, ...], bool]:
    """Return leaf-to-root display names and whether resolution was complete."""

    if trace_id == 0:
        return (), True
    names: list[str] = []
    seen: set[int] = set()
    complete = True
    current = trace_id
    while current:
        if current in seen or current > len(traces):
            complete = False
            break
        seen.add(current)
        node = traces[current - 1]
        if node.ip_id <= 0 or node.ip_id > len(instruction_pointers):
            complete = False
            break
        ip = instruction_pointers[node.ip_id - 1]
        if not ip.frames:
            complete = False
            names.append("[unknown]")
        else:
            for frame in ip.frames:
                names.append(frame.display())
                if frame.function is None or not frame.complete:
                    complete = False
        current = node.parent_id
    return tuple(names), complete


PHASE_BUILD = "writer-under-build_corpus"
PHASE_RUN = "writer-under-run"
PHASE_OTHER = "other"
CATEGORY_WRITER = "writer-helper"
CATEGORY_OPC = "OPC-name-validation"
CATEGORY_ZIP = "ZIP-metadata"
CATEGORY_DEFLATE = "Deflate"
CATEGORY_OTHER = "other"


def _function_part(display_name: str) -> str:
    """Remove the module suffix added by :meth:`Frame.display`."""

    return display_name.split(" in ", 1)[0].lower()


def _has_streaming_function(names: Sequence[str], function: str) -> bool:
    """Match the actual streaming module function, including Rust v0 spelling.

    Interpreted traces commonly retain Rust v0 mangled names such as
    ``...21pptx_streaming_create3run``. Richer symbolization may expose
    ``...::pptx_streaming_create::run``. A broad module substring is
    intentionally insufficient: helper and verification frames stay outside
    the primary operation phase.
    """

    human = f"pptx_streaming_create::{function}"
    # Rust v0 length prefixes are decimal digits (the surrounding numeric
    # fields in the trace are hexadecimal, but symbol lengths are not).
    raw_suffix = f"pptx_streaming_create{len(function)}{function}"
    for name in names:
        part = _function_part(name)
        if part.endswith(human) or part.endswith(raw_suffix):
            return True
    return False


def classify_stack(names: Sequence[str]) -> tuple[str, str]:
    """Classify only by visible stack evidence, with build phase precedence."""

    if _has_streaming_function(names, "build_corpus"):
        phase = PHASE_BUILD
    elif _has_streaming_function(names, "run"):
        phase = PHASE_RUN
    else:
        phase = PHASE_OTHER

    # Categories use the nearest visible matching frame (the stack is stored
    # leaf first).  Looking at the whole stack with writer precedence would
    # swallow nested OPC/ZIP/Deflate work because every such call is reached
    # from the public writer helper.
    category = CATEGORY_OTHER
    for name in names:
        lowered = name.lower()
        if any(
            needle in lowered
            for needle in ("packuri", "partname", "part_name", "membername", "validate_part")
        ):
            category = CATEGORY_OPC
            break
        if any(
            needle in lowered
            for needle in (
                "streamingarchivewriter",
                "central_directory",
                "centraldirectory",
                "zip",
                "directoryentry",
                "member_name",
            )
        ):
            category = CATEGORY_ZIP
            break
        if any(needle in lowered for needle in ("deflate", "miniz", "libz", "compress")):
            category = CATEGORY_DEFLATE
            break
        if any(
            needle in lowered
            for needle in (
                "streamingpresentationwriter",
                "write_pptx_stream",
                "write_text_box",
                "text_box",
            )
        ):
            category = CATEGORY_WRITER
            break
    return phase, category


def _aggregate(mapping: dict[str, Aggregate], key: str) -> Aggregate:
    value = mapping.get(key)
    if value is None:
        value = Aggregate()
        mapping[key] = value
    return value


def _stack_aggregate(mapping: dict[StackKey, Aggregate], key: StackKey) -> Aggregate:
    value = mapping.get(key)
    if value is None:
        value = Aggregate()
        mapping[key] = value
    return value


def _parse_frame_groups(
    fields: Sequence[bytes],
    strings: Sequence[str],
    module_id: int,
    label: str,
) -> tuple[Frame, ...]:
    if len(fields) == 2:
        return ()
    module = _string(strings, module_id)
    frame_fields = fields[2:]

    def frame_from(raw: Sequence[bytes], *, inlined: bool) -> Frame:
        function_id = parse_hex(raw[0], f"{label} function id")
        file_id = parse_hex(raw[1], f"{label} file id") if len(raw) >= 2 else 0
        line = parse_hex(raw[2], f"{label} line") if len(raw) >= 3 else None
        return Frame(
            function=_string(strings, function_id),
            file=_string(strings, file_id) if len(raw) >= 2 else None,
            line=line,
            module=module,
            inlined=inlined,
            complete=len(raw) == 3,
        )

    frames: list[Frame] = []
    if len(frame_fields) < 3:
        # Heaptrack's reader accepts a partially populated direct frame when
        # debug information ends at the function or file field. Preserve its
        # visible function for conservative classification and mark it
        # incomplete instead of rejecting the otherwise valid trace.
        frames.append(frame_from(frame_fields, inlined=False))
        return tuple(frames)
    frames.append(frame_from(frame_fields[:3], inlined=False))
    complete_inline = len(frame_fields[3:]) // 3
    for offset in range(3, 3 + complete_inline * 3, 3):
        frames.append(frame_from(frame_fields[offset : offset + 3], inlined=True))
    # A partial inline group is ignored by Heaptrack's own readFrame loop;
    # the complete direct frame above still remains usable.
    return tuple(frames)


def _parse_print_number(value: str, label: str) -> int:
    cleaned = value.replace(",", "")
    if not cleaned.isdigit():
        fail(f"heaptrack_print {label} is not a decimal integer: {value!r}")
    return int(cleaned, 10)


PRINT_CALLS_RE = re.compile(
    r"^\s*(?P<calls>[0-9][0-9,]*)\s+calls?\s+to\s+allocation functions\s+with\s+"
    r"(?P<peak>[0-9][0-9,]*)B\s+peak\s+consumption\s+from\s*$"
)
PRINT_CALLS_ALT_RE = re.compile(
    r"^\s*(?P<calls>[0-9][0-9,]*)\s+calls?\s+with\s+(?P<peak>[0-9][0-9,]*)B\s+peak\s+consumption\s+from\s*:??\s*$"
)


def parse_heaptrack_print(path: Path) -> dict[str, object]:
    """Parse only the stable top-level call/peak records from heaptrack_print.

    ``heaptrack_print`` does not expose requested bytes for each filtered row;
    therefore this fallback is deliberately separate from raw event totals.
    """

    if path.suffix.lower() == ".zst":
        fail(f"heaptrack_print fallback cannot read zstd directly: {path}")
    opener = gzip.open if path.suffix.lower() == ".gz" else open
    rows: list[dict[str, object]] = []
    pending: dict[str, object] | None = None
    try:
        with opener(path, "rt", encoding="utf-8", errors="strict") as stream:
            for raw in stream:
                line = raw.rstrip("\n")
                match = PRINT_CALLS_RE.match(line) or PRINT_CALLS_ALT_RE.match(line)
                if match:
                    pending = {
                        "allocation_calls": _parse_print_number(
                            match.group("calls"), "calls"
                        ),
                        "peak_live_bytes": _parse_print_number(
                            match.group("peak"), "peak"
                        ),
                        "requested_bytes": None,
                        "stack": [],
                    }
                    rows.append(pending)
                    continue
                if pending is None:
                    continue
                stripped = line.strip()
                if not stripped or stripped.startswith("in "):
                    continue
                if stripped.startswith(("MOST ", "PEAK ", "TEMPORARY ")):
                    pending = None
                    continue
                # heaptrack_print emits a stack as indented symbol lines; keep
                # only non-header lines and let callers see the raw spelling.
                if line.startswith(" "):
                    stack = pending["stack"]
                    assert isinstance(stack, list)
                    stack.append(stripped)
    except (OSError, UnicodeError) as error:
        raise AnalysisError(f"cannot read heaptrack_print fallback {path}: {error}") from error
    return {
        "source": str(path),
        "rows": rows,
        "requested_bytes": "unavailable in heaptrack_print text; use raw + events",
        "scope": "whole-process filtered report; peak values are not phase-local",
    }


def _analyze_trace_full(
    path: Path,
    lane: str,
    *,
    print_path: Path | None = None,
    max_stack_depth: int = 256,
) -> TraceResult:
    """Parse one interpreted v3 stream and aggregate aligned allocation lifetimes."""

    if max_stack_depth < 1:
        fail("max_stack_depth must be positive")
    compressed_sha256, compressed_bytes = file_binding(path)
    strings: list[str] = []
    instruction_pointers: list[InstructionPointer] = []
    traces: list[TraceNode] = []
    allocation_infos: list[AllocationInfo] = []
    active_counts: Counter[int] = Counter()
    stats = ParseStats()
    heaptrack_version: int | None = None
    file_version: int | None = None
    command: str | None = None
    phase: dict[str, Aggregate] = {}
    category: dict[str, Aggregate] = {}
    stacks: dict[StackKey, Aggregate] = {}
    stack_names: dict[int, tuple[str, ...]] = {}
    stack_classification: dict[int, tuple[str, str]] = {}
    unresolved_trace_ids: set[int] = set()
    current_live = 0
    all_peak = 0
    all_peak_event: int | None = None
    final_outstanding = 0
    unclassified_phase_events = 0
    unknown_scope_writer_helper_events = 0
    event_ordinal = 0

    def resolve(trace_id: int) -> tuple[tuple[str, ...], bool]:
        cached = stack_names.get(trace_id)
        if cached is not None:
            return cached, trace_id not in unresolved_trace_ids
        names, complete = trace_frames(
            trace_id, traces, instruction_pointers, strings
        )
        if len(names) > max_stack_depth:
            names = names[:max_stack_depth]
            complete = False
        stack_names[trace_id] = names
        if not complete:
            unresolved_trace_ids.add(trace_id)
        return names, complete

    try:
        stream = open_trace(path)
    except AnalysisError:
        raise
    try:
        with stream:
            for line_number, line in enumerate(stream, start=1):
                stats.lines = checked_add(stats.lines, 1, "trace lines")
                if not line.endswith(b"\n"):
                    fail(f"{path}:{line_number}: unterminated record")
                if not line.strip():
                    fail(f"{path}:{line_number}: blank record")
                if line.startswith(b"#"):
                    stats.comments = checked_add(stats.comments, 1, "comments")
                    continue
                if len(line) < 2 or line[1:2] != b" ":
                    fail(f"{path}:{line_number}: record mode is not followed by one space")
                mode = line[:1]
                if mode == b"v":
                    fields = parse_fields(line, 2, mode)
                    if heaptrack_version is not None:
                        fail(f"{path}:{line_number}: duplicate version record")
                    heaptrack_version = parse_hex(fields[0], "heaptrack version")
                    file_version = parse_hex(fields[1], "file version")
                    if file_version > SUPPORTED_FILE_VERSION:
                        fail(
                            f"{path}:{line_number}: file version {file_version} exceeds supported v3"
                        )
                    if file_version < 3:
                        fail(
                            f"{path}:{line_number}: only sized interpreted v3 strings are supported"
                        )
                elif mode == b"X":
                    if command is not None:
                        fail(f"{path}:{line_number}: duplicate command record")
                    command = decode_string(line[2:].rstrip(b"\n"), "command")
                elif mode == b"s":
                    if file_version is None or file_version < 3:
                        fail(f"{path}:{line_number}: string before v3 version record")
                    strings.append(parse_sized_string(line[2:].rstrip(b"\n"), "string"))
                    stats.strings = checked_add(stats.strings, 1, "strings")
                elif mode == b"i":
                    fields = line[2:].rstrip(b"\n").split(b" ")
                    if len(fields) < 2 or any(field == b"" for field in fields):
                        fail(f"{path}:{line_number}: malformed instruction-pointer record")
                    ip = parse_hex(fields[0], "instruction pointer")
                    module_id = parse_hex(fields[1], "module id")
                    frames = _parse_frame_groups(
                        fields, strings, module_id, "instruction-pointer frame"
                    )
                    instruction_pointers.append(
                        InstructionPointer(module_id=module_id, frames=frames)
                    )
                    stats.instruction_pointers = checked_add(
                        stats.instruction_pointers, 1, "instruction pointers"
                    )
                    del ip  # address identity is not needed for attribution
                elif mode == b"t":
                    fields = parse_fields(line, 2, mode)
                    traces.append(
                        TraceNode(
                            ip_id=parse_hex(fields[0], "trace instruction-pointer id"),
                            parent_id=parse_hex(fields[1], "trace parent id"),
                        )
                    )
                    stats.traces = checked_add(stats.traces, 1, "traces")
                elif mode == b"a":
                    fields = parse_fields(line, 2, mode)
                    allocation_infos.append(
                        AllocationInfo(
                            size=parse_hex(fields[0], "allocation size"),
                            trace_id=parse_hex(fields[1], "allocation trace id"),
                        )
                    )
                    stats.allocation_descriptors = checked_add(
                        stats.allocation_descriptors, 1, "allocation descriptors"
                    )
                elif mode == b"+":
                    fields = parse_fields(line, 1, mode)
                    info_id = parse_hex(fields[0], "allocation descriptor id")
                    if info_id >= len(allocation_infos):
                        fail(f"{path}:{line_number}: allocation descriptor {info_id} is out of bounds")
                    info = allocation_infos[info_id]
                    names, complete = resolve(info.trace_id)
                    if not complete:
                        unresolved_trace_ids.add(info.trace_id)
                    if info.trace_id not in stack_classification:
                        stack_classification[info.trace_id] = classify_stack(names)
                    selected_phase, selected_category = stack_classification[info.trace_id]
                    phase_aggregate = _aggregate(phase, selected_phase)
                    category_aggregate = _aggregate(category, selected_category)
                    stack_aggregate = _stack_aggregate(
                        stacks,
                        StackKey(selected_phase, selected_category, info.trace_id),
                    )
                    for aggregate in (
                        phase_aggregate,
                        category_aggregate,
                        stack_aggregate,
                    ):
                        aggregate.allocate(info.size, event_ordinal)
                    active_counts[info_id] = checked_add(
                        active_counts[info_id], 1, "active allocation count"
                    )
                    current_live = checked_add(current_live, info.size, "whole live bytes")
                    final_outstanding = checked_add(final_outstanding, 1, "outstanding allocations")
                    if current_live > all_peak:
                        all_peak = current_live
                        all_peak_event = event_ordinal
                    stats.allocation_events = checked_add(
                        stats.allocation_events, 1, "allocation events"
                    )
                    if selected_phase == PHASE_OTHER:
                        unclassified_phase_events = checked_add(
                            unclassified_phase_events, 1, "unclassified phase events"
                        )
                        if selected_category == CATEGORY_WRITER:
                            unknown_scope_writer_helper_events = checked_add(
                                unknown_scope_writer_helper_events,
                                1,
                                "unknown-scope writer-helper events",
                            )
                    event_ordinal = checked_add(event_ordinal, 1, "event ordinal")
                elif mode == b"-":
                    fields = parse_fields(line, 1, mode)
                    info_id = parse_hex(fields[0], "deallocation descriptor id")
                    if info_id >= len(allocation_infos):
                        fail(f"{path}:{line_number}: deallocation descriptor {info_id} is out of bounds")
                    if active_counts[info_id] <= 0:
                        fail(f"{path}:{line_number}: deallocation has no active matching allocation")
                    info = allocation_infos[info_id]
                    names, complete = resolve(info.trace_id)
                    if not complete:
                        unresolved_trace_ids.add(info.trace_id)
                    selected_phase, selected_category = stack_classification.get(
                        info.trace_id, classify_stack(names)
                    )
                    for aggregate in (
                        _aggregate(phase, selected_phase),
                        _aggregate(category, selected_category),
                        _stack_aggregate(
                            stacks,
                            StackKey(selected_phase, selected_category, info.trace_id),
                        ),
                    ):
                        aggregate.deallocate(info.size)
                    active_counts[info_id] -= 1
                    current_live -= info.size
                    final_outstanding -= 1
                    stats.deallocation_events = checked_add(
                        stats.deallocation_events, 1, "deallocation events"
                    )
                    event_ordinal = checked_add(event_ordinal, 1, "event ordinal")
                elif mode == b"c":
                    fields = parse_fields(line, 1, mode)
                    timestamp = parse_hex(fields[0], "timestamp")
                    if stats.max_timestamp is not None and timestamp < stats.max_timestamp:
                        fail(f"{path}:{line_number}: timestamps move backwards")
                    stats.max_timestamp = timestamp
                    stats.timestamps = checked_add(stats.timestamps, 1, "timestamps")
                elif mode == b"R":
                    fields = parse_fields(line, 1, mode)
                    rss = parse_hex(fields[0], "RSS")
                    stats.peak_rss = rss if stats.peak_rss is None else max(stats.peak_rss, rss)
                    stats.rss_events = checked_add(stats.rss_events, 1, "RSS events")
                elif mode == b"I":
                    parse_fields(line, 2, mode)
                elif mode == b"S":
                    # Embedded suppression patterns are opaque by design.
                    if len(line) <= 2:
                        fail(f"{path}:{line_number}: empty suppression record")
                elif mode == b"A":
                    # Heaptrack can mark an attached process.  The marker has
                    # no fields and does not alter allocation lifetimes.
                    if line[2:].rstrip(b"\n"):
                        fail(f"{path}:{line_number}: attached marker has fields")
                else:
                    fail(f"{path}:{line_number}: unsupported interpreted record {mode!r}")
    except OSError as error:
        raise AnalysisError(f"cannot read {path}: {error}") from error

    if heaptrack_version is None or file_version is None:
        fail(f"{path}: missing v3 version record")
    if not command:
        fail(f"{path}: missing command record")
    if current_live < 0 or final_outstanding < 0:
        fail(f"{path}: negative final live state")
    limitations = [
        "phase and category labels are conservative substring matches over visible direct/inlined symbols",
        "the event timeline is exact for Heaptrack plus/minus order; it is not an RSS or allocator-region measurement",
        "heaptrack_print filtering cannot replace raw requested-byte attribution or phase-local peaks",
    ]
    if unresolved_trace_ids:
        limitations.append(
            "some allocation traces have unresolved instruction-pointer/frame metadata; their totals remain in other/unresolved rows"
        )
    fallback = parse_heaptrack_print(print_path) if print_path is not None else None
    return TraceResult(
        lane=lane,
        path=str(path),
        compressed_sha256=compressed_sha256,
        compressed_bytes=compressed_bytes,
        file_version=file_version,
        heaptrack_version=heaptrack_version,
        command=command,
        stats=stats,
        phase=phase,
        category=category,
        stacks=stacks,
        stack_names=stack_names,
        unresolved_trace_ids=unresolved_trace_ids,
        unclassified_phase_events=unclassified_phase_events,
        unknown_scope_writer_helper_events=unknown_scope_writer_helper_events,
        all_peak_live_bytes=all_peak,
        all_peak_event=all_peak_event,
        final_live_bytes=current_live,
        final_outstanding_allocations=final_outstanding,
        limitations=limitations,
        fallback=fallback,
    )


METADATA_RECORD_RE = re.compile(rb"(?m)^[vXIsitaSA#](?: [^\n]*)?\n")


def _iter_regex_records(path: Path, pattern: re.Pattern[bytes]) -> Iterator[bytes]:
    """Yield selected complete records while the regex engine skips bulk events.

    The large PPTX profile contains hundreds of millions of short ``+``/``-``
    records.  Splitting every line in Python defeats the purpose of a scoped
    analysis, so this iterator lets the C regex scanner skip those records and
    yields only metadata or selected descriptor events.  A small carry buffer
    preserves records crossing an input chunk boundary.
    """

    carry = b""
    try:
        stream = open_trace(path)
        with stream:
            while True:
                chunk = stream.read(8 * 1024 * 1024)
                if not chunk:
                    break
                data = carry + chunk
                boundary = data.rfind(b"\n")
                if boundary < 0:
                    carry = data
                    continue
                complete = data[: boundary + 1]
                carry = data[boundary + 1 :]
                for match in pattern.finditer(complete):
                    yield match.group(0)
    except OSError as error:
        raise AnalysisError(f"cannot read selected records from {path}: {error}") from error
    if carry:
        fail(f"{path}: unterminated trailing record in selected scan")


@dataclass
class _Metadata:
    strings: list[str] = field(default_factory=list)
    instruction_pointers: list[InstructionPointer] = field(default_factory=list)
    traces: list[TraceNode] = field(default_factory=list)
    allocation_infos: list[AllocationInfo] = field(default_factory=list)
    stats: ParseStats = field(default_factory=ParseStats)
    heaptrack_version: int | None = None
    file_version: int | None = None
    command: str | None = None


def _parse_metadata(path: Path) -> _Metadata:
    """Read only trace/string/descriptor metadata using chunked C regex scans."""

    metadata = _Metadata()
    for line_number, line in enumerate(_iter_regex_records(path, METADATA_RECORD_RE), start=1):
        metadata.stats.lines = checked_add(metadata.stats.lines, 1, "metadata records")
        if line.startswith(b"#"):
            metadata.stats.comments = checked_add(metadata.stats.comments, 1, "comments")
            continue
        if len(line) < 2 or line[1:2] != b" ":
            fail(f"{path}: selected metadata record {line_number} has no mode separator")
        mode = line[:1]
        if mode == b"v":
            fields = parse_fields(line, 2, mode)
            if metadata.heaptrack_version is not None:
                fail(f"{path}: duplicate version record")
            metadata.heaptrack_version = parse_hex(fields[0], "heaptrack version")
            metadata.file_version = parse_hex(fields[1], "file version")
            if metadata.file_version != SUPPORTED_FILE_VERSION:
                fail(f"{path}: fast selected scan requires interpreted file version 3")
        elif mode == b"X":
            if metadata.command is not None:
                fail(f"{path}: duplicate command record")
            metadata.command = decode_string(line[2:].rstrip(b"\n"), "command")
        elif mode == b"s":
            if metadata.file_version != SUPPORTED_FILE_VERSION:
                fail(f"{path}: string before version 3 record")
            metadata.strings.append(
                parse_sized_string(line[2:].rstrip(b"\n"), "string")
            )
            metadata.stats.strings = checked_add(metadata.stats.strings, 1, "strings")
        elif mode == b"i":
            fields = line[2:].rstrip(b"\n").split(b" ")
            if len(fields) < 2 or any(field == b"" for field in fields):
                fail(f"{path}: malformed instruction-pointer record")
            module_id = parse_hex(fields[1], "module id")
            frames = _parse_frame_groups(
                fields, metadata.strings, module_id, "instruction-pointer frame"
            )
            metadata.instruction_pointers.append(
                InstructionPointer(module_id=module_id, frames=frames)
            )
            metadata.stats.instruction_pointers = checked_add(
                metadata.stats.instruction_pointers, 1, "instruction pointers"
            )
        elif mode == b"t":
            fields = parse_fields(line, 2, mode)
            metadata.traces.append(
                TraceNode(
                    ip_id=parse_hex(fields[0], "trace instruction-pointer id"),
                    parent_id=parse_hex(fields[1], "trace parent id"),
                )
            )
            metadata.stats.traces = checked_add(metadata.stats.traces, 1, "traces")
        elif mode == b"a":
            fields = parse_fields(line, 2, mode)
            metadata.allocation_infos.append(
                AllocationInfo(
                    size=parse_hex(fields[0], "allocation size"),
                    trace_id=parse_hex(fields[1], "allocation trace id"),
                )
            )
            metadata.stats.allocation_descriptors = checked_add(
                metadata.stats.allocation_descriptors, 1, "allocation descriptors"
            )
        elif mode == b"I":
            parse_fields(line, 2, mode)
        elif mode == b"S":
            if len(line) <= 2:
                fail(f"{path}: empty suppression record")
        elif mode == b"A":
            if line[2:].rstrip(b"\n"):
                fail(f"{path}: attached marker has fields")
    if metadata.heaptrack_version is None or metadata.file_version is None:
        fail(f"{path}: missing interpreted version record")
    if not metadata.command:
        fail(f"{path}: missing command record")
    return metadata


def descriptor_pattern(identifiers: Iterable[int]) -> bytes:
    """Factor hexadecimal alternatives so skipped events do not scan every ID."""
    trie: dict = {}
    for identifier in identifiers:
        if identifier < 0:
            fail('negative descriptor identifier')
        node = trie
        for character in f'{identifier:x}'.encode('ascii'):
            node = node.setdefault(character, {})
        node[None] = True

    def encode(node: dict) -> bytes:
        branches = [bytes([key]) + encode(node[key]) for key in sorted(k for k in node if k is not None)]
        if not branches:
            return b''
        result = branches[0] if len(branches) == 1 else b'(?:' + b'|'.join(branches) + b')'
        return b'(?:' + result + b')?' if None in node else result

    return encode(trie) if trie else b'(?!)'


def _analyze_trace_fast(
    path: Path,
    lane: str,
    *,
    print_path: Path | None,
    max_stack_depth: int,
) -> TraceResult:
    """Project only descriptors with exact streaming build/run ancestry.

    This is the scalable path for the large profile. It never claims
    whole-process call or requested-byte totals: those require visiting every
    allocation event. The optional heaptrack_print fallback supplies its own
    whole-process calls/peak rows, while this pass computes exact requested
    bytes and event-order peaks for matching writer descriptors.
    """

    compressed_sha256, compressed_bytes = file_binding(path)
    metadata = _parse_metadata(path)
    strings = metadata.strings
    instruction_pointers = metadata.instruction_pointers
    traces = metadata.traces
    stack_names: dict[int, tuple[str, ...]] = {}
    stack_classification: dict[int, tuple[str, str]] = {}
    unresolved_trace_ids: set[int] = set()

    def resolve(trace_id: int) -> tuple[tuple[str, ...], bool]:
        cached = stack_names.get(trace_id)
        if cached is not None:
            return cached, trace_id not in unresolved_trace_ids
        names, complete = trace_frames(
            trace_id, traces, instruction_pointers, strings
        )
        if len(names) > max_stack_depth:
            names = names[:max_stack_depth]
            complete = False
        stack_names[trace_id] = names
        if not complete:
            unresolved_trace_ids.add(trace_id)
        return names, complete

    # The materialization preflight is intentionally not expanded into
    # Python-level per-event work here: the large shape has hundreds of
    # millions of build-corpus allocation events. The exact run projection is
    # small enough to scan losslessly; build-corpus remains an explicit
    # unavailable phase with its descriptors retained in the scope counts.
    target_descriptors: dict[int, tuple[str, str, int]] = {}
    unknown_scope_descriptors = 0
    for info_id, info in enumerate(metadata.allocation_infos):
        names, complete = resolve(info.trace_id)
        if not complete:
            unresolved_trace_ids.add(info.trace_id)
        selected_phase, selected_category = classify_stack(names)
        if selected_phase == PHASE_RUN:
            target_descriptors[info_id] = (
                selected_phase,
                selected_category,
                info.trace_id,
            )
        elif selected_category == CATEGORY_WRITER:
            unknown_scope_descriptors = checked_add(
                unknown_scope_descriptors,
                1,
                "unknown-scope writer descriptors",
            )

    phase: dict[str, Aggregate] = {}
    category: dict[str, Aggregate] = {}
    stacks: dict[StackKey, Aggregate] = {}
    active_counts: Counter[int] = Counter()
    current_live = 0
    all_peak = 0
    all_peak_event: int | None = None
    final_outstanding = 0
    event_ordinal = 0
    if target_descriptors:
        # Require a token boundary before the optional space tail. This keeps
        # descriptor ``a`` from matching valid descriptor ``ab`` events, while
        # still surfacing malformed selected records such as ``+ a extra``.
        event_pattern = re.compile(
            rb"(?m)^([+-]) (" + descriptor_pattern(target_descriptors) + rb")((?: [^\n]*)?)\n"
        )
        for record in _iter_regex_records(path, event_pattern):
            match = event_pattern.fullmatch(record)
            if match is None:
                fail(f"{path}: selected event record is malformed")
            if match.group(3):
                fail(f"{path}: selected event record has trailing fields")
            mode = match.group(1)
            info_id = parse_hex(match.group(2), "selected allocation descriptor id")
            if info_id not in target_descriptors:
                fail(f"{path}: selected event descriptor disappeared during scan")
            selected_phase, selected_category, trace_id = target_descriptors[info_id]
            info = metadata.allocation_infos[info_id]
            stack_key = StackKey(selected_phase, selected_category, trace_id)
            aggregates = (
                _aggregate(phase, selected_phase),
                _aggregate(category, selected_category),
                _stack_aggregate(stacks, stack_key),
            )
            if mode == b"+":
                for aggregate in aggregates:
                    aggregate.allocate(info.size, event_ordinal)
                active_counts[info_id] = checked_add(
                    active_counts[info_id], 1, "selected active allocation count"
                )
                current_live = checked_add(current_live, info.size, "selected live bytes")
                final_outstanding = checked_add(
                    final_outstanding, 1, "selected outstanding allocations"
                )
                if current_live > all_peak:
                    all_peak = current_live
                    all_peak_event = event_ordinal
                metadata.stats.allocation_events = checked_add(
                    metadata.stats.allocation_events, 1, "selected allocation events"
                )
            else:
                if active_counts[info_id] <= 0:
                    fail(f"{path}: selected deallocation has no active matching allocation")
                for aggregate in aggregates:
                    aggregate.deallocate(info.size)
                active_counts[info_id] -= 1
                current_live -= info.size
                final_outstanding -= 1
                metadata.stats.deallocation_events = checked_add(
                    metadata.stats.deallocation_events, 1, "selected deallocation events"
                )
            event_ordinal = checked_add(event_ordinal, 1, "selected event ordinal")

    limitations = [
        "fast scoped mode skips non-matching plus/minus events in the C regex scan; whole-process calls, requested bytes, and peak are unavailable from raw projection",
        "phase attribution measures exact pptx_streaming_create::run ancestry; build-corpus descriptors remain explicit unknown/unavailable scope because their hundreds of millions of events are not expanded in Python",
        "the selected timeline peak is exact for matching descriptors in Heaptrack event order, not whole-process RSS or allocator peak",
        "category labels use the nearest visible direct/inlined frame; Deflate is generic compression evidence and does not claim initialization",
        "heaptrack_print fallback is whole-process and cannot supply requested bytes for a phase",
    ]
    if unresolved_trace_ids:
        limitations.append(
            "some selected traces have unresolved instruction-pointer/frame metadata"
        )
    fallback = parse_heaptrack_print(print_path) if print_path is not None else None
    metadata.stats.timestamps = 0
    metadata.stats.rss_events = 0
    return TraceResult(
        lane=lane,
        path=str(path),
        compressed_sha256=compressed_sha256,
        compressed_bytes=compressed_bytes,
        file_version=metadata.file_version,
        heaptrack_version=metadata.heaptrack_version,
        command=metadata.command,
        stats=metadata.stats,
        phase=phase,
        category=category,
        stacks=stacks,
        stack_names=stack_names,
        unresolved_trace_ids=unresolved_trace_ids,
        unclassified_phase_events=0,
        unknown_scope_writer_helper_events=unknown_scope_descriptors,
        all_peak_live_bytes=all_peak,
        all_peak_event=all_peak_event,
        final_live_bytes=current_live,
        final_outstanding_allocations=final_outstanding,
        limitations=limitations,
        unmeasured_phases={
            PHASE_BUILD: "fast scoped scan intentionally omits build-corpus plus/minus events"
        },
        event_scope="exact_phase_matching_descriptors_only",
        complete_event_scan=False,
        fallback=fallback,
    )


def analyze_trace(
    path: Path,
    lane: str,
    *,
    print_path: Path | None = None,
    max_stack_depth: int = 256,
    full: bool = False,
) -> TraceResult:
    """Analyze one trace; use full=True only for bounded-size traces."""

    if max_stack_depth < 1:
        fail("max_stack_depth must be positive")
    if full:
        return _analyze_trace_full(
            path,
            lane,
            print_path=print_path,
            max_stack_depth=max_stack_depth,
        )
    return _analyze_trace_fast(
        path,
        lane,
        print_path=print_path,
        max_stack_depth=max_stack_depth,
    )


def parse_named(value: str, label: str) -> tuple[str, Path]:
    if "=" not in value:
        fail(f"{label} must use NAME=PATH")
    name, raw_path = value.split("=", 1)
    if not name or not raw_path:
        fail(f"{label} must have non-empty NAME and PATH")
    return name, Path(raw_path)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--trace",
        action="append",
        default=[],
        metavar="LANE=PATH",
        help="interpreted Heaptrack v3 stream, plain or gzip (repeatable)",
    )
    parser.add_argument(
        "--heaptrack-print",
        action="append",
        default=[],
        metavar="LANE=PATH",
        help="optional heaptrack_print text fallback (repeatable)",
    )
    parser.add_argument(
        "--full",
        action="store_true",
        help="visit every plus/minus event; suitable only for bounded traces",
    )
    parser.add_argument("--output", type=Path, help="write JSON here; stdout by default")
    args = parser.parse_args(argv)
    if not args.trace:
        parser.error("at least one --trace=LANE=PATH is required")
    try:
        print_paths = dict(parse_named(item, "--heaptrack-print") for item in args.heaptrack_print)
        traces = []
        for item in args.trace:
            lane, path = parse_named(item, "--trace")
            traces.append(
                analyze_trace(
                    path,
                    lane,
                    print_path=print_paths.get(lane),
                    full=args.full,
                )
            )
        output = {
            "schema": SCHEMA,
            "tool": {
                "name": "heap_analyze.py",
                "python": f"{sys.version_info.major}.{sys.version_info.minor}",
                "format": "interpreted_heaptrack_v3",
                "upstream_format_source": "https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/src/analyze/accumulatedtracedata.cpp",
            },
            "traces": [trace.as_dict() for trace in traces],
        }
        encoded = json.dumps(output, ensure_ascii=False, sort_keys=True, indent=2) + "\n"
        if args.output is None:
            sys.stdout.write(encoded)
        else:
            args.output.write_text(encoded, encoding="utf-8")
    except AnalysisError as error:
        parser.exit(2, f"heap_analyze.py: error: {error}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
