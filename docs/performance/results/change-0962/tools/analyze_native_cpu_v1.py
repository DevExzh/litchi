#!/usr/bin/env python3
"""Conservative offline CPU attribution for the retained 0962 benchmark.

The analyzer consumes a ``perf script --inline`` stream.  With ``--perf-data``
it starts only ``perf script``; it never starts the benchmark, a workload, or a
build.  ``--script`` is provided for sealed exports and synthetic controls.

The result deliberately keeps three ideas separate:

* a qualified row has a physical ``timed_iteration`` frame and an inline
  ``stage_timed_commit`` frame below it in the leaf-first callchain;
* rows with only inline timer evidence are reported as ``inline_only`` and do
  not enter the qualified denominator;
* the direct sampled frame and its mutually exclusive owner category are
  exclusive.  Presence tags are inclusive and may overlap.

Every header record is assigned to exactly one disjoint class.  Sample counts
and event periods are checked against the sum of those classes before output
is written.  Unknown frames, missing source lines, bad source ranges, and
ambiguous marker matches fail closed for phase qualification.
"""

from __future__ import annotations

import argparse
import collections
import decimal
import hashlib
import json
import os
import re
import struct
import subprocess
import sys
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Iterator, TextIO


SCHEMA = "litchi.performance.0962.native-cpu-attribution.v1"
DEFAULT_EVENT = "cycles:u"
DEFAULT_SOURCE_SUFFIX = "xlsx_managed_compression_benchmark.rs"
DEFAULT_STAGE_RANGE = (382, 402)
DEFAULT_TIMED_RANGE = (404, 449)
MAX_LINE_BYTES = 16 * 1024 * 1024
MAX_OUTPUT_BYTES = 1 * 1024 * 1024
MIN_QUALIFIED_SAMPLES = 100
MIN_QUALIFIED_PERIOD = 100_000

# ``perf script -F ...`` prints one header, then leaf-first callchain frames.
# The fields used by this tool intentionally include period and srcline.
SCRIPT_FIELDS = "comm,pid,tid,time,event,period,ip,sym,symoff,dso,srcline"
HEADER_RE = re.compile(
    r"^(?P<comm>.+?)\s+(?P<pid>[0-9]+)/(?P<tid>[0-9]+)\s+"
    r"(?P<time>[^:]+):\s+(?P<period>[0-9]+)\s+"
    r"(?P<event>\S+):\s*$"
)
FRAME_RE = re.compile(r"^\s*(?P<ip>(?:0x)?[0-9a-fA-F]+)(?:\s+(?P<rest>.*?))?\s*$")
INLINE_SUFFIX_RE = re.compile(r"\s+\(inlined\)\s*$")
DSO_SUFFIX_RE = re.compile(r"\s+\((?P<dso>(?:/|\[)[^)]*)\)\s*$")
SOURCE_RE = re.compile(
    r"^\s*(?P<file>.*):(?P<line>[0-9]+)(?P<inline>\s+\(inlined\))?\s*$"
)
BUILD_ID_RE = re.compile(r"^\S+\[[0-9a-fA-F]+\]$")
OFFSET_RE = re.compile(r"\+0x[0-9a-fA-F]+$")


class AnalysisError(RuntimeError):
    """An input or invariant error that must stop the analysis."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AnalysisError(message)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path, label: str) -> dict[str, Any]:
    path = path.expanduser()
    require(path.is_absolute(), f"{label} must be an absolute path: {path}")
    require(path.exists() and path.is_file(), f"{label} is not a regular file: {path}")
    require(not path.is_symlink(), f"{label} must not be a symlink: {path}")
    stat = path.stat()
    return {
        "path": str(path),
        "bytes": stat.st_size,
        "mode": stat.st_mode & 0o777,
        "sha256": sha256_file(path),
    }


def elf_build_id(path: Path) -> str:
    """Read the GNU build-id from an ELF64 little-endian binary.

    This is intentionally a small read-only identity check.  Runtime address
    mapping still belongs to the capture's mmap evidence; a symbol name alone
    never establishes that the sampled DSO is the bound executable.
    """
    with path.open("rb") as stream:
        ident = stream.read(64)
        require(
            len(ident) == 64
            and ident[:4] == b"\x7fELF"
            and ident[4] == 2
            and ident[5] == 1,
            "binary binding must be ELF64 little-endian",
        )
        fields = struct.unpack_from("<HHIQQQIHHHHHH", ident, 16)
        _, _, _, _, phoff, _, _, _, phentsize, phnum, _, _, _ = fields
        require(phentsize >= 56 and phnum > 0, "binary ELF program header table is invalid")
        stream.seek(phoff)
        notes: list[tuple[int, int]] = []
        for _ in range(phnum):
            raw = stream.read(phentsize)
            require(len(raw) >= 56, "truncated binary ELF program header")
            p_type, _, p_offset, _, _, p_filesz, _, _ = struct.unpack_from("<IIQQQQQQ", raw, 0)
            if p_type == 4 and p_filesz:
                notes.append((p_offset, p_filesz))
        for offset, size in notes:
            stream.seek(offset)
            data = stream.read(size)
            position = 0
            while position + 12 <= len(data):
                namesz, descsz, ntype = struct.unpack_from("<III", data, position)
                position += 12
                name_end = position + ((namesz + 3) // 4) * 4
                desc_end = name_end + ((descsz + 3) // 4) * 4
                require(desc_end <= len(data), "truncated binary ELF note")
                name = data[position : position + namesz].rstrip(b"\0")
                desc = data[name_end : name_end + descsz]
                if name == b"GNU" and ntype == 3:
                    return desc.hex()
                position = desc_end
    raise AnalysisError("binary binding has no GNU build-id")


def parse_range(text: str, label: str) -> tuple[int, int]:
    match = re.fullmatch(r"([0-9]+)(?:-([0-9]+))?", text.strip())
    require(match is not None, f"{label} must be LINE or LINE-LINE: {text!r}")
    start = int(match.group(1))
    end = int(match.group(2) or match.group(1))
    require(start > 0 and end >= start, f"invalid {label}: {text!r}")
    return start, end


def parse_hex(text: str, label: str) -> int:
    try:
        value = int(text, 16)
    except ValueError as exc:
        raise AnalysisError(f"invalid {label}: {text!r}") from exc
    require(value >= 0, f"negative {label}: {text!r}")
    return value


@dataclass
class Frame:
    index: int
    ip: int
    symbol: str
    dso: str
    inline: bool
    unknown: bool
    source_file: str | None = None
    source_line: int | None = None
    source_inline: bool = False
    raw: str = ""

    @property
    def known(self) -> bool:
        return not self.unknown and bool(self.symbol)

    @property
    def function(self) -> str:
        return function_name(self.symbol)

    @property
    def inline_evidence(self) -> bool:
        """Preserve either perf's symbol or srcline inline marker."""
        return self.inline or self.source_inline


@dataclass
class Sample:
    ordinal: int
    comm: str
    pid: int
    tid: int
    time: str
    event: str
    period: int
    frames: list[Frame] = field(default_factory=list)


def canonical_time(text: str) -> str:
    """Canonicalize perf's decimal timestamp for cross-rendering joins."""
    try:
        value = decimal.Decimal(text.strip())
    except (decimal.InvalidOperation, ValueError) as exc:
        raise AnalysisError(f"invalid sample timestamp: {text!r}") from exc
    require(value.is_finite() and value >= 0, f"invalid sample timestamp: {text!r}")
    normalized = format(value.normalize(), "f")
    if "." in normalized:
        normalized = normalized.rstrip("0").rstrip(".")
    return normalized or "0"


def sample_key(sample: Sample) -> tuple[str, int, int, str, str, int]:
    """Stable identity shared by perf's inline and no-inline renderings."""
    return (sample.comm, sample.pid, sample.tid, sample.time, sample.event, sample.period)


def function_symbol(symbol: str) -> str:
    """Drop only the terminal instruction offset from a perf symbol."""
    return OFFSET_RE.sub("", symbol.strip())


def function_name(symbol: str) -> str:
    """Return an exact Rust-ish basename without fuzzy substring matching."""
    base = function_symbol(symbol)
    return base.rsplit("::", 1)[-1]


def parse_frame(text: str, line_number: int) -> Frame:
    match = FRAME_RE.fullmatch(text.rstrip("\n"))
    require(match is not None, f"malformed frame at line {line_number}")
    ip = parse_hex(match.group("ip"), f"frame IP at line {line_number}")
    rest = (match.group("rest") or "").strip()
    dso = ""
    dso_match = DSO_SUFFIX_RE.search(rest)
    if dso_match is not None:
        dso = dso_match.group("dso")
        rest = rest[: dso_match.start()].rstrip()
    inline = bool(INLINE_SUFFIX_RE.search(rest))
    if inline:
        rest = INLINE_SUFFIX_RE.sub("", rest).rstrip()
    symbol = rest.strip()
    unknown = "[unknown]" in text.lower() or not symbol
    return Frame(
        index=-1,
        ip=ip,
        symbol=symbol,
        dso=dso,
        inline=inline,
        unknown=unknown,
        raw=text.rstrip("\n"),
    )


def attach_source(frame: Frame, text: str, line_number: int) -> None:
    match = SOURCE_RE.fullmatch(text.rstrip("\n"))
    require(match is not None, f"malformed source location at line {line_number}")
    source_file = match.group("file").strip()
    # ``perf`` uses ``:0`` (and sometimes ``??:0``) for a frame without a
    # usable DWARF source location.  Retain it as missing evidence rather than
    # rejecting the entire stream; the qualification gate will fail closed.
    frame.source_file = source_file or None
    frame.source_line = int(match.group("line"))
    frame.source_inline = bool(match.group("inline"))


def parse_samples(lines: Iterable[str], expected_event: str) -> tuple[list[Sample], dict[str, int]]:
    """Parse a perf-script stream while retaining only one sample at a time."""
    samples: list[Sample] = []
    diagnostics: collections.Counter[str] = collections.Counter()
    current: Sample | None = None
    next_ordinal = 1

    def finish() -> None:
        nonlocal current
        if current is not None:
            samples.append(current)
            current = None

    for line_number, line in enumerate(lines, 1):
        require(len(line.encode("utf-8")) <= MAX_LINE_BYTES, f"line too long at {line_number}")
        stripped = line.strip()
        header = HEADER_RE.fullmatch(line.rstrip("\n"))
        if header is not None:
            finish()
            period = int(header.group("period"))
            require(period > 0, f"nonpositive period at line {line_number}")
            event = header.group("event").strip()
            require(event == expected_event, f"unexpected event {event!r} at line {line_number}; expected {expected_event!r}")
            raw_time = header.group("time").strip()
            current = Sample(
                ordinal=next_ordinal,
                comm=header.group("comm"),
                pid=int(header.group("pid")),
                tid=int(header.group("tid")),
                time=canonical_time(raw_time),
                event=event,
                period=period,
            )
            next_ordinal += 1
            diagnostics["headers"] += 1
            continue

        if not stripped:
            finish()
            continue

        if "PERF_RECORD_LOST" in stripped:
            diagnostics["lost_records"] += 1
            finish()
            continue

        if current is None:
            # ``perf script`` may emit comment metadata only when explicitly
            # asked for headers.  Ignore comments, but never ignore data.
            if stripped.startswith("#"):
                diagnostics["comments"] += 1
                continue
            raise AnalysisError(f"orphan nonempty line {line_number}: {stripped[:160]!r}")

        is_frame = line.startswith("\t") or bool(
            re.match(r"^\s{2,}(?:0x)?[0-9a-fA-F]+(?:\s|$)", line)
        )
        if is_frame:
            frame = parse_frame(line, line_number)
            frame.index = len(current.frames)
            current.frames.append(frame)
            diagnostics["frames"] += 1
            continue

        if current.frames and SOURCE_RE.fullmatch(line.rstrip("\n")) is not None:
            attach_source(current.frames[-1], line, line_number)
            diagnostics["source_locations"] += 1
            continue

        # A DSO build-id line can be emitted after the last frame when perf
        # resolves a symbol.  It is metadata, not another callchain frame.
        if BUILD_ID_RE.fullmatch(stripped) is not None:
            diagnostics["build_id_lines"] += 1
            continue

        raise AnalysisError(f"unrecognized sample line {line_number}: {stripped[:160]!r}")

    finish()
    require(samples, "perf script contained no sample headers")
    return samples, dict(diagnostics)


def stats(samples: int, period: int, denominator: int | None = None) -> dict[str, Any]:
    share = (period / denominator) if denominator else 0.0
    return {"samples": samples, "period": period, "period_share": share}


def add_stats(target: dict[str, int], sample: Sample) -> None:
    target["samples"] += 1
    target["period"] += sample.period


def empty_stats() -> dict[str, int]:
    return {"samples": 0, "period": 0}


def source_path_matches(frame: Frame, suffix: str) -> bool:
    if frame.source_file is None:
        return False
    normalized = frame.source_file.replace("\\", "/")
    expected = suffix.replace("\\", "/").lstrip("/")
    return normalized == expected or normalized.endswith("/" + expected)


def source_check(frame: Frame, suffix: str, line_range: tuple[int, int]) -> str:
    if frame.source_file is None or frame.source_line is None:
        return "missing"
    if not source_path_matches(frame, suffix):
        return "wrong_file"
    if not (line_range[0] <= frame.source_line <= line_range[1]):
        return "wrong_line"
    return "pass"


def physical_source_check(frame: Frame) -> str:
    """Validate source evidence carried by a physical (no-inline) frame.

    The physical rendering is used to prove that the timer exists as a real
    machine-code frame.  Its DWARF file is allowed to differ from the inline
    call-site file (the retained binary reports the timer body from a module
    file), but a missing/unknown file or a zero line is not evidence.
    """
    if frame.source_file is None or frame.source_file in {"??", "[unknown]"}:
        return "missing_file"
    if frame.source_line is None or frame.source_line <= 0:
        return "missing_line"
    return "pass"


def matching_frames(frames: list[Frame], name: str) -> list[Frame]:
    return [frame for frame in frames if frame.function == name]


def marker_kind(sample: Sample, stage_range: tuple[int, int], timed_range: tuple[int, int], source_suffix: str) -> dict[str, Any]:
    stages = matching_frames(sample.frames, "stage_timed_commit")
    timed = matching_frames(sample.frames, "timed_iteration")
    stage_inline = [frame for frame in stages if frame.inline_evidence]
    stage_physical = [frame for frame in stages if not frame.inline_evidence]
    timed_inline = [frame for frame in timed if frame.inline_evidence]
    timed_physical = [frame for frame in timed if not frame.inline_evidence]
    stage_source = [source_check(frame, source_suffix, stage_range) for frame in stage_inline]
    timed_source = [source_check(frame, source_suffix, timed_range) for frame in timed]

    ordered_pairs: list[tuple[Frame, Frame]] = []
    for stage in stage_inline:
        for outer in timed:
            if stage.index < outer.index:
                ordered_pairs.append((stage, outer))

    source_valid_stage = [frame for frame, check in zip(stage_inline, stage_source) if check == "pass"]
    source_valid_timed = [frame for frame, check in zip(timed, timed_source) if check == "pass"]
    valid_pairs = [
        (stage, outer)
        for stage, outer in ordered_pairs
        if source_check(stage, source_suffix, stage_range) == "pass"
        and source_check(outer, source_suffix, timed_range) == "pass"
    ]
    return {
        "stages": stages,
        "timed": timed,
        "stage_inline": stage_inline,
        "stage_physical": stage_physical,
        "timed_inline": timed_inline,
        "timed_physical": timed_physical,
        "stage_source": stage_source,
        "timed_source": timed_source,
        "source_valid_stage": source_valid_stage,
        "source_valid_timed": source_valid_timed,
        "ordered_pairs": ordered_pairs,
        "valid_pairs": valid_pairs,
    }


def classify(sample: Sample, source_suffix: str, stage_range: tuple[int, int], timed_range: tuple[int, int]) -> tuple[str, str, dict[str, Any]]:
    if not sample.frames:
        return "empty_frames", "empty_callchain", marker_kind(sample, stage_range, timed_range, source_suffix)
    evidence = marker_kind(sample, stage_range, timed_range, source_suffix)
    if any(not frame.known for frame in sample.frames):
        return "unknown_frames", "unknown_frame_in_callchain", evidence
    if not evidence["stages"] and not evidence["timed"]:
        return "no_phase_marker", "no_timer_marker", evidence
    if not evidence["stages"]:
        return "unqualified_phase", "stage_marker_missing", evidence
    if evidence["stages"] and not evidence["stage_inline"]:
        return "unqualified_phase", "stage_marker_not_inline", evidence
    if evidence["timed"] and not evidence["timed_physical"] and not evidence["timed_inline"]:
        return "unqualified_phase", "timed_marker_not_inline_or_physical", evidence
    if evidence["stage_inline"] and evidence["stage_physical"]:
        return "ambiguous_phase", "mixed_stage_marker_forms", evidence
    if evidence["timed_inline"] and evidence["timed_physical"]:
        return "ambiguous_phase", "mixed_timed_marker_forms", evidence
    if len(evidence["stage_inline"]) > 1:
        return "ambiguous_phase", "multiple_inline_stage_markers", evidence
    if not evidence["timed"]:
        return "unqualified_phase", "timed_marker_missing", evidence
    if len(evidence["timed_physical"]) > 1:
        return "ambiguous_phase", "multiple_physical_timed_markers", evidence
    if len(evidence["timed_physical"]) == 1:
        if len(evidence["valid_pairs"]) != 1:
            if not evidence["ordered_pairs"]:
                return "unqualified_phase", "stage_is_not_inner_to_timed", evidence
            return "source_invalid", "source_check_failed_for_qualified_pair", evidence
        return "qualified", "physical_timed_with_inline_stage_and_source_checks", evidence
    if len(evidence["timed_inline"]) != 1:
        return "ambiguous_phase", "multiple_inline_timed_markers", evidence
    if len(evidence["valid_pairs"]) != 1:
        if not evidence["ordered_pairs"]:
            return "unqualified_phase", "stage_is_not_inner_to_timed", evidence
        return "source_invalid", "source_check_failed_for_inline_only_pair", evidence
    return "inline_only", "timed_and_stage_are_inline_only", evidence


def normalized_symbol(frame: Frame) -> str:
    symbol = function_symbol(frame.symbol) or "[unknown]"
    dso = frame.dso or "[dso-omitted]"
    source = ""
    if frame.source_file is not None and frame.source_line is not None and frame.source_line > 0:
        source = f" ({frame.source_file}:{frame.source_line})"
    return f"{symbol} [{dso}]{source}"


def normalized_physical_leaf(frame: Frame) -> str:
    """Keep the loaded physical leaf address alongside symbol/source text."""
    return f"0x{frame.ip:x} {normalized_symbol(frame)}"


def _is_sha(blob: str) -> bool:
    return bool(
        re.search(r"(?:^|[\s:<])sha2::", blob)
        or re.search(r"(?:^|[\s:<])sha256(?:$|[_:<(\s])", blob)
    )


def _is_compression(blob: str) -> bool:
    return bool(
        re.search(r"(?:^|[\s:<])(?:zlib(?:_|::)|flate2::|deflate(?:_|::)|inflate(?:_|::)|adler(?:_|::))", blob)
        or re.search(r"(?:^|[/\s])libz(?:\.so|$)", blob)
    )


def _is_allocator(blob: str) -> bool:
    return bool(
        re.search(
            r"(?:^|[\s:<])(?:malloc|calloc|realloc|free|__rust_alloc|__rust_dealloc)"
            r"(?:$|[+:(<\s])",
            blob,
        )
        or re.search(r"(?:^|[\s:<])alloc::alloc::", blob)
    )


def _is_runtime(blob: str) -> bool:
    return bool(
        re.search(r"(?:^|[/\s])libc\.so(?:\.|$)", blob)
        or re.search(r"(?:^|[/\s])ld-linux[^\s]*", blob)
        or re.search(r"(?:^|[\s:<])(?:std::rt::|rustc::|__libc_|rust_begin)", blob)
    )


def owner_candidates(frame: Frame) -> list[str]:
    blob = (frame.symbol + "\n" + frame.dso + "\n" + (frame.source_file or "")).lower()
    if frame.unknown:
        return ["unknown"]
    candidates: list[str] = []
    if _is_sha(blob):
        candidates.append("sha256")
    if _is_compression(blob):
        candidates.append("compression")
    if _is_allocator(blob):
        candidates.append("allocator")
    if "litchi_xlsx" in blob or "crates/litchi-xlsx" in blob:
        candidates.append("litchi_xlsx")
    if "litchi_opc" in blob or "crates/litchi-opc" in blob:
        candidates.append("litchi_opc")
    if "xlsx_managed_compression_benchmark.rs" in blob or "litchi_perf_baseline" in blob:
        candidates.append("benchmark")
    if _is_runtime(blob):
        candidates.append("runtime")
    return sorted(set(candidates)) or ["other"]


def owner_category(frame: Frame) -> str:
    candidates = owner_candidates(frame)
    if len(candidates) == 1:
        return candidates[0]
    return "ambiguous"


def inclusive_tags(frames: list[Frame]) -> list[str]:
    """Return overlapping presence labels, never an additive partition."""
    blobs = [
        (frame.symbol + "\n" + frame.dso + "\n" + (frame.source_file or "")).lower()
        for frame in frames
    ]
    blob = "\n".join(blobs)
    tags: set[str] = set()
    if any(frame.function == "timed_iteration" for frame in frames):
        tags.add("timed_iteration")
    if any(frame.function == "stage_timed_commit" for frame in frames):
        tags.add("stage_timed_commit")
    if "litchi_xlsx" in blob or "crates/litchi-xlsx" in blob:
        tags.add("litchi_xlsx")
    if "litchi_opc" in blob or "crates/litchi-opc" in blob:
        tags.add("litchi_opc")
    if any(_is_compression(frame_blob) for frame_blob in blobs):
        tags.add("compression")
    if any(_is_sha(frame_blob) for frame_blob in blobs):
        tags.add("sha256")
    if any("multi source edit" in token or "multisourceedit" in token for token in blob.splitlines()):
        tags.add("multi_source_edit")
    return sorted(tags)


def add_group(target: dict[str, dict[str, int]], key: str, sample: Sample) -> None:
    value = target.setdefault(key, {"samples": 0, "period": 0})
    value["samples"] += 1
    value["period"] += sample.period


def sorted_groups(groups: dict[str, dict[str, int]], denominator: int) -> list[dict[str, Any]]:
    return [
        {"key": key, **value, "period_share": value["period"] / denominator if denominator else 0.0}
        for key, value in sorted(groups.items(), key=lambda item: (-item[1]["period"], -item[1]["samples"], item[0]))
    ]


def open_input_script(path: Path) -> tuple[Iterator[str], dict[str, Any]]:
    input_desc = descriptor(path, "script input")

    def lines() -> Iterator[str]:
        with path.open("r", encoding="utf-8", errors="strict") as stream:
            yield from stream

    return lines(), {"mode": "script_file", "input": input_desc}


def run_perf_script(path: Path, *, inline: bool) -> tuple[Iterator[str], dict[str, Any]]:
    input_desc = descriptor(path, "perf.data input")
    command = [
        "perf",
        "script",
        "--show-lost-events",
        "--demangle",
        "--full-source-path",
        "--inline" if inline else "--no-inline",
        "-F",
        SCRIPT_FIELDS,
        "-i",
        str(path),
    ]

    def lines() -> Iterator[str]:
        try:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="strict",
            )
        except OSError as exc:
            raise AnalysisError(f"could not start perf script: {exc}") from exc
        assert process.stdout is not None
        assert process.stderr is not None
        stderr_chunks: list[str] = []

        def drain_stderr() -> None:
            assert process.stderr is not None
            stderr_chunks.extend(process.stderr)

        stderr_thread = threading.Thread(target=drain_stderr, name="perf-script-stderr")
        stderr_thread.start()
        try:
            for line in process.stdout:
                yield line
        finally:
            process.stdout.close()
        stderr_thread.join()
        process.stderr.close()
        return_code = process.wait()
        if return_code != 0:
            detail = "".join(stderr_chunks).strip().replace("\n", " ")[:1000]
            raise AnalysisError(f"perf script failed with exit {return_code}: {detail}")

    return lines(), {
        "mode": "perf_script_inline" if inline else "perf_script_no_inline",
        "input": input_desc,
        "command": command,
    }


def perf_version() -> str | None:
    try:
        completed = subprocess.run(
            ["perf", "--version"],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except OSError:
        return None
    value = completed.stdout.strip()
    return value or None


def analyze(
    samples: list[Sample],
    *,
    source_suffix: str,
    stage_range: tuple[int, int],
    timed_range: tuple[int, int],
) -> dict[str, Any]:
    total = empty_stats()
    known = empty_stats()
    qualified = empty_stats()
    inline_only = empty_stats()
    class_totals: dict[str, dict[str, int]] = {}
    reason_totals: dict[str, dict[str, int]] = {}
    comm_totals: dict[str, dict[str, int]] = {}
    self_symbols: dict[str, dict[str, int]] = {}
    owners: dict[str, dict[str, int]] = {}
    owner_candidate_sets: dict[str, dict[str, int]] = {}
    tags: dict[str, dict[str, int]] = {}
    evidence_totals: dict[str, dict[str, int]] = {}
    source_totals: dict[str, dict[str, int]] = {}

    for sample in samples:
        add_stats(total, sample)
        add_group(comm_totals, sample.comm, sample)
        if sample.frames and all(frame.known for frame in sample.frames):
            add_stats(known, sample)
        classification, reason, evidence = classify(sample, source_suffix, stage_range, timed_range)
        add_group(class_totals, classification, sample)
        add_group(reason_totals, reason, sample)

        for key, frames in (
            ("any_stage", evidence["stages"]),
            ("inline_stage", evidence["stage_inline"]),
            ("physical_stage", evidence["stage_physical"]),
            ("any_timed", evidence["timed"]),
            ("inline_timed", evidence["timed_inline"]),
            ("physical_timed", evidence["timed_physical"]),
            ("ordered_stage_timed_pair", evidence["ordered_pairs"]),
            ("source_valid_pair", evidence["valid_pairs"]),
        ):
            if frames:
                add_group(evidence_totals, key, sample)
        for check in evidence["stage_source"] + evidence["timed_source"]:
            source_totals.setdefault(check, {"samples": 0, "period": 0})
            source_totals[check]["samples"] += 1
            source_totals[check]["period"] += sample.period

        if classification == "qualified":
            add_stats(qualified, sample)
            add_group(self_symbols, normalized_symbol(sample.frames[0]), sample)
            add_group(owners, owner_category(sample.frames[0]), sample)
            add_group(owner_candidate_sets, "|".join(owner_candidates(sample.frames[0])), sample)
            for tag in inclusive_tags(sample.frames):
                add_group(tags, tag, sample)
        elif classification == "inline_only":
            add_stats(inline_only, sample)

    class_samples = sum(value["samples"] for value in class_totals.values())
    class_period = sum(value["period"] for value in class_totals.values())
    require(class_samples == total["samples"], f"sample conservation failed: {class_samples} != {total['samples']}")
    require(class_period == total["period"], f"period conservation failed: {class_period} != {total['period']}")

    return {
        "schema": SCHEMA,
        "status": "pass",
        "qualification": "qualified_ancestry_present" if qualified["samples"] else "no_qualified_ancestry",
        "scope": "whole-process cycles:u samples; qualified rows require exact benchmark ancestry and source checks; no latency, speedup, RSS, allocation, or production claim",
        "selection": {
            "leaf_first_callchain": True,
            "timed_function": "timed_iteration",
            "timed_evidence": "physical frame required for qualified scope",
            "stage_function": "stage_timed_commit",
            "stage_evidence": "inline frame required",
            "ancestor_order": "stage_timed_commit must occur before timed_iteration in perf's leaf-first order",
            "unknown_policy": "any unknown frame excludes the row from qualified and inline_only scopes",
            "ambiguous_policy": "multiple exact marker frames fail closed",
            "source_file_suffix": source_suffix,
            "stage_source_line_range": {"start": stage_range[0], "end": stage_range[1]},
            "timed_source_line_range": {"start": timed_range[0], "end": timed_range[1]},
            "physical_source_policy": "physical timed frame must carry a nonempty DWARF file and positive line; its body file may differ from the inline call-site file",
            "owner_category_policy": "a single candidate is exclusive; multiple candidates are classified as ambiguous and candidate sets remain visible",
        },
        "whole_process": stats(total["samples"], total["period"], total["period"]),
        "denominators": {
            "known_nonempty_frames": stats(known["samples"], known["period"], total["period"]),
            "qualified_ancestor": stats(qualified["samples"], qualified["period"], total["period"]),
            "inline_only": stats(inline_only["samples"], inline_only["period"], total["period"]),
        },
        "qualification_gate": {
            "minimum_samples": MIN_QUALIFIED_SAMPLES,
            "minimum_period": MIN_QUALIFIED_PERIOD,
            "qualified_samples": qualified["samples"],
            "qualified_period": qualified["period"],
            "qualified_samples_min_pass": qualified["samples"] >= MIN_QUALIFIED_SAMPLES,
            "qualified_period_min_pass": qualified["period"] >= MIN_QUALIFIED_PERIOD,
            "lost_records_zero": None,
            "admitted_without_lost_record_binding": False,
        },
        "classification": {
            key: {**value, "period_share": value["period"] / total["period"] if total["period"] else 0.0}
            for key, value in sorted(class_totals.items())
        },
        "reasons": {
            key: {**value, "period_share": value["period"] / total["period"] if total["period"] else 0.0}
            for key, value in sorted(reason_totals.items())
        },
        "marker_evidence": {
            key: {**value, "period_share": value["period"] / total["period"] if total["period"] else 0.0}
            for key, value in sorted(evidence_totals.items())
        },
        "source_line_checks": {
            key: {**value, "period_share": value["period"] / total["period"] if total["period"] else 0.0}
            for key, value in sorted(source_totals.items())
        },
        "qualified_attribution": {
            "scope": "qualified ancestor rows only",
            "exclusive_deepest_self_symbol": sorted_groups(self_symbols, qualified["period"]),
            "exclusive_owner_category": sorted_groups(owners, qualified["period"]),
            "owner_candidate_sets": sorted_groups(owner_candidate_sets, qualified["period"]),
            "inclusive_tags": sorted_groups(tags, qualified["period"]),
            "inclusive_tags_are_overlapping": True,
            "inclusive_tag_policy": "only exact phase, litchi package, compression, hash, and multi-source markers are published; generic allocator/runtime names are omitted",
            "owner_category_policy": "a single candidate is exclusive; multiple candidates are classified as ambiguous and candidate sets remain visible",
        },
        "commands": {
            "analysis": "offline parser; no workload or build launched",
            "perf_script_fields": SCRIPT_FIELDS,
        },
        "by_comm": sorted_groups(comm_totals, total["period"]),
        "conservation": {
            "input_samples": total["samples"],
            "input_period": total["period"],
            "classified_samples": class_samples,
            "classified_period": class_period,
            "sample_count_pass": class_samples == total["samples"],
            "period_sum_pass": class_period == total["period"],
        },
        "interpretation": [
            "The qualified denominator is a conservative ancestry population, not a timed phase measurement.",
            "The direct leaf symbol and owner category are disjoint per sample; inclusive tags intentionally overlap and must not be summed.",
            "Rows with inline-only timed_iteration evidence remain descriptive and are excluded from qualified phase rankings.",
        ],
    }


DEFAULT_PHYSICAL_TIMED_PREFIX = "litchi_perf_baseline::xlsx_managed_compression::benchmark::"


def physical_timed_frames(sample: Sample, prefix: str) -> list[Frame]:
    expected = prefix + "timed_iteration"
    return [
        frame
        for frame in sample.frames
        if not frame.inline_evidence
        and function_symbol(frame.symbol) == expected
        and frame.dso
        and frame.dso != "[unknown]"
    ]


def classify_pair(
    inline_sample: Sample,
    physical_sample: Sample,
    *,
    source_suffix: str,
    stage_range: tuple[int, int],
    timed_range: tuple[int, int],
    physical_prefix: str,
    binary_dso: str | None,
) -> tuple[str, str, dict[str, Any]]:
    """Classify one event-identical inline/no-inline sample pair.

    ``perf script --no-inline`` is the authority for the physical timed frame.
    The inline rendering is the authority for the inlined stage callsite and
    its source line.  Requiring the same timed IP in both renderings prevents
    a coincident function name from being attached to the wrong sample.
    """
    require(sample_key(inline_sample) == sample_key(physical_sample), "paired sample keys differ")
    inline_evidence = marker_kind(inline_sample, stage_range, timed_range, source_suffix)
    physical_timed = physical_timed_frames(physical_sample, physical_prefix)
    if binary_dso is not None:
        physical_timed = [
            frame
            for frame in physical_timed
            if frame.dso == binary_dso
        ]
    inline_timed = inline_evidence["timed"]
    inline_stage = inline_evidence["stage_inline"]
    physical_timed_source = [physical_source_check(frame) for frame in physical_timed]
    evidence = {
        **inline_evidence,
        "physical_timed": physical_timed,
        "physical_inline": [
            frame
            for frame in physical_sample.frames
            if frame.inline_evidence
        ],
        "physical_timed_source": physical_timed_source,
        "physical_symbol_identity": len(physical_timed) == 1,
        "timed_ip_match": bool(
            len(inline_timed) == 1
            and len(physical_timed) == 1
            and inline_timed[0].ip == physical_timed[0].ip
        ),
        "physical_frame_count": len(physical_sample.frames),
        "physical_unknown": any(not frame.known for frame in physical_sample.frames),
        "inline_unknown": any(not frame.known for frame in inline_sample.frames),
    }
    if not inline_sample.frames or not physical_sample.frames:
        return "empty_frames", "empty_callchain_in_one_rendering", evidence
    if evidence["physical_inline"]:
        return "ambiguous_phase", "physical_stream_contains_inline_evidence", evidence
    if evidence["inline_unknown"] or evidence["physical_unknown"]:
        return "unknown_frames", "unknown_frame_in_inline_or_physical_callchain", evidence
    if not inline_evidence["stages"] and not inline_evidence["timed"]:
        return "no_phase_marker", "no_timer_marker_in_inline_rendering", evidence
    if not inline_evidence["stages"]:
        return "unqualified_phase", "stage_marker_missing", evidence
    if inline_evidence["stages"] and not inline_stage:
        return "unqualified_phase", "stage_marker_not_inline", evidence
    if inline_evidence["stage_inline"] and inline_evidence["stage_physical"]:
        return "ambiguous_phase", "mixed_stage_marker_forms", evidence
    if inline_evidence["timed_inline"] and inline_evidence["timed_physical"]:
        return "ambiguous_phase", "mixed_timed_marker_forms", evidence
    if len(inline_stage) > 1:
        return "ambiguous_phase", "multiple_inline_stage_markers", evidence
    if not inline_timed:
        return "unqualified_phase", "timed_marker_missing", evidence
    if len(inline_timed) != 1:
        return "ambiguous_phase", "expected_one_inline_timed_marker", evidence
    if len(physical_timed) != 1:
        if not physical_timed:
            if inline_evidence["valid_pairs"]:
                return "inline_only", "inline_pair_lacks_strict_physical_timed_frame", evidence
            return "unqualified_phase", "inline_source_or_ancestor_check_failed", evidence
        return "ambiguous_phase", "multiple_strict_physical_timed_markers", evidence
    if physical_timed_source[0] != "pass":
        return "source_invalid", "physical_timed_source_check_failed", evidence
    if not inline_evidence["valid_pairs"]:
        if not inline_evidence["ordered_pairs"]:
            return "unqualified_phase", "stage_is_not_inner_to_timed", evidence
        return "source_invalid", "source_check_failed_for_paired_context", evidence
    if inline_timed[0].ip != physical_timed[0].ip:
        return "unqualified_phase", "inline_and_physical_timed_IP_mismatch", evidence
    return "qualified", "paired_physical_timed_and_inline_stage_with_source_checks", evidence


def pair_samples(
    inline_samples: list[Sample], physical_samples: list[Sample]
) -> tuple[list[tuple[Sample, Sample]], list[Sample], list[Sample], int]:
    """Join duplicate-safe event records without ever joining by position only."""
    queues: dict[tuple[str, int, int, str, str, int], collections.deque[Sample]] = collections.defaultdict(collections.deque)
    for sample in physical_samples:
        queues[sample_key(sample)].append(sample)
    pairs: list[tuple[Sample, Sample]] = []
    unmatched_inline: list[Sample] = []
    duplicate_keys = sum(max(0, len(queue) - 1) for queue in queues.values())
    for inline_sample in inline_samples:
        queue = queues.get(sample_key(inline_sample))
        if not queue:
            unmatched_inline.append(inline_sample)
            continue
        pairs.append((inline_sample, queue.popleft()))
    unmatched_physical = [sample for queue in queues.values() for sample in queue]
    return pairs, unmatched_inline, unmatched_physical, duplicate_keys


def sample_key_record(key: tuple[str, int, int, str, str, int]) -> dict[str, Any]:
    """Make an event key safe and readable in JSON evidence."""
    comm, pid, tid, time, event, period = key
    return {
        "comm": comm,
        "pid": pid,
        "tid": tid,
        "time": time,
        "event": event,
        "period": period,
    }


def header_multiset_summary(
    inline_samples: list[Sample], physical_samples: list[Sample]
) -> dict[str, Any]:
    """Compare headers as multisets, including duplicate cardinality.

    A timestamp is not a unique event identity: inherited samples can repeat
    every header field.  Counter equality is therefore checked independently
    of the queue join, and duplicate keys are exposed so callers cannot mistake
    a positional or arbitrary queue pairing for proof.
    """
    inline_counts: collections.Counter[tuple[str, int, int, str, str, int]] = collections.Counter(
        sample_key(sample) for sample in inline_samples
    )
    physical_counts: collections.Counter[tuple[str, int, int, str, str, int]] = collections.Counter(
        sample_key(sample) for sample in physical_samples
    )
    inline_duplicates = {
        key: count for key, count in inline_counts.items() if count > 1
    }
    physical_duplicates = {
        key: count for key, count in physical_counts.items() if count > 1
    }

    def duplicate_records(counts: dict[tuple[str, int, int, str, str, int], int]) -> int:
        return sum(count - 1 for count in counts.values())

    def counter_records(
        counts: collections.Counter[tuple[str, int, int, str, str, int]]
    ) -> list[dict[str, Any]]:
        return [
            {"key": sample_key_record(key), "count": counts[key]}
            for key in sorted(counts, key=lambda item: tuple(str(value) for value in item))
        ]

    inline_only = inline_counts - physical_counts
    physical_only = physical_counts - inline_counts
    inline_period = sum(sample.period for sample in inline_samples)
    physical_period = sum(sample.period for sample in physical_samples)
    return {
        "inline_header_count": len(inline_samples),
        "physical_header_count": len(physical_samples),
        "header_count_equal": len(inline_samples) == len(physical_samples),
        "inline_period": inline_period,
        "physical_period": physical_period,
        "header_period_equal": inline_period == physical_period,
        "header_key_multiplicity_equal": inline_counts == physical_counts,
        "header_multiset_equal": inline_counts == physical_counts,
        "duplicate_inline_keys": len(inline_duplicates),
        "duplicate_inline_records": duplicate_records(inline_duplicates),
        "duplicate_physical_keys": len(physical_duplicates),
        "duplicate_physical_records": duplicate_records(physical_duplicates),
        "duplicate_inline_key_details": counter_records(inline_duplicates),
        "duplicate_physical_key_details": counter_records(physical_duplicates),
        "inline_only_key_details": counter_records(inline_only),
        "physical_only_key_details": counter_records(physical_only),
    }


def analyze_paired(
    inline_samples: list[Sample],
    physical_samples: list[Sample],
    *,
    source_suffix: str,
    stage_range: tuple[int, int],
    timed_range: tuple[int, int],
    physical_prefix: str,
    binary_dso: str | None = None,
) -> dict[str, Any]:
    pairs, unmatched_inline, unmatched_physical, duplicate_keys = pair_samples(
        inline_samples, physical_samples
    )
    header_summary = header_multiset_summary(inline_samples, physical_samples)
    duplicate_key_set = {
        tuple(item["key"][field] for field in ("comm", "pid", "tid", "time", "event", "period"))
        for item in header_summary["duplicate_inline_key_details"]
    } | {
        tuple(item["key"][field] for field in ("comm", "pid", "tid", "time", "event", "period"))
        for item in header_summary["duplicate_physical_key_details"]
    }
    total = empty_stats()
    physical_total = empty_stats()
    qualified = empty_stats()
    inline_only = empty_stats()
    class_totals: dict[str, dict[str, int]] = {}
    reason_totals: dict[str, dict[str, int]] = {}
    evidence_totals: dict[str, dict[str, int]] = {}
    source_totals: dict[str, dict[str, int]] = {}
    comm_totals: dict[str, dict[str, int]] = {}
    self_symbols: dict[str, dict[str, int]] = {}
    physical_leaf_symbols: dict[str, dict[str, int]] = {}
    owners: dict[str, dict[str, int]] = {}
    owner_candidate_sets: dict[str, dict[str, int]] = {}
    tags: dict[str, dict[str, int]] = {}

    for sample in inline_samples:
        add_stats(total, sample)
        add_group(comm_totals, sample.comm, sample)
    for sample in physical_samples:
        add_stats(physical_total, sample)
    for inline_sample in unmatched_inline:
        add_group(class_totals, "unpaired_inline", inline_sample)
        add_group(reason_totals, "no_matching_no_inline_record", inline_sample)
    for inline_sample, physical_sample in pairs:
        classification, reason, evidence = classify_pair(
            inline_sample,
            physical_sample,
            source_suffix=source_suffix,
            stage_range=stage_range,
            timed_range=timed_range,
            physical_prefix=physical_prefix,
            binary_dso=binary_dso,
        )
        if sample_key(inline_sample) in duplicate_key_set:
            # A queue can make duplicate headers line up, but equal headers do
            # not identify which callchain payload belongs to which rendering.
            # Keep the row in conservation totals and fail its attribution
            # closed rather than inventing a payload match.
            classification = "ambiguous_pairing"
            reason = "duplicate_event_identity"
            evidence["pairing_key_ambiguous"] = True
        add_group(class_totals, classification, inline_sample)
        add_group(reason_totals, reason, inline_sample)
        if classification == "qualified":
            add_stats(qualified, inline_sample)
            add_group(self_symbols, normalized_symbol(inline_sample.frames[0]), inline_sample)
            add_group(physical_leaf_symbols, normalized_physical_leaf(physical_sample.frames[0]), inline_sample)
            add_group(owners, owner_category(inline_sample.frames[0]), inline_sample)
            add_group(owner_candidate_sets, "|".join(owner_candidates(inline_sample.frames[0])), inline_sample)
            for tag in inclusive_tags(inline_sample.frames):
                add_group(tags, tag, inline_sample)
        elif classification == "inline_only":
            add_stats(inline_only, inline_sample)
        for key, frames in (
            ("any_stage", evidence["stages"]),
            ("inline_stage", evidence["stage_inline"]),
            ("physical_stage", evidence["stage_physical"]),
            ("any_timed", evidence["timed"]),
            ("inline_timed", evidence["timed_inline"]),
            ("physical_timed_inline_rendering", evidence["timed_physical"]),
            ("strict_physical_timed_no_inline", evidence["physical_timed"]),
            ("ordered_stage_timed_pair", evidence["ordered_pairs"]),
            ("source_valid_pair", evidence["valid_pairs"]),
            ("physical_timed_source", evidence["physical_timed_source"]),
        ):
            if frames:
                add_group(evidence_totals, key, inline_sample)
        if evidence["timed_ip_match"]:
            add_group(evidence_totals, "timed_ip_match", inline_sample)
        for check in evidence["stage_source"] + evidence["timed_source"]:
            source_totals.setdefault(check, {"samples": 0, "period": 0})
            source_totals[check]["samples"] += 1
            source_totals[check]["period"] += inline_sample.period
        for check in evidence["physical_timed_source"]:
            key = "physical_" + check
            source_totals.setdefault(key, {"samples": 0, "period": 0})
            source_totals[key]["samples"] += 1
            source_totals[key]["period"] += inline_sample.period

    class_samples = sum(value["samples"] for value in class_totals.values())
    class_period = sum(value["period"] for value in class_totals.values())
    require(class_samples == total["samples"], f"paired sample conservation failed: {class_samples} != {total['samples']}")
    require(class_period == total["period"], f"paired period conservation failed: {class_period} != {total['period']}")
    physical_pair_samples = len(pairs) + len(unmatched_physical)
    paired_exact = (
        not unmatched_inline
        and not unmatched_physical
        and len(pairs) == len(inline_samples) == len(physical_samples)
        and header_summary["header_count_equal"]
        and header_summary["header_period_equal"]
        and header_summary["header_key_multiplicity_equal"]
        and not duplicate_key_set
    )
    result = {
        "schema": SCHEMA,
        "status": "pass" if paired_exact else "failed_closed",
        "qualification": "qualified_ancestry_present" if qualified["samples"] else "no_qualified_ancestry",
        "scope": "paired whole-process cycles:u samples; no-inline physical ancestry is joined to inline source/callsite evidence by exact event identity; no latency, speedup, RSS, allocation, or production claim",
        "selection": {
            "pair_key": ["comm", "pid", "tid", "time", "event", "period"],
            "inline_rendering": "stage_timed_commit must be exactly one inline frame, ordered before exactly one inline timed_iteration, with source checks",
            "physical_rendering": "exactly one non-inline symbol named litchi_perf_baseline::xlsx_managed_compression::benchmark::timed_iteration with nonempty DSO and valid source evidence is required",
            "timed_ip_join": "inline timed_iteration IP must equal the strict physical timed_iteration IP",
            "physical_inline_policy": "any inline-marked frame in the no-inline rendering fails closed",
            "duplicate_key_policy": "duplicate event keys are reported and fail closed because their callchain payloads cannot be identified",
            "unknown_policy": "any unknown frame in either rendering excludes the pair from qualified and inline_only scopes",
            "ambiguous_policy": "multiple exact markers or physical timed frames fail closed",
            "source_file_suffix": source_suffix,
            "stage_source_line_range": {"start": stage_range[0], "end": stage_range[1]},
            "timed_source_line_range": {"start": timed_range[0], "end": timed_range[1]},
            "physical_source_policy": "physical timed frame must carry a nonempty DWARF file and positive line; its body file may differ from the inline call-site file",
            "physical_symbol_prefix": physical_prefix,
            "physical_dso_path": binary_dso,
        },
        "whole_process": stats(total["samples"], total["period"], total["period"]),
        "physical_process": stats(physical_total["samples"], physical_total["period"], physical_total["period"]),
        "denominators": {
            "qualified_ancestor": stats(qualified["samples"], qualified["period"], total["period"]),
            "inline_only": stats(inline_only["samples"], inline_only["period"], total["period"]),
        },
        "qualification_gate": {
            "minimum_samples": MIN_QUALIFIED_SAMPLES,
            "minimum_period": MIN_QUALIFIED_PERIOD,
            "qualified_samples": qualified["samples"],
            "qualified_period": qualified["period"],
            "qualified_samples_min_pass": qualified["samples"] >= MIN_QUALIFIED_SAMPLES,
            "qualified_period_min_pass": qualified["period"] >= MIN_QUALIFIED_PERIOD,
            "lost_records_zero": None,
            "admitted_without_lost_record_binding": False,
        },
        "classification": {
            key: {**value, "period_share": value["period"] / total["period"] if total["period"] else 0.0}
            for key, value in sorted(class_totals.items())
        },
        "reasons": {
            key: {**value, "period_share": value["period"] / total["period"] if total["period"] else 0.0}
            for key, value in sorted(reason_totals.items())
        },
        "marker_evidence": {
            key: {**value, "period_share": value["period"] / total["period"] if total["period"] else 0.0}
            for key, value in sorted(evidence_totals.items())
        },
        "source_line_checks": {
            key: {**value, "period_share": value["period"] / total["period"] if total["period"] else 0.0}
            for key, value in sorted(source_totals.items())
        },
        "qualified_attribution": {
            "scope": "paired qualified ancestor rows only",
            "exclusive_deepest_self_symbol": sorted_groups(self_symbols, qualified["period"]),
            "exclusive_physical_enclosing_leaf": sorted_groups(physical_leaf_symbols, qualified["period"]),
            "exclusive_owner_category": sorted_groups(owners, qualified["period"]),
            "owner_candidate_sets": sorted_groups(owner_candidate_sets, qualified["period"]),
            "inclusive_tags": sorted_groups(tags, qualified["period"]),
            "inclusive_tags_are_overlapping": True,
            "inclusive_tag_policy": "only exact phase, litchi package, compression, hash, and multi-source markers are published; generic allocator/runtime names are omitted",
            "owner_category_policy": "a single candidate is exclusive; multiple candidates are classified as ambiguous and candidate sets remain visible",
        },
        "pairing": {
            "inline_samples": len(inline_samples),
            "physical_samples": len(physical_samples),
            "paired_samples": len(pairs),
            "unmatched_inline_samples": len(unmatched_inline),
            "unmatched_inline_period": sum(sample.period for sample in unmatched_inline),
            "unmatched_physical_samples": len(unmatched_physical),
            "unmatched_physical_period": sum(sample.period for sample in unmatched_physical),
            **header_summary,
            "queue_duplicate_physical_records": duplicate_keys,
            "exact_pairing_pass": paired_exact,
            "physical_pair_population": physical_pair_samples,
        },
        "commands": {
            "analysis": "offline parser; no workload or build launched",
            "perf_script_fields": SCRIPT_FIELDS,
            "pairing": "inline and explicit --no-inline streams are joined by event identity, never by position alone",
        },
        "by_comm": sorted_groups(comm_totals, total["period"]),
        "conservation": {
            "inline_input_samples": total["samples"],
            "inline_input_period": total["period"],
            "inline_classified_samples": class_samples,
            "inline_classified_period": class_period,
            "physical_input_samples": physical_total["samples"],
            "physical_input_period": physical_total["period"],
            "inline_sample_count_pass": class_samples == total["samples"],
            "inline_period_sum_pass": class_period == total["period"],
            "paired_stream_population_pass": paired_exact,
            "header_multiset_pass": header_summary["header_multiset_equal"],
            "physical_pair_population": physical_pair_samples,
        },
        "interpretation": [
            "The qualified denominator is a conservative paired ancestry population, not a timed phase measurement.",
            "The direct inline leaf symbol and owner category are disjoint per qualified sample; inclusive tags intentionally overlap and must not be summed.",
            "Inline-only timer evidence is retained descriptively and is excluded until a matching strict physical timed frame is proven.",
        ],
    }
    return result


def write_json(path: Path, value: dict[str, Any]) -> dict[str, Any]:
    path = path.expanduser()
    require(path.is_absolute(), f"output must be an absolute path: {path}")
    require(not path.exists() and not path.is_symlink(), f"refusing to overwrite output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n").encode("utf-8")
    require(len(encoded) < MAX_OUTPUT_BYTES, f"analysis output exceeds {MAX_OUTPUT_BYTES} bytes")
    with path.open("xb") as stream:
        stream.write(encoded)
        stream.flush()
        os.fsync(stream.fileno())
    return {"path": str(path), "bytes": len(encoded), "sha256": hashlib.sha256(encoded).hexdigest()}


def bind_qualification_gate(result: dict[str, Any], lost_records: int | None) -> None:
    gate = result["qualification_gate"]
    gate["lost_records"] = lost_records
    gate["lost_records_zero"] = lost_records == 0 if lost_records is not None else None
    gate["admitted"] = bool(
        gate["qualified_samples_min_pass"]
        and gate["qualified_period_min_pass"]
        and gate["lost_records_zero"] is True
    )
    if result["qualification"] == "qualified_ancestry_present" and not gate["admitted"]:
        result["qualification"] = "qualified_below_admission_gate"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--perf-data", type=Path, help="sealed perf.data; only perf script is run")
    source.add_argument("--script", type=Path, help="sealed perf script text or synthetic fixture")
    parser.add_argument(
        "--physical-script",
        type=Path,
        help="matching explicit --no-inline export; required with --script for paired attribution",
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--event", default=DEFAULT_EVENT)
    parser.add_argument("--source-file-suffix", default=DEFAULT_SOURCE_SUFFIX)
    parser.add_argument("--stage-lines", default=f"{DEFAULT_STAGE_RANGE[0]}-{DEFAULT_STAGE_RANGE[1]}")
    parser.add_argument("--timed-lines", default=f"{DEFAULT_TIMED_RANGE[0]}-{DEFAULT_TIMED_RANGE[1]}")
    parser.add_argument("--physical-timed-prefix", default=DEFAULT_PHYSICAL_TIMED_PREFIX)
    parser.add_argument("--binary", type=Path, help="optional exact binary binding for physical DSO basename")
    return parser


def run(arguments: argparse.Namespace) -> dict[str, Any]:
    stage_range = parse_range(arguments.stage_lines, "stage source line range")
    timed_range = parse_range(arguments.timed_lines, "timed source line range")
    require(bool(arguments.source_file_suffix.strip()), "source file suffix must not be empty")
    binary_binding = None
    binary_dso = None
    if arguments.binary is not None:
        binary_path = arguments.binary.expanduser().absolute()
        binary_binding = descriptor(binary_path, "binary binding")
        binary_binding["elf_build_id"] = elf_build_id(binary_path)
        binary_dso = str(binary_path)
    if arguments.perf_data is not None:
        require(arguments.physical_script is None, "--physical-script cannot accompany --perf-data")
        inline_lines, inline_info = run_perf_script(arguments.perf_data.expanduser().absolute(), inline=True)
        physical_lines, physical_info = run_perf_script(arguments.perf_data.expanduser().absolute(), inline=False)
        inline_samples, inline_diagnostics = parse_samples(inline_lines, arguments.event)
        physical_samples, physical_diagnostics = parse_samples(physical_lines, arguments.event)
        result = analyze_paired(
            inline_samples,
            physical_samples,
            source_suffix=arguments.source_file_suffix,
            stage_range=stage_range,
            timed_range=timed_range,
            physical_prefix=arguments.physical_timed_prefix,
            binary_dso=binary_dso,
        )
        result["input"] = {"inline": inline_info, "physical": physical_info}
        result["parser"] = {
            "inline": {
                "sample_count": len(inline_samples),
                "diagnostics": inline_diagnostics,
            },
            "physical": {
                "sample_count": len(physical_samples),
                "diagnostics": physical_diagnostics,
            },
            "perf_version": perf_version(),
        }
        if binary_binding is not None:
            result["binary_binding"] = binary_binding
        bind_qualification_gate(
            result,
            inline_diagnostics.get("lost_records", 0) + physical_diagnostics.get("lost_records", 0),
        )
    else:
        inline_lines, inline_info = open_input_script(arguments.script.expanduser().absolute())
        inline_samples, inline_diagnostics = parse_samples(inline_lines, arguments.event)
        if arguments.physical_script is None:
            result = analyze(
                inline_samples,
                source_suffix=arguments.source_file_suffix,
                stage_range=stage_range,
                timed_range=timed_range,
            )
            result["input"] = inline_info
            result["parser"] = {
                "sample_count": len(inline_samples),
                "diagnostics": inline_diagnostics,
                "perf_version": None,
            }
            bind_qualification_gate(result, inline_diagnostics.get("lost_records", 0))
        else:
            physical_lines, physical_info = open_input_script(arguments.physical_script.expanduser().absolute())
            physical_samples, physical_diagnostics = parse_samples(physical_lines, arguments.event)
            result = analyze_paired(
                inline_samples,
                physical_samples,
                source_suffix=arguments.source_file_suffix,
                stage_range=stage_range,
                timed_range=timed_range,
                physical_prefix=arguments.physical_timed_prefix,
                binary_dso=binary_dso,
            )
            result["input"] = {"inline": inline_info, "physical": physical_info}
            result["parser"] = {
                "inline": {
                    "sample_count": len(inline_samples),
                    "diagnostics": inline_diagnostics,
                },
                "physical": {
                    "sample_count": len(physical_samples),
                    "diagnostics": physical_diagnostics,
                },
                "perf_version": None,
            }
            if binary_binding is not None:
                result["binary_binding"] = binary_binding
            bind_qualification_gate(
                result,
                inline_diagnostics.get("lost_records", 0) + physical_diagnostics.get("lost_records", 0),
            )
    result["output"] = write_json(arguments.output, result)
    return result


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    arguments = parser.parse_args(argv)
    try:
        result = run(arguments)
    except (AnalysisError, OSError, UnicodeError, ValueError) as exc:
        print(json.dumps({"schema": SCHEMA, "status": "fail", "error": str(exc)}, sort_keys=True), file=sys.stderr)
        return 1
    print(json.dumps({"status": result["status"], "qualification": result["qualification"], "output": result["output"]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
