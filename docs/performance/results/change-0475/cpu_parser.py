#!/usr/bin/env python3
"""Portable parser and attribution helpers for a ``perf script`` export.

The parser is intentionally independent of the rest of the performance
evidence bundle.  A ``cycles:u`` period is an event weight; it is never
treated as elapsed time.  Accepted samples, malformed blocks, unknown and
unparsed frames, explicit truncation, and perf's lost-sample metadata all
remain visible to callers.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
import gzip
import hashlib
import json
from pathlib import Path
import re
from typing import Any, BinaryIO, Callable, Iterable, Mapping, Sequence, TextIO


EXPECTED_EVENT = "cycles:u"
DEFAULT_TOP = 40

# ``perf script`` has emitted both pid/tid spellings over time.  Keep the
# event token permissive so a malformed period can still be accounted for.
HEADER_RE = re.compile(
    r"^\s*(?P<comm>\S+)\s+(?P<pid>\d+)(?:/(?P<tid>\d+))?\s+"
    r"(?P<time>[^:]+):\s+(?P<period>\S+)\s+(?P<event>\S+)\s*$"
)
HEADER_SEPARATE_TID_RE = re.compile(
    r"^\s*(?P<comm>\S+)\s+(?P<pid>\d+)\s+(?P<tid>\d+)\s+"
    r"(?P<time>[^:]+):\s+(?P<period>\S+)\s+(?P<event>\S+)\s*$"
)
FRAME_RE = re.compile(
    r"^\s*(?:0x)?[0-9a-fA-F]+\s+(?P<symbol>.+?)\+0x[0-9a-fA-F]+\s+\([^)]*\)\s*$"
)
FRAME_RE_NO_OFFSET = re.compile(
    r"^\s*(?:0x)?[0-9a-fA-F]+\s+(?P<symbol>.+?)\s+\([^)]*\)\s*$"
)
FRAME_RE_NO_DSO = re.compile(
    r"^\s*(?:0x)?[0-9a-fA-F]+\s+(?P<symbol>.+?)"
    r"(?:\+0x[0-9a-fA-F]+)?\s*$"
)
UNKNOWN_ONLY_RE = re.compile(
    r"^\s*(?:\?\?|unknown|<unknown>|\[unknown\])"
    r"(?:\s+\([^)]*\))?\s*$",
    re.IGNORECASE,
)
UNKNOWN_ADDRESS_RE = re.compile(
    r"^\s*(?:0x)?[0-9a-fA-F]+\s+"
    r"(?:\?\?|unknown|<unknown>|\[unknown\])"
    r"(?:\s+\([^)]*\))?\s*$",
    re.IGNORECASE,
)
ADDRESS_ONLY_RE = re.compile(
    r"^(?:0x)?[0-9a-f]+(?:\+0x[0-9a-f]+)?$", re.IGNORECASE
)
LOST_RE = re.compile(
    r"lost\s+(?P<first>[\d,]+)\s+samples?"
    r"|(?P<second>[\d,]+)\s+samples?\s+lost"
    r"|samples?\s+lost\s*[:=]\s*(?P<third>[\d,]+)",
    re.IGNORECASE,
)
TRUNCATED_RE = re.compile(
    r"^\s*(?:\.\.\.|<stack\s+truncated>|stack\s+truncated|"
    r"callchain\s+truncated)\s*$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class Sample:
    """One accepted positive ``cycles:u`` sample; frames are leaf first."""

    period: int
    event: str
    symbols: tuple[str, ...]
    unknown_frame_count: int = 0
    unparsed_frame_count: int = 0
    truncated: bool = False
    comm: str | None = None
    pid: int | None = None
    tid: int | None = None
    timestamp: str | None = None


def _is_gzip(path: Path) -> bool:
    return path.suffix.lower() == ".gz"


def _open_binary(path: Path) -> BinaryIO:
    return gzip.open(path, "rb") if _is_gzip(path) else path.open("rb")


def _open_text(path: Path) -> TextIO:
    return gzip.open(path, "rt", encoding="utf-8", errors="replace") if _is_gzip(path) else path.open(
        "r", encoding="utf-8", errors="replace"
    )


def _sha256_stream(path: Path, *, decompressed: bool) -> tuple[str, int]:
    digest = hashlib.sha256()
    stream = _open_binary(path) if decompressed else path.open("rb")
    with stream:
        size = 0
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def sha256_file(path: Path) -> str:
    """Hash the transparent, decompressed input bytes."""

    return _sha256_stream(path, decompressed=True)[0]


def _bundle_relative(path: Path, bundle_root: Path | None) -> tuple[str, bool]:
    if not path.is_absolute():
        candidate = path
        if bundle_root is not None:
            try:
                candidate = path.resolve().relative_to(bundle_root.resolve())
            except ValueError:
                return path.name, True
        return candidate.as_posix(), False
    if bundle_root is not None:
        try:
            return path.resolve().relative_to(bundle_root.resolve()).as_posix(), False
        except ValueError:
            return path.name, True
    return path.name, True


def file_binding(path: Path, bundle_root: Path | None = None) -> dict[str, Any]:
    """Describe an input without embedding an absolute workspace path."""

    resolved = path.resolve()
    relative, outside_bundle = _bundle_relative(path, bundle_root)
    compressed_sha, compressed_bytes = _sha256_stream(resolved, decompressed=False)
    decompressed_sha, decompressed_bytes = _sha256_stream(resolved, decompressed=True)
    return {
        "path": relative,
        "bytes": decompressed_bytes,
        "sha256": decompressed_sha,
        "compression": "gzip" if _is_gzip(resolved) else None,
        "compressed_bytes": compressed_bytes,
        "compressed_sha256": compressed_sha,
        "outside_bundle": outside_bundle,
    }


def normalize_event(event: str) -> str:
    return event.strip().rstrip(":")


def is_unknown_symbol(symbol: str) -> bool:
    value = symbol.strip().lower()
    return (
        value in {"??", "unknown", "<unknown>", "[unknown]"}
        or value.startswith("[unknown ")
        or bool(ADDRESS_ONLY_RE.fullmatch(value))
    )


def symbol_from_frame(line: str) -> str | None:
    """Parse a frame while preserving unresolved addresses as ``[unknown]``."""

    if UNKNOWN_ONLY_RE.match(line) or UNKNOWN_ADDRESS_RE.match(line):
        return "[unknown]"
    if ADDRESS_ONLY_RE.fullmatch(line.strip()):
        return "[unknown]"
    match = FRAME_RE.match(line) or FRAME_RE_NO_OFFSET.match(line)
    if match:
        return match.group("symbol").strip()
    match = FRAME_RE_NO_DSO.match(line)
    if match and match.group("symbol").strip():
        return match.group("symbol").strip()
    return None


def lost_count(line: str) -> int | None:
    match = LOST_RE.search(line)
    if match:
        for name in ("first", "second", "third"):
            value = match.group(name)
            if value is not None:
                return int(value.replace(",", ""))
    return None


def looks_like_header(line: str) -> bool:
    stripped = line.strip()
    if not stripped or stripped.startswith("#"):
        return False
    return "cycles" in stripped.lower() or bool(
        re.search(r"\s+\d+(?:/\d+)?\s+[^:]+:\s+\S+", stripped)
    )


def _header(line: str) -> re.Match[str] | None:
    # Try the explicit separated-pid/tid form first.  The permissive legacy
    # form can otherwise consume the tid as part of its timestamp field.
    return HEADER_SEPARATE_TID_RE.match(line) or HEADER_RE.match(line)


def parse_samples(path: Path) -> tuple[list[Sample], dict[str, Any]]:
    """Parse sample blocks and retain every material coverage denominator.

    A positive, valid cycles:u block is accepted even when one of its frames
    is unresolved or unparsed.  A block ending at EOF, or carrying perf's
    explicit ``...`` stack marker, is accepted with ``truncated=true``.
    Invalid events and periods are excluded from accepted samples but their
    periods and block counts remain in diagnostics.
    """

    samples: list[Sample] = []
    stats: Counter[str] = Counter()
    current_event: str | None = None
    current_period: int | None = None
    current_symbols: list[str] = []
    current_unknown = 0
    current_unparsed = 0
    current_invalid = False
    current_truncated = False
    current_comm: str | None = None
    current_pid: int | None = None
    current_tid: int | None = None
    current_timestamp: str | None = None

    def reset() -> None:
        nonlocal current_event, current_period, current_symbols
        nonlocal current_unknown, current_unparsed, current_invalid, current_truncated
        nonlocal current_comm, current_pid, current_tid, current_timestamp
        current_event = None
        current_period = None
        current_symbols = []
        current_unknown = 0
        current_unparsed = 0
        current_invalid = False
        current_truncated = False
        current_comm = None
        current_pid = None
        current_tid = None
        current_timestamp = None

    def flush(*, truncated: bool = False) -> None:
        nonlocal current_unknown, current_unparsed, current_truncated
        if current_event is None:
            return
        if truncated:
            current_truncated = True
            stats["unterminated_blocks"] += 1
        if current_invalid or current_period is None:
            stats["invalid_sample_blocks"] += 1
            stats["invalid_sample_period"] += current_period or 0
            reset()
            return
        sample = Sample(
            period=current_period,
            event=current_event,
            symbols=tuple(current_symbols),
            unknown_frame_count=current_unknown,
            unparsed_frame_count=current_unparsed,
            truncated=current_truncated,
            comm=current_comm,
            pid=current_pid,
            tid=current_tid,
            timestamp=current_timestamp,
        )
        samples.append(sample)
        stats["sample_blocks_seen"] += 1
        stats["accepted_cycle_period"] += current_period
        stats["parsed_frame_lines"] += len(current_symbols)
        stats["unparsed_frame_lines"] += current_unparsed
        if current_unknown:
            stats["unknown_frame_blocks"] += 1
            stats["unknown_frame_period"] += current_period
            stats["unknown_frame_occurrences"] += current_unknown
            stats["unknown_frame_occurrence_period"] += current_period * current_unknown
            stats["unknown_frame_lines"] += current_unknown
        if current_unparsed:
            stats["unparsed_frame_blocks"] += 1
            stats["unparsed_frame_period"] += current_period
        if not current_symbols:
            stats["empty_stack_blocks"] += 1
            stats["empty_stack_period"] += current_period
        if current_truncated:
            stats["truncated_blocks"] += 1
            stats["truncated_period"] += current_period
        reset()

    with _open_text(path) as source:
        for raw in source:
            stats["input_lines"] += 1
            line = raw.rstrip("\r\n")
            header = _header(line)
            if header:
                if current_event is not None:
                    stats["implicit_block_boundaries"] += 1
                    flush()
                current_comm = header.group("comm")
                current_pid = int(header.group("pid"))
                current_tid = int(header.group("tid") or header.group("pid"))
                current_timestamp = header.group("time").strip()
                current_event = normalize_event(header.group("event"))
                raw_period = header.group("period").replace(",", "")
                try:
                    current_period = int(raw_period)
                except ValueError:
                    current_period = None
                    current_invalid = True
                    stats["malformed_period_headers"] += 1
                if current_event == EXPECTED_EVENT:
                    stats["cycle_headers"] += 1
                    if current_period is not None and current_period <= 0:
                        current_invalid = True
                        stats["invalid_cycle_period_headers"] += 1
                        stats["zero_or_negative_period_samples"] += 1
                else:
                    current_invalid = True
                    stats["non_cycle_headers"] += 1
                    stats["invalid_cycle_event_headers"] += 1
                continue

            if line.strip() == "":
                if current_event is None:
                    stats["blank_lines_outside_blocks"] += 1
                else:
                    stats["block_separator_lines"] += 1
                    flush()
                continue

            count = lost_count(line)
            if count is not None or "lost" in line.lower():
                stats["lost_metadata_lines"] += 1
                if count is None:
                    stats["unquantified_lost_lines"] += 1
                else:
                    stats["lost_sample_count"] += count
                if current_event is not None:
                    stats["lost_metadata_inside_block"] += 1
                continue

            if current_event is None:
                if line.lstrip().startswith("#"):
                    stats["metadata_lines"] += 1
                elif looks_like_header(line):
                    stats["malformed_cycle_headers"] += 1
                else:
                    stats["nonempty_lines_outside_blocks"] += 1
                continue

            if line.lstrip().startswith("#"):
                stats["metadata_lines_inside_blocks"] += 1
                continue
            if TRUNCATED_RE.match(line):
                current_truncated = True
                stats["truncation_markers"] += 1
                continue

            symbol = symbol_from_frame(line)
            if symbol is None:
                current_unparsed += 1
                continue
            current_symbols.append(symbol)
            if is_unknown_symbol(symbol):
                current_unknown += 1

    if current_event is not None:
        flush(truncated=True)
    stats["sample_blocks"] = len(samples)
    stats["total_weighted_event_period"] = sum(sample.period for sample in samples)
    for key in (
        "cycle_headers",
        "non_cycle_headers",
        "sample_blocks_seen",
        "accepted_cycle_period",
        "invalid_cycle_event_headers",
        "invalid_cycle_period_headers",
        "malformed_period_headers",
        "invalid_sample_blocks",
        "invalid_sample_period",
        "zero_or_negative_period_samples",
        "malformed_cycle_headers",
        "parsed_frame_lines",
        "unparsed_frame_lines",
        "unparsed_frame_blocks",
        "unparsed_frame_period",
        "empty_stack_blocks",
        "empty_stack_period",
        "unknown_frame_lines",
        "unknown_frame_blocks",
        "unknown_frame_period",
        "unknown_frame_occurrences",
        "unknown_frame_occurrence_period",
        "lost_sample_count",
        "lost_metadata_lines",
        "unquantified_lost_lines",
        "lost_metadata_inside_block",
        "unterminated_blocks",
        "truncation_markers",
        "truncated_blocks",
        "truncated_period",
    ):
        stats.setdefault(key, 0)
    return samples, dict(stats)


def weighted_metric(
    period: int, blocks: int, denominator: int, whole: int
) -> dict[str, Any]:
    return {
        "weighted_event_period": period,
        "raw_stack_blocks": blocks,
        "share_of_subset_percent": period / denominator * 100.0 if denominator else None,
        "share_of_whole_process_percent": period / whole * 100.0 if whole else None,
    }


def rank_symbols(
    samples: Iterable[Sample],
    predicate: Callable[[Sample], bool],
    subset_denominator: int,
    whole_denominator: int,
    limit: int,
    *,
    leaf: bool,
) -> list[dict[str, Any]]:
    periods: Counter[str] = Counter()
    blocks: Counter[str] = Counter()
    for sample in samples:
        if not predicate(sample):
            continue
        if leaf:
            symbol = sample.symbols[0] if sample.symbols else "<empty-stack>"
            periods[symbol] += sample.period
            blocks[symbol] += 1
        else:
            # Inclusive attribution counts each ancestor once per sample, even
            # when a recursive/repeated frame occurs twice in one callchain.
            for symbol in set(sample.symbols):
                periods[symbol] += sample.period
                blocks[symbol] += 1
    rows = []
    for symbol, period in sorted(
        periods.items(), key=lambda item: (-item[1], item[0])
    )[:limit]:
        rows.append(
            {
                "symbol": symbol,
                **weighted_metric(
                    period, blocks[symbol], subset_denominator, whole_denominator
                ),
            }
        )
    return rows


def marker_matches(symbol: str, literal: str) -> bool:
    """Match a complete Rust symbol marker and closure/generic tails."""

    return (
        symbol == literal
        or symbol.startswith(literal + "::")
        or symbol.startswith(literal + "<")
        or symbol.startswith("<" + literal + ">")
    )


def marker_names(
    symbols: Iterable[str], markers: Mapping[str, Sequence[str]]
) -> set[str]:
    found: set[str] = set()
    for category, literals in markers.items():
        if any(
            marker_matches(symbol, literal)
            for symbol in symbols
            for literal in literals
        ):
            found.add(category)
    return found


def marker_hits(
    symbols: Iterable[str], markers: Mapping[str, Sequence[str]]
) -> list[tuple[int, str, str]]:
    hits: list[tuple[int, str, str]] = []
    for index, symbol in enumerate(symbols):
        for category in sorted(marker_names((symbol,), markers)):
            hits.append((index, category, symbol))
    return hits


def classify_scope(
    symbols: Iterable[str], markers: Mapping[str, Sequence[str]]
) -> str:
    hits = marker_hits(symbols, markers)
    if not hits:
        return "unclassified"
    nearest = min(index for index, _category, _symbol in hits)
    categories = {
        category for index, category, _symbol in hits if index == nearest
    }
    if len(categories) == 1:
        return next(iter(categories))
    return "overlap"


def classify_context_scope(
    symbols: Iterable[str], markers: Mapping[str, Sequence[str]]
) -> str:
    """Classify writer/preflight contexts with an explicit writer priority.

    ``write_pptx_stream`` can be inlined into ``run`` and a future profile
    may retain both a setup marker and the enclosing run marker.  The exact
    public operation marker owns such a sample; it must not be relabeled as
    materialized preflight merely because another marker is present.
    """

    found = marker_names(symbols, markers)
    if "writer_under_run" in found:
        return "writer_under_run"
    return classify_scope(symbols, markers)


def _ranked_scope(
    samples: Sequence[Sample],
    whole: int,
    top: int,
    markers: Mapping[str, Sequence[str]] | None = None,
) -> dict[str, Any]:
    period = sum(sample.period for sample in samples)
    return {
        "scope": weighted_metric(period, len(samples), period, whole),
        "leaf_period_weighted_ranking": rank_symbols(
            samples, lambda _sample: True, period, whole, top, leaf=True
        ),
        "inclusive_period_weighted_ranking": rank_symbols(
            samples, lambda _sample: True, period, whole, top, leaf=False
        ),
        "unknown_frame_blocks": sum(
            sample.unknown_frame_count > 0 for sample in samples
        ),
        "unknown_frame_period": sum(
            sample.period for sample in samples if sample.unknown_frame_count
        ),
        "unparsed_frame_blocks": sum(
            sample.unparsed_frame_count > 0 for sample in samples
        ),
        "unparsed_frame_period": sum(
            sample.period for sample in samples if sample.unparsed_frame_count
        ),
        "truncated_blocks": sum(sample.truncated for sample in samples),
        "truncated_period": sum(sample.period for sample in samples if sample.truncated),
        "observed_marker_symbols": sorted(
            {
                symbol
                for sample in samples
                for symbol in sample.symbols
                if markers is None or marker_names((symbol,), markers)
            }
        ),
    }


def context_summaries(
    samples: Sequence[Sample],
    top: int,
    markers: Mapping[str, Sequence[str]],
) -> dict[str, Any]:
    """Build a disjoint nearest-marker partition for the writer phases."""

    whole = sum(sample.period for sample in samples)
    categories = tuple(dict.fromkeys((*markers.keys(), "overlap", "unclassified")))
    grouped: dict[str, list[Sample]] = {category: [] for category in categories}
    mixed_blocks = 0
    mixed_period = 0
    mixed_categories: Counter[str] = Counter()
    for sample in samples:
        found = marker_names(sample.symbols, markers)
        if len(found) > 1:
            mixed_blocks += 1
            mixed_period += sample.period
            mixed_categories["+".join(sorted(found))] += sample.period
        grouped.setdefault(classify_context_scope(sample.symbols, markers), []).append(sample)

    rows: dict[str, Any] = {}
    partition_period = 0
    partition_blocks = 0
    for category in categories:
        group = grouped[category]
        summary = _ranked_scope(group, whole, top, markers)
        rows[category] = summary
        partition_period += summary["scope"]["weighted_event_period"]
        partition_blocks += summary["scope"]["raw_stack_blocks"]
    return {
        "classification": (
            "one disjoint nearest-marker label per accepted sample block; same-frame "
            "marker ambiguity is retained as overlap"
        ),
        "categories": rows,
        "mixed_marker_stacks": {
            "raw_stack_blocks": mixed_blocks,
            "weighted_event_period": mixed_period,
            "nearest_marker_resolution": True,
            "period_by_marker_set": dict(mixed_categories),
        },
        "partition": {
            "whole_process_weighted_event_period": whole,
            "whole_process_raw_stack_blocks": len(samples),
            "classified_weighted_event_period": partition_period,
            "classified_raw_stack_blocks": partition_blocks,
            "periods_partition_exact": partition_period == whole,
            "blocks_partition_exact": partition_blocks == len(samples),
        },
        "marker_definitions": {key: list(value) for key, value in markers.items()},
    }


def _family_match(family: str, symbol: str, literal: str) -> bool:
    # Owner families are intentionally substring rules because perf demangles
    # generic instantiations as ``<...Type...>`` around the source path.
    if family == "deflate_initialization" and literal in {
        "flate2::deflate::write::DeflateEncoder",
        "soapberry_zip::writer::OwnedCompressor",
        "flate2::mem::Compress",
    }:
        # A generic DeflateEncoder appears in both construction and writes;
        # only the constructor belongs to this family.  The profile's
        # demangled spelling may put the generic arguments between the type
        # name and ``>::new``.
        return literal in symbol and ("::new" in symbol or ">::new" in symbol)
    if family == "deflate_compression" and literal == "flate2::mem::Compress":
        return literal in symbol and "::new" not in symbol and ">::new" not in symbol
    return literal in symbol


def owner_family_summaries(
    samples: Sequence[Sample],
    top: int,
    families: Mapping[str, Sequence[str]],
) -> dict[str, Any]:
    """Rank independent implementation-owner families.

    A sample may belong to several families (for example a deflate write
    under a ZIP entry writer).  These rows deliberately overlap and must not
    be added to estimate a total.
    """

    whole = sum(sample.period for sample in samples)
    rows: dict[str, Any] = {}
    for family, literals in families.items():
        group = [
            sample
            for sample in samples
            if any(
                _family_match(family, symbol, literal)
                for symbol in sample.symbols
                for literal in literals
            )
        ]
        period = sum(sample.period for sample in group)
        rows[family] = {
            "scope": weighted_metric(period, len(group), period, whole),
            "leaf_period_weighted_ranking": rank_symbols(
                group, lambda _sample: True, period, whole, top, leaf=True
            ),
            "inclusive_period_weighted_ranking": rank_symbols(
                group, lambda _sample: True, period, whole, top, leaf=False
            ),
            "observed_marker_symbols": sorted(
                {
                    symbol
                    for sample in group
                    for symbol in sample.symbols
                    if any(
                        _family_match(family, symbol, literal) for literal in literals
                    )
                }
            ),
            "marker_literals": list(literals),
            "unknown_frame_blocks": sum(
                sample.unknown_frame_count > 0 for sample in group
            ),
            "unknown_frame_period": sum(
                sample.period for sample in group if sample.unknown_frame_count
            ),
            "unparsed_frame_blocks": sum(
                sample.unparsed_frame_count > 0 for sample in group
            ),
            "unparsed_frame_period": sum(
                sample.period for sample in group if sample.unparsed_frame_count
            ),
            "truncated_blocks": sum(sample.truncated for sample in group),
            "truncated_period": sum(sample.period for sample in group if sample.truncated),
        }
    return {
        "scope": (
            "independent inclusive owner-family attribution; rows may overlap "
            "within one callchain and their periods must not be summed"
        ),
        "families": rows,
        "overlap_allowed": True,
        "period_sum_is_not_a_partition": True,
    }


def coverage(stats: Mapping[str, Any], samples: Sequence[Sample]) -> dict[str, Any]:
    whole = sum(sample.period for sample in samples)
    resolved_period = sum(
        sample.period
        for sample in samples
        if any(not is_unknown_symbol(symbol) for symbol in sample.symbols)
    )
    truncated_period = sum(sample.period for sample in samples if sample.truncated)
    return {
        "accepted_sample_blocks": len(samples),
        "accepted_weighted_event_period": whole,
        "resolved_frame_weighted_event_period": resolved_period,
        "known_frame_weighted_event_period": resolved_period,
        "unknown_frame_weighted_event_period": stats.get("unknown_frame_period", 0),
        "unparsed_frame_weighted_event_period": stats.get("unparsed_frame_period", 0),
        "empty_stack_weighted_event_period": stats.get("empty_stack_period", 0),
        "truncated_weighted_event_period": truncated_period,
        "unknown_frame_share_of_whole_process_percent": (
            stats.get("unknown_frame_period", 0) / whole * 100.0 if whole else None
        ),
        "unparsed_frame_share_of_whole_process_percent": (
            stats.get("unparsed_frame_period", 0) / whole * 100.0 if whole else None
        ),
        "empty_stack_share_of_whole_process_percent": (
            stats.get("empty_stack_period", 0) / whole * 100.0 if whole else None
        ),
        "truncated_share_of_whole_process_percent": (
            truncated_period / whole * 100.0 if whole else None
        ),
        "coverage_rows_may_overlap": True,
        "lost_sample_count": stats.get("lost_sample_count", 0),
        "lost_period_is_unavailable": True,
        "invalid_cycle_period": stats.get("invalid_sample_period", 0),
        "parser_diagnostics": dict(stats),
    }


def _report_summary(path: Path, bundle_root: Path | None = None) -> dict[str, Any]:
    try:
        with _open_text(path) as stream:
            data = json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        return {
            "file": file_binding(path, bundle_root),
            "status": "invalid",
            "error": str(error),
        }
    if not isinstance(data, dict):
        return {
            "file": file_binding(path, bundle_root),
            "status": "invalid",
            "error": "report is not an object",
        }
    results = data.get("results")
    rows = results if isinstance(results, list) else []
    elapsed_statistics = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        elapsed = row.get("elapsed_ns")
        if isinstance(elapsed, dict) and isinstance(elapsed.get("samples"), list):
            values = elapsed["samples"]
            elapsed_statistics.append(
                {
                    "case": row.get("case"),
                    "sample_count": len(values),
                    "min_ns": elapsed.get("min"),
                    "p50_ns": elapsed.get("p50"),
                    "p95_ns": elapsed.get("p95"),
                    "p99_ns": elapsed.get("p99"),
                    "max_ns": elapsed.get("max"),
                    "mean_ns": elapsed.get("mean"),
                }
            )
    return {
        "file": file_binding(path, bundle_root),
        "status": "parsed",
        "schema_version": data.get("schema_version"),
        "tool": data.get("tool"),
        "binary_identity": data.get("binary_identity"),
        "environment": data.get("environment"),
        "configuration": data.get("configuration"),
        "results": {
            "count": len(rows),
            "cases": [
                row.get("case")
                for row in rows
                if isinstance(row, dict) and isinstance(row.get("case"), str)
            ],
            "elapsed_statistics": elapsed_statistics,
        },
        "timing_use": (
            "elapsed values are retained as workload context only; perf sample "
            "period is never converted to phase time"
        ),
    }
