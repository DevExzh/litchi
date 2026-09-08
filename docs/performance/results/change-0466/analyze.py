#!/usr/bin/env python3
"""Analyze the retained 0466 XLSX ``perf script`` export.

The input is a whole-process ``cycles:u`` sample stream.  A sample period is
an event-period weight, never a wall-clock duration.  The parser keeps
unknown and malformed coverage visible, and all inclusive rankings de-duplicate
an ancestor that appears more than once in one stack.

This helper intentionally does no symbolization and never invokes ``perf``.
It consumes the text export made by the capture postprocessor and optional
normal reports only.  The report data supplies workload identity and elapsed
sample context; it does not turn sampled periods into phase timers.
"""

from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import dataclass
import gzip
import hashlib
import json
from pathlib import Path
import re
import sys
from typing import Any, BinaryIO, Callable, Iterable, Mapping, Sequence, TextIO


SCHEMA = "litchi-0466-xlsx-cpu-attribution-v1"
EXPECTED_EVENT = "cycles:u"
DEFAULT_TOP = 40

# perf prints pid/tid as ``pid/tid`` when both are requested.  Older perf
# versions print them as two adjacent fields, so retain both spellings.
HEADER_RE = re.compile(
    r"^\s*(?P<comm>\S+)\s+(?P<pid>\d+)(?:/(?P<tid>\d+))?\s+"
    r"(?P<time>[^:]+):\s+(?P<period>[+-]?[\d,]+)\s+(?P<event>\S+)\s*$"
)
HEADER_SEPARATE_TID_RE = re.compile(
    r"^\s*(?P<comm>\S+)\s+(?P<pid>\d+)\s+(?P<tid>\d+)\s+"
    r"(?P<time>[^:]+):\s+(?P<period>[+-]?[\d,]+)\s+(?P<event>\S+)\s*$"
)
FRAME_RE = re.compile(
    r"^\s*(?:0x)?[0-9a-fA-F]+\s+(?P<symbol>.+?)"
    r"\+0x[0-9a-fA-F]+\s+\([^)]*\)\s*$"
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
# A symbol that is only an address is an unresolved frame in a few perf
# exports.  Keep it in the unknown denominator instead of presenting it as a
# real symbol (notably the short 0x7fff frames in the original capture).
ADDRESS_ONLY_RE = re.compile(
    r"^(?:0x)?[0-9a-f]+(?:\+0x[0-9a-f]+)?$", re.IGNORECASE
)
LOST_RE = re.compile(
    r"lost\s+(?P<first>[\d,]+)\s+samples?"
    r"|(?P<second>[\d,]+)\s+samples?\s+lost"
    r"|samples?\s+lost\s*[:=]\s*(?P<third>[\d,]+)",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class Sample:
    """One accepted positive ``cycles:u`` sample, leaf frame first."""

    period: int
    event: str
    symbols: tuple[str, ...]
    unknown_frame_count: int = 0
    unparsed_frame_count: int = 0
    truncated: bool = False


def _is_gzip(path: Path) -> bool:
    return path.suffix.lower() == ".gz"


def _open_binary(path: Path) -> BinaryIO:
    return gzip.open(path, "rb") if _is_gzip(path) else path.open("rb")


def _open_text(path: Path) -> TextIO:
    if _is_gzip(path):
        return gzip.open(path, "rt", encoding="utf-8", errors="replace")
    return path.open("r", encoding="utf-8", errors="replace")


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
    """Hash the artifact bytes after transparent gzip decompression."""

    return _sha256_stream(path, decompressed=True)[0]


def _bundle_relative(path: Path, bundle_root: Path | None) -> tuple[str, bool]:
    """Return a stable path identity without leaking an absolute workspace path."""

    if not path.is_absolute():
        candidate = path
        if bundle_root is not None:
            try:
                candidate = path.resolve().relative_to(bundle_root.resolve())
            except ValueError:
                # A caller may intentionally analyze a temporary fixture.  It
                # remains portable as a basename, with the flag documenting
                # that it was outside the evidence bundle.
                return path.name, True
        return candidate.as_posix(), False
    if bundle_root is not None:
        try:
            return path.resolve().relative_to(bundle_root.resolve()).as_posix(), False
        except ValueError:
            return path.name, True
    return path.name, True


def file_binding(path: Path, bundle_root: Path | None = None) -> dict[str, Any]:
    """Describe an artifact by bundle-relative path and decompressed identity."""

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
    """Return a demangled frame symbol, preserving unknown frames."""

    if UNKNOWN_ONLY_RE.match(line) or UNKNOWN_ADDRESS_RE.match(line):
        return "[unknown]"
    if ADDRESS_ONLY_RE.fullmatch(line.strip()):
        return "[unknown]"
    match = FRAME_RE.match(line) or FRAME_RE_NO_OFFSET.match(line)
    if match:
        return match.group("symbol").strip()
    # A no-DSO frame is emitted by some perf versions when the dso is absent.
    # Require an address and at least one non-space symbol character so random
    # metadata cannot silently become a frame.
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
    return HEADER_RE.match(line) or HEADER_SEPARATE_TID_RE.match(line)


def parse_samples(path: Path) -> tuple[list[Sample], dict[str, Any]]:
    """Parse perf blocks and retain every coverage/error denominator.

    Invalid event and period blocks are excluded from the returned samples but
    their periods and counts remain in the diagnostics.  An EOF-terminated
    positive block is retained and marked ``truncated``; this avoids silently
    dropping its weight while making the incomplete boundary visible.
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

    def reset() -> None:
        nonlocal current_event, current_period, current_symbols
        nonlocal current_unknown, current_unparsed, current_invalid
        nonlocal current_truncated
        current_event = None
        current_period = None
        current_symbols = []
        current_unknown = 0
        current_unparsed = 0
        current_invalid = False
        current_truncated = False

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
        )
        samples.append(sample)
        stats["sample_blocks_seen"] += 1
        stats["accepted_cycle_period"] += current_period
        stats["parsed_frame_lines"] += len(current_symbols)
        stats["unparsed_frame_lines"] += current_unparsed
        if current_period <= 0:
            stats["zero_or_negative_period_samples"] += 1
        if not current_symbols:
            stats["empty_stack_blocks"] += 1
            stats["empty_stack_period"] += current_period
        if current_unknown:
            stats["unknown_frame_blocks"] += 1
            stats["unknown_frame_period"] += current_period
            stats["unknown_frame_occurrences"] += current_unknown
            stats["unknown_frame_occurrence_period"] += current_period * current_unknown
            stats["unknown_frame_lines"] += current_unknown
        if current_unparsed:
            stats["unparsed_frame_blocks"] += 1
            stats["unparsed_frame_period"] += current_period
        if current_truncated:
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
                current_event = normalize_event(header.group("event"))
                current_period = int(header.group("period").replace(",", ""))
                if current_event == EXPECTED_EVENT:
                    stats["cycle_headers"] += 1
                    if current_period <= 0:
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
    # Stable zero fields make the summary and negative probes easy to audit.
    for key in (
        "cycle_headers",
        "non_cycle_headers",
        "sample_blocks_seen",
        "accepted_cycle_period",
        "invalid_cycle_event_headers",
        "invalid_cycle_period_headers",
        "invalid_sample_blocks",
        "invalid_sample_period",
        "zero_or_negative_period_samples",
        "malformed_cycle_headers",
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
            # A recursive or repeated ancestor is one inclusive frame in a
            # stack block.  Counting the list directly would inflate both its
            # period and its denominator.
            for symbol in set(sample.symbols):
                periods[symbol] += sample.period
                blocks[symbol] += 1
    rows = []
    for symbol, period in sorted(periods.items(), key=lambda item: (-item[1], item[0]))[:limit]:
        rows.append(
            {
                "symbol": symbol,
                **weighted_metric(period, blocks[symbol], subset_denominator, whole_denominator),
            }
        )
    return rows


# These are observed demangled symbols in the no-inline export.  The aliases
# keep the scope report useful when a method is inlined into the crate writer.
MARKERS: dict[str, tuple[str, ...]] = {
    "edit_commit": (
        "<litchi_xlsx::workbook::edit::semantic::transaction::Edit>::commit",
        "litchi_xlsx::workbook::edit::semantic::transaction::Edit::commit",
    ),
    "workbook_write_to": (
        "<litchi_xlsx::workbook::model::Workbook>::write_to",
        "litchi_xlsx::workbook::model::Workbook::write_to",
        "litchi_xlsx::writer::write_to",
    ),
    # The release/profile build inlines the Workbook::write_to wrapper.  This
    # is the retained sequential writer boundary, so report it separately
    # rather than silently relabeling it as the absent Workbook method.
    "writer_write_to_stream": (
        "<litchi_opc::pkgwriter::PackageWriter>::write_to_stream::<&mut litchi_perf_baseline::CountingSink>",
    ),
    "prepare_open_or_oracle": (
        "litchi_perf_baseline::prepare_xlsx_updates",
        "litchi_perf_baseline::xlsx_expected_output",
        "litchi_perf_baseline::verify_xlsx_cells",
        "litchi_xlsx::workbook::model::Workbook::from_bytes",
        "litchi_xlsx::workbook::model::Workbook::from_bytes_with_limits",
        "<litchi_xlsx::workbook::model::Workbook>::from_bytes_with_limits",
        "<litchi_xlsx::workbook::model::Workbook>::from_package_with_styles",
        "litchi_xlsx::workbook::model::Workbook::from_package_with_styles",
        "<litchi_xlsx::workbook::model::Workbook>::from_bytes",
    ),
}


def marker_matches(symbol: str, literal: str) -> bool:
    """Match a complete method marker, allowing Rust generic/closure tails."""

    return symbol == literal or symbol.startswith(literal + "::") or symbol.startswith(
        literal + "<"
    ) or (literal.startswith("<") and symbol.startswith(literal + "::"))


def marker_names(symbols: Iterable[str], markers: Mapping[str, Sequence[str]] = MARKERS) -> set[str]:
    found: set[str] = set()
    for category, literals in markers.items():
        if any(marker_matches(symbol, literal) for symbol in symbols for literal in literals):
            found.add(category)
    return found


def marker_hits(
    symbols: Iterable[str], markers: Mapping[str, Sequence[str]] = MARKERS
) -> list[tuple[int, str, str]]:
    """Return marker hits as ``(leaf_distance, category, symbol)`` rows."""

    hits: list[tuple[int, str, str]] = []
    for index, symbol in enumerate(symbols):
        for category in sorted(marker_names((symbol,), markers)):
            hits.append((index, category, symbol))
    return hits


def classify_scope(
    symbols: Iterable[str], markers: Mapping[str, Sequence[str]] = MARKERS
) -> str:
    """Return one scope using the nearest (most specific) marker frame.

    ``perf`` callchains are leaf first.  A marker nearer the leaf is the
    operation that owns the sampled work when an ancestor such as
    ``from_package_with_styles`` also appears in the same stack.  A genuine
    same-frame ambiguity remains ``overlap`` so it cannot be silently lost.
    """

    hits = marker_hits(symbols, markers)
    if not hits:
        return "unclassified"
    nearest = min(index for index, _category, _symbol in hits)
    nearest_categories = {
        category for index, category, _symbol in hits if index == nearest
    }
    if len(nearest_categories) == 1:
        return next(iter(nearest_categories))
    if len(nearest_categories) > 1:
        return "overlap"
    return "unclassified"


def scope_summaries(
    samples: Sequence[Sample], top: int, markers: Mapping[str, Sequence[str]] = MARKERS
) -> dict[str, Any]:
    whole = sum(sample.period for sample in samples)
    categories = tuple(dict.fromkeys((*markers.keys(), "overlap", "unclassified")))
    grouped: dict[str, list[Sample]] = {category: [] for category in categories}
    mixed_blocks = 0
    mixed_period = 0
    mixed_categories: Counter[str] = Counter()
    for sample in samples:
        hits = marker_hits(sample.symbols, markers)
        found = {category for _index, category, _symbol in hits}
        if len(found) > 1:
            mixed_blocks += 1
            mixed_period += sample.period
            mixed_categories["+".join(sorted(found))] += sample.period
        grouped.setdefault(classify_scope(sample.symbols, markers), []).append(sample)
    rows: dict[str, Any] = {}
    partition_period = 0
    partition_blocks = 0
    for category in categories:
        group = grouped[category]
        period = sum(sample.period for sample in group)
        partition_period += period
        partition_blocks += len(group)
        rows[category] = {
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
                    if marker_names((symbol,), markers)
                }
            ),
        }
    return {
        "classification": (
            "one disjoint label per accepted sample block; the nearest marker in the "
            "leaf-to-root callchain wins, while same-frame ambiguity is overlap"
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
        "category_descriptions": {
            "edit_commit": "exact litchi_xlsx Edit::commit marker",
            "workbook_write_to": (
                "Workbook::write_to aliases; zero means this wrapper was not retained"
            ),
            "writer_write_to_stream": (
                "PackageWriter::write_to_stream retained writer boundary; the "
                "Workbook::write_to wrapper may be inlined"
            ),
            "prepare_open_or_oracle": (
                "setup, open/from_bytes, expected-output, or verification markers"
            ),
            "overlap": "same-frame marker ambiguity retained explicitly",
            "unclassified": "accepted sample block with no configured marker",
        },
        "marker_definitions": {key: list(value) for key, value in markers.items()},
    }


def exact_contexts(samples: Sequence[Sample]) -> dict[str, Any]:
    """Inclusive method evidence, independent of nearest-marker classification."""
    commit = "<litchi_xlsx::workbook::edit::semantic::transaction::Edit>::commit"
    writer = MARKERS["writer_write_to_stream"][0]
    parser = "<litchi_xlsx::raw::worksheet::model::Parser>::parse"
    attribute = "litchi_ooxml_common::xml::unqualified_attribute_value"
    whole = sum(sample.period for sample in samples)
    commits = [sample for sample in samples if commit in sample.symbols]
    commit_weight = sum(sample.period for sample in commits)
    groups = {
        "commit": (commits, whole),
        "counting_sink_writer": ([s for s in samples if writer in s.symbols], whole),
        "worksheet_parser_under_commit": (
            [s for s in commits if any(marker_matches(v, parser) for v in s.symbols)], commit_weight),
        "attribute_lookup_under_commit": (
            [s for s in commits if attribute in s.symbols], commit_weight),
    }
    return {
        "scope": "whole-process exact ancestor contexts; inclusive rows overlap; not elapsed phases",
        "markers": dict(commit=commit, counting_sink_writer=writer, worksheet_parser=parser, attribute_lookup=attribute),
        "rows": {name: weighted_metric(sum(s.period for s in group), len(group), denominator, whole)
                 for name, (group, denominator) in groups.items()},
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
    config = data.get("configuration")
    results = data.get("results")
    rows = results if isinstance(results, list) else []
    cases = [row.get("case") for row in rows if isinstance(row, dict) and isinstance(row.get("case"), str)]
    elapsed_counts = []
    elapsed_statistics = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        elapsed = row.get("elapsed_ns")
        if isinstance(elapsed, dict) and isinstance(elapsed.get("samples"), list):
            values = elapsed["samples"]
            elapsed_counts.append(len(values))
            # Preserve the report's recorded timer statistics as context for
            # the profile.  These are workload elapsed samples, not sampled
            # phase durations and not operation counters.
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
        "configuration": config,
        "results": {
            "count": len(rows),
            "cases": cases,
            "elapsed_sample_counts": elapsed_counts,
            "elapsed_statistics": elapsed_statistics,
        },
    }


def _coverage(stats: Mapping[str, Any], samples: Sequence[Sample]) -> dict[str, Any]:
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
        "unknown_frame_share_of_whole_process_percent": stats.get("unknown_frame_period", 0) / whole * 100.0 if whole else None,
        "unparsed_frame_share_of_whole_process_percent": stats.get("unparsed_frame_period", 0) / whole * 100.0 if whole else None,
        "empty_stack_share_of_whole_process_percent": stats.get("empty_stack_period", 0) / whole * 100.0 if whole else None,
        "truncated_share_of_whole_process_percent": truncated_period / whole * 100.0 if whole else None,
        "coverage_rows_may_overlap": True,
        "lost_sample_count": stats.get("lost_sample_count", 0),
        "lost_period_is_unavailable": True,
        "invalid_cycle_period": stats.get("invalid_sample_period", 0),
        "parser_diagnostics": dict(stats),
    }


def analyze_script(
    script: Path,
    reports: Sequence[Path] = (),
    *,
    top: int = DEFAULT_TOP,
    markers: Mapping[str, Sequence[str]] = MARKERS,
    bundle_root: Path | None = None,
) -> dict[str, Any]:
    if bundle_root is None:
        bundle_root = Path(__file__).resolve().parent
    samples, stats = parse_samples(script)
    if not samples:
        raise ValueError("no positive cycles:u sample blocks parsed")
    whole = sum(sample.period for sample in samples)
    scope = scope_summaries(samples, top, markers)
    return {
        "schema": SCHEMA,
        "purpose": "descriptive whole-process XLSX CPU attribution; no phase-time or operation-counter claim",
        "timing_semantics": {
            "weight": "cycles:u perf sample period",
            "phase_latency": False,
            "operation_counter": False,
            "callchain_order": "perf frames are leaf to caller/root",
            "inclusive_overlap": "inclusive rows de-duplicate repeated ancestors within a stack but overlap across rows",
            "normal_report_elapsed": (
                "recorded harness elapsed samples are contextual inputs only; they are "
                "not converted into phase time or operation counters"
            ),
        },
        "inputs": {
            "bundle_root": ".",
            "perf_script": file_binding(script, bundle_root),
            "normal_reports": [_report_summary(path, bundle_root) for path in reports],
            "parser": file_binding(Path(__file__).resolve(), bundle_root),
        },
        "sample_parser": _coverage(stats, samples),
        "whole_process": {
            "weighted_event_period": whole,
            "raw_stack_blocks": len(samples),
            "leaf_period_weighted_ranking": rank_symbols(
                samples, lambda _sample: True, whole, whole, top, leaf=True
            ),
            "inclusive_period_weighted_ranking": rank_symbols(
                samples, lambda _sample: True, whole, whole, top, leaf=False
            ),
        },
        "disjoint_scopes": scope,
        "exact_contexts": exact_contexts(samples),
        "limitations": [
            "Sample periods are weighted cycles:u event periods, not wall time or elapsed phase duration.",
            "The whole-process profile includes setup, warmups, retained workload, verification/oracle work, child processes, and teardown whenever sampled.",
            "A missing marker means no matching frame was retained; it does not prove zero work.",
            "Unknown, malformed, empty, unparsed, truncated, and lost coverage remain explicit denominators.",
            "Inclusive scope rows are disjoint by sample classification, while rankings within each scope remain inclusive frame attribution.",
        ],
    }


def _load_markers(path: Path | None) -> Mapping[str, Sequence[str]]:
    if path is None:
        return MARKERS
    value = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(value, dict) and isinstance(value.get("markers"), dict):
        value = value["markers"]
    if not isinstance(value, dict):
        raise ValueError("source anchors must be an object or an object with a markers field")
    result: dict[str, tuple[str, ...]] = {}
    for name, literals in value.items():
        if isinstance(literals, str):
            result[str(name)] = (literals,)
        elif isinstance(literals, list) and all(isinstance(item, str) and item for item in literals):
            result[str(name)] = tuple(literals)
    if not result:
        raise ValueError("source anchors contain no marker literals")
    return result


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--script", type=Path, required=True)
    parser.add_argument("--report", type=Path, action="append", default=[])
    parser.add_argument("--source-anchors", type=Path, default=None)
    parser.add_argument("--top", type=int, default=DEFAULT_TOP)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    if args.top <= 0:
        raise SystemExit("--top must be positive")
    markers = _load_markers(args.source_anchors)
    bundle_root = Path(__file__).resolve().parent
    summary = analyze_script(args.script, args.report, top=args.top, markers=markers, bundle_root=bundle_root)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({
        "output": str(args.output.resolve()),
        "whole_process_blocks": summary["whole_process"]["raw_stack_blocks"],
        "whole_process_period": summary["whole_process"]["weighted_event_period"],
        "scope_partition_exact": summary["disjoint_scopes"]["partition"]["periods_partition_exact"],
        "unknown_frame_period": summary["sample_parser"]["unknown_frame_weighted_event_period"],
        "truncated_period": summary["sample_parser"]["truncated_weighted_event_period"],
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
