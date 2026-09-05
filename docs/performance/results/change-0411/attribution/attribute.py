#!/usr/bin/env python3
"""Attribute bounded XLS CPU captures produced with ``perf script --no-inline``.

This is a read-only diagnostic postprocessor for the 0411 XLS captures.  It
keeps perf's sample period as a weighted event-period value.  The value is
useful for ranking sampled CPU stacks, but it is neither elapsed time nor an
operation-only latency measurement.

The parser intentionally accepts an imperfect ``perf script`` export.  It
retains unknown frames, counts malformed headers and unparsed lines, and
reports those counts with every result.  Production subsets are selected by
an observed demangled ancestor symbol, rather than by treating every
``run_xls_source_backed_case`` descendant as part of the timer.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shlex
import subprocess
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable


SCHEMA = "0411-xls-cpu-attribution-v1"
DEFAULT_ROOT = Path("/home/zhuhe/code/litchi")
DEFAULT_CAPTURE = Path("/tmp/litchi-goal-0411-capture")
DEFAULT_ATTRIBUTION = Path("/tmp/litchi-goal-0411-attribution")
DEFAULT_SCRIPT = DEFAULT_ATTRIBUTION / "perf-script-no-inline.stdout"
DEFAULT_REPORT = DEFAULT_ATTRIBUTION / "perf-all-self.stdout"

EXPECTED_TOOLCHAIN = "1.98.1"
EXPECTED_EVENT = "cycles:u"
EXPECTED_FREQUENCY = 999
EXPECTED_SAMPLES = 1000
EXPECTED_WARMUP = 20

# ``perf script`` emits a header such as:
#   litchi-perf-bas 1292060 126069.5: 1 cycles:u:
# The event can have one or more trailing colons depending on perf version.
HEADER_RE = re.compile(
    r"^\s*\S+\s+\d+\s+[^:]+:\s+(?P<period>\d+)\s+(?P<event>\S+)\s*$"
)
FRAME_RE = re.compile(
    r"^\s*[0-9a-fA-F]+\s+(?P<symbol>.+?)\+0x[0-9a-fA-F]+\s+\([^)]*\)\s*$"
)
FRAME_RE_NO_OFFSET = re.compile(
    r"^\s*[0-9a-fA-F]+\s+(?P<symbol>.+?)\s+\([^)]*\)\s*$"
)
CASE_RE = re.compile(r"\bxls_(?:eager|source_backed|semantic)_\w+\b")
HARNESS_PARENT_LITERAL = "litchi_perf_baseline::run_xls_source_backed_case"


@dataclass(frozen=True)
class Sample:
    """One parsed perf sample; symbols are in perf's leaf-to-root order."""

    period: int
    symbols: tuple[str, ...]


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_binding(path: Path, root: Path | None = None) -> dict[str, Any]:
    resolved = path.resolve()
    item: dict[str, Any] = {
        "path": str(resolved),
        "bytes": resolved.stat().st_size,
        "sha256": sha256_file(resolved),
    }
    if root is not None:
        try:
            item["relative_to_repo"] = str(resolved.relative_to(root.resolve()))
        except ValueError:
            pass
    return item


def normalize_event(event: str) -> str:
    return event.rstrip(":")


def symbol_from_frame(line: str) -> str | None:
    """Return the demangled frame symbol, or None when the line is not a frame."""

    match = FRAME_RE.match(line)
    if match:
        return match.group("symbol").strip()
    match = FRAME_RE_NO_OFFSET.match(line)
    if match:
        return match.group("symbol").strip()
    return None


def parse_samples(path: Path) -> tuple[list[Sample], dict[str, Any]]:
    """Parse cycle samples while retaining diagnostics for malformed input."""

    samples: list[Sample] = []
    stats: Counter[str] = Counter()
    current_period: int | None = None
    current_event: str | None = None
    current_symbols: list[str] = []

    def flush() -> None:
        nonlocal current_period, current_event, current_symbols
        if current_period is None:
            return
        stats["sample_blocks_seen"] += 1
        if current_period <= 0:
            stats["zero_or_negative_period_samples"] += 1
        if not current_symbols:
            stats["empty_stack_blocks"] += 1
        samples.append(Sample(current_period, tuple(current_symbols)))
        current_period = None
        current_event = None
        current_symbols = []

    with path.open("r", encoding="utf-8", errors="replace") as source:
        for raw in source:
            stats["input_lines"] += 1
            line = raw.rstrip("\r\n")
            match = HEADER_RE.match(line)
            if match:
                flush()
                event = normalize_event(match.group("event"))
                if event.startswith("cycles"):
                    current_period = int(match.group("period"))
                    current_event = event
                    stats["cycle_headers"] += 1
                else:
                    stats["non_cycle_headers"] += 1
                    # Leave this block inactive; its frames are metadata from
                    # another event and must not enter the cycle accounting.
                    current_period = None
                    current_event = None
                continue

            if line.strip() == "":
                if current_period is not None:
                    stats["blank_lines_inside_cycle_blocks"] += 1
                    # A perf stack block is terminated by its blank separator.
                    # Flushing here makes malformed text after that separator
                    # visible as outside-block metadata instead of silently
                    # attaching it to the preceding sample.
                    flush()
                continue

            if current_period is None:
                stats["nonempty_lines_outside_cycle_blocks"] += 1
                if "cycles" in line:
                    stats["malformed_cycle_headers"] += 1
                continue

            stats["nonempty_lines_inside_cycle_blocks"] += 1
            symbol = symbol_from_frame(line)
            if symbol is None:
                stats["unparsed_frame_lines"] += 1
                continue
            stats["parsed_frame_lines"] += 1
            if is_unknown_symbol(symbol):
                stats["unknown_frame_lines"] += 1
            current_symbols.append(symbol)

    flush()
    stats["sample_blocks"] = len(samples)
    stats["total_weighted_event_period"] = sum(sample.period for sample in samples)
    stats.setdefault("cycle_headers", 0)
    stats.setdefault("sample_blocks_seen", 0)
    stats.setdefault("empty_stack_blocks", 0)
    stats.setdefault("unknown_frame_lines", 0)
    stats.setdefault("unparsed_frame_lines", 0)
    stats.setdefault("malformed_cycle_headers", 0)
    stats.setdefault("non_cycle_headers", 0)
    return samples, dict(stats)


def is_unknown_symbol(symbol: str) -> bool:
    normalized = symbol.strip().lower()
    return (
        normalized in {"[unknown]", "unknown", "<unknown>", "??"}
        or normalized.startswith("[unknown")
    )


def weighted_metric(
    periods: int,
    blocks: int,
    subset_denominator: int,
    whole_denominator: int,
) -> dict[str, Any]:
    return {
        "weighted_event_period": periods,
        "raw_stack_blocks": blocks,
        "share_of_subset_percent": (
            periods / subset_denominator * 100.0 if subset_denominator else None
        ),
        "share_of_whole_process_percent": (
            periods / whole_denominator * 100.0 if whole_denominator else None
        ),
    }


def sample_metric(
    samples: Iterable[Sample],
    predicate: Callable[[Sample], bool],
    subset_denominator: int,
    whole_denominator: int,
) -> dict[str, Any]:
    periods = 0
    blocks = 0
    for sample in samples:
        if predicate(sample):
            periods += sample.period
            blocks += 1
    return weighted_metric(periods, blocks, subset_denominator, whole_denominator)


def rank_symbols(
    samples: Iterable[Sample],
    predicate: Callable[[Sample], bool],
    subset_denominator: int,
    whole_denominator: int,
    limit: int,
    leaf: bool,
) -> list[dict[str, Any]]:
    periods: Counter[str] = Counter()
    blocks: Counter[str] = Counter()
    for sample in samples:
        if not predicate(sample):
            continue
        if leaf:
            symbol = sample.symbols[0] if sample.symbols else "<no-frame>"
            periods[symbol] += sample.period
            blocks[symbol] += 1
        else:
            for symbol in set(sample.symbols):
                periods[symbol] += sample.period
                blocks[symbol] += 1
    rows: list[dict[str, Any]] = []
    for symbol, period in periods.most_common(limit):
        rows.append(
            {
                "symbol": symbol,
                **weighted_metric(period, blocks[symbol], subset_denominator, whole_denominator),
            }
        )
    return rows


def is_function(symbol: str, owner: str, method: str) -> bool:
    """Match a complete demangled Rust owner/method terminal."""

    return symbol.startswith(owner) and (
        symbol.endswith(f">>::{method}") or symbol.endswith(f"::{method}")
    )


SOURCE_OWNER = "<litchi_xls::workbook::source::SourceBackedWorkbook>"
EAGER_OWNER = "<litchi_xls::workbook::model::Workbook<"
SOURCE_OPEN_LITERAL = f"{SOURCE_OWNER}::from_read_at"
SOURCE_CELL_LITERAL = f"{SOURCE_OWNER}::cell_value_by_index"


def source_open(symbol: str) -> bool:
    return symbol == SOURCE_OPEN_LITERAL or symbol.endswith(
        "<litchi_xls::workbook::source::SourceBackedWorkbook>::from_read_at"
    )


def source_cell(symbol: str) -> bool:
    return symbol == SOURCE_CELL_LITERAL or symbol.endswith(
        "<litchi_xls::workbook::source::SourceBackedWorkbook>::cell_value_by_index"
    )


def eager_new(symbol: str) -> bool:
    return is_function(symbol, EAGER_OWNER, "new") and "litchi_xls::workbook::model" in symbol


def eager_from_ole_file(symbol: str) -> bool:
    return is_function(symbol, EAGER_OWNER, "from_ole_file") and "litchi_xls::workbook::model" in symbol


def harness_parent(symbol: str) -> bool:
    """Match the harness function only for a separate parent observation."""

    return re.search(
        rf"{re.escape(HARNESS_PARENT_LITERAL)}(?:$|::|\{{)", symbol
    ) is not None


def symbol_has_method(symbol: str, owner: str, method: str) -> bool:
    """Allow compiler-generated descendants of a named helper method."""

    if owner not in symbol:
        return False
    # The boundary keeps ``open`` separate from ``open_with_limits`` while
    # still accepting a compiler-generated ``open::{closure#0}`` descendant.
    return re.search(rf"::{re.escape(method)}(?:$|::|\{{)", symbol) is not None


def diagnostic_helpers() -> dict[str, Callable[[str], bool]]:
    return {
        "instrumented_source_read_at": lambda s: (
            "InstrumentedSource" in s and "read_at" in s and "ReadAt" in s
        ),
        "instrumented_source_read_exact_at": lambda s: (
            "InstrumentedSource" in s and "read_exact_at" in s
        ),
        "read_at_trait_helper": lambda s: (
            "litchi_core::source::ReadAt" in s and "read_at" in s
        ),
        "cfb_ole_open": lambda s: symbol_has_method(
            s, "litchi_cfb::file::OleFile", "open"
        ),
        "shared_ole_open": lambda s: symbol_has_method(
            s, "litchi_cfb::shared::SharedOleFile", "open"
        ),
        "shared_ole_open_with_limits": lambda s: symbol_has_method(
            s, "litchi_cfb::shared::SharedOleFile", "open_with_limits"
        ),
        "shared_ole_read_stream_range": lambda s: symbol_has_method(
            s, "litchi_cfb::shared::SharedOleFile", "read_stream_range"
        ),
        "shared_ole_open_stream": lambda s: symbol_has_method(
            s, "litchi_cfb::shared::SharedOleFile", "open_stream"
        ),
        "shared_ole_stream_cursor_at": lambda s: symbol_has_method(
            s, "litchi_cfb::shared::SharedOleFile", "stream_cursor_at"
        ),
        "shared_ole_stream_read_exact": lambda s: symbol_has_method(
            s, "litchi_cfb::shared::SharedOleStreamCursor", "read_exact"
        ),
        "shared_ole_stream_move_forward": lambda s: symbol_has_method(
            s, "litchi_cfb::shared::SharedOleStreamCursor", "move_forward"
        ),
        "cfb_ole_open_with_limits": lambda s: symbol_has_method(
            s, "litchi_cfb::file::OleFile", "open_with_limits"
        ),
        "cfb_read_sector_into": lambda s: symbol_has_method(
            s, "litchi_cfb::file::OleFile", "read_sector_into"
        ),
        "source_query_cell": lambda s: "litchi_xls::workbook::source::query_cell" in s,
        "source_worksheet_next_frame": lambda s: symbol_has_method(
            s, "litchi_xls::workbook::source::WorksheetScan", "next_frame"
        ),
        "source_worksheet_read_payload": lambda s: symbol_has_method(
            s, "litchi_xls::workbook::source::WorksheetScan", "read_payload"
        ),
    }


def first_index(symbols: Iterable[str], predicate: Callable[[str], bool]) -> int | None:
    for index, symbol in enumerate(symbols):
        if predicate(symbol):
            return index
    return None


def production_subsets() -> dict[str, dict[str, Any]]:
    return {
        "source_backed_from_read_at": {
            "label": "SourceBackedWorkbook::from_read_at production ancestors",
            "marker": SOURCE_OPEN_LITERAL,
            "predicate": source_open,
            "match_policy": "exact complete demangled owner/method terminal; generic compiler type arguments are accepted",
        },
        "source_backed_cell_value_by_index": {
            "label": "SourceBackedWorkbook::cell_value_by_index production ancestors",
            "marker": SOURCE_CELL_LITERAL,
            "predicate": source_cell,
            "match_policy": "exact complete demangled owner/method terminal; generic compiler type arguments are accepted",
        },
        "eager_workbook_new": {
            "label": "litchi_xls Workbook::new production ancestors",
            "marker": "<litchi_xls::workbook::model::Workbook<...>>::new",
            "predicate": eager_new,
            "match_policy": "litchi_xls workbook owner plus complete ::new terminal; this excludes unrelated Rust constructors",
        },
        "eager_workbook_from_ole_file": {
            "label": "litchi_xls Workbook::from_ole_file production ancestors",
            "marker": "<litchi_xls::workbook::model::Workbook<...>>::from_ole_file",
            "predicate": eager_from_ole_file,
            "match_policy": "litchi_xls workbook owner plus complete ::from_ole_file terminal",
        },
    }


def subset_summary(
    samples: list[Sample],
    name: str,
    spec: dict[str, Any],
    whole_period: int,
    limit: int,
) -> dict[str, Any]:
    predicate = spec["predicate"]
    matching = [sample for sample in samples if any(predicate(symbol) for symbol in sample.symbols)]
    period = sum(sample.period for sample in matching)
    symbols = sorted(
        {
            symbol
            for sample in matching
            for symbol in sample.symbols
            if predicate(symbol)
        }
    )
    return {
        "name": name,
        "label": spec["label"],
        "marker": spec["marker"],
        "match_policy": spec["match_policy"],
        "observed_marker_symbols": symbols,
        "scope": sample_metric(
            samples,
            lambda sample: any(predicate(symbol) for symbol in sample.symbols),
            period,
            whole_period,
        ),
        "leaf_period_weighted_ranking": rank_symbols(
            samples,
            lambda sample: any(predicate(symbol) for symbol in sample.symbols),
            period,
            whole_period,
            limit,
            leaf=True,
        ),
        "inclusive_period_weighted_ranking": rank_symbols(
            samples,
            lambda sample: any(predicate(symbol) for symbol in sample.symbols),
            period,
            whole_period,
            limit,
            leaf=False,
        ),
        "direct_callers_of_marker": direct_callers(
            matching, predicate, period, whole_period, limit
        ),
    }


def direct_callers(
    samples: Iterable[Sample],
    marker: Callable[[str], bool],
    subset_period: int,
    whole_period: int,
    limit: int,
) -> list[dict[str, Any]]:
    periods: Counter[str] = Counter()
    blocks: Counter[str] = Counter()
    for sample in samples:
        position = first_index(sample.symbols, marker)
        if position is None:
            continue
        caller = sample.symbols[position + 1] if position + 1 < len(sample.symbols) else "<none>"
        periods[caller] += sample.period
        blocks[caller] += 1
    rows: list[dict[str, Any]] = []
    total = sum(periods.values())
    for caller, period in periods.most_common(limit):
        row = weighted_metric(period, blocks[caller], subset_period, whole_period)
        row["caller"] = caller
        row["share_of_marker_period_percent"] = period / total * 100.0 if total else None
        rows.append(row)
    return rows


def helper_summary(
    samples: list[Sample],
    whole_period: int,
    subsets: dict[str, dict[str, Any]],
    limit: int,
) -> dict[str, Any]:
    result: dict[str, Any] = {}
    helpers = diagnostic_helpers()
    for name, predicate in helpers.items():
        matches = lambda sample, predicate=predicate: any(
            predicate(symbol) for symbol in sample.symbols
        )
        matching = [sample for sample in samples if matches(sample)]
        period = sum(sample.period for sample in matching)
        result[name] = {
            "description": helper_description(name),
            "whole_process_scope": sample_metric(
                samples, matches, whole_period, whole_period
            ),
            "whole_process_leaf_period_weighted_ranking": rank_symbols(
                samples, matches, whole_period, whole_period, limit, leaf=True
            ),
            "whole_process_inclusive_period_weighted_ranking": rank_symbols(
                samples, matches, whole_period, whole_period, limit, leaf=False
            ),
            "inside_production_subsets": {
                subset_name: sample_metric(
                    samples,
                    lambda sample, subset_pred=spec["predicate"]: matches(sample)
                    and any(subset_pred(symbol) for symbol in sample.symbols),
                    sum(
                        sample.period
                        for sample in samples
                        if any(spec["predicate"](symbol) for symbol in sample.symbols)
                    ),
                    whole_period,
                )
                for subset_name, spec in subsets.items()
            },
            "observed_helper_symbols": sorted(
                {
                    symbol
                    for sample in matching
                    for symbol in sample.symbols
                    if predicate(symbol)
                }
            ),
            "weighted_event_period": period,
        }
    return result


def helper_description(name: str) -> str:
    descriptions = {
        "instrumented_source_read_at": "perf-baseline InstrumentedSource ReadAt callback; a diagnostic source wrapper, not a physical I/O phase timer",
        "instrumented_source_read_exact_at": "InstrumentedSource exact-range adapter helper; a diagnostic source wrapper",
        "read_at_trait_helper": "ReadAt trait/source adapter frames; attribution may include CFB and source wrappers",
        "cfb_ole_open": "eager CFB OleFile::open path",
        "shared_ole_open": "source-backed CFB SharedOleFile::open path",
        "shared_ole_open_with_limits": "source-backed CFB SharedOleFile::open_with_limits path",
        "shared_ole_read_stream_range": "source-backed CFB stream-range read helper",
        "shared_ole_open_stream": "source-backed CFB stream-open helper",
        "shared_ole_stream_cursor_at": "source-backed CFB stream cursor helper",
        "shared_ole_stream_read_exact": "source-backed CFB stream cursor exact-read helper",
        "shared_ole_stream_move_forward": "source-backed CFB stream cursor movement helper",
        "cfb_ole_open_with_limits": "eager CFB OleFile::open_with_limits path",
        "cfb_read_sector_into": "eager CFB sector read helper",
        "source_query_cell": "source-backed XLS cell query helper",
        "source_worksheet_next_frame": "source-backed XLS worksheet frame scanner",
        "source_worksheet_read_payload": "source-backed XLS worksheet payload read helper",
    }
    return descriptions.get(name, "diagnostic helper symbol")


def recursive_dicts(value: Any) -> Iterable[dict[str, Any]]:
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from recursive_dicts(child)
    elif isinstance(value, list):
        for child in value:
            yield from recursive_dicts(child)


def collect_strings(value: Any) -> Iterable[str]:
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for key, child in value.items():
            if isinstance(key, str):
                yield key
            yield from collect_strings(child)
    elif isinstance(value, list):
        for child in value:
            yield from collect_strings(child)


def command_token_lists(value: Any) -> Iterable[list[str]]:
    """Yield argv-like lists from common capture metadata shapes."""

    if isinstance(value, dict):
        for key, child in value.items():
            key_lower = str(key).lower()
            if key_lower in {"argv", "args", "command", "cmd", "command_line", "commandline"}:
                if isinstance(child, list):
                    yield [str(token) for token in child]
                elif isinstance(child, str):
                    try:
                        yield shlex.split(child)
                    except ValueError:
                        yield child.split()
            yield from command_token_lists(child)
    elif isinstance(value, list):
        if value and all(isinstance(item, (str, int, float)) for item in value):
            yield [str(item) for item in value]
        else:
            for child in value:
                yield from command_token_lists(child)


def append_unique(result: dict[str, Any], key: str, value: Any) -> None:
    if value not in result[key]:
        result[key].append(value)


def observe_toolchain(result: dict[str, Any], value: str) -> None:
    # Environment captures use either RUSTUP_TOOLCHAIN or a full rustc
    # version line.  Store the stable release token so both forms verify the
    # same requested 1.98.1 identity.
    versions = re.findall(r"\b\d+\.\d+\.\d+\b", value)
    if versions:
        for version in versions:
            append_unique(result, "observed_toolchains", version)
    else:
        append_unique(result, "observed_toolchains", value)


def observe_command_tokens(tokens: list[str], result: dict[str, Any]) -> None:
    """Extract event/frequency/sample/warmup facts from argv metadata."""

    for index, token in enumerate(tokens):
        for case in CASE_RE.findall(token):
            append_unique(result, "observed_cases", case)
        if "cycles" in token:
            event = next((part for part in token.split("=") if "cycles" in part), "cycles")
            append_unique(result, "observed_events", event.rstrip(":"))

        def following_number() -> int | None:
            if index + 1 >= len(tokens):
                return None
            try:
                return int(tokens[index + 1].strip(","))
            except ValueError:
                return None

        if token in {"--samples", "--sample-count"}:
            value = following_number()
            if value is not None:
                append_unique(result, "observed_sample_counts", value)
        elif token.startswith("--samples=") or token.startswith("--sample-count="):
            try:
                append_unique(result, "observed_sample_counts", int(token.split("=", 1)[1]))
            except ValueError:
                pass

        if token in {"--warmup", "--warmup-iterations", "--warmup_iterations", "--warmup-iterations-per-case"}:
            value = following_number()
            if value is not None:
                append_unique(result, "observed_warmup_counts", value)
        elif any(token.startswith(prefix) for prefix in ("--warmup=", "--warmup-iterations=", "--warmup_iterations=")):
            try:
                append_unique(result, "observed_warmup_counts", int(token.split("=", 1)[1]))
            except ValueError:
                pass

        if token in {"-F", "--freq", "--frequency"}:
            value = following_number()
            if value is not None:
                append_unique(result, "observed_frequencies", value)
        elif re.fullmatch(r"-F\d+", token):
            append_unique(result, "observed_frequencies", int(token[2:]))
        elif token.startswith("--freq=") or token.startswith("--frequency="):
            try:
                append_unique(result, "observed_frequencies", int(token.split("=", 1)[1]))
            except ValueError:
                pass

        if token == "-e" and index + 1 < len(tokens):
            append_unique(result, "observed_events", tokens[index + 1].rstrip(":"))
        elif token.startswith("-e") and token != "-e" and "cycles" in token:
            append_unique(result, "observed_events", token[2:].rstrip(":"))

        if "--call-graph=fp,127" in token or (
            token == "--call-graph" and index + 1 < len(tokens) and tokens[index + 1] == "fp,127"
        ):
            append_unique(result, "observed_call_graphs", "fp,127")
        if token == "--no-inline":
            result["no_inline_observed"] = True


def load_json(path: Path) -> Any | None:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, UnicodeDecodeError):
        return None


def profile_path_for_script(capture: Path | None, script: Path) -> Path | None:
    """Map ``eager-profile-script.stdout`` to its profile JSON when present."""

    if capture is None or not capture.is_dir():
        return None
    name = script.name
    suffix = "-script.stdout"
    if name.endswith(suffix):
        candidate = capture / f"{name[:-len(suffix)]}.json"
        if candidate.is_file():
            return candidate
    return None


def capture_identity(
    capture: Path | None,
    expected_case: str | None,
    selected_profile: Path | None = None,
) -> dict[str, Any]:
    """Collect identity observations without requiring one capture schema."""

    result: dict[str, Any] = {
        "capture_directory": str(capture) if capture else None,
        "expected_case": expected_case,
        "observed_cases": [],
        "observed_sample_counts": [],
        "observed_warmup_counts": [],
        "observed_events": [],
        "observed_frequencies": [],
        "observed_toolchains": [],
        "observed_call_graphs": [],
        "no_inline_observed": False,
        "selected_profile_file": None,
        "selected_profile_cases": [],
        "selected_profile_sample_counts": [],
        "selected_profile_warmup_counts": [],
        "metadata_files": [],
        "warnings": [],
    }
    if capture is None or not capture.is_dir():
        result["warnings"].append("capture metadata directory is absent; identity is unverified")
        return result

    if selected_profile is not None and selected_profile.is_file():
        result["selected_profile_file"] = file_binding(selected_profile)
        profile_data = load_json(selected_profile)
        if isinstance(profile_data, dict):
            configuration = profile_data.get("configuration")
            if isinstance(configuration, dict):
                for value in configuration.get("cases", []):
                    if isinstance(value, str):
                        result["selected_profile_cases"].extend(CASE_RE.findall(value))
                for key in ("samples", "samples_per_case", "sample_count"):
                    value = configuration.get(key)
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        append_unique(result, "selected_profile_sample_counts", value)
                for key in ("warmup", "warmups", "warmup_iterations_per_case"):
                    value = configuration.get(key)
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        append_unique(result, "selected_profile_warmup_counts", value)
            if isinstance(configuration, dict):
                # Some capture writers put event/frequency details next to
                # the selector.  They are still useful when command logs are
                # incomplete, so retain them in the global observations.
                event = configuration.get("event")
                if isinstance(event, str):
                    append_unique(result, "observed_events", event.rstrip(":"))
                frequency = configuration.get("frequency_hz")
                if isinstance(frequency, (int, float)) and not isinstance(frequency, bool):
                    append_unique(result, "observed_frequencies", frequency)
        result["selected_profile_cases"] = sorted(
            set(result["selected_profile_cases"]), key=str
        )

    paths = sorted(path for path in capture.glob("*.json") if path.is_file())
    for path in paths:
        data = load_json(path)
        if data is None:
            result["warnings"].append(f"could not parse metadata JSON: {path.name}")
            continue
        result["metadata_files"].append(file_binding(path))
        for item in recursive_dicts(data):
            for key, value in item.items():
                key_lower = str(key).lower()
                if isinstance(value, str):
                    for case in CASE_RE.findall(value):
                        if case not in result["observed_cases"]:
                            result["observed_cases"].append(case)
                    if "event" in key_lower and value not in result["observed_events"]:
                        result["observed_events"].append(value.rstrip(":"))
                    if (
                        "toolchain" in key_lower
                        or key_lower in {"rustc", "rustc_version"}
                    ):
                        observe_toolchain(result, value)
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    if (
                        key_lower in {"samples", "sample_count", "record_samples"}
                        or key_lower.endswith("_samples")
                        or key_lower.startswith("samples_")
                    ):
                        if value not in result["observed_sample_counts"]:
                            result["observed_sample_counts"].append(value)
                    if "warmup" in key_lower:
                        if value not in result["observed_warmup_counts"]:
                            result["observed_warmup_counts"].append(value)
                    if key_lower in {"frequency", "freq", "hz"} or key_lower.endswith("_hz"):
                        if value not in result["observed_frequencies"]:
                            result["observed_frequencies"].append(value)
        for text in collect_strings(data):
            for case in CASE_RE.findall(text):
                if case not in result["observed_cases"]:
                    result["observed_cases"].append(case)
            if "cycles" in text:
                event = "cycles:u" if "cycles:u" in text else "cycles"
                if event not in result["observed_events"]:
                    result["observed_events"].append(event)
            if "--call-graph=fp,127" in text or "--call-graph" in text and "fp,127" in text:
                result.setdefault("observed_call_graphs", []).append("fp,127")
            if "--no-inline" in text:
                result["no_inline_observed"] = True
        for tokens in command_token_lists(data):
            observe_command_tokens(tokens, result)

    # Keep lists deterministic while preserving their useful scalar values.
    for key in (
        "observed_cases",
        "observed_events",
        "observed_sample_counts",
        "observed_warmup_counts",
        "observed_frequencies",
        "observed_toolchains",
        "observed_call_graphs",
    ):
        result[key] = sorted(set(result[key]), key=str)
    checks: dict[str, Any] = {}
    checks["case"] = identity_check(
        expected_case,
        result["selected_profile_cases"] or result["observed_cases"],
        "selected profile case",
    )
    checks["event"] = identity_check(EXPECTED_EVENT, result["observed_events"], "event")
    checks["call_graph"] = identity_check(
        "fp,127", result["observed_call_graphs"], "call graph"
    )
    checks["no_inline"] = {
        "status": "verified" if result["no_inline_observed"] else "unverified",
        "expected": True,
        "observed": result["no_inline_observed"],
    }
    checks["toolchain"] = identity_check(
        EXPECTED_TOOLCHAIN, result["observed_toolchains"], "toolchain"
    )
    checks["samples"] = numeric_identity_check(
        EXPECTED_SAMPLES,
        result["selected_profile_sample_counts"] or result["observed_sample_counts"],
        "selected profile sample count",
    )
    checks["warmup"] = numeric_identity_check(
        EXPECTED_WARMUP,
        result["selected_profile_warmup_counts"] or result["observed_warmup_counts"],
        "selected profile warmup count",
    )
    checks["frequency"] = numeric_identity_check(
        EXPECTED_FREQUENCY, result["observed_frequencies"], "frequency"
    )
    result["checks"] = checks
    return result


def identity_check(expected: Any, observed: list[Any], label: str) -> dict[str, Any]:
    if expected is None:
        return {"status": "not_requested", "expected": None, "observed": observed}
    if not observed:
        return {
            "status": "unverified",
            "expected": expected,
            "observed": observed,
            "message": f"{label} was not found in metadata",
        }
    return {
        "status": "verified" if expected in observed else "mismatch",
        "expected": expected,
        "observed": observed,
    }


def numeric_identity_check(expected: int, observed: list[Any], label: str) -> dict[str, Any]:
    normalized = {int(value) for value in observed if isinstance(value, (int, float))}
    return identity_check(expected, sorted(normalized), label)


def source_binding(root: Path, revision: str | None) -> dict[str, Any]:
    specs = [
        (
            "tools/perf-baseline/src/lib.rs",
            {
                "timed_xls_harness": re.compile(r"fn run_xls_source_backed_case"),
                "instrumented_source": re.compile(r"struct InstrumentedSource"),
                "read_at_impl": re.compile(r"impl ReadAt for InstrumentedSource"),
            },
        ),
        (
            "crates/litchi-xls/src/workbook/source.rs",
            {
                "source_backed_open": re.compile(r"pub fn from_read_at\("),
                "source_backed_cell": re.compile(r"pub fn cell_value_by_index\("),
            },
        ),
        (
            "crates/litchi-xls/src/workbook/package.rs",
            {
                "eager_new": re.compile(r"pub fn new\("),
                "eager_from_ole_file": re.compile(r"pub fn from_ole_file\("),
            },
        ),
        (
            "crates/litchi-cfb/src/shared.rs",
            {
                "shared_open": re.compile(r"pub fn open\("),
                "shared_stream_range": re.compile(r"pub fn read_stream_range\("),
                "shared_open_stream": re.compile(r"pub fn open_stream\("),
            },
        ),
        (
            "crates/litchi-cfb/src/file.rs",
            {
                "eager_cfb_open": re.compile(r"pub fn open_with_limits\("),
                "eager_sector_read": re.compile(r"fn read_sector_into\("),
            },
        ),
    ]
    if not revision:
        return {
            "status": "revision_missing",
            "interpretation": "No capture revision was available; current source was not substituted silently.",
            "files": [],
        }

    files: list[dict[str, Any]] = []
    errors: list[str] = []
    for relative, checks in specs:
        try:
            blob = subprocess.check_output(
                ["git", "-C", str(root), "show", f"{revision}:{relative}"],
                stderr=subprocess.STDOUT,
            )
        except (OSError, subprocess.CalledProcessError) as error:
            errors.append(f"{relative}: {error}")
            continue
        lines = blob.decode("utf-8", errors="replace").splitlines()
        found = {
            name: [line_no for line_no, line in enumerate(lines, 1) if pattern.search(line)]
            for name, pattern in checks.items()
        }
        files.append(
            {
                "path": relative,
                "bytes": len(blob),
                "sha256": sha256_bytes(blob),
                "checks": found,
            }
        )
    return {
        "status": "bound" if not errors else "partial",
        "revision": revision,
        "files": files,
        "errors": errors,
        "interpretation": "Git blobs are bound to the capture revision; symbol matching still depends on the captured demangling format.",
    }


def discover_revision(capture: Path | None) -> tuple[str | None, dict[str, Any]]:
    if capture is None or not capture.is_dir():
        return None, {}
    names = ["environment.json", "build-identity.json", "capture-commands.json", "commands.json"]
    names.extend(path.name for path in sorted(capture.glob("*-profile.json")))
    for name in names:
        path = capture / name
        data = load_json(path) if path.is_file() else None
        if not isinstance(data, dict):
            continue
        source_before = data.get("source_before")
        if isinstance(source_before, dict) and isinstance(source_before.get("revision"), str):
            return source_before["revision"], {
                "metadata_file": name,
                "revision": source_before["revision"],
                "source_status": source_before.get("status"),
            }
        for item in recursive_dicts(data):
            for key in ("revision", "commit", "git_revision"):
                value = item.get(key)
                if isinstance(value, str) and re.fullmatch(r"[0-9a-fA-F]{7,64}", value):
                    status = item.get("source_status")
                    if status is None and isinstance(item.get("git_worktree_dirty"), bool):
                        status = "dirty" if item["git_worktree_dirty"] else ""
                    return value, {
                        "metadata_file": name,
                        "revision": value,
                        "source_status": status,
                    }
    return None, {}


def metadata_bindings(capture: Path | None, report: Path | None) -> list[dict[str, Any]]:
    paths: list[Path] = []
    if capture and capture.is_dir():
        paths.extend(sorted(path for path in capture.glob("*.json") if path.is_file()))
    if report and report.is_file():
        paths.append(report)
    return [file_binding(path) for path in paths]


def infer_case(identity: dict[str, Any], requested: str | None) -> str | None:
    if requested:
        return requested
    cases = identity.get("observed_cases", [])
    return cases[0] if len(cases) == 1 else None


def marker_observations(samples: list[Sample], marker: Callable[[str], bool]) -> dict[str, Any]:
    matching_symbols = Counter(
        symbol
        for sample in samples
        for symbol in sample.symbols
        if marker(symbol)
    )
    return {
        "sample_blocks": sum(
            1 for sample in samples if any(marker(symbol) for symbol in sample.symbols)
        ),
        "symbols": [symbol for symbol, _ in matching_symbols.most_common(20)],
    }


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--script", type=Path, default=DEFAULT_SCRIPT)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--capture", type=Path, default=DEFAULT_CAPTURE)
    parser.add_argument("--repo", type=Path, default=DEFAULT_ROOT)
    parser.add_argument("--case", type=str, default=None)
    parser.add_argument("--top", type=int, default=40)
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_ATTRIBUTION / "attribution-summary.json",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    if args.top <= 0:
        raise SystemExit("--top must be positive")
    script_path = args.script.resolve()
    report_path = args.report.resolve()
    capture_path = args.capture.resolve() if args.capture else None
    repo_path = args.repo.resolve()
    output_path = args.output.resolve()
    if not script_path.is_file():
        raise SystemExit(f"missing required perf script input: {script_path}")

    samples, parse_stats = parse_samples(script_path)
    if not samples:
        raise SystemExit("no cycle sample blocks parsed")
    whole_period = sum(sample.period for sample in samples)

    revision, revision_observation = discover_revision(capture_path)
    selected_profile = profile_path_for_script(capture_path, script_path)
    identity = capture_identity(capture_path, args.case, selected_profile)
    selected_case = infer_case(identity, args.case)
    subsets = production_subsets()
    subset_results = {
        name: subset_summary(samples, name, spec, whole_period, args.top)
        for name, spec in subsets.items()
    }
    parent_samples = [
        sample for sample in samples if any(harness_parent(symbol) for symbol in sample.symbols)
    ]
    parent_period = sum(sample.period for sample in parent_samples)
    parent_observation = {
        "marker": HARNESS_PARENT_LITERAL,
        "match_policy": "observed litchi_perf_baseline harness parent symbol; this is a stack observation, not the Instant::now phase",
        "scope": sample_metric(
            samples,
            lambda sample: any(harness_parent(symbol) for symbol in sample.symbols),
            parent_period,
            whole_period,
        ),
        "leaf_period_weighted_ranking": rank_symbols(
            samples,
            lambda sample: any(harness_parent(symbol) for symbol in sample.symbols),
            parent_period,
            whole_period,
            args.top,
            leaf=True,
        ),
        "inclusive_period_weighted_ranking": rank_symbols(
            samples,
            lambda sample: any(harness_parent(symbol) for symbol in sample.symbols),
            parent_period,
            whole_period,
            args.top,
            leaf=False,
        ),
        "interpretation": "Samples under this parent can include build_xls_source_layout setup/oracle work and post-operation checks because those calls execute within the same harness function. Do not read this as an operation-only or timed-phase total.",
    }
    helper_results = helper_summary(samples, whole_period, subsets, args.top)
    unknown_period = sum(
        sample.period
        for sample in samples
        for symbol in {symbol for symbol in sample.symbols if is_unknown_symbol(symbol)}
    )
    unknown_blocks = sum(
        1 for sample in samples if any(is_unknown_symbol(symbol) for symbol in sample.symbols)
    )

    # This is a deliberately explicit warning surface.  It prevents a reader
    # from interpreting a sampled callchain subset as the Instant::now phase.
    limitations = [
        "All weighted values are perf cycle sample periods. They are period-weighted CPU attribution, not wall-clock latency, operation duration, or a phase timer.",
        "Whole-process samples may include process startup, corpus/setup work, source construction, child work, and post-timer oracle/verification work.",
        "A production subset is selected only when its exact observed litchi_xls ancestor appears in the stack. It does not classify every run_xls_source_backed_case descendant as timed.",
        "Setup or oracle calls to the same constructors can appear in the same process and may share indistinguishable callchains; the ancestor subsets therefore cannot fully separate those calls from the timed invocation.",
        "A missing marker means no matching frame was observed or retained by perf; it does not prove that the operation or helper consumed no CPU.",
        "Diagnostic source-helper rows describe stack attribution around InstrumentedSource/CFB adapters. They do not measure physical I/O latency or bytes read by themselves.",
        "This report is descriptive evidence for one capture. It does not establish a before/after speedup, regression, or phase-level cost.",
        "The parsed perf sample-block count is an event-record count and need not equal the harness samples_per_case configuration; that configuration is checked separately from metadata.",
        "No-inline preserves call frames where available but does not guarantee that the compiler kept every source method out of line; absent key-function markers remain unresolved rather than zero-cost proof.",
    ]
    summary: dict[str, Any] = {
        "schema": SCHEMA,
        "purpose": "Bounded XLS CPU attribution for source-backed and eager one-cell captures",
        "timing_semantics": {
            "phase_latency": False,
            "weighted_event": EXPECTED_EVENT,
            "weight": "perf sample period",
            "whole_process_scope": "all parsed process/child samples in the supplied perf script export",
            "production_subset_scope": "inclusive callchain presence of an exact production ancestor marker",
        },
        "requested_capture_contract": {
            "event": EXPECTED_EVENT,
            "frequency_hz": EXPECTED_FREQUENCY,
            "samples": EXPECTED_SAMPLES,
            "warmup_iterations_per_case": EXPECTED_WARMUP,
            "call_graph": "fp,127",
            "perf_script": "--no-inline",
            "toolchain": EXPECTED_TOOLCHAIN,
        },
        "capture_identity": {
            **identity,
            "selected_case": selected_case,
            "revision_observation": revision_observation,
        },
        "inputs": {
            "perf_script": file_binding(script_path, repo_path),
            "perf_report": file_binding(report_path, repo_path) if report_path.is_file() else None,
            "capture_metadata": metadata_bindings(capture_path, report_path),
            "parser": file_binding(Path(__file__).resolve(), repo_path),
        },
        "source_binding": source_binding(repo_path, revision),
        "sample_parser": {
            **parse_stats,
            "whole_process_weighted_event_period": whole_period,
            "whole_process_raw_stack_blocks": len(samples),
            "event": EXPECTED_EVENT,
            "weight_semantics": "perf sample period; weighted event period, not wall time",
            "unknown_frame_period": unknown_period,
            "unknown_frame_blocks": unknown_blocks,
            "unknown_stack_period": unknown_period,
            "unknown_stack_blocks": unknown_blocks,
            "unknown_frame_share_of_whole_process_percent": (
                unknown_period / whole_period * 100.0 if whole_period else None
            ),
        },
        "whole_process": {
            "scope": "all parsed cycles samples in the supplied perf script export",
            "weighted_event_period": whole_period,
            "raw_stack_blocks": len(samples),
            "leaf_period_weighted_ranking": rank_symbols(
                samples, lambda _sample: True, whole_period, whole_period, args.top, leaf=True
            ),
            "inclusive_period_weighted_ranking": rank_symbols(
                samples, lambda _sample: True, whole_period, whole_period, args.top, leaf=False
            ),
        },
        "production_ancestor_subsets": subset_results,
        "harness_parent_observation": parent_observation,
        "diagnostic_source_helpers": helper_results,
        "marker_observations": {
            "harness_parent_run_xls_source_backed_case": marker_observations(samples, harness_parent),
            "source_backed_from_read_at": marker_observations(samples, source_open),
            "source_backed_cell_value_by_index": marker_observations(samples, source_cell),
            "eager_workbook_new": marker_observations(samples, eager_new),
            "eager_workbook_from_ole_file": marker_observations(samples, eager_from_ole_file),
        },
        "limitations": limitations,
        "interpretation": {
            "callchain_order": "perf script frames are interpreted leaf to caller/root",
            "subset_overlap": "Subsets are independent inclusive predicates; one sample can contribute to more than one subset or helper row.",
            "eager_route": "The current one-cell harness calls Workbook::new(Cursor<Vec<u8>>); from_ole_file is retained as a production marker for captures or routes that expose it.",
            "source_route": "SourceBackedWorkbook::from_read_at and cell_value_by_index are exact production ancestor markers, not elapsed phase boundaries.",
        },
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {
                "output": str(output_path),
                "case": selected_case,
                "whole_process_period": whole_period,
                "whole_process_blocks": len(samples),
                "source_open_period": subset_results["source_backed_from_read_at"]["scope"]["weighted_event_period"],
                "source_cell_period": subset_results["source_backed_cell_value_by_index"]["scope"]["weighted_event_period"],
                "eager_new_period": subset_results["eager_workbook_new"]["scope"]["weighted_event_period"],
                "unknown_frame_lines": parse_stats.get("unknown_frame_lines", 0),
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
