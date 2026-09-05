#!/usr/bin/env python3
"""Attribute the 0412 plain OwnedSource XLS CPU capture.

The input is a perf script --no-inline export for the single
xls_owned_source_open_one_cell profile.  Sample periods are retained as
weighted cycles:u event periods.  They are CPU attribution evidence, not
elapsed time or an operation-only latency measurement.

This postprocessor deliberately treats the profile, catalog, command journal,
and candidate build identity as evidence surfaces.  It reports missing,
malformed, unknown, empty, and lost records without turning a valid nullable
source field into a failed identity check.  The parser has no fixed list of
the harness's other selectors; it observes case names from the supplied
metadata.
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

SCHEMA = "0412-xls-owned-source-cpu-attribution-v1"
DEFAULT_ROOT = Path("/home/zhuhe/code/litchi")
DEFAULT_CAPTURE = Path("/tmp/litchi-goal-0412-capture")
DEFAULT_SCRIPT = DEFAULT_CAPTURE / "owned-profile-script.stdout"
DEFAULT_REPORT = DEFAULT_CAPTURE / "owned-profile-self.stdout"
DEFAULT_PROFILE = DEFAULT_CAPTURE / "owned-profile.json"
DEFAULT_CATALOG = DEFAULT_CAPTURE / "owned-profile.catalog.json"
DEFAULT_COMMANDS = DEFAULT_CAPTURE / "commands.json"
DEFAULT_PROTOCOL = DEFAULT_CAPTURE / "protocol.json"
DEFAULT_BUILD_IDENTITY = DEFAULT_CAPTURE / "candidate-build-identity.json"
DEFAULT_OUTPUT = Path("/tmp/litchi-goal-0412-attribution/owned-profile-attribution.json")

SOURCE_OWNER = "litchi_xls::workbook::source::SourceBackedWorkbook"
CASE_RE = re.compile(r"\bxls_[A-Za-z0-9_]+\b")
HASH_RE = re.compile(r"^[0-9a-fA-F]{7,64}$")
HEADER_RE = re.compile(
    r"^\s*(?P<command>.+?)\s+(?P<pid>\d+)\s+(?P<timestamp>[^:]+):\s+"
    r"(?P<period>-?[\d,]+)\s+(?P<event>\S+)\s*$"
)
FRAME_RE = re.compile(
    r"^\s*(?:0x)?[0-9a-fA-F]+\s+(?P<symbol>.+?)"
    r"\+0x[0-9a-fA-F]+\s+\([^)]*\)\s*$"
)
FRAME_RE_NO_OFFSET = re.compile(
    r"^\s*(?:0x)?[0-9a-fA-F]+\s+(?P<symbol>.+?)\s+\([^)]*\)\s*$"
)
UNKNOWN_ONLY_RE = re.compile(
    r"^\s*(?:\?\?|unknown|<unknown>|\[unknown\])(?:\s+\([^)]*\))?\s*$",
    re.IGNORECASE,
)
UNKNOWN_ADDRESS_RE = re.compile(
    r"^\s*(?:0x)?[0-9a-fA-F]+\s+(?:\?\?|unknown|<unknown>|\[unknown\])"
    r"(?:\s+\([^)]*\))?\s*$",
    re.IGNORECASE,
)
LOST_RE = re.compile(
    r"lost\s+(?P<first>[\d,]+)\s+samples?"
    r"|(?P<second>[\d,]+)\s+samples?\s+lost"
    r"|samples?\s+lost\s*[:=]\s*(?P<third>[\d,]+)",
    re.IGNORECASE,
)
REPORT_EVENT_RE = re.compile(r"\bevent\s+['\"](?P<event>[^'\"]+)['\"]", re.IGNORECASE)
REPORT_SAMPLES_RE = re.compile(r"#\s*Samples:\s*(?P<samples>[0-9.,]+[KMG]?)", re.IGNORECASE)
REPORT_ROW_RE = re.compile(
    r"^\s*(?P<percent>\d+(?:\.\d+)?)%\s+"
    r"(?:(?P<self>\d+(?:\.\d+)?)%\s+)?(?P<rest>.+?)\s*$"
)


@dataclass(frozen=True)
class Sample:
    """A cycle sample; symbols retain perf's leaf-to-root order."""

    period: int
    event: str
    symbols: tuple[str, ...]
    unknown_frame_count: int
    unparsed_frame_count: int


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
    result: dict[str, Any] = {
        "path": str(resolved),
        "bytes": resolved.stat().st_size,
        "sha256": sha256_file(resolved),
    }
    if root is not None:
        try:
            result["relative_to_repo"] = str(resolved.relative_to(root.resolve()))
        except ValueError:
            pass
    return result


def normalize_event(event: str) -> str:
    return event.strip().rstrip(":")


def is_unknown_symbol(symbol: str) -> bool:
    normalized = symbol.strip().lower()
    return (
        normalized in {"??", "unknown", "<unknown>", "[unknown]"}
        or normalized.startswith("[unknown")
    )


def symbol_from_frame(line: str) -> str | None:
    """Parse a perf frame while retaining common unknown-frame spellings."""

    if UNKNOWN_ONLY_RE.match(line) or UNKNOWN_ADDRESS_RE.match(line):
        return "[unknown]"
    match = FRAME_RE.match(line)
    if match:
        return match.group("symbol").strip()
    match = FRAME_RE_NO_OFFSET.match(line)
    if match:
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
    if "lost" in line.lower():
        return None
    return None


def looks_like_sample_header(line: str) -> bool:
    stripped = line.strip()
    if not stripped or stripped.startswith("#"):
        return False
    return "cycles" in stripped.lower() or bool(
        re.search(r"\s+\d+\s+[^:]+:\s+-?[\d,]+\s+\S+", stripped)
    )


def parse_samples(path: Path) -> tuple[list[Sample], dict[str, Any]]:
    """Parse cycle blocks and retain diagnostics for imperfect perf output."""

    samples: list[Sample] = []
    stats: Counter[str] = Counter()
    current_event: str | None = None
    current_period: int | None = None
    current_symbols: list[str] = []
    current_unknown = 0
    current_unparsed = 0
    current_invalid = False

    def reset() -> None:
        nonlocal current_event, current_period, current_symbols
        nonlocal current_unknown, current_unparsed, current_invalid
        current_event = None
        current_period = None
        current_symbols = []
        current_unknown = 0
        current_unparsed = 0
        current_invalid = False

    def flush() -> None:
        nonlocal current_unknown, current_unparsed
        if current_event is None:
            return
        if current_invalid:
            stats["invalid_sample_blocks"] += 1
            stats["invalid_sample_period"] += current_period or 0
            reset()
            return
        if current_period is None:
            stats["non_cycle_blocks"] += 1
            reset()
            return
        sample = Sample(
            period=current_period,
            event=current_event,
            symbols=tuple(current_symbols),
            unknown_frame_count=current_unknown,
            unparsed_frame_count=current_unparsed,
        )
        samples.append(sample)
        stats["sample_blocks_seen"] += 1
        if current_period <= 0:
            stats["zero_or_negative_period_samples"] += 1
        if not current_symbols:
            stats["empty_stack_blocks"] += 1
            stats["empty_stack_period"] += current_period
            if current_unparsed:
                stats["unparsed_only_blocks"] += 1
                stats["unparsed_only_period"] += current_period
        if current_unknown:
            stats["unknown_frame_blocks"] += 1
            stats["unknown_frame_period"] += current_period
            stats["unknown_frame_occurrences"] += current_unknown
            stats["unknown_frame_occurrence_period"] += current_period * current_unknown
        stats["unparsed_frame_lines"] += current_unparsed
        reset()

    with path.open("r", encoding="utf-8", errors="replace") as source:
        for raw in source:
            stats["input_lines"] += 1
            line = raw.rstrip("\r\n")
            header = HEADER_RE.match(line)
            if header:
                if current_event is not None:
                    stats["implicit_block_boundaries"] += 1
                flush()
                event = normalize_event(header.group("event"))
                current_event = event
                current_period = int(header.group("period").replace(",", ""))
                if event == "cycles:u":
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
                if current_event is None:
                    continue
                stats["lost_metadata_inside_block"] += 1
                continue

            if current_event is None:
                if line.lstrip().startswith("#"):
                    stats["metadata_lines"] += 1
                elif looks_like_sample_header(line):
                    stats["malformed_cycle_headers"] += 1
                else:
                    stats["nonempty_lines_outside_blocks"] += 1
                continue

            if line.lstrip().startswith("#"):
                stats["metadata_lines_inside_blocks"] += 1
                continue

            if current_period is None:
                stats["non_cycle_frame_lines"] += 1
                continue

            stats["nonempty_lines_inside_cycle_blocks"] += 1
            symbol = symbol_from_frame(line)
            if symbol is None:
                current_unparsed += 1
                continue
            current_symbols.append(symbol)
            stats["parsed_frame_lines"] += 1
            if is_unknown_symbol(symbol):
                current_unknown += 1
                stats["unknown_frame_lines"] += 1

    if current_event is not None:
        stats["unterminated_blocks"] += 1
    flush()
    stats["sample_blocks"] = len(samples)
    stats["total_weighted_event_period"] = sum(sample.period for sample in samples)
    for key in (
        "cycle_headers",
        "non_cycle_headers",
        "sample_blocks_seen",
        "empty_stack_blocks",
        "empty_stack_period",
        "unknown_frame_lines",
        "unknown_frame_blocks",
        "unknown_frame_period",
        "unknown_frame_occurrences",
        "unknown_frame_occurrence_period",
        "unparsed_frame_lines",
        "unparsed_only_blocks",
        "unparsed_only_period",
        "lost_sample_count",
        "lost_metadata_lines",
        "unquantified_lost_lines",
        "malformed_cycle_headers",
        "invalid_cycle_event_headers",
        "invalid_cycle_period_headers",
        "invalid_sample_blocks",
        "invalid_sample_period",
        "zero_or_negative_period_samples",
    ):
        stats.setdefault(key, 0)
    return samples, dict(stats)


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
            symbol = sample.symbols[0] if sample.symbols else "<empty-stack>"
            periods[symbol] += sample.period
            blocks[symbol] += 1
        else:
            for symbol in set(sample.symbols):
                periods[symbol] += sample.period
                blocks[symbol] += 1
    return [
        {
            "symbol": symbol,
            **weighted_metric(
                period,
                blocks[symbol],
                subset_denominator,
                whole_denominator,
            ),
        }
        for symbol, period in periods.most_common(limit)
    ]


def method_marker(symbol: str, method: str) -> bool:
    """Match this owner and exact method terminal, including closure frames."""

    pattern = (
        rf"(?:<)?{re.escape(SOURCE_OWNER)}(?:>|)::"
        rf"{re.escape(method)}(?:$|::|\{{)"
    )
    return re.search(pattern, symbol) is not None


def source_open(symbol: str) -> bool:
    return method_marker(symbol, "from_read_at")


def source_open_profile_method(symbol: str) -> bool:
    """Match the open entry and its generated ``from_read_at_with_limits`` wrapper."""

    pattern = (
        rf"(?:<)?{re.escape(SOURCE_OWNER)}(?:>|)::from_read_at"
        rf"(?:_with_limits)?(?:$|::|\{{)"
    )
    return re.search(pattern, symbol) is not None


def source_cell(symbol: str) -> bool:
    return method_marker(symbol, "cell_value_by_index")


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
        index = next(
            (index for index, symbol in enumerate(sample.symbols) if marker(symbol)),
            None,
        )
        if index is None:
            continue
        caller = sample.symbols[index + 1] if index + 1 < len(sample.symbols) else "<none>"
        periods[caller] += sample.period
        blocks[caller] += 1
    total = sum(periods.values())
    rows = []
    for caller, period in periods.most_common(limit):
        row = weighted_metric(period, blocks[caller], subset_period, whole_period)
        row["caller"] = caller
        row["share_of_marker_period_percent"] = period / total * 100.0 if total else None
        rows.append(row)
    return rows


def subset_summary(
    samples: list[Sample],
    name: str,
    label: str,
    marker: str,
    predicate: Callable[[str], bool],
    whole_period: int,
    limit: int,
) -> dict[str, Any]:
    matching = [
        sample for sample in samples if any(predicate(symbol) for symbol in sample.symbols)
    ]
    period = sum(sample.period for sample in matching)
    marker_symbols = sorted(
        {
            symbol
            for sample in matching
            for symbol in sample.symbols
            if predicate(symbol)
        }
    )
    profile_method_symbols = sorted(
        {
            symbol
            for sample in matching
            for symbol in sample.symbols
            if source_open_profile_method(symbol)
        }
    )
    sample_predicate = lambda sample: any(predicate(symbol) for symbol in sample.symbols)
    return {
        "name": name,
        "label": label,
        "marker": marker,
        "match_policy": (
            "exact demangled SourceBackedWorkbook owner and method terminal; "
            "compiler-generated closure descendants are accepted; *_with_limits is excluded"
        ),
        "observed_marker_symbols": marker_symbols,
        "observed_profile_method_symbols": profile_method_symbols,
        "scope": weighted_metric(period, len(matching), period, whole_period),
        "leaf_period_weighted_ranking": rank_symbols(
            samples, sample_predicate, period, whole_period, limit, leaf=True
        ),
        "inclusive_period_weighted_ranking": rank_symbols(
            samples, sample_predicate, period, whole_period, limit, leaf=False
        ),
        "direct_callers_of_marker": direct_callers(
            matching, predicate, period, whole_period, limit
        ),
        "marker_observed": bool(matching),
        "interpretation": (
            "Inclusive stack presence is an observed CPU subset. A zero scope "
            "means no marker was retained by perf, not proof of zero work."
        ),
    }


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


def add_unique(values: list[Any], value: Any) -> None:
    if value not in values:
        values.append(value)


def case_names(value: Any) -> list[str]:
    values: list[str] = []
    for text in collect_strings(value):
        for case in CASE_RE.findall(text):
            add_unique(values, case)
    return sorted(values)


def numeric_values_for_keys(value: dict[str, Any], names: set[str], contains: str) -> list[Any]:
    values: list[Any] = []
    for key, child in value.items():
        key_lower = str(key).lower()
        if (
            key_lower in names
            or key_lower.endswith("_" + contains)
            or key_lower.startswith(contains + "_")
        ) and isinstance(child, (int, float)) and not isinstance(child, bool):
            add_unique(values, child)
    return values


def normalize_number(value: Any) -> int | float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value) if int(value) == value else value
    return None


def profile_observations(data: Any) -> dict[str, Any]:
    if not isinstance(data, dict):
        return {
            "status": "missing_or_invalid",
            "cases": [],
            "samples": [],
            "warmups": [],
            "events": [],
            "frequencies": [],
            "call_graphs": [],
            "cpus": [],
            "source_field": {},
        }
    config = data.get("configuration")
    if not isinstance(config, dict):
        config = {}
    observations: dict[str, Any] = {
        "status": "parsed",
        "cases": case_names(config.get("cases", [])),
        "samples": numeric_values_for_keys(
            config,
            {"samples", "sample_count", "samples_per_case", "record_samples"},
            "samples",
        ),
        "warmups": numeric_values_for_keys(
            config,
            {"warmup", "warmups", "warmup_iterations_per_case"},
            "warmup",
        ),
        "events": [],
        "frequencies": [],
        "call_graphs": [],
        "cpus": [],
        "source_field": {
            "field_path": "results[*].source",
            "results": 0,
            "present": 0,
            "missing": 0,
            "invalid_results": 0,
            "null": 0,
            "object": 0,
            "other": 0,
            "null_is_accepted": True,
            "missing_is_unverified": True,
        },
    }
    for key, child in config.items():
        key_lower = str(key).lower()
        if isinstance(child, str):
            if "event" in key_lower:
                add_unique(observations["events"], normalize_event(child))
            if "call" in key_lower and "graph" in key_lower:
                add_unique(observations["call_graphs"], child)
        if isinstance(child, (int, float)) and not isinstance(child, bool):
            if key_lower in {"frequency", "freq", "frequency_hz"} or key_lower.endswith("_hz"):
                add_unique(observations["frequencies"], normalize_number(child))
            if key_lower in {"cpu", "cpu_id", "cpu_affinity"}:
                add_unique(observations["cpus"], normalize_number(child))
    for key in ("event", "perf_event"):
        if isinstance(data.get(key), str):
            add_unique(observations["events"], normalize_event(data[key]))
    results = data.get("results")
    if isinstance(results, list):
        observations["source_field"]["results"] = len(results)
        for result in results:
            if not isinstance(result, dict):
                observations["source_field"]["invalid_results"] += 1
                continue
            result_case = result.get("case")
            if isinstance(result_case, str):
                for case in CASE_RE.findall(result_case):
                    add_unique(observations["cases"], case)
            if "source" not in result:
                observations["source_field"]["missing"] += 1
                continue
            observations["source_field"]["present"] += 1
            source = result["source"]
            if source is None:
                observations["source_field"]["null"] += 1
            elif isinstance(source, dict):
                observations["source_field"]["object"] += 1
            else:
                observations["source_field"]["other"] += 1
    for key in ("cases", "events", "call_graphs", "frequencies", "cpus"):
        observations[key] = sorted(
            {value for value in observations[key] if value is not None}, key=str
        )
    return observations


def profile_summary(data: Any, path: Path | None) -> dict[str, Any]:
    if not isinstance(data, dict):
        return {
            "status": "missing_or_invalid",
            "file": file_binding(path) if path and path.is_file() else None,
        }
    config = data.get("configuration")
    results = data.get("results")
    result_cases = []
    if isinstance(results, list):
        result_cases = sorted(
            {
                case
                for item in results
                if isinstance(item, dict)
                for case in CASE_RE.findall(str(item.get("case", "")))
            }
        )
    binary = data.get("binary_identity")
    environment = data.get("environment")
    return {
        "status": "parsed",
        "file": file_binding(path) if path and path.is_file() else None,
        "schema_version": data.get("schema_version"),
        "tool": data.get("tool"),
        "binary_identity": (
            {
                key: binary.get(key)
                for key in ("path", "binary_sha256", "binary_bytes", "profile")
                if key in binary
            }
            if isinstance(binary, dict)
            else None
        ),
        "environment": (
            {
                key: environment.get(key)
                for key in ("git_revision", "git_worktree_dirty", "rustc_version", "cpu_affinity")
                if key in environment
            }
            if isinstance(environment, dict)
            else None
        ),
        "configuration": (
            {
                key: config.get(key)
                for key in (
                    "samples_per_case",
                    "warmup_iterations_per_case",
                    "event",
                    "frequency",
                    "frequency_hz",
                    "call_graph",
                    "cases",
                )
                if key in config
            }
            if isinstance(config, dict)
            else None
        ),
        "result_summary": {
            "count": len(results) if isinstance(results, list) else None,
            "cases": result_cases,
            "source_field": profile_observations(data)["source_field"],
            "source_null_is_valid": True,
        },
    }


def catalog_summary(data: Any, path: Path | None) -> dict[str, Any]:
    if not isinstance(data, dict):
        return {
            "status": "missing_or_invalid",
            "file": file_binding(path) if path and path.is_file() else None,
        }
    build = data.get("build")
    bindings = data.get("case_bindings")
    bound_cases = (
        sorted(
            {
                case
                for item in bindings
                if isinstance(item, dict)
                for case in CASE_RE.findall(str(item.get("case", "")))
            }
        )
        if isinstance(bindings, list)
        else []
    )
    null_sources = 0
    source_fields = 0
    for item in recursive_dicts(data):
        for key, child in item.items():
            if str(key).lower() == "source":
                source_fields += 1
                if child is None:
                    null_sources += 1
    return {
        "status": "parsed",
        "file": file_binding(path) if path and path.is_file() else None,
        "manifest_version": data.get("manifest_version"),
        "manifest_kind": data.get("manifest_kind"),
        "catalog_id": data.get("catalog_id"),
        "catalog_sha256": data.get("catalog_sha256"),
        "content_set_sha256": data.get("content_set_sha256"),
        "build": (
            {
                key: build.get(key)
                for key in (
                    "tool",
                    "tool_version",
                    "git_revision",
                    "git_worktree_dirty",
                    "source_files",
                )
                if key in build
            }
            if isinstance(build, dict)
            else None
        ),
        "case_bindings": bound_cases,
        "source_fields": {
            "count": source_fields,
            "null": null_sources,
            "null_is_accepted": True,
        },
    }


def command_records(value: Any) -> Iterable[dict[str, Any]]:
    if isinstance(value, dict):
        if any(key in value for key in ("argv", "command", "command_line", "cmd")):
            yield value
        for child in value.values():
            yield from command_records(child)
    elif isinstance(value, list):
        for child in value:
            yield from command_records(child)


def record_argv(record: dict[str, Any]) -> list[str]:
    for key in ("argv", "command", "command_line", "cmd"):
        value = record.get(key)
        if isinstance(value, list):
            return [str(token) for token in value]
        if isinstance(value, str):
            try:
                return shlex.split(value)
            except ValueError:
                return value.split()
    return []


def choose_profile_commands(
    records: list[dict[str, Any]], selected_case: str | None
) -> list[dict[str, Any]]:
    preferred = []
    for record in records:
        label = str(record.get("label", "")).lower()
        argv = record_argv(record)
        joined = " ".join(argv)
        if (
            "profile" in label
            or "postprocess" in label
            or "script" in label
            or "self" in label
            or ("perf" in joined and "record" in joined)
        ):
            preferred.append(record)
        if "owned-profile" in label:
            preferred.append(record)
    if preferred:
        unique: list[dict[str, Any]] = []
        seen: set[int] = set()
        for record in preferred:
            identifier = id(record)
            if identifier not in seen:
                seen.add(identifier)
                unique.append(record)
        return unique
    return records


def has_subcommand(argv: list[str], executable: str, subcommand: str) -> bool:
    executable_lower = executable.lower()
    return any(
        Path(token).name.lower() == executable_lower
        and index + 1 < len(argv)
        and argv[index + 1] == subcommand
        for index, token in enumerate(argv)
    )


def strict_identity_check(
    expected: Any, observed: Iterable[Any], label: str
) -> dict[str, Any]:
    """Verify that one command surface contains exactly one requested value."""

    values = sorted(set(observed), key=str)
    if expected is None:
        return {"status": "not_requested", "expected": None, "observed": values}
    if not values:
        return {
            "status": "unverified",
            "expected": expected,
            "observed": values,
            "message": f"{label} was not found in the capture command",
        }
    return {
        "status": "verified" if values == [expected] else "mismatch",
        "expected": expected,
        "observed": values,
    }


def strict_numeric_identity_check(
    expected: Any, observed: Iterable[Any], label: str
) -> dict[str, Any]:
    values = [
        normalized
        for value in observed
        if (normalized := normalize_number(value)) is not None
    ]
    return strict_identity_check(expected, values, label)


def profile_source_absence_check(source_field: dict[str, Any]) -> dict[str, Any]:
    """Validate the optional CaseResult source field without accepting counters."""

    results = source_field.get("results", 0)
    missing = source_field.get("missing", 0)
    invalid_results = source_field.get("invalid_results", 0)
    null = source_field.get("null", 0)
    nonnull = source_field.get("object", 0) + source_field.get("other", 0)
    observed = {
        "results": results,
        "present": source_field.get("present", 0),
        "missing": missing,
        "invalid_results": invalid_results,
        "null": null,
        "object": source_field.get("object", 0),
        "other": source_field.get("other", 0),
    }
    if results != 1 or invalid_results:
        return {
            "status": "unverified" if results == 0 else "mismatch",
            "expected": (
                "exactly one result; direct source may be omitted or explicit null"
            ),
            "observed": observed,
            "null_is_accepted": True,
            "omitted_is_accepted": True,
        }
    if missing == 1 and null == 0 and nonnull == 0:
        status = "accepted_absent"
        absence = "omitted"
    elif missing == 0 and null == 1 and nonnull == 0:
        status = "accepted_null"
        absence = "explicit_null"
    elif nonnull:
        status = "mismatch"
        absence = "nonnull_rejected"
    else:
        status = "unverified"
        absence = "ambiguous"
    return {
        "status": status,
        "expected": (
            "exactly one result; direct source may be omitted or explicit null; "
            "nonnull source/counters are rejected"
        ),
        "observed": observed,
        "absence_state": absence,
        "null_is_accepted": True,
        "omitted_is_accepted": True,
    }


def profile_result_identity(data: Any, selected_case: str | None) -> dict[str, Any]:
    """Require the single owned profile result and its selected case."""

    results = data.get("results") if isinstance(data, dict) else None
    if not isinstance(results, list):
        return {
            "status": "unverified",
            "row_count": {
                "status": "unverified",
                "expected": 1,
                "observed": None,
            },
            "case": strict_identity_check(selected_case, [], "profile result case"),
            "source": profile_source_absence_check({}),
        }
    row_count = {
        "status": "verified" if len(results) == 1 else (
            "mismatch" if len(results) > 1 else "unverified"
        ),
        "expected": 1,
        "observed": len(results),
    }
    cases = [
        item.get("case")
        for item in results
        if isinstance(item, dict) and isinstance(item.get("case"), str)
    ]
    case_check = strict_identity_check(selected_case, cases, "profile result case")
    source_field = profile_observations(data).get("source_field", {})
    source_check = profile_source_absence_check(source_field)
    checks = {"row_count": row_count, "case": case_check, "source": source_check}
    return {
        "status": (
            "verified"
            if row_count["status"] == "verified"
            and case_check["status"] == "verified"
            and source_check["status"] in {"accepted_absent", "accepted_null"}
            else (
                "mismatch"
                if any(check["status"] == "mismatch" for check in checks.values())
                else "unverified"
            )
        ),
        "checks": checks,
        "result_cases": cases,
    }


def capture_command_identity(
    records: list[dict[str, Any]], contract: dict[str, Any], selected_case: str | None
) -> dict[str, Any]:
    """Require the uniquely labelled owned-profile perf record command."""

    candidates = [
        record
        for record in records
        if str(record.get("label", "")).lower() == "owned-profile"
        and has_subcommand(record_argv(record), "perf", "record")
    ]
    observed_records = [
        {
            "label": record.get("label"),
            "argv": record_argv(record),
        }
        for record in candidates
    ]
    result: dict[str, Any] = {
        "status": "unverified",
        "expected": {
            "label": "owned-profile",
            "perf_subcommand": "record",
            "case": selected_case,
            "event": contract.get("event"),
            "frequency": contract.get("frequency"),
            "samples": contract.get("samples"),
            "warmup": contract.get("warmup"),
            "call_graph": contract.get("call_graph"),
            "cpu": contract.get("cpu"),
        },
        "candidate_count": len(candidates),
        "observed_records": observed_records,
    }
    if not candidates:
        result["message"] = (
            "No uniquely labelled owned-profile perf record command was found"
        )
        return result
    if len(candidates) != 1:
        result["status"] = "mismatch"
        result["message"] = "Expected exactly one owned-profile perf record command"
        return result

    observations = command_observations(candidates)
    checks = {
        "case": strict_identity_check(
            selected_case, observations["cases"], "capture case"
        ),
        "event": strict_identity_check(
            contract.get("event"), observations["events"], "capture event"
        ),
        "frequency": strict_numeric_identity_check(
            contract.get("frequency"), observations["frequencies"], "capture frequency"
        ),
        "samples": strict_numeric_identity_check(
            contract.get("samples"), observations["samples"], "capture sample count"
        ),
        "warmup": strict_numeric_identity_check(
            contract.get("warmup"), observations["warmups"], "capture warmup count"
        ),
        "call_graph": strict_identity_check(
            contract.get("call_graph"), observations["call_graphs"], "capture call graph"
        ),
        "cpu": strict_numeric_identity_check(
            contract.get("cpu"), observations["cpus"], "capture CPU affinity"
        ),
    }
    result["checks"] = checks
    result["observations"] = observations
    result["status"] = (
        "verified"
        if all(check["status"] == "verified" for check in checks.values())
        else "mismatch"
    )
    return result


def command_observations(records: list[dict[str, Any]]) -> dict[str, Any]:
    observations: dict[str, Any] = {
        "status": "parsed" if records else "missing_or_empty",
        "records": len(records),
        "selected_labels": [str(record.get("label")) for record in records if "label" in record],
        "cases": [],
        "samples": [],
        "warmups": [],
        "events": [],
        "frequencies": [],
        "call_graphs": [],
        "cpus": [],
        "no_inline": False,
        "revisions": [],
        "source_statuses": [],
        "exit_codes": [],
    }

    def add_case_text(text: str) -> None:
        for case in CASE_RE.findall(text):
            add_unique(observations["cases"], case)

    for record in records:
        argv = record_argv(record)
        for index, token in enumerate(argv):
            add_case_text(token)
            next_token = argv[index + 1] if index + 1 < len(argv) else None
            if token in {"--case", "--cases"} and next_token:
                add_case_text(next_token)
            if token in {"-e", "--event"} and next_token:
                add_unique(observations["events"], normalize_event(next_token))
            elif token.startswith("-e") and len(token) > 2:
                add_unique(observations["events"], normalize_event(token[2:]))
            elif token.startswith("--event="):
                add_unique(observations["events"], normalize_event(token.split("=", 1)[1]))
            if token in {"--samples", "--sample-count"} and next_token:
                try:
                    add_unique(observations["samples"], int(next_token.rstrip(",")))
                except ValueError:
                    pass
            elif token.startswith("--samples=") or token.startswith("--sample-count="):
                try:
                    add_unique(observations["samples"], int(token.split("=", 1)[1]))
                except ValueError:
                    pass
            if token in {"--warmup", "--warmups", "--warmup-iterations", "--warmup_iterations"} and next_token:
                try:
                    add_unique(observations["warmups"], int(next_token.rstrip(",")))
                except ValueError:
                    pass
            elif token.startswith("--warmup=") or token.startswith("--warmup-iterations="):
                try:
                    add_unique(observations["warmups"], int(token.split("=", 1)[1]))
                except ValueError:
                    pass
            if token in {"-F", "--freq", "--frequency"} and next_token:
                try:
                    add_unique(observations["frequencies"], int(next_token.rstrip(",")))
                except ValueError:
                    pass
            elif re.fullmatch(r"-F\d+", token):
                add_unique(observations["frequencies"], int(token[2:]))
            elif token.startswith("--freq=") or token.startswith("--frequency="):
                try:
                    add_unique(observations["frequencies"], int(token.split("=", 1)[1]))
                except ValueError:
                    pass
            if token == "--call-graph" and next_token:
                add_unique(observations["call_graphs"], next_token)
            elif token.startswith("--call-graph="):
                add_unique(observations["call_graphs"], token.split("=", 1)[1])
            if token == "--no-inline":
                observations["no_inline"] = True
            if token == "-c" and next_token and argv[index - 1:index] == ["taskset"]:
                try:
                    add_unique(observations["cpus"], int(next_token))
                except ValueError:
                    pass
            if token in {"--cpu", "--cpu-affinity"} and next_token:
                try:
                    add_unique(observations["cpus"], int(next_token))
                except ValueError:
                    pass
        for key in ("revision", "git_revision", "commit"):
            value = record.get(key)
            if isinstance(value, str):
                add_unique(observations["revisions"], value)
        for key in ("source_status", "status"):
            if key in record:
                add_unique(observations["source_statuses"], record[key])
        if "exit_code" in record:
            observations["exit_codes"].append(record["exit_code"])
    for key in (
        "cases",
        "samples",
        "warmups",
        "events",
        "frequencies",
        "call_graphs",
        "cpus",
        "revisions",
        "source_statuses",
    ):
        observations[key] = sorted(set(observations[key]), key=str)
    return observations


def load_json(path: Path | None) -> tuple[Any | None, str | None]:
    if path is None:
        return None, "path not supplied"
    if not path.is_file():
        return None, "file is absent"
    try:
        if path.suffix == ".zst":
            raw = subprocess.check_output(["zstd", "-dc", str(path)], stderr=subprocess.STDOUT)
            value = json.loads(raw.decode("utf-8"))
        else:
            value = json.loads(path.read_text(encoding="utf-8"))
        return value, None
    except (
        OSError,
        subprocess.CalledProcessError,
        UnicodeDecodeError,
        json.JSONDecodeError,
    ) as error:
        return None, str(error)


def protocol_contract(data: Any) -> dict[str, Any]:
    """Read only explicit protocol fields; missing fields are never defaulted."""

    result: dict[str, Any] = {}
    if not isinstance(data, dict):
        return result
    profile = data.get("profile")
    if isinstance(profile, dict):
        mapping = {
            "case": "case",
            "samples": "samples",
            "warmup": "warmup",
            "event": "event",
            "frequency": "frequency",
            "call_graph": "call_graph",
        }
        for source_key, target_key in mapping.items():
            if source_key in profile:
                value = profile[source_key]
                if target_key == "event" and isinstance(value, str):
                    value = normalize_event(value)
                result[target_key] = value
    if isinstance(data.get("cpu"), int):
        result["cpu"] = data["cpu"]
    return result


def metadata_revision(data: Any) -> str | None:
    """Read the build revision from metadata containers before nested schemas."""

    if not isinstance(data, dict):
        return None
    for container_name in ("environment", "build", "source_before"):
        container = data.get(container_name)
        if isinstance(container, dict):
            for key in ("git_revision", "revision", "commit"):
                value = container.get(key)
                if isinstance(value, str):
                    return value
    for key in ("git_revision", "revision", "commit"):
        value = data.get(key)
        if isinstance(value, str):
            return value
    return None


def candidate_identity_summary(data: Any, path: Path | None) -> dict[str, Any]:
    if not isinstance(data, dict):
        return {
            "status": "missing_or_invalid",
            "file": file_binding(path) if path and path.is_file() else None,
            "source_status_null_is_accepted": True,
        }
    revision = None
    for key in ("revision", "git_revision", "commit"):
        value = data.get(key)
        if isinstance(value, str):
            revision = value
            break
    source_status_present = "source_status" in data or "status" in data
    source_status = data.get("source_status", data.get("status"))
    if not source_status_present:
        source_state = "unverified"
    elif source_status is None:
        source_state = "accepted_null"
    elif source_status == "":
        source_state = "clean"
    else:
        source_state = "observed"
    binaries = data.get("binaries")
    binary_summary = {}
    if isinstance(binaries, dict):
        for name, value in binaries.items():
            if isinstance(value, dict):
                binary_summary[str(name)] = {
                    key: value.get(key)
                    for key in ("path", "bytes", "sha256", "binary_sha256")
                    if key in value
                }
    return {
        "status": "parsed",
        "file": file_binding(path) if path and path.is_file() else None,
        "revision": revision,
        "source_status": source_status,
        "source_status_state": source_state,
        "source_status_null_is_accepted": True,
        "rustc": data.get("rustc"),
        "cargo": data.get("cargo"),
        "build_environment": data.get("build_environment"),
        "binaries": binary_summary,
    }


def source_binding(root: Path, revision: str | None) -> dict[str, Any]:
    specs = [
        (
            "crates/litchi-xls/src/workbook/source.rs",
            {
                "source_backed_type": re.compile(r"\bpub struct SourceBackedWorkbook\b"),
                "from_read_at": re.compile(r"\bpub fn from_read_at\("),
                "cell_value_by_index": re.compile(r"\bpub fn cell_value_by_index\("),
            },
        ),
        (
            "tools/perf-baseline/src/lib.rs",
            {
                "owned_selector_variants": re.compile(
                    r"\bXlsOwnedSourceOpen(?:ListWorksheets|OneCell)?\b"
                ),
                "owned_selector_names": re.compile(
                    r"\bxls_owned_source_open(?:_list_worksheets|_one_cell)?\b"
                ),
                "owned_case_dispatch": re.compile(r"\brun_xls_owned_source_case\("),
                "owned_source_construction": re.compile(r"\bOwnedSource::new\("),
            },
        ),
    ]
    result: dict[str, Any] = {
        "status": "unbound",
        "revision": revision,
        "files": [],
        "errors": [],
        "interpretation": (
            "Git blobs are bound to the candidate identity revision. The "
            "working tree is never substituted for a missing or failed blob."
        ),
    }
    if not revision:
        result["errors"].append("candidate revision is absent")
        return result
    if not HASH_RE.fullmatch(revision):
        result["errors"].append(f"candidate revision is not a Git object id: {revision!r}")
        return result
    for relative, checks in specs:
        try:
            blob = subprocess.check_output(
                ["git", "-C", str(root), "show", f"{revision}:{relative}"],
                stderr=subprocess.STDOUT,
            )
        except (OSError, subprocess.CalledProcessError) as error:
            result["errors"].append(f"{relative}: {error}")
            continue
        lines = blob.decode("utf-8", errors="replace").splitlines()
        found = {
            name: [line_no for line_no, line in enumerate(lines, 1) if pattern.search(line)]
            for name, pattern in checks.items()
        }
        result["files"].append(
            {
                "path": relative,
                "bytes": len(blob),
                "sha256": sha256_bytes(blob),
                "checks": found,
                "all_checks_observed": all(found.values()),
            }
        )
    result["status"] = (
        "bound"
        if len(result["files"]) == len(specs)
        and all(item["all_checks_observed"] for item in result["files"])
        else "partial"
    )
    return result


def parse_report(path: Path | None) -> dict[str, Any]:
    if path is None or not path.is_file():
        return {
            "status": "missing",
            "file": None,
            "warnings": ["perf report self stdout is absent"],
        }
    stats: Counter[str] = Counter()
    events: list[str] = []
    rows: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8", errors="replace") as source:
        for raw in source:
            stats["lines"] += 1
            line = raw.rstrip("\r\n")
            event_match = REPORT_EVENT_RE.search(line)
            if event_match:
                add_unique(events, normalize_event(event_match.group("event")))
            if line.lstrip().startswith("#"):
                stats["metadata_lines"] += 1
                sample_match = REPORT_SAMPLES_RE.search(line)
                if sample_match:
                    stats["sample_count_headers"] += 1
                continue
            match = REPORT_ROW_RE.match(line)
            if not match:
                if line.strip():
                    stats["unparsed_nonempty_lines"] += 1
                continue
            stats["percentage_rows"] += 1
            if len(rows) < 100:
                rows.append(
                    {
                        "percent": float(match.group("percent")),
                        "self_percent": (
                            float(match.group("self")) if match.group("self") else None
                        ),
                        "text": match.group("rest").strip(),
                    }
                )
    return {
        "status": "parsed",
        "file": file_binding(path),
        "events": sorted(events),
        "stats": dict(stats),
        "top_percentage_rows": rows,
    }


def merge_observations(profile: dict[str, Any], commands: dict[str, Any]) -> dict[str, Any]:
    result = {}
    for key in ("cases", "samples", "warmups", "events", "frequencies", "call_graphs", "cpus"):
        result[key] = sorted(
            set(profile.get(key, [])) | set(commands.get(key, [])),
            key=str,
        )
    result["no_inline"] = bool(commands.get("no_inline", False))
    result["revisions"] = commands.get("revisions", [])
    result["source_statuses"] = commands.get("source_statuses", [])
    return result


def resolve_artifact(explicit: Path | None, defaults: list[Path]) -> Path | None:
    if explicit is not None:
        return explicit.resolve()
    for candidate in defaults:
        if candidate.is_file():
            return candidate.resolve()
    return defaults[0].resolve() if defaults else None


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--script", type=Path, default=DEFAULT_SCRIPT)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--capture", type=Path, default=DEFAULT_CAPTURE)
    parser.add_argument("--profile", type=Path, default=None)
    parser.add_argument("--catalog", type=Path, default=None)
    parser.add_argument("--commands", type=Path, default=None)
    parser.add_argument("--postprocess-commands", type=Path, default=None)
    parser.add_argument("--protocol", type=Path, default=None)
    parser.add_argument("--build-identity", type=Path, default=None)
    parser.add_argument("--repo", type=Path, default=DEFAULT_ROOT)
    parser.add_argument("--case", type=str, default=None)
    parser.add_argument("--top", type=int, default=40)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    if args.top <= 0:
        raise SystemExit("--top must be positive")
    script = args.script.resolve()
    report = args.report.resolve()
    capture = args.capture.resolve() if args.capture else None
    repo = args.repo.resolve()
    output = args.output.resolve()
    if not script.is_file():
        raise SystemExit(f"missing required perf script input: {script}")

    profile_path = resolve_artifact(
        args.profile,
        [capture / "owned-profile.json"] if capture else [DEFAULT_PROFILE],
    )
    catalog_path = resolve_artifact(
        args.catalog,
        [capture / "owned-profile.catalog.json"] if capture else [DEFAULT_CATALOG],
    )
    commands_path = resolve_artifact(
        args.commands,
        [capture / "commands.json"] if capture else [DEFAULT_COMMANDS],
    )
    postprocess_commands_path = resolve_artifact(
        args.postprocess_commands,
        [capture / "postprocess-commands.json"] if capture else [],
    )
    protocol_path = resolve_artifact(
        args.protocol,
        [capture / "protocol.json"] if capture else [DEFAULT_PROTOCOL],
    )
    build_path = resolve_artifact(
        args.build_identity,
        [capture / "candidate-build-identity.json"]
        if capture
        else [DEFAULT_BUILD_IDENTITY],
    )

    samples, parser_stats = parse_samples(script)
    if not samples:
        raise SystemExit("no cycles:u sample blocks parsed")
    whole_period = sum(sample.period for sample in samples)

    profile_data, profile_error = load_json(profile_path)
    catalog_data, catalog_error = load_json(catalog_path)
    commands_data, commands_error = load_json(commands_path)
    postprocess_commands_data, postprocess_commands_error = load_json(
        postprocess_commands_path
    )
    protocol_data, protocol_error = load_json(protocol_path)
    build_data, build_error = load_json(build_path)

    contract = protocol_contract(protocol_data)
    selected_case = args.case or contract.get("case") or None
    profile_obs = profile_observations(profile_data)
    all_records = list(command_records(commands_data))
    all_records.extend(command_records(postprocess_commands_data))
    selected_records = choose_profile_commands(all_records, selected_case)
    command_obs = command_observations(selected_records)
    capture_records = [
        record
        for record in selected_records
        if str(record.get("label", "")).lower() == "owned-profile"
        and has_subcommand(record_argv(record), "perf", "record")
    ]
    postprocess_records = [
        record
        for record in selected_records
        if id(record) not in {id(candidate) for candidate in capture_records}
    ]
    capture_obs = command_observations(capture_records)
    postprocess_obs = command_observations(postprocess_records)
    merged_obs = merge_observations(profile_obs, capture_obs)
    # ``--no-inline`` belongs to the separate perf script/report postprocess;
    # keep it visible without mixing report's --call-graph=none into capture
    # identity values.
    merged_obs["no_inline"] = postprocess_obs["no_inline"]
    if selected_case is None and len(merged_obs["cases"]) == 1:
        selected_case = merged_obs["cases"][0]

    candidate_summary = candidate_identity_summary(build_data, build_path)
    candidate_revision = candidate_summary.get("revision")
    profile_revision = metadata_revision(profile_data)
    catalog_build = catalog_data.get("build") if isinstance(catalog_data, dict) else None
    catalog_revision = (
        catalog_build.get("git_revision")
        if isinstance(catalog_build, dict)
        else None
    )
    source = source_binding(repo, candidate_revision)
    capture_command = capture_command_identity(
        selected_records, contract, selected_case
    )
    profile_result = profile_result_identity(profile_data, selected_case)
    candidate_revision_valid = isinstance(candidate_revision, str) and bool(
        HASH_RE.fullmatch(candidate_revision)
    )
    protocol_required_fields = (
        "case",
        "samples",
        "warmup",
        "event",
        "frequency",
        "call_graph",
        "cpu",
    )
    protocol_missing_fields = [
        field for field in protocol_required_fields if field not in contract
    ]
    protocol_is_valid = (
        protocol_error is None
        and not protocol_missing_fields
        and contract.get("event") == "cycles:u"
    )
    protocol_check = {
        "status": (
            "verified"
            if protocol_is_valid
            else ("mismatch" if protocol_error is None else "unverified")
        ),
        "expected_fields": list(protocol_required_fields),
        "missing_fields": protocol_missing_fields,
        "observed": contract,
        "event_must_be": "cycles:u",
    }
    parser_invalid_counts = {
        key: parser_stats.get(key, 0)
        for key in (
            "invalid_cycle_event_headers",
            "invalid_cycle_period_headers",
            "invalid_sample_blocks",
            "invalid_sample_period",
            "malformed_cycle_headers",
            "non_cycle_headers",
            "unparsed_frame_lines",
            "empty_stack_blocks",
            "lost_sample_count",
            "unquantified_lost_lines",
        )
    }
    parser_check = {
        "status": (
            "verified"
            if not any(parser_invalid_counts.values())
            and all(sample.event == "cycles:u" and sample.period > 0 for sample in samples)
            else "mismatch"
        ),
        "expected": "every retained sample is cycles:u with a positive period",
        "observed_invalid_counts": parser_invalid_counts,
        "retained_event_values": sorted({sample.event for sample in samples}),
        "nonpositive_retained_periods": sum(sample.period <= 0 for sample in samples),
    }

    identity_checks = {
        "protocol": protocol_check,
        "sample_parser": parser_check,
        "case": strict_identity_check(
            selected_case,
            capture_obs["cases"],
            "capture case",
        ),
        "event": strict_identity_check(
            contract.get("event"), capture_obs["events"], "capture event"
        ),
        "frequency": strict_numeric_identity_check(
            contract.get("frequency"),
            capture_obs["frequencies"],
            "capture frequency",
        ),
        "samples": strict_numeric_identity_check(
            contract.get("samples"),
            capture_obs["samples"],
            "capture sample count",
        ),
        "warmup": strict_numeric_identity_check(
            contract.get("warmup"),
            capture_obs["warmups"],
            "capture warmup count",
        ),
        "call_graph": strict_identity_check(
            contract.get("call_graph"),
            capture_obs["call_graphs"],
            "capture call graph",
        ),
        "cpu": strict_numeric_identity_check(
            contract.get("cpu"),
            capture_obs["cpus"],
            "capture CPU affinity",
        ),
        "no_inline": {
            "status": "verified" if postprocess_obs["no_inline"] else "unverified",
            "expected": True,
            "observed": postprocess_obs["no_inline"],
            "source": "postprocess command journal",
        },
        "capture_command": capture_command,
        "profile_result": profile_result,
        "candidate_revision": {
            "status": (
                "verified"
                if candidate_revision_valid
                else ("mismatch" if candidate_revision else "unverified")
            ),
            "expected": "candidate-build-identity revision",
            "observed": candidate_revision,
            "valid_git_object_id": candidate_revision_valid,
        },
        "profile_revision": {
            "status": (
                "verified"
                if candidate_revision
                and profile_revision
                and profile_revision == candidate_revision
                else (
                    "mismatch"
                    if candidate_revision and profile_revision
                    else "unverified"
                )
            ),
            "expected": candidate_revision,
            "observed": profile_revision,
            "missing_is_unverified": True,
        },
        "catalog_revision": {
            "status": (
                "verified"
                if candidate_revision
                and catalog_revision
                and catalog_revision == candidate_revision
                else (
                    "mismatch"
                    if candidate_revision and catalog_revision
                    else "unverified"
                )
            ),
            "expected": candidate_revision,
            "observed": catalog_revision,
            "missing_is_unverified": True,
        },
        "source_status": {
            "status": candidate_summary.get("source_status_state", "unverified"),
            "expected": "clean, absent, or explicit null",
            "observed": candidate_summary.get("source_status"),
            "null_is_accepted": True,
        },
        "profile_source_fields": {
            **profile_result["checks"]["source"],
        },
        "source_binding": {
            "status": "verified" if source["status"] == "bound" else "mismatch",
            "expected": "bound candidate-revision Git blobs",
            "observed": source["status"],
        },
    }
    required_identity_names = (
        "protocol",
        "capture_command",
        "sample_parser",
        "profile_result",
        "case",
        "event",
        "frequency",
        "samples",
        "warmup",
        "call_graph",
        "cpu",
        "no_inline",
        "candidate_revision",
        "profile_revision",
        "catalog_revision",
        "source_status",
        "source_binding",
    )
    accepted_identity_statuses = {
        "verified",
        "clean",
        "accepted_absent",
        "accepted_null",
    }
    required_identity_failures = [
        f"{name}: {identity_checks[name].get('status')}"
        for name in required_identity_names
        if identity_checks[name].get("status") not in accepted_identity_statuses
    ]
    required_identity = {
        "status": "verified" if not required_identity_failures else "failed",
        "accepted_statuses": sorted(accepted_identity_statuses),
        "checks": {
            name: identity_checks[name].get("status")
            for name in required_identity_names
        },
        "failures": required_identity_failures,
        "source_binding": source["status"],
    }

    subset_specs = [
        (
            "source_backed_from_read_at",
            "SourceBackedWorkbook::from_read_at production ancestors",
            "<litchi_xls::workbook::source::SourceBackedWorkbook>::from_read_at",
            source_open,
        ),
        (
            "source_backed_cell_value_by_index",
            "SourceBackedWorkbook::cell_value_by_index production ancestors",
            "<litchi_xls::workbook::source::SourceBackedWorkbook>::cell_value_by_index",
            source_cell,
        ),
    ]
    subsets = {
        name: subset_summary(
            samples,
            name,
            label,
            marker,
            predicate,
            whole_period,
            args.top,
        )
        for name, label, marker, predicate in subset_specs
    }

    unknown_period = parser_stats.get("unknown_frame_period", 0)
    empty_period = parser_stats.get("empty_stack_period", 0)
    accounting = {
        **parser_stats,
        "whole_process_weighted_event_period": whole_period,
        "whole_process_raw_stack_blocks": len(samples),
        "unknown_frame_share_of_whole_process_percent": (
            unknown_period / whole_period * 100.0 if whole_period else None
        ),
        "empty_stack_share_of_whole_process_percent": (
            empty_period / whole_period * 100.0 if whole_period else None
        ),
        "lost_is_not_in_sample_period": True,
        "interpretation": (
            "Unknown and empty values are counted explicitly. Lost records have "
            "no period attribution and are reported separately; unquantified "
            "lost diagnostics are not treated as zero."
        ),
    }

    metadata_warnings = []
    for name, error in (
        ("profile", profile_error),
        ("catalog", catalog_error),
        ("command_journal", commands_error),
        ("postprocess_command_journal", postprocess_commands_error),
        ("protocol", protocol_error),
        ("candidate_build_identity", build_error),
    ):
        if error:
            metadata_warnings.append(f"{name}: {error}")

    summary: dict[str, Any] = {
        "schema": SCHEMA,
        "purpose": (
            "Single plain OwnedSource XLS open+one-cell CPU attribution; "
            "descriptive observer/profile evidence only"
        ),
        "timing_semantics": {
            "phase_latency": False,
            "weight": "perf cycles:u sample period",
            "whole_process_scope": (
                "all parsed cycles:u blocks in the supplied perf script export, "
                "including inherited child process samples"
            ),
            "subset_scope": "inclusive presence of exact demangled production ancestor marker",
            "cpu": contract.get("cpu"),
        },
        "requested_capture_contract": contract,
        "inputs": {
            "perf_script": file_binding(script, repo),
            "perf_report": file_binding(report, repo) if report.is_file() else None,
            "profile": (
                file_binding(profile_path)
                if profile_path and profile_path.is_file()
                else None
            ),
            "catalog": (
                file_binding(catalog_path)
                if catalog_path and catalog_path.is_file()
                else None
            ),
            "command_journal": (
                file_binding(commands_path)
                if commands_path and commands_path.is_file()
                else None
            ),
            "postprocess_command_journal": (
                file_binding(postprocess_commands_path)
                if postprocess_commands_path and postprocess_commands_path.is_file()
                else None
            ),
            "candidate_build_identity": (
                file_binding(build_path)
                if build_path and build_path.is_file()
                else None
            ),
            "protocol": (
                file_binding(protocol_path)
                if protocol_path and protocol_path.is_file()
                else None
            ),
            "parser": file_binding(Path(__file__).resolve()),
        },
        "capture_identity": {
            "selected_case": selected_case,
            "profile": profile_summary(profile_data, profile_path),
            "catalog": catalog_summary(catalog_data, catalog_path),
            "command_journal": {
                "file": (
                    file_binding(commands_path)
                    if commands_path and commands_path.is_file()
                    else None
                ),
                "postprocess_file": (
                    file_binding(postprocess_commands_path)
                    if postprocess_commands_path and postprocess_commands_path.is_file()
                    else None
                ),
                "all_records": len(all_records),
                "selected_records": len(selected_records),
                "capture_records": len(capture_records),
                "postprocess_records": len(postprocess_records),
                "capture_observations": capture_obs,
                "postprocess_observations": postprocess_obs,
                "capture_command": capture_command,
                "observations": command_obs,
            },
            "candidate_build_identity": candidate_summary,
            "observed_metadata": {
                "profile": profile_obs,
                "capture_commands": capture_obs,
                "postprocess_commands": postprocess_obs,
                "commands": command_obs,
                "merged": merged_obs,
            },
            "revisions": {
                "candidate_build_identity": candidate_revision,
                "profile": profile_revision,
                "catalog": catalog_revision,
            },
            "checks": identity_checks,
            "required_identity": required_identity,
            "warnings": metadata_warnings,
            "nullable_source_policy": (
                "CaseResult.source is optional metadata: serde omission and an "
                "explicit null both mean source-summary absence and are accepted; "
                "nonnull source objects or synthetic counters fail identity."
            ),
        },
        "source_binding": source,
        "perf_report": parse_report(report),
        "sample_parser": accounting,
        "whole_process": {
            "scope": "all parsed cycles:u samples",
            "weighted_event_period": whole_period,
            "raw_stack_blocks": len(samples),
            "leaf_period_weighted_ranking": rank_symbols(
                samples,
                lambda _sample: True,
                whole_period,
                whole_period,
                args.top,
                leaf=True,
            ),
            "inclusive_period_weighted_ranking": rank_symbols(
                samples,
                lambda _sample: True,
                whole_period,
                whole_period,
                args.top,
                leaf=False,
            ),
        },
        "production_ancestor_subsets": subsets,
        "limitations": [
            "Sample periods are weighted cycles:u event periods, not wall time, operation latency, or an Instant timer phase.",
            "The whole-process export can include process startup, harness setup/oracle work, and inherited fresh children.",
            "A missing production marker means no matching frame was retained; it does not prove zero CPU work.",
            "Unknown, malformed, empty, and lost records are explicit accounting categories and are not silently folded into a production subset.",
            "The OwnedSource selector names are observed from metadata; this parser does not require a fixed six-case harness list.",
            "No source/corpus summary is inferred from a nullable profile result source field.",
        ],
        "interpretation": {
            "callchain_order": "perf script frames are leaf to caller/root",
            "subset_overlap": "Each subset is inclusive and independent; a sample may contribute to both marker subsets.",
            "source_route": "from_read_at and cell_value_by_index are exact production ancestor markers, not timer boundaries.",
        },
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {
                "output": str(output),
                "case": selected_case,
                "whole_process_period": whole_period,
                "whole_process_blocks": len(samples),
                "source_open_period": subsets["source_backed_from_read_at"]["scope"]["weighted_event_period"],
                "source_cell_period": subsets["source_backed_cell_value_by_index"]["scope"]["weighted_event_period"],
                "unknown_frame_period": unknown_period,
                "empty_stack_period": empty_period,
                "lost_sample_count": parser_stats.get("lost_sample_count", 0),
                "source_binding": source["status"],
                "required_identity": required_identity["status"],
                "required_identity_failures": required_identity_failures,
            },
            sort_keys=True,
        )
    )
    return 0 if not required_identity_failures else 2


if __name__ == "__main__":
    raise SystemExit(main())
