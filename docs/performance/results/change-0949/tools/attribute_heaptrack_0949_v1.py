#!/usr/bin/env python3
"""Offline, fail-closed heaptrack attribution for the 0949 diagnostic.

The input is the decoder output from ``decode_heaptrack_0949_v1.py``.  Every
trace is assigned exactly one nearest bucket in separate physical and inline
resolutions.  Publication is a stack-symbol diagnostic selected only when a
retained stack contains an identifiable source-replacement audit symbol; it
is not a time interval.  This helper never launches a workload or profiler.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


SCHEMA = "litchi.performance.0949.heaptrack-attribution.v1"
RUN_SCHEMA = "litchi.performance.0949.heaptrack-attribution-run.v1"
DECODER_SCHEMA = "litchi.performance.0949.heaptrack-raw-decode.v1"
DECODER_RUN_SCHEMA = "litchi.performance.0949.heaptrack-raw-decode-run.v1"
LABEL_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$")
EXPECTED_CATEGORIES = ("window", "namespace", "xml_audit", "writer", "other")


class AttributionError(RuntimeError):
    """An input, conservation, or output-custody check failed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AttributionError(message)


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(),
            f"not a regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(),
            f"not a regular file: {path}")
    return {"path": str(path.absolute()), "bytes": path.stat().st_size,
            "sha256": sha(path)}


def write_json_once(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(),
            f"refusing to overwrite output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False) + "\n").encode()
    with path.open("xb") as stream:
        stream.write(encoded)
        stream.flush()
        os.fsync(stream.fileno())


def write_text_once(path: Path, text: str) -> None:
    require(not path.exists() and not path.is_symlink(),
            f"refusing to overwrite output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write(text.encode("utf-8"))
        stream.flush()
        os.fsync(stream.fileno())


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(),
            f"missing JSON input: {path}")

    def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result

    try:
        return json.loads(path.read_text(encoding="utf-8"),
                          object_pairs_hook=reject_duplicates)
    except (OSError, json.JSONDecodeError) as error:
        raise AttributionError(f"cannot read JSON {path}: {error}") from error


def load_config(path: Path) -> dict[str, Any]:
    config = read_json(path)
    require(isinstance(config, dict)
            and config.get("schema") == "litchi.performance.0949.heaptrack-attribution-config.v1",
            "attribution config schema changed")
    order = config.get("category_order")
    require(order == list(EXPECTED_CATEGORIES),
            "attribution category order changed")
    categories = config.get("categories")
    require(isinstance(categories, dict)
            and all(isinstance(categories.get(name), dict) for name in EXPECTED_CATEGORIES[:-1]),
            "attribution categories are incomplete")
    symbols = config.get("publish_audit_symbols")
    require(isinstance(symbols, list) and symbols
            and all(isinstance(value, str) and value for value in symbols),
            "publish audit symbol list is empty or malformed")
    return config


def function_name(location: Any) -> str | None:
    if not isinstance(location, dict):
        return None
    value = location.get("function")
    return value if isinstance(value, str) and value else None


def file_name(location: Any) -> str | None:
    if not isinstance(location, dict):
        return None
    value = location.get("file")
    return value if isinstance(value, str) and value else None


def frame_locations(frame: Any, mode: str) -> list[dict[str, Any]]:
    if not isinstance(frame, dict):
        return []
    locations = frame.get("locations")
    if not isinstance(locations, list):
        return []
    valid = [location for location in locations if isinstance(location, dict)]
    return valid[:1] if mode == "physical" else valid


def category_for_location(location: dict[str, Any], config: dict[str, Any]) -> str | None:
    function = function_name(location) or ""
    file = file_name(location) or ""
    for category in EXPECTED_CATEGORIES[:-1]:
        matcher = config["categories"][category]
        function_hit = any(token in function for token in matcher.get("function_substrings", []))
        file_hit = any(token in file for token in matcher.get("file_substrings", []))
        if matcher.get("file_only") is True:
            hit = file_hit
        else:
            hit = function_hit or file_hit
        if hit:
            return category
    return None


def classify_stack(stack: list[dict[str, Any]], config: dict[str, Any],
                   mode: str) -> dict[str, Any]:
    require(mode in ("physical", "inline"), f"unknown resolution mode {mode}")
    unknown = False
    for item in stack:
        frame = item.get("frame") if isinstance(item, dict) else None
        locations = frame_locations(frame, mode)
        if not locations:
            unknown = True
            continue
        for location_index, location in enumerate(locations):
            if function_name(location) is None:
                unknown = True
            category = category_for_location(location, config)
            if category is not None:
                return {
                    "category": category,
                    "frame_trace_id": item.get("trace_id"),
                    "location_index": location_index,
                    "function": function_name(location),
                    "file": file_name(location),
                    "line": location.get("line"),
                    "unknown_before_match": unknown,
                }
    return {
        "category": "other",
        "frame_trace_id": None,
        "location_index": None,
        "function": None,
        "file": None,
        "line": None,
        "unknown_before_match": unknown,
    }


def publish_matches(stack: list[dict[str, Any]], symbols: list[str]) -> list[dict[str, Any]]:
    # Longest first prevents verify_source_replacement from hiding its
    # _with_spans specialization in the evidence.
    ordered = sorted(symbols, key=len, reverse=True)
    matches: list[dict[str, Any]] = []
    for item in stack:
        frame = item.get("frame") if isinstance(item, dict) else None
        locations = frame_locations(frame, "inline")
        for location_index, location in enumerate(locations):
            function = function_name(location) or ""
            for symbol in ordered:
                if symbol in function:
                    matches.append({
                        "symbol": symbol,
                        "function": function,
                        "file": file_name(location),
                        "line": location.get("line"),
                        "frame_trace_id": item.get("trace_id"),
                        "location_index": location_index,
                    })
                    break
    return matches


def read_size_aggregates(path: Path) -> dict[int, tuple[int, int]]:
    rows: dict[int, tuple[int, int]] = {}
    require(path.is_file() and not path.is_symlink(), f"missing size aggregate: {path}")
    with path.open(encoding="utf-8") as stream:
        require(stream.readline().rstrip("\n") == "size_bytes\tcalls\trequested_bytes",
                f"unexpected size aggregate header: {path}")
        for line_number, line in enumerate(stream, 2):
            parts = line.rstrip("\n").split("\t")
            require(len(parts) == 3, f"malformed size row {path}:{line_number}")
            try:
                size, calls, requested = (int(part) for part in parts)
            except ValueError as error:
                raise AttributionError(f"non-integer size row {path}:{line_number}") from error
            require(size >= 0 and calls > 0 and requested == size * calls,
                    f"size arithmetic mismatch {path}:{line_number}")
            require(size not in rows, f"duplicate size row {path}:{size}")
            rows[size] = (calls, requested)
    return rows


def read_trace_size_aggregates(path: Path) -> dict[int, dict[int, int]]:
    rows: dict[int, dict[int, int]] = defaultdict(dict)
    require(path.is_file() and not path.is_symlink(), f"missing trace-size aggregate: {path}")
    with path.open(encoding="utf-8") as stream:
        require(stream.readline().rstrip("\n") ==
                "trace_id\tsize_bytes\tcalls\trequested_bytes",
                f"unexpected trace-size aggregate header: {path}")
        for line_number, line in enumerate(stream, 2):
            parts = line.rstrip("\n").split("\t")
            require(len(parts) == 4, f"malformed trace-size row {path}:{line_number}")
            try:
                trace_id = int(parts[0], 16)
                size, calls, requested = (int(part) for part in parts[1:])
            except ValueError as error:
                raise AttributionError(f"non-integer trace-size row {path}:{line_number}") from error
            require(trace_id >= 0 and size >= 0 and calls > 0 and requested == size * calls,
                    f"trace-size arithmetic mismatch {path}:{line_number}")
            require(size not in rows[trace_id],
                    f"duplicate trace-size row {path}:{line_number}")
            rows[trace_id][size] = calls
    return rows


def read_heaptrack_histogram(path: Path) -> dict[int, int]:
    rows: dict[int, int] = {}
    require(path.is_file() and not path.is_symlink(), f"missing heaptrack histogram: {path}")
    with path.open(encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, 1):
            text = line.strip()
            if not text or text.startswith("#"):
                continue
            parts = text.split()
            require(len(parts) == 2, f"malformed heaptrack histogram row {path}:{line_number}")
            try:
                size, calls = (int(part) for part in parts)
            except ValueError as error:
                raise AttributionError(f"non-integer histogram row {path}:{line_number}") from error
            require(size >= 0 and calls >= 0 and size not in rows,
                    f"invalid or duplicate histogram row {path}:{line_number}")
            rows[size] = calls
    return rows


def read_trace_rows(path: Path) -> Iterable[dict[str, Any]]:
    require(path.is_file() and not path.is_symlink(), f"missing trace aggregate: {path}")
    with path.open(encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, 1):
            try:
                def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
                    value: dict[str, Any] = {}
                    for key, item in pairs:
                        require(key not in value,
                                f"duplicate trace key {key!r} at {path}:{line_number}")
                        value[key] = item
                    return value
                row = json.loads(line, object_pairs_hook=reject_duplicates)
            except json.JSONDecodeError as error:
                raise AttributionError(f"malformed trace row {path}:{line_number}") from error
            require(isinstance(row, dict), f"trace row is not an object {path}:{line_number}")
            yield row


def add_total(target: dict[str, int], calls: int, requested: int) -> None:
    target["trace_count"] += 1
    target["calls"] += calls
    target["requested_bytes"] += requested


def new_total() -> dict[str, int]:
    return {"trace_count": 0, "calls": 0, "requested_bytes": 0}


def category_totals() -> dict[str, dict[str, int]]:
    return {category: new_total() for category in EXPECTED_CATEGORIES}


def add_histogram(target: dict[int, int], source: dict[int, int]) -> None:
    for size, calls in source.items():
        target[size] = target.get(size, 0) + calls


def folded_stack(stack: list[dict[str, Any]], mode: str) -> str:
    names: list[str] = []
    for item in reversed(stack):
        frame = item.get("frame") if isinstance(item, dict) else None
        locations = frame_locations(frame, mode)
        if mode == "physical":
            values = [function_name(loc) for loc in locations[:1]]
        else:
            values = [function_name(loc) for loc in locations]
        names.extend(value or "<unknown>" for value in values)
    if not names:
        return "<empty>"
    return ";".join(value.replace(";", "%3B") for value in names)


def write_histogram(path: Path, histogram: dict[int, int]) -> None:
    lines = ["size_bytes\tcalls\n"]
    lines.extend(f"{size}\t{calls}\n" for size, calls in sorted(histogram.items()))
    write_text_once(path, "".join(lines))


def write_folded(path: Path, calls: dict[str, int], bytes_: dict[str, int]) -> None:
    lines = ["stack\tcalls\trequested_bytes\n"]
    keys = sorted(set(calls) | set(bytes_))
    lines.extend(f"{stack}\t{calls.get(stack, 0)}\t{bytes_.get(stack, 0)}\n"
                 for stack in keys)
    write_text_once(path, "".join(lines))


def compare_histograms(decoder: dict[int, tuple[int, int]], external: dict[int, int]) -> dict[str, Any]:
    decoder_calls = {size: calls for size, (calls, _requested) in decoder.items()}
    mismatches = []
    for size in sorted(set(decoder_calls) | set(external)):
        if decoder_calls.get(size) != external.get(size):
            mismatches.append({"size_bytes": size,
                               "decoder_calls": decoder_calls.get(size),
                               "heaptrack_print_calls": external.get(size)})
    return {"status": "pass" if not mismatches else "fail",
            "mismatch_count": len(mismatches), "mismatches": mismatches,
            "decoder_calls": sum(decoder_calls.values()),
            "heaptrack_print_calls": sum(external.values())}


def process_profile(label: str, decoder_root: Path, output_root: Path,
                    config: dict[str, Any], histogram_path: Path | None) -> dict[str, Any]:
    child = decoder_root / label
    summary_path = child / "summary.json"
    summary = read_json(summary_path)
    require(isinstance(summary, dict) and summary.get("schema") == DECODER_SCHEMA
            and summary.get("status") == "completed"
            and summary.get("label") == label,
            f"decoder summary is not completed for {label}")
    trace_path = child / "trace-aggregates.jsonl"
    size_path = child / "size-aggregates.tsv"
    trace_size_path = child / "trace-size-aggregates.tsv"
    decoder_sizes = read_size_aggregates(size_path)
    trace_sizes = read_trace_size_aggregates(trace_size_path)
    require(sum(calls for calls, _requested in decoder_sizes.values()) ==
            int(summary["allocation_events"]), f"decoder size calls mismatch for {label}")
    require(sum(requested for _calls, requested in decoder_sizes.values()) ==
            int(summary["requested_bytes_total"]), f"decoder size bytes mismatch for {label}")
    trace_size_calls = sum(sum(sizes.values()) for sizes in trace_sizes.values())
    trace_size_bytes = sum(sum(size * count for size, count in sizes.items())
                           for sizes in trace_sizes.values())
    require(trace_size_calls == int(summary["allocation_events"])
            and trace_size_bytes == int(summary["requested_bytes_total"]),
            f"decoder trace-size totals mismatch for {label}")
    trace_histogram: dict[int, int] = defaultdict(int)
    for sizes in trace_sizes.values():
        add_histogram(trace_histogram, sizes)
    require(dict(trace_histogram) ==
            {size: calls for size, (calls, _requested) in decoder_sizes.items()},
            f"decoder size histogram does not conserve trace-size rows for {label}")
    external_check = None
    if histogram_path is not None:
        external_check = compare_histograms(decoder_sizes, read_heaptrack_histogram(histogram_path))
        require(external_check["status"] == "pass",
                f"heaptrack histogram mismatch for {label}")

    out = output_root / label
    require(not out.exists() and not out.is_symlink(), f"refusing existing attribution output: {out}")
    out.mkdir(parents=True, exist_ok=False)
    window_trace_path = out / "window-traces.jsonl"
    physical_calls: dict[str, int] = defaultdict(int)
    physical_bytes: dict[str, int] = defaultdict(int)
    inline_calls: dict[str, int] = defaultdict(int)
    inline_bytes: dict[str, int] = defaultdict(int)
    physical_totals = category_totals()
    inline_totals = category_totals()
    publication_physical = category_totals()
    publication_inline = category_totals()
    cross: dict[str, dict[str, int]] = defaultdict(lambda: {"trace_count": 0, "calls": 0,
                                                               "requested_bytes": 0})
    physical_window_sizes: dict[int, int] = defaultdict(int)
    inline_window_sizes: dict[int, int] = defaultdict(int)
    publication_total = new_total()
    whole_total = new_total()
    publish_symbols: dict[str, dict[str, int]] = defaultdict(lambda: {"trace_count": 0,
                                                                         "calls": 0,
                                                                         "requested_bytes": 0})
    trace_count = 0
    with window_trace_path.open("xb") as window_stream:
        for row in read_trace_rows(trace_path):
            trace_count += 1
            trace_id_text = row.get("trace_id")
            require(isinstance(trace_id_text, str), f"trace id missing in {label}")
            try:
                trace_id = int(trace_id_text, 16)
                calls = int(row["calls"])
                requested = int(row["requested_bytes"])
            except (KeyError, TypeError, ValueError) as error:
                raise AttributionError(f"malformed trace totals in {label}: {trace_id_text}") from error
            require(calls > 0 and requested >= 0, f"invalid trace totals in {label}: {trace_id_text}")
            sizes = trace_sizes.get(trace_id)
            require(sizes is not None,
                    f"trace-size aggregate missing trace {trace_id_text} in {label}")
            require(sum(sizes.values()) == calls and
                    sum(size * count for size, count in sizes.items()) == requested,
                    f"trace-size conservation failed for {label}:{trace_id_text}")
            stack = row.get("stack")
            require(isinstance(stack, list), f"trace stack missing in {label}:{trace_id_text}")
            physical = classify_stack(stack, config, "physical")
            inline = classify_stack(stack, config, "inline")
            matches = publish_matches(stack, config["publish_audit_symbols"])
            add_total(whole_total, calls, requested)
            physical_totals[physical["category"]]["calls"] += calls
            physical_totals[physical["category"]]["requested_bytes"] += requested
            physical_totals[physical["category"]]["trace_count"] += 1
            inline_totals[inline["category"]]["calls"] += calls
            inline_totals[inline["category"]]["requested_bytes"] += requested
            inline_totals[inline["category"]]["trace_count"] += 1
            cross_key = f"{physical['category']}->{inline['category']}"
            cross[cross_key]["trace_count"] += 1
            cross[cross_key]["calls"] += calls
            cross[cross_key]["requested_bytes"] += requested
            if matches:
                add_total(publication_total, calls, requested)
                for match in matches:
                    symbol_total = publish_symbols[match["symbol"]]
                    symbol_total["trace_count"] += 1
                    symbol_total["calls"] += calls
                    symbol_total["requested_bytes"] += requested
                for bucket, target in ((physical, publication_physical),
                                       (inline, publication_inline)):
                    total = target[bucket["category"]]
                    total["trace_count"] += 1
                    total["calls"] += calls
                    total["requested_bytes"] += requested
            if physical["category"] == "window":
                add_histogram(physical_window_sizes, sizes)
                stack_key = folded_stack(stack, "physical")
                physical_calls[stack_key] += calls
                physical_bytes[stack_key] += requested
            if inline["category"] == "window":
                add_histogram(inline_window_sizes, sizes)
                stack_key = folded_stack(stack, "inline")
                inline_calls[stack_key] += calls
                inline_bytes[stack_key] += requested
            if physical["category"] == "window" or inline["category"] == "window":
                window_stream.write((json.dumps({
                    "trace_id": trace_id_text,
                    "calls": calls,
                    "requested_bytes": requested,
                    "physical_resolution": physical,
                    "inline_resolution": inline,
                    "publish_audit_matches": matches,
                    "stack": stack,
                }, sort_keys=True) + "\n").encode())
        window_stream.flush()
        os.fsync(window_stream.fileno())

    expected_calls = int(summary["allocation_events"])
    expected_bytes = int(summary["requested_bytes_total"])
    require(whole_total["calls"] == expected_calls and
            whole_total["requested_bytes"] == expected_bytes,
            f"trace total conservation failed for {label}")
    for partition_name, totals in (("physical", physical_totals), ("inline", inline_totals)):
        calls = sum(row["calls"] for row in totals.values())
        requested = sum(row["requested_bytes"] for row in totals.values())
        count = sum(row["trace_count"] for row in totals.values())
        require(calls == expected_calls and requested == expected_bytes and count == trace_count,
                f"{partition_name} category conservation failed for {label}")
    require(sum(row["calls"] for row in cross.values()) == expected_calls and
            sum(row["requested_bytes"] for row in cross.values()) == expected_bytes,
            f"physical/inline cross-resolution conservation failed for {label}")

    write_histogram(out / "window-physical-size.histogram", physical_window_sizes)
    write_histogram(out / "window-inline-size.histogram", inline_window_sizes)
    write_folded(out / "window-physical-stacks.folded", physical_calls, physical_bytes)
    write_folded(out / "window-inline-stacks.folded", inline_calls, inline_bytes)
    write_json_once(out / "publication-symbol-scoped.json", {
        "schema": f"{SCHEMA}.publication-scope.v1",
        "status": "pass",
        "scope": "only traces containing an identifiable publish_audit_symbols stack location",
        "interpretation": "symbol scoped; not temporal and not a claim that all publication allocations are captured",
        "totals": publication_total,
        "symbols": dict(sorted(publish_symbols.items())),
        "physical_categories": publication_physical,
        "inline_categories": publication_inline,
    })
    result = {
        "schema": SCHEMA,
        "status": "pass",
        "label": label,
        "decoder_summary": descriptor(summary_path),
        "config": descriptor(CONFIG_PATH),
        "whole_process": whole_total,
        "publication_symbol_scoped": publication_total,
        "category_partitions": {
            "physical": physical_totals,
            "inline": inline_totals,
        },
        "physical_inline_cross_resolution": dict(sorted(cross.items())),
        "conservation": {
            "decoder_allocation_events": expected_calls,
            "decoder_requested_bytes": expected_bytes,
            "trace_aggregate_count": trace_count,
            "size_aggregate_calls": sum(calls for calls, _requested in decoder_sizes.values()),
            "size_aggregate_requested_bytes": sum(requested for _calls, requested in decoder_sizes.values()),
            "trace_size_aggregate_calls": trace_size_calls,
            "trace_size_aggregate_requested_bytes": trace_size_bytes,
            "external_heaptrack_histogram": external_check,
            "all_exact": True,
        },
        "window_artifacts": {
            "raw_traces": descriptor(window_trace_path),
            "physical_size_histogram": descriptor(out / "window-physical-size.histogram"),
            "inline_size_histogram": descriptor(out / "window-inline-size.histogram"),
            "physical_stacks": descriptor(out / "window-physical-stacks.folded"),
            "inline_stacks": descriptor(out / "window-inline-stacks.folded"),
        },
        "scope_notes": {
            "whole_process": "all decoded allocation events and requested bytes before frees",
            "publication": "stacks containing configured source-replacement audit symbols only",
            "physical": "first location per recorded IP frame; nearest matching category",
            "inline": "all resolved inline locations; nearest matching category",
            "inline_only": "retained as inline evidence and excluded from physical ownership",
            "ownership": "diagnostic attribution; no category is claimed to own allocations solely from inline frames",
        },
    }
    write_json_once(out / "summary.json", result)
    return result


def parse_label_path(value: str, option: str) -> tuple[str, Path]:
    label, separator, raw_path = value.partition("=")
    require(separator and LABEL_RE.fullmatch(label) is not None and raw_path,
            f"{option} must be LABEL=PATH")
    path = Path(raw_path)
    require(path.is_file() and not path.is_symlink(), f"{option} input is not a file: {path}")
    return label, path


def main() -> int:
    global CONFIG_PATH
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--decoder-root", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--config", type=Path,
                        default=Path(__file__).with_name("attribution-config-v1.json"))
    parser.add_argument("--profile", action="append",
                        help="restrict to decoder label; repeat as needed")
    parser.add_argument("--histogram", action="append", default=[],
                        help="optional retained heaptrack histogram as LABEL=PATH")
    args = parser.parse_args()
    decoder_root = args.decoder_root.absolute()
    require(decoder_root.is_dir() and not decoder_root.is_symlink(),
            f"decoder root is not a real directory: {decoder_root}")
    CONFIG_PATH = args.config.absolute()
    config = load_config(CONFIG_PATH)
    terminal_path = decoder_root / "decoder-terminal.json"
    terminal = read_json(terminal_path)
    require(isinstance(terminal, dict)
            and terminal.get("schema") == DECODER_RUN_SCHEMA
            and terminal.get("status") == "terminal",
            "decoder terminal is missing or failed")
    profile_rows = terminal.get("profiles")
    require(isinstance(profile_rows, list), "decoder terminal profiles are malformed")
    available = {row.get("label"): row for row in profile_rows if isinstance(row, dict)}
    labels = args.profile or sorted(available)
    require(labels and len(labels) == len(set(labels)), "attribution labels are empty or duplicated")
    require(all(label in available for label in labels), "requested label is absent from decoder terminal")
    histograms: dict[str, Path] = {}
    for value in args.histogram:
        label, path = parse_label_path(value, "--histogram")
        require(label not in histograms, f"duplicate histogram label {label}")
        histograms[label] = path
    require(set(histograms) <= set(labels), "histogram label is not selected for attribution")
    output_root = args.output_root.absolute()
    require(not output_root.exists() and not output_root.is_symlink(),
            f"refusing existing attribution output root: {output_root}")
    output_root.mkdir(parents=True, exist_ok=False)
    started = now()
    results = []
    try:
        for label in labels:
            results.append(process_profile(label, decoder_root, output_root, config,
                                           histograms.get(label)))
    except Exception as error:
        write_json_once(output_root / "attribution-terminal.json", {
            "schema": RUN_SCHEMA,
            "status": "failed",
            "started_utc": started,
            "ended_utc": now(),
            "profiles_completed": [row.get("label") for row in results],
            "profiles_requested": labels,
            "failure": str(error),
            "scope": "offline attribution only; no workload or profiler launch",
        })
        raise
    write_json_once(output_root / "attribution-terminal.json", {
        "schema": RUN_SCHEMA,
        "status": "terminal",
        "started_utc": started,
        "ended_utc": now(),
        "profiles": results,
        "profiles_requested": labels,
        "scope": "offline attribution only; no workload or profiler launch",
    })
    print(json.dumps({"output": str(output_root), "profiles": len(results)}, sort_keys=True))
    return 0


CONFIG_PATH = Path(__file__).with_name("attribution-config-v1.json")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (AttributionError, OSError, ValueError, KeyError) as error:
        print(f"0949 offline attribution failed: {error}", file=os.sys.stderr)
        raise SystemExit(1)
