#!/usr/bin/env python3
"""Create-only reconciliation for the sealed 0949 XML multi-window traces.

This is offline-only.  It reads decoded heaptrack aggregates and source text,
never launches a workload or decoder, and writes exactly one new JSON output.
The dual publisher/caller frame filter is a symbol/call-site scope, not a
heaptrack time interval.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from collections import Counter
from pathlib import Path
from typing import Any, Iterable

SCHEMA = "litchi.performance.0949.xml-multiwindow-source-profile-reconciliation.v2"
PUBLISHER = "SourceBackedEditor::publish_multi_commit_to_stream"
CALLER = "run_xlsx_cell_values_edit_save"
CALLER_LINE = 44840
ROLES = (
    ("baseline-medium", "medium", "baseline"),
    ("candidate-medium", "medium", "candidate"),
    ("baseline-dense", "dense", "baseline"),
    ("candidate-dense", "dense", "candidate"),
)
EXPECTED_DUAL = {
    "baseline-medium": {"calls": 406, "bytes": 2_495_613, "realloc_calls": 50},
    "candidate-medium": {"calls": 681, "bytes": 2_677_821, "realloc_calls": 54},
    "baseline-dense": {"calls": 406, "bytes": 2_765_840, "realloc_calls": 49},
    "candidate-dense": {"calls": 931, "bytes": 3_172_336, "realloc_calls": 50},
}
EXPECTED_CLASS = {
    "candidate-medium": {
        "parser_windows": (279, 4_464, 3),
        "window_vec": (16, 176_640, 2),
        "open_vec": (8, 1_536, 2),
    },
    "candidate-dense": {
        "parser_windows": (531, 8_496, 3),
        "window_vec": (11, 397_440, 2),
        "open_vec": (6, 1_152, 2),
    },
}
EXPECTED_PAIR = {
    "medium": {"calls": 275, "bytes": 182_208, "realloc_delta": 4},
    "dense": {"calls": 525, "bytes": 406_496, "realloc_delta": 1},
}
EXPECTED_RESIDUAL = {
    "medium": {"calls": -28, "bytes": -432},
    "dense": {"calls": -23, "bytes": -592},
}

SOURCE_FILES = (
    "crates/xml-minifier/src/audit.rs",
    "crates/xml-minifier/src/audit/namespaces.rs",
    "tools/perf-baseline/src/allocation_metrics.rs",
    "tools/perf-baseline/src/bin/support/counting_allocator.rs",
    "crates/litchi-xlsx/src/cell_values/source.rs",
    "tools/perf-baseline/src/lib.rs",
)
SOURCE_TARGETS = (
    ("MultiWindowSearch::consider Vec<Window>::try_reserve", "crates/xml-minifier/src/audit.rs", 2509),
    ("MultiWindowSearch::accepted open.try_reserve", "crates/xml-minifier/src/audit.rs", 2554),
    ("scan State::spaces.try_reserve", "crates/xml-minifier/src/audit.rs", 2106),
    ("prove_windows audit_window loop", "crates/xml-minifier/src/audit.rs", 2651),
    ("verify_source_replacement_with_spans prove_windows call", "crates/xml-minifier/src/audit.rs", 1173),
    ("scan reader.read_event", "crates/xml-minifier/src/audit.rs", 2061),
    ("Namespaces::try_clone", "crates/xml-minifier/src/audit/namespaces.rs", 140),
    ("Namespaces::try_clone_vec", "crates/xml-minifier/src/audit/namespaces.rs", 571),
    ("allocation_metrics reallocation counter", "tools/perf-baseline/src/allocation_metrics.rs", 699),
    ("publication allocator wrapper", "tools/perf-baseline/src/bin/support/counting_allocator.rs", 51),
    ("publisher function", "crates/litchi-xlsx/src/cell_values/source.rs", 649),
    ("publisher benchmark caller", "tools/perf-baseline/src/lib.rs", 44840),
)


def fail(message: str) -> None:
    raise RuntimeError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink():
        fail(f"not a regular file: {path}")
    return {"path": str(path.resolve()), "bytes": path.stat().st_size, "sha256": sha256(path)}


def line_text(path: Path, line_number: int) -> str:
    lines = path.read_text(encoding="utf-8").splitlines()
    if line_number < 1 or line_number > len(lines):
        fail(f"source line out of range: {path}:{line_number}")
    return lines[line_number - 1].strip()


def locations(row: dict[str, Any]) -> Iterable[dict[str, Any]]:
    for stack_row in row.get("stack", []):
        frame = stack_row.get("frame") or {}
        for location in frame.get("locations", []):
            if isinstance(location, dict):
                yield location


def has_function(row: dict[str, Any], needle: str) -> bool:
    return any(needle in (location.get("function") or "") for location in locations(row))


def has_source_line(row: dict[str, Any], path_fragment: str, line_number: int) -> bool:
    return any(
        path_fragment in (location.get("file") or "")
        and location.get("line") == line_number
        for location in locations(row)
    )


def allocator_kind(row: dict[str, Any]) -> str:
    first = (row.get("stack") or [{}])[0].get("frame") or {}
    first_locations = first.get("locations", [])
    if any(
        location.get("line") == 54
        or "::realloc" in (location.get("function") or "")
        for location in first_locations
    ):
        return "realloc"
    for location in first_locations:
        if (location.get("file") or "").endswith("/std/src/sys/alloc/unix.rs"):
            if location.get("line") in (14, 28, 36, 83) or "::alloc" in (location.get("function") or ""):
                return "alloc"
    fail(f"selected row has no resolved alloc/realloc leaf: {row.get('trace_id')}")


def read_rows(path: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    calls = 0
    requested_bytes = 0
    with path.open("rb") as stream:
        for line_number, raw in enumerate(stream, 1):
            try:
                row = json.loads(raw)
            except json.JSONDecodeError as error:
                fail(f"invalid JSON at {path}:{line_number}: {error}")
            if not isinstance(row, dict):
                fail(f"non-object row at {path}:{line_number}")
            if not isinstance(row.get("calls"), int) or not isinstance(row.get("requested_bytes"), int):
                fail(f"bad counters at {path}:{line_number}")
            calls += row["calls"]
            requested_bytes += row["requested_bytes"]
            rows.append(row)
    return rows, {"calls": calls, "bytes": requested_bytes}


def aggregate(rows: Iterable[dict[str, Any]]) -> dict[str, Any]:
    result = {"rows": 0, "calls": 0, "bytes": 0, "alloc_calls": 0, "alloc_bytes": 0,
              "realloc_calls": 0, "realloc_bytes": 0}
    for row in rows:
        kind = allocator_kind(row)
        calls = row["calls"]
        requested = row["requested_bytes"]
        result["rows"] += 1
        result["calls"] += calls
        result["bytes"] += requested
        result[f"{kind}_calls"] += calls
        result[f"{kind}_bytes"] += requested
    return result


def dual(row: dict[str, Any]) -> bool:
    return has_function(row, PUBLISHER) and has_function(row, CALLER) and has_source_line(
        row, "src/lib.rs", CALLER_LINE
    )


def publisher_only(row: dict[str, Any]) -> bool:
    return has_function(row, PUBLISHER)


def class_name(row: dict[str, Any]) -> str | None:
    multiwindow = has_source_line(row, "xml-minifier/src/audit.rs", 2651) and has_source_line(
        row, "xml-minifier/src/audit.rs", 1173
    )
    parser_lines = [
        ("parser_opened_starts", "quick-xml-0.41.0/src/reader/state.rs", 346),
        ("parser_opened_buffer", "quick-xml-0.41.0/src/reader/state.rs", 347),
        ("parser_spaces", "xml-minifier/src/audit.rs", 2106),
    ]
    parser_matches = [name for name, path, line in parser_lines if multiwindow and has_source_line(row, path, line)]
    if len(parser_matches) > 1:
        fail(f"parser class overlap at trace {row.get('trace_id')}: {parser_matches}")
    if parser_matches:
        return "parser_windows"
    window = has_source_line(row, "xml-minifier/src/audit.rs", 2509)
    open_stack = has_source_line(row, "xml-minifier/src/audit.rs", 2554)
    if window and open_stack:
        fail(f"window/open class overlap at trace {row.get('trace_id')}")
    if window:
        return "window_vec"
    if open_stack:
        return "open_vec"
    return None


def source_manifest(source_root: Path) -> dict[str, Any]:
    files: dict[str, Any] = {}
    for relative in SOURCE_FILES:
        path = source_root / relative
        files[relative] = descriptor(path)
    targets: list[dict[str, Any]] = []
    for name, relative, line in SOURCE_TARGETS:
        path = source_root / relative
        targets.append({"name": name, "file": relative, "line": line, "text": line_text(path, line)})
    return {"root": str(source_root.resolve()), "files": files, "targets": targets}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, default=Path("/home/zhuhe/litchi-performance-0949"),
                        help="0949 packet root containing raw-decoder-v1")
    parser.add_argument("--source-root", type=Path,
                        default=Path("/home/zhuhe/litchi-performance-0948/candidate-build-v3"),
                        help="candidate source snapshot used for target line hashes")
    parser.add_argument("--output", type=Path, default=None,
                        help="new JSON path; existing paths are refused")
    args = parser.parse_args()
    packet = args.packet.resolve()
    source_root = args.source_root.resolve()
    output = (args.output or (packet / "review/xml-multiwindow-source-profile-reconciliation-v2.json")).resolve()
    if output.exists() or output.is_symlink():
        fail(f"create-only output already exists: {output}")

    decoder = packet / "tools/decode_heaptrack_0949_v1.py"
    if not decoder.is_file():
        decoder = Path("/home/zhuhe/litchi-performance-0949/tools/decode_heaptrack_0949_v1.py")
    if not decoder.is_file():
        fail(f"decoder not found: {decoder}")

    role_data: dict[str, Any] = {}
    raw_summaries: dict[str, Any] = {}
    for label, density, side in ROLES:
        directory = packet / "raw-decoder-v1" / label
        aggregate_path = directory / "trace-aggregates.jsonl"
        summary_path = directory / "summary.json"
        rows, summed = read_rows(aggregate_path)
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
        if summary.get("schema") != "litchi.performance.0949.heaptrack-raw-decode.v1":
            fail(f"unexpected decoder schema for {label}")
        if summary.get("unresolved_trace_definitions") != 0 or summary.get("plus_without_metadata") != 0:
            fail(f"decoder resolution failure for {label}")
        if summary.get("allocation_events") != summed["calls"] or summary.get("requested_bytes_total") != summed["bytes"]:
            fail(f"aggregate/summary mismatch for {label}")
        raw_summaries[label] = {
            "summary": descriptor(summary_path),
            "trace_aggregates": descriptor(aggregate_path),
            "profile": summary.get("profile"),
            "aggregate": summed,
        }
        dual_rows = [row for row in rows if dual(row)]
        publisher_rows = [row for row in rows if publisher_only(row)]
        dual_summary = aggregate(dual_rows)
        publisher_summary = aggregate(publisher_rows)
        expected = EXPECTED_DUAL[label]
        if dual_summary["calls"] != expected["calls"] or dual_summary["bytes"] != expected["bytes"]:
            fail(f"dual filter mismatch for {label}: {dual_summary}")
        if dual_summary["realloc_calls"] != expected["realloc_calls"]:
            fail(f"dual realloc mismatch for {label}: {dual_summary}")
        role_data[label] = {
            "density": density,
            "side": side,
            "dual_filter": {
                "publisher_function": PUBLISHER,
                "caller_function": CALLER,
                "caller_line": CALLER_LINE,
                "summary": dual_summary,
            },
            "publisher_only": publisher_summary,
            "rows": len(rows),
        }
        if side == "candidate":
            class_rows: dict[str, list[dict[str, Any]]] = {
                "parser_windows": [], "window_vec": [], "open_vec": []
            }
            for row in dual_rows:
                name = class_name(row)
                if name is not None:
                    class_rows[name].append(row)
            class_summary = {name: aggregate(items) for name, items in class_rows.items()}
            expected_classes = EXPECTED_CLASS[label]
            for name, (calls, bytes_, rows_expected) in expected_classes.items():
                measured = class_summary[name]
                if (measured["calls"], measured["bytes"], measured["rows"]) != (calls, bytes_, rows_expected):
                    fail(f"class mismatch for {label}/{name}: {measured}")
            ids = {name: sorted(row["trace_id"] for row in items) for name, items in class_rows.items()}
            all_ids = [trace_id for values in ids.values() for trace_id in values]
            if len(all_ids) != len(set(all_ids)):
                fail(f"site class overlap for {label}: {ids}")
            class_total = aggregate(row for items in class_rows.values() for row in items)
            role_data[label]["disjoint_site_classes"] = {
                "classes": class_summary,
                "trace_ids": ids,
                "subtotal": class_total,
                "unclassified_dual_rows": dual_summary["rows"] - class_total["rows"],
            }

    pair_data: dict[str, Any] = {}
    residual_data: dict[str, Any] = {}
    for density in ("medium", "dense"):
        baseline = role_data[f"baseline-{density}"]["dual_filter"]["summary"]
        candidate = role_data[f"candidate-{density}"]["dual_filter"]["summary"]
        delta = {
            "calls": candidate["calls"] - baseline["calls"],
            "bytes": candidate["bytes"] - baseline["bytes"],
            "realloc_delta": candidate["realloc_calls"] - baseline["realloc_calls"],
        }
        if delta != EXPECTED_PAIR[density]:
            fail(f"pair delta mismatch for {density}: {delta}")
        classes = role_data[f"candidate-{density}"]["disjoint_site_classes"]["subtotal"]
        residual = {"calls": delta["calls"] - classes["calls"], "bytes": delta["bytes"] - classes["bytes"]}
        if residual != EXPECTED_RESIDUAL[density]:
            fail(f"residual mismatch for {density}: {residual}")
        pair_data[density] = {"baseline": baseline, "candidate": candidate, "delta": delta}
        residual_data[density] = {"candidate_class_subtotal": classes, "pair_delta": delta, "residual": residual}

    document = {
        "schema": SCHEMA,
        "verifier": {
            "script": descriptor(Path(__file__).resolve()),
        },
        "scope": {
            "decoder": descriptor(decoder),
            "raw_accounting": "decoded heaptrack + records over the whole profile; requested bytes before frees",
            "dual_filter": "resolved publisher function plus run_xlsx_cell_values_edit_save at src/lib.rs:44840",
            "temporal_claim": False,
            "reallocation_note": "allocation_calls includes successful reallocations; reallocation_calls is a subset and must not be added",
            "create_only": True,
        },
        "source": source_manifest(source_root),
        "raw_profiles": raw_summaries,
        "roles": role_data,
        "pair_deltas": pair_data,
        "residuals_after_disjoint_candidate_classes": residual_data,
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as stream:
        json.dump(document, stream, indent=2, sort_keys=True)
        stream.write("\n")
    print(output)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except RuntimeError as error:
        raise SystemExit(f"error: {error}")
