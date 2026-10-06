#!/usr/bin/env python3
"""Offline, fail-closed conservation and disjoint stack attribution for 0932."""
from __future__ import annotations

import hashlib
import json
import os
import re
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path("/home/zhuhe/litchi-performance-0932")
DECODER = ROOT / "heaptrack-raw-decoder-v2"
ANALYSIS = ROOT / "heaptrack-analysis-v1"
INDEX = ROOT / "capture-index-v1.json"
OUT = ROOT / "heaptrack-attribution-v2.json"
SCRIPT = ROOT / "summarize_heaptrack_v2.py"

EXPECTED = {
    "s8192-c08-exact-shared_snapshot_clones-heaptrack-v1": {
        "decoder_events": 125004757,
        "decoder_bytes": 9178781390,
        "profile_sha256": "5025b7cac8d5f7d9eb2baf1572a4de82d99878747a11ef25ab449020d66deb8d",
    },
    "s8192-c08-exact-independent_sources-heaptrack-v1": {
        "decoder_events": 178052276,
        "decoder_bytes": 13077355214,
        "profile_sha256": "03c7cb95d1f8c86bb25a8155a65d93996ee94eebada27f2274a4acf0f4636686",
    },
}


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def descriptor(path: Path) -> dict:
    if not path.is_file():
        raise RuntimeError(f"missing input: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def fail(message: str) -> None:
    raise RuntimeError(message)


def read_size_tsv(path: Path) -> dict[int, tuple[int, int]]:
    rows: dict[int, tuple[int, int]] = {}
    with path.open(encoding="utf-8") as stream:
        header = stream.readline().rstrip("\n")
        if header != "size_bytes\tcalls\trequested_bytes":
            fail(f"unexpected decoder size header: {path}: {header!r}")
        for line_number, line in enumerate(stream, 2):
            parts = line.rstrip("\n").split("\t")
            if len(parts) != 3:
                fail(f"malformed decoder size row {path}:{line_number}")
            size, calls, requested = (int(part) for part in parts)
            if size in rows:
                fail(f"duplicate decoder size row: {path}:{size}")
            if requested != size * calls:
                fail(f"decoder size arithmetic mismatch: {path}:{line_number}")
            rows[size] = (calls, requested)
    return rows


def read_histogram(path: Path) -> dict[int, int]:
    rows: dict[int, int] = {}
    with path.open(encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, 1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split()
            if len(parts) != 2:
                fail(f"malformed heaptrack histogram row {path}:{line_number}")
            size, calls = (int(part) for part in parts)
            if size in rows:
                fail(f"duplicate heaptrack histogram row: {path}:{size}")
            rows[size] = calls
    return rows


def footer_calls(path: Path) -> int:
    pattern = re.compile(r"^calls to allocation functions: ([0-9]+) \(")
    found = []
    with path.open(encoding="utf-8", errors="replace") as stream:
        for line in stream:
            match = pattern.match(line.strip())
            if match:
                found.append(int(match.group(1)))
    if len(found) != 1:
        fail(f"expected one calls footer in {path}, found {len(found)}")
    return found[0]


def function_name(frame: dict | None) -> str | None:
    if not frame:
        return None
    locations = frame.get("locations") or []
    if not locations:
        return None
    value = locations[0].get("function")
    return value if isinstance(value, str) and value else None


def allocator_frame(name: str | None) -> bool:
    if not name:
        return False
    lowered = name.lower()
    return (
        "rust_alloc" in lowered
        or "__rust_realloc" in lowered
        or "malloc" in lowered
        or "calloc" in lowered
        or "realloc" in lowered
        or "operator new" in lowered
    )


def direct_caller(row: dict) -> tuple[str, str]:
    stack = row.get("stack") or []
    if not stack:
        return "<unknown>", "no_stack"
    first = function_name((stack[0] or {}).get("frame"))
    if allocator_frame(first):
        if len(stack) < 2:
            return "<unknown>", "allocator_only"
        second = function_name((stack[1] or {}).get("frame"))
        return second or "<unknown>", "after_allocator"
    return first or "<unknown>", "allocator_frame_not_found"


def trace_stack_functions(row: dict) -> list[str]:
    values = []
    for item in row.get("stack") or []:
        value = function_name((item or {}).get("frame"))
        values.append(value or "<unknown>")
    return values


def load_json(path: Path) -> dict:
    with path.open(encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        fail(f"expected object: {path}")
    return value


def summarize_child(child_id: str, index_child: dict, binary: dict, source: dict) -> dict:
    decoder_child = DECODER / child_id
    analysis_child = ANALYSIS / child_id
    raw_summary_path = decoder_child / "summary.json"
    trace_path = decoder_child / "trace-aggregates.jsonl"
    size_path = decoder_child / "size-aggregates.tsv"
    histogram_path = analysis_child / "allocation-size.histogram"
    full_summary_path = analysis_child / "full-summary.txt"
    summary = load_json(raw_summary_path)
    if summary.get("status") != "completed":
        fail(f"decoder is not completed: {child_id}")
    expected = EXPECTED[child_id]
    if summary["profile"]["sha256"] != expected["profile_sha256"]:
        fail(f"profile hash mismatch for {child_id}")
    if int(summary["allocation_events"]) != expected["decoder_events"]:
        fail(f"decoder call count mismatch: {child_id}")
    if int(summary["requested_bytes_total"]) != expected["decoder_bytes"]:
        fail(f"decoder byte count mismatch: {child_id}")
    decoder_sizes = read_size_tsv(size_path)
    heaptrack_sizes = read_histogram(histogram_path)
    decoder_histogram = {size: calls for size, (calls, _bytes) in decoder_sizes.items()}
    size_mismatches = []
    for size in sorted(set(decoder_histogram) | set(heaptrack_sizes)):
        if decoder_histogram.get(size) != heaptrack_sizes.get(size):
            size_mismatches.append({
                "size_bytes": size,
                "decoder_calls": decoder_histogram.get(size),
                "heaptrack_print_calls": heaptrack_sizes.get(size),
            })
    trace_calls = 0
    trace_bytes = 0
    trace_count = 0
    groups: dict[str, dict] = defaultdict(lambda: {
        "calls": 0,
        "requested_bytes": 0,
        "trace_count": 0,
        "classification_counts": defaultdict(int),
    })
    top_traces = []
    with trace_path.open(encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, 1):
            row = json.loads(line)
            calls = int(row["calls"])
            requested = int(row["requested_bytes"])
            trace_calls += calls
            trace_bytes += requested
            trace_count += 1
            caller, classification = direct_caller(row)
            group = groups[caller]
            group["calls"] += calls
            group["requested_bytes"] += requested
            group["trace_count"] += 1
            group["classification_counts"][classification] += 1
            if len(top_traces) < 50:
                top_traces.append({
                    "trace_id": row["trace_id"],
                    "calls": calls,
                    "requested_bytes": requested,
                    "unknown_any_frame": bool(row.get("unknown_any_frame")),
                    "unknown_leaf_frame": bool(row.get("unknown_leaf_frame")),
                    "stack": trace_stack_functions(row),
                })
    if trace_calls != int(summary["allocation_events"]) or trace_bytes != int(summary["requested_bytes_total"]):
        fail(f"trace aggregate totals mismatch: {child_id}")
    group_rows = []
    for caller, group in groups.items():
        group_rows.append({
            "caller": caller,
            "calls": group["calls"],
            "requested_bytes": group["requested_bytes"],
            "trace_count": group["trace_count"],
            "classification_counts": dict(sorted(group["classification_counts"].items())),
        })
    group_rows.sort(key=lambda row: (-row["requested_bytes"], -row["calls"], row["caller"]))
    hist_calls = sum(heaptrack_sizes.values())
    hist_bytes = sum(size * calls for size, calls in heaptrack_sizes.items())
    decoder_size_calls = sum(calls for calls, _bytes in decoder_sizes.values())
    decoder_size_bytes = sum(bytes_ for _calls, bytes_ in decoder_sizes.values())
    calls_footer = footer_calls(full_summary_path)
    if calls_footer != int(summary["allocation_events"]):
        fail(f"heaptrack_print footer calls mismatch: {child_id}")
    if hist_calls != int(summary["allocation_events"]) or hist_bytes != int(summary["requested_bytes_total"]):
        fail(f"histogram totals mismatch: {child_id}")
    unknown_group = groups.get("<unknown>", {"calls": 0, "requested_bytes": 0, "trace_count": 0})
    index_expected = next(row for row in load_json(INDEX)["children"] if row["id"] == child_id)
    return {
        "id": child_id,
        "ownership_mode": index_expected["report_identity"]["ownership_mode"],
        "inputs": {
            "profile": descriptor(Path(summary["profile"]["path"])),
            "raw_summary": descriptor(raw_summary_path),
            "trace_aggregates": descriptor(trace_path),
            "decoder_size_aggregates": descriptor(size_path),
            "heaptrack_size_histogram": descriptor(histogram_path),
            "heaptrack_full_summary": descriptor(full_summary_path),
        },
        "conservation": {
            "decoder_allocation_events": int(summary["allocation_events"]),
            "decoder_requested_bytes": int(summary["requested_bytes_total"]),
            "decoder_size_calls": decoder_size_calls,
            "decoder_size_requested_bytes": decoder_size_bytes,
            "heaptrack_print_histogram_calls": hist_calls,
            "heaptrack_print_histogram_requested_bytes": hist_bytes,
            "heaptrack_print_footer_calls": calls_footer,
            "trace_aggregate_count": trace_count,
            "trace_aggregate_calls": trace_calls,
            "trace_aggregate_requested_bytes": trace_bytes,
            "direct_caller_group_count": len(group_rows),
            "direct_caller_calls": sum(row["calls"] for row in group_rows),
            "direct_caller_requested_bytes": sum(row["requested_bytes"] for row in group_rows),
            "size_key_count": len(decoder_sizes),
            "size_histogram_mismatch_count": len(size_mismatches),
            "size_histogram_mismatches": size_mismatches,
            "all_exact_totals_match": (
                decoder_size_calls == hist_calls == trace_calls == calls_footer == int(summary["allocation_events"])
                and decoder_size_bytes == hist_bytes == trace_bytes == int(summary["requested_bytes_total"])
                and not size_mismatches
            ),
            "decoder_status": summary.get("status"),
            "decoder_plus_without_metadata": summary.get("plus_without_metadata"),
            "decoder_a_plus_adjacency_mismatches": summary.get("a_plus_adjacency_mismatches"),
            "decoder_unresolved_trace_definitions": summary.get("unresolved_trace_definitions"),
            "decoder_allocation_minus_deallocation": summary.get("allocation_minus_deallocation"),
        },
        "unknown_or_unresolved": {
            "decoder_unknown_any_frame": summary["unknown_any_frame"],
            "decoder_unknown_leaf_frame": summary["unknown_leaf_frame"],
            "direct_caller_unknown_group": {
                "calls": unknown_group["calls"],
                "requested_bytes": unknown_group["requested_bytes"],
                "trace_count": unknown_group["trace_count"],
            },
        },
        "direct_caller_groups": group_rows,
        "top_trace_rows": top_traces,
        "top_size_rows_by_requested_bytes": [
            {"size_bytes": size, "calls": calls, "requested_bytes": size * calls}
            for size, calls in sorted(heaptrack_sizes.items(), key=lambda item: (-(item[0] * item[1]), -item[1], item[0]))[:50]
        ],
        "scope": {
            "attribution": "whole-process heaptrack allocation events; direct-caller groups are disjoint by trace_id",
            "phase_attribution": "unavailable: run_scenario is inlined into run_capture and heaptrack has no frozen phase markers",
            "requested_bytes": "allocation size from heaptrack a/+ records before frees; not allocator metric deltas or latency",
            "unknown": "unknown caller group is only unresolved/missing stack attribution; whole-process setup, correctness, and release activity remain included",
        },
        "binding": {
            "source": source,
            "binary": binary,
            "capture_index_child": index_expected["id"],
        },
    }


def main() -> int:
    if OUT.exists():
        fail(f"refusing existing output: {OUT}")
    index = load_json(INDEX)
    source = index["source"]
    binary = index["binary"]
    child_ids = sorted(EXPECTED)
    index_ids = {row["id"] for row in index["children"]}
    if set(child_ids) != index_ids:
        fail(f"capture index child mismatch: expected {child_ids}, found {sorted(index_ids)}")
    index_children = {row["id"]: row for row in index["children"]}
    profiles = [summarize_child(child_id, index_children[child_id], binary, source) for child_id in child_ids]
    output = {
        "schema": "litchi.performance.0932.heaptrack-attribution.v2",
        "status": "completed",
        "generated_utc": datetime.now(timezone.utc).isoformat(timespec="microseconds"),
        "capture_index": descriptor(INDEX),
        "capture_binding": {
            "source_revision": source["base_revision"],
            "source_root": source["source_root"],
            "source_receipt_sha256": source["source_receipt_sha256"],
            "overlay_receipt_sha256": source["overlay_receipt_sha256"],
            "full_source_listing_sha256": source["full_source_listing_sha256"],
            "binary_sha256": binary["sha256"],
            "binary_bytes": binary["bytes"],
            "profiler": index["profiler"],
        },
        "method": {
            "decoder": "decode_heaptrack_raw_v2.py: implicit string/IP/trace/allocation-info indices; + references allocation-info index; fail-closed missing references and string-length checks",
            "conservation": "exact size->count comparison against heaptrack_print allocation-size.histogram plus trace, footer-call, and requested-byte totals",
            "direct_caller": "for each disjoint trace aggregate, skip the first allocator/reallocator frame when identified; otherwise retain first frame under allocator_frame_not_found; aggregate each trace exactly once",
            "stack_evidence": "full trace-aggregates.jsonl remains externally bound per child; top 50 trace stacks are copied into this compact receipt",
            "claim_boundary": "descriptive whole-process allocation attribution only; no phase, latency, speedup, or adoption claim",
        },
        "profiles": profiles,
        "all_profiles_exact_conservation": all(
            profile["conservation"]["all_exact_totals_match"] for profile in profiles
        ),
    }
    data = (json.dumps(output, indent=2, sort_keys=True) + "\n").encode()
    if len(data) >= 1048576:
        fail(f"attribution output is too large: {len(data)}")
    with OUT.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())
    print(json.dumps({
        "output": str(OUT),
        "bytes": len(data),
        "sha256": sha(OUT),
        "profiles": len(profiles),
        "all_profiles_exact_conservation": output["all_profiles_exact_conservation"],
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

