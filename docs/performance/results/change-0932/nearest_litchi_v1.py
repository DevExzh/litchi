#!/usr/bin/env python3
"""Offline nearest-litchi-frame attribution from the frozen 0932 trace aggregates."""
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
INDEX = ROOT / "capture-index-v1.json"
PREVIOUS = ROOT / "heaptrack-attribution-v2.json"
OUT = ROOT / "heaptrack-nearest-litchi-v1.json"

EXPECTED = {
    "s8192-c08-exact-shared_snapshot_clones-heaptrack-v1": (125004757, 9178781390),
    "s8192-c08-exact-independent_sources-heaptrack-v1": (178052276, 13077355214),
}
HASH_SUFFIX = re.compile(r"::h[0-9a-fA-F]{16}$")


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def desc(path: Path) -> dict:
    if not path.is_file():
        raise RuntimeError(f"missing input: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def load(path: Path) -> dict:
    with path.open(encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise RuntimeError(f"expected object: {path}")
    return value


def frame_name(item: dict) -> str | None:
    frame = (item or {}).get("frame") or {}
    locations = frame.get("locations") or []
    if not locations:
        return None
    name = locations[0].get("function")
    return name if isinstance(name, str) and name else None


def classify(row: dict) -> tuple[str, str]:
    names = [frame_name(item) for item in row.get("stack") or []]
    for name in names:
        if name and name.startswith("litchi_"):
            return HASH_SUFFIX.sub("", name), "nearest_litchi_frame"
    if any(name is None for name in names):
        return "<unknown>", "no_resolved_litchi_frame_with_unknown"
    return "<other>", "no_litchi_frame"


def main() -> int:
    if OUT.exists():
        raise RuntimeError(f"refusing existing output: {OUT}")
    index = load(INDEX)
    previous = load(PREVIOUS)
    groups_by_child = {}
    for child_id, (expected_calls, expected_bytes) in EXPECTED.items():
        trace_path = DECODER / child_id / "trace-aggregates.jsonl"
        summary = load(DECODER / child_id / "summary.json")
        groups = defaultdict(lambda: {
            "calls": 0,
            "requested_bytes": 0,
            "trace_count": 0,
            "classification": None,
        })
        trace_calls = 0
        trace_bytes = 0
        trace_count = 0
        for line_number, line in enumerate(trace_path.open(encoding="utf-8"), 1):
            row = json.loads(line)
            calls = int(row["calls"])
            requested = int(row["requested_bytes"])
            category, classification = classify(row)
            group = groups[category]
            if group["classification"] is None:
                group["classification"] = classification
            elif group["classification"] != classification:
                raise RuntimeError(f"category classification changed: {child_id}:{category}")
            group["calls"] += calls
            group["requested_bytes"] += requested
            group["trace_count"] += 1
            trace_calls += calls
            trace_bytes += requested
            trace_count += 1
        if (trace_calls, trace_bytes) != (expected_calls, expected_bytes):
            raise RuntimeError(f"trace total mismatch: {child_id}")
        if (int(summary["allocation_events"]), int(summary["requested_bytes_total"])) != (trace_calls, trace_bytes):
            raise RuntimeError(f"decoder total mismatch: {child_id}")
        rows = [
            {
                "category": category,
                "classification": group["classification"],
                "calls": group["calls"],
                "requested_bytes": group["requested_bytes"],
                "trace_count": group["trace_count"],
            }
            for category, group in groups.items()
        ]
        rows.sort(key=lambda row: (-row["requested_bytes"], -row["calls"], row["category"]))
        if sum(row["calls"] for row in rows) != trace_calls:
            raise RuntimeError(f"category call conservation mismatch: {child_id}")
        if sum(row["requested_bytes"] for row in rows) != trace_bytes:
            raise RuntimeError(f"category byte conservation mismatch: {child_id}")
        groups_by_child[child_id] = {
            "inputs": {
                "trace_aggregates": desc(trace_path),
                "decoder_summary": desc(DECODER / child_id / "summary.json"),
            },
            "whole_process": {
                "calls": trace_calls,
                "requested_bytes": trace_bytes,
                "trace_count": trace_count,
            },
            "categories": rows,
            "special_buckets": {
                "other": next((row for row in rows if row["category"] == "<other>"), None),
                "unknown": next((row for row in rows if row["category"] == "<unknown>"), None),
            },
            "exact_conservation": True,
        }
    source = index["source"]
    binary = index["binary"]
    output = {
        "schema": "litchi.performance.0932.heaptrack-nearest-litchi.v1",
        "status": "completed",
        "generated_utc": datetime.now(timezone.utc).isoformat(timespec="microseconds"),
        "inputs": {
            "capture_index": desc(INDEX),
            "immediate_caller_receipt": desc(PREVIOUS),
        },
        "binding": {
            "source_revision": source["base_revision"],
            "source_root": source["source_root"],
            "source_receipt_sha256": source["source_receipt_sha256"],
            "overlay_receipt_sha256": source["overlay_receipt_sha256"],
            "full_source_listing_sha256": source["full_source_listing_sha256"],
            "binary_sha256": binary["sha256"],
            "binary_bytes": binary["bytes"],
        },
        "classifier": {
            "definition": "For each trace aggregate, scan stack frames from leaf toward root. Select the first resolved function whose name starts with litchi_; strip only a terminal Rust ::h plus 16 hex digits. A trace with no such frame is <other> when all frame names resolve, or <unknown> when any frame name is unresolved. Each trace is assigned exactly one category.",
            "scope": "whole-process heaptrack requested allocation bytes before frees; this is a source-symbol attribution, not a phase attribution or speedup/adoption claim",
            "non_litchi_frames": "allocator, RawVec, serde, libc, and other frames are skipped until the nearest resolved litchi_ frame; they are not independently summed",
        },
        "children": groups_by_child,
        "all_children_exact_conservation": all(item["exact_conservation"] for item in groups_by_child.values()),
    }
    data = (json.dumps(output, indent=2, sort_keys=True) + "\n").encode()
    if len(data) >= 1048576:
        raise RuntimeError(f"output too large: {len(data)}")
    with OUT.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())
    print(json.dumps({"output": str(OUT), "bytes": len(data), "sha256": sha(OUT)}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

