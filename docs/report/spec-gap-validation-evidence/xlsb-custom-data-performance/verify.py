#!/usr/bin/env python3
"""Verify raw candidate-only receipts and preservation gates."""

from __future__ import annotations

import argparse
import json
import re
import statistics
from pathlib import Path


SOURCE_COMMIT = "16102fe751d7c5492042330f1bd1f49c304495f0"
LANES = {
    "read",
    "snapshotclone",
    "noop",
    "editpayload",
    "rename-many-to-one",
    "mixedinsertremove",
    "patchinverse",
    "publicapply",
}
LEVELS = {"small", "medium", "large"}


def fail(message: str) -> None:
    raise SystemExit(f"xlsb custom-data profile verification failed: {message}")


def load_rows(path: Path) -> list[dict]:
    rows: list[dict] = []
    for line_number, line in enumerate(path.read_text().splitlines(), 1):
        if not line.strip():
            continue
        try:
            value = json.loads(line)
        except json.JSONDecodeError as error:
            fail(f"{path}:{line_number}: invalid JSON: {error}")
        if not isinstance(value, dict):
            fail(f"{path}:{line_number}: receipt is not an object")
        rows.append(value)
    if not rows:
        fail(f"{path}: no raw samples")
    return rows


def verify_allocator(row: dict) -> None:
    allocator = row.get("allocator")
    if not isinstance(allocator, dict):
        fail(f"{row.get('lane')}/{row.get('size_class')}: allocator object missing")
    required = {
        "allocation_calls",
        "reallocation_calls",
        "deallocation_calls",
        "direct_allocated_bytes",
        "realloc_old_bytes",
        "realloc_new_bytes",
        "requested_alloc_bytes",
        "released_alloc_bytes_operation",
        "released_alloc_bytes_through_drop",
        "live_before",
        "live_after_operation",
        "live_after_drop",
        "logical_peak_bytes",
        "peak_live_delta",
        "allocation_failed",
        "alloc_balance_ok",
        "alloc_invalid",
    }
    missing = required.difference(allocator)
    if missing:
        fail(f"allocator fields missing: {sorted(missing)}")
    if allocator["requested_alloc_bytes"] != (
        allocator["direct_allocated_bytes"] + allocator["realloc_new_bytes"]
    ):
        fail("requested allocation equation failed")
    expected_live = (
        allocator["live_before"]
        + allocator["direct_allocated_bytes"]
        + allocator["realloc_new_bytes"]
        - allocator["realloc_old_bytes"]
        - allocator["released_alloc_bytes_operation"]
    )
    if allocator["live_after_operation"] != expected_live:
        fail("live allocation equation failed")
    if allocator["logical_peak_bytes"] < allocator["live_before"]:
        fail("logical peak is below operation live baseline")
    if not allocator["alloc_balance_ok"] or allocator["alloc_invalid"]:
        fail("allocator balance or underflow gate failed")
    if allocator["allocation_failed"] != 0:
        fail("allocator reported a failed allocation")


def verify_row(row: dict, path: Path) -> None:
    lane = row.get("lane")
    level = row.get("size_class")
    if row.get("schema") != "xlsb-custom-data-performance-v1":
        fail(f"{path}: schema mismatch")
    if row.get("source_commit") != SOURCE_COMMIT:
        fail(f"{path}: source commit mismatch")
    if lane not in LANES or level not in LEVELS:
        fail(f"{path}: unknown lane/class {lane}/{level}")
    if not isinstance(row.get("timing_ns"), int) or row["timing_ns"] <= 0:
        fail(f"{path}: invalid timing")
    fixture = row.get("fixture")
    if not isinstance(fixture, dict):
        fail(f"{path}: fixture metadata missing")
    for field in (
        "source_bytes",
        "source_sha256",
        "source_fnv1a64",
        "connections_bytes",
        "connections_sha256",
        "opaque_bytes",
        "opaque_sha256",
        "storages",
        "references",
        "payload_bytes_per_storage",
    ):
        if field not in fixture:
            fail(f"{path}: fixture field {field} missing")
    if not isinstance(row.get("copies_bytes_observed"), int):
        fail(f"{path}: copy accounting missing")
    candidate_hash = row.get("candidate_package_sha256")
    if not isinstance(candidate_hash, str) or re.fullmatch(r"[0-9a-f]{64}", candidate_hash) is None:
        fail(f"{path}: candidate package hash missing or malformed")
    if row.get("input_preparation") in (None, ""):
        fail(f"{path}: input preparation scope missing")
    forward_hash = row.get("prepared_forward_package_sha256")
    if lane == "patchinverse":
        if not isinstance(forward_hash, str) or re.fullmatch(r"[0-9a-f]{64}", forward_hash) is None:
            fail(f"{path}: prepared patch forward hash missing or malformed")
    elif forward_hash is not None:
        fail(f"{path}: unexpected prepared patch forward hash")
    if lane in {"read", "snapshotclone", "noop", "patchinverse"} and candidate_hash != fixture["source_sha256"]:
        fail(f"{path}: unchanged candidate package hash differs from source")
    assertions = row.get("assertions")
    if not isinstance(assertions, dict) or not assertions.get("semantics_correct"):
        fail(f"{path}: semantic assertion failed")
    if not assertions.get("payloads_exact"):
        fail(f"{path}: exact payload assertion failed")
    if lane in {"noop"} and not assertions.get("source_exact_noop"):
        fail(f"{path}: exact no-op assertion failed")
    if lane in {
        "editpayload",
        "rename-many-to-one",
        "mixedinsertremove",
        "patchinverse",
        "publicapply",
    }:
        if not assertions.get("preserved_unrelated_bytes"):
            fail(f"{path}: unrelated-member preservation assertion failed")
        if not assertions.get("inverse_correct"):
            fail(f"{path}: inverse assertion failed")
    verify_allocator(row)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    args = parser.parse_args()
    raw_dir = args.results / "raw"
    files = sorted(raw_dir.glob("*.jsonl"))
    if len(files) != 24:
        fail(f"expected 24 raw lane/class files, found {len(files)}")
    all_rows: list[dict] = []
    groups: set[tuple[str, str]] = set()
    for path in files:
        rows = load_rows(path)
        if len(rows) != 5:
            fail(f"{path}: expected five measured samples, found {len(rows)}")
        for row in rows:
            verify_row(row, path)
            groups.add((row["lane"], row["size_class"]))
            all_rows.append(row)
    if groups != {(lane, level) for lane in LANES for level in LEVELS}:
        fail("lane/class coverage is incomplete")
    if (args.results / "source-manifest.txt").stat().st_size == 0:
        fail("source manifest is empty")
    if not (args.results / "binary.sha256").is_file():
        fail("binary hash receipt is missing")
    if not (args.results / "toolchain.txt").is_file():
        fail("toolchain receipt is missing")
    if "build_profile=release" not in (args.results / "toolchain.txt").read_text().splitlines():
        fail("toolchain receipt does not identify an optimized release build")
    fixture_hashes = args.results / "fixture-sha256.txt"
    if not fixture_hashes.is_file() or not fixture_hashes.read_text().strip():
        fail("fixture hash receipt is missing")

    summary_lines = [
        "schema=xlsb-custom-data-performance-v1",
        f"samples={len(all_rows)}",
        "candidate_only=true",
        "comparison_baseline=none",
        "native_producer_claim=false",
        "speedup_claim=false",
    ]
    for lane in sorted(LANES):
        for level in sorted(LEVELS):
            values = [
                row["timing_ns"]
                for row in all_rows
                if row["lane"] == lane and row["size_class"] == level
            ]
            summary_lines.append(
                f"{lane}\t{level}\tmedian_ns={int(statistics.median(values))}"
                f"\tmin_ns={min(values)}\tmax_ns={max(values)}"
            )
    (args.results / "summary.tsv").write_text("\n".join(summary_lines) + "\n")
    print(f"verified {len(all_rows)} raw samples across 24 lane/class groups")


if __name__ == "__main__":
    main()
