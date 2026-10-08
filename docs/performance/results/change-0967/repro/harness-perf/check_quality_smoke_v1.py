#!/usr/bin/env python3
"""Fail-closed checker for the external 0967 performance-harness smoke set."""

from __future__ import annotations

import json
import pathlib
import sys


SCHEMA = "litchi.performance.0967.control_profile.v1"
EDIT_CELLS = {
    ("generated-positive", 1),
    ("dense1024", 32),
    ("false-marker", 1),
    ("false-marker", 32),
}
VARIANTS = {"baseline", "candidate"}
LANES = {"native", "allocator"}


def fail(message: str) -> None:
    raise SystemExit(f"quality smoke check failed: {message}")


def check_report(path: pathlib.Path) -> tuple[str, str, str, int | None]:
    try:
        report = json.loads(path.read_text())
    except Exception as error:  # noqa: BLE001 - checker must name malformed input.
        fail(f"{path}: invalid JSON: {error}")
    if not isinstance(report, dict):
        fail(f"{path}: report is not an object")
    if report.get("schema") != SCHEMA or report.get("schema_version") != 1:
        fail(f"{path}: wrong schema")
    variant = report.get("variant")
    lane = report.get("lane")
    control = report.get("control")
    if variant not in VARIANTS or lane not in LANES:
        fail(f"{path}: invalid variant/lane")
    if report.get("timing_claim") is not (lane == "native"):
        fail(f"{path}: timing_claim does not match measured lane")
    if report.get("allocator_timing_claim") is not False:
        fail(f"{path}: allocator timing claim must be false")
    samples = report.get("samples")
    if report.get("warmups") != 0 or report.get("samples_requested") != 1:
        fail(f"{path}: smoke report must be warmup=0/samples=1")
    if not isinstance(samples, list) or len(samples) != 1:
        fail(f"{path}: expected one sample")
    sample = samples[0]
    if sample.get("index") != 0 or sample.get("exact_output_equal_reference") is not True:
        fail(f"{path}: output equality gate failed")
    if report.get("allocation") != sample.get("allocation"):
        fail(f"{path}: report/sample allocation mismatch")
    if lane == "native":
        if report.get("allocation") is not None or sample.get("elapsed_ns") is None:
            fail(f"{path}: native lane shape failed")
    else:
        allocation = report.get("allocation")
        if sample.get("elapsed_ns") is not None or not isinstance(allocation, dict):
            fail(f"{path}: allocator lane shape failed")
        if allocation.get("allocation_calls", 0) <= 0:
            fail(f"{path}: allocator recorded no calls")
    if control == "real-positive":
        if report.get("scope") != "query" or report.get("edits") is not None:
            fail(f"{path}: query shape failed")
        if sample.get("source_unchanged") is not True:
            fail(f"{path}: query changed source")
        if sample.get("public_reopen_succeeded") is not True:
            fail(f"{path}: query reopen failed")
        if sample.get("revision_count") != 20:
            fail(f"{path}: query revision count failed")
        return variant, lane, control, None
    edits = report.get("edits")
    if (control, edits) not in EDIT_CELLS:
        fail(f"{path}: cell outside approved smoke matrix")
    if report.get("scope") != "edit" or report.get("retention") != "default_8mib":
        fail(f"{path}: edit shape failed")
    if sample.get("semantic_change_count") != edits:
        fail(f"{path}: semantic edit count failed")
    if sample.get("public_reopen_succeeded") is not True:
        fail(f"{path}: edit reopen failed")
    if sample.get("source_unchanged") is not False:
        fail(f"{path}: edit source flag failed")
    if sample.get("revision_count") != report.get("reference_revision_count"):
        fail(f"{path}: edit revision count failed")
    return variant, lane, control, edits


def main() -> None:
    if len(sys.argv) != 2:
        fail("usage: check_quality_smoke_v1.py REPORT_DIR")
    root = pathlib.Path(sys.argv[1])
    paths = sorted(root.glob("*.json"))
    if len(paths) != 20:
        fail(f"expected 20 reports, found {len(paths)}")
    seen = {check_report(path) for path in paths}
    expected = {
        (variant, lane, control, edits)
        for variant in VARIANTS
        for lane in LANES
        for control, edits in EDIT_CELLS
    }
    expected |= {
        (variant, lane, "real-positive", None)
        for variant in VARIANTS
        for lane in LANES
    }
    if seen != expected:
        fail(f"report set mismatch: {sorted(seen ^ expected)}")
    print(f"quality smoke check passed: {len(paths)} reports")


if __name__ == "__main__":
    main()
