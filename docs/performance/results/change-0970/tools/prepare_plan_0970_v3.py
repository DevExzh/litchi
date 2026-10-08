#!/usr/bin/env python3
"""Create the immutable, pre-capture 0970 historical-baseline plan."""
from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

from protocol_0970_v3 import (
    BOOTSTRAP_RESAMPLES,
    BOOTSTRAP_SEED,
    CASES,
    CORRECTION_RECEIPT,
    DESIGN_SHA256,
    EDIT_COUNTS,
    EVENTS,
    EXPECTED_ALLOWED_AFFINITY,
    EXPECTED_PIN_CPU,
    HELPER_SEAL_SCHEMA,
    PAIRS,
    PLAN_SCHEMA,
    RETENTION,
    ROOT,
    SAMPLES,
    SOURCE_FILE_MANIFEST,
    SOURCE_FILE_MANIFEST_SHA256,
    SOURCE_MANIFEST,
    SOURCE_MANIFEST_SHA256,
    SOURCE_ROOT,
    TIMEOUT_SECONDS,
    VALIDATOR_PATH,
    VALIDATOR_SHA256,
    WARMUPS,
    REQUIRED_LOCALE,
    binding,
    cells,
    check_binding,
    strict_load_json,
    validate_source_custody,
    write_json_exclusive,
)

BASELINE_BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
BASELINE_SHA256 = "0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb"
FIXTURES = {
    "docnohf": Path("/home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/NoHeadFoot.doc"),
    "docfloat": Path("/home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/FloatingPictures.doc"),
}
BASELINE_REUSE_REVIEW = Path("/home/zhuhe/litchi-performance-0969/review/baseline-reuse-v1.json")
SOURCE_VARIABILITY_REVIEW = ROOT / "source-variability-review-v2.md"
DESIGN = ROOT / "diagnostic-design-v3.md"
HOST_AUDIT = ROOT / "host-audit-v1.md"
TOOL_ADDENDUM = ROOT / "host-tool-addendum-v1.json"
TELEMETRY_PROBE = ROOT / "telemetry-probe-v1.json"
AFFINITY_PROBE = ROOT / "affinity-probe-v1.json"
HELPER_SEAL = ROOT / "receipts/helper-seal-0970-final-v6.json"
OUTPUT_ROOT = ROOT / "capture/diagnostic-v3"
ROOT_GO = ROOT / "receipts/root-go-0970-v4.json"
TIME_FORMAT = '{"user_s":%U,"system_s":%S,"voluntary_context_switches":%w,"involuntary_context_switches":%c,"minor_faults":%R,"major_faults":%F,"max_rss_kb":%M}'


def bound(path: Path) -> dict:
    return binding(path)


def binary_binding(path: Path) -> dict:
    row = binding(path)
    row["executable"] = True
    return row


def check_expected(path: Path, expected: str, label: str) -> dict:
    row = bound(path)
    if row["sha256"] != expected:
        raise ValueError(f"{label}: expected {expected}, got {row['sha256']}")
    return row


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, default=ROOT / "plans/plan-0970-v6.json")
    parser.add_argument("--helper-seal", type=Path, default=HELPER_SEAL)
    args = parser.parse_args()
    output = args.output.resolve()
    if output.exists():
        raise ValueError(f"refusing to overwrite plan: {output}")

    design = check_expected(DESIGN, DESIGN_SHA256, "design v3")
    source_manifest = check_expected(SOURCE_MANIFEST, SOURCE_MANIFEST_SHA256, "source manifest")
    source_file_manifest = check_expected(SOURCE_FILE_MANIFEST, SOURCE_FILE_MANIFEST_SHA256, "source file manifest")
    manifest = strict_load_json(SOURCE_MANIFEST)
    exports = manifest["exports"]
    source = {
        "root": str(SOURCE_ROOT.resolve()),
        "manifest": source_manifest,
        "file_manifest": source_file_manifest,
        "tree_sha256": exports["baseline_tree_sha256"],
        "bytes": exports["baseline_bytes"],
        "file_count": exports["baseline_file_count"],
        "mode_policy": "manifest_mode_with_umask_0022_normalization",
    }
    validate_source_custody(source)
    binary = check_expected(BASELINE_BINARY, BASELINE_SHA256, "historical baseline binary")
    if not BASELINE_BINARY.stat().st_mode & 0o111:
        raise ValueError("historical baseline is not executable")
    fixtures = {}
    for case, path in FIXTURES.items():
        row = bound(path)
        # Preserve the external custody spelling even though test-data is a
        # read-only symlink into the checkout; the hash remains over the bytes.
        row["path"] = str(path.absolute())
        row["case"] = case
        fixtures[case] = row

    reuse = strict_load_json(BASELINE_REUSE_REVIEW)
    if reuse.get("status") != "pass_bounded_output_oracle_and_matched_baseline_reuse":
        raise ValueError("baseline reuse review is not the expected bounded pass")
    addendum = strict_load_json(TOOL_ADDENDUM)
    tool_rows = {row["resolved_path"]: row for row in addendum["tools"]}
    expected_tools = {
        "/usr/bin/perf": tool_rows["/usr/bin/perf"],
        "/usr/bin/time": tool_rows["/usr/bin/time"],
        "/usr/bin/python3.14": tool_rows["/usr/bin/python3.14"],
    }
    tools = {}
    for path, row in expected_tools.items():
        tools[path] = {"path": path, "sha256": row["resolved_path_sha256"], "bytes": row["stat_following_symlink"]["size_bytes"]}

    helper_seal = bound(args.helper_seal.resolve())
    check_binding(helper_seal, "helper seal")
    seal = strict_load_json(args.helper_seal.resolve())
    if seal.get("schema") != HELPER_SEAL_SCHEMA or seal.get("status") != "pass" or seal.get("design", {}).get("sha256") != DESIGN_SHA256:
        raise ValueError("helper seal is not a passing design-v3 seal")
    helper_paths = {}
    for name, row in seal.get("helpers", {}).items():
        path = check_binding(row, f"sealed helper {name}")
        helper_paths[name] = {"path": str(path), "sha256": row["sha256"], "bytes": path.stat().st_size}
    required_helpers = {"launcher", "protocol", "runner", "reducer", "plan_generator"}
    if set(helper_paths) != required_helpers:
        raise ValueError("helper seal binding set drifted")

    plan = {
        "schema": PLAN_SCHEMA,
        "status": "pre_capture",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "no_build_or_workload": True,
        "design": design,
        "reviews": {
            "baseline_reuse": {**bound(BASELINE_REUSE_REVIEW), "status": "pass"},
            "source_variability": bound(SOURCE_VARIABILITY_REVIEW),
            "independent_tool_review": {"status": "pending"},
        },
        "helper_seal": helper_seal,
        "correction_receipt": bound(CORRECTION_RECEIPT),
        "helpers": helper_paths,
        "source": source,
        "validator": {"path": str(VALIDATOR_PATH.resolve()), "sha256": VALIDATOR_SHA256, "schema": "litchi.performance.0969.protocol_common.v4"},
        "binary": binary,
        "fixtures": fixtures,
        "host": {
            "audit": bound(HOST_AUDIT),
            "tool_addendum": bound(TOOL_ADDENDUM),
            "telemetry_probe": bound(TELEMETRY_PROBE),
            "affinity_probe": bound(AFFINITY_PROBE),
            "tools": tools,
        },
        "matrix": {
            "cells": cells(),
            "pairs_per_cell": PAIRS,
            "warmups": WARMUPS,
            "samples": SAMPLES,
            "seed": BOOTSTRAP_SEED,
            "bootstrap_resamples": BOOTSTRAP_RESAMPLES,
            "order": {"odd_pair": ["unpinned", "pinned"], "even_pair": ["pinned", "unpinned"]},
        },
        "affinity": {
            "selector": "sealed_live_mask_exact:[0..31]",
            "pinned": "one_element_mask",
            "unpinned": "fresh_full_allowed_mask",
            "child_self_report": True,
            "requested_actual_must_match": True,
            "expected_allowed": EXPECTED_ALLOWED_AFFINITY,
            "expected_pin_cpu": EXPECTED_PIN_CPU,
        },
        "wrapper": {
            "events": list(EVENTS),
            "time_format": TIME_FORMAT,
            "launcher": str((ROOT / "tools/launch_baseline_v2.py").resolve()),
            "scope": {
                "perf": "perf measures the launched GNU time command and inherited launcher/baseline process tree; perf bookkeeping is excluded",
                "gnu_time": "launcher startup plus same-PID execed historical baseline",
                "harness": "baseline Instant samples only; excludes wrappers and sidecar I/O",
            },
        },
        "execution": {
            "serial": True,
            "retry": False,
            "timeout_seconds": TIMEOUT_SECONDS,
            "process_quiescence_required": True,
            "cargo_processes": 0,
            "binary_rebuild": False,
            "candidate": False,
            "lc_all": REQUIRED_LOCALE,
            "lang": REQUIRED_LOCALE,
            "python_dontwritebytecode": "1",
        },
        "environment": {"LC_ALL": REQUIRED_LOCALE, "LANG": REQUIRED_LOCALE, "PYTHONDONTWRITEBYTECODE": "1"},
        "analysis": {
            "diagnostic_only": True,
            "statuses": ["complete", "incomplete"],
            "no_qualification": True,
            "seed": BOOTSTRAP_SEED,
            "bootstrap_resamples": BOOTSTRAP_RESAMPLES,
            "p99": "descriptive_only",
            "whole_process_perf_ratios": ["cycles/instructions", "cycles/task-clock"],
            "zero_dispersion_ratio": "null_with_reason",
        },
        "outputs": {
            "capture_root": str(OUTPUT_ROOT.resolve()),
            "root_go": {"path": str(ROOT_GO.resolve()), "status": "pending"},
        },
    }
    # The plan is checked in memory by writing an exclusive temporary sibling,
    # so admission and final emission use the same JSON bytes/shape.
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(f".{output.name}.check")
    if temporary.exists():
        raise ValueError(f"temporary plan path already exists: {temporary}")
    write_json_exclusive(temporary, plan)
    try:
        from run_0970_v3 import check_plan
        check_plan(temporary, require_pre_capture=True)
    finally:
        temporary.unlink(missing_ok=True)
    write_json_exclusive(output, plan)
    print(json.dumps({"schema": PLAN_SCHEMA, "status": "pre_capture_plan_sealed", "plan": str(output)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
