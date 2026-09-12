#!/usr/bin/env python3
"""Run bounded negative probes against the 0525 evidence verifier.

The probes use immutable retained receipts, synthetic receipt timelines, and
in-memory report/schema values.  They do not build, capture, run the frozen
analyzers, mutate source, or overwrite an evidence artifact.  A JSON result is
printed to stdout; pass ``--output`` only when a caller wants to retain the
small probe receipt separately.
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import importlib.util
import json
from pathlib import Path
from typing import Any, Callable


HERE = Path(__file__).resolve().parent
VERIFY_PATH = HERE / "verify.py"
NUMERIC_PATH = HERE / "analyze.py"


class ProbeFailure(AssertionError):
    """A corruption probe was unexpectedly accepted."""


def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ProbeFailure(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


VERIFY = load_module(VERIFY_PATH, "xlsx_0525_verify_for_negative_probes")
NUMERIC = load_module(NUMERIC_PATH, "xlsx_0525_numeric_for_negative_probes")


def expect_rejection(name: str, operation: Callable[[], Any]) -> dict[str, Any]:
    try:
        operation()
    except (VERIFY.EvidenceError, NUMERIC.BASE.EvidenceError,
            AssertionError, KeyError, TypeError, ValueError) as error:
        return {"case": name, "rejected": True,
                "reason": str(error) or "evidence invariant rejected"}
    raise ProbeFailure(f"corrupted evidence was accepted: {name}")


def source_mismatch(plan: dict[str, Any]) -> dict[str, Any]:
    path = HERE / "baseline" / "build-normal.receipt.json"
    wrong_manifest = "0" * 64
    return expect_rejection(
        "receipt_source_manifest_mismatch",
        lambda: VERIFY.validate_receipt(
            path, "baseline", plan, wrong_manifest, None,
            expected_binary=None, check_binary=True,
        ),
    )


def quality_rows_for(stage: str) -> dict[str, Any]:
    return {
        "name": "check-xlsx-tests.receipt.json",
        "stage": stage,
        "receipt_sha256": "0" * 64,
        "exit_code": 0,
        "executed_tests": 0,
    }


def with_quality_document(document: dict[str, Any], operation: Callable[[], Any]) -> Any:
    original = VERIFY.read_json

    def read_json(path: Path) -> Any:
        if Path(path) == HERE / "quality-summary.json":
            return document
        return original(path)

    VERIFY.read_json = read_json
    try:
        return operation()
    finally:
        VERIFY.read_json = original


def with_quality_commands(commands: list[tuple[str, list[str]]],
                          operation: Callable[[], Any]) -> Any:
    original = VERIFY._quality_commands
    VERIFY._quality_commands = lambda: commands
    try:
        return operation()
    finally:
        VERIFY._quality_commands = original


def preflight_alias_probes(plan: dict[str, Any]) -> list[dict[str, Any]]:
    stage = "preflight-1"
    stage_dir = HERE / stage
    manifest_sha = VERIFY.sha(stage_dir / "source-manifest.json")
    candidate_sha = VERIFY.sha(HERE / "candidate" / "source-manifest.json")
    failed_receipt = stage_dir / "check-xlsx-tests.receipt.json"
    VERIFY.require(failed_receipt.is_file(),
                   "retained preflight-1 failure receipt is missing")
    receipt = VERIFY.read_json(failed_receipt)
    VERIFY.require(receipt.get("exit_code") != 0,
                   "preflight-1 is no longer a failed retained attempt")

    mismatch_item = {
        "stage": stage,
        "manifest_sha256": manifest_sha,
        "candidate_manifest_equal": False,
        "receipts": [],
    }
    document = {"checks": [quality_rows_for(stage)]}
    one_xlsx_command = [("xlsx-tests", [])]
    mismatch = expect_rejection(
        "preflight_alias_requires_candidate_manifest",
        lambda: with_quality_commands(
            one_xlsx_command,
            lambda: with_quality_document(
                document,
                lambda: VERIFY.validate_check_receipts(
                    plan, "candidate",
                    {"manifest_sha256": candidate_sha},
                    [mismatch_item],
                ),
            ),
        ),
    )

    failed_item = copy.deepcopy(mismatch_item)
    failed_item["candidate_manifest_equal"] = True
    failed = expect_rejection(
        "preflight_failed_receipt_cannot_alias_quality_gate",
        lambda: with_quality_commands(
            one_xlsx_command,
            lambda: with_quality_document(
                document,
                lambda: VERIFY.validate_check_receipts(
                    plan, "candidate",
                    {"manifest_sha256": candidate_sha},
                    [failed_item],
                ),
            ),
        ),
    )
    return [mismatch, failed]


def timeline_rows(plan: dict[str, Any], order: list[tuple[str, int]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    names = set(VERIFY.native_jobs(plan))
    base = dt.datetime(2026, 9, 12, tzinfo=dt.timezone.utc)
    tick = 0
    for stage, repeat in order:
        block = sorted(name for name in names
                       if name.startswith(f"native-r{repeat}-"))
        for name in block:
            tick += 1
            start = base + dt.timedelta(seconds=tick)
            rows.append({
                "name": name,
                "stage": stage,
                "path": f"{stage}/{name}.receipt.json",
                "start_utc": start.isoformat(),
            })
    return rows


def abba_inversion(plan: dict[str, Any]) -> dict[str, Any]:
    inverted = [("baseline", 1), ("candidate", 2),
                ("candidate", 1), ("baseline", 2)]
    return expect_rejection(
        "native_abba_block_inversion",
        lambda: VERIFY.validate_native_order(plan, timeline_rows(plan, inverted)),
    )


def primary_duplicate(plan: dict[str, Any]) -> dict[str, Any]:
    def row(repeat: int, shape: str, elapsed: int, commit: int) -> dict[str, Any]:
        return {
            "kind": "primary", "guard": None, "repeat": repeat,
            "shape": shape,
            "timing": {"elapsed_ns": {"p50": elapsed},
                       "commit_ns": {"p50": commit}},
        }

    baseline_rows = [
        row(1, "medium", 1000, 500),
        row(1, "dense-sparse", 1000, 500),
        row(2, "medium", 1000, 500),
        # Duplicate repeat/shape key; dense-sparse repeat 2 is omitted.
        row(1, "medium", 1000, 500),
    ]
    candidate_rows = [
        row(1, "medium", 900, 400),
        row(1, "dense-sparse", 900, 400),
        row(2, "medium", 900, 400),
        row(2, "dense-sparse", 900, 400),
    ]
    evidence = {
        "baseline": {"native": {"rows": baseline_rows}},
        "candidate": {"native": {"rows": candidate_rows}},
    }
    return expect_rejection(
        "primary_admission_duplicate_row",
        lambda: NUMERIC._native_admission(evidence, plan),
    )


def cleanup_schema(plan: dict[str, Any]) -> list[dict[str, Any]]:
    original = VERIFY.read_json
    malformed = {
        "removed": plan["owned_paths"],
        "accessible_process_references": [],
        "owned_paths_absent": True,
        "python_cache_absent": True,
    }

    def read_json(path: Path) -> Any:
        if Path(path) == HERE / "cleanup.json":
            return malformed
        return original(path)

    VERIFY.read_json = read_json
    try:
        cleanup = expect_rejection(
            "cleanup_missing_plan_binding",
            lambda: VERIFY.validate_cleanup(plan),
        )
    finally:
        VERIFY.read_json = original

    unsafe_stage = expect_rejection(
        "quality_schema_unsafe_stage_path",
        lambda: VERIFY.quality_stage_and_name(
            {"name": "../check-xlsx-tests.receipt.json"}, "candidate"
        ),
    )
    return [cleanup, unsafe_stage]


def run(output: Path | None = None) -> dict[str, Any]:
    plan = VERIFY.plan_data()
    checks = [source_mismatch(plan)]
    checks.extend(preflight_alias_probes(plan))
    checks.append(abba_inversion(plan))
    checks.append(primary_duplicate(plan))
    checks.extend(cleanup_schema(plan))
    result = {
        "schema": "litchi-0525-verifier-negative-probes-v1",
        "status": "pass",
        "valid_controls_pass": True,
        "checks": checks,
        "scope": (
            "Seven bounded custody/schema probes using retained receipts and "
            "in-memory values; no build, capture, analyzer replay, source, "
            "frozen-tool, or retained-artifact mutation."
        ),
    }
    if output is not None:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", nargs="?", type=Path)
    parser.add_argument("--output", dest="output_option", type=Path)
    args = parser.parse_args()
    if args.output is not None and args.output_option is not None:
        parser.error("provide the output path either positionally or with --output")
    output = args.output_option or args.output
    try:
        result = run(output)
    except (VERIFY.EvidenceError, NUMERIC.BASE.EvidenceError,
            AssertionError, KeyError, TypeError, ValueError, OSError) as error:
        print(f"verify_test.py: probe failed: {error}")
        return 2
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
