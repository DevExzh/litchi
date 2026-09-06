#!/usr/bin/env python3
"""Validate the retained 0436 terminal lifecycle and rederive its summary.

The lifecycle gate is intentionally independent from the matrix verifier's
report checks.  It verifies the planned terminal receipts, source custody,
retained pilot/profile artifacts, cleanup gates, and the canonical summary
derivation.  It never launches a workload or rewrites the sealed inventory.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any

ROOT = Path(__file__).resolve().parent
MARKER_NAME = "workload-verify.json"


def module(name: str):
    spec = importlib.util.spec_from_file_location("change0436_lifecycle_" + name, ROOT / (name + ".py"))
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {name}.py")
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


def is_profile_marker(path: Path) -> bool:
    try:
        relative = path.relative_to(ROOT).parts
    except ValueError:
        return False
    if len(relative) < 3 or relative[0] != "profiles" or relative[-1] != MARKER_NAME:
        return False
    if path.read_bytes() != b"VALID\n":
        raise ValueError(f"{path}: profile marker is not exactly VALID")
    return True


def validate_planned(v: Any) -> None:
    planned = v.load(ROOT / "planned-checks.json")
    if not isinstance(planned, dict):
        v.fail("planned-checks.json", "expected an object")
    for key in ("required_pass", "required_review"):
        names = planned.get(key)
        if not isinstance(names, list) or any(not isinstance(name, str) for name in names):
            v.fail(f"planned-checks.json.{key}", "expected a list of bundle-relative paths")
        expected_status = "pass" if key == "required_pass" else "failed"
        for name in names:
            path = v.bundle_path(name, f"planned-checks.{key}")
            receipt = v.load(path, name)
            if receipt.get("status") != expected_status:
                v.fail(name, f"required check has wrong terminal status (expected {expected_status})")


def validate_receipt_artifacts(v: Any, path: Path, row: dict[str, Any]) -> None:
    artifacts = row.get("artifacts")
    if artifacts is None:
        return
    if not isinstance(artifacts, dict):
        v.fail(str(path), "artifact inventory must be an object")
    for name, record in artifacts.items():
        if not isinstance(name, str) or not isinstance(record, dict):
            v.fail(str(path), "artifact inventory entry is malformed")
        v.artifact_path(path, record if "path" in record else dict(record, path=name), f"{path}.artifacts.{name}")


def validate_terminal_receipt(v: Any, path: Path, row: dict[str, Any]) -> None:
    """Validate custody for every retained pass or failure receipt."""

    label = str(path)
    if row.get("status") not in {"pass", "failed"}:
        v.fail(label, "receipt is not terminal")
    if "source_before" not in row or "source_after" not in row:
        v.fail(label, "receipt has incomplete source custody")
    v.check_source_manifest(row["source_before"], label + ".source_before")
    v.check_source_manifest(row["source_after"], label + ".source_after")
    if row["source_before"] != row["source_after"] or (
        "source_unchanged" in row and row.get("source_unchanged") is not True
    ):
        v.fail(label, "receipt does not prove unchanged source custody")
    if row.get("log") is not None:
        v.artifact_path(path, row["log"], label + ".log")
    validate_receipt_artifacts(v, path, row)


def validate_source_receipts(v: Any) -> int:
    """Validate every source-bound terminal receipt and match seal inventory."""

    expected = v.load(ROOT / "expected-checks.json")
    if not isinstance(expected, dict):
        v.fail("expected-checks.json", "expected-checks must be an object")
    actual: dict[str, Any] = {}
    for path in sorted(ROOT.rglob("*.json")):
        if path.name in {"compression.json", "expected-checks.json"}:
            continue
        if is_profile_marker(path):
            continue
        row = v.load(path, str(path))
        if not isinstance(row, dict) or "source_before" not in row:
            continue
        name = path.relative_to(ROOT).with_suffix("").as_posix()
        actual[name] = row.get("status")
        validate_terminal_receipt(v, path, row)
        if path.parent == ROOT / "checks":
            if row.get("driver_sha256") != v.sha(ROOT / "check.py"):
                v.fail(name, "command custody driver differs")
            if row.get("status") == "pass" and row.get("exit_code") != 0:
                v.fail(name, "passing command has nonzero exit")
            if row.get("status") == "pass" and row.get("passed_tests", 1) <= 0:
                v.fail(name, "passing test command ran no tests")
    if actual != expected:
        v.fail("expected-checks.json", "terminal receipt inventory differs")
    return len(actual)


def validate_cleanup_gates(v: Any, stage: str) -> None:
    if stage == "precleanup":
        return
    for name in ("precleanup-portable", "task-cleanup"):
        path = ROOT / "checks" / (name + ".json")
        row = v.load(path, str(path))
        if row.get("status") != "pass" or row.get("exit_code") != 0:
            v.fail(str(path), "required cleanup lifecycle check is not passing")


def validate_pilots_and_profiles(v: Any) -> tuple[int, int]:
    pilots = 0
    pilot_root = ROOT / "pilots"
    if pilot_root.is_dir():
        for path in sorted(pilot_root.rglob("*-receipt.json")):
            row = v.load(path, str(path))
            validate_terminal_receipt(v, path, row)
            if row.get("status") == "pass" and row.get("exit_code") != 0:
                v.fail(str(path), "passing pilot has nonzero exit")
            relative = path.relative_to(pilot_root).parts
            if len(relative) < 3 or relative[0] not in {"before", "after"} or relative[1] != "initial":
                v.fail(str(path), "pilot is outside pilots/{before,after}/initial")
            pilots += 1

    profiles = 0
    profile_root = ROOT / "profiles"
    if profile_root.is_dir():
        for path in sorted(profile_root.rglob("receipt.json")):
            row = v.load(path, str(path))
            validate_terminal_receipt(v, path, row)
            relative = path.relative_to(profile_root).parts
            if len(relative) != 3 or relative[0] not in {"before", "after"} or relative[1] not in {"stat", "record"} or relative[2] != "receipt.json":
                v.fail(str(path), "profile is outside profiles/{before,after}/{stat,record}")
            profiles += 1
    if pilots != 12:
        v.fail("pilots", f"expected 12 initial pilot receipts, found {pilots}")
    if profiles != 4:
        v.fail("profiles", f"expected four formal profile receipts, found {profiles}")
    return pilots, profiles


def validate_summary_shape(v: Any, value: Any) -> None:
    if not isinstance(value, dict):
        v.fail("summary.json", "summary must be an object")
    matrix = value.get("matrix")
    if not isinstance(matrix, dict):
        v.fail("summary.json.matrix", "summary matrix must be an object")
    if matrix.get("formal_reports") != 24 or matrix.get("retained_samples") != 720:
        v.fail("summary.json.matrix", "0436 requires 24 reports and 720 retained samples")


def validate(stage: str) -> dict[str, Any]:
    v = module("verify")
    validate_planned(v)
    validate_cleanup_gates(v, stage)
    terminal_receipts = validate_source_receipts(v)
    pilots, profiles = validate_pilots_and_profiles(v)
    summary = module("summary")
    retained_value = v.load(ROOT / "summary.json")
    derived_value = summary.derive()
    validate_summary_shape(v, retained_value)
    validate_summary_shape(v, derived_value)
    retained = summary.canonical(retained_value)
    derived = summary.canonical(derived_value)
    if retained != derived:
        v.fail("summary.json", "retained summary differs from independent derivation")
    preparatory = module("preparatory-summary")
    preparatory_value = preparatory.derive()
    retained_preparatory = v.load(ROOT / "batching-hypothesis.json")
    if retained_preparatory != preparatory_value:
        v.fail("batching-hypothesis.json", "retained preparatory summary differs from independent derivation")
    decision = module("decision")
    if v.load(ROOT / "decision.json") != decision.derive():
        v.fail("decision.json", "retention decision differs from independent derivation")
    return {
        "status": "pass",
        "change": 436,
        "stage": stage,
        "terminal_receipts": terminal_receipts,
        "pilots": pilots,
        "profiles": profiles,
        "summary_rederived": True,
        "preparatory_summary_rederived": True,
        "formal_reports": 24,
        "retained_samples": 720,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("precleanup", "aftercleanup", "final"), default="final")
    args = parser.parse_args()
    try:
        print(json.dumps(validate(args.stage), sort_keys=True))
    except (OSError, ValueError, KeyError, TypeError, AssertionError) as error:
        print("INVALID: " + str(error), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
