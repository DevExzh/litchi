#!/usr/bin/env python3
"""Aggregate one passing 0891 quality attempt into the packet receipt.

This helper only reads and hashes custody artifacts.  It never invokes Cargo
and creates ``quality.json`` exactly once after all seven gates pass.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import stat
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
QUALITY = PACKET / "quality.json"
FREEZE = PACKET / "candidate-freeze-v3.json"
SOURCE_REVIEW = PACKET / "source-review-v3.json"
FIXTURE_CENSUS = PACKET / "fixture-census.json"
GATES = (
    "focused_x14ac_reuse_tests",
    "xlsx_full_lib_tests",
    "xlsx_all_features_check",
    "xlsx_clippy",
    "xlsx_rustdoc",
    "xlsx_fmt",
    "crate_boundaries",
)


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"expected regular non-symlink file: {path}")
    metadata = path.stat()
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return {"path": str(path), "bytes": metadata.st_size, "mode": stat.S_IMODE(metadata.st_mode), "sha256": digest.hexdigest()}


def read_json(path: Path) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"missing or symlinked JSON: {path}")
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise RuntimeError(f"expected JSON object: {path}")
    return value


def path_from(raw: Any, *, base: Path = PACKET) -> Path:
    if not isinstance(raw, str):
        raise RuntimeError("descriptor path is missing")
    value = Path(raw)
    return (value if value.is_absolute() else base / value).resolve()


def checked_descriptor(raw: dict[str, Any], label: str, *, base: Path = PACKET) -> dict[str, Any]:
    if not isinstance(raw, dict):
        raise RuntimeError(f"{label} descriptor is missing")
    path = path_from(raw.get("path"), base=base)
    actual = descriptor(path)
    for field in ("bytes", "mode", "sha256"):
        expected = raw.get(field)
        if field == "mode" and isinstance(expected, str):
            expected = int(expected, 8) if expected.startswith("0o") else int(expected, 8)
        if field in raw and expected != actual[field]:
            raise RuntimeError(f"{label} {field} changed")
    return actual


def write_once(value: dict[str, Any]) -> None:
    if QUALITY.exists() or QUALITY.is_symlink():
        raise SystemExit(f"refusing to overwrite existing aggregate: {QUALITY}")
    with QUALITY.open("x", encoding="utf-8") as destination:
        json.dump(value, destination, indent=2, sort_keys=True)
        destination.write("\n")
        destination.flush()
        os.fsync(destination.fileno())
    directory = os.open(PACKET, os.O_RDONLY | getattr(os, "O_DIRECTORY", 0))
    try:
        os.fsync(directory)
    finally:
        os.close(directory)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", default="quality-v3", help="successful quality attempt name")
    args = parser.parse_args()
    if not args.attempt or Path(args.attempt).name != args.attempt or args.attempt in {".", ".."}:
        raise SystemExit("--attempt must be one path component")
    attempt = PACKET / "attempts" / args.attempt
    receipt_path = attempt / "quality-receipt.json"
    receipt = read_json(receipt_path)
    if receipt.get("status") != "pass" or receipt.get("performance_claim") is not False:
        raise SystemExit("selected quality attempt is not a passing non-claim receipt")
    commands = receipt.get("commands")
    if not isinstance(commands, list) or [item.get("name") for item in commands] != list(GATES):
        raise SystemExit("quality attempt does not contain the exact seven gates in order")
    if receipt.get("failures") or receipt.get("skipped_after_first_failure"):
        raise SystemExit("quality attempt records a failure or skipped gate")
    for command in commands:
        if command.get("returncode") != 0 or command.get("pass") is not True:
            raise SystemExit(f"quality gate did not pass: {command.get('name')}")
        log = path_from(command.get("log"))
        checked_descriptor(command.get("log_identity", {}), f"{command.get('name')} log")
        if log != path_from(command["log_identity"]["path"]):
            raise SystemExit(f"quality gate log binding mismatch: {command.get('name')}")
    for name in ("focused_x14ac_reuse_tests", "xlsx_full_lib_tests"):
        command = next(item for item in commands if item.get("name") == name)
        if not isinstance(command.get("passed_count"), int) or command["passed_count"] <= 0:
            raise SystemExit(f"quality gate has no positive test count: {name}")
    if receipt.get("source_before") != receipt.get("source_after"):
        raise SystemExit("quality source custody changed during the attempt")
    if receipt.get("root_lock_before") != receipt.get("root_lock_after"):
        raise SystemExit("root Cargo.lock changed during the attempt")

    freeze = read_json(FREEZE)
    if freeze.get("status") != "frozen" or freeze.get("performance_claim") not in (None, False):
        raise SystemExit("candidate-freeze-v3 is not frozen/non-claim")
    source_review = read_json(SOURCE_REVIEW)
    if source_review.get("status") != "approved_for_measurement" or source_review.get("performance_claim") not in (None, False):
        raise SystemExit("source-review-v3 is not approved/non-claim")
    fixture = read_json(FIXTURE_CENSUS)
    if fixture.get("status") != "pass":
        raise SystemExit("fixture census is not passing")

    receipt_descriptor = descriptor(receipt_path)
    runner_raw = receipt.get("runner")
    runner = checked_descriptor(runner_raw, "executed quality runner")
    if path_from(runner_raw.get("path")) != (attempt / "quality_runner.py").resolve():
        raise SystemExit("quality receipt runner is not the selected attempt copy")
    freeze_binding = checked_descriptor(receipt.get("freeze", {}).get("identity_after", {}), "quality freeze")
    if path_from(receipt.get("freeze", {}).get("identity_after", {}).get("path")) != FREEZE.resolve():
        raise SystemExit("quality receipt does not bind candidate-freeze-v3")
    if freeze_binding != descriptor(FREEZE):
        raise SystemExit("quality receipt freeze binding differs from candidate-freeze-v3")
    source_review_descriptor = descriptor(SOURCE_REVIEW)
    fixture_descriptor = descriptor(FIXTURE_CENSUS)
    result = {
        "schema": "litchi.xlsx-x14ac-quality-aggregate-0891.v1",
        "status": "pass",
        "purpose": "aggregate_passing_x14ac_quality_gates",
        "performance_claim": False,
        "created_utc": dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z"),
        "attempt": str(attempt.resolve()),
        "quality_receipt": receipt_descriptor,
        "executed_runner": runner,
        "candidate_freeze": descriptor(FREEZE),
        "source_review": source_review_descriptor,
        "fixture_census": fixture_descriptor,
        "gates": commands,
        "source_before": receipt["source_before"],
        "source_after": receipt["source_after"],
        "root_lock_before": receipt["root_lock_before"],
        "root_lock_after": receipt["root_lock_after"],
    }
    write_once(result)
    print(json.dumps({"status": result["status"], "gates": len(commands), "quality": str(QUALITY)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
