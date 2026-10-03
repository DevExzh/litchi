#!/usr/bin/env python3
"""Verify candidate source, fixture, harness, and executable custody."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import stat


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
SOURCE = Path("/tmp/litchi-0890-correctness-source")
BASELINE_RECEIPT = PACKET / "attempts/baseline-f790c2228c-retry1/baseline-receipt.json"
CANDIDATE_RECEIPT = PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1/candidate-receipt.json"
OVERLAY_RECEIPT = PACKET / "candidate-diagnostic-overlay.json"
FREEZE = PACKET / "candidate-freeze-v2.json"
OUTPUT = PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1/post-candidate-source-verification.json"
TARGET = ROOT / "target/litchi-0890-correctness/debug/xml_attribute_bounds"
BASELINE_COPY = Path("/tmp/litchi-0890-correctness-baseline/xml_attribute_bounds")
ROOT_DIAGNOSTIC = (
    "tools/perf-baseline/src/bin/xml_attribute_bounds.rs",
    "tools/perf-baseline/src/bin/xml_attribute_bounds/boundary_tests.rs",
    "tools/perf-baseline/src/bin/xml_attribute_bounds/input_tests.rs",
    "tools/perf-baseline/src/bin/xml_attribute_bounds/qualification_tests.rs",
)
CANDIDATE_PATHS = (
    "crates/litchi-ooxml-common/src/mce/codec.rs",
    "crates/litchi-ooxml-common/src/mce/tests.rs",
)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def identity(path: Path) -> dict[str, object]:
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "mode": oct(stat.S_IMODE(path.stat().st_mode)),
        "sha256": sha(path),
    }


def main() -> int:
    baseline = json.loads(BASELINE_RECEIPT.read_text())
    candidate = json.loads(CANDIDATE_RECEIPT.read_text())
    overlay = json.loads(OVERLAY_RECEIPT.read_text())
    freeze = json.loads(FREEZE.read_text())
    if baseline.get("status") != "pass" or candidate.get("status") != "pass" or overlay.get("status") != "pass":
        raise SystemExit("baseline, overlay, or candidate receipt is not passing")
    frozen = {entry["path"]: entry for entry in freeze["files"]}
    baseline_files = {entry["path"]: entry for entry in baseline["source"]["files"]}
    if set(frozen) != set(CANDIDATE_PATHS):
        raise SystemExit("candidate freeze path set mismatch")
    if set(overlay["after"]) != set(CANDIDATE_PATHS):
        raise SystemExit("overlay path set mismatch")
    for path, entry in baseline_files.items():
        current = identity(SOURCE / path)
        expected = frozen[path] if path in frozen else entry
        if current["bytes"] != expected["bytes"] or current["sha256"] != expected["sha256"]:
            raise SystemExit(f"isolated candidate source changed after overlay: {path}")
    source_paths = {str(path.relative_to(SOURCE)) for path in SOURCE.rglob("*") if path.is_file() or path.is_symlink()}
    if source_paths != set(baseline_files) | {".owner-0890-source"}:
        raise SystemExit("isolated candidate source file population changed")
    root_candidates = {}
    for path in CANDIDATE_PATHS:
        current = identity(ROOT / path)
        expected = frozen[path]
        if current["bytes"] != expected["bytes"] or current["sha256"] != expected["sha256"]:
            raise SystemExit(f"root candidate source changed after freeze: {path}")
        root_candidates[path] = current
    root_diagnostic = {}
    for path in ROOT_DIAGNOSTIC:
        current = identity(ROOT / path)
        expected = baseline_files[path]
        if current["bytes"] != expected["bytes"] or current["sha256"] != expected["sha256"]:
            raise SystemExit(f"root diagnostic harness changed from f790: {path}")
        root_diagnostic[path] = current
    candidate_binary = identity(TARGET)
    if candidate_binary != candidate.get("candidate_binary"):
        raise SystemExit("candidate binary changed after the root-based test gate")
    if candidate_binary == baseline.get("binary") or candidate_binary == identity(BASELINE_COPY):
        raise SystemExit("candidate executable unexpectedly equals retained baseline")
    test_command = next(item for item in candidate["commands"] if item["name"] == "cargo-test-diagnostic")
    test_log = PACKET / test_command["log"]
    test_text = test_log.read_text()
    if test_command["cwd"] != str(ROOT) or "running 24 tests" not in test_text or "24 passed; 0 failed" not in test_text:
        raise SystemExit("candidate diagnostic test custody does not prove 24 passing tests")
    fixtures = []
    for expected in baseline["fixtures"]:
        path = Path(expected["path"])
        current = identity(path)
        if current["bytes"] != expected["bytes"] or current["sha256"] != expected["sha256"]:
            raise SystemExit(f"fixture changed: {path}")
        fixtures.append(current)
    lock_receipt_path = PACKET / "attempts/baseline-f790c2228c-retry1/root-lock-custody.json"
    lock = json.loads(lock_receipt_path.read_text())
    original_lock = Path(lock["original"]["path"])
    custody_lock = Path(lock["custody"]["path"])
    if sha(original_lock) != lock["original"]["sha256"] or sha(custody_lock) != lock["custody"]["sha256"] or original_lock.read_bytes() != custody_lock.read_bytes():
        raise SystemExit("root Cargo.lock custody changed")
    result = {
        "schema": "litchi.mce-0890-post-candidate-source-verification.v1",
        "status": "pass",
        "candidate_receipt": {"path": str(CANDIDATE_RECEIPT.relative_to(PACKET)), "sha256": sha(CANDIDATE_RECEIPT)},
        "overlay_receipt": {"path": str(OVERLAY_RECEIPT.relative_to(PACKET)), "sha256": sha(OVERLAY_RECEIPT)},
        "candidate_freeze": {"path": str(FREEZE.relative_to(PACKET)), "sha256": sha(FREEZE)},
        "source_root": str(SOURCE),
        "source_files": len(baseline_files),
        "isolated_source_rechecked": True,
        "root_candidates": root_candidates,
        "root_diagnostic_harness": root_diagnostic,
        "candidate_binary": candidate_binary,
        "baseline_binary": baseline["binary"],
        "baseline_copy": identity(BASELINE_COPY),
        "diagnostic_tests": {"count": 24, "passed": 24, "failed": 0, "cwd": test_command["cwd"], "log": test_command["log"]},
        "fixtures": fixtures,
        "root_lock_custody": {"path": str(lock_receipt_path.relative_to(PACKET)), "sha256": sha(lock_receipt_path)},
    }
    OUTPUT.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
