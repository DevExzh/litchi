#!/usr/bin/env python3
"""Remove only the owned 0890 correctness scratch after review authorization.

The command requires ``--remove`` and a passing ``correctness-review.json``.
It writes a create-only pre-removal custody receipt, removes the four exact
marked roots, verifies every retained packet artifact, and atomically writes a
final receipt.  It never traverses or removes native/allocator build roots or
their execution copies.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent

SOURCE = Path("/tmp/litchi-0890-correctness-source")
TARGET = ROOT / "target/litchi-0890-correctness"
BASELINE = Path("/tmp/litchi-0890-correctness-baseline")
ROOT_LOCK = Path("/tmp/litchi-0890-correctness-root-lock")

OWNED = (
    ("diagnostic_source", SOURCE, ".owner-0890-source", "litchi 0890 correctness source archive\n"),
    ("correctness_target", TARGET, ".owner-0890-target", "litchi 0890 correctness target\n"),
    ("baseline_executable_custody", BASELINE, ".owner-0890-baseline", "litchi 0890 baseline executable custody\n"),
    ("root_lock_custody", ROOT_LOCK, ".owner-0890-root-lock", "litchi 0890 root Cargo.lock custody\n"),
)

FORBIDDEN = (
    Path("/tmp/litchi-0890-source"),
    ROOT / "target/litchi-0890-build",
    PACKET / "binaries/native",
    PACKET / "binaries/allocator",
)

REVIEW = PACKET / "correctness-review.json"
PRE_REMOVAL = PACKET / "cleanup-correctness-pre-removal.json"
FINAL = PACKET / "cleanup-correctness.json"

BASELINE_RECEIPT = PACKET / "attempts/baseline-f790c2228c-retry1/baseline-receipt.json"
BASELINE_POSTCHECK = PACKET / "attempts/baseline-f790c2228c-retry1/post-baseline-source-verification.json"
ROOT_LOCK_RECEIPT = PACKET / "attempts/baseline-f790c2228c-retry1/root-lock-custody.json"
CANDIDATE_RECEIPT = PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1/candidate-receipt.json"
CANDIDATE_POSTCHECK = PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1/post-candidate-source-verification.json"
CONTINUATION_RECEIPT = PACKET / "attempts/owner-continuation-v2/owner-continuation-receipt.json"
BINARY_RECEIPT = PACKET / "binaries/correctness-binaries-receipt-v2.json"
FREEZE = PACKET / "candidate-freeze-v2.json"

CORRECTNESS_BEFORE = PACKET / "binaries/correctness-before/xml_attribute_bounds"
CORRECTNESS_AFTER = PACKET / "binaries/correctness-after/xml_attribute_bounds"


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def identity(path: Path, *, relative_to: Path | None = None) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"expected regular file without symlink: {path}")
    metadata = path.stat()
    shown = path if relative_to is None else path.relative_to(relative_to)
    return {
        "path": str(shown),
        "bytes": metadata.st_size,
        "mode": oct(stat.S_IMODE(metadata.st_mode)),
        "sha256": sha256(path),
    }


def write_exclusive(path: Path, value: dict[str, Any]) -> None:
    """Create a receipt exactly once, flushing both file and directory."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x") as destination:
        json.dump(value, destination, indent=2, sort_keys=True)
        destination.write("\n")
        destination.flush()
        os.fsync(destination.fileno())
    directory = os.open(path.parent, os.O_RDONLY | getattr(os, "O_DIRECTORY", 0))
    try:
        os.fsync(directory)
    finally:
        os.close(directory)


def write_atomic(path: Path, value: dict[str, Any]) -> None:
    """Atomically install the final receipt and flush its parent directory."""
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    if temporary.exists() or temporary.is_symlink():
        raise RuntimeError(f"temporary cleanup receipt already exists: {temporary}")
    try:
        with temporary.open("x") as destination:
            json.dump(value, destination, indent=2, sort_keys=True)
            destination.write("\n")
            destination.flush()
            os.fsync(destination.fileno())
        os.replace(temporary, path)
        directory = os.open(path.parent, os.O_RDONLY | getattr(os, "O_DIRECTORY", 0))
        try:
            os.fsync(directory)
        finally:
            os.close(directory)
    finally:
        if temporary.exists() or temporary.is_symlink():
            temporary.unlink()


def read_object(path: Path) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"required receipt is missing or symlinked: {path}")
    value = json.loads(path.read_text())
    if not isinstance(value, dict):
        raise RuntimeError(f"receipt is not a JSON object: {path}")
    return value


def receipt_descriptor(path: Path) -> dict[str, Any]:
    return identity(path, relative_to=PACKET)


def require_pass(path: Path) -> dict[str, Any]:
    value = read_object(path)
    if value.get("status") != "pass":
        raise RuntimeError(f"required receipt is not passing: {path}")
    return value


def tree_inventory(root: Path, marker_name: str, marker_text: str) -> dict[str, Any]:
    if root.is_symlink() or not root.is_dir():
        raise RuntimeError(f"owned cleanup root is missing or symlinked: {root}")
    marker = root / marker_name
    if marker.is_symlink() or not marker.is_file() or marker.read_text() != marker_text:
        raise RuntimeError(f"owned cleanup marker mismatch: {marker}")

    entries: list[dict[str, Any]] = []
    file_bytes = 0
    for path in sorted(root.rglob("*"), key=lambda candidate: candidate.relative_to(root).as_posix()):
        metadata = path.lstat()
        relative = path.relative_to(root).as_posix()
        if stat.S_ISLNK(metadata.st_mode):
            raise RuntimeError(f"refusing cleanup root containing symlink: {root}/{relative}")
        if stat.S_ISDIR(metadata.st_mode):
            kind = "directory"
        elif stat.S_ISREG(metadata.st_mode):
            kind = "file"
            file_bytes += metadata.st_size
        else:
            raise RuntimeError(f"refusing cleanup root containing special file: {root}/{relative}")
        entry: dict[str, Any] = {"path": relative, "kind": kind, "mode": oct(stat.S_IMODE(metadata.st_mode))}
        if kind == "file":
            entry["bytes"] = metadata.st_size
        entries.append(entry)

    return {
        "path": str(root),
        "marker": identity(marker),
        "entry_count": len(entries),
        "file_bytes": file_bytes,
        "entries": entries,
    }


def check_record(path: Path, expected: dict[str, Any]) -> dict[str, Any]:
    actual = identity(path)
    for field in ("bytes", "mode", "sha256"):
        if field in expected and actual[field] != expected[field]:
            raise RuntimeError(f"identity mismatch for {path}: {field}")
    return actual


def candidate_and_lock_checks() -> dict[str, Any]:
    freeze = read_object(FREEZE)
    if freeze.get("base_commit") != "f790c2228cbd9e89c2892ba3ffaeb49802fac43b":
        raise RuntimeError("candidate freeze base mismatch")
    freeze_files = {entry["path"]: entry for entry in freeze.get("files", [])}
    candidate_paths = (
        "crates/litchi-ooxml-common/src/mce/codec.rs",
        "crates/litchi-ooxml-common/src/mce/tests.rs",
    )
    if set(freeze_files) != set(candidate_paths):
        raise RuntimeError("candidate freeze path set mismatch")
    root_candidates: dict[str, Any] = {}
    source_candidates: dict[str, Any] = {}
    for name in candidate_paths:
        root_path = ROOT / name
        source_path = SOURCE / name
        root_candidates[name] = check_record(root_path, freeze_files[name])
        source_candidates[name] = check_record(source_path, freeze_files[name])

    baseline = read_object(BASELINE_RECEIPT)
    candidate = read_object(CANDIDATE_RECEIPT)
    lock = read_object(ROOT_LOCK_RECEIPT)
    baseline_binary = check_record(BASELINE / "xml_attribute_bounds", baseline["binary"])
    target_binary = check_record(TARGET / "debug/xml_attribute_bounds", candidate["candidate_binary"])
    lock_file = check_record(ROOT_LOCK / "Cargo.lock", lock["custody"])
    return {
        "candidate_freeze": receipt_descriptor(FREEZE),
        "root_candidates": root_candidates,
        "source_candidates": source_candidates,
        "baseline_binary": baseline_binary,
        "target_binary": target_binary,
        "root_lock": lock_file,
    }


def binary_copy_checks() -> dict[str, Any]:
    binary = require_pass(BINARY_RECEIPT)
    before = binary["before"]
    after = binary["after"]
    if not before["same_content_size_mode"]["all"] or not after["same_content_size_mode"]["all"]:
        raise RuntimeError("retained correctness binary receipt does not certify matching copies")
    before_copy = check_record(CORRECTNESS_BEFORE, before["copy"])
    after_copy = check_record(CORRECTNESS_AFTER, after["copy"])
    if before_copy["sha256"] == after_copy["sha256"]:
        raise RuntimeError("retained before/after correctness binaries are not distinct")
    return {
        "receipt": receipt_descriptor(BINARY_RECEIPT),
        "before_copy": before_copy,
        "after_copy": after_copy,
    }


def retained_artifacts() -> list[dict[str, Any]]:
    """Capture every packet file except native/allocator copies and cleanup outputs."""
    excluded_roots = {
        (PACKET / "binaries/native").resolve(),
        (PACKET / "binaries/allocator").resolve(),
    }
    excluded_files = {PRE_REMOVAL.resolve(), FINAL.resolve()}
    records: list[dict[str, Any]] = []
    for path in sorted(PACKET.rglob("*"), key=lambda candidate: candidate.relative_to(PACKET).as_posix()):
        if not path.is_file():
            if path.is_symlink():
                raise RuntimeError(f"unexpected packet symlink: {path}")
            continue
        resolved = path.resolve()
        if resolved in excluded_files:
            continue
        if any(root == resolved or root in resolved.parents for root in excluded_roots):
            continue
        records.append(receipt_descriptor(path))
    return records


def verify_retained_artifacts(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    checked: list[dict[str, Any]] = []
    for record in records:
        path = PACKET / record["path"]
        checked.append(check_record(path, record))
    return checked


def owned_root_descriptors() -> list[dict[str, Any]]:
    return [
        {"name": name, "path": str(path), "marker": marker_name, "inventory": tree_inventory(path, marker_name, marker_text)}
        for name, path, marker_name, marker_text in OWNED
    ]


def ensure_forbidden_disjoint() -> list[str]:
    owned_paths = {path.resolve() for _, path, _, _ in OWNED}
    forbidden: list[str] = []
    for path in FORBIDDEN:
        resolved = path.resolve(strict=False)
        if any(root == resolved or root in resolved.parents or resolved in root.parents for root in owned_paths):
            raise RuntimeError(f"forbidden path overlaps cleanup scope: {path}")
        forbidden.append(str(path))
    return forbidden


def remove_owned(roots: list[dict[str, Any]]) -> list[str]:
    removed: list[str] = []
    for record in roots:
        path = Path(record["inventory"]["path"])
        if path.is_symlink() or not path.is_dir():
            raise RuntimeError(f"owned root changed before removal: {path}")
        shutil.rmtree(path)
        if path.exists() or path.is_symlink():
            raise RuntimeError(f"owned root remains after removal: {path}")
        removed.append(str(path))
    return removed


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--remove", action="store_true", help="authorize deletion after all custody checks pass")
    args = parser.parse_args()
    if not args.remove:
        raise SystemExit("refusing cleanup without explicit --remove authorization")
    if PRE_REMOVAL.exists() or FINAL.exists():
        raise SystemExit("cleanup receipts already exist; refusing to restart")

    review = require_pass(REVIEW)
    required = {
        "baseline": require_pass(BASELINE_RECEIPT),
        "baseline_postcheck": require_pass(BASELINE_POSTCHECK),
        "candidate": require_pass(CANDIDATE_RECEIPT),
        "candidate_postcheck": require_pass(CANDIDATE_POSTCHECK),
        "continuation": require_pass(CONTINUATION_RECEIPT),
    }
    forbidden = ensure_forbidden_disjoint()
    roots = owned_root_descriptors()
    critical = candidate_and_lock_checks()
    binary_copies = binary_copy_checks()
    retained = retained_artifacts()

    pre_result: dict[str, Any] = {
        "schema": "litchi.mce-0890-correctness-cleanup-pre-removal.v1",
        "status": "prepared",
        "purpose": "custody_inventory_before_owned_correctness_cleanup",
        "performance_claim": False,
        "authorized_by_flag": True,
        "started_utc": utc_now(),
        "correctness_review": receipt_descriptor(REVIEW),
        "required_receipts": {name: receipt_descriptor(path) for name, path in (
            ("baseline", BASELINE_RECEIPT),
            ("baseline_postcheck", BASELINE_POSTCHECK),
            ("candidate", CANDIDATE_RECEIPT),
            ("candidate_postcheck", CANDIDATE_POSTCHECK),
            ("continuation", CONTINUATION_RECEIPT),
        )},
        "owned_roots": roots,
        "critical_hashes": critical,
        "retained_correctness_binaries": binary_copies,
        "retained_artifacts": retained,
        "protected_paths_not_in_cleanup_scope": forbidden,
        "remove_order": [record["name"] for record in roots],
    }
    write_exclusive(PRE_REMOVAL, pre_result)

    result: dict[str, Any] = {
        "schema": "litchi.mce-0890-correctness-cleanup.v1",
        "status": "running",
        "purpose": "remove_only_owned_correctness_scratch",
        "performance_claim": False,
        "pre_removal_receipt": receipt_descriptor(PRE_REMOVAL),
        "correctness_review": receipt_descriptor(REVIEW),
        "protected_paths_not_in_cleanup_scope": forbidden,
        "removed": [],
    }
    try:
        result["removed"] = remove_owned(roots)
        remaining = [str(path) for _, path, _, _ in OWNED if path.exists() or path.is_symlink()]
        if remaining:
            raise RuntimeError(f"owned roots remain after cleanup: {remaining}")
        result["retained_artifacts"] = verify_retained_artifacts(retained)
        result["retained_correctness_binaries"] = binary_copy_checks()
        result["status"] = "pass"
    except Exception as error:
        result["status"] = "fail"
        result["error"] = repr(error)
    result["finished_utc"] = utc_now()
    write_atomic(FINAL, result)
    if result["status"] != "pass":
        raise SystemExit(result["error"])
    print(json.dumps({"status": result["status"], "removed": result["removed"], "receipt": str(FINAL)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
