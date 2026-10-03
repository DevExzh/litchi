#!/usr/bin/env python3
"""Remove only the owned 0890 release scratch after capture and review.

The command requires ``--remove`` and passing build, verification, and
adoption-decision receipts.  It validates the marked native source/build
roots and all four managed executable copies, writes a create-only
pre-removal inventory, removes exactly those three roots, verifies retained
evidence and binaries, and atomically writes the final result.  Candidate
binaries, the retained source archive, correctness custody, and all packet
evidence remain outside the removal scope.
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

SOURCE = Path("/tmp/litchi-0890-source")
TARGET = ROOT / "target/litchi-0890-build"
EXECUTABLES = PACKET / "executables"
MARKER_NAME = ".owner-0890"
MARKER_TEXT = "MCE candidate 0890\n"

BUILD = PACKET / "build.json"
RELEASE_COMPLETION = PACKET / "release-build-completion.json"
EXECUTION_FILES = PACKET / "execution-files.json"
VERIFICATION = PACKET / "verification.json"
DECISION_CANDIDATES = (PACKET / "decision.json", PACKET / "adoption-decision.json")
CORRECTNESS_REVIEW = PACKET / "correctness-review.json"
METHODOLOGY_REVIEW = PACKET / "methodology-review-v3.json"

NATIVE_BUILD = PACKET / "native-build.json"
ALLOCATOR_BUILD = PACKET / "allocator-build.json"
NATIVE_LOG = PACKET / "native-build.log"
ALLOCATOR_LOG = PACKET / "allocator-build.log"
NATIVE_DEPFILE = PACKET / "native.d"
ALLOCATOR_DEPFILE = PACKET / "allocator.d"

BASELINE_PACKET = ROOT / "docs/performance/results/change-0885"
RETAINED_EXTERNAL = (
    BASELINE_PACKET / "source.tar.gz",
    BASELINE_PACKET / "source.json",
    BASELINE_PACKET / "binaries/native",
    BASELINE_PACKET / "binaries/allocator",
)

PRE_REMOVAL = PACKET / "cleanup-release-pre-removal.json"
FINAL = PACKET / "cleanup-release.json"

FORBIDDEN = (
    Path("/tmp/litchi-0890-correctness-source"),
    ROOT / "target/litchi-0890-correctness",
    Path("/tmp/litchi-0890-correctness-baseline"),
    Path("/tmp/litchi-0890-correctness-root-lock"),
    PACKET / "binaries/correctness-before",
    PACKET / "binaries/correctness-after",
)

ROLES = ("native", "allocator")
SIDES = ("before", "after")
MANAGED_NAMES = {
    ("before", "native"): "before-native",
    ("before", "allocator"): "before-allocator",
    ("after", "native"): "after-native",
    ("after", "allocator"): "after-allocator",
}
PACKET_BINARIES = {
    "native": PACKET / "binaries/native",
    "allocator": PACKET / "binaries/allocator",
}
TARGET_BINARIES = {
    "native": TARGET / "release/litchi-perf-baseline",
    "allocator": TARGET / "release/litchi-perf-baseline-alloc",
}
BASELINE_BINARIES = {
    "native": BASELINE_PACKET / "binaries/native",
    "allocator": BASELINE_PACKET / "binaries/allocator",
}


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


def root_descriptor(path: Path) -> dict[str, Any]:
    return identity(path, relative_to=ROOT)


def require_status(path: Path, statuses: set[str]) -> dict[str, Any]:
    value = read_object(path)
    if value.get("status") not in statuses:
        raise RuntimeError(f"receipt has unacceptable status: {path}")
    return value


def normalized_mode(value: Any) -> str | None:
    if isinstance(value, int):
        return oct(value)
    if isinstance(value, str):
        if value.startswith("0o"):
            return value
        if value.isdigit():
            return oct(int(value))
    return None


def check_record(path: Path, expected: dict[str, Any]) -> dict[str, Any]:
    actual = identity(path)
    for field in ("bytes", "sha256"):
        if field in expected and actual[field] != expected[field]:
            raise RuntimeError(f"identity mismatch for {path}: {field}")
    if "mode" in expected and normalized_mode(expected["mode"]) != actual["mode"]:
        raise RuntimeError(f"identity mismatch for {path}: mode")
    return actual


def tree_inventory(root: Path, *, marker: bool) -> dict[str, Any]:
    if root.is_symlink() or not root.is_dir():
        raise RuntimeError(f"owned release root is missing or symlinked: {root}")
    if marker:
        marker_path = root / MARKER_NAME
        if marker_path.is_symlink() or not marker_path.is_file() or marker_path.read_text() != MARKER_TEXT:
            raise RuntimeError(f"release ownership marker mismatch: {root / MARKER_NAME}")

    entries: list[dict[str, Any]] = []
    file_bytes = 0
    for path in sorted(root.rglob("*"), key=lambda candidate: candidate.relative_to(root).as_posix()):
        metadata = path.lstat()
        relative = path.relative_to(root).as_posix()
        if stat.S_ISLNK(metadata.st_mode):
            raise RuntimeError(f"refusing root containing symlink: {root}/{relative}")
        if stat.S_ISDIR(metadata.st_mode):
            kind = "directory"
        elif stat.S_ISREG(metadata.st_mode):
            kind = "file"
            file_bytes += metadata.st_size
        else:
            raise RuntimeError(f"refusing root containing special file: {root}/{relative}")
        entry: dict[str, Any] = {"path": relative, "kind": kind, "mode": oct(stat.S_IMODE(metadata.st_mode))}
        if kind == "file":
            entry["bytes"] = metadata.st_size
        entries.append(entry)
    return {
        "path": str(root),
        "marker": identity(root / MARKER_NAME) if marker else None,
        "entry_count": len(entries),
        "file_bytes": file_bytes,
        "entries": entries,
    }


def ensure_forbidden_disjoint() -> list[str]:
    removed = {SOURCE.resolve(), TARGET.resolve(), EXECUTABLES.resolve()}
    result: list[str] = []
    for path in FORBIDDEN:
        resolved = path.resolve(strict=False)
        if any(root == resolved or root in resolved.parents or resolved in root.parents for root in removed):
            raise RuntimeError(f"protected path overlaps release cleanup scope: {path}")
        result.append(str(path))
    return result


def build_checks() -> dict[str, Any]:
    build = require_status(BUILD, {"pass"})
    if build.get("schema") != "litchi.mce-candidate-build-0890.v1":
        raise RuntimeError("candidate build schema mismatch")
    if build.get("source_root") != str(SOURCE) or build.get("target_root") != str(TARGET):
        raise RuntimeError("candidate build roots do not match owned release roots")
    freeze = PACKET / "candidate-freeze-v2.json"
    if not isinstance(build.get("source"), dict):
        raise RuntimeError("candidate build does not bind its source freeze")
    check_record(freeze, build["source"])
    completion = require_status(RELEASE_COMPLETION, {"pass"})
    if completion.get("exit_code") != 0:
        raise RuntimeError("release build completion is not zero-exit")
    if not isinstance(completion.get("build"), dict):
        raise RuntimeError("release completion does not bind build.json")
    check_record(BUILD, completion["build"])

    role_receipts: dict[str, Any] = {}
    for role, receipt_path, log_path, depfile_path in (
        ("native", NATIVE_BUILD, NATIVE_LOG, NATIVE_DEPFILE),
        ("allocator", ALLOCATOR_BUILD, ALLOCATOR_LOG, ALLOCATOR_DEPFILE),
    ):
        receipt = read_object(receipt_path)
        if receipt.get("exit_code") != 0:
            raise RuntimeError(f"{role} build receipt is not zero-exit")
        if not isinstance(receipt.get("argv"), list) or "--release" not in receipt["argv"]:
            raise RuntimeError(f"{role} build receipt does not bind a release command")
        completion_receipts = completion.get("receipts", {})
        for name, path in ((f"{role}-build.json", receipt_path), (f"{role}-build.log", log_path)):
            if isinstance(completion_receipts.get(name), dict):
                check_record(path, completion_receipts[name])
        role_receipts[role] = {
            "receipt": receipt_descriptor(receipt_path),
            "log": receipt_descriptor(log_path),
            "depfile": receipt_descriptor(depfile_path),
        }

    binary_checks: dict[str, Any] = {}
    for role in ROLES:
        item = build.get("binaries", {}).get(role)
        if not isinstance(item, dict) or item.get("executable") is not True:
            raise RuntimeError(f"{role} build binary binding is incomplete")
        packet_binary = check_record(PACKET_BINARIES[role], item)
        receipt_binding = item.get("build_receipt")
        if isinstance(receipt_binding, dict):
            check_record(PACKET / f"{role}-build.json", receipt_binding)
        depfile_binding = item.get("depfile")
        if isinstance(depfile_binding, dict):
            check_record(PACKET / f"{role}.d", depfile_binding)
        if not os.access(PACKET_BINARIES[role], os.X_OK):
            raise RuntimeError(f"{role} retained packet binary is not executable")
        target_binary = check_record(TARGET_BINARIES[role], item)
        if target_binary["sha256"] != packet_binary["sha256"] or target_binary["mode"] != packet_binary["mode"]:
            raise RuntimeError(f"{role} target and retained packet binaries differ")
        binary_checks[role] = {"packet": packet_binary, "target": target_binary}

    return {
        "build": receipt_descriptor(BUILD),
        "completion": receipt_descriptor(RELEASE_COMPLETION),
        "roles": role_receipts,
        "binaries": binary_checks,
    }


def managed_copy_checks() -> dict[str, Any]:
    execution = require_status(EXECUTION_FILES, {"pass"})
    inventory = tree_inventory(EXECUTABLES, marker=False)
    top_level = [path for path in EXECUTABLES.iterdir()]
    expected_names = set(MANAGED_NAMES.values())
    if {path.name for path in top_level} != expected_names or not all(path.is_file() and not path.is_symlink() for path in top_level):
        raise RuntimeError("managed executable directory contains unexpected entries")

    checked: dict[str, Any] = {}
    records = execution.get("records")
    if not isinstance(records, dict):
        raise RuntimeError("managed executable records are missing")
    for side in SIDES:
        checked[side] = {}
        for role in ROLES:
            entry = records.get(side, {}).get(role)
            if not isinstance(entry, dict) or not isinstance(entry.get("copy"), dict) or not isinstance(entry.get("source"), dict):
                raise RuntimeError(f"managed executable record is incomplete: {side}/{role}")
            copy = entry["copy"]
            copy_path = Path(copy.get("path", ""))
            expected_copy = EXECUTABLES / MANAGED_NAMES[(side, role)]
            if copy_path.resolve() != expected_copy.resolve():
                raise RuntimeError(f"managed copy path mismatch: {side}/{role}")
            copy_actual = check_record(expected_copy, copy)
            if not os.access(expected_copy, os.X_OK):
                raise RuntimeError(f"managed copy is not executable: {side}/{role}")
            if "device" in copy and "inode" in copy:
                metadata = expected_copy.stat()
                if metadata.st_dev != copy["device"] or metadata.st_ino != copy["inode"]:
                    raise RuntimeError(f"managed copy inode changed: {side}/{role}")

            source_path = Path(entry["source"].get("path", ""))
            expected_source = (BASELINE_BINARIES if side == "before" else PACKET_BINARIES)[role]
            if source_path.resolve() != expected_source.resolve():
                raise RuntimeError(f"managed source path mismatch: {side}/{role}")
            source_actual = check_record(expected_source, entry["source"])
            checked[side][role] = {"source": source_actual, "copy": copy_actual}
    return {"receipt": receipt_descriptor(EXECUTION_FILES), "inventory": inventory, "records": checked}


def review_checks() -> dict[str, Any]:
    verification = require_status(VERIFICATION, {"pass"})
    correctness = require_status(CORRECTNESS_REVIEW, {"pass"})
    methodology = require_status(METHODOLOGY_REVIEW, {"pass"})
    decision_path = next((path for path in DECISION_CANDIDATES if path.is_file()), None)
    if decision_path is None:
        raise RuntimeError("adoption decision receipt is missing")
    decision = read_object(decision_path)
    if decision.get("status") not in {"complete", "pass"}:
        raise RuntimeError("adoption decision receipt is not complete")
    if "decision" not in decision and "adopted_optimization" not in decision:
        raise RuntimeError("adoption decision field is missing")
    return {
        "verification": receipt_descriptor(VERIFICATION),
        "adoption_decision": receipt_descriptor(decision_path),
        "correctness_review": receipt_descriptor(CORRECTNESS_REVIEW),
        "methodology_review": receipt_descriptor(METHODOLOGY_REVIEW),
    }


def retained_artifacts() -> list[dict[str, Any]]:
    """Capture packet evidence except the managed executable directory."""
    excluded_root = EXECUTABLES.resolve()
    excluded_files = {PRE_REMOVAL.resolve(), FINAL.resolve()}
    records: list[dict[str, Any]] = []
    for path in sorted(PACKET.rglob("*"), key=lambda candidate: candidate.relative_to(PACKET).as_posix()):
        if path.is_symlink():
            if excluded_root in path.resolve(strict=False).parents:
                continue
            raise RuntimeError(f"unexpected packet symlink: {path}")
        if not path.is_file():
            continue
        resolved = path.resolve()
        if resolved in excluded_files or resolved == excluded_root or excluded_root in resolved.parents:
            continue
        records.append(receipt_descriptor(path))
    records.extend(root_descriptor(path) for path in RETAINED_EXTERNAL)
    return records


def verify_retained_artifacts(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    checked: list[dict[str, Any]] = []
    for record in records:
        relative = Path(record["path"])
        path = ROOT / relative if not relative.is_absolute() else relative
        checked.append(check_record(path, record))
    return checked


def remove_owned() -> list[str]:
    removed: list[str] = []
    for path in (SOURCE, TARGET, EXECUTABLES):
        if path.is_symlink() or not path.is_dir():
            raise RuntimeError(f"owned release root changed before removal: {path}")
        shutil.rmtree(path)
        if path.exists() or path.is_symlink():
            raise RuntimeError(f"owned release root remains after removal: {path}")
        removed.append(str(path))
    return removed


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--remove", action="store_true", help="authorize deletion after all custody checks pass")
    args = parser.parse_args()
    if not args.remove:
        raise SystemExit("refusing release cleanup without explicit --remove authorization")
    if PRE_REMOVAL.exists() or FINAL.exists():
        raise SystemExit("release cleanup receipts already exist; refusing to restart")

    forbidden = ensure_forbidden_disjoint()
    roots = {
        "source": tree_inventory(SOURCE, marker=True),
        "target": tree_inventory(TARGET, marker=True),
        "managed_executables": managed_copy_checks(),
    }
    build = build_checks()
    reviews = review_checks()
    retained = retained_artifacts()

    pre_result: dict[str, Any] = {
        "schema": "litchi.mce-0890-release-cleanup-pre-removal.v1",
        "status": "prepared",
        "purpose": "custody_inventory_before_owned_release_cleanup",
        "performance_claim": False,
        "authorized_by_flag": True,
        "started_utc": utc_now(),
        "owned_roots": roots,
        "build_custody": build,
        "review_custody": reviews,
        "retained_artifacts": retained,
        "protected_paths_not_in_cleanup_scope": forbidden,
        "remove_order": [str(SOURCE), str(TARGET), str(EXECUTABLES)],
    }
    write_exclusive(PRE_REMOVAL, pre_result)

    result: dict[str, Any] = {
        "schema": "litchi.mce-0890-release-cleanup.v1",
        "status": "running",
        "purpose": "remove_only_owned_release_scratch",
        "performance_claim": False,
        "pre_removal_receipt": receipt_descriptor(PRE_REMOVAL),
        "protected_paths_not_in_cleanup_scope": forbidden,
        "removed": [],
    }
    try:
        result["removed"] = remove_owned()
        remaining = [str(path) for path in (SOURCE, TARGET, EXECUTABLES) if path.exists() or path.is_symlink()]
        if remaining:
            raise RuntimeError(f"owned release roots remain after cleanup: {remaining}")
        result["retained_artifacts"] = verify_retained_artifacts(retained)
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
