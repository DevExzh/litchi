#!/usr/bin/env python3
"""Remove only the marked 0891 quality/release scratch roots.

The command is deliberately create-only and requires ``--remove``.  It
records an absolute-path pre-removal inventory, removes exactly the owned
source, build, quality-target, and managed-executable roots, then creates one
final receipt.  It never touches packet evidence or failed attempts.
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
SOURCE = Path("/tmp/litchi-0891-source")
BUILD_TARGET = ROOT / "target/litchi-0891-build"
QUALITY_TARGET = ROOT / "target/litchi-0891-quality"
EXECUTABLES = PACKET / "executables"
PRE = PACKET / "cleanup-pre-removal.json"
FINAL = PACKET / "cleanup.json"

MARKERS = {
    SOURCE: (".owner-0891", "x14ac candidate 0891\n"),
    BUILD_TARGET: (".owner-0891", "x14ac candidate 0891 target\n"),
    QUALITY_TARGET: (".owner-0891-quality", "litchi 0891 XLSX quality target\n"),
}
MANAGED_NAMES = ("before-native", "before-allocator", "after-native", "after-allocator")

REQUIRED = {
    "build": PACKET / "build.json",
    "quality": PACKET / "quality.json",
    "execution": PACKET / "execution-files.json",
    "plan": PACKET / "plan.json",
    "verification": PACKET / "verification.json",
    "capture_review": PACKET / "capture-review.json",
}


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"expected regular non-symlink file: {path}")
    metadata = path.stat()
    return {
        "path": str(path),
        "bytes": metadata.st_size,
        "mode": stat.S_IMODE(metadata.st_mode),
        "sha256": digest(path),
    }


def read_json(path: Path) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"missing or symlinked receipt: {path}")
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise RuntimeError(f"receipt is not an object: {path}")
    return value


def write_once(path: Path, value: dict[str, Any]) -> None:
    if path.exists() or path.is_symlink():
        raise RuntimeError(f"create-only receipt already exists: {path}")
    with path.open("x", encoding="utf-8") as destination:
        json.dump(value, destination, indent=2, sort_keys=True)
        destination.write("\n")
        destination.flush()
        os.fsync(destination.fileno())
    directory = os.open(path.parent, os.O_RDONLY | getattr(os, "O_DIRECTORY", 0))
    try:
        os.fsync(directory)
    finally:
        os.close(directory)


def check_receipt(path: Path, status: str) -> tuple[dict[str, Any], dict[str, Any]]:
    value = read_json(path)
    if value.get("status") != status:
        raise RuntimeError(f"receipt status is not {status}: {path}")
    return value, descriptor(path)


def marker_inventory(root: Path, marker_name: str, marker_text: str) -> dict[str, Any]:
    root = root.resolve()
    if root.is_symlink() or not root.is_dir():
        raise RuntimeError(f"owned root is missing or symlinked: {root}")
    marker = root / marker_name
    if marker.is_symlink() or not marker.is_file() or marker.read_text(encoding="utf-8") != marker_text:
        raise RuntimeError(f"ownership marker mismatch: {marker}")
    entries: list[dict[str, Any]] = []
    file_bytes = 0
    for path in sorted(root.rglob("*"), key=lambda item: str(item.relative_to(root))):
        metadata = path.lstat()
        if stat.S_ISLNK(metadata.st_mode):
            raise RuntimeError(f"owned root contains symlink: {path}")
        if stat.S_ISDIR(metadata.st_mode):
            kind = "directory"
        elif stat.S_ISREG(metadata.st_mode):
            kind = "file"
            file_bytes += metadata.st_size
        else:
            raise RuntimeError(f"owned root contains special file: {path}")
        item: dict[str, Any] = {"path": str(path.resolve()), "kind": kind, "mode": stat.S_IMODE(metadata.st_mode)}
        if kind == "file":
            item["bytes"] = metadata.st_size
        entries.append(item)
    return {"path": str(root), "marker": descriptor(marker), "entry_count": len(entries), "file_bytes": file_bytes, "entries": entries}


def managed_inventory(execution: dict[str, Any]) -> dict[str, Any]:
    if EXECUTABLES.is_symlink() or not EXECUTABLES.is_dir():
        raise RuntimeError(f"managed executable root is missing or symlinked: {EXECUTABLES}")
    names = sorted(path.name for path in EXECUTABLES.iterdir())
    if names != sorted(MANAGED_NAMES):
        raise RuntimeError(f"managed executable names differ: {names}")
    records: dict[str, Any] = {}
    raw_records = execution.get("records")
    if not isinstance(raw_records, dict):
        raise RuntimeError("execution receipt has no managed records")
    for role in ("before", "after"):
        for mode in ("native", "allocator"):
            name = f"{role}-{mode}"
            raw = raw_records.get(role, {}).get(mode, {})
            copy = raw.get("copy") if isinstance(raw, dict) else None
            if not isinstance(copy, dict):
                raise RuntimeError(f"managed copy descriptor is missing: {name}")
            path = Path(str(copy.get("path", ""))).resolve()
            expected = (EXECUTABLES / name).resolve()
            if path != expected:
                raise RuntimeError(f"managed copy path mismatch: {name}")
            actual = descriptor(path)
            for field in ("bytes", "mode", "sha256"):
                if copy.get(field) != actual[field]:
                    raise RuntimeError(f"managed copy {field} changed: {name}")
            if not (actual["mode"] & 0o111):
                raise RuntimeError(f"managed copy is not executable: {name}")
            records[name] = actual
    return {"path": str(EXECUTABLES.resolve()), "files": records}


def retained_receipts() -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for label, path in REQUIRED.items():
        result[label] = descriptor(path)
    for label, path in {
        "candidate_freeze": PACKET / "candidate-freeze-v3.json",
        "candidate_patch": PACKET / "candidate-v3.patch",
        "source_review": PACKET / "source-review-v3.json",
        "fixture_census": PACKET / "fixture-census.json",
        "quality_runner": PACKET / "quality_runner.py",
        "build_helper": PACKET / "build.py",
        "managed_copy_helper": PACKET / "prepare_exec.py",
    }.items():
        result[label] = descriptor(path)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--remove", action="store_true", help="authorize removal after custody checks")
    args = parser.parse_args()
    if PRE.exists() or FINAL.exists():
        raise SystemExit("cleanup receipts already exist; refusing to restart")

    build, build_descriptor = check_receipt(REQUIRED["build"], "pass")
    quality, quality_descriptor = check_receipt(REQUIRED["quality"], "pass")
    execution, execution_descriptor = check_receipt(REQUIRED["execution"], "pass")
    plan, plan_descriptor = check_receipt(REQUIRED["plan"], "frozen")
    verification, verification_descriptor = check_receipt(REQUIRED["verification"], "pass")
    capture_review, capture_review_descriptor = check_receipt(REQUIRED["capture_review"], "pass")
    if build.get("source_root") != str(SOURCE) or build.get("target_root") != str(BUILD_TARGET):
        raise RuntimeError("build receipt roots do not match owned release roots")
    if quality.get("quality_receipt", {}).get("path") is None:
        raise RuntimeError("quality aggregate does not bind its executed receipt")
    if plan.get("status") != "frozen" or execution.get("status") != "pass":
        raise RuntimeError("plan or managed executable custody is not complete")

    roots = {str(path.resolve()): marker_inventory(path, *marker) for path, marker in MARKERS.items()}
    managed = managed_inventory(execution)
    pre = {
        "schema": "litchi.xlsx-x14ac-cleanup-pre-removal-0891.v1",
        "status": "prepared",
        "purpose": "owned_0891_quality_release_scratch_custody",
        "performance_claim": False,
        "created_utc": utc_now(),
        "remove_requested": args.remove,
        "owned_roots": [str(path.resolve()) for path in (*MARKERS, EXECUTABLES)],
        "marked_root_inventories": roots,
        "managed_executable_inventory": managed,
        "retained_receipts": retained_receipts(),
        "prerequisite_receipts": {
            "build": build_descriptor,
            "quality": quality_descriptor,
            "execution": execution_descriptor,
            "plan": plan_descriptor,
            "verification": verification_descriptor,
            "capture_review": capture_review_descriptor,
        },
    }
    write_once(PRE, pre)
    if not args.remove:
        print(json.dumps({"status": "prepared", "pre_removal": str(PRE)}))
        return 0

    removed: list[str] = []
    try:
        for path in (*MARKERS, EXECUTABLES):
            if path.is_symlink() or not path.is_dir():
                raise RuntimeError(f"owned root changed before removal: {path}")
            shutil.rmtree(path)
            if path.exists() or path.is_symlink():
                raise RuntimeError(f"owned root remains after removal: {path}")
            removed.append(str(path.resolve()))
        result = {
            "schema": "litchi.xlsx-x14ac-cleanup-0891.v1",
            "status": "pass",
            "purpose": "owned_0891_quality_release_scratch_cleanup",
            "performance_claim": False,
            "finished_utc": utc_now(),
            "pre_removal": descriptor(PRE),
            "removed_roots": removed,
            "retained_receipts": pre["retained_receipts"],
        }
    except Exception as error:
        result = {
            "schema": "litchi.xlsx-x14ac-cleanup-0891.v1",
            "status": "fail",
            "purpose": "owned_0891_quality_release_scratch_cleanup",
            "performance_claim": False,
            "finished_utc": utc_now(),
            "pre_removal": descriptor(PRE),
            "removed_roots": removed,
            "error": repr(error),
        }
        write_once(FINAL, result)
        raise
    write_once(FINAL, result)
    print(json.dumps({"status": result["status"], "removed": removed, "final": str(FINAL)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
