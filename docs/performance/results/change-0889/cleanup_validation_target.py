#!/usr/bin/env python3
"""Remove the owned 0889 Cargo target after a successful validation attempt."""

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
TARGET = ROOT / "target" / "litchi-0889-validation"
MARKER = TARGET / ".owner-0889"
MARKER_TEXT = "litchi 0889 validation target\n"
BINARY = TARGET / "debug" / "xml_attribute_bounds"


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def file_identity(path: Path) -> dict[str, Any]:
    mode = stat.S_IMODE(path.stat().st_mode)
    return {
        "path": str(path.relative_to(ROOT)),
        "bytes": path.stat().st_size,
        "mode": oct(mode),
        "sha256": sha256(path),
    }


def write_new(path: Path, value: dict[str, Any]) -> None:
    with path.open("x") as destination:
        json.dump(value, destination, indent=2, sort_keys=True)
        destination.write("\n")
        destination.flush()
        os.fsync(destination.fileno())


def replace_receipt(path: Path, value: dict[str, Any]) -> None:
    temporary = path.with_suffix(".tmp")
    with temporary.open("x") as destination:
        json.dump(value, destination, indent=2, sort_keys=True)
        destination.write("\n")
        destination.flush()
        os.fsync(destination.fileno())
    temporary.replace(path)


def target_inventory() -> dict[str, Any]:
    if TARGET.is_symlink() or not TARGET.is_dir():
        raise SystemExit(f"owned validation target is missing: {TARGET}")
    if not MARKER.is_file() or MARKER.is_symlink():
        raise SystemExit(f"refusing unmarked validation target: {TARGET}")
    if MARKER.read_text() != MARKER_TEXT:
        raise SystemExit(f"validation target marker does not match: {MARKER}")

    entries: list[dict[str, Any]] = []
    total_bytes = 0
    for path in sorted(TARGET.rglob("*")):
        relative = path.relative_to(TARGET)
        metadata = path.lstat()
        if stat.S_ISLNK(metadata.st_mode):
            raise SystemExit(f"refusing target containing symlink: {relative}")
        entry: dict[str, Any] = {
            "path": str(relative),
            "kind": "directory" if path.is_dir() else "file",
            "mode": oct(stat.S_IMODE(metadata.st_mode)),
        }
        if path.is_file():
            entry["bytes"] = metadata.st_size
            total_bytes += metadata.st_size
        entries.append(entry)

    inventory: dict[str, Any] = {
        "target": str(TARGET.relative_to(ROOT)),
        "marker": file_identity(MARKER),
        "entry_count": len(entries),
        "file_bytes": total_bytes,
        "entries": entries,
    }
    if BINARY.is_file():
        inventory["binary"] = file_identity(BINARY)
    return inventory


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", required=True, help="successful attempt name")
    parser.add_argument(
        "--aggregate",
        default=str(PACKET / "qualification-receipt.json"),
        help="sealed aggregate qualification receipt",
    )
    parser.add_argument(
        "--remove",
        action="store_true",
        help="remove the marked target after writing its pre-cleanup identity",
    )
    args = parser.parse_args()
    if not args.attempt or args.attempt in {".", ".."} or "/" in args.attempt:
        raise SystemExit("--attempt must be one path component")

    attempt_dir = PACKET / "attempts" / args.attempt
    receipt_path = attempt_dir / "validation-receipt.json"
    cleanup_path = attempt_dir / "target-cleanup.json"
    if not receipt_path.is_file():
        raise SystemExit(f"validation receipt is missing: {receipt_path}")
    receipt = json.loads(receipt_path.read_text())
    if receipt.get("status") != "pass":
        raise SystemExit("refusing cleanup because validation status is not pass")
    if receipt.get("target_dir") != str(TARGET):
        raise SystemExit("receipt target does not match the owned validation target")
    if receipt.get("target_marker") != str(MARKER.relative_to(ROOT)):
        raise SystemExit("receipt marker does not match the owned validation target")
    aggregate_path = Path(args.aggregate)
    if not aggregate_path.is_absolute():
        aggregate_path = ROOT / aggregate_path
    aggregate = json.loads(aggregate_path.read_text())
    if aggregate.get("status") != "pass":
        raise SystemExit("refusing cleanup because aggregate qualification is not pass")
    if cleanup_path.exists():
        raise SystemExit(f"cleanup receipt already exists: {cleanup_path}")

    inventory = target_inventory()
    expected_binary = (
        receipt.get("after", {}).get("binary")
        or receipt.get("runtime_binary_after")
        or receipt.get("runtime_binary")
        or receipt.get("binary_after_build")
    )
    if expected_binary is None:
        raise SystemExit("successful receipt has no final binary identity")
    if inventory.get("binary") != expected_binary:
        raise SystemExit("target binary identity differs from the successful receipt")
    orphan_cleanup = aggregate.get("orphan_cleanup")
    if orphan_cleanup is None:
        unreadable_ref = aggregate.get("attempts", {}).get("unreadable_package", {}).get("receipt")
        if unreadable_ref:
            unreadable_receipt = json.loads((PACKET / Path(unreadable_ref)).read_text())
            orphan_cleanup = unreadable_receipt.get("orphan_cleanup")
    result: dict[str, Any] = {
        "schema": "litchi.mce-0889-target-cleanup.v1",
        "attempt": str(attempt_dir.relative_to(PACKET)),
        "aggregate": str(aggregate_path.relative_to(PACKET)),
        "started_utc": utc_now(),
        "remove_requested": args.remove,
        "pre_cleanup": inventory,
        "orphan_cleanup": orphan_cleanup,
        "status": "prepared",
        "target_removed": False,
    }
    write_new(cleanup_path, result)
    if args.remove:
        try:
            shutil.rmtree(TARGET)
            result["target_removed"] = not TARGET.exists() and not TARGET.is_symlink()
            if not result["target_removed"]:
                raise RuntimeError(f"target still exists after cleanup: {TARGET}")
            result["status"] = "pass"
        except Exception as error:
            result["status"] = "fail"
            result["error"] = repr(error)
            result["finished_utc"] = utc_now()
            replace_receipt(cleanup_path, result)
            raise
    else:
        result["status"] = "dry_run"
    result["finished_utc"] = utc_now()
    replace_receipt(cleanup_path, result)
    print(json.dumps({"cleanup": str(cleanup_path), "removed": result["target_removed"]}))


if __name__ == "__main__":
    main()
