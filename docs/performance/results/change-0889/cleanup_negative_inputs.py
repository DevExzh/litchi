#!/usr/bin/env python3
"""Remove a sealed failed-attempt mutation copy after identity checks."""

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


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def write_new(path: Path, value: dict[str, Any]) -> None:
    with path.open("x") as destination:
        json.dump(value, destination, indent=2, sort_keys=True)
        destination.write("\n")
        destination.flush()
        os.fsync(destination.fileno())


def replace(path: Path, value: dict[str, Any]) -> None:
    temporary = path.with_suffix(".tmp")
    with temporary.open("x") as destination:
        json.dump(value, destination, indent=2, sort_keys=True)
        destination.write("\n")
        destination.flush()
        os.fsync(destination.fileno())
    temporary.replace(path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", required=True, help="failed attempt owning the mutation copy")
    args = parser.parse_args()
    if not args.attempt or args.attempt in {".", ".."} or "/" in args.attempt:
        raise SystemExit("--attempt must be one path component")

    attempt_dir = PACKET / "attempts" / args.attempt
    receipt_path = attempt_dir / "validation-receipt.json"
    cleanup_path = attempt_dir / "negative-input-cleanup.json"
    require(receipt_path.is_file(), f"receipt is missing: {receipt_path}")
    require(not cleanup_path.exists(), f"cleanup receipt already exists: {cleanup_path}")
    receipt = json.loads(receipt_path.read_text())
    mutation = receipt.get("corrupt_member")
    require(isinstance(mutation, dict), "corrupt mutation recipe is missing")
    command = next(
        (item for item in receipt.get("commands", []) if item.get("name") == "negative_corrupt_member"),
        None,
    )
    require(command is not None, "corrupt-member command is missing")
    log_path = PACKET / command["log"]
    raw_json = attempt_dir / "runtime/corrupt.json"
    archive = attempt_dir / "runtime/corrupt-member.xlsx"
    require(log_path.is_file() and sha256(log_path) == command["log_sha256"],
            "corrupt-member log is missing or changed")
    require(raw_json.is_file(), "corrupt-member JSON is missing")
    metadata = archive.lstat()
    require(stat.S_ISREG(metadata.st_mode) and not stat.S_ISLNK(metadata.st_mode),
            "corrupt-member archive is not a regular file")
    require(metadata.st_uid == os.getuid(), "corrupt-member archive owner differs")
    require(metadata.st_size == mutation["mutated"]["bytes"] and
            sha256(archive) == mutation["mutated"]["sha256"],
            "corrupt-member archive identity differs from its recipe")

    unreadable = attempt_dir / "runtime/unreadable/fixture.xlsx"
    unreadable_info: dict[str, Any] | None = None
    if unreadable.exists():
        unreadable_metadata = unreadable.lstat()
        require(
            stat.S_ISREG(unreadable_metadata.st_mode) and not stat.S_ISLNK(unreadable_metadata.st_mode),
            "unreadable archive is not a regular file",
        )
        require(unreadable_metadata.st_uid == os.getuid(), "unreadable archive owner differs")
        input_identity = next(
            (
                item
                for item in receipt.get("negative_inputs", [])
                if item.get("kind") == "unreadable_package"
            ),
            None,
        )
        require(input_identity is not None, "unreadable archive identity is missing")
        require(
            unreadable_metadata.st_size == input_identity["bytes"]
            and input_identity["sha256"] == receipt.get("negative_inputs", [])[0].get("sha256"),
            "unreadable archive size or recipe identity differs",
        )
        source = ROOT / input_identity["source_fixture"]
        source_sha256 = sha256(source)
        require(input_identity["sha256"] == source_sha256, "unreadable recipe differs from source")
        unreadable.chmod(stat.S_IRUSR | stat.S_IWUSR)
        require(sha256(unreadable) == source_sha256, "unreadable archive content differs from source")
        unreadable_info = {
            "path": str(unreadable.relative_to(PACKET)),
            "bytes": unreadable_metadata.st_size,
            "sha256": source_sha256,
            "source_fixture": input_identity["source_fixture"],
            "restored_mode": oct(stat.S_IMODE(unreadable.stat().st_mode)),
        }

    result: dict[str, Any] = {
        "schema": "litchi.mce-0889-negative-input-cleanup.v1",
        "status": "prepared",
        "attempt": str(attempt_dir.relative_to(PACKET)),
        "started_utc": utc_now(),
        "archive": {
            "path": str(archive.relative_to(PACKET)),
            "bytes": metadata.st_size,
            "sha256": mutation["mutated"]["sha256"],
            "member": mutation.get("member"),
            "offset": mutation.get("offset"),
            "original_byte": mutation.get("original_byte"),
            "mutated_byte": mutation.get("mutated_byte"),
            "raw_read_error": mutation.get("raw_read_error"),
        },
        "retained": {
            "log": str(log_path.relative_to(PACKET)),
            "log_sha256": command["log_sha256"],
            "json": str(raw_json.relative_to(PACKET)),
            "json_bytes": raw_json.stat().st_size,
            "json_sha256": sha256(raw_json),
        },
        "unreadable_archive": unreadable_info,
        "removed": False,
    }
    write_new(cleanup_path, result)
    try:
        archive.unlink()
        if unreadable_info is not None:
            unreadable.unlink()
            unreadable.parent.rmdir()
        result["removed"] = not archive.exists()
        if unreadable_info is not None:
            result["removed"] = result["removed"] and not unreadable.exists()
        if not result["removed"]:
            raise RuntimeError("corrupt-member archive remains after cleanup")
        result["status"] = "pass"
    except Exception as error:
        result["status"] = "fail"
        result["error"] = repr(error)
        result["finished_utc"] = utc_now()
        replace(cleanup_path, result)
        raise
    result["finished_utc"] = utc_now()
    replace(cleanup_path, result)
    print(json.dumps({"cleanup": str(cleanup_path), "removed": result["removed"]}))


if __name__ == "__main__":
    main()
