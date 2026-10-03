#!/usr/bin/env python3
"""Verify release-cleanup custody after the original cleanup has finished.

This verifier is read-only apart from its create-only receipt.  The original
cleanup helper recorded packet paths relative to the 0890 packet and the four
retained baseline paths relative to the repository root; those namespaces are
resolved explicitly here.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import stat
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
PRE = PACKET / "cleanup-release-pre-removal.json"
FINAL = PACKET / "cleanup-release.json"
RECEIPT = PACKET / "cleanup-release-custody-verification.json"
ORIGINAL_HELPER = PACKET / "cleanup_release.py"

REMOVED_ROOTS = (
    Path("/tmp/litchi-0890-source"),
    ROOT / "target/litchi-0890-build",
    PACKET / "executables",
)

EXTERNAL = {
    "docs/performance/results/change-0885/source.tar.gz": ROOT / "docs/performance/results/change-0885/source.tar.gz",
    "docs/performance/results/change-0885/source.json": ROOT / "docs/performance/results/change-0885/source.json",
    "docs/performance/results/change-0885/binaries/native": ROOT / "docs/performance/results/change-0885/binaries/native",
    "docs/performance/results/change-0885/binaries/allocator": ROOT / "docs/performance/results/change-0885/binaries/allocator",
}


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def identity(path: Path) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"expected regular retained file without symlink: {path}")
    metadata = path.stat()
    return {
        "bytes": metadata.st_size,
        "mode": oct(stat.S_IMODE(metadata.st_mode)),
        "sha256": sha256(path),
    }


def descriptor(path: Path, *, relative_to: Path) -> dict[str, Any]:
    result = identity(path)
    result["path"] = str(path.relative_to(relative_to))
    return result


def read_object(path: Path) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"required receipt is missing or symlinked: {path}")
    value = json.loads(path.read_text())
    if not isinstance(value, dict):
        raise RuntimeError(f"receipt is not a JSON object: {path}")
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


def check(path: Path, expected: dict[str, Any]) -> dict[str, Any]:
    actual = identity(path)
    for field in ("bytes", "sha256"):
        if actual[field] != expected.get(field):
            raise RuntimeError(f"identity mismatch for {path}: {field}")
    if normalized_mode(expected.get("mode")) != actual["mode"]:
        raise RuntimeError(f"identity mismatch for {path}: mode")
    return actual


def packet_path(relative: str) -> Path:
    value = Path(relative)
    if value.is_absolute() or ".." in value.parts:
        raise RuntimeError(f"unsafe packet-relative retained path: {relative}")
    path = PACKET / value
    if path.resolve(strict=False).relative_to(PACKET.resolve()) is None:  # pragma: no cover
        raise RuntimeError(f"packet path escaped packet root: {relative}")
    return path


def retained_check(records: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
    checked: list[dict[str, Any]] = []
    by_path: dict[str, dict[str, Any]] = {}
    for record in records:
        relative = record.get("path")
        if not isinstance(relative, str):
            raise RuntimeError("retained descriptor path is not a string")
        if relative in EXTERNAL:
            namespace = "root"
            path = EXTERNAL[relative]
            relative_to = ROOT
        else:
            namespace = "packet"
            path = packet_path(relative)
            relative_to = PACKET
        key = f"{namespace}:{relative}"
        if key in by_path:
            raise RuntimeError(f"duplicate retained descriptor: {key}")
        actual = check(path, record)
        verified = {
            "namespace": namespace,
            "descriptor": record,
            "resolved_path": str(path),
            "actual": descriptor(path, relative_to=relative_to),
        }
        checked.append(verified)
        by_path[key] = verified
    if set(EXTERNAL) != {record["path"] for record in records if record.get("path") in EXTERNAL}:
        raise RuntimeError("retained external baseline set is incomplete or duplicated")
    return checked, by_path


def write_exclusive(value: dict[str, Any]) -> None:
    if RECEIPT.exists() or RECEIPT.is_symlink():
        raise SystemExit(f"refusing to overwrite existing receipt: {RECEIPT}")
    with RECEIPT.open("x") as destination:
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
    pre = read_object(PRE)
    if pre.get("status") != "prepared":
        raise RuntimeError("pre-removal receipt is not prepared")
    if pre.get("remove_order") != [str(path) for path in REMOVED_ROOTS]:
        raise RuntimeError("pre-removal remove order mismatch")

    final = read_object(FINAL)
    if final.get("status") != "fail":
        raise RuntimeError("original cleanup result is not the preserved failed result")
    if final.get("removed") != [str(path) for path in REMOVED_ROOTS]:
        raise RuntimeError("original cleanup removed-root binding mismatch")
    expected_error = "RuntimeError('identity mismatch for /home/zhuhe/code/litchi/.gitattributes: bytes')"
    if final.get("error") != expected_error:
        raise RuntimeError("original cleanup failure binding mismatch")

    pre_binding = final.get("pre_removal_receipt")
    if not isinstance(pre_binding, dict):
        raise RuntimeError("failed cleanup receipt does not bind its pre-removal receipt")
    pre_actual = check(PRE, pre_binding)
    helper_record = next((record for record in pre["retained_artifacts"] if record.get("path") == "cleanup_release.py"), None)
    if not isinstance(helper_record, dict):
        raise RuntimeError("pre-removal inventory does not retain the original helper")
    helper_actual = check(ORIGINAL_HELPER, helper_record)

    absent: list[str] = []
    for path in REMOVED_ROOTS:
        if path.exists() or path.is_symlink():
            raise RuntimeError(f"owned release root remains after cleanup: {path}")
        absent.append(str(path))

    verified, by_path = retained_check(pre["retained_artifacts"])
    binary_paths = {
        "candidate_native": "binaries/native",
        "candidate_allocator": "binaries/allocator",
        "baseline_native": "docs/performance/results/change-0885/binaries/native",
        "baseline_allocator": "docs/performance/results/change-0885/binaries/allocator",
    }
    binaries: dict[str, dict[str, Any]] = {}
    for label, relative in binary_paths.items():
        namespace = "root" if relative in EXTERNAL else "packet"
        key = f"{namespace}:{relative}"
        if key not in by_path:
            raise RuntimeError(f"release binary missing from retained inventory: {relative}")
        binaries[label] = by_path[key]

    result = {
        "schema": "litchi.mce-0890-release-cleanup-custody-verification.v1",
        "status": "pass",
        "purpose": "repair_post_cleanup_retained_custody_namespace_verification",
        "destructive": False,
        "created_utc": now(),
        "original_cleanup_helper": {
            "namespace": "packet",
            "descriptor": helper_record,
            "actual": helper_actual,
        },
        "pre_removal_receipt": {
            "namespace": "packet",
            "descriptor": pre_binding,
            "actual": pre_actual,
        },
        "failed_cleanup_receipt": {
            "namespace": "packet",
            "actual": descriptor(FINAL, relative_to=PACKET),
            "status": final["status"],
            "error": final["error"],
        },
        "removed_roots_absent": absent,
        "namespace_rules": {
            "packet": "all original retained descriptors except the four exact change-0885 paths resolve relative to the 0890 packet",
            "root": sorted(EXTERNAL),
        },
        "retained_artifact_count": len(verified),
        "retained_artifacts": verified,
        "release_binaries": binaries,
    }
    write_exclusive(result)
    print(json.dumps({"status": result["status"], "retained_artifacts": len(verified), "receipt": str(RECEIPT)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
