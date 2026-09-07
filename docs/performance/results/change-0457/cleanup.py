#!/usr/bin/env python3
"""Remove only the owned 0457 staging tree after both proof gates pass."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
TASK = Path("/tmp/litchi-goal-0457")
PRE = ROOT / "precleanup.json"
PORTABLE = ROOT / "portable-verification.json"
INVENTORY = ROOT / "temporary-artifacts.json"
RECEIPT = ROOT / "cleanup.json"


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def fail(message: str) -> "NoReturn":
    raise SystemExit(message)


def regular(path: Path, label: str) -> Path:
    if not path.is_file() or path.is_symlink():
        fail(f"{label} is missing, not regular, or a symlink: {path}")
    return path


def load(path: Path, label: str) -> dict[str, Any]:
    regular(path, label)
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label} is invalid JSON: {error}")
    if not isinstance(value, dict):
        fail(f"{label} must contain a JSON object")
    return value


def reject_symlinks(root: Path) -> None:
    if root.is_symlink():
        fail(f"owned task root is a symlink: {root}")
    for directory, names, files in os.walk(root, followlinks=False):
        for name in (*names, *files):
            path = Path(directory) / name
            if path.is_symlink():
                fail(f"owned task tree contains a symlink: {path}")


def inventory(root: Path) -> tuple[list[dict[str, Any]], int, int]:
    rows: list[dict[str, Any]] = []
    total_bytes = 0
    for directory, names, files in os.walk(root, followlinks=False):
        for name in (*names, *files):
            path = Path(directory) / name
            if path.is_symlink():
                fail(f"owned task tree contains a symlink: {path}")
        for name in sorted(files):
            path = Path(directory) / name
            if not path.is_file():
                fail(f"owned task entry is not a regular file: {path}")
            size = path.stat().st_size
            rows.append({"path": str(path.relative_to(root)), "bytes": size, "sha256": sha(path)})
            total_bytes += size
    rows.sort(key=lambda row: row["path"])
    return rows, len(rows), total_bytes


def require_pass(path: Path, label: str) -> dict[str, Any]:
    value = load(path, label)
    if value.get("status") != "pass" or value.get("exit_code") != 0:
        fail(f"{label} is not a passing proof")
    return value


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.parse_args(sys.argv[1:] if argv is None else argv)
    if INVENTORY.exists() or RECEIPT.exists():
        fail("refusing to overwrite retained cleanup receipts")
    pre = require_pass(PRE, "precleanup receipt")
    portable = require_pass(PORTABLE, "portable verification receipt")
    if portable.get("temporary_directory_absent") is not True:
        fail("portable verifier did not remove its temporary copy")
    if not isinstance(portable.get("copy_seal_sha256"), str) or len(portable["copy_seal_sha256"]) != 64:
        fail("portable verifier did not record a valid copied seal hash")
    verifier_path = pre.get("verifier_path")
    verifier_sha = pre.get("verifier_sha256")
    if not isinstance(verifier_path, str) or not verifier_path or not isinstance(verifier_sha, str) or len(verifier_sha) != 64:
        fail("precleanup verifier identity is incomplete")
    if Path(verifier_path).is_absolute() or ".." in Path(verifier_path).parts:
        fail("precleanup verifier path must be bundle-relative")
    if portable.get("verifier_path") != verifier_path or portable.get("verifier_sha256") != verifier_sha:
        fail("portable proof is bound to a different root verifier")
    regular(ROOT / verifier_path, "root bundle verifier")
    if sha(ROOT / verifier_path) != verifier_sha:
        fail("root bundle verifier changed after proof capture")
    if not TASK.is_dir() or TASK.is_symlink() or TASK != Path("/tmp/litchi-goal-0457"):
        fail(f"refusing to remove unexpected owned task root: {TASK}")
    regular(ROOT / "SHA256SUMS", "root bundle seal")
    reject_symlinks(TASK)
    rows, count, total_bytes = inventory(TASK)
    if count == 0:
        fail("owned task root has no files to inventory")
    inventory_record = {
        "schema": "litchi-0457-temporary-artifacts-v1",
        "change": 457,
        "task": str(TASK),
        "files": count,
        "bytes": total_bytes,
        "artifacts": rows,
    }
    with INVENTORY.open("x", encoding="utf-8") as output:
        output.write(json.dumps(inventory_record, ensure_ascii=False, indent=2) + "\n")
    shutil.rmtree(TASK)
    if TASK.exists() or TASK.is_symlink():
        fail("owned task root remains after cleanup")
    cleanup_record = {
        "schema": "litchi-0457-cleanup-v1",
        "change": 457,
        "status": "pass",
        "task": str(TASK),
        "temporary_directory_absent": True,
        "files_removed": count,
        "bytes_removed": total_bytes,
        "inventory": {
            "path": INVENTORY.name,
            "bytes": INVENTORY.stat().st_size,
            "sha256": sha(INVENTORY),
        },
        "precleanup_sha256": sha(PRE),
        "portable_verification_sha256": sha(PORTABLE),
        "verifier_path": verifier_path,
        "verifier_sha256": verifier_sha,
        "driver_sha256": sha(Path(__file__).resolve()),
    }
    with RECEIPT.open("x", encoding="utf-8") as output:
        output.write(json.dumps(cleanup_record, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({key: cleanup_record[key] for key in ("status", "files_removed", "bytes_removed", "temporary_directory_absent")}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
