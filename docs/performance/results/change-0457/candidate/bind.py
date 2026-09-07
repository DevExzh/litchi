#!/usr/bin/env python3
"""Bind a completed 0457 source-tail build to the candidate capture bundle.

This helper only copies the successful build/source receipts and records the
already-retained executable identities.  It never builds or executes a binary.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
from typing import Any


class BindError(ValueError):
    pass


def fail(message: str) -> None:
    raise BindError(message)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def copy_file(source: Path, destination: Path) -> None:
    if not source.is_file() or source.is_symlink():
        fail(f"{source}: expected a regular file")
    if destination.exists() or destination.is_symlink():
        fail(f"refusing to overwrite existing bound artifact {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, destination)
    if sha(source) != sha(destination):
        fail(f"copy verification failed for {destination}")


def executable(path: Path, label: str) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink() or not path.stat().st_mode & 0o111:
        fail(f"{label}: expected a regular executable")
    return {
        "source_path": str(path),
        "copy_path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
        "profile": "release",
        "executable": True,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--build-receipt", type=Path, required=True)
    parser.add_argument("--repo-root", type=Path, required=True)
    parser.add_argument("--normal-source", type=Path, required=True)
    parser.add_argument("--allocator-source", type=Path, required=True)
    parser.add_argument("--normal-copy", type=Path, required=True)
    parser.add_argument("--allocator-copy", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    try:
        output = args.output_dir.resolve()
        repo = args.repo_root.resolve()
        runs = output / "runs"
        if runs.is_dir() and any(runs.iterdir()):
            fail("capture has started; refusing to rewrite build/source bindings")
        for target in (output / "build-receipt.json", output / "source-manifest.json", output / "binary-bindings.json"):
            if target.exists() or target.is_symlink():
                fail(f"refusing to overwrite existing bound artifact {target}")
        receipt_source = args.build_receipt.resolve()
        build = obj(load(receipt_source), "build receipt")
        if build.get("change") != 457 or build.get("status") != "pass" or build.get("source_unchanged") is not True:
            fail("build receipt must be a successful source-unchanged 0457 build")
        revision = build.get("revision")
        if not isinstance(revision, str) or not revision:
            fail("build receipt revision is missing")
        source_after = obj(build.get("source_after"), "build.source_after")
        source_path = repo.parent / "docs" / "performance" / "results" / "change-0457" / source_after.get("path", "")
        # The build receipt path is relative to change-0457; derive that root
        # from the output bundle so replay remains portable.
        change_root = output.parent
        source_path = change_root / source_after.get("path", "")
        if not source_path.is_file():
            fail(f"source manifest from build receipt is missing: {source_path}")
        copy_file(receipt_source, output / "build-receipt.json")
        copy_file(source_path, output / "source-manifest.json")
        source_manifest_sha = sha(output / "source-manifest.json")
        if source_after.get("sha256") != source_manifest_sha or source_after.get("files") != len(load(output / "source-manifest.json")):
            fail("copied source-manifest.json does not bind build.source_after")
        normal_source = args.normal_source.resolve() if args.normal_source.is_absolute() else (repo / args.normal_source).resolve()
        allocator_source = args.allocator_source.resolve() if args.allocator_source.is_absolute() else (repo / args.allocator_source).resolve()
        normal_copy = args.normal_copy.resolve()
        allocator_copy = args.allocator_copy.resolve()
        for source, copied, label in ((normal_source, normal_copy, "normal"), (allocator_source, allocator_copy, "allocator")):
            if not source.is_file() or source.is_symlink() or not copied.is_file() or copied.is_symlink():
                fail(f"{label} source/copy executable is missing")
            if sha(source) != sha(copied) or source.stat().st_size != copied.stat().st_size:
                fail(f"{label} source/copy executable identity differs")
        def item(source: Path, copied: Path) -> dict[str, Any]:
            if not copied.stat().st_mode & 0o111:
                fail(f"{copied}: copied binary is not executable")
            return {
                "source_path": str(source.relative_to(repo)),
                "copy_path": str(copied),
                "bytes": copied.stat().st_size,
                "sha256": sha(copied),
                "profile": "release",
                "executable": True,
            }
        receipt_record = {"path": "build-receipt.json", "sha256": sha(output / "build-receipt.json")}
        source_record = {"path": "source-manifest.json", "sha256": source_manifest_sha, "files": len(load(output / "source-manifest.json"))}
        bindings = {
            "schema": "litchi-0457-source-tail-candidate-binary-binding-v1",
            "build_receipt": receipt_record,
            "source_manifest": source_record,
            "revision": revision,
            "binaries": {
                "normal": item(normal_source, normal_copy),
                "allocator": item(allocator_source, allocator_copy),
            },
        }
        binary_binding_path = output / "binary-bindings.json"
        if binary_binding_path.exists() or binary_binding_path.is_symlink():
            fail(f"refusing to overwrite existing bound artifact {binary_binding_path}")
        binary_binding_path.write_text(json.dumps(bindings, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({"status": "pass", "revision": revision, "binary_bindings_sha256": sha(binary_binding_path)}))
        return 0
    except (OSError, TypeError, KeyError, BindError) as error:
        print(f"BIND INVALID: {error}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
