#!/usr/bin/env python3
"""Rebuild the bounded 0957 source-manifest-v2 census.

This standalone provenance helper only reads a frozen source tree and a base
manifest and writes a new manifest path. It does not invoke Cargo, capture a
workload, or modify the source tree. Existing output paths are rejected.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
from pathlib import Path

BASE_COMMIT = "02449f394529766d76794332204c27c4022931fc"
OVERLAYS = [
    "tools/perf-baseline/Cargo.toml",
    "tools/perf-baseline/README.md",
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs",
    "tools/perf-baseline/src/xlsx_managed_compression_allocation.rs",
    "tools/perf-baseline/src/bin/xlsx_managed_compression_alloc.rs",
]
ADDED = [
    "tools/perf-baseline/src/bin/xlsx_managed_compression_alloc.rs",
    "tools/perf-baseline/src/xlsx_managed_compression_allocation.rs",
]
CHANGED = [
    "tools/perf-baseline/Cargo.toml",
    "tools/perf-baseline/README.md",
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs",
]


def digest(path: Path) -> tuple[int, str]:
    hasher = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
            size += len(chunk)
    return size, hasher.hexdigest()


def census(root: Path) -> tuple[list[dict[str, object]], int]:
    if not root.is_dir() or root.is_symlink():
        raise SystemExit(f"source root is not a regular directory: {root}")
    rows: list[dict[str, object]] = []
    total = 0
    for directory, directories, names in os.walk(root, topdown=True, followlinks=False):
        directory_path = Path(directory)
        for name in directories:
            path = directory_path / name
            if path.is_symlink():
                raise SystemExit(f"source directory symlink: {path}")
        directories.sort()
        for name in sorted(names):
            path = directory_path / name
            relative = path.relative_to(root).as_posix()
            if path.is_symlink() or not path.is_file():
                raise SystemExit(f"source entry is not a regular file: {path}")
            size, sha256 = digest(path)
            rows.append(
                {
                    "bytes": size,
                    "mode": stat.S_IMODE(path.stat().st_mode),
                    "path": relative,
                    "sha256": sha256,
                }
            )
            total += size
    rows.sort(key=lambda row: str(row["path"]))
    return rows, total


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--base-manifest", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--created-utc", required=True)
    parser.add_argument("--base-commit", default=BASE_COMMIT)
    parser.add_argument("--base-source-root", default="source-base-v1")
    parser.add_argument("--source-root-label", default="source-v1")
    args = parser.parse_args()
    if args.output.exists():
        raise SystemExit(f"refusing to overwrite existing output: {args.output}")
    base_manifest_bytes, base_manifest_sha = digest(args.base_manifest)
    base_manifest = json.loads(args.base_manifest.read_text())
    if base_manifest.get("schema") != "litchi.performance.0957.source-manifest.v1":
        raise SystemExit("unexpected base manifest schema")
    rows, total = census(args.source_root.resolve())
    result = {
        "added_paths": ADDED,
        "base_commit": args.base_commit,
        "base_source_snapshot": {
            "bytes": base_manifest["bytes"],
            "file_count": base_manifest["file_count"],
            "manifest": {
                "bytes": base_manifest_bytes,
                "executable": False,
                "mode": stat.S_IMODE(args.base_manifest.stat().st_mode),
                "path": args.base_manifest.name,
                "sha256": base_manifest_sha,
            },
            "root": args.base_source_root,
            "source_manifest_schema": base_manifest["schema"],
        },
        "bytes": total,
        "changed_paths": CHANGED,
        "created_utc": args.created_utc,
        "deleted_paths": [],
        "file_count": len(rows),
        "files": rows,
        "overlay_count": len(OVERLAYS),
        "overlay_paths": OVERLAYS,
        "schema": "litchi.performance.0957.source-manifest.v2",
        "source_root": args.source_root_label,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"bytes": total, "file_count": len(rows), "output": str(args.output)}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
