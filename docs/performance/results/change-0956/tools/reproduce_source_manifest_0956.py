#!/usr/bin/env python3
"""Regenerate the bounded 0956 source-manifest-v3 census.

This helper only reads a prepared source tree and writes a new manifest path.
It does not invoke Cargo, capture a workload, or modify the source tree.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
from pathlib import Path

DEFAULT_EXCLUDES = ["docs/performance/", "docs/report/"]
DEFAULT_OVERLAYS = [
    "tools/perf-baseline/README.md",
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/xlsx_managed_compression.rs",
    "tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs",
    "tools/perf-baseline/examples/xlsx_managed_compression_benchmark.rs",
]


def digest(path: Path) -> tuple[int, str]:
    h = hashlib.sha256()
    size = 0
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
            size += len(chunk)
    return size, h.hexdigest()


def excluded(relative: str, prefixes: list[str]) -> bool:
    return any(relative.startswith(prefix) for prefix in prefixes)


def census(root: Path, prefixes: list[str]) -> tuple[list[dict[str, object]], int]:
    if not root.is_dir() or root.is_symlink():
        raise SystemExit(f"source root is not a regular directory: {root}")
    rows: list[dict[str, object]] = []
    total = 0
    for directory, dirs, names in os.walk(root, topdown=True, followlinks=False):
        directory_path = Path(directory)
        for name in list(dirs):
            path = directory_path / name
            relative = path.relative_to(root).as_posix() + "/"
            if path.is_symlink():
                raise SystemExit(f"source directory symlink: {path}")
            if excluded(relative, prefixes):
                dirs.remove(name)
        dirs.sort()
        for name in sorted(names):
            path = directory_path / name
            relative = path.relative_to(root).as_posix()
            if excluded(relative, prefixes):
                continue
            if path.is_symlink() or not path.is_file():
                raise SystemExit(f"source entry is not a regular file: {path}")
            size, sha256 = digest(path)
            rows.append({
                "bytes": size,
                "mode": stat.S_IMODE(path.stat().st_mode),
                "path": relative,
                "sha256": sha256,
            })
            total += size
    rows.sort(key=lambda row: str(row["path"]))
    return rows, total


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--base-commit", required=True)
    parser.add_argument("--measurement-source-commit", required=True)
    parser.add_argument("--base-source-root", required=True)
    parser.add_argument("--base-manifest-sha256", required=True)
    parser.add_argument("--prior-integration-manifest-sha256", required=True)
    parser.add_argument("--integration-revision", default="approved-overlay-v3")
    parser.add_argument("--source-root-label", default="source-v1")
    parser.add_argument("--overlay-path", action="append", dest="overlay_paths")
    parser.add_argument("--exclude", action="append", dest="excluded_prefixes")
    args = parser.parse_args()
    if args.output.exists():
        raise SystemExit(f"refusing to overwrite existing output: {args.output}")
    excludes = args.excluded_prefixes or list(DEFAULT_EXCLUDES)
    overlays = args.overlay_paths or list(DEFAULT_OVERLAYS)
    rows, total = census(args.source_root.resolve(), excludes)
    result = {
        "base_commit": args.base_commit,
        "base_manifest_sha256": args.base_manifest_sha256,
        "base_source_root": args.base_source_root,
        "bytes": total,
        "excluded_prefixes": excludes,
        "file_count": len(rows),
        "files": rows,
        "integration_revision": args.integration_revision,
        "measurement_source_commit": args.measurement_source_commit,
        "overlay_paths": overlays,
        "prior_integration_manifest_sha256": args.prior_integration_manifest_sha256,
        "schema": "litchi.performance.0956.source-manifest.v3",
        "source_root": args.source_root_label,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"output": str(args.output), "bytes": total, "file_count": len(rows)}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
