#!/usr/bin/env python3
"""Create a final 0964 candidate source manifest after an explicit scope check.

This helper only reads the two exported source trees and writes new custody
artifacts.  It does not build, run tests, or modify either source tree.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import stat
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

PRODUCTION_ALLOWLIST = {
    "crates/litchi-xlsx/src/cell_values/validation.rs",
    "crates/litchi-xlsx/src/raw/worksheet/codec.rs",
}
EXPECTED_SOURCE_COUNT = 10_933


def fail(message: str) -> None:
    raise RuntimeError(message)


def sha256(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            value.update(block)
    return value.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink():
        fail(f"regular file required: {path}")
    return {
        "path": str(path.resolve()),
        "bytes": path.stat().st_size,
        "mode": stat.S_IMODE(path.stat().st_mode),
        "sha256": sha256(path),
    }


def tree(root: Path) -> dict[str, tuple[int, int, str]]:
    if not root.is_dir() or root.is_symlink():
        fail(f"source root missing or symlinked: {root}")
    result: dict[str, tuple[int, int, str]] = {}
    for path in sorted(root.rglob("*")):
        if path.is_symlink():
            fail(f"symlink is not permitted in source tree: {path}")
        if path.is_file():
            relative = path.relative_to(root).as_posix()
            if relative in result:
                fail(f"duplicate source path: {relative}")
            mode = stat.S_IMODE(path.stat().st_mode)
            result[relative] = (path.stat().st_size, mode, sha256(path))
    return result


def read_manifest(path: Path) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink():
        fail(f"manifest missing: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        fail(f"invalid manifest {path}: {exc}")
    if not isinstance(value, dict) or not isinstance(value.get("files"), list):
        fail(f"invalid manifest shape: {path}")
    if value.get("file_count") != len(value["files"]):
        fail(f"manifest count mismatch: {path}")
    return value


def manifest_tree(path: Path, manifest: dict[str, Any], label: str) -> None:
    expected: dict[str, tuple[Any, ...]] = {}
    for row in manifest["files"]:
        if not isinstance(row, dict) or not isinstance(row.get("path"), str):
            fail(f"invalid {label} manifest row")
        name = row["path"]
        if name in expected:
            fail(f"duplicate {label} manifest path: {name}")
        expected[name] = (row.get("bytes"), row.get("mode"), row.get("sha256"))
    actual = tree(path)
    if expected != actual:
        differences = []
        for name in sorted(set(expected) | set(actual)):
            if expected.get(name) != actual.get(name):
                differences.append(name)
        fail(f"{label} manifest/tree mismatch: {differences[:20]}")


def test_paths(path: Path) -> list[str]:
    if not path.is_file() or path.is_symlink():
        fail(f"test allowlist missing: {path}")
    try:
        value: Any = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        value = [line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
    if isinstance(value, dict):
        value = value.get("paths")
    if not isinstance(value, list) or not all(isinstance(item, str) and item for item in value):
        fail("test allowlist must be a JSON list or one relative path per line")
    paths = sorted(set(value))
    if any(path.startswith("/") or ".." in Path(path).parts for path in paths):
        fail("test allowlist paths must be relative")
    if set(paths) & PRODUCTION_ALLOWLIST:
        fail("test allowlist repeats a production path")
    return paths


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--baseline-source", type=Path, required=True)
    parser.add_argument("--candidate-source", type=Path, required=True)
    parser.add_argument("--baseline-manifest", type=Path, required=True)
    parser.add_argument("--candidate-template", type=Path, required=True)
    parser.add_argument("--test-allowlist-file", type=Path, required=True)
    parser.add_argument("--output-manifest", type=Path, required=True)
    parser.add_argument("--output-receipt", type=Path, required=True)
    args = parser.parse_args()

    if args.output_manifest.exists() or args.output_manifest.is_symlink():
        fail(f"refusing to overwrite manifest: {args.output_manifest}")
    if args.output_receipt.exists() or args.output_receipt.is_symlink():
        fail(f"refusing to overwrite receipt: {args.output_receipt}")

    baseline_manifest = read_manifest(args.baseline_manifest)
    candidate_template = read_manifest(args.candidate_template)
    baseline = tree(args.baseline_source)
    candidate = tree(args.candidate_source)
    if len(baseline) != EXPECTED_SOURCE_COUNT or len(candidate) != EXPECTED_SOURCE_COUNT:
        fail(f"source count must be {EXPECTED_SOURCE_COUNT}: baseline={len(baseline)}, candidate={len(candidate)}")
    manifest_tree(args.baseline_source, baseline_manifest, "baseline")
    changed = sorted(name for name in set(baseline) | set(candidate) if baseline.get(name) != candidate.get(name))
    tests = test_paths(args.test_allowlist_file)
    allowlist = PRODUCTION_ALLOWLIST | set(tests)
    if set(changed) != allowlist:
        fail(f"candidate changed paths {changed!r}; expected {sorted(allowlist)!r}")
    if not PRODUCTION_ALLOWLIST.issubset(set(changed)):
        fail("both approved production paths must change")

    rows = [
        {"bytes": values[0], "mode": values[1], "path": name, "sha256": values[2]}
        for name, values in sorted(candidate.items())
    ]
    output_manifest = dict(candidate_template)
    output_manifest.update({
        "schema": "litchi.performance.0964.source-manifest.v4",
        "created_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_root": str(args.candidate_source.resolve()),
        "role": "candidate",
        "candidate_status": "source_frozen",
        "candidate_changed_paths": changed,
        "files": rows,
        "file_count": len(rows),
        "bytes": sum(row["bytes"] for row in rows),
    })
    output_manifest["root_lock_overlay"] = {
        "path": "Cargo.lock",
        "bytes": candidate["Cargo.lock"][0],
        "mode": candidate["Cargo.lock"][1],
        "sha256": candidate["Cargo.lock"][2],
        "source": "/home/zhuhe/litchi-performance-0961/source-baseline/Cargo.lock",
    }

    args.output_manifest.parent.mkdir(parents=True, exist_ok=True)
    args.output_manifest.write_text(json.dumps(output_manifest, indent=2) + "\n", encoding="utf-8")
    receipt = {
        "schema": "litchi.performance.0964.candidate-freeze.v1",
        "status": "pass",
        "created_utc": output_manifest["created_utc"],
        "source_roots": {
            "baseline": str(args.baseline_source.resolve()),
            "candidate": str(args.candidate_source.resolve()),
        },
        "baseline_manifest": descriptor(args.baseline_manifest),
        "candidate_manifest": descriptor(args.output_manifest),
        "candidate_template": descriptor(args.candidate_template),
        "test_allowlist": descriptor(args.test_allowlist_file),
        "production_paths": sorted(PRODUCTION_ALLOWLIST),
        "test_paths": tests,
        "changed_paths": changed,
        "source_count": len(candidate),
        "build_or_workload_launched": False,
        "quality_and_root_go_required_before_build": True,
    }
    args.output_receipt.parent.mkdir(parents=True, exist_ok=True)
    args.output_receipt.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": "pass", "manifest": str(args.output_manifest), "receipt": str(args.output_receipt), "changed_paths": changed}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, KeyError) as exc:
        print(f"freeze_candidate_scope_v1: FAIL: {exc}")
        raise SystemExit(2)
