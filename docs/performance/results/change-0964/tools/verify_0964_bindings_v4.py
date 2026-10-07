#!/usr/bin/env python3
"""Read-only custody gate for the 0964 reader transport trial."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
import sys
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0964")
PRODUCTION_ALLOWLIST = {
    "crates/litchi-xlsx/src/cell_values/validation.rs",
    "crates/litchi-xlsx/src/raw/worksheet/codec.rs",
}
EXTRA_ROOT_FILES = {"Cargo.lock"}
EXPECTED_SOURCE_COUNT = 10_933
RETAINED_SOURCE_COUNT = 10_932
EXPECTED_PERF_LOCK = "db62bdd9174484613acae1f49f8187cf105d0b4976e2fea4de91df7372a893d5"
EXPECTED_BENCH = "5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3"
EXPECTED_ALLOC = "9ba37a6b9796ad0e316e7571ebe117ba7a5f97bcf5d29153c7cb3d74c485e9a5"


def fail(message: str) -> None:
    raise RuntimeError(message)


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            value.update(block)
    return value.hexdigest()


def files(root: Path) -> dict[str, tuple[int, int, str]]:
    if not root.is_dir() or root.is_symlink():
        fail(f"missing or symlinked source root: {root}")
    # Keep the tuple order identical to manifest_entries: bytes, mode, SHA-256.
    result: dict[str, tuple[int, int, str]] = {}
    for path in sorted(root.rglob("*")):
        if path.is_symlink():
            fail(f"symlink is not permitted in source custody: {path}")
        if path.is_file():
            result[path.relative_to(root).as_posix()] = (
                path.stat().st_size, stat.S_IMODE(path.stat().st_mode), digest(path)
            )
    return result


def manifest_entries(path: Path) -> dict[str, tuple[int, int, str]]:
    if not path.is_file() or path.is_symlink():
        fail(f"source manifest missing: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        fail(f"invalid source manifest {path}: {exc}")
    rows = value.get("files") if isinstance(value, dict) else None
    if not isinstance(rows, list) or value.get("file_count") != len(rows):
        fail(f"invalid source manifest shape: {path}")
    result: dict[str, tuple[int, int, str]] = {}
    for row in rows:
        if not isinstance(row, dict):
            fail(f"invalid source manifest row: {path}")
        name = row.get("path")
        if not isinstance(name, str) or name in result:
            fail(f"invalid or duplicate manifest path: {path}: {name!r}")
        result[name] = (row.get("bytes"), row.get("mode"), row.get("sha256"))
    return result


def parse_allowlist(path: Path | None) -> set[str]:
    if path is None:
        return set()
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
    result = set(value)
    if any(path.startswith("/") or ".." in Path(path).parts for path in result):
        fail("test allowlist paths must be relative")
    if result & PRODUCTION_ALLOWLIST:
        fail("test allowlist repeats a production path")
    return result


def compare(left: dict[str, tuple[Any, ...]], right: dict[str, tuple[Any, ...]],
            *, left_extra: set[str] = set(), changed: set[str] = set()) -> list[str]:
    errors: list[str] = []
    for name in sorted(set(left) | set(right)):
        if name not in right:
            if name not in left_extra:
                errors.append(f"left-only file: {name}")
        elif name not in left:
            if name not in changed and name not in left_extra:
                errors.append(f"right-only file: {name}")
        elif left[name] != right[name] and name not in changed:
            errors.append(f"content/mode mismatch: {name}")
    return errors


def executable(path: Path, expected_sha: str) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink() or not os.access(path, os.X_OK):
        fail(f"missing or non-executable binary: {path}")
    actual = digest(path)
    if actual != expected_sha:
        fail(f"executable hash mismatch: {path}: {actual} != {expected_sha}")
    return {"path": str(path.resolve()), "bytes": path.stat().st_size,
            "mode": stat.S_IMODE(path.stat().st_mode), "executable": True, "sha256": actual}


def binary_descriptor(path: Path) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink() or not os.access(path, os.X_OK):
        fail(f"missing or non-executable candidate binary: {path}")
    return {"path": str(path.resolve()), "bytes": path.stat().st_size,
            "mode": stat.S_IMODE(path.stat().st_mode), "executable": True, "sha256": digest(path)}


def verify_manifest_tree(path: Path, manifest: Path, label: str) -> None:
    expected = manifest_entries(manifest)
    actual = files(path)
    if len(expected) != len(actual):
        fail(f"{label} manifest/tree count differs: {len(expected)} != {len(actual)}")
    errors = []
    for name in sorted(set(expected) | set(actual)):
        if name not in actual:
            errors.append(f"{label} manifest-only: {name}")
        elif name not in expected and name != "Cargo.lock":
            errors.append(f"{label} tree-only: {name}")
        elif expected[name] != actual[name]:
            errors.append(f"{label} manifest/tree mismatch: {name}")
    if errors:
        fail("\n".join(errors[:20]))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=("baseline", "candidate"), default="baseline")
    parser.add_argument("--baseline-source", type=Path, required=True)
    parser.add_argument("--retained-source", type=Path, default=Path("/home/zhuhe/litchi-performance-0957/source-v1"))
    parser.add_argument("--candidate-source", type=Path)
    parser.add_argument("--baseline-manifest", type=Path, default=ROOT / "source-manifest-baseline-v2.json")
    parser.add_argument("--candidate-manifest", type=Path, default=ROOT / "source-manifest-candidate-v4.json")
    parser.add_argument("--test-allowlist-file", type=Path)
    parser.add_argument("--benchmark", type=Path, default=Path("/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_benchmark.v2"))
    parser.add_argument("--allocator", type=Path, default=Path("/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_alloc.v2"))
    parser.add_argument("--candidate-benchmark", type=Path)
    parser.add_argument("--candidate-allocator", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    baseline = files(args.baseline_source)
    retained = files(args.retained_source)
    if len(baseline) != EXPECTED_SOURCE_COUNT:
        fail(f"baseline file count {len(baseline)} != {EXPECTED_SOURCE_COUNT} including required root-lock overlay")
    if len(retained) != RETAINED_SOURCE_COUNT:
        fail(f"retained file count {len(retained)} != {RETAINED_SOURCE_COUNT}")
    errors = compare(baseline, retained, left_extra=EXTRA_ROOT_FILES)
    if errors:
        fail("baseline/retained source mismatch:\n" + "\n".join(errors[:20]))
    perf_lock = args.baseline_source / "tools/perf-baseline/Cargo.lock"
    if digest(perf_lock) != EXPECTED_PERF_LOCK:
        fail(f"perf-baseline Cargo.lock mismatch: {digest(perf_lock)}")
    if args.baseline_manifest:
        verify_manifest_tree(args.baseline_source, args.baseline_manifest, "baseline")

    result: dict[str, Any] = {
        "schema": "litchi.performance.0964.source-binary-binding.v4",
        "status": "pass",
        "mode": args.mode,
        "baseline_source": {"path": str(args.baseline_source.resolve()), "files": len(baseline)},
        "retained_source": {"path": str(args.retained_source.resolve()), "files": len(retained)},
        "root_lock_overlay": ({"path": str((args.baseline_source / "Cargo.lock").resolve()), "sha256": baseline["Cargo.lock"][2], "present": True, "excluded_from_retained_comparison": True}
                              if "Cargo.lock" in baseline else {"present": False, "excluded_from_retained_comparison": True}),
        "perf_baseline_lock_sha256": EXPECTED_PERF_LOCK,
        "retained_executables": [
            executable(args.benchmark, EXPECTED_BENCH), executable(args.allocator, EXPECTED_ALLOC)
        ],
    }
    if args.mode == "candidate":
        if args.candidate_source is None or args.candidate_manifest is None:
            fail("candidate mode requires --candidate-source and --candidate-manifest")
        candidate = files(args.candidate_source)
        if len(candidate) != EXPECTED_SOURCE_COUNT:
            fail(f"candidate file count {len(candidate)} != {EXPECTED_SOURCE_COUNT} including required root-lock overlay")
        if args.candidate_manifest:
            verify_manifest_tree(args.candidate_source, args.candidate_manifest, "candidate")
        test_paths = parse_allowlist(args.test_allowlist_file)
        allowlist = PRODUCTION_ALLOWLIST | test_paths
        errors = compare(baseline, candidate, changed=allowlist)
        if errors:
            fail("candidate has changes outside the allowlist:\n" + "\n".join(errors[:20]))
        changed = sorted(name for name in set(baseline) | set(candidate) if baseline.get(name) != candidate.get(name))
        if set(changed) != allowlist:
            fail(f"candidate changed files {changed!r}; expected {sorted(allowlist)!r}")
        if not PRODUCTION_ALLOWLIST.issubset(set(changed)):
            fail("both approved production hot-loop paths must change")
        result["candidate_source"] = {"path": str(args.candidate_source.resolve()), "files": len(candidate), "changed_files": changed, "test_only_changed_files": sorted(test_paths)}
        if args.candidate_benchmark and args.candidate_allocator:
            result["candidate_executables"] = [
                binary_descriptor(args.candidate_benchmark),
                binary_descriptor(args.candidate_allocator),
            ]

    if args.output.exists() or args.output.is_symlink():
        fail(f"refusing to overwrite receipt: {args.output}")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, KeyError) as exc:
        print(f"verify_0964_bindings_v4: FAIL: {exc}", file=sys.stderr)
        raise SystemExit(2)
