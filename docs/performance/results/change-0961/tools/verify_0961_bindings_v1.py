#!/usr/bin/env python3
"""Read-only custody gate for the 0961 apply_batch trial."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
import sys
from pathlib import Path

CHANGE_ALLOWLIST = {
    "crates/litchi-xlsx/src/cell_values/admission_tests.rs",
    "crates/litchi-xlsx/src/cell_values/source.rs",
}
EXTRA_ROOT_FILES = {"Cargo.lock"}
EXPECTED_SOURCE_COUNT = 10_933
RETAINED_SOURCE_COUNT = 10_932
EXPECTED_PERF_LOCK = "db62bdd9174484613acae1f49f8187cf105d0b4976e2fea4de91df7372a893d5"
EXPECTED_BENCH = "5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3"
EXPECTED_ALLOC = "9ba37a6b9796ad0e316e7571ebe117ba7a5f97bcf5d29153c7cb3d74c485e9a5"


def fail(message: str) -> None:
    raise RuntimeError(message)


def files(root: Path) -> dict[str, tuple[int, str, int]]:
    if not root.is_dir():
        fail(f"missing source root: {root}")
    out: dict[str, tuple[int, str, int]] = {}
    for path in sorted(root.rglob("*")):
        if path.is_symlink():
            fail(f"symlink is not permitted in source custody: {path}")
        if path.is_file():
            data = path.read_bytes()
            out[path.relative_to(root).as_posix()] = (
                len(data), hashlib.sha256(data).hexdigest(), stat.S_IMODE(path.stat().st_mode))
    return out


def compare(left: dict[str, tuple[int, str, int]], right: dict[str, tuple[int, str, int]],
            *, left_extra: set[str] = set(), content: set[str] = set()) -> list[str]:
    errors: list[str] = []
    for name in sorted(set(left) | set(right)):
        if name not in right:
            if name not in left_extra:
                errors.append(f"left-only file: {name}")
        elif name not in left:
            errors.append(f"right-only file: {name}")
        elif left[name] != right[name] and name not in content:
            errors.append(f"content/mode mismatch: {name}")
    return errors


def executable(path: Path, expected: str) -> dict[str, object]:
    if not path.is_file() or path.is_symlink():
        fail(f"missing executable: {path}")
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    if digest != expected:
        fail(f"executable hash mismatch: {path}: {digest} != {expected}")
    if not os.access(path, os.X_OK):
        fail(f"executable bit missing: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": digest}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=("baseline", "candidate"), default="baseline")
    ap.add_argument("--baseline-source", type=Path, required=True)
    ap.add_argument("--retained-source", type=Path,
                    default=Path("/home/zhuhe/litchi-performance-0957/source-v1"))
    ap.add_argument("--candidate-source", type=Path)
    ap.add_argument("--benchmark", type=Path,
                    default=Path("/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_benchmark.v2"))
    ap.add_argument("--allocator", type=Path,
                    default=Path("/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_alloc.v2"))
    ap.add_argument("--output", type=Path, required=True)
    args = ap.parse_args()

    baseline, retained = files(args.baseline_source), files(args.retained_source)
    if len(baseline) != EXPECTED_SOURCE_COUNT:
        fail(f"baseline file count {len(baseline)} != {EXPECTED_SOURCE_COUNT}")
    if len(retained) != RETAINED_SOURCE_COUNT:
        fail(f"retained file count {len(retained)} != {RETAINED_SOURCE_COUNT}")
    errors = compare(baseline, retained, left_extra=EXTRA_ROOT_FILES)
    if errors:
        fail("baseline/retained source mismatch:\n" + "\n".join(errors[:20]))
    lock = args.baseline_source / "tools/perf-baseline/Cargo.lock"
    lock_hash = hashlib.sha256(lock.read_bytes()).hexdigest()
    if lock_hash != EXPECTED_PERF_LOCK:
        fail(f"perf-baseline Cargo.lock mismatch: {lock_hash}")

    result: dict[str, object] = {
        "schema": "litchi.performance.0961.source-binary-binding.v1",
        "status": "pass",
        "mode": args.mode,
        "baseline_source": {"path": str(args.baseline_source), "files": len(baseline)},
        "retained_source": {"path": str(args.retained_source), "files": len(retained)},
        "root_lock_overlay": {
            "path": str(args.baseline_source / "Cargo.lock"),
            "sha256": baseline["Cargo.lock"][1],
            "excluded_from_retained_comparison": True,
        },
        "perf_baseline_lock_sha256": lock_hash,
        "retained_executables": [
            executable(args.benchmark, EXPECTED_BENCH), executable(args.allocator, EXPECTED_ALLOC)
        ],
    }
    if args.mode == "candidate":
        if args.candidate_source is None:
            fail("--candidate-source is required in candidate mode")
        candidate = files(args.candidate_source)
        if len(candidate) != EXPECTED_SOURCE_COUNT:
            fail(f"candidate file count {len(candidate)} != {EXPECTED_SOURCE_COUNT}")
        errors = compare(baseline, candidate, content=CHANGE_ALLOWLIST)
        if errors:
            fail("candidate has changes outside the allowlist:\n" + "\n".join(errors[:20]))
        changed = sorted(name for name in baseline if baseline[name] != candidate.get(name))
        if changed != sorted(CHANGE_ALLOWLIST):
            fail(f"candidate changed files {changed!r}; expected {sorted(CHANGE_ALLOWLIST)!r}")
        result["candidate_source"] = {
            "path": str(args.candidate_source), "files": len(candidate), "changed_files": changed
        }

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, KeyError) as exc:
        print(f"verify_0961_bindings_v1: FAIL: {exc}", file=sys.stderr)
        raise SystemExit(2)
