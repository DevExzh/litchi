#!/usr/bin/env python3
"""Verify raw SVG lifecycle profile receipts and recompute the report."""

from __future__ import annotations

import hashlib
import json
import math
from pathlib import Path

from summarize import LANES, rows

MALFORMED = {
    "limit_small",
    "limit_large",
    "malformed_small",
    "malformed_large",
    "namespace_limit_refusal",
}

REFUSAL_PREFIX = {
    "limit_small": "refused:limit",
    "limit_large": "refused:limit",
    "malformed_small": "refused:invalid",
    "malformed_large": "refused:invalid",
    "namespace_limit_refusal": "refused:namespace-active-limit",
}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def resolve(root: Path, shown: str) -> Path:
    path = Path(shown)
    return path if path.is_absolute() else root / path


def verify_manifest(path: Path, root: Path) -> int:
    lines = path.read_text().splitlines()
    require(lines and lines[0] == "format=pptx-svg-lifecycle-build-source-v1", "manifest format changed")
    packages: dict[tuple[str, str, str], tuple[int, str]] = {}
    files: dict[tuple[str, str, str], list[tuple[str, str]]] = {}
    extras: list[tuple[str, str]] = []
    for line in lines[1:]:
        if line.startswith("metadata_sha256="):
            continue
        parts = line.split("\t")
        if line.startswith("package="):
            require(len(parts) == 7, f"malformed package line: {line}")
            key = (parts[0][len("package=") :], parts[1], parts[3])
            require(key not in packages, f"duplicate package line: {key}")
            packages[key] = (int(parts[5]), parts[6])
        elif line.startswith("file="):
            require(len(parts) == 5, f"malformed file line: {line}")
            key = (parts[0][len("file=") :], parts[1], parts[2])
            files.setdefault(key, []).append((parts[3], parts[4]))
        elif line.startswith("extra=\t"):
            require(len(parts) == 3, f"malformed extra line: {line}")
            extras.append((parts[1], parts[2]))
        else:
            raise AssertionError(f"unknown manifest line: {line}")
    require(set(packages) == set(files), "package/file manifest sets differ")
    checked = 0
    for key, (expected_count, expected_tree) in packages.items():
        entries = files[key]
        require(len(entries) == expected_count, f"package file count changed: {key}")
        tree_payload = "\n".join(f"{shown}\t{digest}" for shown, digest in entries)
        require(hashlib.sha256(tree_payload.encode()).hexdigest() == expected_tree, f"package tree changed: {key}")
        for shown, expected in entries:
            current = resolve(root, shown)
            require(current.is_file(), f"retained manifest path missing: {shown}")
            require(sha256(current) == expected, f"retained manifest hash changed: {shown}")
            checked += 1
    for shown, expected in extras:
        current = resolve(root, shown)
        require(current.is_file(), f"retained extra path missing: {shown}")
        require(sha256(current) == expected, f"retained extra hash changed: {shown}")
        checked += 1
    require(checked >= 10, f"too few retained manifest inputs checked: {checked}")
    return checked


def verify_lanes(results: Path) -> None:
    for lane in LANES:
        paths = sorted(results.glob(f"{lane}-p*.json"))
        require(len(paths) == 3, f"fresh process count mismatch for {lane}")
        expected = lane not in MALFORMED
        input_hashes = set()
        for path in paths:
            value = json.loads(path.read_text())
            require(value["schema"] == "pptx-svg-lifecycle-profile-v1", f"schema mismatch: {path}")
            require(value["lane"] == lane, f"lane mismatch: {path}")
            require(value["warmup"] == 2, f"warm-up mismatch: {path}")
            require(value["sample_count"] == len(value["samples"]) == 20, f"sample count mismatch: {path}")
            require(value["expected_success"] is expected, f"expected status mismatch: {path}")
            input_hashes.add(int(value["input_hash_fnv1a64"]))
            for sample in value["samples"]:
                require(sample["expected_success"] is expected, f"sample expected status mismatch: {path}")
                require(sample["actual_success"] is expected, f"sample actual status mismatch: {path}")
                if expected:
                    require(sample["error"] is None, f"unexpected profile error: {path}")
                else:
                    error = sample["error"]
                    require(isinstance(error, str), f"refusal reason missing: {path}")
                    require(error.startswith(REFUSAL_PREFIX[lane]), f"refusal reason mismatch: {path}")
                require(sample["semantic_ok"] is True, f"semantic gate failed: {path}")
                require(sample["output_exact"] is True, f"output gate failed: {path}")
                direct = int(sample["direct_allocated_bytes"])
                new = int(sample["realloc_new_bytes"])
                old = int(sample["realloc_old_bytes"])
                freed = int(sample["deallocated_bytes"])
                expected_live = int(sample["live_before"]) + direct + new - old - freed
                require(int(sample["requested_alloc_bytes"]) == direct + new, f"allocation equation failed: {path}")
                require(int(sample["live_after"]) == expected_live, f"live equation failed: {path}")
                require(sample["alloc_balance_ok"] is True, f"allocator balance failed: {path}")
                require(sample["alloc_invalid"] is False, f"allocator underflow failed: {path}")
                require(int(sample["alloc_failed"]) == 0, f"allocation failure failed: {path}")
        require(len(input_hashes) == 1, f"fixture hash changed across processes for {lane}")


def verify_report(path: Path, recomputed: list[dict[str, object]]) -> None:
    parsed = {}
    for line in path.read_text().splitlines():
        if not line.startswith("| ") or line.startswith("|---") or line.startswith("| lane"):
            continue
        fields = [field.strip() for field in line.strip().strip("|").split("|")]
        require(len(fields) == 8, f"malformed report row: {line}")
        elapsed = tuple(int(item) for item in fields[4].split("/"))
        alloc = tuple(int(item) for item in fields[5].split("/"))
        peak = tuple(int(item) for item in fields[6].split("/"))
        rss = tuple(int(item) for item in fields[7].replace("–", "-").split("-"))
        parsed[fields[0]] = {
            "processes": int(fields[1]),
            "samples": int(fields[2]),
            "input_bytes": int(fields[3]),
            "elapsed": elapsed,
            "alloc": alloc,
            "peak": peak,
            "rss": rss,
        }
    require(set(parsed) == set(LANES), "report lane set differs")
    for row in recomputed:
        actual = parsed[str(row["lane"])]
        for field in ("processes", "samples", "input_bytes", "elapsed", "alloc", "peak", "rss"):
            require(actual[field] == row[field], f"report {field} differs: {row['lane']}")


def main() -> None:
    here = Path(__file__).resolve().parent
    root = next(path for path in here.parents if (path / "crates").is_dir())
    results = here / "results"
    before = results / "source-manifest-before.txt"
    after = results / "source-manifest-after.txt"
    require(before.read_bytes() == after.read_bytes(), "source manifest changed during profile")
    manifest_sha = sha256(before)
    checked = verify_manifest(before, root)
    verify_lanes(results)
    recomputed = rows(results)
    verify_report(here / "report.md", recomputed)
    result = {
        "passed": True,
        "lanes": len(LANES),
        "processes_per_lane": 3,
        "samples_per_process": 20,
        "source_manifest_sha256": manifest_sha,
        "manifest_inputs_checked": checked,
        "expected_rejections": sorted(MALFORMED),
        "report_recomputed_from_samples": True,
    }
    (here / "verification.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
