#!/usr/bin/env python3
"""Verify the targeted theme-family hardening profile receipt."""

from __future__ import annotations

import hashlib
import json
import math
from pathlib import Path


LANES = (
    "native_read",
    "native_replace",
    "native_remove",
    "native_add",
    "unknown_32",
    "unknown_1000",
    "duplicate",
    "limit_replace",
    "limit_add",
)
EXPECTED_FAILURE = {"duplicate", "limit_replace", "limit_add"}


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


def verify_source_manifest(path: Path, root: Path) -> int:
    lines = path.read_text().splitlines()
    require(lines and lines[0] == "format=xlsb-theme-family-build-source-v2", "manifest format changed")
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
    require(checked >= 100, f"too few retained manifest inputs checked: {checked}")
    return checked


def quantile(values: list[int], percent: int) -> int:
    return sorted(values)[max(0, math.ceil(len(values) * percent / 100) - 1)]


def rss(path: Path) -> int:
    for line in path.read_text().splitlines():
        if line.lstrip().startswith("Maximum resident set size (kbytes):"):
            return int(line.split(":", 1)[1].strip())
    raise AssertionError(f"RSS missing: {path}")


def verify_lanes(results: Path) -> list[dict[str, object]]:
    rows = []
    for lane in LANES:
        paths = sorted(results.glob(f"{lane}-p*.json"))
        require(len(paths) == 3, f"fresh process count mismatch for {lane}")
        expected_success = lane not in EXPECTED_FAILURE
        samples = []
        rss_values = []
        for path in paths:
            value = json.loads(path.read_text())
            require(value["schema"] == "theme-family-hardening-profile-v1", f"schema mismatch in {path}")
            require(value["lane"] == lane, f"lane mismatch in {path}")
            require(value["sample_count"] == len(value["samples"]) == 20, f"sample count mismatch in {path}")
            require(value["expected_success"] is expected_success, f"expected status mismatch in {path}")
            rss_values.append(rss(path.with_suffix(".time.txt")))
            for sample in value["samples"]:
                require(sample["expected_success"] is expected_success, f"sample expectation mismatch in {path}")
                require(sample["actual_success"] is expected_success, f"operation status mismatch in {path}")
                direct = int(sample["direct_allocated_bytes"])
                new = int(sample["realloc_new_bytes"])
                old = int(sample["realloc_old_bytes"])
                freed = int(sample["deallocated_bytes"])
                expected_live = int(sample["live_before"]) + direct + new - old - freed
                require(int(sample["requested_alloc_bytes"]) == direct + new, f"allocation accounting failed in {path}")
                require(int(sample["live_after"]) == expected_live, f"live accounting failed in {path}")
                require(sample["alloc_balance_ok"] is True, f"allocator balance failed in {path}")
                require(sample["alloc_invalid"] is False, f"allocator invalid in {path}")
                require(int(sample["alloc_failed"]) == 0, f"allocator failure in {path}")
                samples.append(sample)
        rows.append(
            {
                "lane": lane,
                "processes": 3,
                "samples": len(samples),
                "elapsed": (
                    quantile([int(s["elapsed_ns"]) for s in samples], 50),
                    quantile([int(s["elapsed_ns"]) for s in samples], 95),
                    quantile([int(s["elapsed_ns"]) for s in samples], 99),
                ),
                "alloc": (
                    quantile([int(s["requested_alloc_bytes"]) for s in samples], 50),
                    quantile([int(s["requested_alloc_bytes"]) for s in samples], 95),
                ),
                "peak": (
                    quantile([int(s["peak_live_delta"]) for s in samples], 50),
                    quantile([int(s["peak_live_delta"]) for s in samples], 95),
                ),
                "rss": (min(rss_values), max(rss_values)),
            }
        )
    return rows


def verify_report(path: Path, rows: list[dict[str, object]]) -> None:
    parsed: dict[str, dict[str, object]] = {}
    for line in path.read_text().splitlines():
        if not line.startswith("| ") or line.startswith("|---") or line.startswith("| lane"):
            continue
        parts = [part.strip() for part in line.strip().strip("|").split("|")]
        require(len(parts) == 7, f"malformed report row: {line}")
        lane = parts[0]
        require(lane in LANES and lane not in parsed, f"unexpected report lane: {lane}")
        elapsed = tuple(int(value.strip()) for value in parts[3].split("/"))
        alloc = tuple(int(value.strip()) for value in parts[4].split("/"))
        peak = tuple(int(value.strip()) for value in parts[5].split("/"))
        rss_values = tuple(int(value.strip()) for value in parts[6].replace("–", "-").split("-"))
        require(len(elapsed) == 3 and len(alloc) == 2 and len(peak) == 2 and len(rss_values) == 2, f"report metrics malformed: {line}")
        parsed[lane] = {"processes": int(parts[1]), "samples": int(parts[2]), "elapsed": elapsed, "alloc": alloc, "peak": peak, "rss": rss_values}
    require(set(parsed) == set(LANES), "report lane set differs")
    for row in rows:
        actual = parsed[str(row["lane"])]
        require(actual["processes"] == row["processes"], f"report process count differs: {row['lane']}")
        require(actual["samples"] == row["samples"], f"report sample count differs: {row['lane']}")
        for field in ("elapsed", "alloc", "peak", "rss"):
            require(actual[field] == row[field], f"report {field} differs: {row['lane']}")


def main() -> None:
    here = Path(__file__).resolve().parent
    root = next(path for path in here.parents if (path / "crates").is_dir())
    results = here / "results"
    before = results / "source-manifest-before.txt"
    after = results / "source-manifest-after.txt"
    require(before.read_bytes() == after.read_bytes(), "source manifest changed during profile")
    manifest_sha = sha256(before)
    manifest_count = verify_source_manifest(before, root)
    provenance = (results / "source-provenance.txt").read_text()
    require(f"source_manifest_before_sha256={manifest_sha}" in provenance, "manifest hash missing")
    require(f"source_manifest_after_sha256={manifest_sha}" in provenance, "post-build manifest hash missing")
    source_hashes: dict[str, str] = {}
    for line in provenance.splitlines():
        if "  " not in line:
            continue
        digest, shown = line.split("  ", 1)
        if len(digest) == 64 and shown.startswith("/"):
            current = Path(shown)
            require(current.is_file(), f"source hash path missing: {shown}")
            require(sha256(current) == digest, f"source hash changed: {shown}")
            source_hashes[shown] = digest
    required_sources = (
        root / "crates/litchi-drawingml/src/theme/family/codec.rs",
        root / "crates/litchi-drawingml/src/theme/family/part.rs",
        root / "crates/litchi-drawingml/tests/theme_part.rs",
    )
    require(all(str(path) in source_hashes for path in required_sources), "required source hashes missing")
    rows = verify_lanes(results)
    verify_report(here / "report.md", rows)
    result = {
        "passed": True,
        "lanes": len(LANES),
        "processes_per_lane": 3,
        "samples_per_process": 20,
        "source_manifest_sha256": manifest_sha,
        "manifest_inputs_checked": manifest_count,
        "source_hashes_checked": len(source_hashes),
        "expected_rejections": sorted(EXPECTED_FAILURE),
        "report_recomputed_from_samples": True,
    }
    (here / "verification.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
