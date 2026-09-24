#!/usr/bin/env python3
"""Independently verify the final XLSB theme-family profile evidence."""

from __future__ import annotations

import hashlib
import json
import math
import re
from pathlib import Path


HERE = Path(__file__).resolve().parent
ROOT = next(path for path in HERE.parents if (path / "crates").is_dir())
FIXTURES = ("native", "opaque")
OPERATIONS = (
    "codec_read",
    "metadata_read",
    "source_read",
    "family_clone",
    "noop",
    "add",
    "update",
    "remove",
    "base_edit",
)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def quantile(values: list[int], percent: int) -> int:
    return sorted(values)[math.ceil(len(values) * percent / 100) - 1]


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def manifest_path(shown: str) -> Path:
    path = Path(shown)
    return path if path.is_absolute() else ROOT / path


def verify_source_manifest(path: Path) -> int:
    lines = path.read_text().splitlines()
    require(lines and lines[0] == "format=xlsb-theme-family-build-source-v2", "unexpected source manifest format")
    require(any(line.startswith("metadata_sha256=") for line in lines), "manifest metadata hash is missing")
    packages: dict[tuple[str, str, str], dict[str, object]] = {}
    files: dict[tuple[str, str, str], list[tuple[str, str]]] = {}
    extras: list[tuple[str, str]] = []
    for line in lines[2:]:
        if line.startswith("package="):
            parts = line.split("\t")
            require(len(parts) == 7, f"malformed package manifest line: {line}")
            key = (parts[0][len("package=") :], parts[1], parts[3])
            require(key not in packages, f"duplicate package manifest record: {key}")
            packages[key] = {"count": int(parts[5]), "tree": parts[6]}
        elif line.startswith("file="):
            parts = line.split("\t")
            require(len(parts) == 5, f"malformed file manifest line: {line}")
            key = (parts[0][len("file=") :], parts[1], parts[2])
            files.setdefault(key, []).append((parts[3], parts[4]))
        elif line.startswith("extra=\t"):
            parts = line.split("\t")
            require(len(parts) == 3, f"malformed extra manifest line: {line}")
            extras.append((parts[1], parts[2]))
        else:
            raise AssertionError(f"unknown source manifest line: {line}")

    checked = 0
    require(set(files) == set(packages), "package/file manifest record sets differ")
    for key, package in packages.items():
        package_files = files[key]
        require(len(package_files) == package["count"], f"package file count changed: {key}")
        tree_payload = "\n".join(f"{shown}\t{digest}" for shown, digest in package_files)
        tree = hashlib.sha256(tree_payload.encode()).hexdigest()
        require(tree == package["tree"], f"package tree hash is inconsistent: {key}")
        for shown, expected in package_files:
            current = manifest_path(shown)
            require(current.is_file(), f"retained manifest path is missing: {shown}")
            require(sha256(current) == expected, f"retained manifest hash changed: {shown}")
            checked += 1
    for shown, expected in extras:
        current = manifest_path(shown)
        require(current.is_file(), f"retained extra path is missing: {shown}")
        require(sha256(current) == expected, f"retained extra hash changed: {shown}")
        checked += 1
    require(checked >= 8, "too few retained manifest paths were checked")
    return checked


def metric_triplet(text: str, field: str) -> tuple[int, int, int]:
    values = tuple(int(part.strip()) for part in text.split("/"))
    require(len(values) == 3, f"report {field} metric is not a triplet: {text}")
    return values


def metric_pair(text: str, field: str) -> tuple[int, int]:
    values = tuple(int(part.strip()) for part in text.split("/"))
    require(len(values) == 2, f"report {field} metric is not a pair: {text}")
    return values


def verify_report(rows: list[dict[str, object]], results: Path, source_manifest_sha: str) -> None:
    report_path = HERE / "report.md"
    report = report_path.read_text()
    binary_sha = (results / "binary.sha256").read_text().split()[0]
    require(f"`{binary_sha}`" in report, "report does not bind the binary SHA-256")
    require(f"`{source_manifest_sha}`" in report, "report does not bind the source manifest SHA-256")
    parsed: dict[tuple[str, str], dict[str, object]] = {}
    for line in report.splitlines():
        parts = [part.strip() for part in line.strip().strip("|").split("|")]
        if len(parts) != 13 or parts[0] not in FIXTURES or parts[1] not in OPERATIONS:
            continue
        key = (parts[0], parts[1])
        require(key not in parsed, f"duplicate pooled report row: {key}")
        elapsed = metric_triplet(parts[7], "elapsed")
        allocation = metric_pair(parts[8], "allocation")
        peak = metric_pair(parts[9], "peak")
        parsed[key] = {
            "processes": int(parts[2]),
            "samples": int(parts[3]),
            "bytes": int(parts[4]),
            "theme_bytes": int(parts[5]),
            "family_bytes": int(parts[6]),
            "p50": elapsed[0],
            "p95": elapsed[1],
            "p99": elapsed[2],
            "alloc50": allocation[0],
            "alloc95": allocation[1],
            "peak50": peak[0],
            "peak95": peak[1],
        }
    expected_keys = {(str(row["fixture"]), str(row["operation"])) for row in rows}
    require(set(parsed) == expected_keys, "pooled report row set differs from verified lanes")
    for row in rows:
        key = (str(row["fixture"]), str(row["operation"]))
        actual = parsed[key]
        for report_field, row_field in (
            ("processes", "processes"),
            ("samples", "samples"),
            ("bytes", "bytes"),
            ("theme_bytes", "theme_bytes"),
            ("family_bytes", "family_bytes"),
            ("p50", "p50_ns"),
            ("p95", "p95_ns"),
            ("p99", "p99_ns"),
            ("alloc50", "requested_alloc_p50"),
            ("alloc95", "requested_alloc_p95"),
            ("peak50", "peak_live_delta_p50"),
            ("peak95", "peak_live_delta_p95"),
        ):
            require(
                actual[report_field] == row[row_field],
                f"pooled report value mismatch for {key}: {report_field}",
            )


def main() -> None:
    results = HERE / "results"
    before = results / "source-manifest-before.txt"
    after = results / "source-manifest-after.txt"
    require(before.read_bytes() == after.read_bytes(), "source manifest changed during build/run")
    source_manifest_sha = sha256(before)
    manifest_paths_checked = verify_source_manifest(before)

    provenance = (results / "source-provenance.txt").read_text()
    require(f"source_manifest_before_sha256={source_manifest_sha}" in provenance, "manifest hash is absent from provenance")
    require(f"source_manifest_after_sha256={source_manifest_sha}" in provenance, "post-build manifest hash is absent from provenance")
    checked_sources = 0
    for expected, name in re.findall(r"^([a-f0-9]{64})  (.+)$", provenance, re.MULTILINE):
        if "crates/" not in name:
            continue
        suffix = name[name.index("crates/") :]
        require(sha256(ROOT / suffix) == expected, f"source hash changed: {suffix}")
        checked_sources += 1
    require(checked_sources >= 8, "too few production source hashes were recorded")

    lanes = 0
    total_samples = 0
    rows = []
    fixture_identities: dict[str, tuple[object, ...]] = {}
    for fixture in FIXTURES:
        for operation in OPERATIONS:
            lane_samples: list[dict[str, object]] = []
            for process in range(1, 4):
                path = results / f"{fixture}-{operation}-p{process}.json"
                value = json.loads(path.read_text())
                require(value["schema"] == "xlsb-theme-family-profile-v1", f"schema: {path}")
                require(value["fixture"] == fixture and value["operation"] == operation, f"lane: {path}")
                require(value["sample_count"] == len(value["samples"]) == 30, f"sample count: {path}")
                expected_foreign_namespaces = 96 if fixture == "opaque" else 0
                require(
                    value["foreign_namespace_declarations"] == expected_foreign_namespaces,
                    f"foreign namespace stress count: {path}",
                )
                identity_fields = (
                    "fixture_path",
                    "package_bytes",
                    "theme_bytes",
                    "family_bytes",
                    "package_digest",
                    "theme_digest",
                    "family_digest",
                    "package_sha256",
                    "theme_sha256",
                    "family_sha256",
                    "foreign_namespace_declarations",
                )
                identity = tuple(value[field] for field in identity_fields)
                if fixture in fixture_identities:
                    require(identity == fixture_identities[fixture], f"fixture identity changed: {path}")
                else:
                    fixture_identities[fixture] = identity
                for key in (
                    "allocator_instrumented",
                    "allocator_self_test",
                    "peak_live_is_incremental_delta",
                    "fixture_shape_ok",
                    "semantic_ok_all",
                    "forward_preservation_ok_all",
                    "preservation_ok_all",
                    "inverse_ok_all",
                    "changed_ok_all",
                    "allocation_balance_all",
                    "source_observation_stable",
                    "source_sharing_gate",
                ):
                    require(value[key] is True, f"{key}: {path}")
                for sample in value["samples"]:
                    direct = int(sample["direct_allocated_bytes"])
                    realloc_new = int(sample["realloc_new_bytes"])
                    realloc_old = int(sample["realloc_old_bytes"])
                    deallocated = int(sample["deallocated_bytes"])
                    live_before = int(sample["live_before"])
                    live_after = int(sample["live_after"])
                    require(
                        int(sample["requested_alloc_bytes"]) == direct + realloc_new,
                        f"requested allocation balance: {path}",
                    )
                    require(
                        live_after == live_before + direct + realloc_new - realloc_old - deallocated,
                        f"live balance: {path}",
                    )
                    require(int(sample["peak_live_delta"]) >= max(0, live_after - live_before), f"peak live: {path}")
                    require(sample["alloc_balance_ok"] is True, f"raw allocation gate: {path}")
                    require(sample["alloc_invalid"] is False and sample["alloc_failed"] == 0, f"allocator error: {path}")
                    require(sample["forward_preservation_ok"] is True, f"forward preservation gate: {path}")
                    if operation in {"family_clone", "noop"}:
                        require(sample["source_shared"] is True, f"source sharing: {path}")
                    if operation == "source_read":
                        require(sample["read_calls"] > 0 and sample["read_returned_bytes"] > 0, f"source I/O: {path}")
                    else:
                        require(
                            sample["read_calls"] == sample["read_requested_bytes"] == sample["read_returned_bytes"] == 0,
                            f"unexpected source I/O: {path}",
                        )
                    lane_samples.append(sample)
            total_samples += len(lane_samples)
            lanes += 1
            process_paths = [results / f"{fixture}-{operation}-p{process}.json" for process in range(1, 4)]
            for path in process_paths:
                process_value = json.loads(path.read_text())
                process_samples = process_value["samples"]
                for percent in (50, 95, 99):
                    require(
                        process_value[f"p{percent}_ns"] == quantile([int(s["elapsed_ns"]) for s in process_samples], percent),
                        f"per-process time percentile: {path}",
                    )
                for prefix, field in (("requested_alloc", "requested_alloc_bytes"), ("peak_live_delta", "peak_live_delta")):
                    for percent in (50, 95):
                        require(
                            process_value[f"{prefix}_p{percent}"] == quantile([int(s[field]) for s in process_samples], percent),
                            f"per-process {prefix} percentile: {path}",
                        )
            representative = json.loads(process_paths[0].read_text())
            rows.append(
                {
                    "fixture": fixture,
                    "operation": operation,
                    "processes": 3,
                    "samples": len(lane_samples),
                    "bytes": int(representative["package_bytes"]),
                    "theme_bytes": int(representative["theme_bytes"]),
                    "family_bytes": int(representative["family_bytes"]),
                    "p50_ns": quantile([int(s["elapsed_ns"]) for s in lane_samples], 50),
                    "p95_ns": quantile([int(s["elapsed_ns"]) for s in lane_samples], 95),
                    "p99_ns": quantile([int(s["elapsed_ns"]) for s in lane_samples], 99),
                    "requested_alloc_p50": quantile([int(s["requested_alloc_bytes"]) for s in lane_samples], 50),
                    "requested_alloc_p95": quantile([int(s["requested_alloc_bytes"]) for s in lane_samples], 95),
                    "peak_live_delta_p50": quantile([int(s["peak_live_delta"]) for s in lane_samples], 50),
                    "peak_live_delta_p95": quantile([int(s["peak_live_delta"]) for s in lane_samples], 95),
                }
            )

    verify_report(rows, results, source_manifest_sha)
    output = {
        "passed": True,
        "source_manifest_sha256": source_manifest_sha,
        "manifest_paths_checked": manifest_paths_checked,
        "source_hashes_checked": checked_sources,
        "lanes": lanes,
        "samples": total_samples,
        "fixture_identities": fixture_identities,
        "rows": rows,
    }
    (HERE / "root-verification.json").write_text(json.dumps(output, indent=2) + "\n")
    print(json.dumps(output, indent=2))


if __name__ == "__main__":
    main()
