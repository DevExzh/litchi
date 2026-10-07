#!/usr/bin/env python3
"""Describe 0960 external snapshots and receipts without copying large evidence.

This is a provenance helper only. It never archives, deletes, or executes a
workload. It refuses to overwrite an output. Directory entries are censused
without following symlinks; source snapshot entries retain their historical
manifest comparison so a post-quality edit cannot be presented as the frozen
snapshot.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
from pathlib import Path
from typing import Any

EXT = Path("/home/zhuhe/litchi-performance-0960")
OLD = Path("/home/zhuhe/litchi-performance-0957")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_descriptor(path: Path, role: str, status: str, command: str) -> dict[str, Any]:
    st = path.lstat()
    if not stat.S_ISREG(st.st_mode):
        raise SystemExit(f"expected regular file: {path}")
    return {
        "kind": "file",
        "path": str(path),
        "role": role,
        "status": status,
        "bytes": st.st_size,
        "mode": stat.S_IMODE(st.st_mode),
        "sha256": sha256_file(path),
        "regeneration": command,
    }


def directory_descriptor(
    path: Path,
    role: str,
    status: str,
    command: str,
    historical: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if not path.is_dir() or path.is_symlink():
        raise SystemExit(f"expected non-symlink directory: {path}")
    files = 0
    bytes_total = 0
    symlinks = 0
    for root, dirs, names in os.walk(path, followlinks=False):
        root_path = Path(root)
        for name in names:
            candidate = root_path / name
            if candidate.is_symlink():
                symlinks += 1
                continue
            if not candidate.is_file():
                raise SystemExit(f"non-regular entry in snapshot: {candidate}")
            files += 1
            bytes_total += candidate.stat().st_size
        for name in dirs:
            if (root_path / name).is_symlink():
                symlinks += 1
    result: dict[str, Any] = {
        "kind": "directory_census",
        "path": str(path),
        "role": role,
        "status": status,
        "regular_file_count": files,
        "regular_bytes": bytes_total,
        "symlink_count": symlinks,
        "regeneration": command,
    }
    if historical is not None:
        result["historical_manifest"] = historical
        result["historical_match"] = {
            "file_count": files == historical["file_count"],
            "bytes": bytes_total == historical["bytes_in_snapshot"],
            "classification": "matches_historical_manifest"
            if files == historical["file_count"] and bytes_total == historical["bytes_in_snapshot"]
            else "post_snapshot_tree_differs_from_historical_manifest",
        }
    return result


def historical_manifest(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text())
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
        "file_count": data["file_count"],
        "bytes_in_snapshot": data["bytes"],
        "source_commit": data.get("source_commit"),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.output.exists():
        raise SystemExit(f"refusing to overwrite: {args.output}")

    historical_source_cmd = (
        "git -C /home/zhuhe/code/litchi archive --format=tar "
        "a6223afa0899d5800d21ed9dc6c49304fbb2e8ca -- . "
        ":(exclude)docs/performance/** :(exclude)docs/report/**"
    )
    source_cmd = (
        "1) git archive the base commit with docs/performance and docs/report excluded; "
        "2) copy the recorded root Cargo.lock overlay into both fresh roots; "
        "3) apply receipts/0960-final-codec-v1.patch to the candidate root; "
        "4) emit sorted path,size,mode,SHA-256 rows for each root"
    )
    entries: list[dict[str, Any]] = []
    old_baseline_manifest = EXT / "source-manifest-baseline-v1.json"
    old_candidate_manifest = EXT / "source-manifest-candidate-v1.json"
    baseline_manifest = EXT / "source-manifest-baseline-v2.json"
    candidate_manifest = EXT / "source-manifest-candidate-v2.json"
    entries.append(file_descriptor(old_baseline_manifest, "pre-quality baseline source manifest", "external_historical", historical_source_cmd))
    entries.append(file_descriptor(old_candidate_manifest, "pre-quality candidate source manifest", "external_historical", historical_source_cmd))
    entries.append(file_descriptor(baseline_manifest, "quality-frozen baseline source manifest", "external_snapshot", source_cmd))
    entries.append(file_descriptor(candidate_manifest, "quality-frozen candidate source manifest", "external_snapshot", source_cmd))
    entries.append(directory_descriptor(
        EXT / "source-baseline-v1", "quality-frozen baseline source snapshot", "external_snapshot",
        source_cmd, historical_manifest(baseline_manifest)))
    entries.append(directory_descriptor(
        EXT / "source-candidate-v1", "quality-frozen candidate source snapshot", "external_snapshot",
        source_cmd, historical_manifest(candidate_manifest)))

    retained_manifest = OLD / "source-manifest-v2.json"
    entries.append(file_descriptor(retained_manifest, "0957 retained runtime source manifest dependency", "external_dependency", "git -C /home/zhuhe/code/litchi archive --format=tar 02449f394529766d76794332204c27c4022931fc"))
    for binary in (
        OLD / "executables/xlsx_managed_compression_alloc.v2",
        OLD / "executables/xlsx_managed_compression_benchmark.v2",
    ):
        entries.append(file_descriptor(binary, "0957 retained binary dependency; execution not performed in 0960", "external_dependency", "rebuild from the retained 0957 source/build receipt; do not commit executable"))

    for binary in (
        EXT / "executables/xlsx_managed_compression_alloc.v1",
        EXT / "executables/xlsx_managed_compression_benchmark.v1",
    ):
        entries.append(file_descriptor(binary, "0960 candidate release executable; external only", "external_candidate_binary", "rebuild with the command and source manifest recorded in receipts/candidate-managed-v1-build-start.json; do not commit executable"))
    for log in (
        EXT / "builds/candidate-managed-v1/build.stdout.log",
        EXT / "builds/candidate-managed-v1/build.stderr.log",
    ):
        entries.append(file_descriptor(log, "0960 candidate build log", "external_build_log", "re-run the command recorded in receipts/candidate-managed-v1-build-start.json"))

    entries.append(file_descriptor(
        EXT / "receipts/0960-final-codec-v1.patch",
        "candidate source patch applied after the base export",
        "external_candidate_source",
        "git -C /home/zhuhe/litchi-performance-0960/source-candidate-v1 apply /home/zhuhe/litchi-performance-0960/receipts/0960-final-codec-v1.patch",
    ))

    final_logs = [
        ("logs/0960-full-xlsx-lib-v2.log", "full XLSX lib test log; observed 1,404 passed", "quality_receipt"),
        ("logs/0960-clippy-xlsx-lib-v2.log", "Clippy v2 log; observed rc=0 on corrected codec hash", "quality_receipt"),
        ("logs/0960-focused-strict-hint-v7.log", "focused strict-lane v7 log; observed 7 passed/1 ignored", "quality_receipt"),
        ("logs/0960-managed-route-census-v1.log", "managed route census log; eight worksheets/27,008 cells", "quality_receipt"),
        ("logs/0960-rustfmt-codec-v1.log", "scoped codec rustfmt log; observed rc=0", "quality_receipt"),
    ]
    historical_logs = [
        ("logs/0960-clippy-xlsx-lib-v1.log", "superseded Clippy failure history", "historical_failure"),
        ("logs/0960-focused-strict-hint-v1.log", "superseded focused failure history", "historical_failure"),
        ("logs/0960-focused-strict-hint-v2.log", "superseded focused compile failure history", "historical_failure"),
        ("logs/0960-focused-strict-hint-v3.log", "superseded focused failure history", "historical_failure"),
        ("logs/0960-focused-strict-hint-v4.log", "superseded focused failure history", "historical_failure"),
        ("logs/0960-focused-strict-hint-v5.log", "superseded focused failure history", "historical_failure"),
    ]
    for rel, role, status in final_logs + historical_logs:
        path = EXT / rel
        entries.append(file_descriptor(path, role, status, f"re-run the command recorded in receipts/{Path(rel).stem}.command"))
    entries.append(directory_descriptor(
        EXT / "route-census-v1", "external eight-worksheet route census inputs", "external_quality_input",
        "recreate the two medium/dense worksheet trees from the route-census input generator; verify receipts/0960-route-census-worksheets.sha256"))

    output = {
        "schema": "litchi.performance.0960.external-artifact-manifest.v1",
        "status": "native_capture_running_allocator_pending",
        "generated_by": "tools/build_external_artifact_manifest_0960_v1.py",
        "scope": "quality/provenance packet; no 0960 benchmark workload or formal capture",
        "large_artifacts_stay_external": True,
        "historical_failures_preserved_by_descriptor": True,
        "entries": entries,
        "notes": [
            "The copied quality receipts are bounded metadata; full logs, source trees, source manifests, route inputs, and binaries remain external.",
            "The v1 source manifests are retained as pre-quality history; v2 manifests and the patch bind the current quality-frozen trees.",
            "The candidate build and authorized measurement plan are separate gates; this manifest records source/quality readiness only.",
        ],
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(output, indent=2, sort_keys=True) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
