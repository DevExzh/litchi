#!/usr/bin/env python3
"""Prepare one immutable DrawingML raw baseline/candidate source pair.

Each raw lane has its own production parent.  The candidate carries the
reviewed protocol adapter; this helper copies only that adapter into a clean
detached baseline and records the exact dirty-path allowlist.  It does not
run Cargo, a native probe, or a timing command.  A later paired runner must
revalidate this receipt before it starts either build or measurement.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any, Mapping


SCHEMA = "litchi.drawingml.raw-preparation.v1"
WORKSPACE_LOCK_SHA256 = "52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6"
WORKSPACE_LOCK_BYTES = 139591
ADAPTER_PATHS = (
    "Cargo.lock",
    "tools/perf-baseline/Cargo.lock",
    "tools/perf-baseline/src/drawingml_raw_probe.rs",
    "tools/perf-baseline/src/bin/drawingml_raw_native.rs",
    "tools/perf-baseline/src/bin/drawingml_raw_allocator.rs",
    "tools/native_qualification/drawingml_raw_validator.py",
    "tools/native_qualification/fixtures/drawingml_raw_fixtures.json",
)

# Filled from the reviewed adapter in this leaf.  Keeping this table in the
# preparer means a future protocol edit cannot silently become historical
# baseline evidence merely because it is copied into a worktree.
STATIC_REVIEWED_ADAPTER_SHA256 = {
    "Cargo.lock": "52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6",
    "tools/perf-baseline/src/drawingml_raw_probe.rs": "7de752332b10cf7b3542efc303c70be193d9872aee043a052f11a3e549084c97",
    "tools/perf-baseline/src/bin/drawingml_raw_native.rs": "faf0d4ecffe07a27b32aa13318bb9e01f22d521351f85f39b6d7bef4b928d392",
    "tools/perf-baseline/src/bin/drawingml_raw_allocator.rs": "7ab157c997f25a8678827d2f4b4384d19cf87ba3ea343ed75a0eb381e79c7d7e",
    "tools/native_qualification/drawingml_raw_validator.py": "cca13390809fe1b1121ad0b2d946d59f9ba78af6a959784590ca23f1f80eb2a4",
    "tools/native_qualification/fixtures/drawingml_raw_fixtures.json": "e72237ec2095535667fa34e962e37950c96ef558eef151839a4e8235a1256925",
}


def _reviewed_adapter_hashes(perf_lock_sha256: str) -> dict[str, str]:
    return {
        **STATIC_REVIEWED_ADAPTER_SHA256,
        "tools/perf-baseline/Cargo.lock": perf_lock_sha256,
    }


class PreparationError(ValueError):
    """A source preparation or receipt invariant failed."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise PreparationError(message)


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: Path) -> tuple[str, int]:
    _expect(not path.is_symlink() and path.is_file(), f"identity path is not a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _git(root: Path, *arguments: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *arguments],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise PreparationError(f"git {' '.join(arguments)} failed in {root}: {result.stderr.strip()}")
    return result.stdout


def _root(value: Path) -> Path:
    root = value.resolve(strict=False)
    _expect(root.is_dir(), f"source root is not a directory: {root}")
    top = Path(_git(root, "rev-parse", "--show-toplevel").strip()).resolve(strict=False)
    _expect(root == top, f"source root is not a worktree top-level: {root}")
    return root


def _status_paths(status: str) -> tuple[str, ...]:
    paths: list[str] = []
    for line in status.splitlines():
        _expect(len(line) >= 4, f"malformed Git status line: {line!r}")
        value = line[3:]
        _expect(" -> " not in value, f"renames are not allowed in a preparation adapter: {line!r}")
        paths.append(value)
    return tuple(sorted(paths))


def _identity(root: Path) -> dict[str, Any]:
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    diff = _git(root, "diff", "--no-ext-diff", "--binary")
    tree = _git(root, "rev-parse", "HEAD^{tree}").strip()
    head = _git(root, "rev-parse", "HEAD").strip()
    hashes = {path: _sha256_file(root / path) for path in ADAPTER_PATHS}
    return {
        "root": str(root),
        "head": head,
        "tree": tree,
        "status_paths": list(_status_paths(status)),
        "status_sha256": _sha256_bytes(status.encode()),
        "diff_sha256": _sha256_bytes(diff.encode()),
        "source_tree_sha256": _sha256_bytes((tree + "\0" + status + "\0" + diff).encode()),
        "adapter_hashes": {
            path: {"sha256": digest, "bytes": size} for path, (digest, size) in hashes.items()
        },
        "workspace_lock_sha256": hashes["Cargo.lock"][0],
        "workspace_lock_bytes": hashes["Cargo.lock"][1],
        "perf_lock_sha256": hashes["tools/perf-baseline/Cargo.lock"][0],
        "perf_lock_bytes": hashes["tools/perf-baseline/Cargo.lock"][1],
    }


def _validate_adapter(identity: Mapping[str, Any], label: str, *, reviewed: bool, perf_lock_sha256: str) -> None:
    statuses = identity.get("status_paths")
    _expect(isinstance(statuses, list), f"{label} status paths are missing")
    _expect(set(statuses).issubset(set(ADAPTER_PATHS)), f"{label} changed outside adapter allowlist: {statuses!r}")
    hashes = identity.get("adapter_hashes")
    _expect(isinstance(hashes, dict) and set(hashes) == set(ADAPTER_PATHS), f"{label} adapter hashes are incomplete")
    _expect(identity.get("workspace_lock_sha256") == WORKSPACE_LOCK_SHA256, f"{label} does not use the historical 52ed workspace lock")
    _expect(identity.get("workspace_lock_bytes") == WORKSPACE_LOCK_BYTES, f"{label} workspace lock byte count differs")
    _expect(identity.get("perf_lock_sha256") == perf_lock_sha256, f"{label} does not use the selected perf-baseline lock")
    if reviewed:
        for path, expected in _reviewed_adapter_hashes(perf_lock_sha256).items():
            observed = hashes[path].get("sha256")
            _expect(observed == expected, f"{label} reviewed adapter hash mismatch: {path}")


def _validate_receipt_data(
    receipt: Mapping[str, Any],
    *,
    baseline_root: Path,
    candidate_root: Path,
    baseline_revision: str,
    candidate_revision: str,
    perf_lock_sha256: str,
) -> dict[str, Any]:
    _expect(receipt.get("schema") == SCHEMA, "preparation receipt schema mismatch")
    _expect(receipt.get("baseline_revision") == baseline_revision, "receipt baseline revision differs")
    _expect(receipt.get("candidate_revision") == candidate_revision, "receipt candidate revision differs")
    _expect(receipt.get("adapter_paths") == list(ADAPTER_PATHS), "adapter path allowlist changed")
    _expect(receipt.get("workspace_lock_sha256") == WORKSPACE_LOCK_SHA256, "receipt workspace lock is not the historic 52ed workspace lock")
    _expect(receipt.get("workspace_lock_bytes") == WORKSPACE_LOCK_BYTES, "receipt workspace lock byte count differs")
    _expect(receipt.get("perf_lock_sha256") == perf_lock_sha256, "receipt perf-baseline lock differs")
    _expect(receipt.get("reviewed_adapter_sha256") == _reviewed_adapter_hashes(perf_lock_sha256), "reviewed adapter hashes changed")
    baseline = receipt.get("baseline")
    candidate = receipt.get("candidate")
    _expect(isinstance(baseline, dict) and isinstance(candidate, dict), "source identities are missing")
    _expect(Path(str(baseline.get("root"))).resolve(strict=False) == baseline_root, "baseline root differs")
    _expect(Path(str(candidate.get("root"))).resolve(strict=False) == candidate_root, "candidate root differs")
    _expect(baseline.get("head") == baseline_revision, "baseline HEAD differs")
    _expect(candidate.get("head") == candidate_revision, "candidate HEAD differs")
    _expect(receipt.get("perf_lock_bytes") == candidate.get("perf_lock_bytes"), "receipt perf-baseline lock byte count differs")
    _validate_adapter(baseline, "baseline", reviewed=True, perf_lock_sha256=perf_lock_sha256)
    _validate_adapter(candidate, "candidate", reviewed=True, perf_lock_sha256=perf_lock_sha256)
    _expect(candidate.get("status_paths") == [], "candidate source must be clean")
    _expect(baseline.get("adapter_hashes") == candidate.get("adapter_hashes"), "adapter bytes differ")
    _expect(receipt.get("baseline_status_paths") == baseline.get("status_paths"), "baseline status receipt differs")
    return dict(receipt)


def validate_preparation_receipt(
    receipt_path: Path,
    *,
    baseline_root: Path,
    candidate_root: Path,
    baseline_revision: str,
    candidate_revision: str,
    perf_lock_sha256: str,
) -> dict[str, Any]:
    receipt_path = receipt_path.resolve(strict=False)
    _expect(not receipt_path.is_symlink() and receipt_path.is_file(), f"receipt is missing: {receipt_path}")
    try:
        receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise PreparationError(f"cannot read preparation receipt: {error}") from error
    _expect(isinstance(receipt, dict), "receipt root is not an object")
    baseline_root = _root(baseline_root)
    candidate_root = _root(candidate_root)
    expected = _validate_receipt_data(
        receipt,
        baseline_root=baseline_root,
        candidate_root=candidate_root,
        baseline_revision=baseline_revision,
        candidate_revision=candidate_revision,
        perf_lock_sha256=perf_lock_sha256,
    )
    for label, current in (("baseline", _identity(baseline_root)), ("candidate", _identity(candidate_root))):
        recorded = expected[label]
        for key in (
            "root",
            "head",
            "tree",
            "status_paths",
            "status_sha256",
            "diff_sha256",
            "source_tree_sha256",
            "adapter_hashes",
            "workspace_lock_sha256",
            "workspace_lock_bytes",
            "perf_lock_sha256",
            "perf_lock_bytes",
        ):
            _expect(current.get(key) == recorded.get(key), f"{label} identity changed: {key}")
    return expected


def prepare(
    baseline_root: Path,
    candidate_root: Path,
    baseline_revision: str,
    candidate_revision: str,
    receipt_path: Path,
    perf_lock_sha256: str,
) -> None:
    baseline_root = _root(baseline_root)
    candidate_root = _root(candidate_root)
    _expect(baseline_root != candidate_root, "baseline and candidate roots must differ")
    _expect(_git(baseline_root, "rev-parse", "HEAD").strip() == baseline_revision, "baseline HEAD differs")
    _expect(_git(candidate_root, "rev-parse", "HEAD").strip() == candidate_revision, "candidate HEAD differs")
    _expect(not _git(candidate_root, "status", "--porcelain=v1", "--untracked-files=all"), "candidate source must be clean")
    before_status = _status_paths(_git(baseline_root, "status", "--porcelain=v1", "--untracked-files=all"))
    _expect(set(before_status).issubset(set(ADAPTER_PATHS)), "baseline already contains an unapproved change")
    candidate = _identity(candidate_root)
    _validate_adapter(candidate, "candidate", reviewed=True, perf_lock_sha256=perf_lock_sha256)
    for relative in ADAPTER_PATHS:
        source = candidate_root / relative
        target = baseline_root / relative
        _expect(not source.is_symlink() and source.is_file(), f"candidate adapter path is missing: {source}")
        _expect(not target.is_symlink(), f"baseline adapter target is a symlink: {target}")
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, target)
    baseline = _identity(baseline_root)
    _validate_adapter(baseline, "baseline", reviewed=True, perf_lock_sha256=perf_lock_sha256)
    _expect(baseline.get("head") == baseline_revision, "baseline HEAD changed during preparation")
    _expect(candidate.get("head") == candidate_revision, "candidate HEAD changed during preparation")
    _expect(baseline.get("adapter_hashes") == candidate.get("adapter_hashes"), "adapter copy differs")
    receipt = {
        "schema": SCHEMA,
        "baseline_revision": baseline_revision,
        "candidate_revision": candidate_revision,
        "adapter_paths": list(ADAPTER_PATHS),
        "baseline_status_paths": list(baseline.get("status_paths", [])),
        "workspace_lock_sha256": WORKSPACE_LOCK_SHA256,
        "workspace_lock_bytes": WORKSPACE_LOCK_BYTES,
        "perf_lock_sha256": perf_lock_sha256,
        "perf_lock_bytes": candidate["perf_lock_bytes"],
        "reviewed_adapter_sha256": _reviewed_adapter_hashes(perf_lock_sha256),
        "baseline": baseline,
        "candidate": candidate,
    }
    receipt_path = receipt_path.resolve(strict=False)
    _expect(not receipt_path.is_symlink() and not receipt_path.exists(), f"refusing to overwrite receipt: {receipt_path}")
    receipt_path.parent.mkdir(parents=True, exist_ok=True)
    with receipt_path.open("x", encoding="utf-8") as stream:
        json.dump(receipt, stream, indent=2, sort_keys=True)
        stream.write("\n")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-root", type=Path, required=True)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--baseline-revision", required=True)
    parser.add_argument("--candidate-revision", required=True)
    parser.add_argument("--perf-lock-sha256", required=True)
    parser.add_argument("--output", type=Path, required=True)
    arguments = parser.parse_args(argv)
    try:
        prepare(
            arguments.baseline_root,
            arguments.candidate_root,
            arguments.baseline_revision,
            arguments.candidate_revision,
            arguments.output,
            arguments.perf_lock_sha256,
        )
    except (OSError, PreparationError) as error:
        parser.error(str(error))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
