#!/usr/bin/env python3
"""Prepare and validate the source pair for the XLSX raw ABBA lane.

The historical source is a real detached Git worktree at the original 0bef
production revision.  Only the reviewed consumer-side probe adapter is
copied from the candidate into that worktree.  The receipt records both Git
identities, every allowed adapter/lock/probe hash, and the resulting dirty
diff so the build and ABBA phases can reject any later substitution.
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


SCHEMA = "litchi.xlsx.raw-preparation.v1"
BASELINE_REVISION = "0befb112d96d2409e6f40c9abfeb4813665905ad"
LOCK_PATHS = ("tools/perf-baseline/unified-xlsx-raw-identity/Cargo.lock",)
PROBE_PATHS = (
    "tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml",
    "tools/perf-baseline/unified-xlsx-raw-identity/src/lib.rs",
    "tools/perf-baseline/unified-xlsx-raw-identity/src/bin/xlsx_raw_identity_alloc.rs",
    "tools/perf-baseline/unified-xlsx-raw-identity/src/bin/xlsx_raw_identity_native.rs",
    "tools/perf-baseline/src/allocation_metrics.rs",
    "tools/perf-baseline/src/xlsx_raw_identity.rs",
    "tools/perf-baseline/src/bin/support/counting_allocator.rs",
)
ADAPTER_PATHS = tuple((*LOCK_PATHS, *PROBE_PATHS))
# These are the reviewed consumer-side inputs used by the historical and
# candidate probes. Equality between the two worktrees is insufficient: a
# pair of equally modified inputs would otherwise pass the preparation gate.
REVIEWED_ADAPTER_SHA256 = {
    "tools/perf-baseline/unified-xlsx-raw-identity/Cargo.lock": "8b61b3c5e17b1c32d89139c97a26145117794afae83dd8fef3c9a7aff8765ee6",
    "tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml": "30ff248844c9e21436fc3913baa4c4d363daffdc617e5e7d03404423410d5c70",
    "tools/perf-baseline/unified-xlsx-raw-identity/src/lib.rs": "d65ce0630a50d1b972a0b2f93044091d1a5c50df1e9c8ef06e9c3ba65cf8a36a",
    "tools/perf-baseline/unified-xlsx-raw-identity/src/bin/xlsx_raw_identity_alloc.rs": "c21cc9b947539078885ff099527ed8d291d631b39a27bf8a299c6553add9007c",
    "tools/perf-baseline/unified-xlsx-raw-identity/src/bin/xlsx_raw_identity_native.rs": "71435d535e65996f2c62e9952ab3f595b8e3bd2342afc0f45092e8a407fc9ee9",
    "tools/perf-baseline/src/allocation_metrics.rs": "0fcec3e5c4972031b23542fd359a38401e25f32a133c77ef4c6e75facf0710d5",
    "tools/perf-baseline/src/xlsx_raw_identity.rs": "a251f1199b5ba442dfd9664806b81f29c7832d071df66fc2d2d5f2ada6c97fd4",
    "tools/perf-baseline/src/bin/support/counting_allocator.rs": "80977e55c2213eeee2af9a3133295d70cff38505e663a68e4757bf7939c60068",
}


class PreparationError(ValueError):
    """A source preparation or receipt invariant failed."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise PreparationError(message)


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: Path) -> tuple[str, int]:
    _expect(not path.is_symlink(), f"identity path may not be a symlink: {path}")
    _expect(path.is_file(), f"identity path is missing: {path}")
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


def _root(root: Path) -> Path:
    root = root.resolve(strict=False)
    _expect(root.is_dir(), f"source root is not a directory: {root}")
    top = Path(_git(root, "rev-parse", "--show-toplevel").strip()).resolve(strict=False)
    _expect(top == root, f"source root is not the Git worktree top-level: {root}")
    return root


def _status_paths(status: str) -> tuple[str, ...]:
    paths: list[str] = []
    for line in status.splitlines():
        _expect(len(line) >= 4, f"malformed Git status line: {line!r}")
        value = line[3:]
        _expect(" -> " not in value, f"renames are not allowed in adapter status: {line!r}")
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
        "status_sha256": _sha256_bytes(status.encode("utf-8")),
        "diff_sha256": _sha256_bytes(diff.encode("utf-8")),
        "source_tree_sha256": _sha256_bytes((tree + "\0" + status + "\0" + diff).encode("utf-8")),
        "adapter_hashes": {
            path: {"sha256": digest, "bytes": size} for path, (digest, size) in hashes.items()
        },
        "lock_hashes": {path: hashes[path][0] for path in LOCK_PATHS},
        "probe_hashes": {path: hashes[path][0] for path in PROBE_PATHS},
    }


def _validate_adapter_paths(
    identity: Mapping[str, Any], *, label: str, require_reviewed_hashes: bool = False
) -> None:
    statuses = identity.get("status_paths")
    _expect(isinstance(statuses, list), f"{label} status paths are missing")
    _expect(
        set(statuses).issubset(set(ADAPTER_PATHS)),
        f"{label} contains changes outside the reviewed adapter: {statuses!r}",
    )
    hashes = identity.get("adapter_hashes")
    _expect(isinstance(hashes, dict) and set(hashes) == set(ADAPTER_PATHS), f"{label} adapter hashes are incomplete")
    if require_reviewed_hashes:
        for path, expected in REVIEWED_ADAPTER_SHA256.items():
            value = hashes.get(path)
            _expect(
                isinstance(value, dict) and value.get("sha256") == expected,
                f"{label} reviewed adapter hash mismatch: {path}",
            )


def _validate_receipt_data(
    receipt: Mapping[str, Any],
    *,
    baseline_root: Path,
    candidate_root: Path,
    candidate_revision: str,
) -> dict[str, Any]:
    _expect(receipt.get("schema") == SCHEMA, "preparation receipt schema mismatch")
    _expect(receipt.get("baseline_revision") == BASELINE_REVISION, "receipt baseline is not original 0bef")
    _expect(receipt.get("candidate_revision") == candidate_revision, "receipt candidate revision mismatch")
    _expect(receipt.get("adapter_paths") == list(ADAPTER_PATHS), "receipt adapter path allowlist changed")
    baseline = receipt.get("baseline")
    candidate = receipt.get("candidate")
    _expect(isinstance(baseline, dict) and isinstance(candidate, dict), "receipt source identities are missing")
    _expect(Path(str(baseline.get("root"))).resolve(strict=False) == baseline_root, "receipt baseline root mismatch")
    _expect(Path(str(candidate.get("root"))).resolve(strict=False) == candidate_root, "receipt candidate root mismatch")
    for label, identity in (("baseline", baseline), ("candidate", candidate)):
        _validate_adapter_paths(identity, label=label, require_reviewed_hashes=True)
    _expect(baseline.get("head") == BASELINE_REVISION, "receipt baseline HEAD is not original 0bef")
    _expect(candidate.get("head") == candidate_revision, "receipt candidate HEAD is not the selected revision")
    _expect(candidate.get("status_paths") == [], "candidate source must be clean")
    _expect(
        set(baseline.get("status_paths", [])).issubset(set(ADAPTER_PATHS)),
        "baseline status contains an unapproved path",
    )
    _expect(receipt.get("allowed_adapter_diff") == list(ADAPTER_PATHS), "receipt allowed diff is not exact")
    _expect(
        receipt.get("reviewed_adapter_sha256") == REVIEWED_ADAPTER_SHA256,
        "receipt reviewed adapter hashes are not the fixed reviewed inputs",
    )
    _expect(
        receipt.get("adapter_diff_paths") == baseline.get("status_paths"),
        "receipt adapter diff paths differ from the recorded baseline worktree",
    )
    baseline_hashes = baseline.get("adapter_hashes", {})
    candidate_hashes = candidate.get("adapter_hashes", {})
    _expect(baseline_hashes == candidate_hashes, "baseline adapter bytes differ from reviewed candidate bytes")
    _expect(receipt.get("lock_paths") == list(LOCK_PATHS), "receipt lock path set changed")
    _expect(receipt.get("probe_paths") == list(PROBE_PATHS), "receipt probe path set changed")
    return dict(receipt)


def validate_preparation_receipt(
    receipt_path: Path,
    *,
    baseline_root: Path,
    candidate_root: Path,
    candidate_revision: str,
) -> dict[str, Any]:
    """Re-read a receipt and compare it with both current source worktrees."""

    receipt_path = receipt_path.resolve(strict=False)
    _expect(not receipt_path.is_symlink() and receipt_path.is_file(), f"preparation receipt is missing: {receipt_path}")
    try:
        receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise PreparationError(f"cannot read preparation receipt: {error}") from error
    _expect(isinstance(receipt, dict), "preparation receipt root is not an object")
    baseline_root = _root(baseline_root)
    candidate_root = _root(candidate_root)
    current_baseline = _identity(baseline_root)
    current_candidate = _identity(candidate_root)
    result = _validate_receipt_data(
        receipt,
        baseline_root=baseline_root,
        candidate_root=candidate_root,
        candidate_revision=candidate_revision,
    )
    for label, current in (("baseline", current_baseline), ("candidate", current_candidate)):
        recorded = result[label]
        for key in (
            "root",
            "head",
            "tree",
            "status_paths",
            "status_sha256",
            "diff_sha256",
            "source_tree_sha256",
            "adapter_hashes",
            "lock_hashes",
            "probe_hashes",
        ):
            _expect(current.get(key) == recorded.get(key), f"{label} preparation identity changed: {key}")
    return result


def prepare(
    baseline_root: Path,
    candidate_root: Path,
    candidate_revision: str,
    receipt_path: Path,
) -> None:
    """Copy only the reviewed adapter into a detached 0bef worktree."""

    _expect(
        len(candidate_revision) == 40 and all(char in "0123456789abcdefABCDEF" for char in candidate_revision),
        "candidate revision must be a 40-character hexadecimal revision",
    )
    _expect(candidate_revision != BASELINE_REVISION, "candidate revision must differ from original 0bef")
    baseline_root = _root(baseline_root)
    candidate_root = _root(candidate_root)
    _expect(baseline_root != candidate_root, "baseline and candidate roots must differ")
    _expect(_git(baseline_root, "rev-parse", "HEAD").strip() == BASELINE_REVISION, "baseline worktree must be detached at original 0bef")
    _expect(_git(candidate_root, "rev-parse", "HEAD").strip() == candidate_revision, "candidate worktree HEAD differs from selected revision")
    _expect(not _git(candidate_root, "status", "--porcelain=v1", "--untracked-files=all"), "candidate source must be clean")
    before_status = _status_paths(_git(baseline_root, "status", "--porcelain=v1", "--untracked-files=all"))
    _expect(set(before_status).issubset(set(ADAPTER_PATHS)), "baseline has unapproved pre-existing changes")
    for relative in ADAPTER_PATHS:
        source = candidate_root / relative
        target = baseline_root / relative
        _expect(not source.is_symlink() and source.is_file(), f"candidate adapter path is missing: {source}")
        _expect(not target.is_symlink(), f"baseline adapter path may not be a symlink: {target}")
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, target)
    baseline = _identity(baseline_root)
    candidate = _identity(candidate_root)
    _validate_adapter_paths(candidate, label="candidate", require_reviewed_hashes=True)
    _validate_adapter_paths(baseline, label="baseline", require_reviewed_hashes=True)
    _expect(baseline["head"] == BASELINE_REVISION, "baseline HEAD changed during preparation")
    _expect(candidate["head"] == candidate_revision, "candidate HEAD changed during preparation")
    _expect(candidate["status_paths"] == [], "candidate became dirty during preparation")
    _expect(baseline["adapter_hashes"] == candidate["adapter_hashes"], "adapter copy does not match candidate")
    _expect(set(baseline["status_paths"]).issubset(set(ADAPTER_PATHS)), "baseline adapter diff escaped its allowlist")
    receipt = {
        "schema": SCHEMA,
        "baseline_revision": BASELINE_REVISION,
        "candidate_revision": candidate_revision,
        "adapter_paths": list(ADAPTER_PATHS),
        "allowed_adapter_diff": list(ADAPTER_PATHS),
        "reviewed_adapter_sha256": dict(REVIEWED_ADAPTER_SHA256),
        "adapter_diff_paths": list(baseline["status_paths"]),
        "lock_paths": list(LOCK_PATHS),
        "probe_paths": list(PROBE_PATHS),
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
    parser.add_argument("--candidate-revision", required=True)
    parser.add_argument("--output", type=Path, required=True)
    arguments = parser.parse_args(argv)
    try:
        prepare(
            arguments.baseline_root,
            arguments.candidate_root,
            arguments.candidate_revision,
            arguments.output,
        )
    except (OSError, PreparationError) as error:
        parser.error(str(error))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
