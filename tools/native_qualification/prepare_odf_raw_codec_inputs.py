#!/usr/bin/env python3
"""Prepare the pristine ODF raw-codec baseline overlay.

This is source preparation only.  It never invokes Cargo, Rust, a probe, or
timing code.  The baseline is the exact pre-migration commit; its only
overlay is the reviewed consumer adapter and a separately hashed lockfile.
Production files are never copied from the candidate.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
from typing import Any


SCHEMA = "litchi.odf.raw-codec-qualification.preparation.v1"
LANE = "odf-raw-codec"
BASELINE_COMMIT = "70d64634c9d1d4328447d8724d3b076114a91fd0"
HISTORICAL_CANDIDATE = "15ab9324ef4f8c927acd6ecfc19d9707c3f886ee"
PROBE_PATH = Path("tools/perf-baseline/examples/odf_raw_codec_qualification.rs")
CONSUMER_MANIFEST = Path("tools/perf-baseline/Cargo.toml")
LOCK_PATH = Path("tools/perf-baseline/Cargo.lock")
FULL_SHA = re.compile(r"^[0-9a-f]{40}$")
HASH = re.compile(r"^[0-9a-f]{64}$")
UPSTREAM_COMPONENT = "3rd" + "party"


class PreparationError(RuntimeError):
    """A source or overlay identity is not admissible."""


def _resolve(path: str | os.PathLike[str], *, base: Path | None = None) -> Path:
    candidate = Path(path)
    if not candidate.is_absolute():
        if base is None:
            raise PreparationError(f"path must be absolute: {path!r}")
        candidate = base / candidate
    probe = Path(candidate.anchor)
    for part in candidate.parts[1:]:
        probe /= part
        if probe.is_symlink():
            raise PreparationError(f"symlink path component is forbidden: {candidate}")
    resolved = candidate.expanduser().resolve(strict=False)
    if UPSTREAM_COMPONENT in resolved.parts:
        raise PreparationError(f"untracked upstream checkout path is forbidden: {resolved}")
    return resolved


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _sha256(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise PreparationError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise PreparationError(f"git {' '.join(args)} failed in {root}: {result.stderr.strip()}")
    return result.stdout


def _source_state(root: Path) -> dict[str, Any]:
    root = _resolve(root)
    top = _resolve(_git(root, "rev-parse", "--show-toplevel").strip())
    if top != root:
        raise PreparationError(f"source is not a worktree top-level: {root}")
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    sparse_result = subprocess.run(
        ["git", "-C", str(root), "config", "--get", "core.sparseCheckout"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if sparse_result.returncode not in (0, 1):
        raise PreparationError(f"cannot inspect sparse-checkout configuration: {sparse_result.stderr.strip()}")
    sparse_config = sparse_result.stdout.strip()
    tracked_flags = _git(root, "ls-files", "-v").splitlines()
    skip_worktree_count = sum(line.startswith("S") for line in tracked_flags)
    return {
        "root": str(root),
        "head": _git(root, "rev-parse", "HEAD").strip(),
        "tree": _git(root, "rev-parse", "HEAD^{tree}").strip(),
        "status": status,
        "status_sha256": hashlib.sha256(status.encode()).hexdigest(),
        "clean": not bool(status),
        "sparse_checkout": sparse_config,
        "skip_worktree_count": skip_worktree_count,
    }


def _production_tree(root: Path) -> str:
    return hashlib.sha256(_git(root, "ls-tree", "-r", "--full-tree", "HEAD").encode()).hexdigest()


def _status_paths(status: str) -> list[str]:
    paths = []
    for line in status.splitlines():
        if len(line) < 4:
            raise PreparationError(f"malformed git status line: {line!r}")
        path = line[3:]
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        paths.append(path)
    return paths


def _require_sha(value: str, label: str) -> str:
    if not FULL_SHA.fullmatch(value):
        raise PreparationError(f"{label} must be a full lowercase commit SHA")
    return value


def _require_hash(value: str, label: str) -> str:
    if not HASH.fullmatch(value):
        raise PreparationError(f"{label} must be a full lowercase SHA-256")
    return value


def prepare(
    *,
    baseline_root: Path,
    candidate_root: Path,
    candidate_ref: str,
    candidate_lock_sha256: str,
    baseline_lock: Path,
    output: Path,
) -> dict[str, Any]:
    baseline_root = _resolve(baseline_root)
    candidate_root = _resolve(candidate_root)
    baseline = _source_state(baseline_root)
    candidate = _source_state(candidate_root)
    _require_sha(candidate_ref, "candidate_ref")
    if candidate_ref == HISTORICAL_CANDIDATE:
        raise PreparationError("historical 15ab candidate is not admissible")
    _require_hash(candidate_lock_sha256, "candidate_lock_sha256")
    if baseline["head"] != BASELINE_COMMIT:
        raise PreparationError(f"baseline HEAD must be {BASELINE_COMMIT}")
    if not baseline["clean"]:
        raise PreparationError("baseline must be clean before overlay")
    if not candidate["clean"] or candidate["head"] != candidate_ref:
        raise PreparationError("candidate must be clean and at candidate_ref")
    if baseline_root == candidate_root or _under(baseline_root, candidate_root) or _under(candidate_root, baseline_root):
        raise PreparationError("baseline and candidate worktrees must be disjoint")
    if baseline["skip_worktree_count"] or candidate["skip_worktree_count"]:
        raise PreparationError("sparse/skip-worktree source is not admissible for preparation")

    candidate_probe = _resolve(PROBE_PATH, base=candidate_root)
    candidate_consumer = _resolve(CONSUMER_MANIFEST, base=candidate_root)
    candidate_lock = _resolve(LOCK_PATH, base=candidate_root)
    baseline_probe = _resolve(PROBE_PATH, base=baseline_root)
    baseline_consumer = _resolve(CONSUMER_MANIFEST, base=baseline_root)
    baseline_lock_destination = _resolve(LOCK_PATH, base=baseline_root)
    external_lock = _resolve(baseline_lock)
    candidate_probe_sha, candidate_probe_bytes = _sha256(candidate_probe)
    candidate_consumer_sha, candidate_consumer_bytes = _sha256(candidate_consumer)
    observed_candidate_lock_sha, candidate_lock_bytes = _sha256(candidate_lock)
    if observed_candidate_lock_sha != candidate_lock_sha256:
        raise PreparationError("candidate Cargo.lock hash does not match supplied identity")
    baseline_lock_sha, baseline_lock_bytes = _sha256(external_lock)
    if baseline_probe.exists() or baseline_probe.is_symlink():
        raise PreparationError(f"baseline probe destination must be absent: {baseline_probe}")
    baseline_probe.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(candidate_probe, baseline_probe)
    if baseline_consumer.is_symlink():
        raise PreparationError("baseline consumer manifest destination may not be a symlink")
    baseline_consumer.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(candidate_consumer, baseline_consumer)
    if baseline_lock_destination.is_symlink():
        raise PreparationError("baseline Cargo.lock destination may not be a symlink")
    baseline_lock_destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(external_lock, baseline_lock_destination)
    prepared_probe_sha, prepared_probe_bytes = _sha256(baseline_probe)
    prepared_consumer_sha, prepared_consumer_bytes = _sha256(baseline_consumer)
    prepared_lock_sha, prepared_lock_size = _sha256(baseline_lock_destination)
    if (prepared_probe_sha, prepared_probe_bytes) != (candidate_probe_sha, candidate_probe_bytes):
        raise PreparationError("baseline adapter differs from candidate probe")
    if (prepared_consumer_sha, prepared_consumer_bytes) != (candidate_consumer_sha, candidate_consumer_bytes):
        raise PreparationError("baseline consumer manifest differs from candidate")
    if (prepared_lock_sha, prepared_lock_size) != (baseline_lock_sha, baseline_lock_bytes):
        raise PreparationError("baseline Cargo.lock differs from pinned external lock")

    after = _source_state(baseline_root)
    allowed_status = [str(CONSUMER_MANIFEST), str(PROBE_PATH)]
    if str(LOCK_PATH) in _status_paths(after["status"]):
        allowed_status.append(str(LOCK_PATH))
    if sorted(_status_paths(after["status"])) != sorted(allowed_status):
        raise PreparationError(f"baseline overlay changed unexpected paths: {_status_paths(after['status'])!r}")
    if after["tree"] != baseline["tree"]:
        raise PreparationError("baseline tracked production tree changed during overlay")
    output = _resolve(output)
    if _under(output, baseline_root) or _under(output, candidate_root):
        raise PreparationError("preparation output must be outside both source worktrees")
    if output.exists():
        raise PreparationError(f"preparation output already exists: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    preparation = {
        "schema": SCHEMA,
        "lane": LANE,
        "baseline": {
            "before": baseline,
            "after_overlay": after,
            "commit": BASELINE_COMMIT,
            "production_tree_sha256": _production_tree(baseline_root),
            "overlay_paths": allowed_status,
            "probe": {"source_path": str(candidate_probe), "destination_path": str(baseline_probe), "sha256": prepared_probe_sha, "bytes": prepared_probe_bytes},
            "consumer_manifest": {"source_path": str(candidate_consumer), "destination_path": str(baseline_consumer), "sha256": prepared_consumer_sha, "bytes": prepared_consumer_bytes},
            "lock": {"source_path": str(external_lock), "destination_path": str(baseline_lock_destination), "sha256": prepared_lock_sha, "bytes": prepared_lock_size},
        },
        "candidate": {
            "root": str(candidate_root),
            "commit": candidate_ref,
            "consumer_manifest": {"path": str(candidate_consumer), "sha256": candidate_consumer_sha, "bytes": candidate_consumer_bytes},
            "probe": {"path": str(candidate_probe), "sha256": candidate_probe_sha, "bytes": candidate_probe_bytes},
            "lock": {"path": str(candidate_lock), "sha256": observed_candidate_lock_sha, "bytes": candidate_lock_bytes},
        },
        "claim_scope": "source preparation identity only; no build, timing, or acceptance claim",
    }
    output.write_text(json.dumps(preparation, indent=2) + "\n", encoding="utf-8")
    return preparation


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-root", type=Path, required=True)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--candidate-ref", required=True)
    parser.add_argument("--candidate-lock-sha256", required=True)
    parser.add_argument("--baseline-lock", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        prepare(
            baseline_root=args.baseline_root,
            candidate_root=args.candidate_root,
            candidate_ref=args.candidate_ref,
            candidate_lock_sha256=args.candidate_lock_sha256,
            baseline_lock=args.baseline_lock,
            output=args.output,
        )
        return 0
    except (OSError, PreparationError) as error:
        print(f"error: {error}", file=os.sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
