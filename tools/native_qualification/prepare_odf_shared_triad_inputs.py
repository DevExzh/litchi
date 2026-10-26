#!/usr/bin/env python3
"""Prepare the pristine ODF shared-triad baseline overlay.

This command is source preparation only.  It never invokes Cargo, Rust, a
probe, or a timing command.  The baseline is the exact 173dd tree; the only
permitted overlay is the separately hashed reviewed probe adapter plus an
external, separately hashed Cargo.lock.  Candidate production files are
never copied into the baseline.
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


SCHEMA = "litchi.odf.shared-triad-qualification.preparation.v1"
BASELINE_COMMIT = "173dd5330476713714871007316794b9a3a1bdaa"
HISTORICAL_CANDIDATE = "15ab9324ef4f8c927acd6ecfc19d9707c3f886ee"
PROBE_PATH = Path("crates/litchi-ops/examples/odf_shared_triad_qualification.rs")
LOCK_PATH = Path("Cargo.lock")
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
    return {
        "root": str(root),
        "head": _git(root, "rev-parse", "HEAD").strip(),
        "tree": _git(root, "rev-parse", "HEAD^{tree}").strip(),
        "status": status,
        "status_sha256": hashlib.sha256(status.encode()).hexdigest(),
        "clean": not bool(status),
    }


def _production_tree(root: Path) -> str:
    return hashlib.sha256(
        _git(root, "ls-tree", "-r", "--full-tree", "HEAD").encode()
    ).hexdigest()


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

    candidate_probe = _resolve(PROBE_PATH, base=candidate_root)
    candidate_lock = _resolve(LOCK_PATH, base=candidate_root)
    baseline_probe = _resolve(PROBE_PATH, base=baseline_root)
    baseline_lock_destination = _resolve(LOCK_PATH, base=baseline_root)
    external_lock = _resolve(baseline_lock)
    candidate_probe_sha, candidate_probe_bytes = _sha256(candidate_probe)
    observed_candidate_lock_sha, candidate_lock_bytes = _sha256(candidate_lock)
    if observed_candidate_lock_sha != candidate_lock_sha256:
        raise PreparationError("candidate Cargo.lock hash does not match supplied identity")
    baseline_lock_sha, baseline_lock_bytes = _sha256(external_lock)
    if baseline_probe.exists() or baseline_probe.is_symlink():
        raise PreparationError(f"baseline probe destination must be absent: {baseline_probe}")
    baseline_probe.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(candidate_probe, baseline_probe)
    if baseline_lock_destination.exists() and not baseline_lock_destination.is_symlink():
        existing_sha, _ = _sha256(baseline_lock_destination)
        if existing_sha != baseline_lock_sha:
            raise PreparationError("baseline Cargo.lock already exists with a different hash")
    else:
        if baseline_lock_destination.is_symlink():
            raise PreparationError("baseline Cargo.lock destination may not be a symlink")
        baseline_lock_destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(external_lock, baseline_lock_destination)
    prepared_probe_sha, prepared_probe_bytes = _sha256(baseline_probe)
    prepared_lock_sha, prepared_lock_size = _sha256(baseline_lock_destination)
    if (prepared_probe_sha, prepared_probe_bytes) != (candidate_probe_sha, candidate_probe_bytes):
        raise PreparationError("baseline adapter differs from candidate probe")
    if (prepared_lock_sha, prepared_lock_size) != (baseline_lock_sha, baseline_lock_bytes):
        raise PreparationError("baseline Cargo.lock differs from pinned external lock")

    after = _source_state(baseline_root)
    allowed_status = [str(PROBE_PATH)]
    if "Cargo.lock" in _status_paths(after["status"]):
        allowed_status.append(str(LOCK_PATH))
    if sorted(_status_paths(after["status"])) != sorted(allowed_status):
        raise PreparationError(
            f"baseline overlay changed unexpected paths: {_status_paths(after['status'])!r}"
        )
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
        "lane": "odf-shared-triad",
        "baseline": {
            "before": baseline,
            "after_overlay": after,
            "commit": BASELINE_COMMIT,
            "production_tree_sha256": _production_tree(baseline_root),
            "overlay_paths": allowed_status,
            "probe": {
                "source_path": str(candidate_probe),
                "destination_path": str(baseline_probe),
                "sha256": prepared_probe_sha,
                "bytes": prepared_probe_bytes,
            },
            "lock": {
                "source_path": str(external_lock),
                "destination_path": str(baseline_lock_destination),
                "sha256": prepared_lock_sha,
                "bytes": prepared_lock_size,
            },
        },
        "candidate": {
            "root": str(candidate_root),
            "commit": candidate_ref,
            "probe": {"path": str(candidate_probe), "sha256": candidate_probe_sha, "bytes": candidate_probe_bytes},
            "lock": {"path": str(candidate_lock), "sha256": observed_candidate_lock_sha, "bytes": candidate_lock_bytes},
        },
        "claim_scope": "source preparation identity only; no build, timing, or acceptance claim",
    }
    output.write_text(json.dumps(preparation, indent=2) + "\n", encoding="utf-8")
    return preparation


def reuse_approved_overlay(
    *,
    baseline_root: Path,
    candidate_root: Path,
    candidate_ref: str,
    candidate_lock_sha256: str,
    baseline_lock: Path,
    approved_preparation: Path,
    output: Path,
) -> dict[str, Any]:
    """Bind a retained reviewed overlay without modifying either endpoint.

    The ordinary ``prepare`` path requires a clean baseline before creating
    its overlay.  A reviewed preparation already has that overlay in place,
    so this successor path checks the recorded clean-before-overlay state,
    production tree, and exact bytes, then binds a new candidate receipt
    externally.  It never copies, rewrites, or removes files under either
    source root.
    """
    baseline_root = _resolve(baseline_root)
    candidate_root = _resolve(candidate_root)
    baseline_lock = _resolve(baseline_lock)
    approved_preparation = _resolve(approved_preparation)
    output = _resolve(output)
    if not approved_preparation.is_file() or approved_preparation.is_symlink():
        raise PreparationError(f"approved preparation must be a regular file: {approved_preparation}")
    if _under(output, baseline_root) or _under(output, candidate_root):
        raise PreparationError("preparation output must be outside both source worktrees")
    if output.exists():
        raise PreparationError(f"preparation output already exists: {output}")
    if baseline_root == candidate_root or _under(baseline_root, candidate_root) or _under(candidate_root, baseline_root):
        raise PreparationError("baseline and candidate worktrees must be disjoint")
    try:
        approved = json.loads(approved_preparation.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise PreparationError(f"cannot read approved preparation: {error}") from error
    if not isinstance(approved, dict) or approved.get("schema") != SCHEMA or approved.get("lane") != "odf-shared-triad":
        raise PreparationError("approved preparation schema/lane differs")
    approved_baseline = approved.get("baseline")
    if not isinstance(approved_baseline, dict):
        raise PreparationError("approved baseline preparation is missing")
    before = approved_baseline.get("before")
    approved_after = approved_baseline.get("after_overlay")
    if not isinstance(before, dict) or not isinstance(approved_after, dict):
        raise PreparationError("approved preparation source states are missing")

    baseline = _source_state(baseline_root)
    candidate = _source_state(candidate_root)
    _require_sha(candidate_ref, "candidate_ref")
    if candidate_ref == HISTORICAL_CANDIDATE:
        raise PreparationError("historical 15ab candidate is not admissible")
    _require_hash(candidate_lock_sha256, "candidate_lock_sha256")
    if baseline["head"] != BASELINE_COMMIT:
        raise PreparationError(f"baseline HEAD must be {BASELINE_COMMIT}")
    if candidate["head"] != candidate_ref or not candidate["clean"]:
        raise PreparationError("candidate must be clean and at candidate_ref")
    if _resolve(str(before.get("root"))) != baseline_root or before.get("head") != BASELINE_COMMIT:
        raise PreparationError("approved clean-before-overlay root or commit differs")
    if before.get("clean") is not True or before.get("tree") != baseline["tree"] or before.get("status") != "":
        raise PreparationError("approved preparation lacks a clean pre-overlay baseline")
    if _resolve(str(approved_after.get("root"))) != baseline_root or approved_after.get("tree") != baseline["tree"]:
        raise PreparationError("approved after-overlay baseline differs from the retained root")
    approved_paths = _status_paths(str(approved_after.get("status", "")))
    current_paths = _status_paths(baseline["status"])
    if sorted(approved_paths) != sorted(current_paths) or not approved_paths:
        raise PreparationError("retained baseline status differs from the approved overlay")
    if any(path not in {str(PROBE_PATH), str(LOCK_PATH)} for path in current_paths):
        raise PreparationError(f"baseline has paths outside the reviewed overlay: {current_paths!r}")
    if approved_baseline.get("commit") != BASELINE_COMMIT:
        raise PreparationError("approved baseline commit record differs")
    production_tree = _production_tree(baseline_root)
    if approved_baseline.get("production_tree_sha256") != production_tree:
        raise PreparationError("retained baseline production tree differs from approved receipt")

    candidate_probe = _resolve(PROBE_PATH, base=candidate_root)
    candidate_lock = _resolve(LOCK_PATH, base=candidate_root)
    baseline_probe = _resolve(PROBE_PATH, base=baseline_root)
    baseline_lock_destination = _resolve(LOCK_PATH, base=baseline_root)
    candidate_probe_sha, candidate_probe_bytes = _sha256(candidate_probe)
    observed_candidate_lock_sha, candidate_lock_bytes = _sha256(candidate_lock)
    if observed_candidate_lock_sha != candidate_lock_sha256:
        raise PreparationError("candidate Cargo.lock hash does not match supplied identity")
    approved_probe = approved_baseline.get("probe")
    approved_lock = approved_baseline.get("lock")
    if not isinstance(approved_probe, dict) or not isinstance(approved_lock, dict):
        raise PreparationError("approved overlay probe/lock identities are missing")
    try:
        approved_probe_sha = _require_hash(str(approved_probe["sha256"]), "approved probe SHA")
        approved_probe_bytes = int(approved_probe["bytes"])
        approved_lock_sha = _require_hash(str(approved_lock["sha256"]), "approved lock SHA")
        approved_lock_bytes = int(approved_lock["bytes"])
    except (KeyError, TypeError, ValueError) as error:
        raise PreparationError(f"approved overlay identity is malformed: {error}") from error
    if approved_probe_bytes <= 0 or approved_lock_bytes <= 0:
        raise PreparationError("approved overlay identity byte counts must be positive")
    retained_probe_sha, retained_probe_bytes = _sha256(baseline_probe)
    retained_lock_sha, retained_lock_bytes = _sha256(baseline_lock_destination)
    external_lock_sha, external_lock_bytes = _sha256(baseline_lock)
    if (candidate_probe_sha, candidate_probe_bytes) != (approved_probe_sha, approved_probe_bytes):
        raise PreparationError("candidate probe differs from the approved overlay bytes")
    if (retained_probe_sha, retained_probe_bytes) != (approved_probe_sha, approved_probe_bytes):
        raise PreparationError("retained baseline probe differs from the approved overlay bytes")
    if (retained_lock_sha, retained_lock_bytes) != (approved_lock_sha, approved_lock_bytes):
        raise PreparationError("retained baseline Cargo.lock differs from the approved overlay bytes")
    if (external_lock_sha, external_lock_bytes) != (approved_lock_sha, approved_lock_bytes):
        raise PreparationError("supplied baseline lock differs from the approved overlay bytes")

    preparation = {
        "schema": SCHEMA,
        "lane": "odf-shared-triad",
        "baseline": {
            "before": before,
            "after_overlay": baseline,
            "commit": BASELINE_COMMIT,
            "production_tree_sha256": production_tree,
            "overlay_paths": current_paths,
            "probe": {
                "source_path": str(candidate_probe),
                "destination_path": str(baseline_probe),
                "sha256": retained_probe_sha,
                "bytes": retained_probe_bytes,
            },
            "lock": {
                "source_path": str(baseline_lock),
                "destination_path": str(baseline_lock_destination),
                "sha256": retained_lock_sha,
                "bytes": retained_lock_bytes,
            },
        },
        "candidate": {
            "root": str(candidate_root),
            "commit": candidate_ref,
            "probe": {"path": str(candidate_probe), "sha256": candidate_probe_sha, "bytes": candidate_probe_bytes},
            "lock": {"path": str(candidate_lock), "sha256": observed_candidate_lock_sha, "bytes": candidate_lock_bytes},
        },
        "claim_scope": "retained approved source overlay rebound to selected candidate; no build, timing, or acceptance claim",
    }
    output.parent.mkdir(parents=True, exist_ok=True)
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
    parser.add_argument("--reuse-approved-overlay", action="store_true")
    parser.add_argument("--approved-preparation", type=Path)
    args = parser.parse_args(argv)
    try:
        common = {
            "baseline_root": _resolve(args.baseline_root),
            "candidate_root": _resolve(args.candidate_root),
            "candidate_ref": args.candidate_ref,
            "candidate_lock_sha256": args.candidate_lock_sha256,
            "baseline_lock": _resolve(args.baseline_lock),
            "output": args.output,
        }
        if args.reuse_approved_overlay:
            if args.approved_preparation is None:
                parser.error("--approved-preparation is required with --reuse-approved-overlay")
            reuse_approved_overlay(
                **common,
                approved_preparation=_resolve(args.approved_preparation),
            )
        else:
            if args.approved_preparation is not None:
                parser.error("--approved-preparation requires --reuse-approved-overlay")
            prepare(**common)
        return 0
    except (OSError, PreparationError) as error:
        print(f"error: {error}", file=os.sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
