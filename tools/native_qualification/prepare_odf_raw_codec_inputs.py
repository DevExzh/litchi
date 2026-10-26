#!/usr/bin/env python3
"""Prepare the pristine ODF raw-codec baseline overlay.

This is source preparation only.  It never invokes Cargo, Rust, a probe, or
timing code.  The baseline is the exact pre-migration commit; its only
overlay is the reviewed consumer adapter and a separately hashed lockfile.
The candidate remains pinned to C33 and receives only that adapter overlay;
production files are never copied from the adapter endpoint.
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
import tempfile
from typing import Any


SCHEMA = "litchi.odf.raw-codec-qualification.preparation.v2"
LANE = "odf-raw-codec"
BASELINE_COMMIT = "70d64634c9d1d4328447d8724d3b076114a91fd0"
CANDIDATE_COMMIT = "33a0a63d3553712c1276da592f7846a0c7dcabd6"
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


def _production_source_tree(root: Path) -> str:
    """Bind production/workspace provenance without binding the consumer overlay."""

    return hashlib.sha256(
        _git(
            root,
            "ls-tree",
            "-r",
            "--full-tree",
            "HEAD",
            "--",
            "Cargo.toml",
            "Cargo.lock",
            ".cargo",
            "rust-toolchain.toml",
            "rust-toolchain",
            "crates",
        ).encode()
    ).hexdigest()


def _external_output(path: Path, roots: tuple[Path, ...]) -> Path:
    lexical = path.absolute()
    cursor = Path(lexical.anchor)
    for component in lexical.parts[1:]:
        cursor /= component
        if cursor.is_symlink():
            raise PreparationError(f"preparation receipt path contains a symlink component: {cursor}")
    if lexical.exists() or lexical.is_symlink():
        raise PreparationError(f"refusing to overwrite preparation receipt: {path}")
    resolved = lexical.parent.resolve() / lexical.name
    for root in roots:
        root = root.resolve()
        if resolved == root or root in resolved.parents:
            raise PreparationError(f"preparation receipt must be outside source roots: {path}")
    if not resolved.parent.is_dir():
        raise PreparationError(f"preparation receipt parent is not a directory: {resolved.parent}")
    return resolved


def _write_exclusive_json(path: Path, value: dict[str, Any]) -> None:
    payload = (json.dumps(value, indent=2) + "\n").encode("utf-8")
    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        with temporary.open("wb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        os.link(temporary, path)
    except FileExistsError as error:
        raise PreparationError(f"refusing to overwrite preparation receipt: {path}") from error
    finally:
        temporary.unlink(missing_ok=True)


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
    adapter_root: Path,
    candidate_ref: str,
    candidate_lock_sha256: str,
    baseline_lock: Path,
    output: Path,
) -> dict[str, Any]:
    baseline_root = _resolve(baseline_root)
    candidate_root = _resolve(candidate_root)
    adapter_root = _resolve(adapter_root)
    output = _external_output(output, (baseline_root, candidate_root, adapter_root))
    baseline = _source_state(baseline_root)
    candidate = _source_state(candidate_root)
    adapter = _source_state(adapter_root)
    _require_sha(candidate_ref, "candidate_ref")
    if candidate_ref == HISTORICAL_CANDIDATE:
        raise PreparationError("historical 15ab candidate is not admissible")
    if candidate_ref != CANDIDATE_COMMIT:
        raise PreparationError(f"candidate production pin must remain {CANDIDATE_COMMIT}")
    _require_hash(candidate_lock_sha256, "candidate_lock_sha256")
    if baseline["head"] != BASELINE_COMMIT:
        raise PreparationError(f"baseline HEAD must be {BASELINE_COMMIT}")
    if not baseline["clean"]:
        raise PreparationError("baseline must be clean before overlay")
    if not candidate["clean"] or candidate["head"] != candidate_ref:
        raise PreparationError("candidate production endpoint must be clean and at the pinned C33 commit")
    if not adapter["clean"]:
        raise PreparationError("consumer adapter endpoint must be clean")
    if baseline["skip_worktree_count"] or candidate["skip_worktree_count"] or adapter["skip_worktree_count"]:
        raise PreparationError("sparse/skip-worktree source is not admissible for preparation")
    if len({baseline_root, candidate_root, adapter_root}) != 3:
        raise PreparationError("baseline, candidate production, and adapter endpoints must be disjoint")
    if _production_source_tree(adapter_root) != _production_source_tree(candidate_root):
        raise PreparationError("adapter production source does not match the pinned C33 endpoint")

    adapter_probe = _resolve(PROBE_PATH, base=adapter_root)
    adapter_consumer = _resolve(CONSUMER_MANIFEST, base=adapter_root)
    candidate_probe = _resolve(PROBE_PATH, base=candidate_root)
    candidate_consumer = _resolve(CONSUMER_MANIFEST, base=candidate_root)
    candidate_lock = _resolve(LOCK_PATH, base=candidate_root)
    baseline_probe = _resolve(PROBE_PATH, base=baseline_root)
    baseline_consumer = _resolve(CONSUMER_MANIFEST, base=baseline_root)
    baseline_lock_destination = _resolve(LOCK_PATH, base=baseline_root)
    external_lock = _resolve(baseline_lock)
    if _under(external_lock, baseline_root) or _under(external_lock, candidate_root) or _under(external_lock, adapter_root):
        raise PreparationError("baseline lock source must be external to every source endpoint")
    adapter_probe_sha, adapter_probe_bytes = _sha256(adapter_probe)
    adapter_consumer_sha, adapter_consumer_bytes = _sha256(adapter_consumer)
    observed_candidate_lock_sha, candidate_lock_bytes = _sha256(candidate_lock)
    if observed_candidate_lock_sha != candidate_lock_sha256:
        raise PreparationError("candidate Cargo.lock hash does not match supplied identity")
    if candidate_probe.exists() or candidate_probe.is_symlink():
        raise PreparationError(f"candidate probe destination must be absent: {candidate_probe}")
    baseline_lock_sha, baseline_lock_bytes = _sha256(external_lock)
    if baseline_probe.exists() or baseline_probe.is_symlink():
        raise PreparationError(f"baseline probe destination must be absent: {baseline_probe}")
    if candidate_consumer.is_symlink() or baseline_consumer.is_symlink() or baseline_lock_destination.is_symlink():
        raise PreparationError("consumer/lock overlay destination may not be a symlink")

    destinations = (baseline_probe, baseline_consumer, baseline_lock_destination, candidate_probe, candidate_consumer)
    for destination in destinations:
        if destination.exists() and not destination.is_file():
            raise PreparationError(f"overlay destination must be a regular file or absent: {destination}")
    backups: dict[Path, bytes | None] = {
        path: path.read_bytes() if path.is_file() else None for path in destinations
    }
    try:
        baseline_probe.parent.mkdir(parents=True, exist_ok=True)
        candidate_probe.parent.mkdir(parents=True, exist_ok=True)
        baseline_consumer.parent.mkdir(parents=True, exist_ok=True)
        candidate_consumer.parent.mkdir(parents=True, exist_ok=True)
        baseline_lock_destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(adapter_probe, baseline_probe)
        shutil.copyfile(adapter_probe, candidate_probe)
        shutil.copyfile(adapter_consumer, baseline_consumer)
        shutil.copyfile(adapter_consumer, candidate_consumer)
        shutil.copyfile(external_lock, baseline_lock_destination)
    except BaseException:
        for path, content in backups.items():
            if content is None:
                path.unlink(missing_ok=True)
            else:
                path.write_bytes(content)
        raise
    prepared_probe_sha, prepared_probe_bytes = _sha256(baseline_probe)
    prepared_consumer_sha, prepared_consumer_bytes = _sha256(baseline_consumer)
    prepared_lock_sha, prepared_lock_size = _sha256(baseline_lock_destination)
    if (prepared_probe_sha, prepared_probe_bytes) != (adapter_probe_sha, adapter_probe_bytes):
        raise PreparationError("baseline adapter differs from the reviewed adapter probe")
    if (prepared_consumer_sha, prepared_consumer_bytes) != (adapter_consumer_sha, adapter_consumer_bytes):
        raise PreparationError("baseline consumer manifest differs from the reviewed adapter")
    if (prepared_lock_sha, prepared_lock_size) != (baseline_lock_sha, baseline_lock_bytes):
        raise PreparationError("baseline Cargo.lock differs from pinned external lock")

    after = _source_state(baseline_root)
    candidate_after = _source_state(candidate_root)
    allowed_status = [str(CONSUMER_MANIFEST), str(PROBE_PATH)]
    if str(LOCK_PATH) in _status_paths(after["status"]):
        allowed_status.append(str(LOCK_PATH))
    if sorted(_status_paths(after["status"])) != sorted(allowed_status):
        raise PreparationError(f"baseline overlay changed unexpected paths: {_status_paths(after['status'])!r}")
    if sorted(_status_paths(candidate_after["status"])) != sorted([str(CONSUMER_MANIFEST), str(PROBE_PATH)]):
        raise PreparationError(f"candidate overlay changed unexpected paths: {_status_paths(candidate_after['status'])!r}")
    if after["tree"] != baseline["tree"]:
        raise PreparationError("baseline tracked production tree changed during overlay")
    preparation = {
        "schema": SCHEMA,
        "lane": LANE,
        "baseline": {
            "before": baseline,
            "after_overlay": after,
            "commit": BASELINE_COMMIT,
            "production_tree_sha256": _production_tree(baseline_root),
            "overlay_paths": allowed_status,
            "probe": {"source_path": str(adapter_probe), "destination_path": str(baseline_probe), "sha256": prepared_probe_sha, "bytes": prepared_probe_bytes},
            "consumer_manifest": {"source_path": str(adapter_consumer), "destination_path": str(baseline_consumer), "sha256": prepared_consumer_sha, "bytes": prepared_consumer_bytes},
            "lock": {"source_path": str(external_lock), "destination_path": str(baseline_lock_destination), "sha256": prepared_lock_sha, "bytes": prepared_lock_size},
        },
        "candidate": {
            "root": str(candidate_root),
            "commit": candidate_ref,
            "before": candidate,
            "after_overlay": candidate_after,
            "production_tree_sha256": _production_tree(candidate_root),
            "production_source_tree_sha256": _production_source_tree(candidate_root),
            "overlay_paths": [str(CONSUMER_MANIFEST), str(PROBE_PATH)],
            "consumer_manifest": {"source_path": str(adapter_consumer), "path": str(candidate_consumer), "sha256": adapter_consumer_sha, "bytes": adapter_consumer_bytes},
            "probe": {"source_path": str(adapter_probe), "path": str(candidate_probe), "sha256": adapter_probe_sha, "bytes": adapter_probe_bytes},
            "lock": {"path": str(candidate_lock), "sha256": observed_candidate_lock_sha, "bytes": candidate_lock_bytes},
        },
        "adapter": {
            "root": str(adapter_root),
            "head": adapter["head"],
            "tree": adapter["tree"],
            "production_source_tree_sha256": _production_source_tree(adapter_root),
            "consumer_manifest": {"path": str(adapter_consumer), "sha256": adapter_consumer_sha, "bytes": adapter_consumer_bytes},
            "probe": {"path": str(adapter_probe), "sha256": adapter_probe_sha, "bytes": adapter_probe_bytes},
        },
        "claim_scope": "source preparation identity only; no native acceptance claim; no build or timing claim",
    }
    _write_exclusive_json(output, preparation)
    return preparation


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-root", type=Path, required=True)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--adapter-root", type=Path, required=True)
    parser.add_argument("--candidate-ref", required=True)
    parser.add_argument("--candidate-lock-sha256", required=True)
    parser.add_argument("--baseline-lock", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        prepare(
            baseline_root=args.baseline_root,
            candidate_root=args.candidate_root,
            adapter_root=args.adapter_root,
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
