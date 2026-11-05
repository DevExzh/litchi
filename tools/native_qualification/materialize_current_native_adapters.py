#!/usr/bin/env python3
"""Materialize the reviewed consumer-only overlays for the frozen f2 source.

The f2 checkout is deliberately kept pristine in Git.  This helper operates
on a separate detached worktree, verifies its original tree first, copies only
the three fixed adapter files, and emits a small provenance receipt.  It does
not alter production source, regenerate a lockfile, invoke Cargo, or accept a
caller-provided adapter path or digest.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import sys
from typing import Any

from native_qualification import current_native_policy_v2 as policy


MATERIALIZATION_SCHEMA = "litchi.current-native-adapter-materialization.v1"


class MaterializationError(ValueError):
    """The external f2 adapter worktree is not safe to materialize."""


def _canonical(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True) + "\n").encode("utf-8")


def _safe_receipt_path(path: Path, source: Path) -> Path:
    lexical = path.absolute()
    policy._reject_symlink_components(lexical, "adapter materialization receipt")
    if lexical.exists() or lexical.is_symlink():
        raise MaterializationError(f"refusing to overwrite adapter receipt: {path}")
    resolved = lexical.parent.resolve() / lexical.name
    protocol = Path(__file__).resolve().parents[2]
    source = source.resolve()
    if resolved == source or source in resolved.parents or resolved == protocol or protocol in resolved.parents:
        raise MaterializationError("adapter receipt must be outside the source and protocol checkouts")
    resolved.parent.mkdir(parents=True, exist_ok=True)
    return resolved


def _status_rows(source: Path) -> list[dict[str, str]]:
    raw = policy._git(source, "status", "--porcelain=v1", "--ignored=matching", "--untracked-files=all")
    rows: list[dict[str, str]] = []
    for line in raw.splitlines():
        if len(line) < 4:
            raise MaterializationError(f"malformed source status row: {line!r}")
        relative = line[3:]
        if relative.startswith('"') and relative.endswith('"'):
            relative = relative[1:-1]
        rows.append({"status": line[:2], "path": relative.rstrip("/")})
    return rows


def _assert_pristine_f2(source: Path) -> tuple[tuple[dict[str, str], ...], str]:
    head = policy._git(source, "rev-parse", "HEAD")
    tree = policy._git(source, "rev-parse", "HEAD^{tree}")
    if head != policy.BASELINE_COMMIT or tree != policy.BASELINE_GIT_TREE:
        raise MaterializationError(
            f"adapter destination must be the frozen f2 source: {head}/{tree}"
        )
    allowed_lock = source / policy.ROLE_LOCK_PATHS["workspace"]
    lock_digest, lock_size = policy._hash_file(allowed_lock, "approved root Cargo.lock")
    if lock_digest != policy.BASELINE_LOCK_SHA256 or lock_size != policy.BASELINE_LOCK_BYTES:
        raise MaterializationError("approved root Cargo.lock differs from the frozen f2 identity")
    rows = _status_rows(source)
    for row in rows:
        if row["path"] != policy.ROLE_LOCK_PATHS["workspace"] or row["status"] != "!!":
            raise MaterializationError(f"source must be pristine before adapter materialization: {row}")
    entries = policy._git_tree_entries(source)
    policy._tracked_materialization(
        source,
        entries,
        (policy.TRUSTED_POLICY.baseline.lock_for_role("workspace"),),
    )
    return entries, hashlib.sha256(policy._git_tree_payload(source)).hexdigest()


def materialize_baseline_adapters(source: Path, receipt: Path | None = None) -> dict[str, Any]:
    """Copy the reviewed consumer overlays into an external f2 worktree."""

    source = policy._safe_source(Path(source), "adapter destination source")
    policy._reject_protocol_source(source, "adapter destination source")
    if not policy.TRUSTED_POLICY.baseline.complete():
        raise MaterializationError("trusted baseline adapter policy is incomplete")
    entries, tree_sha256 = _assert_pristine_f2(source)
    protocol_root = Path(__file__).resolve().parents[2]
    original_manifest = {
        item.relative_path: item
        for item in policy.TRUSTED_POLICY.baseline.source_manifests
    }
    overlay_rows: list[dict[str, Any]] = []
    for overlay in policy.TRUSTED_POLICY.baseline.adapter_overlays:
        adapter_relative = policy.ADAPTER_SOURCE_PATHS[overlay.relative_path]
        adapter_source = protocol_root / adapter_relative
        source_digest, source_size = policy._hash_file(adapter_source, f"adapter source {adapter_relative}")
        if source_digest != overlay.sha256 or source_size != overlay.bytes:
            raise MaterializationError(
                f"reviewed adapter source changed: {adapter_relative}"
            )
        destination = source / overlay.relative_path
        if destination.exists() or destination.is_symlink():
            if destination.is_symlink() or not destination.is_file():
                raise MaterializationError(f"adapter destination is not a regular file: {destination}")
            if overlay.relative_path in original_manifest:
                policy._validate_git_pin(
                    source, original_manifest[overlay.relative_path],
                    f"original source manifest {overlay.relative_path}",
                )
            existing_digest, existing_size = policy._hash_file(destination, f"existing adapter {overlay.relative_path}")
            if existing_digest == overlay.sha256 and existing_size == overlay.bytes:
                pass
            elif overlay.relative_path in original_manifest:
                # A tracked manifest may only be replaced from its untouched
                # f2 blob, never from an arbitrary dirty file.
                original_payload = policy._git_blob_payload(source, overlay.relative_path)
                if destination.read_bytes() != original_payload:
                    raise MaterializationError(f"existing adapter path is not the original f2 blob: {destination}")
            else:
                raise MaterializationError(f"existing adapter path differs from the reviewed overlay: {destination}")
        destination.parent.mkdir(parents=True, exist_ok=True)
        if not destination.exists():
            destination.write_bytes(adapter_source.read_bytes())
        elif policy._hash_file(destination, f"adapter destination {overlay.relative_path}") != (overlay.sha256, overlay.bytes):
            destination.write_bytes(adapter_source.read_bytes())
        final_digest, final_size = policy._hash_file(destination, f"adapter destination {overlay.relative_path}")
        if (final_digest, final_size) != (overlay.sha256, overlay.bytes):
            raise MaterializationError(f"adapter overlay was not materialized exactly: {destination}")
        overlay_rows.append({
            "relative_path": overlay.relative_path,
            "source_path": str(adapter_source.resolve()),
            "path": str(destination.resolve()),
            "sha256": final_digest,
            "bytes": final_size,
        })

    # Recheck the complete tracked tree after copying.  This catches a dirty
    # path, sparse omission, symlink, or unreviewed overlay before any caller
    # can build the endpoint.
    tracked = policy._tracked_materialization(
        source,
        entries,
        (
            policy.TRUSTED_POLICY.baseline.lock_for_role("workspace"),
            *policy.TRUSTED_POLICY.baseline.adapter_overlays,
        ),
        require_overlays=True,
    )
    rows = _status_rows(source)
    expected_paths = {item.relative_path for item in policy.TRUSTED_POLICY.baseline.adapter_overlays}
    actual_paths = {row["path"] for row in rows}
    if not expected_paths.issubset(actual_paths):
        raise MaterializationError("adapter status does not expose every reviewed overlay")
    result: dict[str, Any] = {
        "schema": MATERIALIZATION_SCHEMA,
        "policy_schema": policy.POLICY_SCHEMA,
        "policy_sha256": policy.TRUSTED_POLICY.digest(),
        "source_root": str(source),
        "commit": policy._git(source, "rev-parse", "HEAD"),
        "git_tree": policy._git(source, "rev-parse", "HEAD^{tree}"),
        "tree_sha256": tree_sha256,
        "original_tree": {
            "commit": policy.BASELINE_COMMIT,
            "git_tree": policy.BASELINE_GIT_TREE,
            "tree_sha256": policy.BASELINE_TREE_SHA256,
            "tracked_files": len(entries),
        },
        "adapter_overlays": overlay_rows,
        "dirty_paths": rows,
        "tracked_materialization": tracked,
        "build_boundary": {
            "source_archive_represents": "original f2 Git tree only",
            "overlay_files_are": "reviewed consumer adapters only",
            "production_source_mutated": False,
            "lock_regenerated": False,
        },
    }
    if receipt is not None:
        destination = _safe_receipt_path(Path(receipt), source)
        destination.write_bytes(_canonical(result))
        result["receipt"] = {"path": str(destination), "sha256": hashlib.sha256(destination.read_bytes()).hexdigest()}
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--receipt", type=Path)
    args = parser.parse_args(argv)
    try:
        result = materialize_baseline_adapters(args.source, args.receipt)
    except (OSError, policy.PolicyError, MaterializationError) as error:
        print(f"current-native adapter materialization refused: {error}", file=sys.stderr)
        return 1
    print(f"current-native adapters materialized: {result['source_root']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
