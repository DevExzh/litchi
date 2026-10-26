#!/usr/bin/env python3
"""Prepare the pristine ODF raw-codec source pair.

This is source preparation only.  It never invokes Cargo, Rust, a probe, or
timing code.  The baseline is the exact pre-migration commit; its only
overlay is the reviewed consumer adapter and a separately hashed baseline
lockfile.  The candidate is an already-integrated final-C worktree.  Its
commit, tree, production closure, and three consumer files are supplied by a
separate protocol-P1-sealed identity.  P1 is independently pinned and may
have a later clean P2 binding installed on the protocol worktree; the
preparer never writes to either source endpoint.
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
from typing import Any, Mapping


SCHEMA = "litchi.odf.raw-codec-qualification.preparation.v5"
IDENTITY_SCHEMA = "litchi.odf.raw-codec-qualification.candidate-identity.v3"
PROTOCOL_P_PIN_SCHEMA = "litchi.odf.raw-codec-qualification.protocol-p1-pin.v1"
CANDIDATE_MODE = "integrated_final_c"
LANE = "odf-raw-codec"
BASELINE_COMMIT = "70d64634c9d1d4328447d8724d3b076114a91fd0"
HISTORICAL_CANDIDATE = "15ab9324ef4f8c927acd6ecfc19d9707c3f886ee"
# The probe remains one canonical reviewed source under perf-baseline.  Only
# the manifest and lock move into the isolated ODF consumer package.
PROBE_PATH = Path("tools/perf-baseline/examples/odf_raw_codec_qualification.rs")
CONSUMER_MANIFEST = Path("tools/odf-raw-codec-consumer/Cargo.toml")
LOCK_PATH = Path("tools/odf-raw-codec-consumer/Cargo.lock")
CONSUMER_LOCK_KEY = "consumer_lock"
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
    non_normal_flags = [line for line in tracked_flags if not line.startswith("H")]
    if non_normal_flags:
        raise PreparationError(f"source has sparse or assume-unchanged tracked entries: {non_normal_flags[:3]!r}")
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


def _canonical_json(value: Mapping[str, Any]) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode("utf-8")


def _candidate_identity_digest(candidate: Mapping[str, Any]) -> str:
    return hashlib.sha256(_canonical_json(candidate)).hexdigest()


def _file_identity(value: Any, label: str, expected_path: Path) -> dict[str, Any]:
    if not isinstance(value, Mapping) or set(value) != {"path", "sha256", "bytes"}:
        raise PreparationError(f"{label} identity must contain only path, sha256 and bytes")
    if value.get("path") != str(expected_path):
        raise PreparationError(f"{label} path differs from the canonical consumer path")
    digest = _require_hash(value.get("sha256"), f"{label}.sha256")
    size = value.get("bytes")
    if isinstance(size, bool) or not isinstance(size, int) or size <= 0:
        raise PreparationError(f"{label}.bytes must be a positive integer")
    return {"path": str(expected_path), "sha256": digest, "bytes": size}


def validate_candidate_identity(value: Mapping[str, Any]) -> dict[str, Any]:
    """Validate the small identity envelope sealed by protocol P1.

    The seal is deliberately a digest over the candidate object, not over the
    candidate worktree itself.  The preparer then re-measures every field
    against Git and the three regular consumer files, so a caller cannot turn
    arbitrary metadata into source admission.
    """

    if not isinstance(value, Mapping) or set(value) != {"schema", "sealed_by", "candidate"}:
        raise PreparationError("candidate identity envelope has unexpected fields")
    if value.get("schema") != IDENTITY_SCHEMA:
        raise PreparationError("candidate identity schema differs")
    sealed_by = value.get("sealed_by")
    if not isinstance(sealed_by, Mapping) or set(sealed_by) != {"role", "commit", "tree", "p1_pin_sha256", "candidate_sha256"}:
        raise PreparationError("candidate identity seal is incomplete")
    if sealed_by.get("role") != "protocol-P1":
        raise PreparationError("candidate identity must be sealed by protocol P1")
    _require_sha(sealed_by.get("commit"), "candidate identity seal commit")
    _require_sha(sealed_by.get("tree"), "candidate identity seal tree")
    _require_hash(sealed_by.get("p1_pin_sha256"), "candidate identity P1 pin digest")
    sealed_digest = _require_hash(sealed_by.get("candidate_sha256"), "candidate identity seal digest")
    candidate = value.get("candidate")
    if not isinstance(candidate, Mapping) or set(candidate) != {
        "mode",
        "commit",
        "tree",
        "production_source_tree_sha256",
        "consumer_manifest",
        "probe",
        CONSUMER_LOCK_KEY,
    }:
        raise PreparationError("candidate identity fields are incomplete")
    if candidate.get("mode") != CANDIDATE_MODE:
        raise PreparationError("candidate identity mode is not integrated final C")
    commit = _require_sha(candidate.get("commit"), "candidate identity candidate commit")
    if commit == HISTORICAL_CANDIDATE:
        raise PreparationError("historical 15ab candidate is not admissible")
    _require_sha(candidate.get("tree"), "candidate identity candidate tree")
    _require_hash(candidate.get("production_source_tree_sha256"), "candidate identity production source closure")
    _file_identity(candidate.get("consumer_manifest"), "candidate consumer manifest", CONSUMER_MANIFEST)
    _file_identity(candidate.get("probe"), "candidate probe", PROBE_PATH)
    _file_identity(candidate.get(CONSUMER_LOCK_KEY), "candidate consumer lock", LOCK_PATH)
    if sealed_digest != _candidate_identity_digest(candidate):
        raise PreparationError("candidate identity seal digest does not match the candidate object")
    return {
        "schema": IDENTITY_SCHEMA,
        "sealed_by": dict(sealed_by),
        "candidate": {
            "mode": CANDIDATE_MODE,
            "commit": commit,
            "tree": str(candidate["tree"]),
            "production_source_tree_sha256": str(candidate["production_source_tree_sha256"]),
            "consumer_manifest": dict(candidate["consumer_manifest"]),
            "probe": dict(candidate["probe"]),
            CONSUMER_LOCK_KEY: dict(candidate[CONSUMER_LOCK_KEY]),
        },
    }


def _load_candidate_identity(path: Path, roots: tuple[Path, ...]) -> tuple[dict[str, Any], str, int, Path]:
    path = _resolve(path)
    if any(_under(path, root) or path == root for root in roots):
        raise PreparationError("candidate identity must be outside every source endpoint")
    if path.is_symlink() or not path.is_file():
        raise PreparationError(f"candidate identity must be a regular file: {path}")
    try:
        raw = path.read_bytes()
        value = json.loads(raw)
    except (OSError, json.JSONDecodeError) as error:
        raise PreparationError(f"cannot read candidate identity: {error}") from error
    identity = validate_candidate_identity(value)
    return identity, hashlib.sha256(raw).hexdigest(), len(raw), path


def _load_protocol_p_pin(path: Path, protocol_p_root: Path, roots: tuple[Path, ...]) -> tuple[dict[str, Any], str, int, Path]:
    path = _resolve(path)
    if any(_under(path, root) or path == root for root in roots):
        raise PreparationError("protocol-P1 pin must be outside every source endpoint")
    if path.is_symlink() or not path.is_file():
        raise PreparationError(f"protocol-P1 pin must be a regular file: {path}")
    try:
        raw = path.read_bytes()
        value = json.loads(raw)
    except (OSError, json.JSONDecodeError) as error:
        raise PreparationError(f"cannot read protocol-P1 pin: {error}") from error
    if not isinstance(value, Mapping) or set(value) != {"schema", "root", "source_candidate", "p1"}:
        raise PreparationError("protocol-P1 pin envelope has unexpected fields")
    if value.get("schema") != PROTOCOL_P_PIN_SCHEMA or value.get("root") != str(protocol_p_root):
        raise PreparationError("protocol-P1 pin schema or root differs")
    source_candidate = value.get("source_candidate")
    p1 = value.get("p1")
    if not isinstance(source_candidate, Mapping) or set(source_candidate) != {"commit", "tree"}:
        raise PreparationError("protocol-P1 source candidate pin is incomplete")
    if not isinstance(p1, Mapping) or set(p1) != {"commit", "tree"}:
        raise PreparationError("protocol-P1 object pin is incomplete")
    normalized = {
        "schema": PROTOCOL_P_PIN_SCHEMA,
        "root": str(protocol_p_root),
        "source_candidate": {
            "commit": _require_sha(source_candidate.get("commit"), "protocol-P1 source candidate commit"),
            "tree": _require_sha(source_candidate.get("tree"), "protocol-P1 source candidate tree"),
        },
        "p1": {
            "commit": _require_sha(p1.get("commit"), "protocol-P1 commit"),
            "tree": _require_sha(p1.get("tree"), "protocol-P1 tree"),
        },
    }
    return normalized, hashlib.sha256(raw).hexdigest(), len(raw), path


def _require_tracked_regular(root: Path, relative: Path, label: str) -> Path:
    path = _resolve(relative, base=root)
    if path.is_symlink() or not path.is_file():
        raise PreparationError(f"{label} must be a regular file: {path}")
    tracked = _git(root, "ls-files", "--error-unmatch", "--", str(relative)).strip().splitlines()
    if tracked != [str(relative)]:
        raise PreparationError(f"{label} must be tracked exactly at {relative}")
    return path


def _verify_protocol_p_pin(protocol_p_root: Path, pin: Mapping[str, Any], candidate: Mapping[str, Any]) -> dict[str, str]:
    protocol_p_root = _resolve(protocol_p_root)
    top = _resolve(_git(protocol_p_root, "rev-parse", "--show-toplevel").strip())
    if top != protocol_p_root:
        raise PreparationError("protocol-P1 pin root is not a worktree top-level")
    if pin.get("root") != str(protocol_p_root):
        raise PreparationError("protocol-P1 pin root differs from the selected endpoint")
    source_candidate = pin["source_candidate"]
    p1 = pin["p1"]
    candidate_commit = _require_sha(candidate.get("commit"), "candidate identity candidate commit")
    candidate_tree = _require_sha(candidate.get("tree"), "candidate identity candidate tree")
    if source_candidate["commit"] != candidate_commit or source_candidate["tree"] != candidate_tree:
        raise PreparationError("protocol-P1 pin source candidate differs from the sealed final-C identity")
    p1_commit = _require_sha(p1["commit"], "protocol-P1 commit")
    p1_tree = _require_sha(p1["tree"], "protocol-P1 tree")
    _git(protocol_p_root, "cat-file", "-e", f"{p1_commit}^{{commit}}")
    actual_p1_tree = _git(protocol_p_root, "rev-parse", f"{p1_commit}^{{tree}}").strip()
    if actual_p1_tree != p1_tree:
        raise PreparationError("protocol-P1 commit tree differs from the supplied pin")
    _git(protocol_p_root, "cat-file", "-e", f"{candidate_commit}^{{commit}}")
    actual_candidate_tree = _git(protocol_p_root, "rev-parse", f"{candidate_commit}^{{tree}}").strip()
    if actual_candidate_tree != candidate_tree:
        raise PreparationError("protocol-P1 source candidate tree differs from the sealed identity")
    ancestry = subprocess.run(
        ["git", "-C", str(protocol_p_root), "merge-base", "--is-ancestor", candidate_commit, p1_commit],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if ancestry.returncode != 0:
        raise PreparationError("protocol-P1 is not derived from the frozen final-C candidate")
    current_commit = _require_sha(_git(protocol_p_root, "rev-parse", "HEAD").strip(), "current protocol-P commit")
    current_tree = _require_sha(_git(protocol_p_root, "rev-parse", "HEAD^{tree}").strip(), "current protocol-P tree")
    current_ancestry = subprocess.run(
        ["git", "-C", str(protocol_p_root), "merge-base", "--is-ancestor", p1_commit, current_commit],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if current_ancestry.returncode != 0:
        raise PreparationError("current protocol-P endpoint is not P1 or a descendant P2")
    if _git(protocol_p_root, "status", "--porcelain=v1", "--untracked-files=all"):
        raise PreparationError("protocol-P endpoint must be clean and pinned")
    return {
        "root": str(protocol_p_root),
        "p1_commit": p1_commit,
        "p1_tree": p1_tree,
        "source_candidate_commit": candidate_commit,
        "source_candidate_tree": candidate_tree,
        "current_commit": current_commit,
        "current_tree": current_tree,
    }


def prepare(
    *,
    baseline_root: Path,
    candidate_root: Path,
    adapter_root: Path,
    protocol_p_root: Path,
    protocol_p_pin_path: Path,
    candidate_identity_path: Path,
    baseline_lock: Path,
    output: Path,
) -> dict[str, Any]:
    baseline_root = _resolve(baseline_root)
    candidate_root = _resolve(candidate_root)
    adapter_root = _resolve(adapter_root)
    protocol_p_root = _resolve(protocol_p_root)
    protocol_p_pin, protocol_p_pin_sha256, protocol_p_pin_bytes, protocol_p_pin_path = _load_protocol_p_pin(
        protocol_p_pin_path,
        protocol_p_root,
        (baseline_root, candidate_root, adapter_root, protocol_p_root),
    )
    output = _external_output(output, (baseline_root, candidate_root, adapter_root, protocol_p_root))
    identity, identity_sha256, identity_bytes, identity_path = _load_candidate_identity(
        candidate_identity_path,
        (baseline_root, candidate_root, adapter_root, protocol_p_root),
    )
    candidate_identity = identity["candidate"]
    if identity["sealed_by"]["p1_pin_sha256"] != protocol_p_pin_sha256:
        raise PreparationError("candidate identity is sealed against a different protocol-P1 pin")
    if (
        identity["sealed_by"]["commit"] != protocol_p_pin["p1"]["commit"]
        or identity["sealed_by"]["tree"] != protocol_p_pin["p1"]["tree"]
    ):
        raise PreparationError("candidate identity seal does not name the pinned protocol-P1 object")
    protocol_p_anchor = _verify_protocol_p_pin(protocol_p_root, protocol_p_pin, candidate_identity)
    baseline = _source_state(baseline_root)
    candidate = _source_state(candidate_root)
    adapter = _source_state(adapter_root)
    if baseline["head"] != BASELINE_COMMIT:
        raise PreparationError(f"baseline HEAD must be {BASELINE_COMMIT}")
    if not baseline["clean"]:
        raise PreparationError("baseline must be clean before overlay")
    if (
        not candidate["clean"]
        or candidate["head"] != candidate_identity["commit"]
        or candidate["tree"] != candidate_identity["tree"]
    ):
        raise PreparationError("candidate production endpoint is not the clean P-sealed final-C source")
    if not adapter["clean"]:
        raise PreparationError("consumer adapter endpoint must be clean")
    if baseline["skip_worktree_count"] or candidate["skip_worktree_count"] or adapter["skip_worktree_count"]:
        raise PreparationError("sparse/skip-worktree source is not admissible for preparation")
    if len({baseline_root, candidate_root, adapter_root}) != 3:
        raise PreparationError("baseline, candidate production, and adapter endpoints must be disjoint")
    if any(protocol_p_root == root or _under(protocol_p_root, root) or _under(root, protocol_p_root) for root in (baseline_root, candidate_root, adapter_root)):
        raise PreparationError("protocol-P anchor endpoint must be disjoint from source endpoints")
    candidate_production_source_tree = _production_source_tree(candidate_root)
    if candidate_production_source_tree != candidate_identity["production_source_tree_sha256"]:
        raise PreparationError("candidate production source closure differs from the P-sealed identity")

    adapter_probe = _require_tracked_regular(adapter_root, PROBE_PATH, "adapter probe")
    adapter_consumer = _require_tracked_regular(adapter_root, CONSUMER_MANIFEST, "adapter consumer manifest")
    adapter_lock = _require_tracked_regular(adapter_root, LOCK_PATH, "adapter performance lock")
    candidate_probe = _require_tracked_regular(candidate_root, PROBE_PATH, "candidate probe")
    candidate_consumer = _require_tracked_regular(candidate_root, CONSUMER_MANIFEST, "candidate consumer manifest")
    candidate_lock = _require_tracked_regular(candidate_root, LOCK_PATH, "candidate performance lock")
    baseline_probe = _resolve(PROBE_PATH, base=baseline_root)
    baseline_consumer = _resolve(CONSUMER_MANIFEST, base=baseline_root)
    baseline_lock_destination = _resolve(LOCK_PATH, base=baseline_root)
    external_lock = _resolve(baseline_lock)
    if _under(external_lock, baseline_root) or _under(external_lock, candidate_root) or _under(external_lock, adapter_root):
        raise PreparationError("baseline lock source must be external to every source endpoint")
    adapter_probe_sha, adapter_probe_bytes = _sha256(adapter_probe)
    adapter_consumer_sha, adapter_consumer_bytes = _sha256(adapter_consumer)
    adapter_lock_sha, adapter_lock_bytes = _sha256(adapter_lock)
    for label, observed, expected in (
        ("consumer manifest", (adapter_consumer_sha, adapter_consumer_bytes), candidate_identity["consumer_manifest"]),
        ("probe", (adapter_probe_sha, adapter_probe_bytes), candidate_identity["probe"]),
        ("consumer lock", (adapter_lock_sha, adapter_lock_bytes), candidate_identity[CONSUMER_LOCK_KEY]),
    ):
        if observed != (expected["sha256"], expected["bytes"]):
            raise PreparationError(f"adapter {label} does not match the P-sealed final-C consumer identity")
    for path, label, expected in (
        (candidate_consumer, "candidate consumer manifest", candidate_identity["consumer_manifest"]),
        (candidate_probe, "candidate probe", candidate_identity["probe"]),
        (candidate_lock, "candidate consumer lock", candidate_identity[CONSUMER_LOCK_KEY]),
    ):
        observed = _sha256(path)
        if observed != (expected["sha256"], expected["bytes"]):
            raise PreparationError(f"{label} differs from the P-sealed final-C consumer identity")
    baseline_lock_sha, baseline_lock_bytes = _sha256(external_lock)
    if baseline_probe.exists() or baseline_probe.is_symlink():
        raise PreparationError(f"baseline probe destination must be absent: {baseline_probe}")
    if baseline_consumer.is_symlink() or baseline_lock_destination.is_symlink():
        raise PreparationError("baseline consumer/lock overlay destination may not be a symlink")

    destinations = (baseline_probe, baseline_consumer, baseline_lock_destination)
    for destination in destinations:
        if destination.exists() and not destination.is_file():
            raise PreparationError(f"overlay destination must be a regular file or absent: {destination}")
    backups: dict[Path, bytes | None] = {
        path: path.read_bytes() if path.is_file() else None for path in destinations
    }
    try:
        baseline_probe.parent.mkdir(parents=True, exist_ok=True)
        baseline_consumer.parent.mkdir(parents=True, exist_ok=True)
        baseline_lock_destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(adapter_probe, baseline_probe)
        shutil.copyfile(adapter_consumer, baseline_consumer)
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
    candidate_production_source_tree_after = _production_source_tree(candidate_root)
    allowed_status = [str(CONSUMER_MANIFEST), str(PROBE_PATH)]
    if str(LOCK_PATH) in _status_paths(after["status"]):
        allowed_status.append(str(LOCK_PATH))
    if sorted(_status_paths(after["status"])) != sorted(allowed_status):
        raise PreparationError(f"baseline overlay changed unexpected paths: {_status_paths(after['status'])!r}")
    if candidate_after != candidate:
        raise PreparationError("candidate source changed during final-C admission")
    if candidate_production_source_tree_after != candidate_production_source_tree:
        raise PreparationError("candidate production source closure changed during admission")
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
            "mode": CANDIDATE_MODE,
            "commit": candidate_identity["commit"],
            "tree": candidate_identity["tree"],
            "before": candidate,
            "after": candidate_after,
            "production_tree_sha256": _production_tree(candidate_root),
            "production_source_tree_before_sha256": candidate_production_source_tree,
            "production_source_tree_sha256": candidate_production_source_tree_after,
            "overlay_paths": [],
            "consumer_manifest": {"source_path": str(candidate_consumer), "path": str(candidate_consumer), "sha256": candidate_identity["consumer_manifest"]["sha256"], "bytes": candidate_identity["consumer_manifest"]["bytes"]},
            "probe": {"source_path": str(candidate_probe), "path": str(candidate_probe), "sha256": candidate_identity["probe"]["sha256"], "bytes": candidate_identity["probe"]["bytes"]},
            "lock": {"source_path": str(candidate_lock), "path": str(candidate_lock), "sha256": candidate_identity[CONSUMER_LOCK_KEY]["sha256"], "bytes": candidate_identity[CONSUMER_LOCK_KEY]["bytes"]},
        },
        "adapter": {
            "root": str(adapter_root),
            "head": adapter["head"],
            "tree": adapter["tree"],
            "consumer_provenance_scope": "three consumer files only; production closure is not equated to final C",
            "consumer_manifest": {"path": str(adapter_consumer), "sha256": adapter_consumer_sha, "bytes": adapter_consumer_bytes},
            "probe": {"path": str(adapter_probe), "sha256": adapter_probe_sha, "bytes": adapter_probe_bytes},
            "lock": {"path": str(adapter_lock), "sha256": adapter_lock_sha, "bytes": adapter_lock_bytes},
        },
        "candidate_identity": {
            "path": str(identity_path),
            "sha256": identity_sha256,
            "bytes": identity_bytes,
            "value": identity,
        },
        "protocol_p_pin": {
            "path": str(protocol_p_pin_path),
            "sha256": protocol_p_pin_sha256,
            "bytes": protocol_p_pin_bytes,
            "value": protocol_p_pin,
        },
        "protocol_p_anchor": protocol_p_anchor,
        "claim_scope": "source preparation identity only; no native acceptance claim; no build or timing claim",
    }
    _write_exclusive_json(output, preparation)
    return preparation


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-root", type=Path, required=True)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--adapter-root", type=Path, required=True)
    parser.add_argument("--protocol-p-root", type=Path, required=True)
    parser.add_argument("--protocol-p-pin", type=Path, required=True)
    parser.add_argument("--candidate-identity", type=Path, required=True)
    parser.add_argument("--baseline-lock", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        prepare(
            baseline_root=args.baseline_root,
            candidate_root=args.candidate_root,
            adapter_root=args.adapter_root,
            protocol_p_root=args.protocol_p_root,
            protocol_p_pin_path=args.protocol_p_pin,
            candidate_identity_path=args.candidate_identity,
            baseline_lock=args.baseline_lock,
            output=args.output,
        )
        return 0
    except (OSError, PreparationError) as error:
        print(f"error: {error}", file=os.sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
