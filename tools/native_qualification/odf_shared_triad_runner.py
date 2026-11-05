#!/usr/bin/env python3
"""Fail-closed runner and receipt validator for the ODF shared-triad lane.

This lane is deliberately independent of the historical W4b/W5 runner.  A
manifest may leave the candidate commit unresolved until the final assembly;
the command line must then supply the selected full commit SHA.  The runner
binds that source tree, its probe and lock file, the exact feature set and
route/phase contract, and the complete probe receipt before accepting output.
It runs no Cargo command during manifest or receipt validation.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import resource
import subprocess
import sys
from typing import Any, Mapping, Sequence


SCHEMA = "litchi.odf.shared-triad-qualification.manifest.v2"
RECEIPT_SCHEMA = "litchi.odf.shared-triad-qualification.receipt.v1"
OUTPUT_SCHEMA = "litchi.odf.shared-triad-qualification.v1"
FULL_SHA = re.compile(r"^[0-9a-f]{40}$")
HASH = re.compile(r"^[0-9a-f]{64}$")
HISTORICAL_CANDIDATE = "15ab9324ef4f8c927acd6ecfc19d9707c3f886ee"
MAX_OUTPUT_BYTES = 1_048_576
TARGET_PREFLIGHT_BYTES = 19_200_000_000
TARGET_HARD_LIMIT_BYTES = 20_000_000_000
MEMORY_HARD_LIMIT_BYTES = 20_000_000_000
UPSTREAM_COMPONENT = "3rd" + "party"
FAMILIES = ("odc", "odi", "oth", "odm")
ROUTE_IDS = (
    "odc_chart_title",
    "odi_frame_geometry",
    "oth_paragraph_text",
    "odm_title",
)
ROUTE_DESCRIPTORS = (
    {
        "id": "odc_chart_title",
        "family": "odc",
        "source": "litchi_odc::Package::new",
        "leaf": "Package::edit -> replace_chart(title) -> commit",
        "facade": "File::edit -> set(/, title) -> commit",
        "changed_marker": "shared-triad qualification",
        "no_op": "Package::edit/File::edit with an empty closure",
    },
    {
        "id": "odi_frame_geometry",
        "family": "odi",
        "source": "litchi_odi::Package::new",
        "leaf": "Package::edit -> set_geometry(frame 0) -> commit",
        "facade": "File::edit -> set(/frame[1], geometry) -> commit",
        "changed_marker": "1cm,2cm,3cm,4cm frame geometry",
        "no_op": "Package::edit/File::edit with an empty closure",
    },
    {
        "id": "oth_paragraph_text",
        "family": "oth",
        "source": "litchi_oth::Builder::paragraph(before)",
        "leaf": "Package::edit -> set_paragraph_text(position 0) -> commit",
        "facade": "File::edit -> set(/body/p[1], text) -> commit",
        "changed_marker": "shared-triad qualification",
        "no_op": "Package::edit/File::edit with an empty closure",
    },
    {
        "id": "odm_title",
        "family": "odm",
        "source": "litchi_odm::Builder::new + timestamp-free Structure::default_meta_xml",
        "leaf": "Package::edit -> set_title -> commit",
        "facade": "File::edit -> set(/metadata, title) -> commit",
        "changed_marker": "shared-triad qualification",
        "no_op": "Package::edit/File::edit with an empty closure",
    },
)
PHASES = ("commit_only", "commit_plus_diagnostics", "publication")
SETUP_PHASES = (
    "deterministic builder or tracked fixture read",
    "Package::from_bytes and File::open_bytes",
    "typed edit inputs, props and edit target construction",
)
CLAIM_SCOPE = "correctness only unless mode=timed; no acceptance claim"


class RunnerError(RuntimeError):
    """A manifest, source, output, or process binding failed."""


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise RunnerError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while True:
            block = stream.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _resolve(path: str | os.PathLike[str], *, base: Path | None = None) -> Path:
    candidate = Path(path)
    if not candidate.is_absolute():
        if base is None:
            raise RunnerError(f"path must be absolute: {path!r}")
        candidate = base / candidate
    probe = Path(candidate.anchor)
    for part in candidate.parts[1:]:
        probe /= part
        if probe.is_symlink():
            raise RunnerError(f"symlink path component is forbidden: {candidate}")
    resolved = candidate.expanduser().resolve(strict=False)
    if UPSTREAM_COMPONENT in resolved.parts:
        raise RunnerError(f"untracked upstream checkout path is forbidden: {resolved}")
    return resolved


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _ensure_external(path: Path, source: Path, label: str) -> None:
    if _under(path, source):
        raise RunnerError(f"{label} must be outside the source worktree: {path}")


def _reject_symlink_ancestors(path: Path, label: str) -> None:
    current = path.parent
    ancestors = []
    while True:
        ancestors.append(current)
        if current == current.parent:
            break
        current = current.parent
    for ancestor in reversed(ancestors):
        if ancestor.is_symlink():
            raise RunnerError(f"{label} ancestor may not be a symlink: {ancestor}")


def _directory_size(path: Path) -> int:
    if not path.exists():
        return 0
    if path.is_symlink():
        raise RunnerError(f"directory may not be a symlink: {path}")
    if path.is_file():
        return path.stat().st_size
    total = 0
    for root, dirs, files in os.walk(path, followlinks=False):
        dirs[:] = [name for name in dirs if not (Path(root) / name).is_symlink()]
        for name in files:
            item = Path(root) / name
            if not item.is_symlink():
                total += item.stat().st_size
    return total


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise RunnerError(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.stdout


def _sparse_checkout_value(root: Path) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), "config", "--bool", "core.sparseCheckout"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode not in (0, 1) or result.stderr.strip():
        raise RunnerError(f"cannot inspect sparse-checkout configuration: {result.stderr.strip()}")
    return result.stdout.strip().lower()


def _tracked_closure_identity(root: Path) -> dict[str, Any]:
    if _sparse_checkout_value(root) == "true":
        raise RunnerError(f"sparse-checkout is enabled for source endpoint: {root}")
    head_records = _git(root, "ls-tree", "-r", "-z", "--full-tree", "HEAD").split("\0")
    head_paths: set[str] = set()
    for record in head_records:
        if not record:
            continue
        header, separator, relative = record.partition("\t")
        if not separator:
            raise RunnerError(f"malformed HEAD tree record: {record!r}")
        mode = header.split(" ", 1)[0]
        if mode not in {"100644", "100755"}:
            raise RunnerError(f"source closure contains unsupported tracked mode {mode}: {relative}")
        head_paths.add(relative)
    status_raw = _git(root, "ls-files", "-v", "-z")
    index_raw = _git(root, "ls-files", "--stage", "-z")
    seen: set[str] = set()
    for record in status_raw.split("\0"):
        if not record:
            continue
        marker, separator, relative = record.partition(" ")
        if not separator or marker != "H":
            raise RunnerError(f"source endpoint has sparse/skip-worktree/assume-unchanged entry: {record!r}")
        seen.add(relative)
    digest = hashlib.sha256()
    files = []
    for record in index_raw.split("\0"):
        if not record:
            continue
        header, separator, relative = record.partition("\t")
        if not separator:
            raise RunnerError(f"malformed tracked-file record: {record!r}")
        mode = header.split(" ", 1)[0]
        if mode not in {"100644", "100755"}:
            raise RunnerError(f"source closure contains unsupported tracked mode {mode}: {relative}")
        if relative not in seen:
            raise RunnerError(f"tracked file is missing from full worktree status: {relative}")
        path = _resolve(relative, base=root)
        file_hash, file_bytes = _sha256_file(path)
        digest.update(relative.encode())
        digest.update(b"\0")
        digest.update(mode.encode())
        digest.update(b"\0")
        digest.update(file_hash.encode())
        digest.update(b"\0")
        digest.update(str(file_bytes).encode())
        files.append({"path": relative, "mode": mode, "sha256": file_hash, "bytes": file_bytes})
    if seen != head_paths:
        raise RunnerError("source worktree is not a materialized full HEAD tree")
    if len(files) != len(seen):
        raise RunnerError("source worktree tracked-file closure is incomplete")
    return {
        "sha256": digest.hexdigest(),
        "file_count": len(files),
        "flags_sha256": _sha256_bytes(status_raw.encode()),
        "index_sha256": _sha256_bytes(index_raw.encode()),
    }


def _cargo_config_paths(root: Path) -> list[Path]:
    paths: list[Path] = []
    current = root
    while True:
        cargo = current / ".cargo"
        for name in ("config.toml", "config"):
            candidate = cargo / name
            if candidate.exists() or candidate.is_symlink():
                paths.append(candidate)
        if current.parent == current:
            break
        current = current.parent
    home_cargo = Path(os.environ.get("CARGO_HOME", str(Path.home() / ".cargo")))
    for name in ("config.toml", "config"):
        candidate = home_cargo / name
        if candidate.exists() or candidate.is_symlink():
            paths.append(candidate)
    configured = os.environ.get("CARGO_CONFIG")
    if configured:
        paths.append(Path(configured))
    unique: list[Path] = []
    for path in paths:
        resolved = _resolve(path)
        if resolved not in unique:
            unique.append(resolved)
    return sorted(unique)


def _cargo_config_identity(root: Path) -> dict[str, Any]:
    files = []
    digest = hashlib.sha256()
    for path in _cargo_config_paths(root):
        file_hash, file_bytes = _sha256_file(path)
        digest.update(str(path).encode())
        digest.update(b"\0")
        digest.update(file_hash.encode())
        digest.update(b"\0")
        digest.update(str(file_bytes).encode())
        files.append({"path": str(path), "sha256": file_hash, "bytes": file_bytes})
    return {"sha256": digest.hexdigest(), "files": files}


def _require_sha(value: Any, label: str) -> str:
    if not isinstance(value, str) or not FULL_SHA.fullmatch(value):
        raise RunnerError(f"{label} must be a full 40-character lowercase commit SHA")
    return value


def _require_hash(value: Any, label: str) -> str:
    if not isinstance(value, str) or not HASH.fullmatch(value):
        raise RunnerError(f"{label} must be a full 64-character lowercase SHA-256")
    return value


def validate_manifest(data: Mapping[str, Any]) -> dict[str, Any]:
    """Validate and return a detached manifest mapping.

    Candidate selection and lock identity are intentionally external inputs;
    unresolved values are accepted here but rejected by ``source_identity``.
    This lets the committed manifest remain fail-closed until final assembly.
    """

    if data.get("schema") != SCHEMA:
        raise RunnerError(f"manifest schema must be {SCHEMA!r}")
    if data.get("lane") != "odf-shared-triad":
        raise RunnerError("manifest lane must be odf-shared-triad")
    _require_sha(data.get("baseline_commit"), "baseline_commit")
    candidate = data.get("candidate_commit")
    if candidate is not None:
        candidate = _require_sha(candidate, "candidate_commit")
        if candidate == HISTORICAL_CANDIDATE:
            raise RunnerError("historical 15ab candidate is not a final candidate")

    probe = data.get("probe")
    if not isinstance(probe, Mapping):
        raise RunnerError("probe must be an object")
    if probe.get("path") != "crates/litchi-ops/examples/odf_shared_triad_qualification.rs":
        raise RunnerError("probe path is not the reviewed shared-triad example")
    features = probe.get("features")
    if features != list(FAMILIES):
        raise RunnerError(f"probe features must be {list(FAMILIES)!r}")
    if probe.get("modes") != ["validate", "timed", "allocator"]:
        raise RunnerError("probe modes must be validate, timed and allocator")
    if probe.get("max_iterations") != 128:
        raise RunnerError("probe max_iterations must remain 128")
    _require_hash(probe.get("sha256"), "probe.sha256")

    source = data.get("source")
    if not isinstance(source, Mapping):
        raise RunnerError("source must be an object")
    if source.get("package_manifest") != "crates/litchi-ops/Cargo.toml":
        raise RunnerError("source package_manifest must be crates/litchi-ops/Cargo.toml")
    if source.get("lock_path") != "Cargo.lock":
        raise RunnerError("source lock_path must be Cargo.lock")
    lock_sha = source.get("lock_sha256")
    if lock_sha is not None:
        _require_hash(lock_sha, "source.lock_sha256")

    routes = data.get("routes")
    if not isinstance(routes, list) or any(not isinstance(item, Mapping) for item in routes):
        raise RunnerError("routes must be an ordered list of objects")
    if routes != list(ROUTE_DESCRIPTORS):
        raise RunnerError("routes do not match the exact reviewed descriptors")

    phases = data.get("phases")
    if not isinstance(phases, Mapping):
        raise RunnerError("phases must be an object")
    if tuple(phases.get("inside_clock", ())) != PHASES:
        raise RunnerError("phases.inside_clock does not match the shared-triad contract")
    if tuple(phases.get("setup_outside_clock", ())) != SETUP_PHASES:
        raise RunnerError("phases.setup_outside_clock does not match the timed boundary")
    allocation = phases.get("allocation_observation")
    if (
        not isinstance(allocation, Mapping)
        or allocation.get("facade") != "Commit::diagnostics().snapshot_bytes"
        or allocation.get("leaf") != "not exposed by litchi_core::edit::Diagnostics"
        or allocation.get("process_allocator")
        != "paired runner operation-scoped system allocator observer"
    ):
        raise RunnerError("allocation observation is not the reviewed boundary")

    output = data.get("output")
    if not isinstance(output, Mapping):
        raise RunnerError("output must be an object")
    if output.get("schema") != OUTPUT_SCHEMA or output.get("families") != list(FAMILIES):
        raise RunnerError("output schema/family contract is incomplete")
    if output.get("odm_changed_output_decisions") != [
        "byte_exact",
        "archive_member_order_and_manifest_listing",
    ]:
        raise RunnerError("ODM archive decision contract is incomplete")
    if data.get("claim_scope") != "validation and reproducible phase protocol only; no native acceptance claim":
        raise RunnerError("manifest claim scope must remain non-acceptance")
    return dict(data)


def load_manifest(path: Path) -> tuple[dict[str, Any], bytes]:
    try:
        raw = path.read_bytes()
        data = json.loads(raw)
    except (OSError, json.JSONDecodeError) as error:
        raise RunnerError(f"cannot read manifest {path}: {error}") from error
    if not isinstance(data, Mapping):
        raise RunnerError("manifest root must be an object")
    return validate_manifest(data), raw


def source_identity(
    manifest: Mapping[str, Any],
    *,
    role: str,
    source_root: Path,
    source_commit: str | None,
    lock_sha256: str | None,
) -> dict[str, Any]:
    root = _resolve(source_root)
    if not root.is_dir():
        raise RunnerError(f"source root is not a directory: {root}")
    top = _resolve(_git(root, "rev-parse", "--show-toplevel").strip())
    if top != root:
        raise RunnerError(f"source root is not the worktree top-level: {root}")
    if role not in {"baseline", "candidate"}:
        raise RunnerError("role must be baseline or candidate")
    expected = manifest["baseline_commit"] if role == "baseline" else _require_sha(source_commit, "candidate source_commit")
    if role == "candidate" and expected == HISTORICAL_CANDIDATE:
        raise RunnerError("historical 15ab candidate is not a final candidate")
    if role == "candidate" and manifest.get("candidate_commit") is not None and expected != manifest["candidate_commit"]:
        raise RunnerError("selected candidate SHA differs from the manifest candidate")
    head = _git(root, "rev-parse", "HEAD").strip()
    if head != expected:
        raise RunnerError(f"{role} HEAD {head} does not equal selected commit {expected}")
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    if status:
        raise RunnerError(f"source worktree is dirty:\n{status}")
    # A clean porcelain status does not expose sparse/skip-worktree entries,
    # and Cargo can also be redirected through configuration outside the
    # repository.  Capture both closures from the actual endpoint before
    # constructing the receipt; callers must not be able to self-report them.
    tracked_closure = _tracked_closure_identity(root)
    cargo_configs = _cargo_config_identity(root)
    tree = _git(root, "rev-parse", "HEAD^{tree}").strip()
    diff = _git(root, "diff", "--no-ext-diff", "--binary")

    source_spec = manifest["source"]
    package_manifest = _resolve(str(source_spec["package_manifest"]), base=root)
    probe = _resolve(str(manifest["probe"]["path"]), base=root)
    lock = _resolve(str(source_spec["lock_path"]), base=root)
    for label, path in (("package manifest", package_manifest), ("probe", probe), ("lock", lock)):
        if path.is_symlink() or not path.is_file():
            raise RunnerError(f"{label} is not a regular file: {path}")
    probe_hash, probe_bytes = _sha256_file(probe)
    if probe_hash != manifest["probe"]["sha256"]:
        raise RunnerError("probe SHA-256 does not match the manifest")
    observed_lock_hash, lock_bytes = _sha256_file(lock)
    selected_lock_hash = lock_sha256 or source_spec.get("lock_sha256")
    if selected_lock_hash is None:
        raise RunnerError("lock SHA-256 is unresolved; supply --lock-sha256")
    _require_hash(selected_lock_hash, "lock_sha256")
    if source_spec.get("lock_sha256") is not None and selected_lock_hash != source_spec["lock_sha256"]:
        raise RunnerError("selected lock SHA differs from the manifest lock identity")
    if observed_lock_hash != selected_lock_hash:
        raise RunnerError("Cargo.lock SHA-256 does not match the selected identity")
    try:
        import tomllib

        package = tomllib.loads(package_manifest.read_text(encoding="utf-8"))["package"]
    except (OSError, KeyError, tomllib.TOMLDecodeError) as error:
        raise RunnerError(f"cannot read litchi-ops package manifest: {error}") from error
    if package.get("name") != "litchi-ops":
        raise RunnerError("source package manifest is not litchi-ops")
    package_hash, package_bytes = _sha256_file(package_manifest)
    source_tree_sha256 = _sha256_bytes(
        (
            tree
            + "\0"
            + status
            + "\0"
            + diff
            + "\0"
            + tracked_closure["sha256"]
            + "\0"
            + cargo_configs["sha256"]
        ).encode()
    )
    return {
        "role": role,
        "root": str(root),
        "commit": head,
        "tree": tree,
        "source_tree_sha256": source_tree_sha256,
        "tracked_closure": tracked_closure,
        "cargo_configs": cargo_configs,
        "clean": True,
        "package_manifest": {
            "path": str(package_manifest),
            "sha256": package_hash,
            "bytes": package_bytes,
            "name": package["name"],
        },
        "probe": {"path": str(probe), "sha256": probe_hash, "bytes": probe_bytes},
        "lock": {"path": str(lock), "sha256": observed_lock_hash, "bytes": lock_bytes},
    }


def _require_bool(mapping: Mapping[str, Any], key: str, expected: bool, label: str) -> None:
    value = mapping.get(key)
    if not isinstance(value, bool) or value is not expected:
        raise RunnerError(f"{label}.{key} must be {str(expected).lower()}")


def _validate_side(side: Mapping[str, Any], expected_changed: bool, label: str) -> None:
    for key, expected in (
        ("changed", expected_changed),
        ("patch_noop", not expected_changed),
        ("patch_base_match", True),
        ("diagnostics_changed", expected_changed),
        ("source_unchanged", True),
        ("output_differs_from_source", expected_changed),
        ("publication_repeat_exact", True),
        ("forward_replay_exact", True),
        ("inverse_replay_restores_source", True),
    ):
        _require_bool(side, key, expected, label)
    if not isinstance(side.get("source_bytes"), int) or side["source_bytes"] <= 0:
        raise RunnerError(f"{label}.source_bytes must be positive")
    if not isinstance(side.get("output_bytes"), int) or side["output_bytes"] <= 0:
        raise RunnerError(f"{label}.output_bytes must be positive")
    _require_hash(side.get("output_sha256"), f"{label}.output_sha256")


def _validate_timing(timing: Mapping[str, Any], mode: str, label: str) -> None:
    if not isinstance(timing, Mapping):
        raise RunnerError(f"{label}.timing is missing")
    if mode == "validate":
        if timing.get("status") != "not_run":
            raise RunnerError(f"{label}.timing status must be not_run in validate mode")
        if tuple(timing.get("clock_phases", ())) != PHASES:
            raise RunnerError(f"{label}.timing clock phases differ")
        if timing.get("max_iterations") != 128:
            raise RunnerError(f"{label}.timing max_iterations differs")
        if tuple(timing.get("setup_outside_clock", ())) != SETUP_PHASES:
            raise RunnerError(f"{label}.timing setup boundary differs")
        return
    if mode == "allocator":
        if timing.get("status") != "allocator_observed":
            raise RunnerError(f"{label}.timing status must be allocator_observed in allocator mode")
        if timing.get("elapsed_claimable") is not False or timing.get("allocator_enabled") is not True:
            raise RunnerError(f"{label}.timing allocator elapsed/enablement contract is invalid")
    elif timing.get("status") != "measured":
        raise RunnerError(f"{label}.timing status must be measured in timed mode")
    elif timing.get("elapsed_claimable") is not True or timing.get("allocator_enabled") is not False:
        raise RunnerError(f"{label}.timing normal elapsed/allocator contract is invalid")
    for side_name in ("leaf", "facade"):
        measured = timing.get(side_name)
        if not isinstance(measured, Mapping) or not isinstance(measured.get("iterations"), int):
            raise RunnerError(f"{label}.timing.{side_name} is incomplete")
        if measured["iterations"] != 128:
            raise RunnerError(f"{label}.timing.{side_name} iteration count must be exactly 128")
        phases = measured.get("ns_per_operation")
        if not isinstance(phases, Mapping) or tuple(phases) != PHASES:
            raise RunnerError(f"{label}.timing.{side_name} phases differ")
        if any(not isinstance(value, int) or value < 0 for value in phases.values()):
            raise RunnerError(f"{label}.timing.{side_name} contains invalid values")
        if mode == "allocator":
            samples = measured.get("operation_allocator_samples")
            if not isinstance(samples, Mapping) or tuple(samples) != PHASES:
                raise RunnerError(f"{label}.timing.{side_name} allocator phases differ")
            for phase in PHASES:
                values = samples[phase]
                if not isinstance(values, list) or len(values) != measured["iterations"]:
                    raise RunnerError(f"{label}.timing.{side_name}.{phase} allocator samples are incomplete")
                for sample in values:
                    if not isinstance(sample, Mapping):
                        raise RunnerError(f"{label}.timing.{side_name}.{phase} allocator sample is malformed")
                    for counter in (
                        "allocations",
                        "deallocations",
                        "allocated_bytes",
                        "deallocated_bytes",
                        "reallocations",
                    ):
                        if not isinstance(sample.get(counter), int) or sample[counter] < 0:
                            raise RunnerError(
                                f"{label}.timing.{side_name}.{phase}.{counter} is not a non-negative integer"
                            )
        elif measured.get("operation_allocator_samples") != {
            "commit_only": [],
            "commit_plus_diagnostics": [],
            "publication": [],
        }:
            raise RunnerError(f"{label}.timing.{side_name} unexpectedly contains allocator samples")


_MANIFEST_SEMANTIC_KEYS = {
    "root_attributes",
    "ordered_entries",
    "normalized_entries",
    "other_events",
}
_ARCHIVE_MEMBER_KEYS = {
    "name",
    "name_raw_sha256",
    "content_sha256",
    "content_bytes",
    "compression",
    "compressed_bytes",
    "crc32",
    "version_needed",
    "flags",
    "compression_options",
    "external_attributes",
    "is_utf8",
    "using_data_descriptor",
    "large_file",
    "central_extra_sha256",
    "last_modified",
    "extra_sha256",
    "comment",
    "unix_mode",
    "version_made_by",
    "encrypted",
    "content_text",
}


def _validate_manifest_semantics(value: Any, label: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping) or set(value) != _MANIFEST_SEMANTIC_KEYS:
        raise RunnerError(f"{label} has an incomplete normalized manifest schema")
    root = value["root_attributes"]
    if not isinstance(root, list) or not root:
        raise RunnerError(f"{label}.root_attributes must be non-empty")

    def attributes(items: Any, location: str) -> None:
        if not isinstance(items, list) or not items:
            raise RunnerError(f"{location} must be a non-empty attribute list")
        keys: set[str] = set()
        for pair in items:
            if not isinstance(pair, list) or len(pair) != 2 or not all(isinstance(item, str) for item in pair):
                raise RunnerError(f"{location} contains a malformed attribute")
            if pair[0] in keys:
                raise RunnerError(f"{location} contains a duplicate attribute")
            keys.add(pair[0])

    attributes(root, f"{label}.root_attributes")
    ordered = value["ordered_entries"]
    normalized = value["normalized_entries"]
    if not isinstance(ordered, list) or not ordered:
        raise RunnerError(f"{label}.ordered_entries must be non-empty")
    if not isinstance(normalized, list) or normalized != sorted(ordered):
        raise RunnerError(f"{label}.normalized_entries is not the complete sorted entry set")
    paths: list[str] = []
    for index, entry in enumerate(ordered):
        attributes(entry, f"{label}.ordered_entries[{index}]")
        full_paths = [value for key, value in entry if key == "manifest:full-path"]
        if len(full_paths) != 1:
            raise RunnerError(f"{label}.ordered_entries[{index}] has no manifest:full-path")
        paths.append(full_paths[0])
    if len(set(paths)) != len(paths):
        raise RunnerError(f"{label}.ordered_entries contains duplicate manifest:full-path values")
    other = value["other_events"]
    if not isinstance(other, list) or any(not isinstance(item, str) for item in other):
        raise RunnerError(f"{label}.other_events is malformed")
    return value


def _validate_archive_member_identity(value: Any, label: str, expected_name: str) -> None:
    if not isinstance(value, Mapping) or set(value) != _ARCHIVE_MEMBER_KEYS:
        raise RunnerError(f"{label} has an incomplete ZIP member identity")
    if value.get("name") != expected_name or not isinstance(value.get("name"), str):
        raise RunnerError(f"{label}.name differs from the audited member")
    for key in ("name_raw_sha256", "content_sha256"):
        _require_hash(value.get(key), f"{label}.{key}")
    for key in ("content_bytes", "compressed_bytes", "version_needed", "flags", "external_attributes"):
        if not isinstance(value.get(key), int) or value[key] < 0:
            raise RunnerError(f"{label}.{key} is invalid")
    if not isinstance(value.get("compression"), str) or not isinstance(value.get("compression_options"), str):
        raise RunnerError(f"{label} compression identity is incomplete")
    if not isinstance(value.get("crc32"), str) or not re.fullmatch(r"[0-9a-f]{8}", value["crc32"]):
        raise RunnerError(f"{label}.crc32 is invalid")
    if not isinstance(value.get("comment"), str):
        raise RunnerError(f"{label}.comment is invalid")
    if value.get("last_modified") is not None and not isinstance(value.get("last_modified"), str):
        raise RunnerError(f"{label}.last_modified is invalid")
    for key in ("extra_sha256", "central_extra_sha256"):
        if value.get(key) is not None:
            _require_hash(value.get(key), f"{label}.{key}")
    if value.get("unix_mode") is not None and not isinstance(value.get("unix_mode"), int):
        raise RunnerError(f"{label}.unix_mode is invalid")
    version = value.get("version_made_by")
    if not isinstance(version, list) or len(version) != 2 or any(not isinstance(item, int) or item < 0 for item in version):
        raise RunnerError(f"{label}.version_made_by is invalid")
    for key in ("is_utf8", "using_data_descriptor", "large_file", "encrypted"):
        if not isinstance(value.get(key), bool):
            raise RunnerError(f"{label}.{key} is invalid")
    if value.get("content_text") is not None and not isinstance(value.get("content_text"), str):
        raise RunnerError(f"{label}.content_text is invalid")


def _validate_archive_member_pair(
    leaf: Mapping[str, Any],
    facade: Mapping[str, Any],
    *,
    name: str,
    decision: str,
    label: str,
) -> None:
    differences = {key for key in _ARCHIVE_MEMBER_KEYS if leaf[key] != facade[key]}
    if decision == "byte_exact":
        allowed: set[str] = set()
    elif name == "META-INF/manifest.xml":
        # Reordering manifest entries necessarily changes the XML payload and
        # its derived CRC/size fields.  Every other ZIP identity must remain
        # exactly equal; this is the only admitted order-only delta.
        allowed = {
            "content_sha256",
            "content_bytes",
            "compressed_bytes",
            "crc32",
            "content_text",
        }
        if not {"content_sha256", "content_text"}.issubset(differences):
            raise RunnerError(f"{label} manifest payload difference is not proven")
    else:
        allowed = set()
    unexpected = differences - allowed
    if unexpected:
        raise RunnerError(f"{label} ZIP identity differs in unadmitted fields: {sorted(unexpected)}")


def _manifest_paths(semantics: Mapping[str, Any], *, include_root: bool) -> list[str]:
    paths = [
        next(value for key, value in entry if key == "manifest:full-path")
        for entry in semantics["ordered_entries"]
    ]
    if include_root:
        return paths
    return [path for path in paths if path != "/"]


def _archive_payload_order(order: list[str]) -> list[str]:
    return [name for name in order if name not in {"mimetype", "META-INF/manifest.xml"}]


def _validate_archive_audit(audit: Mapping[str, Any], label: str) -> None:
    if audit.get("status") != "audited" or audit.get("allowed") is not True:
        raise RunnerError(f"{label}.archive_audit is not an allowed audit")
    decision = audit.get("decision")
    if decision not in {
        "byte_exact",
        "archive_member_order_and_manifest_listing",
    }:
        raise RunnerError(f"{label}.archive_audit decision is unknown")
    for key in (
        "archive_comment_equal",
        "meta_stored_metadata_exact",
        "non_manifest_stored_metadata_exact",
        "non_manifest_content_exact",
        "non_manifest_archive_record_exact",
    ):
        _require_bool(audit, key, True, f"{label}.archive_audit")
    for key in ("leaf_archive_comment_sha256", "facade_archive_comment_sha256"):
        _require_hash(audit.get(key), f"{label}.archive_audit.{key}")
    if audit["leaf_archive_comment_sha256"] != audit["facade_archive_comment_sha256"]:
        raise RunnerError(f"{label}.archive_audit archive comments differ")
    _require_bool(
        audit,
        "non_meta_stored_metadata_exact",
        decision != "archive_member_order_and_manifest_listing",
        f"{label}.archive_audit",
    )
    _require_bool(
        audit,
        "meta_archive_record_exact",
        True,
        f"{label}.archive_audit",
    )
    _require_bool(audit, "manifest_semantics_equal", True, f"{label}.archive_audit")
    leaf_semantics = _validate_manifest_semantics(
        audit.get("leaf_manifest_semantics"), f"{label}.archive_audit.leaf_manifest_semantics"
    )
    facade_semantics = _validate_manifest_semantics(
        audit.get("facade_manifest_semantics"), f"{label}.archive_audit.facade_manifest_semantics"
    )
    for key in ("root_attributes", "normalized_entries", "other_events"):
        if leaf_semantics[key] != facade_semantics[key]:
            raise RunnerError(f"{label}.archive_audit manifest {key} differs")
    ordered_equal = leaf_semantics["ordered_entries"] == facade_semantics["ordered_entries"]
    differences = audit.get("content_differences")
    expected_differences = {
        "byte_exact": [],
        "archive_member_order_and_manifest_listing": ["META-INF/manifest.xml"],
    }[decision]
    if differences != expected_differences:
        raise RunnerError(f"{label}.archive_audit content difference decision is inconsistent")
    if decision == "archive_member_order_and_manifest_listing":
        if ordered_equal:
            raise RunnerError(f"{label}.archive_audit order-only manifest entries did not differ")
    elif not ordered_equal:
        raise RunnerError(f"{label}.archive_audit byte-exact manifest entries differ")
    if audit.get("leaf_member_count") != audit.get("facade_member_count"):
        raise RunnerError(f"{label}.archive_audit member counts differ")
    leaf_order = audit.get("leaf_order")
    facade_order = audit.get("facade_order")
    if not isinstance(leaf_order, list) or not isinstance(facade_order, list):
        raise RunnerError(f"{label}.archive_audit member orders are missing")
    if len(leaf_order) != len(facade_order) or len(leaf_order) != audit["leaf_member_count"]:
        raise RunnerError(f"{label}.archive_audit member orders are incomplete")
    if len(set(leaf_order)) != len(leaf_order) or len(set(facade_order)) != len(facade_order):
        raise RunnerError(f"{label}.archive_audit member orders contain duplicates")
    if set(leaf_order) != set(facade_order):
        raise RunnerError(f"{label}.archive_audit member names differ")
    leaf_manifest_paths = _manifest_paths(leaf_semantics, include_root=True)
    facade_manifest_paths = _manifest_paths(facade_semantics, include_root=True)
    leaf_manifest_payload_paths = _manifest_paths(leaf_semantics, include_root=False)
    facade_manifest_payload_paths = _manifest_paths(facade_semantics, include_root=False)
    actual_manifest_listing_tracks_order = leaf_manifest_paths == facade_manifest_paths
    actual_leaf_manifest_tracks_order = leaf_manifest_payload_paths == _archive_payload_order(leaf_order)
    actual_facade_manifest_tracks_order = facade_manifest_payload_paths == _archive_payload_order(facade_order)
    for key, actual in (
        ("manifest_listing_tracks_order", actual_manifest_listing_tracks_order),
        ("leaf_manifest_tracks_order", actual_leaf_manifest_tracks_order),
        ("facade_manifest_tracks_order", actual_facade_manifest_tracks_order),
    ):
        _require_bool(audit, key, actual, f"{label}.archive_audit")
    _require_bool(
        audit,
        "archive_member_order_and_manifest_listing",
        decision == "archive_member_order_and_manifest_listing",
        f"{label}.archive_audit",
    )
    _require_bool(
        audit,
        "member_order_exact",
        decision != "archive_member_order_and_manifest_listing",
        f"{label}.archive_audit",
    )
    _require_bool(
        audit,
        "non_meta_content_exact",
        decision != "archive_member_order_and_manifest_listing",
        f"{label}.archive_audit",
    )
    if decision == "archive_member_order_and_manifest_listing":
        if leaf_order == facade_order or not actual_leaf_manifest_tracks_order or not actual_facade_manifest_tracks_order:
            raise RunnerError(f"{label}.archive_audit order-only decision has no order difference")
    elif leaf_order != facade_order:
        raise RunnerError(f"{label}.archive_audit non-order decision has an order difference")
    members = audit.get("members")
    if not isinstance(members, list) or not members:
        raise RunnerError(f"{label}.archive_audit has no member rows")
    names = [item.get("name") for item in members if isinstance(item, Mapping)]
    facade_names = [item.get("facade_name") for item in members if isinstance(item, Mapping)]
    if names != leaf_order or facade_names != leaf_order:
        raise RunnerError(f"{label}.archive_audit member rows do not follow the leaf order")
    if names.count("meta.xml") != 1:
        raise RunnerError(f"{label}.archive_audit must contain one meta.xml row")
    for item in members:
        if not isinstance(item, Mapping):
            raise RunnerError(f"{label}.archive_audit member row is malformed")
        if item.get("name") != item.get("facade_name"):
            raise RunnerError(f"{label}.archive_audit member order/name differs")
        if not isinstance(item.get("name"), str):
            raise RunnerError(f"{label}.archive_audit member name is malformed")
        _validate_archive_member_identity(item.get("leaf"), f"{label}.archive_audit.member.leaf", item["name"])
        _validate_archive_member_identity(item.get("facade"), f"{label}.archive_audit.member.facade", item["name"])
        _validate_archive_member_pair(
            item["leaf"],
            item["facade"],
            name=item["name"],
            decision=decision,
            label=f"{label}.archive_audit.member[{item['name']}]",
        )
        expected_content = {
            "byte_exact": True,
            "archive_member_order_and_manifest_listing": item["name"] != "META-INF/manifest.xml",
        }[audit["decision"]]
        expected_record = not (
            audit["decision"] == "archive_member_order_and_manifest_listing"
            and item["name"] == "META-INF/manifest.xml"
        )
        _require_bool(item, "content_equal", expected_content, f"{label}.archive_audit.member")
        expected_metadata = not (
            audit["decision"] == "archive_member_order_and_manifest_listing"
            and item["name"] == "META-INF/manifest.xml"
        )
        _require_bool(item, "stored_metadata_equal", expected_metadata, f"{label}.archive_audit.member")
        _require_bool(item, "archive_record_equal", expected_record, f"{label}.archive_audit.member")


def validate_probe_output(
    manifest: Mapping[str, Any],
    payload: Mapping[str, Any],
    expected_mode: str | None = None,
) -> None:
    if payload.get("schema") != OUTPUT_SCHEMA:
        raise RunnerError("probe output schema differs from the manifest")
    mode = payload.get("mode")
    if mode not in ("validate", "timed", "allocator"):
        raise RunnerError("probe output mode is invalid")
    if expected_mode is not None and mode != expected_mode:
        raise RunnerError(f"probe output mode {mode!r} differs from requested mode {expected_mode!r}")
    iterations = payload.get("iterations")
    if not isinstance(iterations, int) or not 1 <= iterations <= 128:
        raise RunnerError("probe output iterations must be an integer in 1..=128")
    if mode in ("timed", "allocator") and iterations != 128:
        raise RunnerError("timed and allocator probe outputs must use exactly 128 iterations")
    if payload.get("claim_scope") != CLAIM_SCOPE:
        raise RunnerError("probe output claim scope is not non-acceptance")
    rows = payload.get("families")
    if not isinstance(rows, list) or len(rows) != len(FAMILIES):
        raise RunnerError("probe output must contain exactly four family rows")
    route_map = {item["family"]: item for item in manifest["routes"]}
    for index, row in enumerate(rows):
        if not isinstance(row, Mapping) or row.get("family") != FAMILIES[index]:
            raise RunnerError("probe output family order differs from the manifest")
        family = row["family"]
        route = route_map[family]
        if row.get("route") != route["id"] or row.get("source") != route["source"]:
            raise RunnerError(f"{family} route/source identity differs")
        if row.get("descriptor") != route:
            raise RunnerError(f"{family} route descriptor differs")
        if row.get("source_pair_exact") is not True:
            raise RunnerError(f"{family} native/facade source bytes were not proven equal")
        _require_hash(row.get("source_sha256"), f"{family}.source_sha256")
        if not isinstance(row.get("source_bytes"), int) or row["source_bytes"] <= 0:
            raise RunnerError(f"{family}.source_bytes is invalid")
        _require_hash(row.get("facade_source_sha256"), f"{family}.facade_source_sha256")
        if not isinstance(row.get("facade_source_bytes"), int) or row["facade_source_bytes"] <= 0:
            raise RunnerError(f"{family}.facade_source_bytes is invalid")
        source_pair_exact = (
            row["source_sha256"] == row["facade_source_sha256"]
            and row["source_bytes"] == row["facade_source_bytes"]
        )
        if row.get("source_pair_exact") is not source_pair_exact:
            raise RunnerError(f"{family} source_pair_exact is not bound to both source identities")
        for side_name in ("leaf", "facade"):
            side = row.get(side_name)
            if not isinstance(side, Mapping):
                raise RunnerError(f"{family}.{side_name} is missing")
            changed_side = side.get("changed")
            noop_side = side.get("no_op")
            if not isinstance(changed_side, Mapping) or not isinstance(noop_side, Mapping):
                raise RunnerError(f"{family}.{side_name} side rows are missing")
            _validate_side(changed_side, True, f"{family}.{side_name}.changed")
            _validate_side(noop_side, False, f"{family}.{side_name}.no_op")
        paired = row.get("paired")
        if not isinstance(paired, Mapping):
            raise RunnerError(f"{family}.paired is missing")
        _require_bool(paired, "semantic_equality", True, f"{family}.paired")
        leaf_changed = row["leaf"]["changed"]
        facade_changed = row["facade"]["changed"]
        _require_hash(paired.get("changed_output_sha256"), f"{family}.paired.changed_output_sha256")
        _require_hash(paired.get("facade_changed_output_sha256"), f"{family}.paired.facade_changed_output_sha256")
        if paired["changed_output_sha256"] != leaf_changed["output_sha256"]:
            raise RunnerError(f"{family}.paired.changed_output_sha256 is not the leaf changed output")
        if paired["facade_changed_output_sha256"] != facade_changed["output_sha256"]:
            raise RunnerError(f"{family}.paired.facade_changed_output_sha256 is not the facade changed output")
        changed_output_exact = (
            leaf_changed["output_sha256"] == facade_changed["output_sha256"]
            and leaf_changed["output_bytes"] == facade_changed["output_bytes"]
        )
        if paired.get("changed_output_exact") is not changed_output_exact:
            raise RunnerError(f"{family}.paired.changed_output_exact is not bound to output identities")
        if family != "odm" and not changed_output_exact:
            raise RunnerError(f"{family} changed outputs are not byte-identical")
        if family == "odm":
            audit = paired.get("archive_audit")
            if not isinstance(audit, Mapping):
                raise RunnerError("ODM archive audit is missing")
            _validate_archive_audit(audit, family)
        elif paired.get("archive_audit") is not None:
            raise RunnerError(f"{family} must not claim an ODM archive audit")
        _validate_timing(row.get("timing", {}), mode, family)


def _fresh_external(path: Path, source: Path, label: str) -> None:
    _ensure_external(path, source, label)
    _reject_symlink_ancestors(path, label)
    if path.exists():
        if path.is_symlink() or (path.is_dir() and any(path.iterdir())):
            raise RunnerError(f"{label} must be absent or empty: {path}")
        if path.is_file():
            raise RunnerError(f"{label} must be a fresh directory: {path}")
    path.mkdir(parents=True, exist_ok=True)


def _write_new_text(path: Path, text: str, source: Path, label: str) -> None:
    _ensure_external(path, source, label)
    _reject_symlink_ancestors(path, label)
    if path.is_symlink() or path.exists():
        raise RunnerError(f"{label} already exists: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        stream.write(text)
        stream.flush()
        os.fsync(stream.fileno())


def _assert_disjoint(paths: Sequence[tuple[Path, str]]) -> None:
    for index, (left, left_label) in enumerate(paths):
        for right, right_label in paths[index + 1 :]:
            if left == right or _under(left, right) or _under(right, left):
                raise RunnerError(f"{left_label} and {right_label} paths overlap")


def _cargo_command(manifest: Mapping[str, Any], mode: str) -> list[str]:
    probe_name = Path(manifest["probe"]["path"]).stem
    features = ",".join(manifest["probe"]["features"])
    return [
        "cargo",
        "run",
        "--offline",
        "--locked",
        "-p",
        "litchi-ops",
        "--example",
        probe_name,
        "--features",
        features,
        "--",
        "--mode",
        mode,
        "--family",
        "all",
    ]


def _limited_child() -> None:
    resource.setrlimit(resource.RLIMIT_AS, (MEMORY_HARD_LIMIT_BYTES, MEMORY_HARD_LIMIT_BYTES))


def run_probe(
    manifest: Mapping[str, Any],
    *,
    role: str,
    source: dict[str, Any],
    mode: str,
    target_dir: Path,
    tmpdir: Path,
    output_path: Path,
    stderr_path: Path,
) -> dict[str, Any]:
    root = Path(source["root"])
    _fresh_external(target_dir, root, "target directory")
    _fresh_external(tmpdir, root, "temporary directory")
    for path, label in ((output_path, "probe output"), (stderr_path, "probe stderr")):
        _ensure_external(path, root, label)
        _reject_symlink_ancestors(path, label)
        if path.exists():
            raise RunnerError(f"{label} already exists: {path}")
        path.parent.mkdir(parents=True, exist_ok=True)
    command = _cargo_command(manifest, mode)
    environment = os.environ.copy()
    environment.update(
        {
            "CARGO_TARGET_DIR": str(target_dir),
            "CARGO_BUILD_JOBS": "1",
            "CARGO_INCREMENTAL": "0",
            "TMPDIR": str(tmpdir),
        }
    )
    completed = subprocess.run(
        command,
        cwd=root,
        env=environment,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
        preexec_fn=_limited_child,
    )
    with stderr_path.open("xb") as stream:
        stream.write(completed.stderr)
        stream.flush()
        os.fsync(stream.fileno())
    with output_path.open("xb") as stream:
        stream.write(completed.stdout)
        stream.flush()
        os.fsync(stream.fileno())
    if completed.returncode:
        raise RunnerError(f"probe exited {completed.returncode}; see {stderr_path}")
    if len(completed.stdout) == 0 or len(completed.stdout) > MAX_OUTPUT_BYTES:
        raise RunnerError("probe stdout is empty or exceeds the reduced receipt bound")
    try:
        payload = json.loads(completed.stdout)
    except json.JSONDecodeError as error:
        raise RunnerError(f"probe stdout is not one JSON object: {error}") from error
    if not isinstance(payload, Mapping):
        raise RunnerError("probe stdout root must be an object")
    validate_probe_output(manifest, payload, mode)
    target_bytes = _directory_size(target_dir)
    if target_bytes > TARGET_PREFLIGHT_BYTES:
        raise RunnerError("target directory exceeded the fixed 19.2 GB preflight limit")
    if target_bytes > TARGET_HARD_LIMIT_BYTES:
        raise RunnerError("target directory exceeded the fixed 20 GB hard limit")
    output_hash, output_bytes = _sha256_file(output_path)
    return {
        "command": command,
        "phase": "run",
        "mode": mode,
        "output": {"path": str(output_path), "sha256": output_hash, "bytes": output_bytes},
        "stderr": {
            "path": str(stderr_path),
            "sha256": _sha256_bytes(completed.stderr),
            "bytes": len(completed.stderr),
        },
        "target": {"path": str(target_dir), "bytes": _directory_size(target_dir)},
    }


def make_receipt(manifest_bytes: bytes, source: Mapping[str, Any], run: Mapping[str, Any] | None) -> dict[str, Any]:
    return {
        "schema": RECEIPT_SCHEMA,
        "lane": "odf-shared-triad",
        "manifest_sha256": _sha256_bytes(manifest_bytes),
        "source": source,
        "run": run,
        "claim_scope": "orchestration identity only; no native acceptance claim",
    }


def validate_receipt_bindings(
    manifest: Mapping[str, Any],
    receipt: Mapping[str, Any],
    manifest_bytes: bytes | None = None,
) -> None:
    if receipt.get("schema") != RECEIPT_SCHEMA or receipt.get("lane") != "odf-shared-triad":
        raise RunnerError("receipt schema/lane differs")
    _require_hash(receipt.get("manifest_sha256"), "receipt.manifest_sha256")
    if manifest_bytes is not None and receipt["manifest_sha256"] != _sha256_bytes(manifest_bytes):
        raise RunnerError("receipt manifest identity differs from the supplied manifest")
    source = receipt.get("source")
    if not isinstance(source, Mapping) or source.get("clean") is not True:
        raise RunnerError("receipt source identity is not clean")
    if not FULL_SHA.fullmatch(str(source.get("commit", ""))) or not FULL_SHA.fullmatch(str(source.get("tree", ""))):
        raise RunnerError("receipt source commit/tree identity is malformed")
    _require_hash(source.get("source_tree_sha256"), "receipt.source_tree_sha256")
    tracked_closure = source.get("tracked_closure")
    if not isinstance(tracked_closure, Mapping):
        raise RunnerError("receipt tracked source closure is missing")
    for key in ("sha256", "flags_sha256", "index_sha256"):
        _require_hash(tracked_closure.get(key), f"receipt.tracked_closure.{key}")
    if not isinstance(tracked_closure.get("file_count"), int) or tracked_closure["file_count"] <= 0:
        raise RunnerError("receipt.tracked_closure.file_count is invalid")
    cargo_configs = source.get("cargo_configs")
    if not isinstance(cargo_configs, Mapping):
        raise RunnerError("receipt Cargo configuration closure is missing")
    _require_hash(cargo_configs.get("sha256"), "receipt.cargo_configs.sha256")
    config_files = cargo_configs.get("files")
    if not isinstance(config_files, list):
        raise RunnerError("receipt Cargo configuration files are malformed")
    for index, config in enumerate(config_files):
        if not isinstance(config, Mapping):
            raise RunnerError(f"receipt.cargo_configs.files[{index}] is malformed")
        path = config.get("path")
        if not isinstance(path, str) or not Path(path).is_absolute() or UPSTREAM_COMPONENT in Path(path).parts:
            raise RunnerError(f"receipt.cargo_configs.files[{index}] path is invalid")
        _require_hash(config.get("sha256"), f"receipt.cargo_configs.files[{index}].sha256")
        if not isinstance(config.get("bytes"), int) or config["bytes"] < 0:
            raise RunnerError(f"receipt.cargo_configs.files[{index}].bytes is invalid")
    package_manifest = source.get("package_manifest")
    if not isinstance(package_manifest, Mapping) or package_manifest.get("name") != "litchi-ops":
        raise RunnerError("receipt package manifest identity is missing")
    _require_hash(package_manifest.get("sha256"), "receipt.package_manifest.sha256")
    if not isinstance(package_manifest.get("path"), str) or not package_manifest["path"].endswith(
        manifest["source"]["package_manifest"]
    ):
        raise RunnerError("receipt package manifest path identity differs from manifest")
    probe = source.get("probe")
    lock = source.get("lock")
    if not isinstance(probe, Mapping) or probe.get("sha256") != manifest["probe"]["sha256"]:
        raise RunnerError("receipt probe identity differs from manifest")
    if not isinstance(probe.get("path"), str) or not probe["path"].endswith(manifest["probe"]["path"]):
        raise RunnerError("receipt probe path identity differs from manifest")
    if not isinstance(lock, Mapping) or not HASH.fullmatch(str(lock.get("sha256", ""))):
        raise RunnerError("receipt lock identity is missing")
    if not isinstance(lock.get("path"), str) or not lock["path"].endswith(manifest["source"]["lock_path"]):
        raise RunnerError("receipt lock path identity differs from manifest")
    run = receipt.get("run")
    if run is not None:
        if not isinstance(run, Mapping) or run.get("mode") not in ("validate", "timed", "allocator"):
            raise RunnerError("receipt run identity is malformed")
        output = run.get("output")
        if not isinstance(output, Mapping) or not HASH.fullmatch(str(output.get("sha256", ""))):
            raise RunnerError("receipt output identity is missing")
        output_path = _resolve(str(output.get("path")))
        observed_hash, observed_bytes = _sha256_file(output_path)
        if observed_hash != output["sha256"] or observed_bytes != output.get("bytes"):
            raise RunnerError("receipt output identity does not match the output file")
        if run.get("phase") == "run":
            stderr = run.get("stderr")
            if not isinstance(stderr, Mapping) or not HASH.fullmatch(str(stderr.get("sha256", ""))):
                raise RunnerError("receipt stderr identity is missing")
            stderr_path = _resolve(str(stderr.get("path")))
            observed_stderr_hash, observed_stderr_bytes = _sha256_file(stderr_path)
            if observed_stderr_hash != stderr["sha256"] or observed_stderr_bytes != stderr.get("bytes"):
                raise RunnerError("receipt stderr identity does not match the stderr file")
        elif run.get("phase") != "validate":
            raise RunnerError("receipt run phase identity is missing")


def _parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--phase", choices=("validate", "run"), required=True)
    parser.add_argument("--role", choices=("baseline", "candidate"), required=True)
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--source-commit")
    parser.add_argument("--lock-sha256")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--receipt", type=Path, required=True)
    parser.add_argument("--target-dir", type=Path)
    parser.add_argument("--tmpdir", type=Path)
    parser.add_argument("--stderr", type=Path)
    parser.add_argument("--mode", choices=("validate", "timed", "allocator"), default="validate")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = _parse_args(argv)
    try:
        manifest_path = _resolve(
            args.manifest,
            base=Path.cwd() if not args.manifest.is_absolute() else None,
        )
        manifest, manifest_bytes = load_manifest(manifest_path)
        source = source_identity(
            manifest,
            role=args.role,
            source_root=args.source_root,
            source_commit=args.source_commit,
            lock_sha256=args.lock_sha256,
        )
        if not _under(manifest_path, Path(source["root"])):
            raise RunnerError("manifest must be inside the selected source worktree")
        output_path = _resolve(args.output)
        receipt_path = _resolve(args.receipt)
        _ensure_external(output_path, Path(source["root"]), "probe output")
        _assert_disjoint(((output_path, "probe output"), (receipt_path, "receipt")))
        if args.phase == "run":
            if args.target_dir is None or args.tmpdir is None or args.stderr is None:
                raise RunnerError("run phase requires --target-dir, --tmpdir and --stderr")
            target_path = _resolve(args.target_dir)
            tmp_path = _resolve(args.tmpdir)
            stderr_path = _resolve(args.stderr)
            _assert_disjoint(
                (
                    (target_path, "target directory"),
                    (tmp_path, "temporary directory"),
                    (output_path, "probe output"),
                    (stderr_path, "probe stderr"),
                    (receipt_path, "receipt"),
                )
            )
            run = run_probe(
                manifest,
                role=args.role,
                source=source,
                mode=args.mode,
                target_dir=target_path,
                tmpdir=tmp_path,
                output_path=output_path,
                stderr_path=stderr_path,
            )
        else:
            output_bytes = output_path.read_bytes()
            if len(output_bytes) == 0 or len(output_bytes) > MAX_OUTPUT_BYTES:
                raise RunnerError("probe output is empty or exceeds the reduced receipt bound")
            payload = json.loads(output_bytes)
            if not isinstance(payload, Mapping):
                raise RunnerError("probe output root must be an object")
            validate_probe_output(manifest, payload, args.mode)
            run = {
                "phase": "validate",
                "mode": payload["mode"],
                "output": {
                    "path": str(output_path),
                    "sha256": _sha256_bytes(output_bytes),
                    "bytes": len(output_bytes),
                },
            }
        receipt = make_receipt(manifest_bytes, source, run)
        validate_receipt_bindings(manifest, receipt, manifest_bytes)
        _ensure_external(receipt_path, Path(source["root"]), "receipt")
        _write_new_text(receipt_path, json.dumps(receipt, indent=2) + "\n", Path(source["root"]), "receipt")
        return 0
    except (RunnerError, OSError, json.JSONDecodeError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
