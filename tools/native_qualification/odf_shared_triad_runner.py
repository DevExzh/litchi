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
    if probe.get("modes") != ["validate", "timed"]:
        raise RunnerError("probe modes must be validate and timed")
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
    if tuple(item.get("id") for item in routes) != ROUTE_IDS:
        raise RunnerError(f"routes must be the reviewed ordered routes {list(ROUTE_IDS)!r}")
    if tuple(item.get("family") for item in routes) != FAMILIES:
        raise RunnerError("route families must match the reviewed order")
    for item in routes:
        if not isinstance(item, Mapping) or not isinstance(item.get("leaf"), str) or not isinstance(item.get("facade"), str):
            raise RunnerError("each route must declare leaf and facade lifecycles")

    phases = data.get("phases")
    if not isinstance(phases, Mapping):
        raise RunnerError("phases must be an object")
    if tuple(phases.get("inside_clock", ())) != PHASES:
        raise RunnerError("phases.inside_clock does not match the shared-triad contract")
    if tuple(phases.get("setup_outside_clock", ())) != SETUP_PHASES:
        raise RunnerError("phases.setup_outside_clock does not match the timed boundary")
    allocation = phases.get("allocation_observation")
    if not isinstance(allocation, Mapping) or allocation.get("facade") != "Commit::diagnostics().snapshot_bytes":
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
    source_tree_sha256 = _sha256_bytes((tree + "\0" + status + "\0" + diff).encode())
    return {
        "role": role,
        "root": str(root),
        "commit": head,
        "tree": tree,
        "source_tree_sha256": source_tree_sha256,
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
    if timing.get("status") != "measured":
        raise RunnerError(f"{label}.timing status must be measured in timed mode")
    for side_name in ("leaf", "facade"):
        measured = timing.get(side_name)
        if not isinstance(measured, Mapping) or not isinstance(measured.get("iterations"), int):
            raise RunnerError(f"{label}.timing.{side_name} is incomplete")
        phases = measured.get("ns_per_operation")
        if not isinstance(phases, Mapping) or tuple(phases) != PHASES:
            raise RunnerError(f"{label}.timing.{side_name} phases differ")
        if any(not isinstance(value, int) or value < 0 for value in phases.values()):
            raise RunnerError(f"{label}.timing.{side_name} contains invalid values")


def _validate_archive_audit(audit: Mapping[str, Any], label: str) -> None:
    if audit.get("status") != "audited" or audit.get("allowed") is not True:
        raise RunnerError(f"{label}.archive_audit is not an allowed audit")
    decision = audit.get("decision")
    if decision not in {
        "byte_exact",
        "metadata_member_only",
        "archive_member_order_and_manifest_listing",
    }:
        raise RunnerError(f"{label}.archive_audit decision is unknown")
    for key in (
        "archive_comment_equal",
        "non_meta_stored_metadata_exact",
        "meta_stored_metadata_exact",
        "non_manifest_content_exact",
        "non_manifest_archive_record_exact",
    ):
        _require_bool(audit, key, True, f"{label}.archive_audit")
    _require_bool(
        audit,
        "meta_archive_record_exact",
        decision != "metadata_member_only",
        f"{label}.archive_audit",
    )
    _require_bool(audit, "member_order_exact", decision != "archive_member_order_and_manifest_listing", f"{label}.archive_audit")
    _require_bool(audit, "non_meta_content_exact", decision != "archive_member_order_and_manifest_listing", f"{label}.archive_audit")
    differences = audit.get("content_differences")
    expected_differences = {
        "byte_exact": [],
        "metadata_member_only": ["meta.xml"],
        "archive_member_order_and_manifest_listing": ["META-INF/manifest.xml"],
    }[decision]
    if differences != expected_differences:
        raise RunnerError(f"{label}.archive_audit content difference decision is inconsistent")
    if decision == "archive_member_order_and_manifest_listing":
        for key, expected in (
            ("manifest_listing_tracks_order", False),
            ("leaf_manifest_tracks_order", True),
            ("facade_manifest_tracks_order", True),
            ("archive_member_order_and_manifest_listing", True),
        ):
            _require_bool(audit, key, expected, f"{label}.archive_audit")
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
    if decision == "archive_member_order_and_manifest_listing":
        if leaf_order == facade_order:
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
        expected_content = (
            audit["decision"] == "byte_exact"
            or item["name"] not in {"meta.xml", "META-INF/manifest.xml"}
        )
        expected_record = item["name"] != "meta.xml" or audit["decision"] != "metadata_member_only"
        if audit["decision"] == "archive_member_order_and_manifest_listing" and item["name"] == "META-INF/manifest.xml":
            expected_record = False
        _require_bool(item, "content_equal", expected_content, f"{label}.archive_audit.member")
        _require_bool(item, "stored_metadata_equal", True, f"{label}.archive_audit.member")
        _require_bool(item, "archive_record_equal", expected_record, f"{label}.archive_audit.member")
    if audit["decision"] == "metadata_member_only" and differences != ["meta.xml"]:
        raise RunnerError(f"{label}.archive_audit metadata-member decision is not exact")


def validate_probe_output(manifest: Mapping[str, Any], payload: Mapping[str, Any]) -> None:
    if payload.get("schema") != OUTPUT_SCHEMA:
        raise RunnerError("probe output schema differs from the manifest")
    mode = payload.get("mode")
    if mode not in ("validate", "timed"):
        raise RunnerError("probe output mode is invalid")
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
        if row.get("source_pair_exact") is not True:
            raise RunnerError(f"{family} native/facade source bytes were not proven equal")
        _require_hash(row.get("source_sha256"), f"{family}.source_sha256")
        if not isinstance(row.get("source_bytes"), int) or row["source_bytes"] <= 0:
            raise RunnerError(f"{family}.source_bytes is invalid")
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
        if family != "odm":
            _require_bool(paired, "changed_output_exact", True, f"{family}.paired")
        elif not isinstance(paired.get("changed_output_exact"), bool):
            raise RunnerError("ODM paired.changed_output_exact must be a boolean")
        _require_hash(paired.get("changed_output_sha256"), f"{family}.paired.changed_output_sha256")
        _require_hash(paired.get("facade_changed_output_sha256"), f"{family}.paired.facade_changed_output_sha256")
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
    if path.exists():
        if path.is_symlink() or (path.is_dir() and any(path.iterdir())):
            raise RunnerError(f"{label} must be absent or empty: {path}")
        if path.is_file():
            raise RunnerError(f"{label} must be a fresh directory: {path}")
    path.mkdir(parents=True, exist_ok=True)


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
    stderr_path.write_bytes(completed.stderr)
    output_path.write_bytes(completed.stdout)
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
    validate_probe_output(manifest, payload)
    target_bytes = _directory_size(target_dir)
    if target_bytes > TARGET_PREFLIGHT_BYTES:
        raise RunnerError("target directory exceeded the fixed 19.2 GB preflight limit")
    if target_bytes > TARGET_HARD_LIMIT_BYTES:
        raise RunnerError("target directory exceeded the fixed 20 GB hard limit")
    output_hash, output_bytes = _sha256_file(output_path)
    return {
        "command": command,
        "mode": mode,
        "output": {"path": str(output_path), "sha256": output_hash, "bytes": output_bytes},
        "stderr": {"path": str(stderr_path), "bytes": len(completed.stderr)},
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
        if not isinstance(run, Mapping) or run.get("mode") not in ("validate", "timed"):
            raise RunnerError("receipt run identity is malformed")
        output = run.get("output")
        if not isinstance(output, Mapping) or not HASH.fullmatch(str(output.get("sha256", ""))):
            raise RunnerError("receipt output identity is missing")
        output_path = _resolve(str(output.get("path")))
        observed_hash, observed_bytes = _sha256_file(output_path)
        if observed_hash != output["sha256"] or observed_bytes != output.get("bytes"):
            raise RunnerError("receipt output identity does not match the output file")


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
    parser.add_argument("--mode", choices=("validate", "timed"), default="validate")
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
            validate_probe_output(manifest, payload)
            run = {
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
        receipt_path.parent.mkdir(parents=True, exist_ok=True)
        receipt_path.write_text(json.dumps(receipt, indent=2) + "\n", encoding="utf-8")
        return 0
    except (RunnerError, OSError, json.JSONDecodeError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
