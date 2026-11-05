#!/usr/bin/env python3
"""Fail-closed matched native qualification for the ODG joined API.

The historical endpoint exposes ``Package::apply_joined``. The candidate
endpoint exposes ``Edit::join`` followed by the ordinary commit triad. This
runner never builds either endpoint: a frozen manifest supplies source-specific
probes and records their external provenance. It reuses the generic runner's
strict JSON, Git, process and load guards while validating the ODG protocol
independently.

``--plan`` validates only the manifest shape. ``--correctness-only`` runs both
untimed endpoint proofs. The default run performs correctness, calibration,
and pinned A1/B1/B2/A2 ABBA for both measured scopes of every route.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from dataclasses import replace
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import time
from typing import Any, Mapping

try:
    from tools.odp_native_qualification import (
        MAX_CHILD_OUTPUT,
        Protocol,
        RunFailure,
        _digest,
        _load_json,
        _nonnegative_int,
        _object,
        _path,
        _positive_float,
        _positive_int,
        _parse_child_json,
        _run_git,
        _string,
        _require_taskset,
        _validate_external_target,
        binary_identity,
        quiet_guard,
    )
    from tools.odg_joined_qualification_bindings import (
        REVIEWED_FIXTURE_IDENTITY,
        REVIEWED_FIXTURE_PROVENANCE,
        REVIEWED_FIXTURE_RELATIVE,
        REVIEWED_PROBE_IDENTITIES,
        REVIEWED_PROBE_RELATIVES,
        REVIEWED_REPO_ROOT,
    )
except ImportError:  # pragma: no cover - supports direct test-module loading.
    from odp_native_qualification import (  # type: ignore[no-redef]
        MAX_CHILD_OUTPUT,
        Protocol,
        RunFailure,
        _digest,
        _load_json,
        _nonnegative_int,
        _object,
        _path,
        _positive_float,
        _positive_int,
        _parse_child_json,
        _run_git,
        _string,
        _require_taskset,
        _validate_external_target,
        binary_identity,
        quiet_guard,
    )
    from odg_joined_qualification_bindings import (  # type: ignore[no-redef]
        REVIEWED_FIXTURE_IDENTITY,
        REVIEWED_FIXTURE_PROVENANCE,
        REVIEWED_FIXTURE_RELATIVE,
        REVIEWED_PROBE_IDENTITIES,
        REVIEWED_PROBE_RELATIVES,
        REVIEWED_REPO_ROOT,
    )


INPUT_SCHEMA = "litchi.odg.joined-native-qualification.input.v3"
REPORT_SCHEMA = "litchi.odg.joined-native-qualification.v3"
UNTRACKED_UPSTREAM_COMPONENT = "3rd" + "party"
BASELINE_SCHEMA = "litchi.odg.joined-native-qualification.legacy-probe.v2"
CANDIDATE_SCHEMA = "litchi.odg.joined-native-qualification.probe.v2"
BASELINE_PROFILE = "legacy-apply-joined"
CANDIDATE_PROFILE = "current-edit-join-commit"
ROUTES = ("package_replace_one", "semantic_two_disjoint", "semantic_four_disjoint", "empty_join")
SCOPES = ("join_apply", "join_apply_publish")
ABBA_ORDER = ("a1", "b1", "b2", "a2")
ROLE_ENDPOINT = {"a1": "baseline", "a2": "baseline", "b1": "candidate", "b2": "candidate"}
REVISION = re.compile(r"[0-9a-f]{40}\Z")
REQUIRED_BLOCKS = 2
REQUIRED_WARMUPS = 20
REQUIRED_SAMPLES = 1
CALIBRATION_POLICY = "adaptive-doubling-v1"
MAX_MANIFEST_BYTES = 256 * 1024
MAX_FIXTURE_BYTES = 4 * 1024 * 1024
TARGET_PREFLIGHT_BYTES = 19_200_000_000
TARGET_HARD_LIMIT_BYTES = 20_000_000_000
EXPECTED_LIMITS = {"max_sub_edits": 8, "max_effects_per_sub_edit": 8, "max_total_effects": 32, "max_conflicts": 16}
EXPECTED_PROBE_DESTINATIONS = {
    "common": "crates/litchi-odg/examples/odg_joined_qualification_common.rs",
    "allocator": "crates/litchi-odg/examples/odg_joined_qualification_allocator.rs",
    "baseline": "crates/litchi-odg/examples/litchi-perf-odg-joined-qualification.rs",
    "candidate": "crates/litchi-odg/examples/litchi-perf-odg-joined-qualification.rs",
}


def _exact_keys(value: Mapping[str, Any], expected: set[str], label: str) -> None:
    if set(value) != expected:
        raise RunFailure(f"{label} fields differ from the qualification contract")


def _identity_file(value: Any, label: str, base: Path) -> IdentityFile:
    item = _object(value, label)
    _exact_keys(item, {"path", "bytes", "sha256"}, label)
    return IdentityFile(
        path=_path(item.get("path"), f"{label}.path", base),
        bytes=_positive_int(item.get("bytes"), f"{label}.bytes"),
        sha256=_digest(item.get("sha256"), f"{label}.sha256"),
    )


def _tree_identity(value: Any, label: str) -> dict[str, Any]:
    item = _object(value, label)
    _exact_keys(item, {"entries", "bytes", "sha256"}, label)
    entries = _positive_int(item.get("entries"), f"{label}.entries")
    bytes_value = _positive_int(item.get("bytes"), f"{label}.bytes")
    return {"entries": entries, "bytes": bytes_value, "sha256": _digest(item.get("sha256"), f"{label}.sha256")}


def _digest_binding(value: Any, label: str) -> dict[str, Any]:
    item = _object(value, label)
    _exact_keys(item, {"bytes", "sha256"}, label)
    return {"bytes": _positive_int(item.get("bytes"), f"{label}.bytes"),
            "sha256": _digest(item.get("sha256"), f"{label}.sha256")}


def _reviewed_file(path: Path, relative: Path, expected: Mapping[str, Any], label: str) -> dict[str, Any]:
    approved = (REVIEWED_REPO_ROOT / relative).absolute()
    if path.absolute() != approved:
        raise RunFailure(f"{label} must use reviewed repository path {relative}")
    if path.is_symlink() or not path.is_file():
        raise RunFailure(f"{label} is not a regular reviewed file")
    try:
        status = _run_git(REVIEWED_REPO_ROOT, "ls-files", "-v", "--", str(relative))
    except RunFailure as error:
        raise RunFailure(f"{label} is not tracked by the reviewed repository") from error
    if status != f"H {relative}":
        raise RunFailure(f"{label} is not an ordinary tracked worktree file")
    actual_sha, actual_bytes = _sha256_file(path)
    if (actual_bytes, actual_sha) != (expected["bytes"], expected["sha256"]):
        raise RunFailure(f"{label} differs from the fixed reviewed bytes")
    return {"path": str(path), "bytes": actual_bytes, "sha256": actual_sha}


@dataclass(frozen=True)
class Fixture:
    path: Path
    bytes: int
    sha256: str
    provenance: str


@dataclass(frozen=True)
class IdentityFile:
    path: Path
    bytes: int
    sha256: str


@dataclass(frozen=True)
class Endpoint:
    name: str
    root: Path
    requested_ref: str
    expected_tree: str
    binary: Path
    target_dir: Path
    profile: str
    lock: IdentityFile
    toolchain: IdentityFile
    cargo_config: IdentityFile
    cargo_manifest: IdentityFile
    source_manifest: dict[str, Any]
    source_closure: dict[str, Any]
    features: tuple[str, ...]
    allocator_binary: Path
    allocator_target_dir: Path
    binary_identity: dict[str, Any]
    allocator_binary_identity: dict[str, Any]


@dataclass(frozen=True)
class BuildPaths:
    prepare_receipt: Path
    receipt: Path
    receipt_identity: IdentityFile
    final_candidate_ref: str
    cpu: int


@dataclass(frozen=True)
class ProbeSource:
    name: str
    path: Path
    bytes: int
    sha256: str
    destination: str


@dataclass(frozen=True)
class Manifest:
    path: Path
    sha256: str
    fixture: Fixture
    baseline: Endpoint
    candidate: Endpoint
    probe_sources: tuple[ProbeSource, ...]
    protocol: Protocol
    tmpdir: Path
    limits: dict[str, int]
    build: BuildPaths


@dataclass(frozen=True)
class ResolvedEndpoint:
    endpoint: Endpoint
    resolved_ref: str
    tree: str
    binary: dict[str, Any]
    allocator_binary: dict[str, Any]
    lock: dict[str, Any]
    toolchain: dict[str, Any]
    cargo_config: dict[str, Any]
    cargo_manifest: dict[str, Any]
    source_closure: dict[str, Any]
    environment: dict[str, Any]


def _sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink():
        raise RunFailure(f"refusing to hash symlinked file: {path}")
    digest = hashlib.sha256()
    total = 0
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                total += len(block)
    except OSError as error:
        raise RunFailure(f"cannot hash {path}: {error}") from error
    return digest.hexdigest(), total


def _no_follow_path(path: Path, label: str) -> Path:
    path = path.absolute()
    current = Path(path.anchor)
    for component in path.parts[1:]:
        current /= component
        try:
            current.lstat()
        except FileNotFoundError:
            raise RunFailure(f"{label} disappeared: {current}")
        except OSError as error:
            raise RunFailure(f"cannot inspect {label} {current}: {error}") from error
        if current.is_symlink():
            raise RunFailure(f"{label} contains a symlink component: {current}")
        if current != path and not current.is_dir():
            raise RunFailure(f"{label} has a non-directory component: {current}")
    return path


def _directory_bytes(path: Path) -> int:
    try:
        path = _no_follow_path(path, "target directory")
        if path.is_symlink() or not path.exists():
            raise RunFailure(f"target directory is missing or symlinked: {path}")
        if not path.is_dir():
            raise RunFailure(f"target path is not a directory: {path}")
    except OSError as error:
        raise RunFailure(f"cannot inspect target directory {path}: {error}") from error
    total = 0
    pending = [path]
    while pending:
        current = pending.pop()
        try:
            entries = list(os.scandir(current))
        except OSError as error:
            raise RunFailure(f"cannot inspect target directory {current}: {error}") from error
        for entry in entries:
            try:
                if entry.is_symlink():
                    raise RunFailure(f"target contains a symlink: {entry.path}")
                if entry.is_dir(follow_symlinks=False):
                    pending.append(Path(entry.path))
                elif entry.is_file(follow_symlinks=False):
                    total += entry.stat(follow_symlinks=False).st_size
                else:
                    raise RunFailure(f"target contains a non-regular entry: {entry.path}")
            except FileNotFoundError as error:
                raise RunFailure(f"target entry disappeared during inspection: {entry.path}") from error
            except OSError as error:
                raise RunFailure(f"cannot inspect target entry {entry.path}: {error}") from error
    return total


def _target_guard(endpoints: Mapping[str, ResolvedEndpoint], phase: str) -> dict[str, Any]:
    result: dict[str, Any] = {"phase": phase, "checked_unix_ns": time.time_ns(), "targets": {}}
    for name, resolved in endpoints.items():
        seen: set[Path] = set()
        for kind, target in (
            ("normal", resolved.endpoint.target_dir),
            ("allocator", resolved.endpoint.allocator_target_dir),
        ):
            if target in seen:
                continue
            seen.add(target)
            observed = _directory_bytes(target)
            if phase.startswith("before-") and observed > TARGET_PREFLIGHT_BYTES:
                raise RunFailure(f"{phase}: {name}/{kind} target starts above preflight limit: {observed}")
            if observed > TARGET_HARD_LIMIT_BYTES:
                raise RunFailure(f"{phase}: {name}/{kind} target exceeds hard limit: {observed}")
            result["targets"][f"{name}/{kind}"] = {
                "path": str(target),
                "bytes": observed,
                "preflight_limit": TARGET_PREFLIGHT_BYTES,
                "hard_limit": TARGET_HARD_LIMIT_BYTES,
            }
    return result


def _resource_guard(manifest: Manifest, endpoints: Mapping[str, ResolvedEndpoint], phase: str) -> dict[str, Any]:
    result = _target_guard(endpoints, phase)
    observed = _directory_bytes(manifest.tmpdir)
    if phase.startswith("before-") and observed > TARGET_PREFLIGHT_BYTES:
        raise RunFailure(f"{phase}: temporary directory starts above preflight limit: {observed}")
    if observed > TARGET_HARD_LIMIT_BYTES:
        raise RunFailure(f"{phase}: temporary directory exceeds hard limit: {observed}")
    result["temporary"] = {
        "path": str(manifest.tmpdir),
        "bytes": observed,
        "preflight_limit": TARGET_PREFLIGHT_BYTES,
        "hard_limit": TARGET_HARD_LIMIT_BYTES,
    }
    return result


def _manifest_identity(path: Path) -> str:
    return _sha256_file(path)[0]


def _load_manifest_file(path: Path) -> dict[str, Any]:
    try:
        if path.stat().st_size > MAX_MANIFEST_BYTES:
            raise RunFailure(f"manifest exceeds {MAX_MANIFEST_BYTES} bytes")
    except OSError as error:
        raise RunFailure(f"cannot stat manifest {path}: {error}") from error
    value = _load_json(path)
    if "__sha256" in value:
        raise RunFailure("manifest uses the reserved __sha256 field")
    value["__sha256"] = _manifest_identity(path)
    return value


def load_manifest(path: Path) -> Manifest:
    path = path.resolve()
    value = _load_manifest_file(path)
    _exact_keys(value, {
        "schema", "fixture", "baseline", "candidate", "probe_sources", "protocol",
        "runtime", "composition_limits", "build", "__sha256",
    }, "manifest")
    if value.get("schema") != INPUT_SCHEMA:
        raise RunFailure(f"manifest.schema must be {INPUT_SCHEMA!r}")
    base = path.parent
    fixture_value = _object(value.get("fixture"), "manifest.fixture")
    _exact_keys(fixture_value, {"path", "bytes", "sha256", "provenance"}, "manifest.fixture")
    fixture = Fixture(
        path=_path(fixture_value.get("path"), "manifest.fixture.path", base),
        bytes=_positive_int(fixture_value.get("bytes"), "manifest.fixture.bytes"),
        sha256=_digest(fixture_value.get("sha256"), "manifest.fixture.sha256"),
        provenance=_string(fixture_value.get("provenance"), "manifest.fixture.provenance"),
    )
    if fixture.bytes > MAX_FIXTURE_BYTES:
        raise RunFailure(f"manifest.fixture.bytes exceeds {MAX_FIXTURE_BYTES}")

    def endpoint(name: str, profile: str) -> Endpoint:
        item = _object(value.get(name), f"manifest.{name}")
        _exact_keys(item, {
            "root", "ref", "tree", "profile", "binary", "target_dir", "lock", "toolchain",
            "cargo_config", "cargo_manifest", "source_manifest", "source_closure", "allocator_binary", "allocator_target_dir",
            "features", "binary_identity", "allocator_binary_identity",
        }, f"manifest.{name}")
        requested_ref = _string(item.get("ref"), f"manifest.{name}.ref")
        expected_tree = _string(item.get("tree"), f"manifest.{name}.tree")
        if REVISION.fullmatch(requested_ref) is None or REVISION.fullmatch(expected_tree) is None:
            raise RunFailure(f"manifest.{name}.ref and tree must be immutable 40-character revisions")
        actual = _string(item.get("profile"), f"manifest.{name}.profile")
        if actual != profile:
            raise RunFailure(f"manifest.{name}.profile must be {profile!r}")
        features_value = item.get("features")
        if (not isinstance(features_value, list)
                or any(not isinstance(feature, str) or not feature for feature in features_value)
                or features_value != sorted(set(features_value))):
            raise RunFailure(f"manifest.{name}.features must be a sorted unique string list")
        return Endpoint(
            name=name,
            root=_path(item.get("root"), f"manifest.{name}.root", base),
            requested_ref=requested_ref,
            expected_tree=expected_tree,
            binary=_path(item.get("binary"), f"manifest.{name}.binary", base),
            target_dir=_path(item.get("target_dir"), f"manifest.{name}.target_dir", base),
            profile=actual,
            lock=_identity_file(item.get("lock"), f"manifest.{name}.lock", base),
            toolchain=_identity_file(item.get("toolchain"), f"manifest.{name}.toolchain", base),
            cargo_config=_identity_file(item.get("cargo_config"), f"manifest.{name}.cargo_config", base),
            cargo_manifest=_identity_file(item.get("cargo_manifest"), f"manifest.{name}.cargo_manifest", base),
            source_manifest=_tree_identity(item.get("source_manifest"), f"manifest.{name}.source_manifest"),
            source_closure=_tree_identity(item.get("source_closure"), f"manifest.{name}.source_closure"),
            features=tuple(features_value),
            allocator_binary=_path(item.get("allocator_binary"), f"manifest.{name}.allocator_binary", base),
            allocator_target_dir=_path(item.get("allocator_target_dir"), f"manifest.{name}.allocator_target_dir", base),
            binary_identity=_digest_binding(item.get("binary_identity"), f"manifest.{name}.binary_identity"),
            allocator_binary_identity=_digest_binding(
                item.get("allocator_binary_identity"), f"manifest.{name}.allocator_binary_identity"
            ),
        )

    baseline = endpoint("baseline", BASELINE_PROFILE)
    candidate = endpoint("candidate", CANDIDATE_PROFILE)
    if baseline.root == candidate.root:
        raise RunFailure("baseline and candidate roots must be distinct")
    source_value = _object(value.get("probe_sources"), "manifest.probe_sources")
    _exact_keys(source_value, {"common", "allocator", "baseline", "candidate"}, "manifest.probe_sources")
    probe_sources: list[ProbeSource] = []
    for name in ("common", "allocator", "baseline", "candidate"):
        item = _object(source_value.get(name), f"manifest.probe_sources.{name}")
        _exact_keys(item, {"path", "bytes", "sha256", "destination"}, f"manifest.probe_sources.{name}")
        destination = _string(item.get("destination"), f"manifest.probe_sources.{name}.destination")
        if destination != EXPECTED_PROBE_DESTINATIONS[name]:
            raise RunFailure(
                f"manifest.probe_sources.{name}.destination must be {EXPECTED_PROBE_DESTINATIONS[name]!r}"
            )
        probe_sources.append(ProbeSource(
            name=name,
            path=_path(item.get("path"), f"manifest.probe_sources.{name}.path", base),
            bytes=_positive_int(item.get("bytes"), f"manifest.probe_sources.{name}.bytes"),
            sha256=_digest(item.get("sha256"), f"manifest.probe_sources.{name}.sha256"),
            destination=destination,
        ))

    protocol_value = _object(value.get("protocol"), "manifest.protocol")
    _exact_keys(protocol_value, {
        "cpu", "blocks", "warmups", "target_ms", "minimum_ms", "max_load",
        "max_iterations", "timeout_s",
    }, "manifest.protocol")
    cpu = _nonnegative_int(protocol_value.get("cpu"), "manifest.protocol.cpu")
    blocks = _positive_int(protocol_value.get("blocks"), "manifest.protocol.blocks")
    warmups = _nonnegative_int(protocol_value.get("warmups"), "manifest.protocol.warmups")
    if blocks != REQUIRED_BLOCKS or warmups != REQUIRED_WARMUPS:
        raise RunFailure(
            f"ODG protocol requires exactly {REQUIRED_BLOCKS} ABBA blocks and "
            f"{REQUIRED_WARMUPS} warmups"
        )
    target_ms = _positive_float(protocol_value.get("target_ms"), "manifest.protocol.target_ms")
    minimum_ms = _positive_float(protocol_value.get("minimum_ms"), "manifest.protocol.minimum_ms")
    if target_ms < 500 or minimum_ms < 150 or target_ms < minimum_ms:
        raise RunFailure("ODG protocol requires target_ms >= 500, minimum_ms >= 150, and target >= minimum")
    protocol = Protocol(
        cpu=cpu,
        blocks=blocks,
        warmups=warmups,
        target_ns=round(target_ms * 1_000_000),
        minimum_ns=round(minimum_ms * 1_000_000),
        max_load=_positive_float(protocol_value.get("max_load"), "manifest.protocol.max_load"),
        max_iterations=_positive_int(protocol_value.get("max_iterations"), "manifest.protocol.max_iterations"),
        timeout_s=_positive_float(protocol_value.get("timeout_s"), "manifest.protocol.timeout_s"),
    )
    runtime = _object(value.get("runtime"), "manifest.runtime")
    _exact_keys(runtime, {"tmpdir"}, "manifest.runtime")
    tmpdir = _path(runtime.get("tmpdir"), "manifest.runtime.tmpdir", base)
    limit_value = _object(value.get("composition_limits"), "manifest.composition_limits")
    _exact_keys(limit_value, set(EXPECTED_LIMITS), "manifest.composition_limits")
    limits = {key: _positive_int(limit_value.get(key), f"manifest.composition_limits.{key}") for key in EXPECTED_LIMITS}
    if limits != EXPECTED_LIMITS:
        raise RunFailure(f"composition_limits must equal {EXPECTED_LIMITS!r}")
    build_value = _object(value.get("build"), "manifest.build")
    _exact_keys(build_value, {"prepare_receipt", "receipt", "receipt_identity", "final_candidate_ref", "cpu"}, "manifest.build")
    final_candidate_ref = _string(build_value.get("final_candidate_ref"), "manifest.build.final_candidate_ref")
    if REVISION.fullmatch(final_candidate_ref) is None or final_candidate_ref != candidate.requested_ref:
        raise RunFailure("manifest.build.final_candidate_ref must equal the immutable candidate ref")
    receipt = _path(build_value.get("receipt"), "manifest.build.receipt", base)
    receipt_identity = _identity_file(
        build_value.get("receipt_identity"), "manifest.build.receipt_identity", base
    )
    if receipt_identity.path != receipt:
        raise RunFailure("manifest.build.receipt_identity.path must equal manifest.build.receipt")
    build = BuildPaths(
        prepare_receipt=_path(build_value.get("prepare_receipt"), "manifest.build.prepare_receipt", base),
        receipt=receipt,
        receipt_identity=receipt_identity,
        final_candidate_ref=final_candidate_ref,
        cpu=_nonnegative_int(build_value.get("cpu"), "manifest.build.cpu"),
    )
    if build.prepare_receipt == build.receipt:
        raise RunFailure("manifest build receipts must be distinct")
    return Manifest(path, value["__sha256"], fixture, baseline, candidate, tuple(probe_sources), protocol, tmpdir, limits, build)


def validate_fixture(fixture: Fixture) -> dict[str, Any]:
    if fixture.path.absolute() != (REVIEWED_REPO_ROOT / REVIEWED_FIXTURE_RELATIVE).absolute():
        raise RunFailure("fixture path is not the fixed reviewed repository fixture")
    if fixture.provenance != REVIEWED_FIXTURE_PROVENANCE:
        raise RunFailure("fixture provenance is not the fixed reviewed provenance")
    if not fixture.path.is_file() or fixture.path.is_symlink():
        raise RunFailure(f"fixture is not a regular reviewed file: {fixture.path}")
    actual_sha, actual_bytes = _sha256_file(fixture.path)
    if (actual_bytes, actual_sha) != (fixture.bytes, fixture.sha256) or {
        "bytes": actual_bytes, "sha256": actual_sha
    } != REVIEWED_FIXTURE_IDENTITY:
        raise RunFailure("fixture identity differs from its manifest")
    content = fixture.path.read_bytes()
    if b"<office:document-content" not in content or b"<draw:page" not in content:
        raise RunFailure("fixture is not the required ODG content XML")
    return {"path": str(fixture.path), "bytes": actual_bytes, "sha256": actual_sha, "provenance": fixture.provenance}


def validate_probe_sources(sources: tuple[ProbeSource, ...]) -> dict[str, Any]:
    if len({source.path for source in sources}) != len(sources):
        raise RunFailure("probe source overlays must be separate files")
    records = {}
    for source in sources:
        relative = REVIEWED_PROBE_RELATIVES.get(source.name)
        if relative is None or source.path.absolute() != (REVIEWED_REPO_ROOT / relative).absolute():
            raise RunFailure(f"probe source {source.name} is not the fixed reviewed repository adapter")
        if source.path.is_symlink() or not source.path.is_file():
            raise RunFailure(f"probe source {source.name} is not a regular reviewed file")
        status = _run_git(REVIEWED_REPO_ROOT, "ls-files", "-v", "--", str(relative))
        if status != f"H {relative}":
            raise RunFailure(f"probe source {source.name} is not an ordinary tracked worktree file")
        actual_sha, actual_bytes = _sha256_file(source.path)
        expected = REVIEWED_PROBE_IDENTITIES[source.name]
        if ((actual_bytes, actual_sha) != (source.bytes, source.sha256)
                or (actual_bytes, actual_sha) != (expected["bytes"], expected["sha256"])):
            raise RunFailure(f"probe source {source.name} identity differs from its manifest")
        records[source.name] = {"path": str(source.path), "bytes": actual_bytes, "sha256": actual_sha,
                                "destination": source.destination}
    return records


def _resolve_endpoint(
    endpoint: Endpoint,
    roots: tuple[Path, ...],
    *,
    environment_overrides: Mapping[str, str] | None = None,
) -> ResolvedEndpoint:
    if not endpoint.root.is_dir():
        raise RunFailure(f"{endpoint.name}.root is not a directory: {endpoint.root}")
    try:
        from odg_joined_qualification_builder import (
            _check_worktree_shape,
            _environment_identity,
            _source_closure,
            _tree_manifest,
        )
    except ImportError:  # pragma: no cover
        from tools.odg_joined_qualification_builder import (  # type: ignore[no-redef]
            _check_worktree_shape,
            _environment_identity,
            _source_closure,
            _tree_manifest,
        )
    _check_worktree_shape(endpoint.root, endpoint.name)
    status = _run_git(endpoint.root, "status", "--porcelain", "--untracked-files=all")
    if status:
        raise RunFailure(f"{endpoint.name}.root is dirty")
    resolved = _run_git(endpoint.root, "rev-parse", "--verify", f"{endpoint.requested_ref}^{{commit}}")
    head = _run_git(endpoint.root, "rev-parse", "--verify", "HEAD")
    if (REVISION.fullmatch(resolved) is None or REVISION.fullmatch(head) is None
            or resolved != head or resolved != endpoint.requested_ref):
        raise RunFailure(f"{endpoint.name}: requested ref is not the clean worktree HEAD")
    tree = _run_git(endpoint.root, "rev-parse", "--verify", f"{resolved}^{{tree}}")
    if REVISION.fullmatch(tree) is None or tree != endpoint.expected_tree:
        raise RunFailure(f"{endpoint.name}: resolved tree differs from the manifest")
    expected_binary = endpoint.target_dir / "release" / "examples" / "litchi-perf-odg-joined-qualification"
    expected_allocator_binary = (
        endpoint.allocator_target_dir / "release" / "examples" / "litchi-perf-odg-joined-qualification"
    )
    if endpoint.binary != expected_binary:
        raise RunFailure(f"{endpoint.name}.binary must be the release example artifact")
    if endpoint.allocator_binary != expected_allocator_binary:
        raise RunFailure(f"{endpoint.name}.allocator_binary must be the release example artifact")
    if not endpoint.binary.is_file() or endpoint.binary.is_symlink() or not endpoint.binary.stat().st_mode & 0o111:
        raise RunFailure(f"{endpoint.name}.binary is not executable")
    if endpoint.binary.name != "litchi-perf-odg-joined-qualification":
        raise RunFailure(
            f"{endpoint.name}.binary must be the paired litchi-perf-odg-joined-qualification probe"
        )
    if not endpoint.target_dir.is_dir():
        raise RunFailure(f"{endpoint.name}.target_dir is not a directory")
    if not endpoint.allocator_binary.is_file() or endpoint.allocator_binary.is_symlink() or not endpoint.allocator_binary.stat().st_mode & 0o111:
        raise RunFailure(f"{endpoint.name}.allocator_binary is not executable")
    if endpoint.allocator_binary.name != "litchi-perf-odg-joined-qualification":
        raise RunFailure(f"{endpoint.name}.allocator_binary has an unexpected basename")
    if not endpoint.allocator_target_dir.is_dir():
        raise RunFailure(f"{endpoint.name}.allocator_target_dir is not a directory")
    _validate_external_target(endpoint.target_dir, roots, f"{endpoint.name}.target_dir")
    _validate_external_target(endpoint.allocator_target_dir, roots, f"{endpoint.name}.allocator_target_dir")
    try:
        endpoint.binary.relative_to(endpoint.target_dir)
    except ValueError as error:
        raise RunFailure(f"{endpoint.name}.binary must be below target_dir") from error
    try:
        endpoint.allocator_binary.relative_to(endpoint.allocator_target_dir)
    except ValueError as error:
        raise RunFailure(f"{endpoint.name}.allocator_binary must be below allocator_target_dir") from error
    def identity_file(value: IdentityFile, label: str) -> dict[str, Any]:
        try:
            value.path.relative_to(endpoint.root)
        except ValueError as error:
            raise RunFailure(f"{label} must be below {endpoint.name}.root") from error
        actual_sha, actual_bytes = _sha256_file(value.path)
        if (actual_bytes, actual_sha) != (value.bytes, value.sha256):
            raise RunFailure(f"{label} identity differs from its manifest")
        return {"path": str(value.path), "bytes": actual_bytes, "sha256": actual_sha}
    lock = identity_file(endpoint.lock, f"{endpoint.name}.lock")
    toolchain = identity_file(endpoint.toolchain, f"{endpoint.name}.toolchain")
    cargo_config = identity_file(endpoint.cargo_config, f"{endpoint.name}.cargo_config")
    cargo_manifest = identity_file(endpoint.cargo_manifest, f"{endpoint.name}.cargo_manifest")
    actual_binary = binary_identity(endpoint.binary)
    actual_allocator = binary_identity(endpoint.allocator_binary)
    if (actual_binary["bytes"], actual_binary["sha256"]) != (
        endpoint.binary_identity["bytes"], endpoint.binary_identity["sha256"]
    ):
        raise RunFailure(f"{endpoint.name}.binary differs from its external binding")
    if (actual_allocator["bytes"], actual_allocator["sha256"]) != (
        endpoint.allocator_binary_identity["bytes"], endpoint.allocator_binary_identity["sha256"]
    ):
        raise RunFailure(f"{endpoint.name}.allocator_binary differs from its external binding")
    if _tree_manifest(endpoint.root, resolved) != endpoint.source_manifest:
        raise RunFailure(f"{endpoint.name}: Git tree metadata differs from its manifest")
    source_closure = _source_closure(endpoint.root, resolved, endpoint.source_closure)
    environment = _environment_identity(
        endpoint.root,
        ref=resolved,
        tree=tree,
        source_closure=source_closure["binding"],
        overrides=environment_overrides,
    )
    return ResolvedEndpoint(endpoint, resolved, tree, actual_binary, actual_allocator,
                            lock, toolchain, cargo_config, cargo_manifest,
                            source_closure, environment)


def _under(left: Path, right: Path) -> bool:
    return left == right or right in left.parents


def _validate_layout(manifest: Manifest, out_dir: Path) -> None:
    roots = (manifest.baseline.root, manifest.candidate.root)
    try:
        from odg_joined_qualification_builder import _raw_path, _safe_path
    except ImportError:  # pragma: no cover
        from tools.odg_joined_qualification_builder import _raw_path, _safe_path  # type: ignore[no-redef]
    _safe_path(out_dir.absolute(), "artifact directory")
    raw = _load_json(manifest.path)
    runtime_value = _object(raw.get("runtime"), "manifest.runtime")
    _safe_path(_raw_path(runtime_value.get("tmpdir"), "manifest.runtime.tmpdir", manifest.path.parent),
               "temporary directory")
    _validate_external_target(out_dir, roots, "artifact directory")
    _validate_external_target(manifest.tmpdir, roots, "temporary directory")
    _validate_external_target(manifest.build.prepare_receipt, roots, "prepare receipt")
    _validate_external_target(manifest.build.receipt, roots, "build receipt")
    _validate_external_target(manifest.build.receipt_identity.path, roots, "build receipt identity")
    if (UNTRACKED_UPSTREAM_COMPONENT in manifest.build.receipt.parts
            or UNTRACKED_UPSTREAM_COMPONENT in manifest.build.prepare_receipt.parts):
        raise RunFailure(
            f"build receipts must not reference {UNTRACKED_UPSTREAM_COMPONENT}"
        )
    if _under(out_dir, manifest.tmpdir) or _under(manifest.tmpdir, out_dir):
        raise RunFailure("artifact and temporary directories must be disjoint")
    pairs = (
        (out_dir, manifest.baseline.target_dir),
        (out_dir, manifest.candidate.target_dir),
        (out_dir, manifest.baseline.allocator_target_dir),
        (out_dir, manifest.candidate.allocator_target_dir),
        (manifest.baseline.target_dir, manifest.candidate.target_dir),
        (manifest.baseline.allocator_target_dir, manifest.candidate.allocator_target_dir),
        (manifest.baseline.target_dir, manifest.baseline.allocator_target_dir),
        (manifest.candidate.target_dir, manifest.candidate.allocator_target_dir),
        (manifest.tmpdir, manifest.baseline.target_dir),
        (manifest.tmpdir, manifest.candidate.target_dir),
        (manifest.tmpdir, manifest.baseline.allocator_target_dir),
        (manifest.tmpdir, manifest.candidate.allocator_target_dir),
    )
    if any(_under(left, right) or _under(right, left) for left, right in pairs):
        raise RunFailure("target and temporary directories must be pairwise disjoint")


def runtime_identity(manifest: Manifest, out_dir: Path, *, require_build_receipt: bool = True) -> dict[str, Any]:
    _validate_layout(manifest, out_dir)
    if require_build_receipt:
        try:
            from odg_joined_qualification_builder import validate_build_receipt
        except ImportError:  # pragma: no cover - direct module loading in tests.
            from tools.odg_joined_qualification_builder import validate_build_receipt  # type: ignore[no-redef]
        validate_build_receipt(manifest.path)
    fixture = validate_fixture(manifest.fixture)
    probes = validate_probe_sources(manifest.probe_sources)
    roots = (manifest.baseline.root, manifest.candidate.root)
    environment_overrides = {"TMPDIR": str(manifest.tmpdir)}
    baseline = _resolve_endpoint(manifest.baseline, roots, environment_overrides=environment_overrides)
    candidate = _resolve_endpoint(manifest.candidate, roots, environment_overrides=environment_overrides)
    if baseline.resolved_ref == candidate.resolved_ref or baseline.tree == candidate.tree:
        raise RunFailure("baseline and candidate source identities must differ")
    if _directory_bytes(manifest.tmpdir) > TARGET_HARD_LIMIT_BYTES:
        raise RunFailure("temporary directory exceeds target hard limit")
    return {
        "fixture": fixture,
        "probe_sources": probes,
        "baseline": {"root": str(baseline.endpoint.root), "requested_ref": baseline.endpoint.requested_ref,
                     "resolved_ref": baseline.resolved_ref, "tree": baseline.tree,
                     "profile": baseline.endpoint.profile, "features": list(baseline.endpoint.features),
                     "lock": baseline.lock, "toolchain": baseline.toolchain,
                     "cargo_config": baseline.cargo_config, "cargo_manifest": baseline.cargo_manifest,
                     "source_closure": baseline.source_closure["binding"],
                     "environment": baseline.environment,
                     "binary": baseline.binary, "allocator_binary": baseline.allocator_binary,
                     "target_dir": str(baseline.endpoint.target_dir),
                     "allocator_target_dir": str(baseline.endpoint.allocator_target_dir)},
        "candidate": {"root": str(candidate.endpoint.root), "requested_ref": candidate.endpoint.requested_ref,
                      "resolved_ref": candidate.resolved_ref, "tree": candidate.tree,
                      "profile": candidate.endpoint.profile, "features": list(candidate.endpoint.features),
                      "lock": candidate.lock, "toolchain": candidate.toolchain,
                      "cargo_config": candidate.cargo_config, "cargo_manifest": candidate.cargo_manifest,
                      "source_closure": candidate.source_closure["binding"],
                      "environment": candidate.environment,
                      "binary": candidate.binary, "allocator_binary": candidate.allocator_binary,
                      "target_dir": str(candidate.endpoint.target_dir),
                      "allocator_target_dir": str(candidate.endpoint.allocator_target_dir)},
        "endpoints": {"baseline": baseline, "candidate": candidate},
    }


def _assert_unchanged(manifest: Manifest, endpoints: Mapping[str, ResolvedEndpoint], phase: str) -> dict[str, Any]:
    if _manifest_identity(manifest.path) != manifest.sha256:
        raise RunFailure(f"{phase}: manifest changed")
    result: dict[str, Any] = {"phase": phase, "fixture": validate_fixture(manifest.fixture),
                              "probe_sources": validate_probe_sources(manifest.probe_sources)}
    full_source_check = phase in {"after-correctness", "after-allocation", "after-calibration", "after-abba"}
    for name, resolved in endpoints.items():
        root = resolved.endpoint.root
        try:
            from odg_joined_qualification_builder import (
                _check_worktree_shape,
                _environment_identity,
                _validate_source_closure,
            )
        except ImportError:  # pragma: no cover
            from tools.odg_joined_qualification_builder import (  # type: ignore[no-redef]
                _check_worktree_shape,
                _environment_identity,
                _validate_source_closure,
            )
        _check_worktree_shape(root, name)
        status = _run_git(root, "status", "--porcelain", "--untracked-files=all")
        head = _run_git(root, "rev-parse", "--verify", "HEAD")
        requested = _run_git(root, "rev-parse", "--verify", f"{resolved.endpoint.requested_ref}^{{commit}}")
        tree = _run_git(root, "rev-parse", "--verify", f"{resolved.resolved_ref}^{{tree}}")
        if status or head != resolved.resolved_ref or requested != resolved.resolved_ref or tree != resolved.tree:
            raise RunFailure(f"{phase}: {name} source ref or cleanliness changed")
        _validate_source_closure(
            root,
            resolved.source_closure,
            resolved.endpoint.source_closure,
            force_full=full_source_check,
        )
        current_environment = _environment_identity(
            root,
            ref=resolved.resolved_ref,
            tree=resolved.tree,
            source_closure=resolved.source_closure["binding"],
            overrides={"TMPDIR": str(manifest.tmpdir)},
        )
        if current_environment != resolved.environment:
            raise RunFailure(f"{phase}: {name} effective build environment changed")
        actual = binary_identity(resolved.endpoint.binary)
        if actual != resolved.binary:
            raise RunFailure(f"{phase}: {name} binary changed")
        actual_allocator = binary_identity(resolved.endpoint.allocator_binary)
        if actual_allocator != resolved.allocator_binary:
            raise RunFailure(f"{phase}: {name} allocator binary changed")
        for label, value, expected in (
            ("lock", resolved.endpoint.lock, resolved.lock),
            ("toolchain", resolved.endpoint.toolchain, resolved.toolchain),
            ("cargo_config", resolved.endpoint.cargo_config, resolved.cargo_config),
            ("cargo_manifest", resolved.endpoint.cargo_manifest, resolved.cargo_manifest),
        ):
            actual_sha, actual_bytes = _sha256_file(value.path)
            if (actual_bytes, actual_sha) != (value.bytes, value.sha256) or {
                "path": str(value.path), "bytes": actual_bytes, "sha256": actual_sha
            } != expected:
                raise RunFailure(f"{phase}: {name} {label} changed")
        result[name] = {"head": head, "tree": tree, "requested_ref": requested,
                        "features": list(resolved.endpoint.features), "binary": actual,
                        "allocator_binary": actual_allocator,
                        "lock": resolved.lock, "toolchain": resolved.toolchain,
                        "cargo_config": resolved.cargo_config, "cargo_manifest": resolved.cargo_manifest,
                        "source_closure": resolved.source_closure["binding"],
                        "environment": resolved.environment}
    result["checked_unix_ns"] = time.time_ns()
    return result


def _probe_command(endpoint: ResolvedEndpoint, manifest: Manifest, mode: str,
                   *, binary: Path | None = None, **kwargs: Any) -> list[str]:
    taskset = _require_taskset(manifest.protocol.cpu)
    command = [taskset, "--cpu-list", str(manifest.protocol.cpu),
               str(binary or endpoint.endpoint.binary), "--mode", mode, "--fixture", str(manifest.fixture.path)]
    for key in ("route", "scope", "iterations", "samples", "warmups", "target_ns"):
        if kwargs.get(key) is not None:
            command.extend((f"--{key.replace('_', '-')}", str(kwargs[key])))
    command.extend(("--max-iterations", str(manifest.protocol.max_iterations)))
    return command


def _run_probe(endpoint: ResolvedEndpoint, manifest: Manifest, out_dir: Path, mode: str, label: str,
               *, binary: Path | None = None, target_dir: Path | None = None, **kwargs: Any) -> tuple[dict[str, Any], dict[str, Any]]:
    command = _probe_command(endpoint, manifest, mode, binary=binary, **kwargs)
    environment_identity = endpoint.environment
    environment = environment_identity["variables"]
    try:
        from odg_joined_qualification_builder import _stream_command
    except ImportError:  # pragma: no cover - direct module loading in tests.
        from tools.odg_joined_qualification_builder import _stream_command  # type: ignore[no-redef]
    guard_before = quiet_guard(manifest.protocol, f"before-{label}")
    event = _stream_command(
        command,
        endpoint.endpoint.root,
        out_dir,
        label,
        target_dir or endpoint.endpoint.target_dir,
        manifest.protocol.cpu,
        manifest.protocol.timeout_s,
        environment,
        MAX_CHILD_OUTPUT,
        manifest.protocol.max_load,
    )
    guard_after = quiet_guard(manifest.protocol, f"after-{label}")
    logs = {
        "command": command, "returncode": event["returncode"],
        "endpoint": endpoint.endpoint.name, "profile": endpoint.endpoint.profile,
        "resolved_ref": endpoint.resolved_ref,
        "binary": binary_identity(binary) if binary is not None else endpoint.binary,
        "environment": environment_identity,
        "binding": {"mode": mode, "route": kwargs.get("route"), "scope": kwargs.get("scope")},
        "stdout": event["stdout"], "stderr": event["stderr"],
        "guard": {"before": guard_before, "live": event["monitor"]["samples"], "after": guard_after,
                   "memory_limit_bytes": event["monitor"]["memory_limit_bytes"],
                   "max_output_bytes": event["monitor"]["max_output_bytes"],
                   "max_load": event["monitor"]["max_load"]},
        "target": event["target"],
    }
    stdout_path = Path(event["stdout"]["path"])
    report = _parse_child_json(stdout_path.read_bytes(), label)
    return report, logs


def _validate_log_binding(logs: Mapping[str, Any], endpoint: ResolvedEndpoint, mode: str,
                          route: str | None, scope: str | None, label: str) -> None:
    _exact_keys(logs, {"command", "returncode", "endpoint", "profile", "resolved_ref", "binary",
                       "environment", "binding", "stdout", "stderr", "guard", "target"}, label)
    if (logs.get("returncode") != 0 or logs.get("endpoint") != endpoint.endpoint.name
            or logs.get("profile") != endpoint.endpoint.profile
            or logs.get("resolved_ref") != endpoint.resolved_ref
            or logs.get("binary") != endpoint.binary
            or logs.get("environment") != endpoint.environment):
        raise RunFailure(f"{label}: child log is bound to a different endpoint or binary")
    binding = _object(logs.get("binding"), f"{label}.binding")
    _exact_keys(binding, {"mode", "route", "scope"}, f"{label}.binding")
    if binding != {"mode": mode, "route": route, "scope": scope}:
        raise RunFailure(f"{label}: child log route/phase binding differs from the invocation")
    for stream in ("stdout", "stderr"):
        item = _object(logs.get(stream), f"{label}.{stream}")
        _exact_keys(item, {"path", "bytes", "sha256"}, f"{label}.{stream}")
        path = Path(_string(item.get("path"), f"{label}.{stream}.path"))
        if not path.is_file() or path.is_symlink():
            raise RunFailure(f"{label}.{stream}: log path is not a regular file")
        actual_sha, actual_bytes = _sha256_file(path)
        if item.get("bytes") != actual_bytes or item.get("sha256") != actual_sha:
            raise RunFailure(f"{label}.{stream}: log bytes or hash changed")
        if actual_bytes > MAX_CHILD_OUTPUT:
            raise RunFailure(f"{label}.{stream}: log exceeds bounded child output")
    guard = _object(logs.get("guard"), f"{label}.guard")
    _exact_keys(guard, {"before", "live", "after", "memory_limit_bytes", "max_output_bytes", "max_load"}, f"{label}.guard")
    if guard["memory_limit_bytes"] != 20_000_000_000 or guard["max_output_bytes"] != MAX_CHILD_OUTPUT:
        raise RunFailure(f"{label}.guard resource limits differ from the fixed protocol")
    for phase in ("before", "after"):
        value = _object(guard.get(phase), f"{label}.guard.{phase}")
        _exact_keys(value, {"phase", "load_1m", "max_load", "heavy_processes", "checked_unix_ns"},
                    f"{label}.guard.{phase}")
        if not isinstance(value["heavy_processes"], list):
            raise RunFailure(f"{label}.guard.{phase}: process list is malformed")
        if (type(value["load_1m"]) not in (int, float) or value["load_1m"] < 0
                or value["load_1m"] > guard["max_load"]):
            raise RunFailure(f"{label}.guard.{phase}: load bound failed")
        if value["max_load"] != guard["max_load"]:
            raise RunFailure(f"{label}.guard.{phase}: load ceiling identity differs")
    live = guard["live"]
    if not isinstance(live, list) or not live:
        raise RunFailure(f"{label}.guard.live is incomplete")
    for index, value in enumerate(live):
        sample = _object(value, f"{label}.guard.live[{index}]")
        _exact_keys(sample, {"checked_unix_ns", "output_bytes", "rss_bytes", "load_1m",
                             "outside_processes", "owned_processes"},
                    f"{label}.guard.live[{index}]")
        _positive_int(sample["checked_unix_ns"], f"{label}.guard.live[{index}].checked_unix_ns")
        if type(sample["output_bytes"]) is not int or sample["output_bytes"] < 0 or sample["output_bytes"] > MAX_CHILD_OUTPUT:
            raise RunFailure(f"{label}.guard.live[{index}]: output bound failed")
        if type(sample["rss_bytes"]) is not int or sample["rss_bytes"] < 0 or sample["rss_bytes"] > TARGET_HARD_LIMIT_BYTES:
            raise RunFailure(f"{label}.guard.live[{index}]: memory bound failed")
        if (type(sample["load_1m"]) not in (int, float)
                or sample["load_1m"] < 0 or sample["load_1m"] > guard["max_load"]):
            raise RunFailure(f"{label}.guard.live[{index}]: load bound failed")
        if sample["outside_processes"] != []:
            raise RunFailure(f"{label}.guard.live[{index}]: unrelated process observed")
        if not isinstance(sample["owned_processes"], list):
            raise RunFailure(f"{label}.guard.live[{index}]: owned process identity is malformed")
        for process in sample["owned_processes"]:
            item = _object(process, f"{label}.guard.live[{index}].owned_process")
            _exact_keys(item, {"pid", "start_time", "process_group"},
                        f"{label}.guard.live[{index}].owned_process")
            _positive_int(item["pid"], f"{label}.guard.live[{index}].owned_process.pid")
            _positive_int(item["start_time"], f"{label}.guard.live[{index}].owned_process.start_time")
            _positive_int(item["process_group"], f"{label}.guard.live[{index}].owned_process.process_group")
    if live[-1].get("owned_processes") != []:
        raise RunFailure(f"{label}.guard.live: owned descendants remained after child exit")
    target = _object(logs.get("target"), f"{label}.target")
    _exact_keys(target, {"path", "bytes_before", "bytes_after", "hard_limit_bytes"}, f"{label}.target")
    expected_target = str(endpoint.endpoint.target_dir)
    if target["path"] != expected_target:
        raise RunFailure(f"{label}.target path differs from the endpoint target")
    if (type(target["bytes_before"]) is not int or type(target["bytes_after"]) is not int
            or target["bytes_before"] < 0 or target["bytes_after"] < 0
            or target["bytes_before"] > TARGET_HARD_LIMIT_BYTES
            or target["bytes_after"] > TARGET_HARD_LIMIT_BYTES):
        raise RunFailure(f"{label}.target byte accounting is malformed")
    if target["hard_limit_bytes"] != TARGET_HARD_LIMIT_BYTES:
        raise RunFailure(f"{label}.target hard limit identity differs")
    if target["bytes_before"] > target["hard_limit_bytes"] or target["bytes_after"] > target["hard_limit_bytes"]:
        raise RunFailure(f"{label}.target exceeded its hard bound")


def _digest_object(value: Any, label: str) -> dict[str, Any]:
    item = _object(value, label)
    if set(item) != {"bytes", "sha256"} or _nonnegative_int(item.get("bytes"), f"{label}.bytes") == 0:
        raise RunFailure(f"{label} is not a positive digest object")
    return {"bytes": item["bytes"], "sha256": _digest(item.get("sha256"), f"{label}.sha256")}


def _validate_route(route: str, value: Mapping[str, Any], candidate: bool) -> dict[str, Any]:
    proof_name = "candidate" if candidate else "baseline"
    if set(value) != {"output", "published", "semantic", "checks", proof_name}:
        raise RunFailure(f"{route}: route proof fields are incomplete")
    output = _digest_object(value["output"], f"{route}.output")
    published = _digest_object(value["published"], f"{route}.published")
    checks = _object(value["checks"], f"{route}.checks")
    required_checks = {"source_unchanged", "publication_readback_exact", "route_semantics", "changed_output",
                       "exact_noop", "deterministic_repeat"}
    if candidate:
        required_checks.add("diagnostics_changed")
    if set(checks) != required_checks or any(item is not True for item in checks.values()):
        raise RunFailure(f"{route}: route checks are incomplete or false")
    proof = _object(value[proof_name], f"{route}.{proof_name}")
    if candidate:
        required = {"diagnostics_changed", "diagnostics_change_count", "diagnostics_resource_change_count",
                    "patch_applicable", "patch_noop", "direct_forward_exact", "durable_forward_exact",
                    "inverse_exact", "durable_inverse_exact", "content_matching_forward_exact"}
        if set(proof) != required:
            raise RunFailure(f"{route}: candidate proof fields are incomplete")
        booleans = required - {"diagnostics_change_count", "diagnostics_resource_change_count", "patch_noop"}
        if any(proof[key] is not True for key in booleans) or not isinstance(proof["patch_noop"], bool):
            raise RunFailure(f"{route}: candidate commit proof is false or malformed")
        _nonnegative_int(proof["diagnostics_change_count"], f"{route}.diagnostics_change_count")
        _nonnegative_int(proof["diagnostics_resource_change_count"], f"{route}.diagnostics_resource_change_count")
        if proof["patch_noop"] is not (route == "empty_join"):
            raise RunFailure(f"{route}: patch noop state is incorrect")
    elif proof != {"apply_joined": True, "commit_triads": False, "patch_replay": False}:
        raise RunFailure(f"{route}: baseline proof is incomplete or overstated")
    semantic = _object(value["semantic"], f"{route}.semantic")
    pages = semantic.get("pages")
    if not isinstance(pages, list):
        raise RunFailure(f"{route}: semantic.pages is not a list")
    def shape(page_index: int, shape_index: int) -> Mapping[str, Any]:
        if page_index >= len(pages):
            raise RunFailure(f"{route}: semantic page {page_index} is missing")
        page = _object(pages[page_index], f"{route}.semantic.pages[{page_index}]")
        shapes = page.get("shapes")
        if not isinstance(shapes, list) or shape_index >= len(shapes):
            raise RunFailure(f"{route}: semantic shape {page_index}/{shape_index} is missing")
        return _object(shapes[shape_index], f"{route}.semantic.pages[{page_index}].shapes[{shape_index}]")
    if route == "package_replace_one":
        if len(pages) != 3 or _object(pages[-1], f"{route}.semantic.pages[-1]").get("name") != "Replacement Page":
            raise RunFailure(f"{route}: replacement semantics are missing")
    elif route == "semantic_two_disjoint":
        if shape(0, 0).get("text") != "Joined page one" or shape(1, 1).get("name") != "Joined page two":
            raise RunFailure(f"{route}: two-operation semantics are missing")
    elif route == "semantic_four_disjoint":
        expected = ((0, 0, "text", "Joined page one"), (0, 1, "name", "Joined page one name"),
                    (1, 0, "text", "Joined page two"), (1, 1, "name", "Joined page two name"))
        if any(shape(page, index).get(field) != expected_value
               for page, index, field, expected_value in expected):
            raise RunFailure(f"{route}: four-operation semantics are missing")
    return {"output": output, "published": published, "semantic": value["semantic"]}


def validate_verify_report(report: Mapping[str, Any], endpoint_name: str, fixture: Mapping[str, Any]) -> dict[str, Any]:
    candidate = endpoint_name == "candidate"
    schema = CANDIDATE_SCHEMA if candidate else BASELINE_SCHEMA
    expected_endpoint = "candidate-edit-join-commit" if candidate else "baseline-apply-joined"
    _exact_keys(report, {"schema", "endpoint", "mode", "fixture", "source_package", "routes",
                         "refusals", "capabilities", "checks"}, f"{endpoint_name}.report")
    if (report.get("schema") != schema or report.get("endpoint") != expected_endpoint
            or report.get("mode") != "verify"):
        raise RunFailure(f"{endpoint_name}: correctness report identity mismatch")
    if _digest_object(report.get("fixture"), f"{endpoint_name}.fixture") != {"bytes": fixture["bytes"], "sha256": fixture["sha256"]}:
        raise RunFailure(f"{endpoint_name}: fixture proof differs from manifest")
    source = _digest_object(report.get("source_package"), f"{endpoint_name}.source_package")
    route_value = _object(report.get("routes"), f"{endpoint_name}.routes")
    if set(route_value) != set(ROUTES):
        raise RunFailure(f"{endpoint_name}: route set is incomplete")
    routes = {route: _validate_route(route, _object(route_value[route], f"{endpoint_name}.{route}"), candidate) for route in ROUTES}
    for route in ROUTES:
        if route == "empty_join":
            if routes[route]["output"] != source or routes[route]["published"] != source:
                raise RunFailure(f"{endpoint_name}/{route}: no-op did not preserve exact source bytes")
        elif routes[route]["output"] == source:
            raise RunFailure(f"{endpoint_name}/{route}: changed route produced the source bytes")
    refusals = _object(report.get("refusals"), f"{endpoint_name}.refusals")
    if set(refusals) != {"stale_lineage_refused", "conflicting_effect_refused",
                         "changed_recipient_refused", "changed_recipient_scope"}:
        raise RunFailure(f"{endpoint_name}: refusal proof fields are incomplete")
    if refusals["stale_lineage_refused"] is not True or refusals["conflicting_effect_refused"] is not True:
        raise RunFailure(f"{endpoint_name}: stale/conflict refusal proof failed")
    if refusals["changed_recipient_scope"] != "candidate-only":
        raise RunFailure(f"{endpoint_name}: changed-recipient scope is not explicit")
    if refusals["changed_recipient_refused"] is not candidate:
        raise RunFailure(f"{endpoint_name}: changed-recipient capability boundary is incorrect")
    expected_capabilities = {"apply_joined": not candidate, "edit_join_commit": candidate, "commit_triads": candidate,
                             "patch_replay": candidate, "inverse_replay": candidate,
                             "content_matching_replay": candidate, "changed_recipient_refusal": candidate,
                             "changed_recipient_refusal_scope": "candidate-only"}
    if _object(report.get("capabilities"), f"{endpoint_name}.capabilities") != expected_capabilities:
        raise RunFailure(f"{endpoint_name}: capability boundary is incomplete or overstated")
    aggregate = _object(report.get("checks"), f"{endpoint_name}.checks")
    expected_aggregate = {"source_fixture_deterministic", "all_routes_checked", "all_refusal_proofs",
                          "all_candidate_commit_proofs" if candidate else "all_baseline_apply_proofs"}
    if set(aggregate) != expected_aggregate or any(item is not True for item in aggregate.values()):
        raise RunFailure(f"{endpoint_name}: aggregate correctness checks are incomplete")
    return {"source": source, "routes": routes}


def _expected_proof(correctness: Mapping[str, Any], route: str, scope: str) -> dict[str, Any]:
    value = correctness["routes"][route]
    return {"output": value["output" if scope == "join_apply" else "published"], "semantic": value["semantic"]}


def validate_calibration(report: Mapping[str, Any], endpoint: str, route: str, scope: str,
                         protocol: Protocol, expected: Mapping[str, Any]) -> int:
    schema = CANDIDATE_SCHEMA if endpoint == "candidate" else BASELINE_SCHEMA
    expected_endpoint = "candidate-edit-join-commit" if endpoint == "candidate" else "baseline-apply-joined"
    _exact_keys(report, {"schema", "endpoint", "mode", "route", "scope", "target_ns", "iterations",
                         "elapsed_ns", "attempts", "iteration_policy", "initial_iterations",
                         "max_iterations"}, f"{endpoint}/{route}/{scope}.calibration")
    if (report.get("schema") != schema or report.get("endpoint") != expected_endpoint
            or report.get("mode") != "calibrate" or report.get("route") != route
            or report.get("scope") != scope):
        raise RunFailure(f"{endpoint}/{route}/{scope}: calibration identity mismatch")
    if report.get("target_ns") != protocol.target_ns:
        raise RunFailure(f"{endpoint}/{route}/{scope}: calibration target mismatch")
    if (report.get("iteration_policy") != CALIBRATION_POLICY
            or report.get("initial_iterations") != 1
            or report.get("max_iterations") != protocol.max_iterations):
        raise RunFailure(f"{endpoint}/{route}/{scope}: calibration iteration contract mismatch")
    iterations = _positive_int(report.get("iterations"), "calibration.iterations")
    elapsed = _positive_int(report.get("elapsed_ns"), "calibration.elapsed_ns")
    if iterations > protocol.max_iterations or elapsed < protocol.target_ns:
        raise RunFailure(f"{endpoint}/{route}/{scope}: calibration target not reached")
    attempts = report.get("attempts")
    if not isinstance(attempts, list) or not attempts:
        raise RunFailure(f"{endpoint}/{route}/{scope}: calibration attempts missing")
    previous = 0
    for index, item in enumerate(attempts):
        attempt = _object(item, f"{endpoint}/{route}/{scope}.attempts[{index}]")
        _exact_keys(attempt, {"iterations", "elapsed_ns", "proof"},
                    f"{endpoint}/{route}/{scope}.attempts[{index}]")
        count = _positive_int(attempt.get("iterations"), "calibration attempt iterations")
        duration = _positive_int(attempt.get("elapsed_ns"), "calibration attempt elapsed_ns")
        expected_count = 1 if index == 0 else min(previous * 2, protocol.max_iterations)
        if (count != expected_count or count > protocol.max_iterations or duration <= 0
                or attempt.get("proof") != expected):
            raise RunFailure(f"{endpoint}/{route}/{scope}: malformed calibration attempt")
        previous = count
        if index == len(attempts) - 1 and (count != iterations or duration != elapsed):
            raise RunFailure(f"{endpoint}/{route}/{scope}: final calibration attempt mismatch")
    return iterations


def validate_timed(report: Mapping[str, Any], endpoint: str, route: str, scope: str,
                   iterations: int, protocol: Protocol, expected: Mapping[str, Any]) -> int:
    schema = CANDIDATE_SCHEMA if endpoint == "candidate" else BASELINE_SCHEMA
    expected_endpoint = "candidate-edit-join-commit" if endpoint == "candidate" else "baseline-apply-joined"
    _exact_keys(report, {"schema", "endpoint", "mode", "route", "scope", "iterations", "warmups",
                         "samples", "proof"}, f"{endpoint}/{route}/{scope}.timed")
    if (report.get("schema") != schema or report.get("endpoint") != expected_endpoint
            or report.get("mode") != "timed" or report.get("route") != route
            or report.get("scope") != scope):
        raise RunFailure(f"{endpoint}/{route}/{scope}: timed identity mismatch")
    if report.get("iterations") != iterations or report.get("warmups") != protocol.warmups or report.get("proof") != expected:
        raise RunFailure(f"{endpoint}/{route}/{scope}: timed normalization/proof mismatch")
    samples = report.get("samples")
    if not isinstance(samples, list) or len(samples) != REQUIRED_SAMPLES:
        raise RunFailure(
            f"{endpoint}/{route}/{scope}: timed output must contain exactly "
            f"{REQUIRED_SAMPLES} sample"
        )
    sample = _positive_int(samples[0], "timed sample")
    if sample < protocol.minimum_ns:
        raise RunFailure(f"{endpoint}/{route}/{scope}: timed sample below minimum")
    return sample


def validate_allocation(report: Mapping[str, Any], endpoint: str, route: str, scope: str,
                        protocol: Protocol, expected: Mapping[str, Any]) -> dict[str, Any]:
    schema = CANDIDATE_SCHEMA if endpoint == "candidate" else BASELINE_SCHEMA
    expected_endpoint = "candidate-edit-join-commit" if endpoint == "candidate" else "baseline-apply-joined"
    _exact_keys(report, {"schema", "endpoint", "mode", "route", "scope", "iterations", "warmups", "samples"},
                f"{endpoint}/{route}/{scope}.allocation")
    if (report.get("schema") != schema or report.get("endpoint") != expected_endpoint
            or report.get("mode") != "allocation" or report.get("route") != route
            or report.get("scope") != scope or report.get("iterations") != 1
            or report.get("warmups") != REQUIRED_WARMUPS):
        raise RunFailure(f"{endpoint}/{route}/{scope}: allocation identity mismatch")
    samples = report.get("samples")
    if not isinstance(samples, list) or len(samples) != REQUIRED_SAMPLES:
        raise RunFailure(f"{endpoint}/{route}/{scope}: allocation sample count is not fixed")
    sample = _object(samples[0], f"{endpoint}/{route}/{scope}.allocation.samples[0]")
    _exact_keys(sample, {"metrics", "proof"}, f"{endpoint}/{route}/{scope}.allocation.sample")
    if sample["proof"] != expected:
        raise RunFailure(f"{endpoint}/{route}/{scope}: allocation proof differs from correctness")
    metrics = _object(sample["metrics"], f"{endpoint}/{route}/{scope}.allocation.metrics")
    _exact_keys(metrics, {"tool", "allocation", "latency_claim"}, f"{endpoint}/{route}/{scope}.allocation.metrics")
    if metrics["latency_claim"] != "allocator_instrumented_elapsed_not_latency_claim":
        raise RunFailure(f"{endpoint}/{route}/{scope}: allocation report made a latency claim")
    tool = _object(metrics["tool"], f"{endpoint}/{route}/{scope}.allocation.tool")
    _exact_keys(tool, {"binary", "instrumentation", "allocator_counter_revision"},
                f"{endpoint}/{route}/{scope}.allocation.tool")
    if (tool["binary"] != "litchi-perf-odg-joined-qualification-alloc"
            or tool["instrumentation"] != "system_allocator_operation_scoped"
            or tool["allocator_counter_revision"] != "serialized_region_peak_v3"):
        raise RunFailure(f"{endpoint}/{route}/{scope}: allocator tool identity is incorrect")
    allocation = _object(metrics["allocation"], f"{endpoint}/{route}/{scope}.allocation.allocation")
    fields = {"status", "scope", "allocation_calls", "deallocation_calls", "reallocation_calls",
              "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before",
              "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes"}
    fields.add("live_underflow")
    _exact_keys(allocation, fields, f"{endpoint}/{route}/{scope}.allocation.allocation")
    if allocation["status"] != "measured" or allocation["scope"] != "operation_global_system_allocator":
        raise RunFailure(f"{endpoint}/{route}/{scope}: allocation observation is not measured")
    numbers: dict[str, int] = {}
    for field in fields - {"status", "scope"}:
        if field == "live_underflow":
            if type(allocation.get(field)) is not bool:
                raise RunFailure(f"{endpoint}/{route}/{scope}: allocator underflow flag is malformed")
        else:
            numbers[field] = _nonnegative_int(allocation.get(field), f"allocation.{field}")
    if allocation["live_underflow"]:
        raise RunFailure(f"{endpoint}/{route}/{scope}: allocator live-byte underflow was observed")
    if (numbers["live_bytes_after"] - numbers["live_bytes_before"]
            != numbers["allocated_bytes"] - numbers["deallocated_bytes"]):
        raise RunFailure(f"{endpoint}/{route}/{scope}: allocation live-byte conservation failed")
    if (numbers["peak_live_bytes_before"] < numbers["live_bytes_before"]
            or numbers["peak_live_bytes_after"] < numbers["peak_live_bytes_before"]
            or numbers["peak_live_bytes_after"] < numbers["live_bytes_after"]
            or numbers["region_peak_live_bytes"] > numbers["peak_live_bytes_after"]
            or numbers["region_peak_live_bytes"] < numbers["live_bytes_before"]
            or numbers["region_peak_live_bytes"] < numbers["live_bytes_after"]
            or numbers["reallocation_calls"] > numbers["allocation_calls"]):
        raise RunFailure(f"{endpoint}/{route}/{scope}: allocation peak bounds failed")
    return {"metrics": metrics, "proof": expected}


def validate_allocation_rows(rows: Any, key: str) -> list[dict[str, Any]]:
    if not isinstance(rows, list) or len(rows) != len(ABBA_ORDER):
        raise RunFailure(f"{key}: allocation ABBA must contain exactly four legs")
    normalized: list[dict[str, Any]] = []
    for index, (value, role) in enumerate(zip(rows, ABBA_ORDER)):
        row = _object(value, f"{key}.rows[{index}]")
        _exact_keys(row, {"role", "endpoint", "guard", "target", "logs", "allocation", "proof"},
                    f"{key}.rows[{index}]")
        if row["role"] != role or row["endpoint"] != ROLE_ENDPOINT[role]:
            raise RunFailure(f"{key}: allocation leg {index} has the wrong endpoint")
        _object(row["guard"], f"{key}.rows[{index}].guard")
        _object(row["target"], f"{key}.rows[{index}].target")
        _object(row["logs"], f"{key}.rows[{index}].logs")
        _object(row["allocation"], f"{key}.rows[{index}].allocation")
        _object(row["proof"], f"{key}.rows[{index}].proof")
        normalized.append(row)
    return normalized


def validate_abba_rows(rows: Any, key: str) -> list[dict[str, Any]]:
    if not isinstance(rows, list) or len(rows) != len(ABBA_ORDER):
        raise RunFailure(f"{key}: ABBA must contain exactly {len(ABBA_ORDER)} legs")
    normalized: list[dict[str, Any]] = []
    expected_fields = {"role", "endpoint", "iterations", "sample_ns", "per_operation_ns",
                       "guard", "target", "logs", "proof"}
    for index, (row_value, role) in enumerate(zip(rows, ABBA_ORDER)):
        row = _object(row_value, f"{key}.rows[{index}]")
        _exact_keys(row, expected_fields, f"{key}.rows[{index}]")
        if row.get("role") != role or row.get("endpoint") != ROLE_ENDPOINT[role]:
            raise RunFailure(f"{key}: ABBA leg {index} has the wrong role or endpoint")
        _positive_int(row.get("iterations"), f"{key}.rows[{index}].iterations")
        _positive_int(row.get("sample_ns"), f"{key}.rows[{index}].sample_ns")
        _positive_float(row.get("per_operation_ns"), f"{key}.rows[{index}].per_operation_ns")
        guard = _object(row.get("guard"), f"{key}.rows[{index}].guard")
        _exact_keys(guard, {"phase", "load_1m", "max_load", "heavy_processes", "checked_unix_ns"},
                    f"{key}.rows[{index}].guard")
        target = _object(row.get("target"), f"{key}.rows[{index}].target")
        _exact_keys(target, {"targets", "temporary", "phase", "checked_unix_ns"}, f"{key}.rows[{index}].target")
        targets = target["targets"]
        if not isinstance(targets, dict) or not targets:
            raise RunFailure(f"{key}.rows[{index}].target target set is malformed")
        for target_name, target_value in targets.items():
            target_item = _object(target_value, f"{key}.rows[{index}].target.targets[{target_name}]")
            _exact_keys(target_item, {"path", "bytes", "preflight_limit", "hard_limit"},
                        f"{key}.rows[{index}].target.targets[{target_name}]")
            if (type(target_item["bytes"]) is not int or target_item["bytes"] < 0
                    or target_item["preflight_limit"] != TARGET_PREFLIGHT_BYTES
                    or target_item["hard_limit"] != TARGET_HARD_LIMIT_BYTES):
                raise RunFailure(f"{key}.rows[{index}].target target identity is malformed")
        temporary = _object(target["temporary"], f"{key}.rows[{index}].target.temporary")
        _exact_keys(temporary, {"path", "bytes", "preflight_limit", "hard_limit"},
                    f"{key}.rows[{index}].target.temporary")
        if (type(temporary["bytes"]) is not int or temporary["bytes"] < 0
                or temporary["preflight_limit"] != TARGET_PREFLIGHT_BYTES
                or temporary["hard_limit"] != TARGET_HARD_LIMIT_BYTES):
            raise RunFailure(f"{key}.rows[{index}].target temporary identity is malformed")
        logs = _object(row.get("logs"), f"{key}.rows[{index}].logs")
        _exact_keys(logs, {"command", "returncode", "endpoint", "profile", "resolved_ref", "binary",
                           "environment", "binding", "stdout", "stderr", "guard", "target"},
                    f"{key}.rows[{index}].logs")
        for stream in ("stdout", "stderr"):
            _digest_object(logs[stream], f"{key}.rows[{index}].logs.{stream}")
        proof = _object(row.get("proof"), f"{key}.rows[{index}].proof")
        if set(proof) != {"output", "semantic"}:
            raise RunFailure(f"{key}.rows[{index}].proof identity is incomplete")
        _digest_object(proof["output"], f"{key}.rows[{index}].proof.output")
        _object(proof["semantic"], f"{key}.rows[{index}].proof.semantic")
        normalized.append(row)
    return normalized


def _write_json(path: Path, value: Any) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    if path.exists() or temporary.exists() or path.is_symlink() or temporary.is_symlink():
        raise RunFailure(f"refusing to overwrite existing JSON receipt: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with temporary.open("xb") as stream:
            stream.write((json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8"))
            stream.flush()
            os.fsync(stream.fileno())
        # Link publication is exclusive on the same filesystem; unlike
        # `replace`, it cannot silently overwrite a concurrently-created
        # receipt.
        os.link(temporary, path)
        temporary.unlink()
    except OSError as error:
        raise RunFailure(f"cannot publish JSON receipt {path}: {error}") from error


def _plan(manifest: Manifest, out_dir: Path) -> dict[str, Any]:
    _validate_layout(manifest, out_dir)
    return {"schema": REPORT_SCHEMA, "claim_scope": "plan-only", "manifest": str(manifest.path),
            "manifest_sha256": manifest.sha256, "profiles": {"baseline": BASELINE_PROFILE, "candidate": CANDIDATE_PROFILE},
            "routes": list(ROUTES), "scopes": list(SCOPES), "abba_order": list(ABBA_ORDER),
            "composition_limits": manifest.limits,
            "protocol": {"cpu": manifest.protocol.cpu, "blocks": manifest.protocol.blocks,
                          "warmups": manifest.protocol.warmups, "target_ns": manifest.protocol.target_ns,
                          "minimum_ns": manifest.protocol.minimum_ns, "max_iterations": manifest.protocol.max_iterations,
                          "calibration_policy": CALIBRATION_POLICY}}


def _allocator_endpoints(endpoints: Mapping[str, ResolvedEndpoint]) -> dict[str, ResolvedEndpoint]:
    return {
        name: replace(
            resolved,
            endpoint=replace(
                resolved.endpoint,
                binary=resolved.endpoint.allocator_binary,
                target_dir=resolved.endpoint.allocator_target_dir,
            ),
            binary=resolved.allocator_binary,
        )
        for name, resolved in endpoints.items()
    }


def _execute(manifest: Manifest, out_dir: Path, correctness_only: bool, allocation_only: bool = False) -> dict[str, Any]:
    if out_dir.exists():
        raise RunFailure(f"output directory must not already exist: {out_dir}")
    if manifest.tmpdir.exists():
        raise RunFailure(f"temporary directory must not already exist: {manifest.tmpdir}")
    out_dir.mkdir(parents=True, exist_ok=False)
    manifest.tmpdir.mkdir(parents=True, exist_ok=False)
    identity = runtime_identity(manifest, out_dir)
    endpoints: Mapping[str, ResolvedEndpoint] = identity["endpoints"]
    target_before = _resource_guard(manifest, endpoints, "before-correctness")
    correctness_dir = out_dir / "correctness"
    correctness_dir.mkdir()
    baseline_report, baseline_logs = _run_probe(endpoints["baseline"], manifest, correctness_dir, "verify", "baseline-verify")
    candidate_report, candidate_logs = _run_probe(endpoints["candidate"], manifest, correctness_dir, "verify", "candidate-verify")
    _validate_log_binding(baseline_logs, endpoints["baseline"], "verify", None, None, "baseline-verify")
    _validate_log_binding(candidate_logs, endpoints["candidate"], "verify", None, None, "candidate-verify")
    fixture = identity["fixture"]
    baseline = validate_verify_report(baseline_report, "baseline", fixture)
    candidate = validate_verify_report(candidate_report, "candidate", fixture)
    if baseline["source"] != candidate["source"]:
        raise RunFailure("baseline and candidate constructed different source packages")
    for route in ROUTES:
        if baseline["routes"][route] != candidate["routes"][route]:
            raise RunFailure(f"{route}: output or semantic proof differs between endpoints")
    target_after = _resource_guard(manifest, endpoints, "after-correctness")
    result: dict[str, Any] = {
        "schema": REPORT_SCHEMA,
        "claim_scope": "correctness-only" if correctness_only else "correctness-calibration-abba",
        "manifest": str(manifest.path), "manifest_sha256": manifest.sha256,
        "routes": list(ROUTES), "scopes": list(SCOPES), "abba_order": list(ABBA_ORDER),
        "composition_limits": manifest.limits,
        "protocol": {"cpu": manifest.protocol.cpu, "blocks": manifest.protocol.blocks,
                      "warmups": manifest.protocol.warmups, "target_ns": manifest.protocol.target_ns,
                      "minimum_ns": manifest.protocol.minimum_ns,
                      "max_iterations": manifest.protocol.max_iterations,
                      "calibration_policy": CALIBRATION_POLICY},
        "identity": {key: value for key, value in identity.items() if key != "endpoints"},
        "target_guards": [target_before, target_after],
        "correctness": {"baseline": baseline, "candidate": candidate,
                         "logs": {"baseline": baseline_logs, "candidate": candidate_logs}},
        "identity_checks": [_assert_unchanged(manifest, endpoints, "after-correctness")],
    }
    if correctness_only:
        _write_json(out_dir / "report.json", result)
        return result

    if allocation_only:
        allocation_endpoints = _allocator_endpoints(endpoints)
        allocation_dir = out_dir / "allocation"
        allocation_dir.mkdir()
        allocation: dict[str, Any] = {}
        for route in ROUTES:
            for scope in SCOPES:
                key = f"{route}/{scope}"
                blocks = []
                for block in range(REQUIRED_BLOCKS):
                    rows = []
                    for role in ABBA_ORDER:
                        endpoint_name = ROLE_ENDPOINT[role]
                        guard = quiet_guard(manifest.protocol, f"before-allocation-{block}-{role}-{key}")
                        target = _resource_guard(manifest, allocation_endpoints, f"before-allocation-{block}-{role}-{key}")
                        report, logs = _run_probe(
                            allocation_endpoints[endpoint_name], manifest, allocation_dir, "allocation",
                            f"allocation-block-{block}-{role}-{route}-{scope}",
                            route=route, scope=scope, samples=REQUIRED_SAMPLES,
                            warmups=REQUIRED_WARMUPS, iterations=1,
                        )
                        _validate_log_binding(logs, allocation_endpoints[endpoint_name], "allocation", route, scope,
                                             f"allocation-{role}-{route}-{scope}")
                        expected = _expected_proof(baseline if endpoint_name == "baseline" else candidate, route, scope)
                        metrics = validate_allocation(report, endpoint_name, route, scope, manifest.protocol, expected)
                        rows.append({"role": role, "endpoint": endpoint_name, "guard": guard, "target": target,
                                     "logs": logs, "allocation": metrics["metrics"], "proof": expected})
                        _assert_unchanged(manifest, endpoints, f"after-allocation-{block}-{role}-{key}")
                    blocks.append({"block": block, "order": list(ABBA_ORDER),
                                   "rows": validate_allocation_rows(rows, f"{key}/block-{block}")})
                allocation[key] = {"blocks": blocks, "order": list(ABBA_ORDER)}
        result["claim_scope"] = "correctness-allocation"
        result["allocation"] = allocation
        result["identity_checks"].append(_assert_unchanged(manifest, endpoints, "after-allocation"))
        _write_json(out_dir / "report.json", result)
        return result

    calibration_dir = out_dir / "calibration"
    calibration_dir.mkdir()
    calibrations: dict[str, Any] = {}
    for route in ROUTES:
        for scope in SCOPES:
            key = f"{route}/{scope}"
            calibrations[key] = {}
            for endpoint_name in ("baseline", "candidate"):
                guard = quiet_guard(manifest.protocol, f"before-calibration-{endpoint_name}-{key}")
                target = _resource_guard(manifest, endpoints, f"before-calibration-{endpoint_name}-{key}")
                report, logs = _run_probe(endpoints[endpoint_name], manifest, calibration_dir, "calibrate",
                                          f"{endpoint_name}-{route}-{scope}", route=route, scope=scope,
                                          warmups=manifest.protocol.warmups, target_ns=manifest.protocol.target_ns)
                _validate_log_binding(logs, endpoints[endpoint_name], "calibrate", route, scope,
                                     f"{endpoint_name}-{route}-{scope}")
                expected = _expected_proof(baseline if endpoint_name == "baseline" else candidate, route, scope)
                iterations = validate_calibration(report, endpoint_name, route, scope, manifest.protocol, expected)
                calibrations[key][endpoint_name] = {"iterations": iterations, "guard": guard,
                                                     "target": target, "logs": logs, "proof": expected}
                _assert_unchanged(manifest, endpoints, f"after-calibration-{endpoint_name}-{key}")
    result["calibration"] = calibrations
    result["identity_checks"].append(_assert_unchanged(manifest, endpoints, "after-calibration"))

    abba_dir = out_dir / "abba"
    abba_dir.mkdir()
    abba: dict[str, Any] = {}
    for route in ROUTES:
        for scope in SCOPES:
            key = f"{route}/{scope}"
            blocks = []
            for block in range(REQUIRED_BLOCKS):
                rows = []
                for role in ABBA_ORDER:
                    endpoint_name = ROLE_ENDPOINT[role]
                    guard = quiet_guard(manifest.protocol, f"before-abba-{block}-{role}-{key}")
                    target = _resource_guard(manifest, endpoints, f"before-abba-{block}-{role}-{key}")
                    iterations = calibrations[key][endpoint_name]["iterations"]
                    report, logs = _run_probe(
                        endpoints[endpoint_name], manifest, abba_dir, "timed",
                        f"block-{block}-{role}-{route}-{scope}", route=route, scope=scope,
                        iterations=iterations, samples=1, warmups=manifest.protocol.warmups,
                    )
                    _validate_log_binding(logs, endpoints[endpoint_name], "timed", route, scope,
                                         f"{role}-{route}-{scope}")
                    expected = _expected_proof(baseline if endpoint_name == "baseline" else candidate, route, scope)
                    sample = validate_timed(report, endpoint_name, route, scope, iterations,
                                           manifest.protocol, expected)
                    rows.append({"role": role, "endpoint": endpoint_name, "iterations": iterations,
                                 "sample_ns": sample, "per_operation_ns": sample / iterations,
                                 "guard": guard, "target": target, "logs": logs, "proof": expected})
                    _assert_unchanged(manifest, endpoints, f"after-abba-{block}-{role}-{key}")
                normalized_rows = validate_abba_rows(rows, f"{key}/block-{block}")
                blocks.append({"block": block, "order": list(ABBA_ORDER), "rows": normalized_rows})
            abba[key] = {"blocks": blocks, "order": list(ABBA_ORDER)}
    result["abba"] = abba
    result["identity_checks"].append(_assert_unchanged(manifest, endpoints, "after-abba"))
    _write_json(out_dir / "report.json", result)
    return result


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", required=True, type=Path)
    parser.add_argument("--out-dir", required=True, type=Path)
    parser.add_argument("--plan", action="store_true")
    parser.add_argument("--prepare", action="store_true", help="create a fresh source/build preparation receipt")
    parser.add_argument("--build", action="store_true", help="build both paired probe kinds and write a receipt")
    parser.add_argument("--validate-build", action="store_true", help="rehash and validate the build receipt")
    parser.add_argument("--correctness-only", action="store_true")
    parser.add_argument("--allocation-only", action="store_true")
    args = parser.parse_args(argv)
    selected = sum(bool(getattr(args, name)) for name in (
        "plan", "prepare", "build", "validate_build", "correctness_only", "allocation_only"
    ))
    if selected > 1:
        parser.error("qualification phase flags are mutually exclusive")
    return args


def _failure_receipt(path: Path, error: BaseException) -> None:
    try:
        if path.exists():
            # A failure must never turn an existing artifact directory into a
            # mutable log sink or overwrite a prior receipt.
            return
        path.mkdir(parents=True, exist_ok=False)
        receipt = path / "failure.json"
        _write_json(receipt, {"schema": REPORT_SCHEMA, "status": "failed", "error": str(error)})
    except (OSError, RunFailure):
        return


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    out_dir = args.out_dir.absolute()
    try:
        manifest = load_manifest(args.manifest)
        if args.plan:
            if out_dir.exists():
                raise RunFailure(f"output directory must not already exist: {out_dir}")
            out_dir.mkdir(parents=True, exist_ok=False)
            _write_json(out_dir / "plan.json", _plan(manifest, out_dir))
        elif args.prepare:
            try:
                from odg_joined_qualification_builder import prepare
            except ImportError:  # pragma: no cover
                from tools.odg_joined_qualification_builder import prepare  # type: ignore[no-redef]
            prepare(manifest.path)
        elif args.build:
            try:
                from odg_joined_qualification_builder import build
            except ImportError:  # pragma: no cover
                from tools.odg_joined_qualification_builder import build  # type: ignore[no-redef]
            build(manifest.path)
        elif args.validate_build:
            try:
                from odg_joined_qualification_builder import validate_build_receipt
            except ImportError:  # pragma: no cover
                from tools.odg_joined_qualification_builder import validate_build_receipt  # type: ignore[no-redef]
            validate_build_receipt(manifest.path)
        else:
            _execute(manifest, out_dir, args.correctness_only, args.allocation_only)
    except (RunFailure, OSError, ValueError) as error:
        _failure_receipt(out_dir, error)
        print(str(error), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
