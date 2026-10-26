#!/usr/bin/env python3
"""Fail-closed native qualification runner for the PPT lazy-entry lane.

The manifest names two already-built endpoint pairs.  This runner never builds
or edits a source tree: it checks each clean Git ref, the separate perf lock,
the probe sources, both normal/allocator executable identities, and external
target directories before running a correctness preflight and pinned ABBA.
The checked-in manifest is deliberately parameterized until the final
candidate is frozen; ``--plan`` is the only mode that accepts that template.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import os
from pathlib import Path
import re
import subprocess
import sys
import time
from typing import Any, Mapping, Sequence

try:
    # Prefer package-qualified imports even when another tooling test has
    # inserted ``tools/`` at the front of ``sys.path``.  Mixing a top-level
    # ``odp_native_qualification`` with the package-qualified builder creates
    # two distinct RunFailure classes; combined unittest runs then fail to
    # catch the builder's intentional guard refusals.
    from tools.odg_joined_qualification_builder import (
        _check_worktree_shape,
        _directory_bytes,
        _quiet_guard,
        _stream_command,
    )
    from tools.odp_native_qualification import RunFailure, _digest, binary_identity
except ImportError:  # pragma: no cover - direct package/module loading.
    from odg_joined_qualification_builder import (  # type: ignore[no-redef]
        _check_worktree_shape,
        _directory_bytes,
        _quiet_guard,
        _stream_command,
    )
    from odp_native_qualification import (  # type: ignore[no-redef]
        RunFailure,
        _digest,
        binary_identity,
    )


INPUT_SCHEMA = "litchi.ppt.lazy-native-qualification.input.v1"
REPORT_SCHEMA = "litchi.ppt.lazy-native-qualification.v1"
PROBE_SCHEMA = "litchi.ppt.lazy-native.probe.v1"
BASELINE_REF = "8779d059add6c90c74552801372c44f770db9f0f"
BASELINE_TREE = "27e0b56e2f8e3ab11b00c119604b5636f358dfe2"
BASELINE_LOCK_BYTES = 57_158
BASELINE_LOCK_SHA256 = "4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184"
BASELINE_COMPILE_FIXTURES = (
    (
        "test-data/poi/test-data/slideshow/ppt_with_embeded.ppt",
        180_736,
        "afa1dd5adc36765e6e40725a133b0330e5ad1febd1506ca93ceae8c5c5400b53",
    ),
    (
        "test-data/unified-ops/0294/video.ppt",
        43_520,
        "f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a",
    ),
    (
        "test-data/poi/test-data/slideshow/bug61881.ppt",
        74_752,
        "ba2bf551e87f5fd7667dec52ac19005396bdf9af00f0006f9a247a9fede2e8ee",
    ),
    (
        "test-data/poi/test-data/spreadsheet/54016.xls",
        984_576,
        "2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a",
    ),
    (
        "test-data/rtf/watermark.rtf",
        69_471,
        "48d62dcd959e737b06ebb8255780bcaaf1e88056ff9c3d5a21d3ff5cd3ddf9cb",
    ),
)
BASELINE_OVERLAY_STATUS = (
    "M tools/perf-baseline/src/lib.rs\n"
    "?? test-data/unified-ops/ppt-lazy-native-fixture.ppt\n"
    "?? tools/perf-baseline/src/bin/ppt_lazy_native.rs\n"
    "?? tools/perf-baseline/src/bin/ppt_lazy_native_alloc.rs\n"
    "?? tools/perf-baseline/src/ppt_lazy_native.rs"
)
BASELINE_PROBE_PATHS = (
    "tools/perf-baseline/src/ppt_lazy_native.rs",
    "tools/perf-baseline/src/bin/ppt_lazy_native.rs",
    "tools/perf-baseline/src/bin/ppt_lazy_native_alloc.rs",
    "tools/perf-baseline/src/lib.rs",
)
ROUTES = (
    "primary_entry_drop",
    "primary_noop_commit",
    "metadata_entry_drop",
    "metadata_noop_commit",
    "animations_entry_drop",
    "animations_noop_commit",
    "broadcast_entry_drop",
    "broadcast_noop_commit",
    "client_data_entry_drop",
    "client_data_noop_commit",
    "diagrams_entry_drop",
    "diagrams_noop_commit",
    "document_comparison_entry_drop",
    "document_comparison_noop_commit",
    "objects_entry_drop",
    "objects_noop_commit",
    "external_media_entry_drop",
    "external_media_noop_commit",
    "fonts_entry_drop",
    "fonts_noop_commit",
    "master_metadata_entry_drop",
    "master_metadata_noop_commit",
    "master_layout_entry_drop",
    "master_layout_noop_commit",
    "picture_bullets_entry_drop",
    "picture_bullets_noop_commit",
    "slide_sync_entry_drop",
    "slide_sync_noop_commit",
)
FIXTURE_GENERATOR = "litchi-ppt-lazy-native-fixture-v2"
FIXTURE_KEYS = ("default", "objects", "external_media", "picture_bullets")
ROUTE_FIXTURE_KEYS = {
    route: (
        "objects"
        if route.startswith("objects_")
        else "external_media"
        if route.startswith("external_media_")
        else "picture_bullets"
        if route.startswith("picture_bullets_")
        else "default"
    )
    for route in ROUTES
}
ABBA_ORDER = ("a1", "b1", "b2", "a2")
ROLE_ENDPOINT = {"a1": "baseline", "a2": "baseline", "b1": "candidate", "b2": "candidate"}
MAX_MANIFEST_BYTES = 256 * 1024
MAX_CHILD_OUTPUT = 4 * 1024 * 1024
TARGET_PREFLIGHT_BYTES = 19_200_000_000
TARGET_HARD_LIMIT_BYTES = 20_000_000_000
MEMORY_HARD_LIMIT_BYTES = 20_000_000_000
REVISION = re.compile(r"[0-9a-f]{40}\Z")
HEX64 = re.compile(r"[0-9a-f]{64}\Z")


def _object(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise RunFailure(f"{label} must be an object")
    return value


def _string(value: Any, label: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str) or (nonempty and not value):
        raise RunFailure(f"{label} must be a non-empty string")
    return value


def _digest_value(value: Any, label: str) -> str:
    value = _string(value, label).lower()
    if HEX64.fullmatch(value) is None:
        raise RunFailure(f"{label} must be lowercase SHA-256")
    return value


def _positive_int(value: Any, label: str) -> int:
    if type(value) is not int or value <= 0:
        raise RunFailure(f"{label} must be a positive integer")
    return value


def _nonnegative_int(value: Any, label: str) -> int:
    if type(value) is not int or value < 0:
        raise RunFailure(f"{label} must be a nonnegative integer")
    return value


def _finite_number(value: Any, label: str, *, minimum: float | None = None) -> float | int:
    if type(value) not in (int, float):
        raise RunFailure(f"{label} must be a finite number")
    try:
        finite = math.isfinite(float(value))
    except (OverflowError, ValueError) as error:
        raise RunFailure(f"{label} must be a finite number") from error
    if not finite:
        raise RunFailure(f"{label} must be a finite number")
    if minimum is not None and value < minimum:
        raise RunFailure(f"{label} must be at least {minimum}")
    return value


def _strict_json(raw: bytes, label: str) -> Any:
    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise RunFailure(f"{label} contains duplicate key {key!r}")
            result[key] = value
        return result

    def reject_constant(constant: str) -> Any:
        raise RunFailure(f"{label} contains non-finite JSON constant {constant!r}")

    try:
        value = json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=unique,
            parse_constant=reject_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RunFailure(f"{label} is not strict UTF-8 JSON: {error}") from error
    def reject_nonfinite(value: Any) -> None:
        if isinstance(value, float) and not math.isfinite(value):
            raise RunFailure(f"{label} contains a non-finite JSON number")
        if isinstance(value, dict):
            for nested in value.values():
                reject_nonfinite(nested)
        elif isinstance(value, list):
            for nested in value:
                reject_nonfinite(nested)
    reject_nonfinite(value)
    return value


def _canonical_json(value: Mapping[str, Any]) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")


def _build_input_projection_value(manifest: Mapping[str, Any]) -> dict[str, Any]:
    """Return the canonical manifest with only receipt content identity removed."""

    value = copy.deepcopy(dict(manifest))
    try:
        receipt = value["build"]["receipt"]
    except (KeyError, TypeError) as error:
        raise RunFailure("manifest.build.receipt is missing") from error
    if not isinstance(receipt, dict) or set(receipt) != {"path", "bytes", "sha256"}:
        raise RunFailure("manifest.build.receipt identity fields differ")
    del receipt["bytes"]
    del receipt["sha256"]
    return value


def _validate_build_input_projection(
    manifest: Mapping[str, Any],
    manifest_path: Path,
    projection_path: Path | None,
) -> dict[str, Any] | None:
    if projection_path is None:
        return None
    projection_path = _safe_path(projection_path.resolve(), "build-input projection", must_exist=True)
    raw = projection_path.read_bytes()
    projection = _strict_json(raw, "build-input projection")
    if not isinstance(projection, dict):
        raise RunFailure("build-input projection must be an object")
    expected = _build_input_projection_value(manifest)
    expected_raw = _canonical_json(expected)
    if raw != expected_raw:
        raise RunFailure("build-input projection differs outside receipt identity fields or is not canonical")
    identity = _identity(projection_path)
    return {
        "path": str(projection_path),
        "bytes": identity["bytes"],
        "sha256": identity["sha256"],
    }


def _load_json(path: Path) -> dict[str, Any]:
    try:
        raw = path.read_bytes()
    except OSError as error:
        raise RunFailure(f"cannot read manifest {path}: {error}") from error
    if len(raw) > MAX_MANIFEST_BYTES:
        raise RunFailure(f"manifest exceeds {MAX_MANIFEST_BYTES} bytes")

    value = _strict_json(raw, "manifest")
    return _object(value, "manifest")


def _exact_keys(value: Mapping[str, Any], expected: set[str], label: str) -> None:
    if set(value) != expected:
        raise RunFailure(f"{label} fields differ from the protocol")


def _identity_spec(value: Any, label: str, *, plan: bool) -> dict[str, Any]:
    item = _object(value, label)
    _exact_keys(item, {"path", "bytes", "sha256"}, label)
    _string(item.get("path"), f"{label}.path")
    if plan and item.get("bytes") is None and item.get("sha256") is None:
        return item
    _positive_int(item.get("bytes"), f"{label}.bytes")
    _digest_value(item.get("sha256"), f"{label}.sha256")
    return item


def _endpoint_spec(value: Any, label: str, *, plan: bool) -> dict[str, Any]:
    item = _object(value, label)
    expected = {
        "root",
        "ref",
        "tree",
        "lock",
        "normal_binary",
        "allocator_binary",
        "target_dir",
        "allocator_target_dir",
    }
    _exact_keys(item, expected, label)
    _string(item.get("root"), f"{label}.root")
    ref = _string(item.get("ref"), f"{label}.ref")
    if plan and ref == "PARAMETERIZED_FINAL_CANDIDATE":
        pass
    elif REVISION.fullmatch(ref) is None:
        raise RunFailure(f"{label}.ref must be a full commit SHA")
    tree = _string(item.get("tree"), f"{label}.tree")
    if plan and tree == "PARAMETERIZED_FINAL_TREE":
        pass
    elif REVISION.fullmatch(tree) is None:
        raise RunFailure(f"{label}.tree must be a full tree SHA")
    lock = _identity_spec(item.get("lock"), f"{label}.lock", plan=plan)
    if lock["path"] != "tools/perf-baseline/Cargo.lock":
        raise RunFailure(f"{label}.lock.path must be tools/perf-baseline/Cargo.lock")
    _identity_spec(item.get("normal_binary"), f"{label}.normal_binary", plan=plan)
    _identity_spec(item.get("allocator_binary"), f"{label}.allocator_binary", plan=plan)
    _string(item.get("target_dir"), f"{label}.target_dir")
    _string(item.get("allocator_target_dir"), f"{label}.allocator_target_dir")
    return item


def _build_spec(value: Any, *, plan: bool) -> dict[str, Any]:
    build = _object(value, "manifest.build")
    _exact_keys(
        build,
        {
            "preparation",
            "receipt",
            "source_archive",
            "fixtures",
            "fixture_routes",
            "toolchain",
            "cargo_config",
            "profile",
            "target_subdir",
            "normal_command",
            "allocator_command",
            "features",
            "environment",
        },
        "manifest.build",
    )
    for name in ("preparation", "receipt", "source_archive", "toolchain", "cargo_config"):
        _identity_spec(build.get(name), f"manifest.build.{name}", plan=plan)
    if build.get("profile") != "release" or build.get("target_subdir") != "target/release":
        raise RunFailure("manifest.build must pin the release profile and target/release")
    expected_commands = {
        "normal_command": [
            "cargo",
            "build",
            "--release",
            "--offline",
            "--locked",
            "--manifest-path",
            "tools/perf-baseline/Cargo.toml",
            "--bin",
            "ppt_lazy_native",
        ],
        "allocator_command": [
            "cargo",
            "build",
            "--release",
            "--offline",
            "--locked",
            "--manifest-path",
            "tools/perf-baseline/Cargo.toml",
            "--bin",
            "ppt_lazy_native_alloc",
        ],
    }
    for name, expected in expected_commands.items():
        if build.get(name) != expected:
            raise RunFailure(f"manifest.build.{name} differs from the fixed command")
    features = _object(build.get("features"), "manifest.build.features")
    _exact_keys(features, {"normal", "allocator"}, "manifest.build.features")
    expected_features = {
        "normal": ["current-api", "current-xlsx-raw-opc"],
        "allocator": ["current-api", "current-xlsx-raw-opc"],
    }
    if features != expected_features:
        raise RunFailure("manifest.build.features differs from the fixed native probes")
    fixtures = _object(build.get("fixtures"), "manifest.build.fixtures")
    _exact_keys(fixtures, set(FIXTURE_KEYS), "manifest.build.fixtures")
    for key in FIXTURE_KEYS:
        _identity_spec(fixtures.get(key), f"manifest.build.fixtures.{key}", plan=plan)
    fixture_routes = _object(build.get("fixture_routes"), "manifest.build.fixture_routes")
    _exact_keys(fixture_routes, set(ROUTES), "manifest.build.fixture_routes")
    if fixture_routes != ROUTE_FIXTURE_KEYS:
        raise RunFailure("manifest.build.fixture_routes differs from the reviewed route matrix")
    environment = _object(build.get("environment"), "manifest.build.environment")
    if environment != {
        "CARGO_BUILD_JOBS": "1",
        "CARGO_INCREMENTAL": "0",
        "CARGO_NET_OFFLINE": "true",
        "CARGO_PROFILE_RELEASE_DEBUG": "0",
    }:
        raise RunFailure("manifest.build.environment differs from the fixed build environment")
    return build


def _load_manifest(path: Path, *, plan: bool) -> dict[str, Any]:
    value = _load_json(path)
    _exact_keys(
        value,
        {
            "schema",
            "baseline",
            "candidate",
            "probe_sources",
            "baseline_adapter",
            "build",
            "protocol",
            "claim_scope",
        },
        "manifest",
    )
    if value.get("schema") != INPUT_SCHEMA:
        raise RunFailure(f"manifest.schema must be {INPUT_SCHEMA!r}")
    baseline = _endpoint_spec(value.get("baseline"), "manifest.baseline", plan=plan)
    if baseline["ref"] != BASELINE_REF:
        raise RunFailure(f"manifest.baseline.ref must be the true pre-lazy parent {BASELINE_REF}")
    if baseline["tree"] != BASELINE_TREE:
        raise RunFailure(f"manifest.baseline.tree must be {BASELINE_TREE}")
    lock = baseline["lock"]
    if lock["path"] != "tools/perf-baseline/Cargo.lock":
        raise RunFailure("manifest.baseline.lock.path must be tools/perf-baseline/Cargo.lock")
    if lock["bytes"] != BASELINE_LOCK_BYTES or lock["sha256"] != BASELINE_LOCK_SHA256:
        raise RunFailure("manifest.baseline lock identity is not the pinned pre-lazy lock")

    candidate_value = value.get("candidate")
    if candidate_value is None:
        if not plan:
            raise RunFailure("candidate endpoint is still parameterized; final freeze is required")
    else:
        _endpoint_spec(candidate_value, "manifest.candidate", plan=plan)
        if candidate_value["ref"] == BASELINE_REF:
            raise RunFailure("candidate must differ from the true pre-lazy baseline")

    sources = value.get("probe_sources")
    if not isinstance(sources, list) or not sources:
        raise RunFailure("manifest.probe_sources must be a non-empty list")
    if [source.get("path") for source in sources] != list(BASELINE_PROBE_PATHS):
        raise RunFailure("manifest.probe_sources differs from the fixed probe allowlist")
    for index, source in enumerate(sources):
        _identity_spec(source, f"manifest.probe_sources[{index}]", plan=plan)
    adapter = _object(value.get("baseline_adapter"), "manifest.baseline_adapter")
    _exact_keys(
        adapter,
        {"overlay_paths", "registration", "fixture", "compile_fixtures", "status"},
        "manifest.baseline_adapter",
    )
    if adapter.get("status") != BASELINE_OVERLAY_STATUS:
        raise RunFailure("manifest.baseline_adapter.status differs from the reviewed overlay")
    if adapter.get("overlay_paths") != [
        "tools/perf-baseline/src/ppt_lazy_native.rs",
        "tools/perf-baseline/src/bin/ppt_lazy_native.rs",
        "tools/perf-baseline/src/bin/ppt_lazy_native_alloc.rs",
    ]:
        raise RunFailure("manifest.baseline_adapter.overlay_paths differs from the reviewed overlay")
    registration = _object(adapter.get("registration"), "manifest.baseline_adapter.registration")
    _exact_keys(registration, {"path", "line"}, "manifest.baseline_adapter.registration")
    if registration != {
        "path": "tools/perf-baseline/src/lib.rs",
        "line": "pub mod ppt_lazy_native;",
    }:
        raise RunFailure("manifest.baseline_adapter.registration differs from the reviewed adapter")
    fixture = _object(adapter.get("fixture"), "manifest.baseline_adapter.fixture")
    _exact_keys(fixture, {"path", "bytes", "sha256"}, "manifest.baseline_adapter.fixture")
    if fixture != {
        "path": "test-data/unified-ops/ppt-lazy-native-fixture.ppt",
        "bytes": 14_848,
        "sha256": "02fa44ca76903943e80a3a2da1f9dbd326c7a4a047540c3aa9cf211e34ef617c",
    }:
        raise RunFailure("manifest.baseline_adapter.fixture differs from the reviewed fixture")
    compile_fixtures = adapter.get("compile_fixtures")
    if not isinstance(compile_fixtures, list):
        raise RunFailure("manifest.baseline_adapter.compile_fixtures must be a list")
    expected_compile_fixtures = [
        {"path": path, "bytes": bytes_, "sha256": sha256}
        for path, bytes_, sha256 in BASELINE_COMPILE_FIXTURES
    ]
    if compile_fixtures != expected_compile_fixtures:
        raise RunFailure("manifest.baseline_adapter.compile_fixtures differs from the reviewed closure")
    _build_spec(value.get("build"), plan=plan)
    protocol = _object(value.get("protocol"), "manifest.protocol")
    _exact_keys(
        protocol,
        {"cpu", "blocks", "warmups", "target_ms", "minimum_ms", "max_load", "max_iterations", "timeout_s"},
        "manifest.protocol",
    )
    if _nonnegative_int(protocol.get("cpu"), "manifest.protocol.cpu") == 0:
        raise RunFailure("manifest.protocol.cpu must be a nonzero pinned CPU")
    if protocol.get("blocks") != 2 or protocol.get("warmups") != 3:
        raise RunFailure("PPT lazy protocol requires two blocks and three warmups")
    _finite_number(protocol.get("target_ms"), "manifest.protocol.target_ms", minimum=500)
    if type(protocol.get("target_ms")) not in (int, float) or protocol["target_ms"] < 500:
        raise RunFailure("PPT lazy protocol requires target_ms >= 500")
    _finite_number(protocol.get("minimum_ms"), "manifest.protocol.minimum_ms", minimum=150)
    if type(protocol.get("minimum_ms")) not in (int, float) or protocol["minimum_ms"] < 150:
        raise RunFailure("PPT lazy protocol requires minimum_ms >= 150")
    if protocol["target_ms"] < protocol["minimum_ms"]:
        raise RunFailure("target_ms must not be below minimum_ms")
    _finite_number(protocol.get("max_load"), "manifest.protocol.max_load", minimum=0)
    if type(protocol.get("max_load")) not in (int, float) or protocol["max_load"] <= 0:
        raise RunFailure("protocol max_load must be positive")
    _positive_int(protocol.get("max_iterations"), "manifest.protocol.max_iterations")
    _finite_number(protocol.get("timeout_s"), "manifest.protocol.timeout_s", minimum=0)
    if type(protocol.get("timeout_s")) not in (int, float) or protocol["timeout_s"] <= 0:
        raise RunFailure("manifest.protocol.timeout_s must be positive")
    if value.get("claim_scope") != (
        "preparation and reproducible native qualification protocol only; no native timing claim"
    ):
        raise RunFailure("manifest.claim_scope is not the source-only preparation claim")
    return value


def _safe_path(value: str, base: Path | None = None, *, must_exist: bool = False) -> Path:
    path = Path(value)
    if not path.is_absolute():
        if base is None:
            raise RunFailure(f"path must be absolute: {value!r}")
        path = base / path
    current = Path(path.anchor)
    for part in path.parts[1:]:
        current /= part
        if current.is_symlink():
            raise RunFailure(f"symlink path component is forbidden: {path}")
    if "3rd" + "party" in path.parts:
        raise RunFailure(f"untracked upstream checkout is forbidden: {path}")
    path = path.resolve(strict=False)
    if must_exist and not path.exists():
        raise RunFailure(f"path does not exist: {path}")
    return path


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _overlap(first: Path, second: Path) -> bool:
    return _under(first, second) or _under(second, first)


def _git(root: Path, *args: str, no_optional_locks: bool = False) -> str:
    import subprocess

    command = ["git", "-C", str(root)]
    if no_optional_locks:
        command.append("--no-optional-locks")
    command.extend(args)
    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        check=False,
        timeout=30,
    )
    if result.returncode:
        raise RunFailure(f"git {' '.join(args)} failed in {root}: {result.stderr.strip()}")
    return result.stdout.strip()


def _sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise RunFailure(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def _identity(path: Path) -> dict[str, Any]:
    sha, size = _sha256_file(path)
    return {"path": str(path), "bytes": size, "sha256": sha}


def _check_identity(spec: Mapping[str, Any], label: str, base: Path, *, executable: bool = False) -> dict[str, Any]:
    path = _safe_path(_string(spec.get("path"), f"{label}.path"), base, must_exist=True)
    actual = _identity(path)
    expected = {"path": str(path), "bytes": spec["bytes"], "sha256": spec["sha256"]}
    if actual != expected:
        raise RunFailure(f"{label} identity differs from the manifest")
    if executable and not os.access(path, os.X_OK):
        raise RunFailure(f"{label} is not executable: {path}")
    return actual


def _check_lock(spec: Mapping[str, Any], label: str, root: Path) -> dict[str, Any]:
    if spec.get("path") != "tools/perf-baseline/Cargo.lock":
        raise RunFailure(f"{label}.path must be tools/perf-baseline/Cargo.lock")
    expected = (root / "tools/perf-baseline/Cargo.lock").resolve(strict=False)
    actual = _check_identity(spec, label, root)
    if Path(actual["path"]) != expected:
        raise RunFailure(f"{label}.path must resolve below its endpoint root")
    return actual


def _baseline_overlay_paths(adapter: Mapping[str, Any]) -> list[str]:
    return sorted(
        [
            adapter["fixture"]["path"],
            adapter["registration"]["path"],
            *adapter["overlay_paths"],
            *[item["path"] for item in adapter["compile_fixtures"]],
        ]
    )


def _check_baseline_overlay_shape(root: Path, label: str, adapter: Mapping[str, Any]) -> dict[str, Any]:
    """Validate the intentionally sparse baseline without admitting a dirty root."""

    top = _safe_path(Path(_git(root, "rev-parse", "--show-toplevel")), f"{label}.root")
    if top != root:
        raise RunFailure(f"{label}.root must be the Git worktree root")
    sparse = _git(root, "config", "--get", "core.sparseCheckout")
    if sparse.lower() not in {"true", "1", "yes", "on"}:
        raise RunFailure(f"{label}.root must use the reviewed sparse baseline checkout")
    status = _git(
        root,
        "status",
        "--porcelain",
        "--untracked-files=all",
        no_optional_locks=True,
    )
    if status != adapter["status"]:
        raise RunFailure(f"{label}.root changed outside the reviewed overlay closure")
    for relative in _baseline_overlay_paths(adapter):
        _safe_path(root / relative, f"{label} overlay {relative}", must_exist=True)
    indexed = subprocess.run(
        ["git", "-C", str(root), "ls-files", "-v", "-z"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
        timeout=30,
    )
    if indexed.returncode != 0:
        raise RunFailure(f"{label}.root cannot inspect sparse index")
    for entry in indexed.stdout.split(b"\0"):
        if entry and entry[:1] == b"h":
            raise RunFailure(f"{label}.root contains an assume-unchanged tracked entry")
    return {"root": str(root), "status": status, "overlay_paths": _baseline_overlay_paths(adapter)}


def _resolve_endpoint(
    raw: Mapping[str, Any],
    label: str,
    manifest_path: Path,
    roots: Sequence[Path],
    *,
    baseline_adapter: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    root = _safe_path(_string(raw["root"], f"{label}.root"), manifest_path.parent, must_exist=True)
    if not root.is_dir():
        raise RunFailure(f"{label}.root is not a directory")
    top = _safe_path(_git(root, "rev-parse", "--show-toplevel"))
    if top != root:
        raise RunFailure(f"{label}.root is not a worktree top-level")
    # Reuse the common worktree-shape validator for the complete candidate.
    # The baseline is intentionally sparse, but only through its exact
    # manifest-bound overlay closure.
    overlay_shape = (
        _check_baseline_overlay_shape(root, label, baseline_adapter)
        if baseline_adapter is not None
        else _check_worktree_shape(root, label)
    )
    head = _git(root, "rev-parse", "HEAD")
    tree = _git(root, "rev-parse", "HEAD^{tree}")
    if head != raw["ref"] or tree != raw["tree"]:
        raise RunFailure(f"{label} HEAD/tree differs from its manifest binding")
    if label == "baseline" and head != BASELINE_REF:
        raise RunFailure("baseline is not the true pre-lazy parent")
    lock = _check_lock(raw["lock"], f"{label}.lock", root)
    if label == "baseline" and (
        lock["bytes"] != BASELINE_LOCK_BYTES or lock["sha256"] != BASELINE_LOCK_SHA256
    ):
        raise RunFailure("baseline lock differs from the pinned pre-lazy lock")
    normal_binary = _check_identity(raw["normal_binary"], f"{label}.normal_binary", manifest_path.parent, executable=True)
    allocator_binary = _check_identity(raw["allocator_binary"], f"{label}.allocator_binary", manifest_path.parent, executable=True)
    target = _safe_path(_string(raw["target_dir"], f"{label}.target_dir"), manifest_path.parent)
    allocator_target = _safe_path(_string(raw["allocator_target_dir"], f"{label}.allocator_target_dir"), manifest_path.parent)
    if not target.is_dir() or not allocator_target.is_dir():
        raise RunFailure(f"{label} target directories must exist")
    for path, name in ((target, "target_dir"), (allocator_target, "allocator_target_dir")):
        if _overlap(path, root) or any(_overlap(path, other) for other in roots if other != root):
            raise RunFailure(f"{label}.{name} must be outside both source worktrees")
    if _under(target, allocator_target) or _under(allocator_target, target):
        raise RunFailure(f"{label} normal and allocator targets overlap")
    if not _under(Path(normal_binary["path"]), target):
        raise RunFailure(f"{label}.normal_binary must be below target_dir")
    if not _under(Path(allocator_binary["path"]), allocator_target):
        raise RunFailure(f"{label}.allocator_binary must be below allocator_target_dir")
    _check_target(target, f"{label}.target_dir")
    _check_target(allocator_target, f"{label}.allocator_target_dir")
    return {
        "name": label,
        "raw": raw,
        "root": root,
        "head": head,
        "tree": tree,
        "lock": lock,
        "normal_binary": normal_binary,
        "allocator_binary": allocator_binary,
        "target_dir": target,
        "allocator_target_dir": allocator_target,
        "overlay_shape": overlay_shape,
        "overlay_adapter": baseline_adapter,
    }


def _check_probe_sources(
    raw_sources: Sequence[Mapping[str, Any]],
    endpoint: Mapping[str, Any],
) -> list[dict[str, Any]]:
    result = []
    for index, spec in enumerate(raw_sources):
        relative = _string(spec.get("path"), f"probe_sources[{index}].path")
        if Path(relative).is_absolute() or "3rd" + "party" in Path(relative).parts:
            raise RunFailure(f"probe_sources[{index}].path must be a reviewed relative source path")
        path = _safe_path(relative, endpoint["root"], must_exist=True)
        if not _under(path, endpoint["root"]):
            raise RunFailure(f"probe_sources[{index}] escapes the candidate source root")
        tracked = _git(endpoint["root"], "ls-files", "-v", "--", relative)
        if tracked != f"H {relative}":
            raise RunFailure(f"probe_sources[{index}] is not an ordinary tracked candidate file")
        actual = _identity(path)
        expected = {"path": relative, "bytes": spec["bytes"], "sha256": spec["sha256"]}
        if spec["bytes"] is None or spec["sha256"] is None:
            raise RunFailure(f"probe_sources[{index}] is not frozen to candidate source bytes")
        if actual != {"path": str(path), "bytes": expected["bytes"], "sha256": expected["sha256"]}:
            raise RunFailure(f"probe_sources[{index}] identity differs from candidate source")
        result.append({"path": str(path), "bytes": actual["bytes"], "sha256": actual["sha256"]})
    return result


def _relative_identity(path: Path, root: Path, label: str) -> dict[str, Any]:
    try:
        relative = path.relative_to(root)
    except ValueError as error:
        raise RunFailure(f"{label} is outside its declared source root") from error
    if relative == Path("."):
        raise RunFailure(f"{label} must name a file below its source root")
    identity = _identity(path)
    return {"path": relative.as_posix(), "bytes": identity["bytes"], "sha256": identity["sha256"]}


def _validate_baseline_overlay_content(
    manifest: Mapping[str, Any],
    receipt: Mapping[str, Any],
    baseline: Mapping[str, Any],
    candidate: Mapping[str, Any],
) -> dict[str, Any]:
    """Bind every reviewed overlay byte at both endpoint roots."""

    adapter = _object(receipt.get("baseline_adapter"), "build receipt baseline_adapter")
    expected = _object(manifest["baseline_adapter"], "manifest.baseline_adapter")
    if adapter.get("overlay_paths") is None or adapter.get("compile_fixtures") is None:
        raise RunFailure("build receipt baseline overlay closure is missing")
    if sorted(
        [
            expected["fixture"]["path"],
            expected["registration"]["path"],
            *expected["overlay_paths"],
            *[item["path"] for item in expected["compile_fixtures"]],
        ]
    ) != _baseline_overlay_paths(expected):
        raise RunFailure("manifest baseline overlay closure is not canonical")
    if baseline.get("overlay_shape", {}).get("status") != expected["status"]:
        raise RunFailure("baseline overlay status is not the manifest-bound status")
    root = baseline["root"]
    candidate_root = candidate["root"]

    def expected_identity(relative: str, bytes_: int, sha256: str, root_path: Path) -> dict[str, Any]:
        return {"path": str(root_path / relative), "bytes": bytes_, "sha256": sha256}

    checked: list[str] = []
    registration = _object(adapter.get("registration"), "build receipt baseline_adapter.registration")
    before = _object(registration.get("before"), "build receipt baseline registration.before")
    actual = _identity(root / registration["path"])
    if actual != expected_identity(registration["path"], before["bytes"], before["sha256"], root):
        raise RunFailure("baseline registration overlay bytes differ from the receipt")
    candidate_registration = _identity(candidate_root / registration["path"])
    after = _object(registration.get("after"), "build receipt baseline registration.after")
    if candidate_registration != expected_identity(registration["path"], after["bytes"], after["sha256"], candidate_root):
        raise RunFailure("candidate registration overlay bytes differ from the receipt")
    checked.append(registration["path"])

    fixture = _object(adapter.get("fixture"), "build receipt baseline_adapter.fixture")
    for root_path, name in ((root, "baseline"), (candidate_root, "candidate")):
        actual = _identity(root_path / fixture["path"])
        if actual != expected_identity(fixture["path"], fixture["bytes"], fixture["sha256"], root_path):
            raise RunFailure(f"{name} PPT fixture overlay bytes differ from the receipt")
    checked.append(fixture["path"])

    for item in [*adapter["overlay_paths"], *adapter["compile_fixtures"]]:
        relative = item["path"]
        for root_path, name in ((root, "baseline"), (candidate_root, "candidate")):
            actual = _identity(root_path / relative)
            if actual != expected_identity(relative, item["bytes"], item["sha256"], root_path):
                raise RunFailure(f"{name} overlay bytes differ for {relative}")
        checked.append(relative)
    return {"status": baseline["overlay_shape"]["status"], "paths": sorted(checked)}


def _validate_build_receipt(
    manifest: Mapping[str, Any],
    manifest_path: Path,
    baseline: Mapping[str, Any],
    candidate: Mapping[str, Any],
    probe_sources: Sequence[Mapping[str, Any]],
    *,
    receipt_manifest_path: Path | None = None,
) -> dict[str, Any]:
    build = _object(manifest.get("build"), "manifest.build")
    preparation_spec = _identity_spec(build.get("preparation"), "manifest.build.preparation", plan=False)
    preparation_path = _safe_path(preparation_spec["path"], manifest_path.parent, must_exist=True)
    if _overlap(preparation_path, baseline["root"]) or _overlap(preparation_path, candidate["root"]):
        raise RunFailure("build preparation receipt must be outside both source roots")
    preparation_identity = _identity(preparation_path)
    if preparation_identity != {
        "path": str(preparation_path),
        "bytes": preparation_spec["bytes"],
        "sha256": preparation_spec["sha256"],
    }:
        raise RunFailure("build preparation receipt identity differs from the manifest")
    receipt_spec = _identity_spec(build.get("receipt"), "manifest.build.receipt", plan=False)
    receipt_path = _safe_path(receipt_spec["path"], manifest_path.parent, must_exist=True)
    if _overlap(receipt_path, baseline["root"]) or _overlap(receipt_path, candidate["root"]):
        raise RunFailure("build receipt must be outside both source roots")
    receipt_identity = _identity(receipt_path)
    if receipt_identity != {
        "path": str(receipt_path),
        "bytes": receipt_spec["bytes"],
        "sha256": receipt_spec["sha256"],
    }:
        raise RunFailure("build receipt identity differs from the manifest")
    value = _strict_json(receipt_path.read_bytes(), "build receipt")
    receipt = _object(value, "build receipt")
    expected_keys = {
        "schema",
        "status",
        "manifest",
        "manifest_sha256",
        "preparation",
        "source_archive",
        "toolchain",
        "cargo_config",
        "baseline",
        "candidate",
        "fixtures",
        "fixture_routes",
        "probe_sources",
        "baseline_adapter",
        "commands",
        "environment",
        "profile",
        "target_subdir",
        "features",
        "targets",
        "binaries",
        "completed_unix_ns",
    }
    _exact_keys(receipt, expected_keys, "build receipt")
    if receipt.get("schema") != "litchi.ppt.lazy-native.build.v1" or receipt.get("status") != "passed":
        raise RunFailure("build receipt is not a passed PPT lazy-native build receipt")
    receipt_manifest_path = (receipt_manifest_path or manifest_path).resolve()
    receipt_manifest_identity = _identity(receipt_manifest_path)
    if receipt.get("manifest") != str(receipt_manifest_path) or receipt.get("manifest_sha256") != receipt_manifest_identity["sha256"]:
        raise RunFailure("build receipt is bound to a different manifest")
    if receipt["preparation"] != preparation_identity:
        raise RunFailure("build receipt preparation identity differs from the manifest")
    preparation = _object(_strict_json(preparation_path.read_bytes(), "PPT preparation receipt"), "PPT preparation receipt")
    _exact_keys(
        preparation,
        {"schema", "status", "lane", "baseline", "candidate", "claim_scope"},
        "PPT preparation receipt",
    )
    if preparation["schema"] != "litchi.ppt.lazy-native-qualification.preparation.v1" or preparation["status"] != "prepared" or preparation["lane"] != "ppt-lazy-native":
        raise RunFailure("PPT preparation receipt is not the reviewed prepared lane")
    prepared_baseline = _object(preparation["baseline"], "PPT preparation receipt.baseline")
    prepared_candidate = _object(preparation["candidate"], "PPT preparation receipt.candidate")
    if prepared_baseline.get("ref") != baseline["head"] or prepared_baseline.get("tree") != baseline["tree"]:
        raise RunFailure("PPT preparation receipt baseline source identity differs")
    if prepared_candidate.get("ref") != candidate["head"] or prepared_candidate.get("tree") != candidate["tree"]:
        raise RunFailure("PPT preparation receipt candidate source identity differs")
    prepared_paths = prepared_baseline.get("overlay_paths")
    expected_overlay_paths = _baseline_overlay_paths(manifest["baseline_adapter"])
    if prepared_paths != expected_overlay_paths:
        raise RunFailure("PPT preparation receipt overlay paths differ from the reviewed closure")
    prepared_fixture = _object(prepared_baseline.get("fixture"), "PPT preparation receipt.baseline.fixture")
    fixture_source = _object(prepared_fixture.get("source"), "PPT preparation receipt.baseline.fixture.source")
    fixture_destination = _object(prepared_fixture.get("destination"), "PPT preparation receipt.baseline.fixture.destination")
    expected_fixture = manifest["baseline_adapter"]["fixture"]
    expected_fixture_path = candidate["root"] / expected_fixture["path"]
    if fixture_source != {
        "path": str(expected_fixture_path),
        "bytes": expected_fixture["bytes"],
        "sha256": expected_fixture["sha256"],
    }:
        raise RunFailure("PPT preparation receipt fixture source differs from the reviewed fixture")
    if fixture_destination != {
        "path": str(baseline["root"] / expected_fixture["path"]),
        "bytes": expected_fixture["bytes"],
        "sha256": expected_fixture["sha256"],
    }:
        raise RunFailure("PPT preparation receipt fixture destination differs from the baseline endpoint")
    prepared_candidate_fixture = _object(prepared_candidate.get("fixture"), "PPT preparation receipt.candidate.fixture")
    if prepared_candidate_fixture != {
        "path": str(candidate["root"] / expected_fixture["path"]),
        "bytes": expected_fixture["bytes"],
        "sha256": expected_fixture["sha256"],
    }:
        raise RunFailure("PPT preparation receipt candidate fixture differs from the reviewed fixture")
    prepared_compile_fixtures = prepared_baseline.get("compile_fixtures")
    prepared_candidate_compile_fixtures = prepared_candidate.get("compile_fixtures")
    if not isinstance(prepared_compile_fixtures, list) or not isinstance(prepared_candidate_compile_fixtures, list):
        raise RunFailure("PPT preparation receipt compile fixture closure is missing")
    expected_compile_fixtures = manifest["baseline_adapter"]["compile_fixtures"]
    if len(prepared_compile_fixtures) != len(expected_compile_fixtures) or len(prepared_candidate_compile_fixtures) != len(expected_compile_fixtures):
        raise RunFailure("PPT preparation receipt compile fixture closure length differs")
    for index, expected_compile_fixture in enumerate(expected_compile_fixtures):
        prepared_compile = _object(prepared_compile_fixtures[index], f"PPT preparation receipt.baseline.compile_fixtures[{index}]")
        _exact_keys(prepared_compile, {"source", "destination"}, f"PPT preparation receipt.baseline.compile_fixtures[{index}]")
        expected_source = {
            "path": str(candidate["root"] / expected_compile_fixture["path"]),
            "bytes": expected_compile_fixture["bytes"],
            "sha256": expected_compile_fixture["sha256"],
        }
        expected_destination = {
            "path": str(baseline["root"] / expected_compile_fixture["path"]),
            "bytes": expected_compile_fixture["bytes"],
            "sha256": expected_compile_fixture["sha256"],
        }
        if prepared_compile.get("source") != expected_source or prepared_compile.get("destination") != expected_destination:
            raise RunFailure(f"PPT preparation receipt compile fixture {index} differs from the reviewed closure")
        if prepared_candidate_compile_fixtures[index] != expected_source:
            raise RunFailure(f"PPT preparation receipt candidate compile fixture {index} differs")

    roots = (baseline["root"], candidate["root"])
    external_inputs: dict[str, dict[str, Any]] = {}
    for name in ("source_archive", "toolchain", "cargo_config"):
        spec = _identity_spec(build.get(name), f"manifest.build.{name}", plan=False)
        path = _safe_path(spec["path"], manifest_path.parent, must_exist=True)
        if any(_overlap(path, root) for root in roots):
            raise RunFailure(f"manifest.build.{name} must be outside both source roots")
        actual = _identity(path)
        expected = {"path": str(path), "bytes": spec["bytes"], "sha256": spec["sha256"]}
        if actual != expected:
            raise RunFailure(f"manifest.build.{name} identity differs from the manifest")
        if receipt[name] != actual:
            raise RunFailure(f"build receipt {name} identity differs from the manifest")
        external_inputs[name] = actual

    fixture_identities: dict[str, dict[str, Any]] = {}
    for key in FIXTURE_KEYS:
        fixture_spec = _identity_spec(
            build["fixtures"].get(key), f"manifest.build.fixtures.{key}", plan=False
        )
        fixture_path = _safe_path(fixture_spec["path"], manifest_path.parent, must_exist=True)
        if any(_overlap(fixture_path, root) for root in roots):
            raise RunFailure(f"manifest.build.fixtures.{key} must be outside both source roots")
        fixture_identity = _identity(fixture_path)
        expected_fixture = {
            "path": str(fixture_path),
            "bytes": fixture_spec["bytes"],
            "sha256": fixture_spec["sha256"],
        }
        if fixture_identity != expected_fixture:
            raise RunFailure(f"qualification fixture {key} identity differs from the manifest")
        fixture_identities[key] = fixture_identity
    if receipt["fixtures"] != fixture_identities:
        raise RunFailure("build receipt fixture identities differ from the manifest")
    if receipt["fixture_routes"] != build["fixture_routes"]:
        raise RunFailure("build receipt fixture route bindings differ from the manifest")

    for name, endpoint in (("baseline", baseline), ("candidate", candidate)):
        expected_endpoint = {"ref": endpoint["head"], "tree": endpoint["tree"], "lock": endpoint["lock"]}
        if receipt[name] != expected_endpoint:
            raise RunFailure(f"build receipt {name} source/lock identity differs")
    expected_probe = [
        {"path": Path(spec["path"]).as_posix(), "bytes": spec["bytes"], "sha256": spec["sha256"]}
        for spec in manifest["probe_sources"]
    ]
    actual_probe = [
        {
            "path": str(_safe_path(spec["path"], candidate["root"])),
            "bytes": spec["bytes"],
            "sha256": spec["sha256"],
        }
        for spec in manifest["probe_sources"]
    ]
    if list(probe_sources) != actual_probe:
        raise RunFailure("candidate probe source checks were not bound to the receipt")
    if receipt["probe_sources"] != expected_probe:
        raise RunFailure("build receipt probe source identities differ from the manifest")
    adapter = _object(receipt["baseline_adapter"], "build receipt baseline_adapter")
    _exact_keys(
        adapter,
        {"overlay_paths", "registration", "fixture", "compile_fixtures"},
        "build receipt baseline_adapter",
    )
    if adapter["overlay_paths"] != expected_probe[:3]:
        raise RunFailure("build receipt baseline overlay identities differ from the reviewed probe")
    adapter_fixture = _object(adapter["fixture"], "build receipt baseline_adapter.fixture")
    if adapter_fixture != manifest["baseline_adapter"]["fixture"]:
        raise RunFailure("build receipt baseline fixture identity differs from the reviewed fixture")
    if adapter["compile_fixtures"] != manifest["baseline_adapter"]["compile_fixtures"]:
        raise RunFailure("build receipt baseline compile fixture closure differs from the reviewed closure")
    registration = _object(adapter["registration"], "build receipt baseline_adapter.registration")
    _exact_keys(registration, {"path", "line", "before", "after"}, "build receipt baseline_adapter.registration")
    if registration["path"] != manifest["baseline_adapter"]["registration"]["path"] or registration["line"] != manifest["baseline_adapter"]["registration"]["line"]:
        raise RunFailure("build receipt adapter registration differs from the reviewed diff")
    before_path = _safe_path(registration["before"]["path"], baseline["root"], must_exist=True)
    after_path = _safe_path(registration["after"]["path"], candidate["root"], must_exist=True)
    if before_path != baseline["root"] / registration["path"] or after_path != candidate["root"] / registration["path"]:
        raise RunFailure("build receipt adapter registration paths differ from endpoint roots")
    before_identity = _relative_identity(before_path, baseline["root"], "adapter registration before")
    after_identity = _relative_identity(after_path, candidate["root"], "adapter registration after")
    if registration["before"] != before_identity or registration["after"] != after_identity:
        raise RunFailure("build receipt adapter registration bytes changed")
    if after_identity != expected_probe[3]:
        raise RunFailure("build receipt adapter registration does not bind candidate src/lib.rs")

    if receipt["commands"] != {
        "normal": build["normal_command"],
        "allocator": build["allocator_command"],
    } or receipt["environment"] != build["environment"] or receipt["profile"] != build["profile"] or receipt["target_subdir"] != build["target_subdir"] or receipt["features"] != build["features"]:
        raise RunFailure("build receipt command/configuration identity differs from the manifest")
    if _positive_int(receipt.get("completed_unix_ns"), "build receipt.completed_unix_ns") <= 0:
        raise RunFailure("build receipt completion timestamp is invalid")

    expected_targets = {
        name: {
            "normal": str(endpoint["target_dir"]),
            "allocator": str(endpoint["allocator_target_dir"]),
        }
        for name, endpoint in (("baseline", baseline), ("candidate", candidate))
    }
    if receipt["targets"] != expected_targets:
        raise RunFailure("build receipt target paths differ from the endpoints")
    expected_binaries = {
        name: {
            "normal": endpoint["normal_binary"],
            "allocator": endpoint["allocator_binary"],
        }
        for name, endpoint in (("baseline", baseline), ("candidate", candidate))
    }
    if receipt["binaries"] != expected_binaries:
        raise RunFailure("build receipt binary identities differ from the endpoints")
    overlay = _validate_baseline_overlay_content(manifest, receipt, baseline, candidate)
    return {
        "path": str(receipt_path),
        "bytes": receipt_identity["bytes"],
        "sha256": receipt_identity["sha256"],
        "source_archive": external_inputs["source_archive"],
        "preparation": preparation_identity,
        "toolchain": external_inputs["toolchain"],
        "cargo_config": external_inputs["cargo_config"],
        "fixtures": fixture_identities,
        "baseline_overlay": overlay,
        "baseline_adapter": receipt["baseline_adapter"],
        "manifest_binding": {
            "path": str(receipt_manifest_path),
            "bytes": receipt_manifest_identity["bytes"],
            "sha256": receipt_manifest_identity["sha256"],
        },
    }


def _check_target(path: Path, label: str) -> dict[str, Any]:
    if not path.is_dir():
        raise RunFailure(f"{label} is no longer a directory")
    size = _directory_bytes(path)
    if size > TARGET_PREFLIGHT_BYTES:
        raise RunFailure(f"{label} exceeds the fixed 19.2 GB preflight ceiling")
    return {
        "path": str(path),
        "bytes": size,
        "preflight_limit_bytes": TARGET_PREFLIGHT_BYTES,
        "hard_limit_bytes": TARGET_HARD_LIMIT_BYTES,
    }


def _parse_child(path: Path, label: str) -> dict[str, Any]:
    try:
        raw = path.read_bytes()
    except OSError as error:
        raise RunFailure(f"{label}: cannot read child output: {error}") from error
    if len(raw) > MAX_CHILD_OUTPUT:
        raise RunFailure(f"{label}: child output exceeds {MAX_CHILD_OUTPUT} bytes")
    value = _strict_json(raw, label)
    return _object(value, label)


def _validate_quiet_guard(value: Any, label: str, max_load: float) -> dict[str, Any]:
    guard = _object(value, label)
    _exact_keys(guard, {"phase", "load_1m", "max_load", "heavy_processes", "checked_unix_ns"}, label)
    if guard["max_load"] != max_load:
        raise RunFailure(f"{label}.max_load differs from the manifest")
    load = guard["load_1m"]
    _finite_number(load, f"{label}.load_1m", minimum=0)
    if load > max_load:
        raise RunFailure(f"{label}.load_1m exceeds the quiet-host ceiling")
    if guard["heavy_processes"] != []:
        raise RunFailure(f"{label}.heavy_processes is not empty")
    _positive_int(guard["checked_unix_ns"], f"{label}.checked_unix_ns")
    return guard


def _validate_child_event(
    event: Mapping[str, Any],
    *,
    endpoint: Mapping[str, Any],
    out_dir: Path,
    label: str,
    binary: Path,
    target_dir: Path,
    args: Sequence[str],
    protocol: Mapping[str, Any],
) -> dict[str, Any]:
    _exact_keys(
        event,
        {"label", "command", "cwd", "returncode", "wall_ns", "target", "stdout", "stderr", "monitor"},
        f"{label}.event",
    )
    if (event["label"] != label
            or event["command"] != [str(binary), *args]
            or event["cwd"] != str(endpoint["root"])
            or event["returncode"] != 0):
        raise RunFailure(f"{label}: child event identity differs")
    _positive_int(event["wall_ns"], f"{label}.event.wall_ns")

    target = _object(event["target"], f"{label}.event.target")
    _exact_keys(target, {"path", "bytes_before", "bytes_after", "hard_limit_bytes"}, f"{label}.event.target")
    if (target["path"] != str(target_dir)
            or type(target["bytes_before"]) is not int
            or type(target["bytes_after"]) is not int
            or target["bytes_before"] < 0
            or target["bytes_after"] < 0
            or target["bytes_before"] > TARGET_PREFLIGHT_BYTES
            or target["bytes_after"] > TARGET_HARD_LIMIT_BYTES
            or target["hard_limit_bytes"] != TARGET_HARD_LIMIT_BYTES):
        raise RunFailure(f"{label}: target accounting differs from the fixed ceilings")

    expected_stdout = out_dir / f"{label}.stdout"
    expected_stderr = out_dir / f"{label}.stderr"
    for stream, expected in (("stdout", expected_stdout), ("stderr", expected_stderr)):
        identity = _object(event[stream], f"{label}.event.{stream}")
        _exact_keys(identity, {"path", "bytes", "sha256"}, f"{label}.event.{stream}")
        if identity["path"] != str(expected):
            raise RunFailure(f"{label}: {stream} log path differs from the output directory")
        _nonnegative_int(identity["bytes"], f"{label}.event.{stream}.bytes")
        _digest_value(identity["sha256"], f"{label}.event.{stream}.sha256")
        if _identity(expected) != identity:
            raise RunFailure(f"{label}: {stream} log identity changed after the child exited")

    monitor = _object(event["monitor"], f"{label}.event.monitor")
    _exact_keys(
        monitor,
        {"samples", "memory_limit_bytes", "max_output_bytes", "max_load", "owned_processes"},
        f"{label}.event.monitor",
    )
    if (monitor["memory_limit_bytes"] != MEMORY_HARD_LIMIT_BYTES
            or monitor["max_output_bytes"] != MAX_CHILD_OUTPUT
            or monitor["max_load"] != protocol["max_load"]
            or monitor["owned_processes"] != []):
        raise RunFailure(f"{label}: child monitor limits or ownership differ")
    samples = monitor["samples"]
    if not isinstance(samples, list) or not samples:
        raise RunFailure(f"{label}: child monitor samples are empty")
    for index, sample_value in enumerate(samples):
        sample = _object(sample_value, f"{label}.event.monitor.samples[{index}]")
        _exact_keys(
            sample,
            {"checked_unix_ns", "output_bytes", "rss_bytes", "load_1m", "outside_processes", "owned_processes"},
            f"{label}.event.monitor.samples[{index}]",
        )
        _positive_int(sample["checked_unix_ns"], f"{label}.event.monitor.samples[{index}].checked_unix_ns")
        _nonnegative_int(sample["output_bytes"], f"{label}.event.monitor.samples[{index}].output_bytes")
        if sample["output_bytes"] > MAX_CHILD_OUTPUT:
            raise RunFailure(f"{label}: child output exceeded the fixed cap")
        _nonnegative_int(sample["rss_bytes"], f"{label}.event.monitor.samples[{index}].rss_bytes")
        if sample["rss_bytes"] > MEMORY_HARD_LIMIT_BYTES:
            raise RunFailure(f"{label}: child RSS exceeded the fixed cap")
        load = sample["load_1m"]
        _finite_number(load, f"{label}.event.monitor.samples[{index}].load_1m", minimum=0)
        if load > protocol["max_load"]:
            raise RunFailure(f"{label}: child sample load exceeded the quiet-host ceiling")
        if sample["outside_processes"] != []:
            raise RunFailure(f"{label}: child process ownership/host guard was not clean")
        owned = sample["owned_processes"]
        if not isinstance(owned, list):
            raise RunFailure(f"{label}: child ownership sample is malformed")
        for owned_index, owned_value in enumerate(owned):
            owned_item = _object(owned_value, f"{label}.event.monitor.samples[{index}].owned_processes[{owned_index}]")
            _exact_keys(owned_item, {"pid", "start_time", "process_group"}, f"{label}.owned_processes[{owned_index}]")
            _positive_int(owned_item["pid"], f"{label}.owned_processes[{owned_index}].pid")
            _positive_int(owned_item["start_time"], f"{label}.owned_processes[{owned_index}].start_time")
            _positive_int(owned_item["process_group"], f"{label}.owned_processes[{owned_index}].process_group")
    return dict(event)


def _run_child(
    endpoint: Mapping[str, Any],
    protocol: Mapping[str, Any],
    out_dir: Path,
    label: str,
    binary: Path,
    target_dir: Path,
    args: Sequence[str],
) -> tuple[dict[str, Any], dict[str, Any]]:
    import shutil

    expected_binary = endpoint["normal_binary"]
    if binary == Path(endpoint["allocator_binary"]["path"]):
        expected_binary = endpoint["allocator_binary"]
    if _identity(binary) != expected_binary:
        raise RunFailure(f"{label}: executable identity changed before launch")
    if shutil.which("taskset") is None:
        raise RunFailure("taskset is required for the pinned native lane")
    cpu = _nonnegative_int(protocol["cpu"], "protocol.cpu")
    try:
        affinity = os.sched_getaffinity(0)
    except (AttributeError, OSError) as error:
        raise RunFailure(f"cannot inspect runner CPU affinity: {error}") from error
    if cpu not in affinity:
        raise RunFailure(f"protocol CPU {cpu} is not available in runner affinity {sorted(affinity)}")
    guard_before = _quiet_guard(float(protocol["max_load"]), f"before-{label}")
    environment = dict(os.environ)
    environment["TMPDIR"] = str(out_dir / "tmp")
    event = _stream_command(
        [str(binary), *args],
        endpoint["root"],
        out_dir,
        label,
        target_dir,
        cpu,
        float(protocol["timeout_s"]),
        environment,
        MAX_CHILD_OUTPUT,
        float(protocol["max_load"]),
    )
    guard_after = _quiet_guard(float(protocol["max_load"]), f"after-{label}")
    _validate_quiet_guard(guard_before, f"{label}.guard.before", float(protocol["max_load"]))
    _validate_quiet_guard(guard_after, f"{label}.guard.after", float(protocol["max_load"]))
    _validate_child_event(
        event,
        endpoint=endpoint,
        out_dir=out_dir,
        label=label,
        binary=binary,
        target_dir=target_dir,
        args=args,
        protocol=protocol,
    )
    report = _parse_child(Path(event["stdout"]["path"]), label)
    return report, {
        "label": label,
        "endpoint": endpoint["name"],
        "binary": binary_identity(binary),
        "args": list(args),
        "guard": {
            "before": guard_before,
            "live": event["monitor"]["samples"],
            "after": guard_after,
            "memory_limit_bytes": MEMORY_HARD_LIMIT_BYTES,
            "max_output_bytes": MAX_CHILD_OUTPUT,
            "max_load": protocol["max_load"],
        },
        "target": event["target"],
        "event": event,
    }


def _digest_object(value: Any, label: str) -> dict[str, Any]:
    item = _object(value, label)
    if set(item) != {"bytes", "sha256"}:
        raise RunFailure(f"{label} digest fields differ from the protocol")
    return {"bytes": _positive_int(item.get("bytes"), f"{label}.bytes"), "sha256": _digest_value(item.get("sha256"), f"{label}.sha256")}


def _validate_route_proof(
    proof_value: Any,
    route: str,
    fixture: Mapping[str, Any],
    fixture_key: str,
    label: str,
) -> dict[str, Any]:
    proof = _object(proof_value, label)
    expected_fields = {
        "output",
        "source_exact",
        "changed",
        "patch_noop",
        "operation_count",
        "patch_before_exact",
        "patch_after_exact",
        "entry",
        "fixture",
    }
    if set(proof) != expected_fields:
        raise RunFailure(f"{label}: proof fields differ from the protocol")
    if _digest_object(proof["output"], f"{label}.output") != dict(fixture):
        raise RunFailure(f"{label}: output differs from the reviewed fixture")
    proof_fixture = _object(proof["fixture"], f"{label}.fixture")
    _exact_keys(proof_fixture, {"key", "output"}, f"{label}.fixture")
    if proof_fixture["key"] != fixture_key:
        raise RunFailure(f"{label}: route is bound to the wrong fixture")
    if _digest_object(proof_fixture["output"], f"{label}.fixture.output") != dict(fixture):
        raise RunFailure(f"{label}: route fixture digest differs")
    expected_entry = "commit" if route.endswith("noop_commit") else "drop"
    if proof["source_exact"] is not True or proof["changed"] is not False or proof["entry"] != expected_entry:
        raise RunFailure(f"{label}: exact no-op/drop semantics failed")
    if expected_entry == "commit" and proof["patch_noop"] is not True:
        raise RunFailure(f"{label}: commit did not report an exact no-op")
    if expected_entry == "drop" and proof["patch_noop"] is not None:
        raise RunFailure(f"{label}: drop route unexpectedly has a patch")
    for name in ("operation_count", "patch_before_exact", "patch_after_exact"):
        if proof[name] is not None and name == "operation_count" and proof[name] != 0:
            raise RunFailure(f"{label}: operation count is not zero")
        if proof[name] is not None and name != "operation_count" and proof[name] is not True:
            raise RunFailure(f"{label}: exact patch endpoint proof failed")
    return proof


def _validate_verify(
    report: Mapping[str, Any],
    label: str,
    expected_fixtures: Mapping[str, Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    if set(report) != {"schema", "mode", "fixtures", "routes", "allocation"}:
        raise RunFailure(f"{label}: verify report fields differ from the protocol")
    if report.get("schema") != PROBE_SCHEMA or report.get("mode") != "verify":
        raise RunFailure(f"{label}: verify report identity is invalid")
    fixtures = _object(report.get("fixtures"), f"{label}.fixtures")
    if set(fixtures) != set(FIXTURE_KEYS):
        raise RunFailure(f"{label}: fixture set differs from the reviewed matrix")
    sources: dict[str, dict[str, Any]] = {}
    for key in FIXTURE_KEYS:
        fixture = _object(fixtures[key], f"{label}.fixtures.{key}")
        if set(fixture) != {"generator", "output"} or fixture.get("generator") != FIXTURE_GENERATOR:
            raise RunFailure(f"{label}: reviewed fixture {key} identity is invalid")
        sources[key] = _digest_object(fixture.get("output"), f"{label}.fixtures.{key}.output")
        if expected_fixtures is not None:
            expected = expected_fixtures[key]
            expected_digest = {
                "bytes": _positive_int(expected.get("bytes"), f"manifest.build.fixtures.{key}.bytes"),
                "sha256": _digest_value(expected.get("sha256"), f"manifest.build.fixtures.{key}.sha256"),
            }
            if sources[key] != expected_digest:
                raise RunFailure(f"{label}: fixture {key} output differs from its manifest source")
    if report.get("allocation") != {"status": "unavailable", "scope": "operation_global_system_allocator"}:
        raise RunFailure(f"{label}: normal verify probe must report allocator metrics as unavailable")
    routes = _object(report.get("routes"), f"{label}.routes")
    if set(routes) != set(ROUTES):
        raise RunFailure(f"{label}: route set differs from {ROUTES!r}")
    for route in ROUTES:
        fixture_key = ROUTE_FIXTURE_KEYS[route]
        _validate_route_proof(routes[route], route, sources[fixture_key], fixture_key, f"{label}.{route}")
    return {"fixtures": sources, "routes": routes}


def _validate_calibration(
    report: Mapping[str, Any],
    route: str,
    protocol: Mapping[str, Any],
    fixture: Mapping[str, Any],
    fixture_key: str,
    label: str,
) -> int:
    if set(report) != {"schema", "mode", "route", "target_ns", "iterations", "elapsed_ns", "fixture", "postcondition"}:
        raise RunFailure(f"{label}: calibration fields differ from the protocol")
    if report.get("schema") != PROBE_SCHEMA or report.get("mode") != "calibrate" or report.get("route") != route:
        raise RunFailure(f"{label}: calibration identity differs")
    iterations = _positive_int(report.get("iterations"), f"{label}.iterations")
    max_iterations = _positive_int(protocol["max_iterations"], "protocol.max_iterations")
    if iterations > max_iterations or report.get("target_ns") != round(float(protocol["target_ms"]) * 1_000_000):
        raise RunFailure(f"{label}: calibration target or bound differs")
    if _positive_int(report.get("elapsed_ns"), f"{label}.elapsed_ns") < report["target_ns"]:
        raise RunFailure(f"{label}: calibration did not reach its target")
    report_fixture = _object(report.get("fixture"), f"{label}.fixture")
    _exact_keys(report_fixture, {"key", "output"}, f"{label}.fixture")
    if report_fixture["key"] != fixture_key or _digest_object(report_fixture["output"], f"{label}.fixture.output") != dict(fixture):
        raise RunFailure(f"{label}: calibration fixture differs from correctness")
    _validate_route_proof(report.get("postcondition"), route, fixture, fixture_key, f"{label}.postcondition")
    return iterations


def _validate_timed(
    report: Mapping[str, Any],
    route: str,
    iterations: int,
    allocator: bool,
    protocol: Mapping[str, Any],
    fixture: Mapping[str, Any],
    fixture_key: str,
    label: str,
) -> dict[str, Any]:
    if set(report) != {"schema", "mode", "route", "samples", "fixture", "postcondition"}:
        raise RunFailure(f"{label}: timed fields differ from the protocol")
    if report.get("schema") != PROBE_SCHEMA or report.get("mode") != "timed" or report.get("route") != route:
        raise RunFailure(f"{label}: timed identity differs")
    report_fixture = _object(report.get("fixture"), f"{label}.fixture")
    _exact_keys(report_fixture, {"key", "output"}, f"{label}.fixture")
    if report_fixture["key"] != fixture_key or _digest_object(report_fixture["output"], f"{label}.fixture.output") != dict(fixture):
        raise RunFailure(f"{label}: timed fixture differs from correctness")
    _validate_route_proof(report.get("postcondition"), route, fixture, fixture_key, f"{label}.postcondition")
    samples = report.get("samples")
    if not isinstance(samples, list) or not samples:
        raise RunFailure(f"{label}: timed samples are empty")
    rows = []
    for index, value in enumerate(samples):
        row = _object(value, f"{label}.samples[{index}]")
        if set(row) != {"elapsed_ns", "iterations", "allocation"}:
            raise RunFailure(f"{label}.samples[{index}]: fields differ")
        elapsed_ns = _positive_int(row.get("elapsed_ns"), f"{label}.samples[{index}].elapsed_ns")
        minimum_ns = round(float(protocol["minimum_ms"]) * 1_000_000)
        if row.get("iterations") != iterations or elapsed_ns < minimum_ns:
            raise RunFailure(f"{label}.samples[{index}]: timing identity is invalid")
        allocation = _object(row.get("allocation"), f"{label}.samples[{index}].allocation")
        if allocator:
            required = {"status", "scope", "allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after"}
            if set(allocation) != required or allocation["status"] != "measured" or allocation["scope"] != "operation_global_system_allocator":
                raise RunFailure(f"{label}.samples[{index}]: allocation observation is not measured")
            for metric in required - {"status", "scope"}:
                _nonnegative_int(allocation[metric], f"{label}.samples[{index}].allocation.{metric}")
            if (
                allocation["live_bytes_after"] - allocation["live_bytes_before"]
                != allocation["allocated_bytes"] - allocation["deallocated_bytes"]
                or allocation["peak_live_bytes_before"] < allocation["live_bytes_before"]
                or allocation["peak_live_bytes_after"] < allocation["peak_live_bytes_before"]
                or allocation["peak_live_bytes_after"] < allocation["live_bytes_after"]
            ):
                raise RunFailure(f"{label}.samples[{index}]: allocator peak regressed")
        else:
            if allocation != {"status": "unavailable", "scope": "operation_global_system_allocator"}:
                raise RunFailure(f"{label}.samples[{index}]: normal probe reported allocator metrics")
        rows.append(row)
    return {"samples": rows}


def _recheck(
    endpoint: Mapping[str, Any],
    protocol_sources: Sequence[Mapping[str, Any]],
    manifest_path: Path,
    manifest_identity: Mapping[str, Any],
    phase: str,
    projection_path: Path | None = None,
    projection_identity: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    raw = endpoint["raw"]
    root = endpoint["root"]
    actual_manifest = _identity(manifest_path)
    if actual_manifest != manifest_identity:
        raise RunFailure(f"{phase}: qualification manifest changed")
    if (projection_path is None) != (projection_identity is None):
        raise RunFailure(f"{phase}: build-input projection binding is incomplete")
    if projection_path is not None and _identity(projection_path) != projection_identity:
        raise RunFailure(f"{phase}: build-input projection changed")
    probe_sources = (
        _check_probe_sources(protocol_sources, endpoint)
        if endpoint["name"] == "candidate"
        else []
    )
    if endpoint.get("overlay_adapter") is not None:
        _check_baseline_overlay_shape(root, endpoint["name"], endpoint["overlay_adapter"])
        if endpoint.get("overlay_receipt") is not None and endpoint.get("overlay_candidate") is not None:
            _validate_baseline_overlay_content(
                {"baseline_adapter": endpoint["overlay_adapter"]},
                {"baseline_adapter": endpoint["overlay_receipt"]},
                endpoint,
                endpoint["overlay_candidate"],
            )
        status = endpoint["overlay_shape"]["status"]
    else:
        _check_worktree_shape(root, endpoint["name"])
        status = _git(
            root,
            "status",
            "--porcelain",
            "--untracked-files=all",
            no_optional_locks=True,
        )
    head = _git(root, "rev-parse", "HEAD")
    tree = _git(root, "rev-parse", "HEAD^{tree}")
    if (status and endpoint.get("overlay_adapter") is None) or head != endpoint["head"] or tree != endpoint["tree"]:
        raise RunFailure(f"{phase}: {endpoint['name']} source changed")
    lock = _check_lock(raw["lock"], f"{phase}.{endpoint['name']}.lock", root)
    normal = _check_identity(raw["normal_binary"], f"{phase}.{endpoint['name']}.normal_binary", manifest_path.parent, executable=True)
    allocator = _check_identity(raw["allocator_binary"], f"{phase}.{endpoint['name']}.allocator_binary", manifest_path.parent, executable=True)
    targets = {
        "normal": _check_target(endpoint["target_dir"], f"{phase}.{endpoint['name']}.target_dir"),
        "allocator": _check_target(endpoint["allocator_target_dir"], f"{phase}.{endpoint['name']}.allocator_target_dir"),
    }
    return {
        "phase": phase,
        "head": head,
        "tree": tree,
        "lock": lock,
        "normal_binary": normal,
        "allocator_binary": allocator,
        "targets": targets,
        "probe_sources": probe_sources,
    }


def _prepare_output(path: Path, roots: Sequence[Path]) -> Path:
    path = _safe_path(str(path))
    if any(_overlap(path, root) for root in roots):
        raise RunFailure("output directory must be outside source worktrees")
    if path.exists():
        if not path.is_dir() or any(path.iterdir()):
            raise RunFailure("output directory must be new and empty")
    else:
        path.mkdir(parents=True)
    (path / "tmp").mkdir()
    return path


def _bind_inputs(
    manifest: Mapping[str, Any],
    manifest_path: Path,
    projection_path: Path | None,
) -> tuple[dict[str, Any], dict[str, Any], list[dict[str, Any]], dict[str, Any], dict[str, Any] | None]:
    candidate_raw = _object(manifest.get("candidate"), "manifest.candidate")
    projection_identity = _validate_build_input_projection(manifest, manifest_path, projection_path)
    roots = [
        _safe_path(_string(manifest["baseline"]["root"], "baseline.root"), manifest_path.parent),
        _safe_path(_string(candidate_raw["root"], "candidate.root"), manifest_path.parent),
    ]
    if roots[0] == roots[1] or _overlap(roots[0], roots[1]):
        raise RunFailure("baseline and candidate source roots must be distinct and non-nested")
    baseline = _resolve_endpoint(
        manifest["baseline"],
        "baseline",
        manifest_path,
        roots,
        baseline_adapter=manifest["baseline_adapter"],
    )
    candidate = _resolve_endpoint(candidate_raw, "candidate", manifest_path, roots)
    if baseline["head"] == candidate["head"] or baseline["tree"] == candidate["tree"]:
        raise RunFailure("baseline and candidate source identities must differ")
    probe_sources = _check_probe_sources(manifest["probe_sources"], candidate)
    build_receipt = _validate_build_receipt(
        manifest,
        manifest_path,
        baseline,
        candidate,
        probe_sources,
        receipt_manifest_path=projection_path,
    )
    baseline["overlay_receipt"] = build_receipt["baseline_adapter"]
    baseline["overlay_candidate"] = candidate
    return manifest, {"baseline": baseline, "candidate": candidate}, probe_sources, build_receipt, projection_identity


def validate_inputs(manifest_path: Path, projection_path: Path | None = None) -> dict[str, Any]:
    """Run the public source/receipt/endpoint lifecycle without a native child."""

    manifest_path = manifest_path.resolve()
    manifest = _load_manifest(manifest_path, plan=False)
    _, endpoints, probe_sources, build_receipt, projection_identity = _bind_inputs(
        manifest, manifest_path, projection_path.resolve() if projection_path is not None else None
    )
    return {
        "manifest": _identity(manifest_path),
        "build_input_projection": projection_identity,
        "baseline": endpoints["baseline"],
        "candidate": endpoints["candidate"],
        "probe_sources": probe_sources,
        "build_receipt": build_receipt,
    }


def _plan(manifest: Mapping[str, Any], path: Path, out_dir: Path) -> None:
    plan = {
        "schema": REPORT_SCHEMA,
        "status": "plan_only",
        "claim_scope": "preparation only; no source, binary, allocation, or timing execution",
        "manifest": str(path.resolve()),
        "baseline": {"ref": manifest["baseline"]["ref"], "tree": manifest["baseline"]["tree"], "lock": manifest["baseline"]["lock"]},
        "candidate": manifest["candidate"],
        "preparation": manifest["build"]["preparation"],
        "routes": list(ROUTES),
        "fixtures": manifest["build"]["fixtures"],
        "fixture_routes": manifest["build"]["fixture_routes"],
        "abba_order": list(ABBA_ORDER),
        "allocation": {"binary": "ppt_lazy_native_alloc", "scope": "operation_global_system_allocator"},
        "limits": {"target_preflight_bytes": TARGET_PREFLIGHT_BYTES, "target_hard_limit_bytes": TARGET_HARD_LIMIT_BYTES, "memory_hard_limit_bytes": MEMORY_HARD_LIMIT_BYTES, "max_child_output": MAX_CHILD_OUTPUT},
        "protocol": manifest["protocol"],
    }
    (out_dir / "plan.json").write_text(json.dumps(plan, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def _run(
    manifest: Mapping[str, Any],
    manifest_path: Path,
    out_dir: Path,
    correctness_only: bool,
    projection_path: Path | None,
) -> dict[str, Any]:
    manifest_identity = _identity(manifest_path)
    _, endpoints, probe_sources, build_receipt, projection_identity = _bind_inputs(
        manifest, manifest_path, projection_path
    )
    baseline = endpoints["baseline"]
    candidate = endpoints["candidate"]
    roots = [baseline["root"], candidate["root"]]
    protocol = manifest["protocol"]
    endpoints = {"baseline": baseline, "candidate": candidate}
    resource_paths = [out_dir]
    for endpoint in endpoints.values():
        resource_paths.extend((endpoint["target_dir"], endpoint["allocator_target_dir"]))
    for resource in resource_paths:
        if any(_overlap(resource, root) for root in roots):
            raise RunFailure(f"qualification output or target overlaps a source root: {resource}")
    for index, first in enumerate(resource_paths):
        for second in resource_paths[index + 1 :]:
            if _under(first, second) or _under(second, first):
                raise RunFailure(f"qualification output and target directories overlap: {first} / {second}")
    identity = {
        "manifest": manifest_identity,
        "before_correctness": {
            name: _recheck(
                endpoint,
                manifest["probe_sources"],
                manifest_path,
                manifest_identity,
                "before_correctness",
                projection_path,
                projection_identity,
            )
            for name, endpoint in endpoints.items()
        },
    }
    correctness: dict[str, Any] = {}
    for name, endpoint in endpoints.items():
        report, event = _run_child(
            endpoint,
            protocol,
            out_dir,
            f"correctness-{name}",
            Path(endpoint["normal_binary"]["path"]),
            endpoint["target_dir"],
            ("--mode", "verify"),
        )
        correctness[name] = {
            "report": _validate_verify(report, f"correctness.{name}", manifest["build"]["fixtures"]),
            "event": event,
        }
    if correctness["baseline"]["report"] != correctness["candidate"]["report"]:
        raise RunFailure("baseline and candidate correctness proofs differ")
    identity["after_correctness"] = {
        name: _recheck(
            endpoint,
            manifest["probe_sources"],
            manifest_path,
            manifest_identity,
            "after_correctness",
            projection_path,
            projection_identity,
        )
        for name, endpoint in endpoints.items()
    }
    if correctness_only:
        return {
            "schema": REPORT_SCHEMA,
            "status": "correctness_only",
            "claim_scope": "untimed correctness preflight only; no native timing or allocation claim",
            "manifest": str(manifest_path.resolve()),
            "build_input_projection": projection_identity,
            "routes": list(ROUTES),
            "abba_order": list(ABBA_ORDER),
            "probe_sources": probe_sources,
            "correctness": correctness,
            "build_receipt": build_receipt,
            "identity": identity,
        }
    calibrations: dict[str, dict[str, int]] = {"baseline": {}, "candidate": {}}
    calibration_events = []
    target_ns = round(float(protocol["target_ms"]) * 1_000_000)
    for name, endpoint in endpoints.items():
        for route in ROUTES:
            report, event = _run_child(
                endpoint,
                protocol,
                out_dir,
                f"calibrate-{name}-{route}",
                Path(endpoint["normal_binary"]["path"]),
                endpoint["target_dir"],
                ("--mode", "calibrate", "--route", route, "--target-ns", str(target_ns), "--max-iterations", str(protocol["max_iterations"])),
            )
            calibrations[name][route] = _validate_calibration(
                report,
                route,
                protocol,
                correctness["baseline"]["report"]["fixtures"][ROUTE_FIXTURE_KEYS[route]],
                ROUTE_FIXTURE_KEYS[route],
                f"calibration.{name}.{route}",
            )
            calibration_events.append(event)
    identity["after_calibration"] = {
        name: _recheck(
            endpoint,
            manifest["probe_sources"],
            manifest_path,
            manifest_identity,
            "after_calibration",
            projection_path,
            projection_identity,
        )
        for name, endpoint in endpoints.items()
    }
    routes: dict[str, Any] = {}
    for route in ROUTES:
        blocks = []
        for block in range(protocol["blocks"]):
            normal_rows = []
            allocation_rows = []
            for role in ABBA_ORDER:
                name = ROLE_ENDPOINT[role]
                endpoint = endpoints[name]
                iterations = calibrations[name][route]
                normal_report, normal_event = _run_child(
                    endpoint,
                    protocol,
                    out_dir,
                    f"normal-{route}-block-{block}-{role}",
                    Path(endpoint["normal_binary"]["path"]),
                    endpoint["target_dir"],
                    ("--mode", "timed", "--route", route, "--iterations", str(iterations), "--samples", "1", "--warmups", str(protocol["warmups"]), "--max-iterations", str(protocol["max_iterations"])),
                )
                normal_rows.append({"role": role, "endpoint": name, "report": _validate_timed(normal_report, route, iterations, False, protocol, correctness["baseline"]["report"]["fixtures"][ROUTE_FIXTURE_KEYS[route]], ROUTE_FIXTURE_KEYS[route], f"timed.{route}.{role}.{block}"), "event": normal_event})
                allocation_report, allocation_event = _run_child(
                    endpoint,
                    protocol,
                    out_dir,
                    f"allocation-{route}-block-{block}-{role}",
                    Path(endpoint["allocator_binary"]["path"]),
                    endpoint["allocator_target_dir"],
                    ("--mode", "timed", "--route", route, "--iterations", str(iterations), "--samples", "1", "--warmups", str(protocol["warmups"]), "--max-iterations", str(protocol["max_iterations"])),
                )
                allocation_rows.append({"role": role, "endpoint": name, "report": _validate_timed(allocation_report, route, iterations, True, protocol, correctness["baseline"]["report"]["fixtures"][ROUTE_FIXTURE_KEYS[route]], ROUTE_FIXTURE_KEYS[route], f"allocation.{route}.{role}.{block}"), "event": allocation_event})
            blocks.append({"block": block, "abba_order": list(ABBA_ORDER), "normal": normal_rows, "allocation": allocation_rows})
        routes[route] = {"blocks": blocks}
    identity["after_abba"] = {
        name: _recheck(
            endpoint,
            manifest["probe_sources"],
            manifest_path,
            manifest_identity,
            "after_abba",
            projection_path,
            projection_identity,
        )
        for name, endpoint in endpoints.items()
    }
    return {
        "schema": REPORT_SCHEMA,
        "status": "complete",
        "claim_scope": "native elapsed and operation-scoped allocator observations after exact correctness preflight",
        "manifest": str(manifest_path.resolve()),
        "build_input_projection": projection_identity,
        "routes": list(ROUTES),
        "abba_order": list(ABBA_ORDER),
        "probe_sources": probe_sources,
        "correctness": correctness,
        "build_receipt": build_receipt,
        "calibration": calibrations,
        "calibration_events": calibration_events,
        "routes_evidence": routes,
        "identity": identity,
    }


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument(
        "--build-input-projection",
        type=Path,
        help="canonical manifest projection with only build.receipt bytes/sha256 removed",
    )
    parser.add_argument("--out-dir", type=Path, required=True)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--plan", action="store_true")
    mode.add_argument("--correctness-only", action="store_true")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        manifest_path = args.manifest.resolve()
        manifest = _load_manifest(manifest_path, plan=args.plan)
        if args.plan:
            if args.build_input_projection is not None:
                _validate_build_input_projection(manifest, manifest_path, args.build_input_projection.resolve())
            out_dir = _prepare_output(args.out_dir, ())
            _plan(manifest, manifest_path, out_dir)
            return 0
        candidate = _object(manifest.get("candidate"), "manifest.candidate")
        roots = [
            _safe_path(_string(manifest["baseline"]["root"], "baseline.root"), manifest_path.parent),
            _safe_path(_string(candidate["root"], "candidate.root"), manifest_path.parent),
        ]
        if roots[0] == roots[1] or _overlap(roots[0], roots[1]):
            raise RunFailure("baseline and candidate source roots must be distinct and non-nested")
        out_dir = _prepare_output(args.out_dir, roots)
        report = _run(
            manifest,
            manifest_path,
            out_dir,
            args.correctness_only,
            args.build_input_projection.resolve() if args.build_input_projection is not None else None,
        )
        (out_dir / "report.json").write_text(json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        return 0
    except RunFailure as error:
        print(f"PPT lazy native qualification failed: {error}", file=sys.stderr)
        return 1
    except Exception as error:  # keep the CLI fail-closed and content-light
        print(f"PPT lazy native qualification failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
