#!/usr/bin/env python3
"""Independently validate one completed W5 OLE owner qualification run.

The paired runner is an orchestrator, not the authority that closes a native
qualification.  This validator deliberately re-reads every receipt, journal,
log, executable, oracle artifact, and measurement from disk and repeats the
identity/conservation checks without calling runner parsing helpers.  It is
safe to run after a build or against a build-free fixture; it never invokes
Cargo or a native binary.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import struct
import tomllib
from typing import Any, Mapping

TOOLS_ROOT = Path(__file__).resolve().parents[1]
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))


SCHEMA = "litchi.ole-owner-qualification.runner.v1"
MANIFEST_SCHEMA = "litchi.ole-owner-qualification.manifest.v1"
MEASUREMENTS_SCHEMA = "litchi.ole-owner-qualification.measurements.v1"
ROLES = ("before", "after")
VARIANTS = ("normal", "allocator")
BUILD_KEYS = tuple(f"{role}/{variant}" for role in ROLES for variant in VARIANTS)
ABBA_ORDER = ("before", "after", "after", "before")
ROUTES = (
    "open", "begin_edit", "noop_commit", "replace_stream", "read", "edit",
    "doc_tracked_revision", "xls_cells", "xls_cells_stage", "doc_history", "xls_history",
)
ALLOCATOR_ROUTES = ("read", "edit", "xls_cells", "xls_cells_stage", "doc_history", "xls_history")
ALLOCATOR_EVIDENCE_FIELDS = (
    "route", "alloc_calls", "alloc_bytes", "dealloc_calls", "dealloc_bytes",
    "realloc_calls", "realloc_old_bytes", "realloc_new_bytes", "base_live_calls",
    "base_live_bytes", "live_calls", "live_bytes", "peak_calls", "peak_bytes",
    "retained_packages", "retained_patches",
)
ORACLE_FILES = (
    "ole-source.bin", "ole-noop.bin", "ole-replacement.bin", "properties.tsv",
    "doc-source.bin", "doc-edit.bin", "doc-proof.tsv", "xls-source.bin", "xls-edit.bin",
    "xls-proof.tsv", "xls-opaque-source.bin", "xls-opaque-edit.bin",
    "doc-history-final.bin", "xls-history-final.bin", "history-proof.tsv", "ole-inventory.tsv",
    "oracle.tsv", "allocator.tsv",
)
MIN_SPAN_NS = 250_000_000
BLOCKS = 5
WARMUP = 2
SAMPLES = 3
TARGET_PREFLIGHT_BYTES = 19_200_000_000
TARGET_HARD_BYTES = 20_000_000_000
MEMORY_BYTES = 20_000_000_000
EXPECTED_FEATURES = {
    "before/normal": "legacy-owner-api,legacy-property-api",
    "before/allocator": "legacy-owner-api,legacy-property-api,w5-counting-allocator",
    "after/normal": "current-api",
    "after/allocator": "current-api,w5-counting-allocator",
}
CLAIM_FREE = {
    "build": "orchestration and source-bound evidence only; no performance acceptance claim",
    "abba": "raw paired evidence only; no performance acceptance claim",
    "measurements": "raw paired evidence only; no performance acceptance claim",
    "oracle": "correctness and identity preflight only; no performance acceptance claim",
}

# These are deliberately repeated here instead of importing the runner's
# source parser.  The post-run check must be able to reject a receipt even if
# the orchestration code has accidentally accepted a stale or overlaid source.
BASELINE_CFB_OPEN = "pub fn open(reader: R) -> Result<Self, OleError>"
BASELINE_CFB_LIMITED_OPEN = "pub fn open_with_limits(mut reader: R, limits: OleFileLimits)"
CANDIDATE_CFB_OPEN = "pub fn open(mut reader: R, limits: OleFileLimits) -> Result<Self, OleError>"
CANDIDATE_CFB_NO_DEFAULT_OPEN = "pub fn open(reader: R) -> Result<Self, OleError>"
CFB_SOURCE_PATH = Path("crates/litchi-cfb/src/file.rs")
REVIEWED_LOCK_SHA256 = "bc5a133f0668142d18d637b2b196097e855c8890697cb3bbf43aeb937145f04a"
BASELINE_LOCK_SHA256 = "450f43c6b74536ef552bd5e9a19e321a3c0dea745f645fecc0aa9abe37868e14"
CANDIDATE_LOCK_SHA256 = "5a07adc1d62dea06dbac60986faa020007882506523dde49e4605708e1276aaf"
LOCK_SOURCE = (
    "external matched-harness lock retained by 0359j W5 qualification; "
    "each driver lock is the reviewed base plus only its pinned local dependency edges"
)
LOCK_PATCH_POLICY: dict[str, dict[str, dict[str, list[str]]]] = {
    "baseline": {
        "litchi-perf-baseline": {"added": ["sha2 0.11.0"], "removed": []},
    },
    "candidate": {
        "litchi-odraw": {"added": ["litchi-core"], "removed": []},
        "litchi-perf-baseline": {"added": ["sha2 0.11.0"], "removed": []},
    },
}
LOCK_HASH = re.compile(r"^[0-9a-f]{64}$")
PROBE_NAME = "ole_w5_probe.rs"
CLONE_NAME = "ole_package_clone_probe.rs"
ORACLE_NAME = "ole_owner_oracle.rs"
TARGET_NAME = "ole_w5_edit_target.bin"
REQUIRED_PROBE_MARKERS = (
    "const SIBLINGS: usize = 64;",
    "const DOC_HISTORY_STEPS: usize = 128;",
    "const XLS_HISTORY_STEPS: usize = 128;",
    '"open"', '"begin_edit"', '"noop_commit"', '"replace_stream"',
    '"read"', '"edit"', '"doc_tracked_revision"', '"xls_cells"',
    '"xls_cells_stage"', '"doc_history"', '"xls_history"',
    "verify_doc_history", "verify_xls_history", "patches", "packages",
    "inverse()", "record_dealloc", "record_realloc", "#[global_allocator]",
)
REQUIRED_CLONE_MARKERS = (
    "const SMALL_SIBLINGS: usize = 1;",
    "const LARGE_SIBLINGS: usize = 1_024;",
    "clone_drop", "black_box(&clone)",
)


class ValidationError(RuntimeError):
    """A receipt or measurement is not admissible."""


def _sha256(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise ValidationError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _json(path: Path, label: str) -> tuple[dict[str, Any], str, int]:
    digest, size = _sha256(path)
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read {label}: {error}") from error
    if not isinstance(value, dict):
        raise ValidationError(f"{label} must contain an object")
    return value, digest, size


def _mapping(value: Any, label: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise ValidationError(f"{label} must be an object")
    return value


def _same(value: Any, expected: Any, label: str) -> None:
    if value != expected:
        raise ValidationError(f"{label} differs")


def _path(value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value:
        raise ValidationError(f"{label} path is missing")
    return Path(value).resolve(strict=False)


def _hash_record(record: Mapping[str, Any], label: str) -> None:
    path = _path(record.get("path"), label)
    digest, size = _sha256(path)
    _same(record.get("sha256"), digest, f"{label} SHA-256")
    _same(record.get("bytes"), size, f"{label} byte count")


def _lock_document(path: Path, label: str) -> dict[str, Any]:
    try:
        value = tomllib.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, tomllib.TOMLDecodeError) as error:
        raise ValidationError(f"cannot read {label} as Cargo.lock TOML: {error}") from error
    if not isinstance(value, dict):
        raise ValidationError(f"{label} must contain a TOML object")
    return value


def _lock_packages(document: Mapping[str, Any], label: str) -> dict[tuple[str, str, str | None], Mapping[str, Any]]:
    packages = document.get("package")
    if not isinstance(packages, list) or not packages:
        raise ValidationError(f"{label} has no Cargo.lock packages")
    result: dict[tuple[str, str, str | None], Mapping[str, Any]] = {}
    for index, package in enumerate(packages):
        if not isinstance(package, Mapping):
            raise ValidationError(f"{label} package {index} is not an object")
        name = package.get("name")
        version = package.get("version")
        source = package.get("source")
        if not isinstance(name, str) or not isinstance(version, str) or (source is not None and not isinstance(source, str)):
            raise ValidationError(f"{label} package {index} identity is malformed")
        key = (name, version, source)
        if key in result:
            raise ValidationError(f"{label} contains duplicate package identity {key!r}")
        result[key] = package
    return result


def _lock_registry_identity(packages: Mapping[tuple[str, str, str | None], Mapping[str, Any]], label: str) -> tuple[tuple[str, str, str, str], ...]:
    identity: list[tuple[str, str, str, str]] = []
    for (name, version, source), package in packages.items():
        if source is None:
            continue
        if not source.startswith("registry+"):
            raise ValidationError(f"{label} has an unreviewed non-registry source: {name} {version}")
        checksum = package.get("checksum")
        if not isinstance(checksum, str) or not LOCK_HASH.fullmatch(checksum):
            raise ValidationError(f"{label} registry package lacks a valid checksum: {name} {version}")
        identity.append((name, version, source, checksum))
    return tuple(sorted(identity))


def _validate_lock_patch(base: Path, selected: Path, role: str) -> dict[str, Any]:
    if role not in LOCK_PATCH_POLICY:
        raise ValidationError(f"unknown driver lock role: {role}")
    base_packages = _lock_packages(_lock_document(base, "reviewed base lock"), "reviewed base lock")
    selected_packages = _lock_packages(_lock_document(selected, f"{role} driver lock"), f"{role} driver lock")
    if _lock_registry_identity(base_packages, "reviewed base lock") != _lock_registry_identity(selected_packages, f"{role} driver lock"):
        raise ValidationError(f"{role} driver lock changes a registry package identity")
    if set(base_packages) != set(selected_packages):
        raise ValidationError(f"{role} driver lock changes the package identity set")
    expected = LOCK_PATCH_POLICY[role]
    observed: dict[str, dict[str, list[str]]] = {}
    for key in base_packages:
        base_package = base_packages[key]
        selected_package = selected_packages[key]
        if {name: value for name, value in base_package.items() if name != "dependencies"} != {
            name: value for name, value in selected_package.items() if name != "dependencies"
        }:
            raise ValidationError(f"{role} driver lock changes package metadata for {key[0]} {key[1]}")
        base_dependencies = base_package.get("dependencies", [])
        selected_dependencies = selected_package.get("dependencies", [])
        if not isinstance(base_dependencies, list) or not isinstance(selected_dependencies, list):
            raise ValidationError(f"{role} driver lock dependency list is malformed for {key[0]} {key[1]}")
        allowed = expected.get(key[0], {"added": [], "removed": []})
        added = sorted(set(selected_dependencies) - set(base_dependencies))
        removed = sorted(set(base_dependencies) - set(selected_dependencies))
        observed[key[0]] = {"added": added, "removed": removed}
        if added != sorted(allowed["added"]) or removed != sorted(allowed["removed"]):
            raise ValidationError(
                f"{role} driver lock dependency patch differs for {key[0]}: added={added!r}, removed={removed!r}"
            )
        if [item for item in selected_dependencies if item not in allowed["added"]] != base_dependencies:
            raise ValidationError(f"{role} driver lock reorders or duplicates dependencies for {key[0]}")
    return {
        "base_sha256": REVIEWED_LOCK_SHA256,
        "role": role,
        "patches": observed,
        "policy": expected,
    }


def _validate_toolchain(toolchain: Mapping[str, Any]) -> None:
    """Validate the selected tools and the PATH that selected them.

    ``path`` is the exact absolute dispatcher passed to Cargo (important for
    rustup's cargo/rustc symlinks); ``resolved_path`` is its underlying regular
    file whose hash and size are attested.  The runner also removes compiler
    wrappers and writes these values into every child event, so a changed
    receipt cannot silently select another compiler.
    """

    path_env = toolchain.get("path_env")
    if not isinstance(path_env, str) or not path_env:
        raise ValidationError("build toolchain PATH is missing")
    for name in ("cargo", "rustc"):
        record = _mapping(toolchain.get(name), f"toolchain.{name}")
        _hash_record(record, f"toolchain.{name}")
        path_value = record.get("path")
        resolved_value = record.get("resolved_path")
        if not isinstance(path_value, str) or not Path(path_value).is_absolute():
            raise ValidationError(f"toolchain.{name} path is not absolute")
        if not isinstance(resolved_value, str) or not Path(resolved_value).is_absolute():
            raise ValidationError(f"toolchain.{name} resolved path is not absolute")
        _same(
            _path(resolved_value, f"toolchain.{name} resolved path"),
            _path(path_value, f"toolchain.{name} path"),
            f"toolchain.{name} resolved path binding",
        )
        if not isinstance(record.get("version"), str) or not record["version"].strip():
            raise ValidationError(f"toolchain.{name} version is missing")


def _toolchain_environment(
    toolchain: Mapping[str, Any], target: Path, label: str,
) -> dict[str, Any]:
    _validate_toolchain(toolchain)
    cargo = _mapping(toolchain.get("cargo"), "toolchain.cargo")
    rustc = _mapping(toolchain.get("rustc"), "toolchain.rustc")
    return {
        "CARGO_TARGET_DIR": str(target),
        "CARGO_BUILD_JOBS": "1",
        "CARGO_INCREMENTAL": "0",
        "LITCHI_MEM_MAX": "20G",
        "PATH": toolchain["path_env"],
        "CARGO": cargo["path"],
        "RUSTC": rustc["path"],
        "RUSTC_WRAPPER": None,
        "RUSTC_WORKSPACE_WRAPPER": None,
        "CARGO_BUILD_RUSTC": None,
        "RUSTUP_TOOLCHAIN": None,
    }


def _validate_manifest_policy(manifest_path: Path, manifest: Mapping[str, Any], manifest_hash: str) -> None:
    """Use the runner's canonical policy parser before checking run evidence.

    Keeping this call at the independent boundary prevents the validator from
    silently drifting on routes, ABBA bounds, allocator policy, quiet limits,
    or runtime candidate parameterization.  Preparation validation below then
    binds that parameterization to the actual selected candidate worktree.
    """

    try:
        from native_qualification import ole_owner_paired_runner as paired_runner
        canonical, canonical_hash = paired_runner._load_manifest(manifest_path)
    except ImportError as error:
        raise ValidationError(f"manifest policy validator is unavailable: {error}") from error
    except paired_runner.RunnerError as error:
        raise ValidationError(f"manifest policy validation failed: {error}") from error
    _same(canonical_hash, manifest_hash, "manifest policy file identity")
    _same(canonical, manifest, "manifest policy document")


def _events(path: Path, label: str) -> tuple[list[dict[str, Any]], str, int]:
    digest, size = _sha256(path)
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
        values = [json.loads(line) for line in lines]
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read {label}: {error}") from error
    if not all(isinstance(value, dict) for value in values):
        raise ValidationError(f"{label} contains a non-object event")
    return values, digest, size


def _event_hash(events: list[Mapping[str, Any]]) -> str:
    return hashlib.sha256(json.dumps(events, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def _validate_log(event: Mapping[str, Any], field: str, output_root: Path) -> None:
    record = _mapping(event.get(field), f"event.{field}")
    path = _path(record.get("path"), f"event.{field}")
    try:
        path.relative_to(output_root / "commands")
    except ValueError:
        try:
            path.relative_to(output_root / "runs")
        except ValueError as error:
            raise ValidationError(f"event {field} escapes the output log roots: {path}") from error
    _hash_record(record, f"event.{field}")


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise ValidationError(f"git {' '.join(args)} failed in {root}: {result.stderr.strip()}")
    return result.stdout


def _status_paths(status: str) -> list[str]:
    paths: list[str] = []
    for line in status.splitlines():
        if len(line) < 4:
            raise ValidationError(f"malformed git status line: {line!r}")
        path = line[3:]
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        paths.append(path)
    return paths


def _source_state(root: Path) -> dict[str, Any]:
    root = root.resolve(strict=False)
    if not root.is_dir():
        raise ValidationError(f"source root is not a directory: {root}")
    top = Path(_git(root, "rev-parse", "--show-toplevel").strip()).resolve(strict=False)
    if top != root:
        raise ValidationError(f"source root is not a worktree top-level: {root}")
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    diff = _git(root, "diff", "--no-ext-diff", "--binary")
    return {
        "root": str(root),
        "head": _git(root, "rev-parse", "HEAD").strip(),
        "tree": _git(root, "rev-parse", "HEAD^{tree}").strip(),
        "status": status,
        "status_paths": _status_paths(status),
        "clean": not bool(status),
        "status_sha256": hashlib.sha256(status.encode()).hexdigest(),
        "diff_sha256": hashlib.sha256(diff.encode()).hexdigest(),
        "production_tree_sha256": hashlib.sha256(
            _git(root, "ls-tree", "-r", "--full-tree", "HEAD").encode()
        ).hexdigest(),
    }


def _read_text(path: Path, label: str) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as error:
        raise ValidationError(f"cannot read {label}: {error}") from error


def _validate_source_api(root: Path, *, baseline: bool) -> dict[str, Any]:
    source = root / CFB_SOURCE_PATH
    text = _read_text(source, "selected CFB source")
    if baseline:
        if BASELINE_CFB_OPEN not in text or BASELINE_CFB_LIMITED_OPEN not in text:
            raise ValidationError("baseline CFB source does not expose the reviewed legacy API")
        if CANDIDATE_CFB_OPEN in text:
            raise ValidationError("baseline CFB source exposes the candidate two-argument API")
        return {
            "source_path": str(source.resolve(strict=False)),
            "open_signature": BASELINE_CFB_OPEN,
            "limited_open_signature": BASELINE_CFB_LIMITED_OPEN,
            "api_generation": "pre-owner explicit open_with_limits",
        }
    if CANDIDATE_CFB_OPEN not in text or CANDIDATE_CFB_NO_DEFAULT_OPEN in text:
        raise ValidationError("candidate CFB source does not expose only explicit-limits open")
    return {
        "source_path": str(source.resolve(strict=False)),
        "open_signature": CANDIDATE_CFB_OPEN,
        "api_generation": "current explicit-limits open",
    }


def _validate_probe_api(path: Path, *, baseline: bool) -> dict[str, Any]:
    text = _read_text(path, "selected probe adapter")
    missing = [marker for marker in REQUIRED_PROBE_MARKERS if marker not in text]
    if missing:
        raise ValidationError(f"selected probe is missing workload markers: {missing!r}")
    if baseline:
        if "OleFile::open(Cursor::new(package.into_inner()))" not in text:
            raise ValidationError("baseline probe does not use one-argument CFB open")
        if "OleFileLimits::default()" in text:
            raise ValidationError("baseline probe unexpectedly passes CFB limits")
        return {"open_call": "OleFile::open(reader)", "adapter": "immutable historical baseline"}
    if text.count("OleFileLimits::default()") < 2 or "OleFile::open(\n        Cursor::new(package.into_inner())," not in text:
        raise ValidationError("candidate probe is not the explicit-CFB-limits adapter")
    return {"open_call": "OleFile::open(reader, OleFileLimits::default())", "adapter": "CFB limits arity only"}


def _validate_clone_input(path: Path) -> None:
    text = _read_text(path, "selected clone probe")
    missing = [marker for marker in REQUIRED_CLONE_MARKERS if marker not in text]
    if missing:
        raise ValidationError(f"selected clone probe is missing workload markers: {missing!r}")


def _validate_driver(endpoint: Mapping[str, Any], *, baseline: bool) -> None:
    source_root = _path(endpoint.get("root"), "preparation source root")
    driver = _mapping(endpoint.get("driver"), "preparation driver")
    driver_root = _path(driver.get("root"), "preparation driver root")
    files = driver.get("files")
    if not isinstance(files, list) or len(files) != 5:
        raise ValidationError("each prepared endpoint must bind exactly five driver files")
    expected = {"Cargo.toml", "Cargo.lock", PROBE_NAME, CLONE_NAME, ORACLE_NAME}
    actual: set[str] = set()
    for index, raw in enumerate(files):
        item = _mapping(raw, f"driver file {index}")
        path = _path(item.get("path"), f"driver file {index}")
        try:
            path.relative_to(driver_root)
        except ValueError as error:
            raise ValidationError(f"driver file escapes its root: {path}") from error
        if path.name in actual:
            raise ValidationError(f"driver file names are duplicated: {path.name}")
        actual.add(path.name)
        _hash_record(item, f"driver file {index}")
    _same(actual, expected, "prepared driver file names")
    existing = {
        path.name
        for path in driver_root.rglob("*")
        if path.is_file() and not path.is_symlink()
    }
    _same(existing, expected, "prepared driver contents")
    manifest = driver_root / "Cargo.toml"
    manifest_text = _read_text(manifest, "prepared Cargo.toml")
    if "@SOURCE_ROOT@" in manifest_text or str(source_root) not in manifest_text:
        raise ValidationError("prepared Cargo.toml is not bound to its selected source root")
    if "[workspace]" not in manifest_text or "litchi-cfb" not in manifest_text:
        raise ValidationError("prepared Cargo.toml is not the reviewed standalone driver")
    probe = driver_root / "examples" / PROBE_NAME
    clone = driver_root / "examples" / CLONE_NAME
    oracle = driver_root / "examples" / ORACLE_NAME
    _validate_probe_api(probe, baseline=baseline)
    _validate_clone_input(clone)
    oracle_text = _read_text(oracle, "prepared oracle")
    for marker in ("fn check_ole_outputs", "allocator_evidence", "owner_target"):
        if marker not in oracle_text:
            raise ValidationError(f"prepared oracle is missing marker {marker!r}")


def _validate_preparation(preparation: Mapping[str, Any], manifest: Mapping[str, Any]) -> None:
    _same(preparation.get("schema"), "litchi.ole-owner-qualification.preparation.v1", "preparation schema")
    _same(preparation.get("lane"), "w5-ole-owner", "preparation lane")
    _same(preparation.get("claim_scope"), "external source and consumer preparation only; no Cargo, timing, or acceptance claim", "preparation claim scope")
    overlay = _mapping(preparation.get("baseline_overlay"), "preparation.baseline_overlay")
    _same(overlay.get("allowed_paths"), [], "baseline overlay")
    _same(overlay.get("production_tree"), "unchanged; driver files are external", "baseline production-tree policy")
    shared = _mapping(preparation.get("shared_inputs"), "preparation.shared_inputs")
    _same(shared.get("routes"), list(ROUTES), "preparation routes")
    _same(shared.get("history_retention"), {"doc_steps": 128, "xls_steps": 128, "packages": 129, "patches": 128}, "preparation history retention")
    _same(shared.get("clone_sizes"), {"small_siblings": 1, "large_siblings": 1024}, "preparation clone sizes")
    api = _mapping(shared.get("api_binding"), "preparation API binding")
    _same(api.get("policy"), "the baseline one-argument open and candidate explicit-limits open are checked against source text; no implementation overlay", "preparation API binding policy")
    pins = _mapping(manifest.get("input_pins"), "manifest input pins")
    policy = _mapping(manifest.get("source_policy"), "manifest source policy")
    _same(policy.get("baseline_commit"), _mapping(preparation.get("baseline"), "preparation.baseline").get("commit"), "baseline source commit pin")
    for endpoint in ROLES:
        source_key = "baseline" if endpoint == "before" else "candidate"
        _mapping(api.get(source_key), f"preparation API binding.{endpoint}")
        source = _mapping(preparation.get(source_key), f"preparation.{endpoint}")
        _same(source.get("clean"), True, f"preparation.{endpoint} clean")
        source_root = _path(source.get("root"), f"preparation.{endpoint}.root")
        current = _source_state(source_root)
        for key in (
            "root", "head", "tree", "status", "status_paths", "clean",
            "status_sha256", "diff_sha256", "production_tree_sha256",
        ):
            _same(current.get(key), source.get(key), f"preparation.{endpoint} current {key}")
        if endpoint == "before":
            _same(source.get("after_preparation"), current, "preparation.baseline post-driver source identity")
        _same(current.get("head"), source.get("commit"), f"preparation.{endpoint} commit")
        _same(current.get("clean"), True, f"preparation.{endpoint} source cleanliness")
        driver = _mapping(source.get("driver"), f"preparation.{endpoint}.driver")
        files = driver.get("files")
        if not isinstance(files, list) or len(files) != 5:
            raise ValidationError(f"preparation.{endpoint} does not bind five driver inputs")
        names = {Path(str(_mapping(item, "driver file").get("path"))).name for item in files}
        _same(names, {"Cargo.toml", "Cargo.lock", "ole_w5_probe.rs", "ole_package_clone_probe.rs", "ole_owner_oracle.rs"}, f"preparation.{endpoint} driver files")
        for index, item in enumerate(files):
            _hash_record(_mapping(item, "driver file"), f"preparation.{endpoint}.driver[{index}]")
        role = "baseline" if endpoint == "before" else "candidate"
        driver_locks = _mapping(shared.get("driver_locks"), "preparation.driver_locks")
        _same(set(driver_locks), {"baseline", "candidate"}, "preparation driver lock roles")
        driver_lock = _mapping(driver_locks.get(role), f"preparation.{role} driver lock")
        expected_lock_hash = BASELINE_LOCK_SHA256 if role == "baseline" else CANDIDATE_LOCK_SHA256
        _same(driver_lock.get("provenance"), LOCK_SOURCE, f"preparation.{role} driver lock provenance")
        _same(driver_lock.get("sha256"), expected_lock_hash, f"preparation.{role} driver lock pin")
        _same(driver_lock.get("base_sha256"), REVIEWED_LOCK_SHA256, f"preparation.{role} lock base pin")
        _same(driver_lock.get("role"), role, f"preparation.{role} lock role")
        _same(driver_lock.get("policy"), LOCK_PATCH_POLICY[role], f"preparation.{role} lock policy")
        driver_lock_path = _path(driver_lock.get("path"), f"preparation.{role} driver lock")
        _hash_record(driver_lock, f"preparation.{role} driver lock")
        _same(
            _validate_lock_patch(
                _path(_mapping(shared.get("reviewed_lock"), "preparation.reviewed_lock").get("path"), "reviewed lock"),
                driver_lock_path,
                role,
            ),
            {key: driver_lock[key] for key in ("base_sha256", "role", "patches", "policy")},
            f"preparation.{role} lock graph",
        )
        shared_hashes = {
            "Cargo.lock": expected_lock_hash,
            CLONE_NAME: _mapping(shared.get("clone_probe"), "preparation.clone_probe").get("sha256"),
            ORACLE_NAME: _mapping(shared.get("oracle_probe"), "preparation.oracle_probe").get("sha256"),
            PROBE_NAME: (
                _mapping(shared.get("reviewed_probe"), "preparation.reviewed_probe").get("sha256")
                if endpoint == "before" else
                _mapping(source.get("probe"), "preparation.candidate.probe").get("sha256")
            ),
        }
        for item in files:
            record = _mapping(item, "driver file")
            name = Path(str(record.get("path"))).name
            if name != "Cargo.toml":
                _same(record.get("sha256"), shared_hashes[name], f"preparation.{endpoint} driver {name} pin")
        if endpoint == "after":
            _same(
                _path(_mapping(source.get("probe"), "preparation.candidate.probe").get("path"), "preparation.candidate.probe path"),
                source_root / Path("tools/perf-baseline/examples/ole_w5_probe.rs"),
                "candidate probe source path",
            )
        _validate_driver(source, baseline=endpoint == "before")
        actual_api = _validate_source_api(source_root, baseline=endpoint == "before")
        _same(actual_api, api.get(source_key, {}).get("cfb"), f"preparation.{endpoint} CFB API binding")
        adapter_path = (
            _path(_mapping(shared.get("reviewed_probe"), "preparation.reviewed_probe").get("path"), "reviewed probe")
            if endpoint == "before"
            else _path(_mapping(source.get("probe"), "preparation.candidate.probe").get("path"), "candidate probe")
        )
        actual_probe = _validate_probe_api(adapter_path, baseline=endpoint == "before")
        _same(actual_probe, api.get(source_key, {}).get("probe"), f"preparation.{endpoint} probe API binding")
    for name, provenance in (
        ("reviewed_probe", "external matched-harness artifact from 0359j; historical source a84650cd70390641f7734dc1f3362bcb9cd1e025 tools/perf-baseline/examples/ole_w5_probe.rs"),
        ("clone_probe", "external matched-harness artifact from 0359j; historical source a84650cd70390641f7734dc1f3362bcb9cd1e025 tools/perf-baseline/examples/ole_package_clone_probe.rs"),
        ("reviewed_lock", LOCK_SOURCE),
    ):
        item = _mapping(shared.get(name), f"preparation.{name}")
        _same(item.get("provenance"), provenance, f"preparation.{name} provenance")
        _hash_record(item, f"preparation.{name}")
        pin_name = {
            "reviewed_probe": "reviewed_probe_sha256",
            "clone_probe": "clone_probe_sha256",
            "reviewed_lock": "reviewed_lock_sha256",
        }[name]
        _same(item.get("sha256"), pins.get(pin_name), f"manifest {pin_name}")
    oracle = _mapping(shared.get("oracle_probe"), "preparation.oracle_probe")
    _same(oracle.get("provenance"), "tracked source-only correctness oracle; compiled in both endpoint drivers", "preparation.oracle provenance")
    _hash_record(oracle, "preparation.oracle")
    _same(oracle.get("sha256"), pins.get("oracle_source_sha256"), "manifest oracle_source_sha256")
    candidate_probe = _mapping(_mapping(preparation.get("candidate"), "preparation.candidate").get("probe"), "preparation.candidate.probe")
    _hash_record(candidate_probe, "preparation candidate probe")
    _same(candidate_probe.get("sha256"), pins.get("candidate_adapter_sha256"), "manifest candidate_adapter_sha256")


def _source_identity(preparation: Mapping[str, Any], role: str) -> dict[str, Any]:
    endpoint_name = "baseline" if role == "before" else "candidate"
    endpoint = _mapping(preparation.get(endpoint_name), f"preparation.{endpoint_name}")
    source = {key: endpoint.get(key) for key in (
        "root", "head", "tree", "status", "status_paths", "clean",
        "status_sha256", "diff_sha256", "production_tree_sha256",
    )}
    driver = _mapping(endpoint.get("driver"), f"preparation.{endpoint_name}.driver")
    files = driver.get("files")
    if not isinstance(files, list):
        raise ValidationError(f"preparation.{endpoint_name}.driver files are missing")
    driver_files: list[dict[str, Any]] = []
    for raw in files:
        item = _mapping(raw, "driver file")
        path = _path(item.get("path"), "driver file path")
        digest, size = _sha256(path)
        _same(item.get("sha256"), digest, f"{role} driver file hash")
        _same(item.get("bytes"), size, f"{role} driver file size")
        driver_files.append({"path": str(path), "bytes": size, "sha256": digest})
    identity: dict[str, Any] = {
        "source": source,
        "driver_root": str(_path(driver.get("root"), f"{role} driver root")),
        "driver_files": driver_files,
    }
    identity["identity_sha256"] = hashlib.sha256(
        json.dumps(identity, sort_keys=True, separators=(",", ":")).encode()
    ).hexdigest()
    return identity


def _expected_build_command(
    item: Mapping[str, Any], preparation: Mapping[str, Any], role: str, variant: str,
    cargo_path: Path, cargo_dispatch_path: str,
) -> list[str]:
    command = item.get("command")
    if not isinstance(command, list) or not command:
        raise ValidationError(f"builds.{role}/{variant} command is missing")
    cargo = command[0]
    if not isinstance(cargo, str) or not cargo:
        raise ValidationError(f"builds.{role}/{variant} cargo executable is not explicit")
    _same(cargo, cargo_dispatch_path, f"builds.{role}/{variant} Cargo dispatch path")
    _same(_path(cargo, f"builds.{role}/{variant} cargo executable"), cargo_path, f"builds.{role}/{variant} Cargo executable binding")
    endpoint = _mapping(
        preparation.get("baseline" if role == "before" else "candidate"),
        f"preparation.{role}",
    )
    driver_root = _path(
        _mapping(endpoint.get("driver"), f"preparation.{role}.driver").get("root"),
        f"{role} driver root",
    )
    return [
        cargo, "build", "--release", "--offline", "--locked", "--manifest-path",
        str(driver_root / "Cargo.toml"), "--no-default-features", "--features",
        EXPECTED_FEATURES[f"{role}/{variant}"], "--example", "ole_w5_probe",
        "--example", "ole_package_clone_probe", "--example", "ole_owner_oracle",
    ]


def _validate_build(
    manifest: Mapping[str, Any], preparation: Mapping[str, Any], build: Mapping[str, Any],
    build_path: Path, runner_path: Path, manifest_hash: str, preparation_hash: str,
) -> tuple[str, Path, list[dict[str, Any]]]:
    _same(build.get("schema"), SCHEMA, "build receipt schema")
    _same(build.get("phase"), "build", "build receipt phase")
    _same(build.get("status"), "passed", "build receipt status")
    _same(build.get("claim_scope"), CLAIM_FREE["build"], "build claim scope")
    _same(_mapping(build.get("manifest"), "build manifest").get("sha256"), manifest_hash, "build manifest identity")
    _same(_mapping(build.get("preparation"), "build preparation").get("sha256"), preparation_hash, "build preparation identity")
    receipt_runner = _mapping(build.get("runner"), "build runner")
    _same(_path(receipt_runner.get("path"), "build runner"), runner_path, "build runner path")
    _hash_record(receipt_runner, "build runner")
    limits = _mapping(build.get("limits"), "build limits")
    _same(limits, {"target_preflight_bytes": TARGET_PREFLIGHT_BYTES, "target_hard_bytes": TARGET_HARD_BYTES, "memory_bytes": MEMORY_BYTES}, "build limits")
    toolchain = _mapping(build.get("toolchain"), "build toolchain")
    _validate_toolchain(toolchain)
    root = build_path.parent
    journal = _mapping(build.get("event_journal"), "build event journal")
    journal_path = _path(journal.get("path"), "build event journal")
    _same(journal_path, root / "events-build.jsonl", "build event journal path")
    events, digest, size = _events(journal_path, "build event journal")
    _same(journal.get("count"), len(events), "build journal count")
    _same(journal.get("sha256"), _event_hash(events), "build journal semantic hash")
    _same(journal.get("file_sha256"), digest, "build journal file hash")
    _same(journal.get("bytes"), size, "build journal byte count")
    _same([event.get("event_id") for event in events], [f"build-{key.replace('/', '-')}" for key in BUILD_KEYS], "build event order")
    builds = _mapping(build.get("builds"), "builds")
    identities = _mapping(build.get("identities"), "build identities")
    target_root = _path(build.get("target_root"), "build target root")
    expected_identities = {
        role: _source_identity(preparation, role) for role in ROLES
    }
    for role in ROLES:
        _same(identities.get(role), expected_identities[role], f"build identity {role}")
    for key in BUILD_KEYS:
        role, variant = key.split("/")
        item = _mapping(builds.get(key), f"builds.{key}")
        _same(item.get("features"), EXPECTED_FEATURES[key], f"builds.{key} features")
        _same(item.get("role"), role, f"builds.{key} role")
        _same(item.get("variant"), variant, f"builds.{key} variant")
        expected_target = target_root / f"{role}-{variant}"
        _same(_path(item.get("target"), f"builds.{key} target"), expected_target, f"builds.{key} target")
        cargo_record = _mapping(toolchain.get("cargo"), "toolchain.cargo")
        cargo_dispatch_path = cargo_record.get("path")
        if not isinstance(cargo_dispatch_path, str):
            raise ValidationError("toolchain.cargo path is missing")
        cargo_path = _path(cargo_dispatch_path, "toolchain.cargo path")
        expected_command = _expected_build_command(
            item, preparation, role, variant, cargo_path, cargo_dispatch_path,
        )
        _same(item.get("command"), expected_command, f"builds.{key} command")
        event = events[BUILD_KEYS.index(key)]
        _same(event.get("phase"), "build", f"build event {key} phase")
        _same(event.get("role"), role, f"build event {key} role")
        _same(event.get("variant"), variant, f"build event {key} variant")
        identity = _mapping(identities.get(role), f"build identities.{role}")
        _same(event.get("source_identity_sha256"), identity.get("identity_sha256"), f"build event {key} source identity hash")
        _same(event.get("source_identity_before"), identity, f"build event {key} source identity before")
        _same(event.get("source_identity_after"), identity, f"build event {key} source identity after")
        _same(event.get("argv"), expected_command, f"build event {key} argv")
        _same(event.get("cwd"), str(_path(_mapping(
            _mapping(preparation.get("baseline" if role == "before" else "candidate"),
                     f"preparation.{role}").get("driver"), f"preparation.{role}.driver").get("root"),
            f"{role} driver root",
        )), f"build event {key} cwd")
        _same(event.get("returncode"), 0, f"build event {key} returncode")
        _same(event.get("runner_error"), None, f"build event {key} runner error")
        if event.get("process_guard_violations") or event.get("post_exit_descendants") or event.get("post_exit_offenders"):
            raise ValidationError(f"build event {key} was not quiet")
        expected_environment = _toolchain_environment(toolchain, expected_target, f"build event {key}")
        environment = _mapping(event.get("environment"), f"build event {key} environment")
        _same(environment, expected_environment, f"build event {key} environment")
        _same(event.get("toolchain"), toolchain, f"build event {key} toolchain identity")
        target = _mapping(event.get("target"), f"build event {key} target")
        _same(_path(target.get("path"), f"build event {key} target path"), expected_target, f"build event {key} target path")
        _validate_log(event, "stdout", root)
        _validate_log(event, "stderr", root)
        outputs = event.get("outputs")
        artifacts = _mapping(item.get("artifacts"), f"builds.{key}.artifacts")
        expected_output_paths = [
            str(expected_target / "release" / "examples" / name)
            for name in ("ole_w5_probe", "ole_package_clone_probe", "ole_owner_oracle")
        ]
        if not isinstance(outputs, list) or not all(isinstance(value, Mapping) for value in outputs) or [value.get("path") for value in outputs] != expected_output_paths:
            raise ValidationError(f"build event {key} output set differs")
        for name in ("probe", "clone", "oracle"):
            artifact = _mapping(artifacts.get(name), f"builds.{key}.{name}")
            path = _path(artifact.get("path"), f"builds.{key}.{name}")
            expected_artifact = expected_target / "release" / "examples" / {
                "probe": "ole_w5_probe",
                "clone": "ole_package_clone_probe",
                "oracle": "ole_owner_oracle",
            }[name]
            _same(path, expected_artifact, f"artifact {key}/{name} path")
            digest, bytes_ = _sha256(path)
            _same(artifact.get("sha256"), digest, f"artifact {key}/{name} hash")
            _same(artifact.get("bytes"), bytes_, f"artifact {key}/{name} size")
            if not path.stat().st_mode & 0o111:
                raise ValidationError(f"artifact {key}/{name} is not executable")
            event_output = next((value for value in outputs if isinstance(value, Mapping) and value.get("path") == str(path)), None)
            if event_output is None or event_output.get("sha256") != digest or event_output.get("bytes") != bytes_ or event_output.get("exists") is not True:
                raise ValidationError(f"build event {key}/{name} artifact handoff differs")
    return _sha256_file(build_path)[0], root, events


def _sha256_file(path: Path) -> tuple[str, int]:
    return _sha256(path)


def _hash_json(value: Mapping[str, Any]) -> str:
    # Receipts bind the bytes on disk, not a re-serialized object.  Callers
    # replace this helper with the path hash below when they have a path.
    raw = json.dumps(value, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(raw).hexdigest()


def _event_index(events: list[dict[str, Any]]) -> dict[str, Mapping[str, Any]]:
    index: dict[str, Mapping[str, Any]] = {}
    for event in events:
        event_id = event.get("event_id")
        if not isinstance(event_id, str) or not event_id or event_id in index:
            raise ValidationError("ABBA event IDs must be unique and nonempty")
        index[event_id] = event
    return index


def _endpoint(preparation: Mapping[str, Any], role: str) -> Mapping[str, Any]:
    return _mapping(preparation.get("baseline" if role == "before" else "candidate"), f"preparation.{role}")


def _event_common(
    event: Mapping[str, Any], *, event_id: str, role: str, variant: str,
    preparation: Mapping[str, Any], build: Mapping[str, Any], output_root: Path,
) -> None:
    _same(event.get("event_id"), event_id, f"ABBA event {event_id} ID")
    _same(event.get("phase"), "abba", f"ABBA event {event_id} phase")
    _same(event.get("role"), role, f"ABBA event {event_id} role")
    _same(event.get("variant"), variant, f"ABBA event {event_id} variant")
    _same(event.get("returncode"), 0, f"ABBA event {event_id} returncode")
    _same(event.get("runner_error"), None, f"ABBA event {event_id} runner error")
    if event.get("process_guard_violations") or event.get("post_exit_descendants") or event.get("post_exit_offenders"):
        raise ValidationError(f"ABBA event {event_id} was not quiet")
    identity = _mapping(_mapping(build.get("identities"), "build identities").get(role), f"build identity {role}")
    _same(event.get("source_identity_sha256"), identity.get("identity_sha256"), f"ABBA event {event_id} source identity")
    _same(event.get("source_identity_before"), identity, f"ABBA event {event_id} source identity before")
    _same(event.get("source_identity_after"), identity, f"ABBA event {event_id} source identity after")
    driver_root = _path(_mapping(_endpoint(preparation, role).get("driver"), f"{role} driver").get("root"), f"{role} driver root")
    _same(_path(event.get("cwd"), f"ABBA event {event_id} cwd"), driver_root, f"ABBA event {event_id} cwd")
    target_root = _path(build.get("target_root"), "build target root")
    expected_target = target_root / f"{role}-{variant}"
    target = _mapping(event.get("target"), f"ABBA event {event_id} target")
    _same(_path(target.get("path"), f"ABBA event {event_id} target path"), expected_target, f"ABBA event {event_id} target path")
    environment = _mapping(event.get("environment"), f"ABBA event {event_id} environment")
    toolchain = _mapping(build.get("toolchain"), "build toolchain")
    _same(
        environment,
        _toolchain_environment(toolchain, expected_target, f"ABBA event {event_id}"),
        f"ABBA event {event_id} environment",
    )
    _same(event.get("toolchain"), toolchain, f"ABBA event {event_id} toolchain identity")
    _validate_log(event, "stdout", output_root)
    _validate_log(event, "stderr", output_root)


def _stdout_lines(event: Mapping[str, Any], label: str) -> list[str]:
    record = _mapping(event.get("stdout"), f"{label}.stdout")
    path = _path(record.get("path"), f"{label}.stdout")
    try:
        return path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeDecodeError) as error:
        raise ValidationError(f"cannot read {label} stdout: {error}") from error


def _validate_probe_output(event: Mapping[str, Any], rows: list[Mapping[str, Any]], *, clone: bool, label: str) -> None:
    lines = _stdout_lines(event, label)
    if clone:
        if not lines or lines[0] != "route,size,sample,nanoseconds,batch,archive_bytes":
            raise ValidationError(f"{label} clone output header differs")
        actual: list[dict[str, int | str]] = []
        for line in lines[1:]:
            fields = line.split(",")
            if len(fields) != 6 or fields[0] != "clone_drop" or fields[1] not in {"small", "large"}:
                raise ValidationError(f"{label} clone output row differs")
            if not all(field.isdigit() for field in fields[2:]):
                raise ValidationError(f"{label} clone output contains a non-integer")
            actual.append({"size": fields[1], "sample": int(fields[2]), "nanoseconds": int(fields[3]), "batch": int(fields[4]), "output_bytes": int(fields[5])})
        expected = [
            {"size": row.get("size"), "sample": row.get("sample"), "nanoseconds": row.get("nanoseconds"), "output_bytes": row.get("output_bytes")}
            for row in rows
        ]
        selected = [{key: value for key, value in item.items() if key != "batch"} for item in actual]
        if selected != expected:
            raise ValidationError(f"{label} clone output is not the measured output")
        return
    if not lines or lines[0] != "route,sample,nanoseconds,output_bytes":
        raise ValidationError(f"{label} route output header differs")
    actual_rows: list[dict[str, int]] = []
    for line in lines[1:]:
        fields = line.split(",")
        if len(fields) != 4 or fields[0] != rows[0].get("route", fields[0]) or not all(field.isdigit() for field in fields[1:]):
            # The measurement rows do not repeat the route; the caller binds
            # the route in the command and supplies it in the expected rows.
            if len(fields) != 4 or not all(field.isdigit() for field in fields[1:]):
                raise ValidationError(f"{label} route output row differs")
        actual_rows.append({"sample": int(fields[1]), "nanoseconds": int(fields[2]), "output_bytes": int(fields[3])})
    expected = [{key: int(row[key]) for key in ("sample", "nanoseconds", "output_bytes")} for row in rows]
    if actual_rows != expected:
        raise ValidationError(f"{label} route output is not the measured output")


def _validate_measurements(
    measurements: Mapping[str, Any], measurements_path: Path, *, events: Mapping[str, Mapping[str, Any]],
    preparation: Mapping[str, Any], build: Mapping[str, Any], output_root: Path,
) -> set[str]:
    _same(measurements.get("schema"), MEASUREMENTS_SCHEMA, "measurements schema")
    _same(measurements.get("status"), "complete; coordinator review pending", "measurements status")
    _same(measurements.get("claim_scope"), CLAIM_FREE["measurements"], "measurements claim scope")
    _same(measurements.get("routes"), list(ROUTES), "measurements routes")
    _same(measurements.get("abba_order"), list(ABBA_ORDER), "measurements ABBA order")
    _same(measurements.get("blocks"), BLOCKS, "measurements block count")
    _same(measurements.get("warmup"), WARMUP, "measurements warmup")
    _same(measurements.get("samples"), SAMPLES, "measurements samples")
    _same(measurements.get("minimum_span_ns"), MIN_SPAN_NS, "measurements minimum span")
    timing = _mapping(measurements.get("timing"), "measurements timing")
    batches = _mapping(timing.get("batches"), "measurements timing batches")
    clone_batch = _mapping(batches.get("clone"), "measurements clone batches").get("run")
    if not isinstance(clone_batch, int) or clone_batch <= 0:
        raise ValidationError("clone run batch is missing")
    clone_blocks = timing.get("clone")
    if not isinstance(clone_blocks, list) or len(clone_blocks) != BLOCKS:
        raise ValidationError("timing clone does not contain five blocks")
    referenced: set[str] = set()
    for block_index, block in enumerate(clone_blocks):
        legs = block.get("legs") if isinstance(block, Mapping) else None
        if not isinstance(legs, list) or not all(isinstance(leg, Mapping) for leg in legs) or [leg.get("role") for leg in legs] != list(ABBA_ORDER):
            raise ValidationError("timing clone ABBA legs differ")
        for leg_index, leg in enumerate(legs):
            rows = leg.get("rows")
            if not isinstance(rows, list) or not all(isinstance(row, Mapping) for row in rows) or len(rows) != 2 * SAMPLES:
                raise ValidationError("timing clone sample count differs")
            if any(int(row.get("nanoseconds", 0)) < MIN_SPAN_NS for row in rows if isinstance(row, Mapping)):
                raise ValidationError("timing clone leg is shorter than minimum span")
            event_id = leg.get("event_id")
            expected_id = f"run-clone-b{block_index:02d}-{leg_index}-{leg.get('role')}"
            if not isinstance(event_id, str) or event_id in referenced:
                raise ValidationError("timing clone event links are incomplete or duplicated")
            _same(event_id, expected_id, "timing clone event ID")
            event = events.get(event_id)
            if event is None:
                raise ValidationError(f"timing clone references missing event {event_id}")
            _event_common(event, event_id=event_id, role=str(leg.get("role")), variant="normal", preparation=preparation, build=build, output_root=output_root)
            clone_build = _mapping(_mapping(build.get("builds"), "builds").get(f"{leg.get('role')}/normal"), "clone build")
            clone_artifacts = _mapping(clone_build.get("artifacts"), "clone artifacts")
            binary_path = _path(_mapping(clone_artifacts.get("clone"), "clone artifact").get("path"), "clone artifact")
            expected_argv = [str(binary_path), "--warmup", str(WARMUP), "--samples", str(SAMPLES), "--batch", str(clone_batch)]
            _same(event.get("argv"), expected_argv, f"clone event {event_id} argv")
            _validate_probe_output(event, rows, clone=True, label=event_id)
            referenced.add(event_id)
    route_reports = timing.get("routes")
    if not isinstance(route_reports, Mapping) or set(route_reports) != set(ROUTES):
        raise ValidationError("timing route set differs")
    for route in ROUTES:
        blocks = route_reports[route]
        if not isinstance(blocks, list) or len(blocks) != BLOCKS:
            raise ValidationError(f"timing route {route} does not contain five blocks")
        for block_index, block in enumerate(blocks):
            legs = block.get("legs") if isinstance(block, Mapping) else None
            if not isinstance(legs, list) or not all(isinstance(leg, Mapping) for leg in legs) or [leg.get("role") for leg in legs] != list(ABBA_ORDER):
                raise ValidationError(f"timing route {route} ABBA legs differ")
            for leg_index, leg in enumerate(legs):
                rows = leg.get("rows")
                if not isinstance(rows, list) or not all(isinstance(row, Mapping) for row in rows) or len(rows) != SAMPLES:
                    raise ValidationError(f"timing route {route} sample count differs")
                if any(int(row.get("nanoseconds", 0)) < MIN_SPAN_NS for row in rows if isinstance(row, Mapping)):
                    raise ValidationError(f"timing route {route} leg is shorter than minimum span")
                event_id = leg.get("event_id")
                expected_id = f"run-route-{route}-b{block_index:02d}-{leg_index}-{leg.get('role')}"
                if not isinstance(event_id, str) or event_id in referenced:
                    raise ValidationError("timing route event links are incomplete or duplicated")
                _same(event_id, expected_id, "timing route event ID")
                event = events.get(event_id)
                if event is None:
                    raise ValidationError(f"timing route references missing event {event_id}")
                role = str(leg.get("role"))
                _event_common(event, event_id=event_id, role=role, variant="normal", preparation=preparation, build=build, output_root=output_root)
                build_item = _mapping(_mapping(build.get("builds"), "builds").get(f"{role}/normal"), "route build")
                probe_path = _path(_mapping(_mapping(build_item.get("artifacts"), "route artifacts").get("probe"), "route probe").get("path"), "route probe")
                route_batch = _mapping(batches.get(route), f"{route} batches").get("run")
                if not isinstance(route_batch, int) or route_batch <= 0:
                    raise ValidationError(f"{route} run batch is missing")
                expected_argv = [str(probe_path), "--route", route, "--warmup", str(WARMUP), "--samples", str(SAMPLES), "--batch", str(route_batch)]
                _same(event.get("argv"), expected_argv, f"route event {event_id} argv")
                expected_rows = [dict(row, route=route) for row in rows]
                _validate_probe_output(event, expected_rows, clone=False, label=event_id)
                referenced.add(event_id)
    allocation = _mapping(measurements.get("allocation"), "measurements allocation")
    _same(allocation.get("abba_order"), list(ABBA_ORDER), "allocation ABBA order")
    _same(allocation.get("elapsed_claimable"), False, "allocation claim policy")
    routes = _mapping(allocation.get("routes"), "allocation routes")
    if set(routes) != set(ALLOCATOR_ROUTES):
        raise ValidationError("allocation route set differs")
    for route in ALLOCATOR_ROUTES:
        legs = _mapping(routes[route], f"allocation.{route}").get("legs")
        if not isinstance(legs, list) or not all(isinstance(leg, Mapping) for leg in legs) or [leg.get("role") for leg in legs] != list(ABBA_ORDER):
            raise ValidationError(f"allocation {route} ABBA legs differ")
        for leg_index, leg in enumerate(legs):
            event_id = leg.get("event_id")
            expected_id = f"run-alloc-{route}-{leg_index}-{leg.get('role')}"
            if not isinstance(event_id, str) or event_id in referenced:
                raise ValidationError("allocation event links are incomplete or duplicated")
            _same(event_id, expected_id, "allocation event ID")
            event = events.get(event_id)
            if event is None:
                raise ValidationError(f"allocation references missing event {event_id}")
            role = str(leg.get("role"))
            _event_common(event, event_id=event_id, role=role, variant="allocator", preparation=preparation, build=build, output_root=output_root)
            build_item = _mapping(_mapping(build.get("builds"), "builds").get(f"{role}/allocator"), "allocator build")
            probe_path = _path(_mapping(_mapping(build_item.get("artifacts"), "allocator artifacts").get("probe"), "allocator probe").get("path"), "allocator probe")
            expected_argv = [str(probe_path), "--route", route, "--warmup", "0", "--samples", "1", "--batch", "1"]
            _same(event.get("argv"), expected_argv, f"allocator event {event_id} argv")
            lines = _stdout_lines(event, event_id)
            if not lines or lines[0] != "route,sample,nanoseconds,output_bytes":
                raise ValidationError(f"allocator event {event_id} output header differs")
            normal = [line for line in lines[1:] if line and not line.startswith("w5_alloc,")]
            if len(normal) != 1:
                raise ValidationError(f"allocator event {event_id} normal output cardinality differs")
            fields = normal[0].split(",")
            if len(fields) != 4 or fields[0] != route or fields[1] != "0" or not all(value.isdigit() for value in fields[1:]):
                raise ValidationError(f"allocator event {event_id} output row differs")
            _same(int(fields[3]), _mapping(leg.get("report"), "allocation report").get("output_bytes"), f"allocator event {event_id} output extent")
            referenced.add(event_id)
    return referenced


def _validate_calibration_events(
    events: Mapping[str, Mapping[str, Any]], *, referenced: set[str], preparation: Mapping[str, Any],
    build: Mapping[str, Any], output_root: Path,
) -> None:
    extras = set(events) - referenced
    if not extras:
        raise ValidationError("ABBA journal has no calibration events")
    categories: dict[tuple[str, str], list[int]] = {}
    for event_id in sorted(extras):
        parts = event_id.split("-")
        if len(parts) != 5 or parts[:2] != ["run", "cal"] or parts[2] not in ROLES or (parts[3] != "clone" and parts[3] not in ROUTES):
            raise ValidationError(f"unreferenced ABBA event is not a calibration event: {event_id}")
        role, operation = parts[2], parts[3]
        try:
            attempt = int(parts[4])
        except ValueError as error:
            raise ValidationError(f"calibration event attempt is not numeric: {event_id}") from error
        if attempt <= 0:
            raise ValidationError(f"calibration event attempt is invalid: {event_id}")
        category = (role, operation)
        categories.setdefault(category, []).append(attempt)
        event = events[event_id]
        _event_common(event, event_id=event_id, role=role, variant="normal", preparation=preparation, build=build, output_root=output_root)
        item = _mapping(_mapping(build.get("builds"), "builds").get(f"{role}/normal"), "calibration build")
        artifacts = _mapping(item.get("artifacts"), "calibration artifacts")
        if operation == "clone":
            binary = _path(_mapping(artifacts.get("clone"), "calibration clone artifact").get("path"), "calibration clone binary")
            argv = event.get("argv")
            if not isinstance(argv, list) or len(argv) != 7 or argv[:5] != [str(binary), "--warmup", "1", "--samples", "1"] or argv[5] != "--batch" or not str(argv[6]).isdigit() or int(argv[6]) <= 0:
                raise ValidationError(f"calibration event {event_id} clone argv differs")
            lines = _stdout_lines(event, event_id)
            if not lines or lines[0] != "route,size,sample,nanoseconds,batch,archive_bytes":
                raise ValidationError(f"calibration event {event_id} clone output header differs")
        else:
            binary = _path(_mapping(artifacts.get("probe"), "calibration route artifact").get("path"), "calibration route binary")
            argv = event.get("argv")
            if not isinstance(argv, list) or len(argv) != 9 or argv[:3] != [str(binary), "--route", operation] or argv[3:7] != ["--warmup", "1", "--samples", "1"] or argv[7] != "--batch" or not str(argv[8]).isdigit() or int(argv[8]) <= 0:
                raise ValidationError(f"calibration event {event_id} route argv differs")
            lines = _stdout_lines(event, event_id)
            if not lines or lines[0] != "route,sample,nanoseconds,output_bytes":
                raise ValidationError(f"calibration event {event_id} route output header differs")
    expected_categories = {(role, operation) for role in ROLES for operation in ("clone", *ROUTES)}
    _same(set(categories), expected_categories, "ABBA calibration categories")
    for category, attempts in categories.items():
        _same(sorted(attempts), list(range(1, len(attempts) + 1)), f"ABBA calibration attempts {category}")


def _validate_oracle_stdout(
    event: Mapping[str, Any], *, directory: Path, files: Mapping[str, Any], label: str,
) -> None:
    lines = _stdout_lines(event, label)
    if not lines or lines[0] != f"ole_owner_oracle,{directory}":
        raise ValidationError(f"{label} does not identify its selected output directory")
    records: dict[str, tuple[int, str]] = {}
    for line in lines[1:]:
        fields = line.split("\t")
        if len(fields) != 4 or fields[0] != "oracle_output" or fields[1] in records:
            raise ValidationError(f"{label} oracle output manifest row differs")
        try:
            size = int(fields[2])
        except ValueError as error:
            raise ValidationError(f"{label} oracle output manifest size is not an integer") from error
        records[fields[1]] = (size, fields[3])
    _same(set(records), set(ORACLE_FILES), f"{label} oracle output manifest files")
    for name in ORACLE_FILES:
        record = _mapping(files[name], f"{label}.{name}")
        _same(records[name], (record.get("bytes"), record.get("sha256")), f"{label} oracle output manifest {name}")


def _validate_oracle(
    measurements: Mapping[str, Any], output_root: Path, *, events: Mapping[str, Mapping[str, Any]],
    preparation: Mapping[str, Any], build: Mapping[str, Any], referenced: set[str],
) -> None:
    oracle = _mapping(measurements.get("oracle"), "measurements oracle")
    _same(oracle.get("status"), "preflight-validated", "oracle status")
    _same(oracle.get("claim_scope"), CLAIM_FREE["oracle"], "oracle claim scope")
    reports = _mapping(oracle.get("files"), "oracle files")
    if set(reports) != set(BUILD_KEYS):
        raise ValidationError("oracle endpoint set differs")
    selected: dict[str, Mapping[str, Any]] = {}
    for key in BUILD_KEYS:
        report = _mapping(reports[key], f"oracle.{key}")
        directory = _path(report.get("directory"), f"oracle.{key}.directory")
        expected_directory = (output_root / "oracle" / key.replace("/", "-")).resolve(strict=False)
        _same(directory, expected_directory, f"oracle.{key} selected directory")
        if not directory.is_dir() or {
            path.name for path in directory.iterdir() if path.is_file() and not path.is_symlink()
        } != set(ORACLE_FILES):
            raise ValidationError(f"oracle.{key} directory contains an unexpected artifact")
        event_id = report.get("event_id")
        _same(event_id, f"oracle-{key.replace('/', '-')}", f"oracle.{key} event ID")
        if event_id in referenced:
            raise ValidationError(f"oracle event {event_id} is referenced more than once")
        event = events.get(str(event_id))
        if event is None:
            raise ValidationError(f"oracle.{key} references a missing ABBA event")
        role, variant = key.split("/")
        _event_common(event, event_id=str(event_id), role=role, variant=variant, preparation=preparation, build=build, output_root=output_root)
        files = _mapping(report.get("files"), f"oracle.{key}.files")
        if set(files) != set(ORACLE_FILES):
            raise ValidationError(f"oracle.{key} file set differs")
        for name in ORACLE_FILES:
            record = _mapping(files[name], f"oracle.{key}.{name}")
            _hash_record(record, f"oracle.{key}.{name}")
            record_path = _path(record.get("path"), f"oracle.{key}.{name}")
            if record_path.parent != directory:
                raise ValidationError(f"oracle.{key}.{name} escapes its endpoint directory")
        build_item = _mapping(_mapping(build.get("builds"), "builds").get(key), f"builds.{key}")
        artifacts = _mapping(build_item.get("artifacts"), f"builds.{key}.artifacts")
        artifact = _mapping(artifacts.get("oracle"), f"oracle.{key} binary")
        binary = _path(artifact.get("path"), f"oracle.{key} binary")
        _same(event.get("argv"), [str(binary), "--output", str(directory)], f"oracle.{key} argv")
        outputs = event.get("outputs")
        if not isinstance(outputs, list) or [item.get("path") for item in outputs if isinstance(item, Mapping)] != [str(directory / name) for name in ORACLE_FILES]:
            raise ValidationError(f"oracle.{key} event output paths differ")
        for name, output in zip(ORACLE_FILES, outputs):
            output = _mapping(output, f"oracle.{key} output {name}")
            record = _mapping(files[name], f"oracle.{key}.{name}")
            _same(output.get("path"), record.get("path"), f"oracle.{key} output {name} path")
            _same(output.get("sha256"), record.get("sha256"), f"oracle.{key} output {name} SHA-256")
            _same(output.get("bytes"), record.get("bytes"), f"oracle.{key} output {name} bytes")
            _same(output.get("exists"), True, f"oracle.{key} output {name} existence")
        _validate_oracle_stdout(event, directory=directory, files=files, label=f"oracle.{key}")
        _validate_oracle_proofs(files, f"oracle.{key}")
        selected[key] = files
        allocation = (_path(files["allocator.tsv"]["path"], "oracle allocator")).read_text(encoding="utf-8").splitlines()
        if key.endswith("/normal"):
            _same(allocation, ["status\tnative-uninstrumented"], f"oracle.{key} native allocator marker")
        else:
            _validate_allocator_rows(allocation, f"oracle.{key}.allocator")
        referenced.add(str(event_id))
    for key in ("before/allocator", "after/normal", "after/allocator"):
        for name in ORACLE_FILES[:-1]:
            _same(selected[key][name].get("sha256"), selected["before/normal"][name].get("sha256"), f"oracle cross-endpoint digest {key}/{name}")
    for key in BUILD_KEYS:
        _same(selected[key]["ole-source.bin"].get("sha256"), selected[key]["ole-noop.bin"].get("sha256"), f"oracle no-op digest {key}")


def _validate_oracle_proofs(files: Mapping[str, Any], label: str) -> None:
    metadata = (_path(files["oracle.tsv"]["path"], f"{label}.metadata")).read_text(encoding="utf-8").splitlines()
    _same(metadata, [
        "schema\t1", "owner_streams\t65", f"owner_target\t{TARGET_NAME}",
        "owner_original_bytes\t15", "owner_replacement_bytes\t18", "properties\tvalidated",
        "doc\trevision\tInsertion\t0\t24\tW5Oracle", "xls\tCells\t0\t0\t4242",
    ], f"{label} metadata")
    properties = (_path(files["properties.tsv"]["path"], f"{label}.properties")).read_text(encoding="utf-8").splitlines()
    _same(properties, [
        "link_base\thttps://example.invalid/base/",
        "hyperlinks\t1",
        "hyperlink\t7\t11\t0\thttps://example.invalid/target\tbookmark",
        "opaque_property\t42\t128",
    ], f"{label} property values")
    inventory = (_path(files["ole-inventory.tsv"]["path"], f"{label}.inventory")).read_text(encoding="utf-8").splitlines()
    _same(inventory[0].split("\t") if inventory else [], [
        "kind", "name", "entry_type", "sid", "sid_left", "sid_right", "sid_child",
        "clsid", "state_bits", "creation_time", "modified_time", "size", "payload_hex",
    ], f"{label} inventory header")
    if len(inventory) != 1 + 2 * (64 + 2):
        raise ValidationError(f"{label} inventory cardinality differs")
    rows: dict[str, dict[str, list[str]]] = {"source": {}, "replacement": {}}
    for line in inventory[1:]:
        fields = line.split("\t")
        if len(fields) != 13 or fields[0] not in rows or fields[1] in rows[fields[0]]:
            raise ValidationError(f"{label} inventory row differs")
        rows[fields[0]][fields[1]] = fields[2:]
    root_name = "Root Entry"
    for kind, side in rows.items():
        if root_name not in side or side[root_name][0] != "5":
            raise ValidationError(f"{label} {kind} inventory omits the Root Entry storage identity")
    source_nonstreams = {name for name, fields in rows["source"].items() if fields[0] != "2"}
    replacement_nonstreams = {name for name, fields in rows["replacement"].items() if fields[0] != "2"}
    _same(source_nonstreams, replacement_nonstreams, f"{label} non-stream identity")
    if not source_nonstreams:
        raise ValidationError(f"{label} inventory omits root/storage metadata")
    for name in source_nonstreams:
        _same(rows["source"][name], rows["replacement"][name], f"{label} non-stream metadata {name}")
    expected_streams = {f"ole_w5_sibling_{index:03}" for index in range(64)} | {TARGET_NAME}
    if any(not expected_streams.issubset(side) for side in rows.values()):
        raise ValidationError(f"{label} inventory omits an expected stream")
    for name in expected_streams:
        before = rows["source"][name]
        after = rows["replacement"][name]
        try:
            before_payload = bytes.fromhex(before[-1])
            after_payload = bytes.fromhex(after[-1])
        except ValueError as error:
            raise ValidationError(f"{label} inventory payload is not hex") from error
        if before[-2] != str(len(before_payload)) or after[-2] != str(len(after_payload)):
            raise ValidationError(f"{label} stream size differs from its payload: {name}")
        if name == TARGET_NAME:
            if before_payload != b"ole-w5-original" or after_payload != b"ole-w5-replacement":
                raise ValidationError(f"{label} target inventory proof differs")
        elif before != after or before_payload != f"opaque sibling payload {name[-3:]}".encode():
            raise ValidationError(f"{label} untouched sibling inventory proof differs")
    doc_proof = (_path(files["doc-proof.tsv"]["path"], f"{label}.doc-proof")).read_text(encoding="utf-8").splitlines()
    if len(doc_proof) != 1 + 2 * 256 or doc_proof[0] != "kind\tindex\ttext":
        raise ValidationError(f"{label} DOC proof cardinality differs")
    doc_indices: dict[str, list[int]] = {"source": [], "edit": []}
    for line in doc_proof[1:]:
        fields = line.split("\t")
        if len(fields) != 3 or fields[0] not in {"source", "edit"} or not fields[1].isdigit():
            raise ValidationError(f"{label} DOC proof row differs")
        index = int(fields[1])
        doc_indices[fields[0]].append(index)
        expected = f"ole-w5 tracked paragraph {index:03}"
        if fields[0] == "edit" and index == 0:
            expected = f"ole-w5-tracked-revision {expected}"
        if index >= 256 or fields[2] != expected:
            raise ValidationError(f"{label} DOC untouched paragraph proof differs")
    for kind in ("source", "edit"):
        _same(doc_indices[kind], list(range(256)), f"{label} DOC {kind} indices")
    xls_proof = (_path(files["xls-proof.tsv"]["path"], f"{label}.xls-proof")).read_text(encoding="utf-8").splitlines()
    if len(xls_proof) != 1 + 2 * 256 + 1 or xls_proof[0] != "kind\trow\tvalue_bits":
        raise ValidationError(f"{label} XLS proof cardinality differs")
    xls_indices: dict[str, list[int]] = {"source": [], "edit": []}
    for line in xls_proof[1:-1]:
        fields = line.split("\t")
        if len(fields) != 3 or fields[0] not in {"source", "edit"} or not fields[1].isdigit() or not fields[2].isdigit():
            raise ValidationError(f"{label} XLS proof row differs")
        row = int(fields[1])
        xls_indices[fields[0]].append(row)
        value = 4242.0 if fields[0] == "edit" and row == 0 else float(row)
        expected_bits = struct.unpack("<Q", struct.pack("<d", value))[0]
        if row >= 256 or int(fields[2]) != expected_bits:
            raise ValidationError(f"{label} XLS untouched-cell proof differs")
    for kind in ("source", "edit"):
        _same(xls_indices[kind], list(range(256)), f"{label} XLS {kind} indices")
    _same(xls_proof[-1] if xls_proof else None, "opaque\tbytes\t98304", f"{label} XLS opaque proof")
    opaque_source = _path(files["xls-opaque-source.bin"]["path"], f"{label}.opaque-source")
    opaque_output = _path(files["xls-opaque-edit.bin"]["path"], f"{label}.opaque-edit")
    if opaque_source.stat().st_size != 98304 or opaque_output.read_bytes() != opaque_source.read_bytes():
        raise ValidationError(f"{label} opaque stream bytes differ")
    history = (_path(files["history-proof.tsv"]["path"], f"{label}.history")).read_text(encoding="utf-8").splitlines()
    _same(history, [
        "kind\tpackages\tpatches\tforward\tinverse",
        "doc\t129\t128\tvalidated\tvalidated",
        "xls\t129\t128\tvalidated\tvalidated",
    ], f"{label} history proof")


def _validate_allocator_rows(lines: list[str], label: str) -> None:
    header = lines[0].split("\t") if lines else []
    expected = [
        "route", "alloc_calls", "alloc_bytes", "dealloc_calls", "dealloc_bytes",
        "realloc_calls", "realloc_old_bytes", "realloc_new_bytes", "base_live_calls",
        "base_live_bytes", "live_calls", "live_bytes", "peak_calls", "peak_bytes",
        "retained_packages", "retained_patches",
    ]
    _same(header, expected, f"{label} header")
    if len(lines) != 1 + len(ALLOCATOR_ROUTES):
        raise ValidationError(f"{label} route count differs")
    routes: list[str] = []
    for line in lines[1:]:
        fields = line.split("\t")
        if len(fields) != len(expected) or fields[0] not in ALLOCATOR_ROUTES:
            raise ValidationError(f"{label} route row differs")
        routes.append(fields[0])
        values = {key: int(value) for key, value in zip(expected[1:], fields[1:])}
        if any(value < 0 for value in values.values()):
            raise ValidationError(f"{label} has a negative counter")
        expected_calls = values["base_live_calls"] + values["alloc_calls"] - values["dealloc_calls"]
        expected_bytes = (
            values["base_live_bytes"] + values["alloc_bytes"] - values["dealloc_bytes"]
            - values["realloc_old_bytes"] + values["realloc_new_bytes"]
        )
        _same(values["live_calls"], expected_calls, f"{label} live-call conservation")
        _same(values["live_bytes"], expected_bytes, f"{label} live-byte conservation")
        if values["peak_calls"] < values["live_calls"] or values["peak_bytes"] < values["live_bytes"]:
            raise ValidationError(f"{label} peak is below live state")
        expected_retained = (129, 128) if fields[0] in {"doc_history", "xls_history"} else (0, 0)
        _same((values["retained_packages"], values["retained_patches"]), expected_retained, f"{label} retained history count")
    _same(routes, list(ALLOCATOR_ROUTES), f"{label} route order and uniqueness")


def validate(
    *, manifest_path: Path, preparation_path: Path, build_path: Path,
    abba_path: Path, measurements_path: Path, runner_path: Path,
) -> dict[str, Any]:
    manifest, manifest_hash, _ = _json(manifest_path, "manifest")
    _same(manifest.get("schema"), MANIFEST_SCHEMA, "manifest schema")
    _same(manifest.get("lane"), "w5-ole-owner", "manifest lane")
    _validate_manifest_policy(manifest_path.resolve(), manifest, manifest_hash)
    manifest_allocator = _mapping(manifest.get("allocator"), "manifest allocator")
    _same(manifest_allocator.get("oracle_evidence_file"), "allocator.tsv", "manifest allocator evidence file")
    _same(manifest_allocator.get("oracle_evidence_fields"), list(ALLOCATOR_EVIDENCE_FIELDS), "manifest allocator evidence fields")
    _same(
        manifest_allocator.get("conservation"),
        "live_calls=base_live_calls+alloc_calls-dealloc_calls; live_bytes=base_live_bytes+alloc_bytes-dealloc_bytes-realloc_old_bytes+realloc_new_bytes; peak is at least live; retained history is 129 packages and 128 patches only on history routes",
        "manifest allocator conservation",
    )
    preparation, preparation_hash, _ = _json(preparation_path, "preparation")
    _validate_preparation(preparation, manifest)
    build, _, _ = _json(build_path, "build receipt")
    build_file_hash, build_root, _ = _validate_build(
        manifest, preparation, build, build_path, runner_path.resolve(), manifest_hash, preparation_hash
    )
    abba, _, _ = _json(abba_path, "ABBA receipt")
    _same(abba.get("schema"), SCHEMA, "ABBA receipt schema")
    _same(abba.get("phase"), "abba", "ABBA receipt phase")
    _same(abba.get("status"), "passed", "ABBA receipt status")
    _same(abba.get("claim_scope"), "orchestration and source-bound evidence only; no performance acceptance claim", "ABBA claim scope")
    _same(_mapping(abba.get("manifest"), "ABBA manifest").get("sha256"), manifest_hash, "ABBA manifest identity")
    _same(_mapping(abba.get("preparation"), "ABBA preparation").get("sha256"), preparation_hash, "ABBA preparation identity")
    _same(abba.get("build_receipt_sha256"), build_file_hash, "immutable build receipt hash")
    receipt_runner = _mapping(abba.get("runner"), "ABBA runner")
    _hash_record(receipt_runner, "ABBA runner")
    _same(_path(receipt_runner.get("path"), "ABBA runner"), runner_path.resolve(), "ABBA runner path")
    journal = _mapping(abba.get("event_journal"), "ABBA event journal")
    journal_path = _path(journal.get("path"), "ABBA event journal")
    _same(journal_path, build_root / "events-abba.jsonl", "ABBA event journal path")
    events, digest, size = _events(journal_path, "ABBA event journal")
    _same(journal.get("count"), len(events), "ABBA journal count")
    _same(journal.get("sha256"), _event_hash(events), "ABBA journal semantic hash")
    _same(journal.get("file_sha256"), digest, "ABBA journal file hash")
    _same(journal.get("bytes"), size, "ABBA journal byte count")
    _same(journal.get("event_ids"), [event.get("event_id") for event in events], "ABBA journal event IDs")
    log_paths: set[Path] = set()
    for event in events:
        _same(event.get("phase"), "abba", "ABBA event phase")
        _same(event.get("returncode"), 0, "ABBA event returncode")
        _same(event.get("runner_error"), None, "ABBA event runner error")
        if event.get("process_guard_violations") or event.get("post_exit_descendants") or event.get("post_exit_offenders"):
            raise ValidationError("ABBA event was not quiet")
        _validate_log(event, "stdout", build_root)
        _validate_log(event, "stderr", build_root)
        for field in ("stdout", "stderr"):
            path = _path(_mapping(event.get(field), f"ABBA event {field}").get("path"), f"ABBA event {field}")
            if path in log_paths:
                raise ValidationError(f"ABBA event log path is reused: {path}")
            log_paths.add(path)
    measurements, measurements_hash, measurements_bytes = _json(measurements_path, "measurements")
    _same(_mapping(abba.get("measurements"), "ABBA measurements").get("path"), str(measurements_path.resolve()), "ABBA measurements path")
    _same(_mapping(abba.get("measurements"), "ABBA measurements").get("sha256"), measurements_hash, "ABBA measurements hash")
    _same(_mapping(abba.get("measurements"), "ABBA measurements").get("bytes"), measurements_bytes, "ABBA measurements byte count")
    _same(measurements_path.parent, build_root, "measurements selected run root")
    event_index = _event_index(events)
    referenced = _validate_measurements(
        measurements, measurements_path, events=event_index, preparation=preparation,
        build=build, output_root=build_root,
    )
    _validate_oracle(
        measurements, build_root, events=event_index, preparation=preparation,
        build=build, referenced=referenced,
    )
    if len(referenced) != 4 + BLOCKS * len(ABBA_ORDER) * (1 + len(ROUTES)) + len(ALLOCATOR_ROUTES) * len(ABBA_ORDER):
        raise ValidationError("ABBA referenced event cardinality differs")
    _validate_calibration_events(
        event_index, referenced=referenced, preparation=preparation,
        build=build, output_root=build_root,
    )
    _same(
        [event.get("event_id") for event in events[:4]],
        [f"oracle-{key.replace('/', '-')}" for key in BUILD_KEYS],
        "ABBA oracle event order",
    )
    expected_event_order = [f"oracle-{key.replace('/', '-')}" for key in BUILD_KEYS]
    for role in ROLES:
        expected_event_order.extend(
            event.get("event_id") for event in events
            if str(event.get("event_id", "")).startswith(f"run-cal-{role}-clone-")
        )
    expected_event_order.extend(
        f"run-clone-b{block:02d}-{index}-{role}"
        for block in range(BLOCKS)
        for index, role in enumerate(ABBA_ORDER)
    )
    for route in ROUTES:
        for role in ROLES:
            expected_event_order.extend(
                event.get("event_id") for event in events
                if str(event.get("event_id", "")).startswith(f"run-cal-{role}-{route}-")
            )
    expected_event_order.extend(
        f"run-route-{route}-b{block:02d}-{index}-{role}"
        for route in ROUTES
        for block in range(BLOCKS)
        for index, role in enumerate(ABBA_ORDER)
    )
    expected_event_order.extend(
        f"run-alloc-{route}-{index}-{role}"
        for route in ALLOCATOR_ROUTES
        for index, role in enumerate(ABBA_ORDER)
    )
    _same([event.get("event_id") for event in events], expected_event_order, "ABBA event order")
    return {
        "status": "validated",
        "manifest_sha256": manifest_hash,
        "preparation_sha256": preparation_hash,
        "build_receipt_sha256": build_file_hash,
        "measurements_sha256": measurements_hash,
        "abba_events": len(events),
        "claim_scope": "validation only; no performance acceptance claim",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--preparation", type=Path, required=True)
    parser.add_argument("--build-receipt", type=Path, required=True)
    parser.add_argument("--abba-receipt", type=Path, required=True)
    parser.add_argument("--measurements", type=Path, required=True)
    parser.add_argument("--runner", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        result = validate(
            manifest_path=args.manifest,
            preparation_path=args.preparation,
            build_path=args.build_receipt,
            abba_path=args.abba_receipt,
            measurements_path=args.measurements,
            runner_path=args.runner,
        )
    except (OSError, ValidationError, ValueError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
