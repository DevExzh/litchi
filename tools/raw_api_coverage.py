#!/usr/bin/env python3
"""Audit the public callable surface against an explicit raw-API census.

The glossary gate checks rules that a profile has registered.  This companion
tool answers the preceding question: has every public callable in a supplied
rustdoc profile been registered at all?  It deliberately consumes rustdoc
JSON and exact manifest paths; it never guesses that a name containing
``parse`` or ``codec`` is a raw boundary.

The manifest is source-bound.  Every disposition has an exact source file,
blob, line, and symbol anchor.  Raw/record/codec entries additionally pin the
complete rustdoc input list and the nominal limits type.  Owner and semantic
entries still need an explicit disposition and evidence; they are not a
crate-wide exemption.

The caller supplies the manifest with ``--manifest``; the repository tracks
none.  Normal validation reports an incomplete manifest's state.  The final
``--require-complete`` mode fails unless every expected profile has a rustdoc
input, every public callable has one exact entry, all source identities and
required signatures match, and no active/unclassified/stale disposition
remains.
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass
import hashlib
import json
import os
import re
from pathlib import Path
import stat
import subprocess
import sys
import tomllib
from typing import Any, Iterable, Mapping

# The source-overlay validator must not create an ignored pyc while importing
# its local policy modules.  Final capture receipts also bind this setting.
sys.dont_write_bytecode = True

try:
    from . import check_api_glossary as glossary_gate
    from .rustdoc_return_types import ReturnTypes
except ImportError:  # Direct command-line execution.
    import check_api_glossary as glossary_gate
    from rustdoc_return_types import ReturnTypes


SCHEMA_VERSION = 1
CAPTURE_RECEIPT_SCHEMA = 1
CAPTURE_POLICY_ENV = {
    "CARGO_NET_OFFLINE": "true",
    "CARGO_BUILD_JOBS": "1",
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_DEV_DEBUG": "0",
    "CARGO_PROFILE_TEST_DEBUG": "0",
    "CARGO_PROFILE_RELEASE_DEBUG": "0",
    "RUSTC_BOOTSTRAP": "1",
    "LITCHI_MEM_MAX": "20G",
    "RUST_TEST_THREADS": "1",
    "CARGO_TERM_COLOR": "never",
    "LITCHI_CAPPED": "",
    "LITCHI_GATE_ATTEMPT": "f5",
    "RUSTDOCFLAGS": "",
    "PYTHONDONTWRITEBYTECODE": "1",
}
CAPTURE_POLICY_GUARDS = {
    "target_pre_max_bytes": 18_000_000_000,
    "target_live_max_bytes": 19_000_000_000,
    "target_hard_max_bytes": 20_000_000_000,
    "free_min_bytes": 12_000_000_000,
    "memory_max": "20G",
    "memory_max_bytes": 20 * 1024**3,
    "build_jobs": 1,
    "test_threads": 1,
    "incremental": 0,
    "offline": True,
    "locked": True,
}
CAPTURE_RECEIPT_ENV_KEYS = frozenset({
    *CAPTURE_POLICY_ENV,
    "CARGO_TARGET_DIR",
    "LITCHI_GATE_TARGET_ROOT",
    "LITCHI_GATE_TARGET_TRACE",
    "TMPDIR",
})
CAPTURE_RECEIPT_GUARD_KEYS = frozenset(CAPTURE_POLICY_GUARDS)
CAPTURE_RESOURCE_KEYS = frozenset({
    "target_pre_bytes", "target_live_peak_bytes", "target_post_bytes", "peak_rss_bytes",
    "free_bytes_after", "target_live_guard_triggered",
})
DISPOSITIONS = frozenset({
    "required_raw_limits",
    "configured_owner",
    "retained_owner",
    "semantic_value",
    "fixed_scalar",
    "private_removed",
    "active_lane",
})
RAW_KINDS = frozenset({"raw_function", "record_parser", "codec_function", "raw_method"})
NON_RAW_KINDS = frozenset({"owner_constructor", "semantic_method", "scalar_conversion",
                           "removed_callable"})
HEX40 = re.compile(r"^[0-9a-f]{40}$")
HEX64 = re.compile(r"^[0-9a-f]{64}$")
IDENTIFIER = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
FEATURE = re.compile(r"^[A-Za-z0-9_/-]+$")
PROFILE_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.:/-]*$")


class CoverageError(ValueError):
    """The manifest or a rustdoc input cannot be audited safely."""


@dataclass(frozen=True)
class Profile:
    profile_id: str
    package: str
    default_features: bool
    features: tuple[str, ...]

    @property
    def crate(self) -> str:
        return self.package.replace("-", "_")


@dataclass(frozen=True)
class SourceAnchor:
    path: str
    blob: str
    line: int
    symbol: str
    symbol_kind: str


def _nonempty_string(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise CoverageError(f"{label} must be a nonempty string")
    return value


def _path_string(value: Any, label: str) -> str:
    path = _nonempty_string(value, label)
    path_obj = Path(path)
    if path_obj.is_absolute() or ".." in path_obj.parts:
        raise CoverageError(f"{label} must be a repository-relative path")
    if "3rdparty" in path_obj.parts:
        raise CoverageError(f"{label} must not reference 3rdparty/")
    return path


def _policy_external_path(value: Any, label: str) -> str:
    path = _nonempty_string(value, label)
    path_obj = Path(path)
    if not path_obj.is_absolute():
        raise CoverageError(f"{label} must be an absolute path")
    resolved = path_obj.resolve()
    if resolved == Path("/tmp") or Path("/tmp") in resolved.parents:
        raise CoverageError(f"{label} must not be under /tmp")
    return str(resolved)


def _capture_policy(value: Any, label: str = "coverage.capture_policy") -> dict[str, Any]:
    required = {
        "runner_path", "runner_sha256", "cap_wrapper_path", "cap_wrapper_sha256",
        "underlying_cargo_path", "underlying_cargo_sha256", "environment", "guards",
    }
    if not isinstance(value, dict) or set(value) != required:
        raise CoverageError(f"{label} requires exact runner, wrapper, environment and guard fields")
    result = dict(value)
    for field in ("runner_path", "cap_wrapper_path", "underlying_cargo_path"):
        result[field] = _policy_external_path(result[field], f"{label}.{field}")
    for field in ("runner_sha256", "cap_wrapper_sha256", "underlying_cargo_sha256"):
        digest = _nonempty_string(result[field], f"{label}.{field}").lower()
        if not HEX64.fullmatch(digest):
            raise CoverageError(f"{label}.{field} must be a full lowercase SHA-256")
        result[field] = digest
    environment = result["environment"]
    if not isinstance(environment, dict) or set(environment) != set(CAPTURE_POLICY_ENV):
        raise CoverageError(f"{label}.environment must match the fixed f5 capture environment")
    if any(not isinstance(item, str) for item in environment.values()):
        raise CoverageError(f"{label}.environment values must be strings")
    if environment != CAPTURE_POLICY_ENV:
        raise CoverageError(f"{label}.environment differs from the fixed f5 capture environment")
    result["environment"] = dict(environment)
    guards = result["guards"]
    if not isinstance(guards, dict) or set(guards) != set(CAPTURE_POLICY_GUARDS):
        raise CoverageError(f"{label}.guards must match the fixed resource protocol")
    for field, expected in CAPTURE_POLICY_GUARDS.items():
        if type(guards[field]) is not type(expected) or guards[field] != expected:
            raise CoverageError(f"{label}.guards.{field} differs from the fixed resource protocol")
    result["guards"] = dict(guards)
    return result


def _source_anchor(value: Any, label: str) -> SourceAnchor:
    required = {"path", "blob", "line", "symbol", "symbol_kind"}
    if not isinstance(value, dict) or set(value) != required:
        raise CoverageError(f"{label} requires path, blob, line, symbol and symbol_kind")
    path = _path_string(value["path"], f"{label}.path")
    blob = _nonempty_string(value["blob"], f"{label}.blob").lower()
    if not HEX40.fullmatch(blob):
        raise CoverageError(f"{label}.blob must be a full lowercase git blob SHA-1")
    line = value["line"]
    if type(line) is not int or line < 1:
        raise CoverageError(f"{label}.line must be a positive integer")
    symbol = _nonempty_string(value["symbol"], f"{label}.symbol")
    symbol_kind = value["symbol_kind"]
    if symbol_kind not in {"rust_item", "descriptive"}:
        raise CoverageError(f"{label}.symbol_kind must be rust_item or descriptive")
    if symbol_kind == "rust_item":
        leaf = symbol.rsplit("::", 1)[-1]
        if not IDENTIFIER.fullmatch(leaf):
            raise CoverageError(f"{label}.symbol must end in a Rust identifier for rust_item")
    return SourceAnchor(path, blob, line, symbol, symbol_kind)


def _profile(value: Any, index: int) -> Profile:
    label = f"profiles[{index}]"
    if not isinstance(value, dict):
        raise CoverageError(f"{label} must be an object")
    required = {"id", "package", "default_features", "features"}
    if set(value) != required:
        raise CoverageError(f"{label} requires exactly {sorted(required)}")
    profile_id = _nonempty_string(value["id"], f"{label}.id")
    if not PROFILE_ID.fullmatch(profile_id):
        raise CoverageError(f"{label}.id contains invalid characters")
    package = _nonempty_string(value["package"], f"{label}.package")
    if type(value["default_features"]) is not bool:
        raise CoverageError(f"{label}.default_features must be boolean")
    features = value["features"]
    if not isinstance(features, list) or any(
        not isinstance(feature, str) or not FEATURE.fullmatch(feature)
        for feature in features
    ) or len(set(features)) != len(features):
        raise CoverageError(f"{label}.features must be unique feature identifiers")
    return Profile(profile_id, package, value["default_features"], tuple(features))


def _signature(value: Any, label: str) -> dict[str, Any]:
    required = {"inputs", "limits_argument", "limits_type"}
    if not isinstance(value, dict) or set(value) != required:
        raise CoverageError(f"{label} requires exactly {sorted(required)}")
    inputs = value["inputs"]
    if (not isinstance(inputs, list) or not inputs
            or any(not isinstance(name, str) or not IDENTIFIER.fullmatch(name)
                   for name in inputs)
            or len(set(inputs)) != len(inputs)):
        raise CoverageError(f"{label}.inputs must be unique Rust identifiers")
    limits_argument = value["limits_argument"]
    if (not isinstance(limits_argument, str)
            or limits_argument not in inputs):
        raise CoverageError(f"{label}.limits_argument must name an input")
    limits_type = _nonempty_string(value["limits_type"], f"{label}.limits_type")
    if any(not IDENTIFIER.fullmatch(part) for part in limits_type.split("::")):
        raise CoverageError(f"{label}.limits_type must be a fully qualified Rust path")
    return {"inputs": inputs, "limits_argument": limits_argument, "limits_type": limits_type}


def _entry(value: Any, index: int, profile_ids: set[str]) -> dict[str, Any]:
    label = f"entries[{index}]"
    if not isinstance(value, dict):
        raise CoverageError(f"{label} must be an object")
    required = {"profile", "path", "kind", "disposition", "source", "evidence"}
    if not required <= set(value):
        raise CoverageError(f"{label} requires {sorted(required)}")
    allowed = required | {"required_signature", "lane"}
    if set(value) - allowed:
        raise CoverageError(f"{label} has unknown fields: {sorted(set(value) - allowed)}")
    profile = _nonempty_string(value["profile"], f"{label}.profile")
    if profile not in profile_ids:
        raise CoverageError(f"{label}.profile {profile!r} is not declared")
    path = _nonempty_string(value["path"], f"{label}.path")
    if any(not IDENTIFIER.fullmatch(part) for part in path.split("::")):
        raise CoverageError(f"{label}.path must be a fully qualified Rust path")
    kind = _nonempty_string(value["kind"], f"{label}.kind")
    disposition = value["disposition"]
    if disposition not in DISPOSITIONS:
        raise CoverageError(f"{label}.disposition {disposition!r} is not allowed")
    anchor = _source_anchor(value["source"], f"{label}.source")
    if anchor.symbol_kind != "rust_item":
        raise CoverageError(f"{label}.source.symbol_kind must be rust_item for callable entries")
    _nonempty_string(value["evidence"], f"{label}.evidence")
    if disposition == "required_raw_limits":
        if kind not in RAW_KINDS:
            raise CoverageError(f"{label}: raw limits require a raw/record/codec kind")
        if "required_signature" not in value:
            raise CoverageError(f"{label}: required_raw_limits needs required_signature")
        _signature(value["required_signature"], f"{label}.required_signature")
    elif "required_signature" in value:
        raise CoverageError(f"{label}: only required_raw_limits may carry required_signature")
    if disposition in {"configured_owner", "retained_owner"} and kind not in NON_RAW_KINDS:
        raise CoverageError(f"{label}: owner disposition needs an owner kind")
    if disposition == "semantic_value" and kind not in {"semantic_method", "scalar_conversion"}:
        raise CoverageError(f"{label}: semantic_value needs a semantic kind")
    if disposition == "fixed_scalar" and kind != "scalar_conversion":
        raise CoverageError(f"{label}: fixed_scalar needs scalar_conversion kind")
    if disposition == "private_removed" and kind != "removed_callable":
        raise CoverageError(f"{label}: private_removed needs removed_callable kind")
    if disposition == "active_lane":
        lane = value.get("lane")
        if not isinstance(lane, str) or not lane.strip():
            raise CoverageError(f"{label}: active_lane needs a nonempty lane")
    elif "lane" in value:
        raise CoverageError(f"{label}: lane is only valid for active_lane")
    return value


def _lane(value: Any, index: int) -> dict[str, Any]:
    label = f"lanes[{index}]"
    if not isinstance(value, dict):
        raise CoverageError(f"{label} must be an object")
    required = {"id", "package", "scope", "disposition", "source", "evidence"}
    if set(value) != required:
        raise CoverageError(f"{label} requires exactly {sorted(required)}")
    _nonempty_string(value["id"], f"{label}.id")
    _nonempty_string(value["package"], f"{label}.package")
    _nonempty_string(value["scope"], f"{label}.scope")
    if value["disposition"] != "active_lane":
        raise CoverageError(f"{label}.disposition must be active_lane")
    _source_anchor(value["source"], f"{label}.source")
    _nonempty_string(value["evidence"], f"{label}.evidence")
    return value


def load_manifest(path: Path) -> dict[str, Any]:
    try:
        manifest = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise CoverageError(f"cannot read manifest {path}: {error}") from error
    if not isinstance(manifest, dict):
        raise CoverageError("manifest must be an object")
    required = {"schema_version", "source_commit", "coverage", "profiles", "entries", "lanes"}
    if set(manifest) != required:
        raise CoverageError(f"manifest requires exactly {sorted(required)}")
    if manifest["schema_version"] != SCHEMA_VERSION:
        raise CoverageError(f"unsupported manifest schema {manifest['schema_version']!r}")
    source_commit = _nonempty_string(manifest["source_commit"], "source_commit").lower()
    if not HEX40.fullmatch(source_commit):
        raise CoverageError("source_commit must be a full lowercase commit SHA-1")
    coverage = manifest["coverage"]
    coverage_fields = {
        "state", "reason", "expected_profile_source", "expected_profile_source_blob",
        "expected_workspace_source", "expected_workspace_source_blob",
        "approved_cargo_lock_sha256", "capture_policy",
    }
    if not isinstance(coverage, dict) or set(coverage) != coverage_fields:
        raise CoverageError(
            "coverage requires state, reason, expected source paths and trusted blobs"
        )
    if coverage["state"] not in {"incomplete", "complete"}:
        raise CoverageError("coverage.state must be incomplete or complete")
    _nonempty_string(coverage["reason"], "coverage.reason")
    _path_string(coverage["expected_profile_source"], "coverage.expected_profile_source")
    profile_blob = _nonempty_string(
        coverage["expected_profile_source_blob"], "coverage.expected_profile_source_blob"
    ).lower()
    if not HEX40.fullmatch(profile_blob):
        raise CoverageError("coverage.expected_profile_source_blob must be a full lowercase SHA-1")
    coverage["expected_profile_source_blob"] = profile_blob
    _path_string(coverage["expected_workspace_source"], "coverage.expected_workspace_source")
    workspace_blob = _nonempty_string(
        coverage["expected_workspace_source_blob"], "coverage.expected_workspace_source_blob"
    ).lower()
    if not HEX40.fullmatch(workspace_blob):
        raise CoverageError("coverage.expected_workspace_source_blob must be a full lowercase SHA-1")
    coverage["expected_workspace_source_blob"] = workspace_blob
    lock_blob = _nonempty_string(
        coverage["approved_cargo_lock_sha256"], "coverage.approved_cargo_lock_sha256"
    ).lower()
    if not HEX64.fullmatch(lock_blob):
        raise CoverageError("coverage.approved_cargo_lock_sha256 must be a full lowercase SHA-256")
    coverage["approved_cargo_lock_sha256"] = lock_blob
    coverage["capture_policy"] = _capture_policy(coverage["capture_policy"])
    profiles = manifest["profiles"]
    if not isinstance(profiles, list):
        raise CoverageError("profiles must be an array")
    parsed_profiles = [_profile(value, index) for index, value in enumerate(profiles)]
    profile_ids = {profile.profile_id for profile in parsed_profiles}
    if len(profile_ids) != len(parsed_profiles):
        raise CoverageError("profiles contain duplicate ids")
    entries = manifest["entries"]
    if not isinstance(entries, list):
        raise CoverageError("entries must be an array")
    parsed_entries = [_entry(value, index, profile_ids) for index, value in enumerate(entries)]
    identities = [(entry["profile"], entry["path"]) for entry in parsed_entries]
    if len(set(identities)) != len(identities):
        raise CoverageError("entries contain duplicate profile/path identities")
    lanes = manifest["lanes"]
    if not isinstance(lanes, list):
        raise CoverageError("lanes must be an array")
    [_lane(value, index) for index, value in enumerate(lanes)]
    lane_ids = {lane["id"] for lane in lanes}
    if len(lane_ids) != len(lanes):
        raise CoverageError("lanes contain duplicate ids")
    for index, entry in enumerate(parsed_entries):
        if entry["disposition"] == "active_lane" and entry["lane"] not in lane_ids:
            raise CoverageError(
                f"entries[{index}] active lane {entry['lane']!r} is not declared in lanes"
            )
    manifest["source_commit"] = source_commit
    manifest["_profiles"] = {profile.profile_id: profile for profile in parsed_profiles}
    manifest["_entries"] = {(entry["profile"], entry["path"]): entry for entry in parsed_entries}
    return manifest


def _git_type(root: Path, object_id: str) -> str | None:
    result = subprocess.run(
        ["git", "cat-file", "-t", object_id],
        cwd=root, check=False, capture_output=True, text=True,
    )
    object_type = result.stdout.strip()
    return object_type if result.returncode == 0 and object_type else None


def _git_object(root: Path, commit: str, relative_path: str) -> tuple[str, str] | None:
    result = subprocess.run(
        ["git", "rev-parse", "--verify", f"{commit}:{relative_path}"],
        cwd=root, check=False, capture_output=True, text=True,
    )
    if result.returncode != 0:
        return None
    object_id = result.stdout.strip()
    if not HEX40.fullmatch(object_id):
        return None
    object_type = _git_type(root, object_id)
    return (object_id, object_type) if object_type is not None else None


def _git_blob(root: Path, commit: str, relative_path: str) -> str | None:
    resolved = _git_object(root, commit, relative_path)
    if resolved is None or resolved[1] != "blob":
        return None
    return resolved[0]


def _git_blob_bytes(root: Path, blob: str) -> bytes | None:
    result = subprocess.run(
        ["git", "cat-file", "blob", blob],
        cwd=root, check=False, capture_output=True,
    )
    return result.stdout if result.returncode == 0 else None


def _anchor_failure(root: Path, commit: str, label: str, anchor: Mapping[str, Any]) -> str | None:
    resolved = _git_object(root, commit, anchor["path"])
    if resolved is None:
        return f"{label} source {anchor['path']} is missing at {commit}"
    object_id, object_type = resolved
    if object_type != "blob":
        return f"{label} source {anchor['path']} resolves to {object_type}, not a blob"
    if object_id != anchor["blob"]:
        return (
            f"{label} source {anchor['path']} blob {anchor['blob']} != "
            f"{object_id} at {commit}"
        )
    content = _git_blob_bytes(root, object_id)
    if content is None:
        return f"{label} source {anchor['path']} blob {object_id} cannot be read"
    lines = content.splitlines()
    if anchor["line"] > len(lines):
        return (
            f"{label} source {anchor['path']} line {anchor['line']} exceeds "
            f"the {len(lines)}-line blob at {commit}"
        )
    if anchor["symbol_kind"] == "rust_item":
        line = lines[anchor["line"] - 1].decode("utf-8", errors="replace")
        leaf = anchor["symbol"].rsplit("::", 1)[-1]
        declaration = (
            rf"\b(?:fn|struct|enum|trait|type|const|static|mod)\s+"
            rf"{re.escape(leaf)}\b"
        )
        if not re.search(declaration, line):
            return (
                f"{label} source {anchor['path']} line {anchor['line']} does not "
                f"contain rust item {anchor['symbol']!r}"
            )
    return None


def _git_head(root: Path) -> str | None:
    result = subprocess.run(
        ["git", "rev-parse", "--verify", "HEAD^{commit}"],
        cwd=root, check=False, capture_output=True, text=True,
    )
    head = result.stdout.strip()
    return head if HEX40.fullmatch(head) else None


def _git_tree(root: Path, commit: str) -> str | None:
    result = subprocess.run(
        ["git", "rev-parse", "--verify", f"{commit}^{{tree}}"],
        cwd=root, check=False, capture_output=True, text=True,
    )
    tree = result.stdout.strip()
    return tree if HEX40.fullmatch(tree) else None


def _sha256_file(path: Path) -> str | None:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError:
        return None


def _regular_file_with_components(path: Path) -> bool:
    """Require a regular file and reject symlinks in every path component."""
    absolute = path.absolute()
    current = Path(absolute.anchor)
    components = absolute.parts[1:]
    for component in components:
        current /= component
        try:
            mode = current.lstat().st_mode
        except OSError:
            return False
        if stat.S_ISLNK(mode):
            return False
        if current != absolute and not stat.S_ISDIR(mode):
            return False
    try:
        return stat.S_ISREG(absolute.lstat().st_mode)
    except OSError:
        return False


def _source_blob_failure(root: Path, commit: str, label: str, path: str,
                         expected_blob: str) -> str | None:
    resolved = _git_object(root, commit, path)
    if resolved is None:
        return f"{label} source {path} is missing at {commit}"
    object_id, object_type = resolved
    if object_type != "blob":
        return f"{label} source {path} resolves to {object_type}, not a blob"
    if object_id != expected_blob:
        return f"{label} source {path} blob {expected_blob} != {object_id} at {commit}"
    return None


def _source_overlay_failures(root: Path, cargo_lock_sha256: str | None) -> list[str]:
    """Reject every checkout overlay except the exact seeded root lockfile."""
    result = subprocess.run(
        ["git", "status", "--porcelain=v1", "--untracked-files=all", "--ignored=traditional"],
        cwd=root, check=False, capture_output=True, text=True,
    )
    if result.returncode != 0:
        return [f"cannot inspect source checkout overlays: {result.stderr.strip() or 'git status failed'}"]
    failures: list[str] = []
    for raw in result.stdout.splitlines():
        if len(raw) < 3:
            continue
        state, raw_path = raw[:2], raw[3:]
        path = raw_path.strip()
        if path.startswith('"') and path.endswith('"'):
            try:
                path = json.loads(path)
            except json.JSONDecodeError:
                pass
        if path == "Cargo.lock" and cargo_lock_sha256 is not None:
            lock = root / path
            actual = _sha256_file(lock)
            if _regular_file_with_components(lock) and actual == cargo_lock_sha256:
                continue
        failures.append(f"source checkout overlay {path!r} ({state})")
        candidate = root / path.rstrip("/")
        if candidate.is_dir() and not candidate.is_symlink():
            try:
                children = sorted(candidate.rglob("*"))
            except OSError:
                children = []
            for child in children:
                try:
                    relative = child.relative_to(root)
                except ValueError:
                    continue
                failures.append(f"source checkout overlay {relative!s} ({state})")
    return failures


def _external_path(value: str, label: str, root: Path) -> Path:
    path = Path(value)
    if not path.is_absolute():
        raise CoverageError(f"{label} must be absolute")
    resolved = path.resolve()
    source_root = root.resolve()
    try:
        resolved.relative_to(source_root)
    except ValueError:
        pass
    else:
        raise CoverageError(f"{label} must be outside the source root")
    if resolved == Path("/tmp") or Path("/tmp") in resolved.parents:
        raise CoverageError(f"{label} must not be under /tmp")
    return resolved


def _capture_receipt(value: Any, label: str) -> dict[str, Any]:
    required = {
        "schema", "profile", "package", "default_features", "features", "argv",
        "source_commit", "source_tree", "source_root", "target_dir",
        "cargo_lock_path", "cargo_lock_sha256", "rust_toolchain_path",
        "rust_toolchain_sha256", "cargo_path", "cargo_sha256", "rustc_path",
        "rustc_sha256", "artifact", "runner_path", "runner_sha256",
        "cap_wrapper_path", "cap_wrapper_sha256", "underlying_cargo_path",
        "underlying_cargo_sha256", "capture_env", "guards", "resource",
        "execution_trace",
    }
    if not isinstance(value, dict) or set(value) != required:
        raise CoverageError(f"{label} requires exact capture identity fields")
    if value["schema"] != CAPTURE_RECEIPT_SCHEMA:
        raise CoverageError(f"{label}.schema is unsupported")
    for field in ("profile", "package", "source_root", "target_dir", "cargo_path", "rustc_path"):
        _nonempty_string(value[field], f"{label}.{field}")
    if type(value["default_features"]) is not bool:
        raise CoverageError(f"{label}.default_features must be boolean")
    features = value["features"]
    if not isinstance(features, list) or any(
        not isinstance(feature, str) or not FEATURE.fullmatch(feature)
        for feature in features
    ) or len(set(features)) != len(features):
        raise CoverageError(f"{label}.features must be unique feature identifiers")
    argv = value["argv"]
    if not isinstance(argv, list) or not argv or any(not isinstance(item, str) or not item for item in argv):
        raise CoverageError(f"{label}.argv must be a nonempty string array")
    for field in ("source_commit", "source_tree", "cargo_lock_sha256", "rust_toolchain_sha256",
                  "cargo_sha256", "rustc_sha256"):
        digest = _nonempty_string(value[field], f"{label}.{field}").lower()
        pattern = HEX40 if field in {"source_commit", "source_tree"} else HEX64
        if not pattern.fullmatch(digest):
            raise CoverageError(f"{label}.{field} has the wrong digest shape")
        value[field] = digest
    if value["cargo_lock_path"] != "Cargo.lock":
        raise CoverageError(f"{label}.cargo_lock_path must be exactly Cargo.lock")
    if value["rust_toolchain_path"] not in {"rust-toolchain.toml", "rust-toolchain"}:
        raise CoverageError(f"{label}.rust_toolchain_path is not an approved root toolchain path")
    for field in ("runner_path", "cap_wrapper_path", "underlying_cargo_path"):
        _nonempty_string(value[field], f"{label}.{field}")
    for field in ("runner_sha256", "cap_wrapper_sha256", "underlying_cargo_sha256"):
        digest = _nonempty_string(value[field], f"{label}.{field}").lower()
        if not HEX64.fullmatch(digest):
            raise CoverageError(f"{label}.{field} must be a full lowercase SHA-256")
        value[field] = digest
    capture_env = value["capture_env"]
    if not isinstance(capture_env, dict) or set(capture_env) != CAPTURE_RECEIPT_ENV_KEYS:
        raise CoverageError(f"{label}.capture_env has an incomplete or extra environment identity")
    if any(not isinstance(item, str) for item in capture_env.values()):
        raise CoverageError(f"{label}.capture_env values must be strings")
    guards = value["guards"]
    if not isinstance(guards, dict) or set(guards) != CAPTURE_RECEIPT_GUARD_KEYS:
        raise CoverageError(f"{label}.guards has an incomplete or extra resource identity")
    for field, expected in CAPTURE_POLICY_GUARDS.items():
        if type(guards[field]) is not type(expected):
            raise CoverageError(f"{label}.guards.{field} has the wrong type")
    resource = value["resource"]
    if not isinstance(resource, dict) or set(resource) != CAPTURE_RESOURCE_KEYS:
        raise CoverageError(f"{label}.resource requires exact target, RSS and free-space fields")
    for field in CAPTURE_RESOURCE_KEYS - {"target_live_guard_triggered"}:
        if type(resource[field]) is not int or resource[field] < 0:
            raise CoverageError(f"{label}.resource.{field} must be a nonnegative integer")
    if type(resource["target_live_guard_triggered"]) is not bool:
        raise CoverageError(f"{label}.resource.target_live_guard_triggered must be boolean")
    trace = value["execution_trace"]
    if not isinstance(trace, dict) or set(trace) != {"path", "bytes", "sha256", "rows"}:
        raise CoverageError(f"{label}.execution_trace requires path, bytes, sha256 and rows")
    _nonempty_string(trace["path"], f"{label}.execution_trace.path")
    if type(trace["bytes"]) is not int or trace["bytes"] < 1:
        raise CoverageError(f"{label}.execution_trace.bytes must be positive")
    trace_hash = _nonempty_string(trace["sha256"], f"{label}.execution_trace.sha256").lower()
    if not HEX64.fullmatch(trace_hash):
        raise CoverageError(f"{label}.execution_trace.sha256 must be a full lowercase SHA-256")
    trace["sha256"] = trace_hash
    rows = trace["rows"]
    if not isinstance(rows, list) or not rows:
        raise CoverageError(f"{label}.execution_trace.rows must be nonempty")
    for index, row in enumerate(rows):
        if not isinstance(row, dict) or set(row) != {"pid", "command", "target"}:
            raise CoverageError(f"{label}.execution_trace.rows[{index}] requires pid, command and target")
        if not isinstance(row["pid"], str) or not row["pid"].isdigit() or int(row["pid"]) < 1:
            raise CoverageError(f"{label}.execution_trace.rows[{index}].pid is invalid")
        for field in ("command", "target"):
            _nonempty_string(row[field], f"{label}.execution_trace.rows[{index}].{field}")
    artifact = value["artifact"]
    if not isinstance(artifact, dict) or set(artifact) != {"path", "bytes", "sha256"}:
        raise CoverageError(f"{label}.artifact requires path, bytes and sha256")
    _nonempty_string(artifact["path"], f"{label}.artifact.path")
    if type(artifact["bytes"]) is not int or artifact["bytes"] < 1:
        raise CoverageError(f"{label}.artifact.bytes must be positive")
    artifact_hash = _nonempty_string(artifact["sha256"], f"{label}.artifact.sha256").lower()
    if not HEX64.fullmatch(artifact_hash):
        raise CoverageError(f"{label}.artifact.sha256 must be a full lowercase SHA-256")
    artifact["sha256"] = artifact_hash
    return value


def _validate_capture_receipt(manifest: Mapping[str, Any], profile: Profile,
                              rustdoc_path: Path, receipt_path: Path, root: Path,
                              receipt: Mapping[str, Any],
                              overlay_failures: list[str]) -> list[str]:
    failures: list[str] = []
    try:
        source_root = Path(receipt["source_root"]).resolve()
        if source_root != root.resolve():
            failures.append(f"{profile.profile_id}: receipt source_root {source_root} != {root.resolve()}")
        target = _external_path(str(receipt["target_dir"]), f"{profile.profile_id}.target_dir", root)
        if not target.is_dir():
            failures.append(f"{profile.profile_id}: receipt target_dir is not a directory: {target}")
        artifact = _external_path(str(receipt["artifact"]["path"]), f"{profile.profile_id}.artifact.path", root)
        if artifact != rustdoc_path.resolve():
            failures.append(f"{profile.profile_id}: receipt artifact path {artifact} != {rustdoc_path.resolve()}")
        _external_path(str(receipt_path.resolve()), f"{profile.profile_id}.receipt_path", root)
        cargo_path = _external_path(str(receipt["cargo_path"]), f"{profile.profile_id}.cargo_path", root)
        rustc_path = _external_path(str(receipt["rustc_path"]), f"{profile.profile_id}.rustc_path", root)
        runner_path = _external_path(str(receipt["runner_path"]), f"{profile.profile_id}.runner_path", root)
        cap_wrapper_path = _external_path(
            str(receipt["cap_wrapper_path"]), f"{profile.profile_id}.cap_wrapper_path", root
        )
        underlying_cargo_path = _external_path(
            str(receipt["underlying_cargo_path"]),
            f"{profile.profile_id}.underlying_cargo_path", root,
        )
        trace_path = _external_path(
            str(receipt["execution_trace"]["path"]),
            f"{profile.profile_id}.execution_trace.path", root,
        )
        tmpdir = _external_path(str(receipt["capture_env"]["TMPDIR"]), f"{profile.profile_id}.TMPDIR", root)
    except CoverageError as error:
        failures.append(f"{profile.profile_id}: {error}")
        return failures
    if receipt["profile"] != profile.profile_id:
        failures.append(f"{profile.profile_id}: receipt profile is {receipt['profile']!r}")
    if receipt["package"] != profile.package:
        failures.append(f"{profile.profile_id}: receipt package is {receipt['package']!r}")
    if receipt["default_features"] is not profile.default_features:
        failures.append(f"{profile.profile_id}: receipt default_features differs from manifest")
    if tuple(receipt["features"]) != profile.features:
        failures.append(f"{profile.profile_id}: receipt features differ from manifest")
    policy = manifest["coverage"]["capture_policy"]
    for field in ("runner_path", "cap_wrapper_path", "underlying_cargo_path"):
        if str(Path(receipt[field]).resolve()) != policy[field]:
            failures.append(f"{profile.profile_id}: receipt {field} differs from the approved f5 runner")
    for field in ("runner_sha256", "cap_wrapper_sha256", "underlying_cargo_sha256"):
        if receipt[field] != policy[field]:
            failures.append(f"{profile.profile_id}: receipt {field} differs from the approved f5 runner")
    if str(cargo_path) != str(cap_wrapper_path):
        failures.append(f"{profile.profile_id}: cargo_path is not the f5 cap wrapper")
    if not _regular_file_with_components(runner_path):
        failures.append(f"{profile.profile_id}: runner is not a regular file: {runner_path}")
    for label, path in (("cap wrapper", cap_wrapper_path), ("underlying Cargo", underlying_cargo_path)):
        if not _regular_file_with_components(path) or not os.access(path, os.X_OK):
            failures.append(f"{profile.profile_id}: {label} is not a regular executable: {path}")
    if not _regular_file_with_components(trace_path):
        failures.append(f"{profile.profile_id}: execution trace is not a regular file: {trace_path}")
    for label, path, expected in (
        ("cargo", cargo_path, receipt["cargo_sha256"]),
        ("rustc", rustc_path, receipt["rustc_sha256"]),
        ("runner", runner_path, receipt["runner_sha256"]),
        ("cap wrapper", cap_wrapper_path, receipt["cap_wrapper_sha256"]),
        ("underlying Cargo", underlying_cargo_path, receipt["underlying_cargo_sha256"]),
    ):
        actual = _sha256_file(path)
        if actual != expected:
            failures.append(f"{profile.profile_id}: {label} identity differs from receipt")
    try:
        wrapper_source = cap_wrapper_path.read_text(encoding="utf-8")
        underlying_source = underlying_cargo_path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as error:
        failures.append(f"{profile.profile_id}: cannot read f5 cap wrapper identity: {error}")
    else:
        trace_printf = "printf '%s\\t%s\\t%s\\n' \"$$\" \"$command\" \"$requested_target\""
        if ("LITCHI_GATE_TARGET_TRACE" not in wrapper_source
                or "requested_target" not in wrapper_source
                or trace_printf not in wrapper_source):
            failures.append(f"{profile.profile_id}: cap wrapper does not publish target trace rows")
        if "systemd-run" not in underlying_source or "MemoryMax=${LITCHI_MEM_MAX" not in underlying_source:
            failures.append(f"{profile.profile_id}: underlying Cargo is not the memory-capped shim")
    expected_argv = [
        str(cargo_path), "rustdoc", "--package", profile.package, "--lib",
        "--target-dir", str(Path(receipt["target_dir"]).resolve()),
    ]
    if not profile.default_features:
        expected_argv.append("--no-default-features")
    if profile.features:
        expected_argv += ["--features", ",".join(profile.features)]
    expected_argv += [
        "--offline", "--locked", "--", "-Zunstable-options", "--output-format", "json",
        "--document-hidden-items",
    ]
    if receipt["argv"] != expected_argv:
        failures.append(f"{profile.profile_id}: receipt argv does not exactly match profile capture")
    expected_env = dict(policy["environment"])
    expected_env.update({
        "CARGO_TARGET_DIR": str(target),
        "LITCHI_GATE_TARGET_ROOT": str(target),
        "LITCHI_GATE_TARGET_TRACE": str(trace_path),
        "TMPDIR": str(tmpdir),
    })
    if receipt["capture_env"] != expected_env:
        failures.append(f"{profile.profile_id}: capture environment differs from the f5 policy")
    if receipt["guards"] != policy["guards"]:
        failures.append(f"{profile.profile_id}: resource guards differ from the f5 policy")
    resource = receipt["resource"]
    guards = policy["guards"]
    if resource["target_pre_bytes"] > guards["target_pre_max_bytes"]:
        failures.append(f"{profile.profile_id}: target preflight exceeded the f5 guard")
    if resource["target_live_peak_bytes"] > guards["target_live_max_bytes"]:
        failures.append(f"{profile.profile_id}: target live guard exceeded the f5 limit")
    if resource["target_live_guard_triggered"]:
        failures.append(f"{profile.profile_id}: target live guard reported an overflow")
    if resource["target_post_bytes"] > guards["target_hard_max_bytes"]:
        failures.append(f"{profile.profile_id}: target post-size exceeded the f5 hard cap")
    if resource["peak_rss_bytes"] < 1 or resource["peak_rss_bytes"] >= guards["memory_max_bytes"]:
        failures.append(f"{profile.profile_id}: peak RSS exceeded or omitted the f5 memory cap")
    if resource["free_bytes_after"] < guards["free_min_bytes"]:
        failures.append(f"{profile.profile_id}: free-space receipt is below the f5 guard")
    try:
        trace_bytes = trace_path.read_bytes()
        actual_trace_rows = []
        for line in trace_bytes.decode("utf-8").splitlines():
            fields = line.split("\t")
            if len(fields) != 3 or not fields[0].isdigit() or int(fields[0]) < 1:
                raise ValueError("trace row has invalid pid or field count")
            actual_trace_rows.append({"pid": fields[0], "command": fields[1], "target": fields[2]})
    except (OSError, UnicodeDecodeError, ValueError) as error:
        failures.append(f"{profile.profile_id}: invalid f5 execution trace: {error}")
    else:
        trace = receipt["execution_trace"]
        if len(trace_bytes) != trace["bytes"] or hashlib.sha256(trace_bytes).hexdigest() != trace["sha256"]:
            failures.append(f"{profile.profile_id}: execution trace bytes/SHA differs from receipt")
        if actual_trace_rows != trace["rows"]:
            failures.append(f"{profile.profile_id}: execution trace rows differ from receipt")
        if any(row["command"] != "rustdoc" or row["target"] != str(target) for row in actual_trace_rows):
            failures.append(f"{profile.profile_id}: execution trace does not match the cap wrapper rustdoc target")
    head = _git_head(root)
    approved_lock = manifest["coverage"]["approved_cargo_lock_sha256"]
    if receipt["cargo_lock_sha256"] != approved_lock:
        failures.append(f"{profile.profile_id}: receipt Cargo.lock is not the approved seeded lock")
    if receipt["source_commit"] != manifest["source_commit"] or receipt["source_commit"] != head:
        failures.append(f"{profile.profile_id}: receipt source_commit is stale or differs from frozen HEAD")
    tree = _git_tree(root, receipt["source_commit"])
    if tree != receipt["source_tree"]:
        failures.append(f"{profile.profile_id}: receipt source_tree differs from source commit")
    lock = root / receipt["cargo_lock_path"]
    if not _regular_file_with_components(lock):
        failures.append(f"{profile.profile_id}: approved Cargo.lock is not a regular non-symlink path")
    if _sha256_file(lock) != receipt["cargo_lock_sha256"]:
        failures.append(f"{profile.profile_id}: seeded Cargo.lock identity differs from receipt")
    selected_toolchain = next(
        (root / name for name in ("rust-toolchain.toml", "rust-toolchain") if (root / name).is_file()),
        None,
    )
    if selected_toolchain is None or receipt["rust_toolchain_path"] != selected_toolchain.name:
        failures.append(f"{profile.profile_id}: receipt toolchain path is not the selected root pin")
    elif _sha256_file(selected_toolchain) != receipt["rust_toolchain_sha256"]:
        failures.append(f"{profile.profile_id}: rust-toolchain identity differs from receipt")
    artifact_path = rustdoc_path.resolve()
    artifact = receipt["artifact"]
    actual_bytes = artifact_path.stat().st_size if artifact_path.is_file() else None
    actual_hash = _sha256_file(artifact_path)
    if actual_bytes != artifact["bytes"] or actual_hash != artifact["sha256"]:
        failures.append(f"{profile.profile_id}: rustdoc artifact bytes/SHA differs from receipt")
    failures.extend(overlay_failures)
    return failures


def verify_source_anchors(manifest: Mapping[str, Any], root: Path, *, current: bool = False) -> list[str]:
    failures: list[str] = []
    commit = _git_head(root) if current else manifest["source_commit"]
    if commit is None:
        return ["cannot resolve a current HEAD for source-anchor verification"]
    if not current and _git_type(root, commit) != "commit":
        return [f"source_commit {commit} does not resolve to a Git commit"]
    coverage = manifest["coverage"]
    for path_key, blob_key in (
        ("expected_profile_source", "expected_profile_source_blob"),
        ("expected_workspace_source", "expected_workspace_source_blob"),
    ):
        failure = _source_blob_failure(
            root, commit, f"coverage.{path_key}", coverage[path_key], coverage[blob_key]
        )
        if failure:
            failures.append(failure)
    for index, entry in enumerate(manifest["entries"]):
        anchor = entry["source"]
        failure = _anchor_failure(root, commit, f"entries[{index}] {entry['path']}", anchor)
        if failure:
            failures.append(failure)
    for index, lane in enumerate(manifest["lanes"]):
        anchor = lane["source"]
        failure = _anchor_failure(root, commit, f"lanes[{index}] {lane['id']}", anchor)
        if failure:
            failures.append(failure)
    return failures


def _glossary_profiles(path: Path) -> dict[str, Profile]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise CoverageError(f"cannot read glossary {path}: {error}") from error
    if not isinstance(value, dict) or not isinstance(value.get("profiles"), list):
        raise CoverageError("glossary must contain a profiles array")
    profiles: dict[str, Profile] = {}
    for index, raw in enumerate(value["profiles"]):
        if not isinstance(raw, dict):
            raise CoverageError(f"glossary profiles[{index}] is not an object")
        profile = Profile(
            _nonempty_string(raw.get("id"), f"glossary profiles[{index}].id"),
            _nonempty_string(raw.get("package"), f"glossary profiles[{index}].package"),
            raw.get("default_features") is True,
            tuple(raw.get("features", [])),
        )
        if profile.profile_id in profiles:
            raise CoverageError(f"glossary contains duplicate profile {profile.profile_id}")
        profiles[profile.profile_id] = profile
    return profiles


def _workspace_packages(root: Path) -> set[str]:
    """Enumerate direct library packages without invoking Cargo."""
    packages: set[str] = set()
    crates = root / "crates"
    if not crates.is_dir():
        return packages
    for manifest in sorted(crates.glob("*/Cargo.toml")):
        try:
            document = tomllib.loads(manifest.read_text(encoding="utf-8"))
        except (OSError, tomllib.TOMLDecodeError) as error:
            raise CoverageError(f"cannot read {manifest}: {error}") from error
        package = document.get("package")
        if not isinstance(package, dict) or not isinstance(package.get("name"), str):
            continue
        lib = document.get("lib")
        lib_path = manifest.parent / (lib.get("path", "src/lib.rs") if isinstance(lib, dict) else "src/lib.rs")
        if lib_path.is_file():
            packages.add(package["name"])
    return packages


def _parse_rustdoc(path: Path) -> dict[str, Any]:
    try:
        document = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise CoverageError(f"cannot read rustdoc JSON {path}: {error}") from error
    if not isinstance(document, dict):
        raise CoverageError(f"rustdoc JSON {path} must be an object")
    return document


def parse_rustdoc_args(values: Iterable[str]) -> dict[str, Path]:
    result: dict[str, Path] = {}
    for value in values:
        if "=" not in value:
            raise CoverageError(f"--rustdoc expects PROFILE=PATH, got {value!r}")
        profile, path = value.split("=", 1)
        if not PROFILE_ID.fullmatch(profile) or not path:
            raise CoverageError(f"invalid --rustdoc binding {value!r}")
        if profile in result:
            raise CoverageError(f"duplicate rustdoc profile {profile}")
        result[profile] = Path(path)
    return result


def parse_receipt_args(values: Iterable[str]) -> dict[str, Path]:
    result: dict[str, Path] = {}
    for value in values:
        if "=" not in value:
            raise CoverageError(f"--receipt expects PROFILE=PATH, got {value!r}")
        profile, path = value.split("=", 1)
        if not PROFILE_ID.fullmatch(profile) or not path:
            raise CoverageError(f"invalid --receipt binding {value!r}")
        if profile in result:
            raise CoverageError(f"duplicate receipt profile {profile}")
        result[profile] = Path(path)
    return result


def _function(surface: glossary_gate.Surface, path: tuple[str, ...]) -> dict[str, Any] | None:
    if surface.exports.get(path) != "function":
        return None
    item_id = surface.item_ids.get(path)
    item = surface.index.get(item_id) if item_id is not None else None
    inner = item.get("inner", {}) if isinstance(item, dict) else {}
    function = inner.get("function") if isinstance(inner, dict) else None
    return function if isinstance(function, dict) else None


def _input_names(function: Mapping[str, Any]) -> list[str] | None:
    inputs = function.get("sig", {}).get("inputs") if isinstance(function.get("sig"), dict) else None
    if not isinstance(inputs, list):
        return None
    names: list[str] = []
    for item in inputs:
        if not isinstance(item, list) or len(item) != 2 or not isinstance(item[0], str):
            return None
        names.append(item[0])
    return names


def _actual_limit_identity(surface: glossary_gate.Surface, function: Mapping[str, Any], index: int):
    inputs = function.get("sig", {}).get("inputs") if isinstance(function.get("sig"), dict) else None
    if not isinstance(inputs, list) or index >= len(inputs):
        return None
    value = inputs[index][1]
    resolved = ReturnTypes(surface).resolve(value)
    while isinstance(resolved, tuple) and resolved and resolved[0] == "ref":
        resolved = resolved[2]
    if isinstance(resolved, tuple) and len(resolved) == 3 and resolved[0] == "nominal":
        return ReturnTypes(surface)._resolve_nominal_identity(resolved[1])
    return None


def _expected_limit_identity(surface: glossary_gate.Surface, path: str):
    resolver = ReturnTypes(surface)
    expected_path = tuple(path.split("::"))
    expected = resolver.nominal_identity(expected_path)
    return expected if expected is not None else resolver.nominal_identity(expected_path, canonical=True)


def check_required_signature(surface: glossary_gate.Surface, entry: Mapping[str, Any]) -> str | None:
    path = tuple(entry["path"].split("::"))
    function = _function(surface, path)
    if function is None:
        return f"{entry['path']} is not an exported function"
    signature = entry["required_signature"]
    names = _input_names(function)
    if names != signature["inputs"]:
        return f"{entry['path']} inputs {names!r} != {signature['inputs']!r}"
    index = signature["inputs"].index(signature["limits_argument"])
    actual = _actual_limit_identity(surface, function, index)
    expected = _expected_limit_identity(surface, signature["limits_type"])
    if actual is None or expected is None or actual != expected:
        return f"{entry['path']} limits type identity {actual!r} != {expected!r}"
    return None


def _profile_from_rustdoc(manifest: Mapping[str, Any], profile_id: str,
                          document: Mapping[str, Any]) -> glossary_gate.Surface:
    profile = manifest["_profiles"].get(profile_id)
    if profile is None:
        raise CoverageError(f"rustdoc profile {profile_id!r} is not declared")
    return glossary_gate.Surface(document, profile.crate)


def census_profile(surface: glossary_gate.Surface) -> set[str]:
    """Return exact caller-visible public callable paths from one rustdoc graph."""
    return {
        "::".join(path)
        for path, kind in surface.exports.items()
        if kind == "function"
    }


def validate_manifest(manifest: Mapping[str, Any], *, root: Path,
                      glossary_path: Path, rustdoc: Mapping[str, Path] | None = None,
                      receipts: Mapping[str, Path] | None = None,
                      require_complete: bool = False,
                      verify_sources: bool = True) -> dict[str, Any]:
    """Validate a manifest and return a machine-readable finite report."""
    rustdoc = dict(rustdoc or {})
    receipts = dict(receipts or {})
    expected = _glossary_profiles(glossary_path)
    declared: dict[str, Profile] = manifest["_profiles"]
    report: dict[str, Any] = {
        "source_commit": manifest["source_commit"],
        "manifest_state": manifest["coverage"]["state"],
        "declared_profiles": sorted(declared),
        "expected_profiles": sorted(expected),
        "missing_profiles": sorted(set(expected) - set(declared)),
        # A package may have an explicit source profile before its glossary
        # row exists.  Keep those profiles visible, but do not mistake them
        # for missing/invalid callables; adding the eventual glossary row is a
        # separate completeness obligation.
        "custom_profiles": sorted(set(declared) - set(expected)),
        "profile_mismatches": [],
        "missing_workspace_packages": [],
        "missing_inputs": [],
        "extra_inputs": sorted(set(rustdoc) - set(declared)),
        "missing_receipts": sorted(set(declared) - set(receipts)) if verify_sources else [],
        "extra_receipts": sorted(set(receipts) - set(declared)) if verify_sources else [],
        "receipt_failures": [],
        "capture_overlay_failures": [],
        "loaded_profiles": [],
        "callable_counts": {},
        "unclassified": [],
        "stale_entries": [],
        "signature_failures": [],
        "source_failures": verify_source_anchors(manifest, root) if verify_sources else [],
        "current_source_failures": (
            verify_source_anchors(manifest, root, current=True) if verify_sources else []
        ),
        "active_entries": sorted(
            f"{entry['profile']}::{entry['path']}"
            for entry in manifest["entries"] if entry["disposition"] == "active_lane"
        ),
        "active_lanes": sorted(lane["id"] for lane in manifest["lanes"]),
    }
    # The coverage source paths are verified by their exact Git blob identities
    # above.  Do not require a sparse checkout to materialize them: the source
    # identity is the contract, and a worktree-only existence check would reject
    # a valid capture before rustdoc inputs are even considered.
    workspace = _workspace_packages(root)
    declared_packages = {profile.package for profile in declared.values()}
    report["missing_workspace_packages"] = sorted(workspace - declared_packages)
    for profile_id in sorted(set(declared) & set(expected)):
        actual = declared[profile_id]
        wanted = expected[profile_id]
        if actual != wanted:
            report["profile_mismatches"].append(
                f"{profile_id}: manifest {actual!r} != glossary {wanted!r}"
            )
    for profile_id, profile in declared.items():
        if verify_sources:
            receipt_path = receipts.get(profile_id)
            rustdoc_path = rustdoc.get(profile_id)
            if receipt_path is not None:
                if not receipt_path.is_file():
                    report["receipt_failures"].append(f"{profile_id}: missing receipt {receipt_path}")
                else:
                    try:
                        receipt_value = json.loads(receipt_path.read_text(encoding="utf-8"))
                        receipt = _capture_receipt(receipt_value, f"receipt[{profile_id}]")
                    except (OSError, json.JSONDecodeError, CoverageError) as error:
                        report["receipt_failures"].append(f"{profile_id}: invalid receipt: {error}")
                    else:
                        overlay_failures = _source_overlay_failures(
                            root, manifest["coverage"]["approved_cargo_lock_sha256"]
                        )
                        report["capture_overlay_failures"].extend(overlay_failures)
                        if rustdoc_path is None:
                            report["receipt_failures"].append(
                                f"{profile_id}: receipt has no matching rustdoc input"
                            )
                        else:
                            report["receipt_failures"].extend(
                                _validate_capture_receipt(
                                    manifest, profile, rustdoc_path, receipt_path, root,
                                    receipt, overlay_failures,
                                )
                            )
        path = rustdoc.get(profile_id)
        if path is None:
            report["missing_inputs"].append(profile_id)
            continue
        if not path.is_file():
            report["missing_inputs"].append(f"{profile_id}: {path}")
            continue
        try:
            surface = _profile_from_rustdoc(manifest, profile_id, _parse_rustdoc(path))
        except (CoverageError, ValueError) as error:
            report["stale_entries"].append(f"{profile_id}: rustdoc input: {error}")
            continue
        report["loaded_profiles"].append(profile_id)
        callables = census_profile(surface)
        report["callable_counts"][profile_id] = len(callables)
        entries = {
            entry["path"]: entry for entry in manifest["entries"]
            if entry["profile"] == profile_id
        }
        for path in sorted(callables - set(entries)):
            report["unclassified"].append(f"{profile_id}::{path}")
        for path in sorted(set(entries) - callables):
            entry = entries[path]
            if entry["disposition"] != "private_removed":
                report["stale_entries"].append(
                    f"{profile_id}::{path}: manifest entry is not exported by rustdoc"
                )
        for path, entry in entries.items():
            if path not in callables:
                continue
            if entry["disposition"] == "private_removed":
                report["stale_entries"].append(
                    f"{profile_id}::{path}: private_removed item remains public"
                )
            if entry["disposition"] == "required_raw_limits":
                failure = check_required_signature(surface, entry)
                if failure:
                    report["signature_failures"].append(failure)
    report["missing_inputs"] = sorted(report["missing_inputs"])
    report["loaded_profiles"] = sorted(report["loaded_profiles"])
    report["unclassified"] = sorted(report["unclassified"])
    report["stale_entries"] = sorted(report["stale_entries"])
    report["signature_failures"] = sorted(report["signature_failures"])
    report["missing_receipts"] = sorted(report["missing_receipts"])
    report["extra_receipts"] = sorted(report["extra_receipts"])
    report["receipt_failures"] = sorted(set(report["receipt_failures"]))
    report["capture_overlay_failures"] = sorted(set(report["capture_overlay_failures"]))
    report["complete"] = not any(
        report[key] for key in (
            "missing_profiles", "missing_workspace_packages",
            "profile_mismatches", "missing_inputs", "extra_inputs", "missing_receipts",
            "extra_receipts", "receipt_failures", "capture_overlay_failures",
            "unclassified", "stale_entries", "signature_failures",
            "source_failures", "current_source_failures", "active_entries", "active_lanes",
        )
    ) and manifest["coverage"]["state"] == "complete"
    if require_complete and not report["complete"]:
        report["failure"] = "raw API coverage is incomplete; see report fields"
    return report


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--glossary", type=Path, default=Path("tools/api_glossary.json"))
    parser.add_argument("--rustdoc", action="append", default=[], metavar="PROFILE=PATH",
                        help="bind an existing rustdoc JSON input; repeat per profile")
    parser.add_argument("--receipt", action="append", default=[], metavar="PROFILE=PATH",
                        help="bind the exact capture receipt for a rustdoc input; repeat per profile")
    parser.add_argument("--require-complete", action="store_true",
                        help="fail unless every profile/callable/disposition is complete")
    parser.add_argument("--json", action="store_true", help="print the full report as JSON")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _parser()
    args = parser.parse_args(argv)
    try:
        manifest = load_manifest(args.manifest)
        rustdoc = parse_rustdoc_args(args.rustdoc)
        receipts = parse_receipt_args(args.receipt)
        report = validate_manifest(
            manifest, root=args.root, glossary_path=args.glossary,
            rustdoc=rustdoc, receipts=receipts, require_complete=args.require_complete,
        )
    except CoverageError as error:
        parser.error(str(error))
    if args.json:
        print(json.dumps(report, indent=2, sort_keys=True))
    else:
        print(
            f"raw-api-coverage: complete={report['complete']} "
            f"loaded={len(report['loaded_profiles'])} "
            f"unclassified={len(report['unclassified'])} "
            f"active={len(report['active_entries']) + len(report['active_lanes'])} "
            f"stale={len(report['stale_entries']) + len(report['signature_failures'])}"
        )
        if not report["complete"]:
            print("raw-api-coverage: incomplete manifest; use --json for details")
    return 1 if args.require_complete and not report["complete"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
