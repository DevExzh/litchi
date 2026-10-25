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

The initial repository manifest is intentionally incomplete while the raw
limit waves are in flight.  Normal validation reports that state.  The final
``--require-complete`` mode fails unless every expected profile has a rustdoc
input, every public callable has one exact entry, all source identities and
required signatures match, and no active/unclassified/stale disposition
remains.
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
import re
from pathlib import Path
import subprocess
import tomllib
from typing import Any, Iterable, Mapping

try:
    from . import check_api_glossary as glossary_gate
    from .rustdoc_return_types import ReturnTypes
except ImportError:  # Direct command-line execution.
    import check_api_glossary as glossary_gate
    from rustdoc_return_types import ReturnTypes


SCHEMA_VERSION = 1
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
    if (not isinstance(coverage, dict)
            or set(coverage) != {"state", "reason", "expected_profile_source", "expected_workspace_source"}):
        raise CoverageError("coverage requires state, reason, expected_profile_source and expected_workspace_source")
    if coverage["state"] not in {"incomplete", "complete"}:
        raise CoverageError("coverage.state must be incomplete or complete")
    _nonempty_string(coverage["reason"], "coverage.reason")
    _path_string(coverage["expected_profile_source"], "coverage.expected_profile_source")
    _path_string(coverage["expected_workspace_source"], "coverage.expected_workspace_source")
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


def verify_source_anchors(manifest: Mapping[str, Any], root: Path, *, current: bool = False) -> list[str]:
    failures: list[str] = []
    commit = _git_head(root) if current else manifest["source_commit"]
    if commit is None:
        return ["cannot resolve a current HEAD for source-anchor verification"]
    if not current and _git_type(root, commit) != "commit":
        return [f"source_commit {commit} does not resolve to a Git commit"]
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
                      require_complete: bool = False,
                      verify_sources: bool = True) -> dict[str, Any]:
    """Validate a manifest and return a machine-readable finite report."""
    rustdoc = dict(rustdoc or {})
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
    if verify_sources:
        for source_key in ("expected_profile_source", "expected_workspace_source"):
            source_path = root / manifest["coverage"][source_key]
            if not source_path.is_file():
                report["source_failures"].append(
                    f"coverage.{source_key} is missing: {manifest['coverage'][source_key]}"
                )
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
    report["complete"] = not any(
        report[key] for key in (
            "missing_profiles", "missing_workspace_packages",
            "profile_mismatches", "missing_inputs", "extra_inputs", "unclassified", "stale_entries", "signature_failures",
            "source_failures", "current_source_failures", "active_entries", "active_lanes",
        )
    ) and manifest["coverage"]["state"] == "complete"
    if require_complete and not report["complete"]:
        report["failure"] = "raw API coverage is incomplete; see report fields"
    return report


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, default=Path("tools/raw_api_coverage.json"))
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--glossary", type=Path, default=Path("tools/api_glossary.json"))
    parser.add_argument("--rustdoc", action="append", default=[], metavar="PROFILE=PATH",
                        help="bind an existing rustdoc JSON input; repeat per profile")
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
        report = validate_manifest(
            manifest, root=args.root, glossary_path=args.glossary,
            rustdoc=rustdoc, require_complete=args.require_complete,
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
