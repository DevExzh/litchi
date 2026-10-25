#!/usr/bin/env python3
"""Prepare the deterministic current raw-API profile matrix.

The checked-in ``raw_api_coverage.json`` is a historical, source-bound
coverage manifest.  This tool does not rewrite it or manufacture entries,
captures, dispositions, lanes, or receipts.  It derives the next matrix from
the current glossary and every direct ``crates/*`` package, carries forward
only the approved F5 capture-policy description, and marks the result
incomplete until rustdoc inputs and callable dispositions are supplied.

The output is a preparation artifact for the final source freeze.  It is
deliberately a different schema from ``raw_api_coverage.json`` so historical
evidence cannot be relabeled as evidence for a new candidate.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tomllib
from typing import Any


SCHEMA = "litchi.raw-api-census-matrix.v1"
GLOSSARY_PATH = "tools/api_glossary.json"
WORKSPACE_PATH = "Cargo.toml"
LEGACY_MANIFEST_PATH = "tools/raw_api_coverage.json"
PACKAGE_ONLY_FEATURES = {
    "litchi-eval": ("web_functions",),
    "litchi-odf": ("all",),
}
PACKAGE_ONLY_ID_SUFFIX = {
    "litchi-eval": "-full",
    "litchi-odf": "-full",
}


class CensusError(ValueError):
    """The source tree cannot produce a safe deterministic profile matrix."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise CensusError(message)


def _sha256(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def _git(root: Path, *arguments: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *arguments],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise CensusError(f"git {' '.join(arguments)} failed: {result.stderr.strip()}")
    return result.stdout.strip()


def _root(value: Path) -> Path:
    root = value.resolve(strict=False)
    _expect(root.is_dir(), f"root is not a directory: {root}")
    top = Path(_git(root, "rev-parse", "--show-toplevel")).resolve(strict=False)
    _expect(top == root, f"root is not a Git worktree top-level: {root}")
    return root


def _git_blob_at(root: Path, revision: str, relative: str) -> str:
    output = _git(root, "ls-tree", revision, "--", relative)
    fields = output.split()
    _expect(len(fields) == 4 and fields[1] == "blob" and fields[3] == relative, f"missing tracked blob: {relative}")
    return fields[2]


def _git_blob(root: Path, relative: str) -> str:
    return _git_blob_at(root, "HEAD", relative)


def _worktree_blob(root: Path, relative: str) -> str:
    """Return the blob hash of the bytes currently present in the worktree."""
    return _git(root, "hash-object", f"--path={relative}", relative)


def _file_identity(root: Path, relative: str) -> dict[str, Any]:
    path = root / relative
    _expect(path.is_file() and not path.is_symlink(), f"required source file is missing or symlinked: {path}")
    payload = path.read_bytes()
    blob = _git_blob(root, relative)
    _expect(
        _worktree_blob(root, relative) == blob,
        f"working-tree bytes differ from the tracked blob: {relative}",
    )
    return {
        "path": relative,
        "blob": blob,
        "bytes": len(payload),
        "sha256": _sha256(payload),
    }


def _load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise CensusError(f"cannot read {label}: {error}") from error
    _expect(isinstance(value, dict), f"{label} must be an object")
    return value


def _features(document: dict[str, Any], package: str) -> dict[str, Any]:
    features = document.get("features", {})
    _expect(isinstance(features, dict), f"{package} features table is invalid")
    return features


def _package_manifests(root: Path) -> dict[str, tuple[str, dict[str, Any]]]:
    result: dict[str, tuple[str, dict[str, Any]]] = {}
    for manifest in sorted((root / "crates").glob("*/Cargo.toml")):
        relative = manifest.relative_to(root).as_posix()
        try:
            document = tomllib.loads(manifest.read_text(encoding="utf-8"))
        except (OSError, tomllib.TOMLDecodeError) as error:
            raise CensusError(f"cannot read package manifest {relative}: {error}") from error
        package = document.get("package")
        if not isinstance(package, dict) or not isinstance(package.get("name"), str):
            continue
        name = package["name"]
        _expect(name not in result, f"duplicate direct package: {name}")
        result[name] = (relative, document)
    _expect(len(result) == 68, f"expected 68 direct package manifests, found {len(result)}")
    return result


def _glossary_profiles(value: dict[str, Any]) -> list[dict[str, Any]]:
    profiles = value.get("profiles")
    _expect(isinstance(profiles, list), "glossary profiles must be an array")
    result: list[dict[str, Any]] = []
    seen: set[str] = set()
    for index, profile in enumerate(profiles):
        _expect(isinstance(profile, dict), f"glossary profile {index} is not an object")
        profile_id = profile.get("id")
        package = profile.get("package")
        default_features = profile.get("default_features")
        features = profile.get("features")
        _expect(isinstance(profile_id, str) and profile_id, f"glossary profile {index} has no id")
        _expect(profile_id not in seen, f"duplicate glossary profile: {profile_id}")
        _expect(isinstance(package, str) and package, f"glossary profile {profile_id} has no package")
        _expect(isinstance(default_features, bool), f"glossary profile {profile_id} default_features is invalid")
        _expect(isinstance(features, list) and all(isinstance(item, str) for item in features), f"glossary profile {profile_id} features are invalid")
        seen.add(profile_id)
        result.append({
            "id": profile_id,
            "package": package,
            "default_features": default_features,
            "features": list(features),
            "origin": "glossary",
        })
    _expect(len(result) == 71, f"expected 71 glossary profiles, found {len(result)}")
    return result


def _package_only_profiles(
    manifests: dict[str, tuple[str, dict[str, Any]]],
    glossary_packages: set[str],
) -> list[dict[str, Any]]:
    package_only: list[dict[str, Any]] = []
    for package in sorted(set(manifests) - glossary_packages):
        relative, document = manifests[package]
        features = PACKAGE_ONLY_FEATURES.get(package, ())
        available = _features(document, package)
        for feature in features:
            _expect(feature in available, f"package-only feature {package}:{feature} is not declared")
        suffix = PACKAGE_ONLY_ID_SUFFIX.get(package, "")
        package_only.append({
            "id": f"workspace-{package}{suffix}",
            "package": package,
            "default_features": False,
            "features": list(features),
            "origin": "package_only",
        })
    _expect(len(package_only) == 13, f"expected 13 package-only profiles, found {len(package_only)}")
    return package_only


def build_matrix(root: Path, *, source_commit: str | None = None) -> dict[str, Any]:
    root = _root(root)
    glossary = _load_json(root / GLOSSARY_PATH, "glossary")
    legacy = _load_json(root / LEGACY_MANIFEST_PATH, "legacy coverage manifest")
    manifests = _package_manifests(root)
    glossary_profiles = _glossary_profiles(glossary)
    glossary_packages = {profile["package"] for profile in glossary_profiles}
    for package in glossary_packages:
        _expect(package in manifests, f"glossary package is not a direct package: {package}")
    package_only = _package_only_profiles(manifests, glossary_packages)

    profile_values = glossary_profiles + package_only
    profile_ids = [profile["id"] for profile in profile_values]
    _expect(len(profile_ids) == len(set(profile_ids)), "matrix profile ids are not unique")
    for profile in profile_values:
        relative, _ = manifests[profile["package"]]
        profile["manifest"] = relative
        profile["manifest_identity"] = _file_identity(root, relative)

    glossary_identity = _file_identity(root, GLOSSARY_PATH)
    workspace_identity = _file_identity(root, WORKSPACE_PATH)
    legacy_identity = _file_identity(root, LEGACY_MANIFEST_PATH)
    if source_commit is None:
        source_commit = _git(root, "rev-parse", "HEAD^{commit}")
    _expect(len(source_commit) == 40, "source commit is not a full SHA-1")
    _expect(
        _git(root, "cat-file", "-e", f"{source_commit}^{{commit}}") == "",
        f"source commit is not a commit: {source_commit}",
    )
    for relative, identity in (
        (GLOSSARY_PATH, glossary_identity),
        (WORKSPACE_PATH, workspace_identity),
        (LEGACY_MANIFEST_PATH, legacy_identity),
    ):
        _expect(
            _git_blob_at(root, source_commit, relative) == identity["blob"],
            f"{relative} differs from source commit {source_commit}; rebind the matrix",
        )
    for profile in profile_values:
        identity = profile["manifest_identity"]
        _expect(
            _git_blob_at(root, source_commit, identity["path"]) == identity["blob"],
            f"{identity['path']} differs from source commit {source_commit}; rebind the matrix",
        )
    legacy_profiles = legacy.get("profiles", [])
    legacy_entries = legacy.get("entries", [])
    legacy_lanes = legacy.get("lanes", [])
    _expect(isinstance(legacy_profiles, list), "legacy manifest profiles are invalid")
    _expect(isinstance(legacy_entries, list), "legacy manifest entries are invalid")
    _expect(isinstance(legacy_lanes, list), "legacy manifest lanes are invalid")
    coverage = legacy.get("coverage")
    _expect(isinstance(coverage, dict), "legacy manifest coverage policy is missing")
    capture_policy = coverage.get("capture_policy")
    _expect(isinstance(capture_policy, dict), "legacy F5 capture policy is missing")

    return {
        "schema": SCHEMA,
        "state": "incomplete",
        "reason": (
            "84 current profile inputs are prepared from 71 glossary profiles and "
            "13 direct package-only profiles; rustdoc inputs, receipts, callable "
            "entries, and dispositions are intentionally not supplied"
        ),
        "source": {
            "commit": source_commit,
            "glossary": glossary_identity,
            "workspace": workspace_identity,
        },
        "counts": {
            "profiles": len(profile_values),
            "glossary_profiles": len(glossary_profiles),
            "package_only_profiles": len(package_only),
            "direct_packages": len(manifests),
        },
        "profiles": profile_values,
        "legacy_coverage": {
            "manifest": legacy_identity,
            "source_commit": legacy.get("source_commit"),
            "profiles": len(legacy_profiles),
            "entries": len(legacy_entries),
            "lanes": len(legacy_lanes),
        },
        "capture_policy": capture_policy,
        "entries": [],
        "lanes": [],
        "captures": [],
    }


def write_matrix(matrix: dict[str, Any], output: Path, *, check: bool = False) -> bool:
    payload = json.dumps(matrix, indent=2, sort_keys=True) + "\n"
    if check:
        _expect(output.is_file(), f"matrix output is missing: {output}")
        actual = output.read_text(encoding="utf-8")
        if actual != payload:
            raise CensusError(f"matrix output is stale: {output}")
        return False
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(payload, encoding="utf-8")
    return True


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path("."))
    parser.add_argument("--output", type=Path, default=Path("tools/raw_api_census_matrix.json"))
    parser.add_argument(
        "--source-commit",
        help="bind the input files to this commit (defaults to the current HEAD)",
    )
    parser.add_argument("--check", action="store_true")
    arguments = parser.parse_args(argv)
    try:
        output = arguments.output
        source_commit = arguments.source_commit
        if arguments.check and source_commit is None:
            existing = _load_json(output, "existing census matrix")
            source = existing.get("source")
            source_commit = source.get("commit") if isinstance(source, dict) else None
            _expect(isinstance(source_commit, str), "existing census matrix has no source commit")
        matrix = build_matrix(arguments.root, source_commit=source_commit)
        write_matrix(matrix, output, check=arguments.check)
    except (OSError, CensusError) as error:
        parser.error(str(error))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
