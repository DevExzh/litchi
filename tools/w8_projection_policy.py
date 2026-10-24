#!/usr/bin/env python3
"""Verify the reviewed original-to-final W8 dependency projection pair.

The normal W8 runner requires one identical dependency projection at both
ends.  That rule is intentionally unchanged.  This module implements the
separate, opt-in original-to-final protocol: the real f2 endpoint keeps its
own manifest and lock, the final endpoint keeps its own manifest and lock,
and the difference is accepted only when it is the reviewed fixed graph
delta.  The policy and its pre/post images are checked in and pinned by this
module; callers cannot supply a replacement policy or a set of convenient
hashes.

The module does not build or execute a probe.  Endpoint manifests have
already passed the ordinary artifact and 109-route checks before the runner
calls the in-memory verifier.  The standalone ``pair`` command performs the
same artifact checks while producing the pair receipt before a timing run.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tools import w8_file_provenance as provenance


PACKAGE_NAME = "litchi-ops"
POLICY_RELATIVE = Path("docs/unified-ops/results/0362yh/original-to-final-policy.json")
POLICY_SCHEMA = "litchi_w8_original_to_final_projection_policy.v1"
PAIR_SCHEMA = "litchi_w8_original_to_final_pair.v1"
POLICY_VERSION = "0362yh-v1"

# Filled after the reviewed policy file is assembled.  A fixed digest is
# deliberate: an arbitrary policy path or edited JSON cannot become an
# accepted graph rule merely by being passed on the command line.
REVIEWED_POLICY_SHA256 = "5f198d0e4462529783538b6a7fa2f5a8cab025bc191f9ce582f150a9cd696d2e"

BASELINE_COMMIT = "f2bbbc8d07ab5da26c180f1ff74813d39f632d54"
FINAL_COMMIT = "f6fe430a4d19d94aa360c0e9ed03ebd757265225"
TARGET = "x86_64-unknown-linux-gnu"
PROFILE = "release"
PROBE_SHA256 = "21c8e3833c96411a612a2631f98093bd70cf0a5c28b2952c3d15873f464e54b0"
PROBE_BYTES = 75_035
FIXTURE_SHA256 = "b7439d899a940b4135dd44081554a4aa9d7e70c19e2ccd3339440baad5ca7e09"
FIXTURE_BYTES = 4_231
TOOLCHAIN_SHA256 = "e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310"

FEATURES = (
    "default", "doc", "docx", "encryption", "eval", "formula",
    "iwork-adapter", "keynote", "markdown", "numbers", "odb", "odc",
    "odf-formula", "odf-objects", "odg", "odi", "odm", "odp", "ods",
    "odt", "opc", "oth", "pages", "ppt", "pptx", "rtf", "test-support",
    "xls", "xlsb", "xlsx", "yaml",
)

REGISTRY_PACKAGE_COUNT = 518

# This is the complete reviewed local-edge delta.  The registry closure is
# compared separately, package by package, including checksums and dependency
# lists.  Keeping the local lists here makes an omitted or invented edge
# fail closed even if a future caller edits a receipt.
LOCAL_EDGE_DELTAS: tuple[dict[str, Any], ...] = (
    {
        "package": ["litchi-iwa-core", "0.0.1"],
        "before": [
            "litchi-iwa-common", "litchi-iwa-protos", "prost", "snap",
            "thiserror 2.0.18",
        ],
        "after": [
            "litchi-iwa-common", "litchi-iwa-protos", "prost", "smallvec 1.16.3",
            "snap", "thiserror 2.0.18",
        ],
    },
    {
        "package": ["litchi-odraw", "0.0.1"],
        "before": ["bitflags 2.13.2", "zerocopy", "zerocopy-derive"],
        "after": ["bitflags 2.13.2", "litchi-core", "zerocopy", "zerocopy-derive"],
    },
    {
        "package": ["litchi-ops", "0.0.1"],
        "before": [
            "base64", "caseless", "chrono", "litchi-cfb", "litchi-core",
            "litchi-crypto", "litchi-doc", "litchi-docx", "litchi-drawingml",
            "litchi-eval", "litchi-formula", "litchi-iwa-archive",
            "litchi-iwa-detect", "litchi-keynote", "litchi-markdown",
            "litchi-numbers", "litchi-odb", "litchi-odc", "litchi-odf-common",
            "litchi-odf-formula", "litchi-odg", "litchi-odi", "litchi-odm",
            "litchi-odp", "litchi-ods", "litchi-odt", "litchi-ole-common",
            "litchi-ooxml-common", "litchi-opc", "litchi-oth", "litchi-pages",
            "litchi-ppt", "litchi-pptx", "litchi-rtf", "litchi-sheet-text",
            "litchi-sign", "litchi-xls", "litchi-xlsb", "litchi-xlsx", "memchr",
            "proptest", "quick-xml 0.41.0", "serde", "serde-saphyr", "serde_json",
            "tempfile", "unicode-normalization", "zip",
        ],
        "after": [
            "base64", "caseless", "chrono", "litchi-biff", "litchi-cfb",
            "litchi-core", "litchi-crypto", "litchi-doc", "litchi-docx",
            "litchi-drawingml", "litchi-eval", "litchi-formula",
            "litchi-iwa-archive", "litchi-iwa-detect", "litchi-keynote",
            "litchi-markdown", "litchi-numbers", "litchi-odb", "litchi-odc",
            "litchi-odf-common", "litchi-odf-formula", "litchi-odg", "litchi-odi",
            "litchi-odm", "litchi-odp", "litchi-ods", "litchi-odt", "litchi-ograph",
            "litchi-ole-common", "litchi-ooxml-common", "litchi-opc", "litchi-oth",
            "litchi-pages", "litchi-ppt", "litchi-pptx", "litchi-rtf",
            "litchi-sheet-text", "litchi-sign", "litchi-xls", "litchi-xlsb",
            "litchi-xlsx", "memchr", "proptest", "quick-xml 0.41.0", "serde",
            "serde-saphyr", "serde_json", "sha2 0.11.1", "soapberry-zip", "tempfile",
            "unicode-normalization", "zip",
        ],
    },
    {
        "package": ["litchi-ppt", "0.0.1"],
        "before": [
            "bitflags 2.13.2", "chrono", "flate2", "litchi-cfb", "litchi-codepage",
            "litchi-core", "litchi-crypto", "litchi-fonts", "litchi-formula",
            "litchi-odraw", "litchi-ograph", "litchi-ole-common", "litchi-opc",
            "litchi-sign", "litchi-vba", "once_cell", "quick-xml 0.41.0",
            "rand 0.10.3", "serde_json", "sha1 0.11.0", "sha2 0.11.1",
            "soapberry-zip", "url", "zerocopy", "zerocopy-derive", "zeroize",
        ],
        "after": [
            "bitflags 2.13.2", "chrono", "flate2", "litchi-cfb", "litchi-codepage",
            "litchi-core", "litchi-crypto", "litchi-fonts", "litchi-formula",
            "litchi-odraw", "litchi-ograph", "litchi-ole-common", "litchi-opc",
            "litchi-sign", "litchi-vba", "once_cell", "quick-xml 0.41.0",
            "rand 0.10.3", "serde_json", "sha1 0.11.0", "sha2 0.11.1", "siphasher",
            "soapberry-zip", "url", "zerocopy", "zerocopy-derive", "zeroize",
        ],
    },
    {
        "package": ["litchi-xlsb", "0.0.1"],
        "before": [
            "bitflags 2.13.2", "chrono", "litchi-core", "litchi-drawingml",
            "litchi-ooxml-common", "litchi-opc", "litchi-sheet",
            "litchi-sign", "litchi-spreadsheet-drawing", "litchi-vba", "litchi-xldm",
            "once_cell", "quick-xml 0.41.0", "soapberry-zip", "tempfile",
            "thiserror 2.0.18",
        ],
        "after": [
            "bitflags 2.13.2", "chrono", "litchi-core", "litchi-drawingml",
            "litchi-ooxml-common", "litchi-opc", "litchi-sheet",
            "litchi-sign", "litchi-spreadsheet-drawing", "litchi-vba", "litchi-xldm",
            "once_cell", "quick-xml 0.41.0", "sha2 0.11.1", "soapberry-zip", "tempfile",
            "thiserror 2.0.18",
        ],
    },
    {
        "package": ["litchi-xlsx", "0.0.1"],
        "before": [
            "base64", "bitflags 2.13.2", "caseless", "chrono", "itoa", "litchi-core",
            "litchi-crypto", "litchi-drawingml", "litchi-ooxml-common", "litchi-opc",
            "litchi-sheet", "litchi-spreadsheet-drawing", "litchi-vba", "litchi-xldm",
            "memchr", "once_cell", "quick-xml 0.41.0", "ryu", "serde_json",
            "smallvec 1.16.3", "soapberry-zip", "tempfile", "thiserror 2.0.18",
            "unicode-normalization",
        ],
        "after": [
            "base64", "bitflags 2.13.2", "caseless", "chrono", "itoa", "litchi-core",
            "litchi-crypto", "litchi-drawingml", "litchi-ooxml-common", "litchi-opc",
            "litchi-sheet", "litchi-spreadsheet-drawing", "litchi-vba", "litchi-xldm",
            "memchr", "once_cell", "quick-xml 0.41.0", "ryu", "serde_json",
            "sha2 0.11.1", "smallvec 1.16.3", "soapberry-zip", "tempfile",
            "thiserror 2.0.18", "unicode-normalization",
        ],
    },
)

FIXED_INPUTS: dict[str, dict[str, Any]] = {
    "before": {
        "source_commit": BASELINE_COMMIT,
        "source": {"bytes": 6_614_528_000, "sha256": "de625729ec8e6f8055e619443ad31bf1a91719cd31465975f951d54dbc1c2280"},
        "workspace_manifest": {"path": "docs/unified-ops/results/0362yh/images/before/Cargo.toml", "bytes": 10_608, "sha256": "66a8f4d2b0ffb6b9ba4eb87d84fee8a3049e8ef7e4380b6ee48acaaf372dd7b5"},
        "package_manifest": {"path": "docs/unified-ops/results/0362yh/images/before/litchi-ops.Cargo.toml", "bytes": 12_490, "sha256": "80b231a0a052f5848ca7358a870847672fe39d8a6fb56b03e8881b8dbd840db2"},
        "toolchain": {"path": "docs/unified-ops/results/0362yh/images/before/rust-toolchain.toml", "bytes": 86, "sha256": TOOLCHAIN_SHA256},
        "lock": {"path": "docs/unified-ops/results/0362yh/images/before/Cargo.lock", "bytes": 139_441, "sha256": "1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c"},
        "projection": {"path": "docs/unified-ops/results/0362yh/images/before/projection.json", "bytes": 145_931, "sha256": "b070ddb516a0223f9f21509de6564c7c5690d26e955047a922d4d837b7522335"},
        "source_assembly": {"overlay_kind": "compatibility", "overlay_base_commit": BASELINE_COMMIT, "patch_bytes": 153_482, "patch_sha256": "3e05d33946c7a9a2fea5c84cda33c799fbf8cd3c92fbcf890812736641ed01cd"},
    },
    "after": {
        "source_commit": FINAL_COMMIT,
        "source": {"bytes": 6_629_089_280, "sha256": "cf533108be38b84f8c2ac4de536d6196e16c701c84f9a70645d05810db6d7f3b"},
        "workspace_manifest": {"path": "docs/unified-ops/results/0362yh/images/after/Cargo.toml", "bytes": 10_672, "sha256": "f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de"},
        "package_manifest": {"path": "docs/unified-ops/results/0362yh/images/after/litchi-ops.Cargo.toml", "bytes": 13_489, "sha256": "b22e9ff3a31366cd0bf85bd0d65b8add84cb89c612535ae905ea578b9a0eaf1d"},
        "toolchain": {"path": "docs/unified-ops/results/0362yh/images/after/rust-toolchain.toml", "bytes": 86, "sha256": TOOLCHAIN_SHA256},
        "lock": {"path": "docs/unified-ops/results/0362yh/images/after/Cargo.lock", "bytes": 139_591, "sha256": "52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6"},
        "projection": {"path": "docs/unified-ops/results/0362yh/images/after/projection.json", "bytes": 146_453, "sha256": "586cbbe06a8591322488a9650bbc098ad4b9f2f137c95e2ea80617728cbcd2e5"},
        "source_assembly": {"overlay_kind": "none"},
    },
}


class ProjectionPolicyError(RuntimeError):
    """A reviewed policy or endpoint failed closed."""


def fail(message: str) -> ProjectionPolicyError:
    return ProjectionPolicyError(message)


def canonical(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest(path: Path) -> str:
    value = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                value.update(block)
    except OSError as error:
        raise fail(f"cannot read {path}: {error}") from error
    return value.hexdigest()


def artifact(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if not path.is_file() or path.is_symlink():
        raise fail(f"artifact is not a regular file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": digest(path)}


def checked_commit(value: object, label: str) -> str:
    if not isinstance(value, str) or not re.fullmatch(r"[0-9a-f]{40}", value):
        raise fail(f"{label} is not a full source commit")
    return value


def _fixed_path(relative: str, label: str) -> Path:
    path = (ROOT / relative).resolve()
    try:
        path.relative_to(ROOT.resolve())
    except ValueError as error:
        raise fail(f"{label} escapes the reviewed repository root") from error
    return path


def _check_expected_file(path: Path, expected: dict[str, Any], label: str) -> dict[str, Any]:
    actual = artifact(path)
    if actual["bytes"] != expected["bytes"] or actual["sha256"] != expected["sha256"]:
        raise fail(f"{label} differs from the reviewed fixed pre/post image")
    return actual


def _policy_summary() -> dict[str, Any]:
    return {
        "schema": POLICY_SCHEMA,
        "policy_version": POLICY_VERSION,
        "original_commit": BASELINE_COMMIT,
        "final_commit": FINAL_COMMIT,
        "target": TARGET,
        "profile": PROFILE,
        "features": list(FEATURES),
        "probe": {"bytes": PROBE_BYTES, "sha256": PROBE_SHA256},
        "fixture": {"bytes": FIXTURE_BYTES, "sha256": FIXTURE_SHA256, "route_count": 109},
        "registry": {"count": REGISTRY_PACKAGE_COUNT, "checksums_and_dependencies": True},
        "local_edge_deltas": list(LOCAL_EDGE_DELTAS),
        "inputs": FIXED_INPUTS,
    }


def load_reviewed_policy(path: Path | None = None) -> dict[str, Any]:
    expected_path = (ROOT / POLICY_RELATIVE).resolve()
    actual_path = expected_path if path is None else path.resolve()
    if actual_path != expected_path:
        raise fail(f"policy path must be the reviewed {POLICY_RELATIVE}")
    try:
        payload = actual_path.read_bytes()
    except OSError as error:
        raise fail(f"cannot read reviewed policy: {actual_path}") from error
    if REVIEWED_POLICY_SHA256.startswith("__"):
        raise fail("reviewed policy digest is not registered")
    if hashlib.sha256(payload).hexdigest() != REVIEWED_POLICY_SHA256:
        raise fail("reviewed policy bytes are not the registered policy")
    try:
        value = json.loads(payload)
    except json.JSONDecodeError as error:
        raise fail(f"reviewed policy is invalid JSON: {error}") from error
    if not isinstance(value, dict) or canonical(value) != payload:
        raise fail("reviewed policy is not canonical JSON")
    if value != _policy_summary():
        raise fail("reviewed policy does not match the fixed source/graph contract")
    for side, expected in FIXED_INPUTS.items():
        for name in ("workspace_manifest", "package_manifest", "toolchain", "lock", "projection"):
            image = expected[name]
            _check_expected_file(_fixed_path(image["path"], f"{side}.{name}"), image, f"{side}.{name}")
    return value


def _load_fixed_projection(side: str) -> dict[str, Any]:
    expected = FIXED_INPUTS[side]["projection"]
    path = _fixed_path(expected["path"], f"{side}.projection")
    _check_expected_file(path, expected, f"{side}.projection")
    try:
        value = json.loads(path.read_bytes())
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise fail(f"{side} fixed projection is invalid JSON") from error
    if not isinstance(value, dict) or canonical(value) != path.read_bytes():
        raise fail(f"{side} fixed projection is not canonical JSON")
    return value


def _projection_packages(value: dict[str, Any], side: str) -> list[dict[str, Any]]:
    lock = value.get("lock")
    packages = lock.get("packages") if isinstance(lock, dict) else None
    if not isinstance(packages, list) or not packages:
        raise fail(f"{side} projection has no lock package list")
    result: list[dict[str, Any]] = []
    for index, package in enumerate(packages):
        if (
            not isinstance(package, dict)
            or not isinstance(package.get("name"), str)
            or not isinstance(package.get("version"), str)
            or not isinstance(package.get("dependencies"), list)
        ):
            raise fail(f"{side} projection package {index} is incomplete")
        result.append(package)
    return result


def _package_key(package: dict[str, Any]) -> tuple[str, str, object, object]:
    return (package["name"], package["version"], package.get("source"), package.get("checksum"))


def _package_identity(package: dict[str, Any]) -> tuple[str, str]:
    return (package["name"], package["version"])


def _validate_registry_pair(before: list[dict[str, Any]], after: list[dict[str, Any]]) -> None:
    before_registry = [item for item in before if isinstance(item.get("source"), str) and item["source"].startswith("registry+")]
    after_registry = [item for item in after if isinstance(item.get("source"), str) and item["source"].startswith("registry+")]
    if len(before_registry) != REGISTRY_PACKAGE_COUNT or len(after_registry) != REGISTRY_PACKAGE_COUNT:
        raise fail("registry closure does not contain exactly the reviewed 518 packages")
    for side, packages in (("before", before_registry), ("after", after_registry)):
        keys = set()
        for package in packages:
            checksum = package.get("checksum")
            if not isinstance(checksum, str) or not re.fullmatch(r"[0-9a-f]{64}", checksum):
                raise fail(f"{side} registry package has no valid checksum")
            key = _package_key(package)
            if key in keys:
                raise fail(f"{side} registry closure contains a duplicate package identity")
            keys.add(key)
    before_keys = {_package_key(item) for item in before_registry}
    after_keys = {_package_key(item) for item in after_registry}
    if before_keys != after_keys:
        raise fail("registry package identities or checksums differ")
    before_by_key = {_package_key(item): item for item in before_registry}
    after_by_key = {_package_key(item): item for item in after_registry}
    for key in sorted(before_by_key):
        if before_by_key[key]["dependencies"] != after_by_key[key]["dependencies"]:
            raise fail(f"registry dependency edge changed for {key[0]} {key[1]}")


def _validate_local_delta(before: list[dict[str, Any]], after: list[dict[str, Any]]) -> None:
    before_local = {_package_identity(item): item for item in before if item.get("source") is None}
    after_local = {_package_identity(item): item for item in after if item.get("source") is None}
    if set(before_local) != set(after_local):
        raise fail("local package identities differ from the reviewed pair")
    expected = {(tuple(item["package"]), tuple(item["before"]), tuple(item["after"])) for item in LOCAL_EDGE_DELTAS}
    actual = {
        (key, tuple(before_local[key]["dependencies"]), tuple(after_local[key]["dependencies"]))
        for key in before_local
        if before_local[key]["dependencies"] != after_local[key]["dependencies"]
    }
    if actual != expected:
        raise fail("local dependency edges differ from the exact reviewed six-package delta")


def validate_projection_pair(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    """Validate two already-loaded projections against the fixed graph rule."""
    if before.get("schema") != "litchi_w8_dependency_projection.v1" or after.get("schema") != "litchi_w8_dependency_projection.v1":
        raise fail("projection schema is not the reviewed W8 projection")
    for side, projection in (("before", before), ("after", after)):
        fixed = _load_fixed_projection(side)
        if projection != fixed:
            raise fail(f"{side} projection does not match the reviewed fixed pre/post image")
        if projection.get("features") != {"package": "litchi-ops", "selected": list(FEATURES), "all_features": True}:
            raise fail(f"{side} feature selection differs from the reviewed all-features set")
        build = projection.get("build")
        if not isinstance(build, dict) or build.get("target") != TARGET or build.get("profile") != PROFILE:
            raise fail(f"{side} target/profile differs from the reviewed build")
        toolchain = build.get("toolchain")
        if not isinstance(toolchain, dict) or toolchain.get("sha256") != TOOLCHAIN_SHA256:
            raise fail(f"{side} toolchain differs from the reviewed build")
        workload = projection.get("workload")
        if workload != {
            "probe_sha256": PROBE_SHA256,
            "fixture_sha256": FIXTURE_SHA256,
            "fixture_schema": "litchi_w8_fixture_descriptor.v1",
            "route_count": 109,
        }:
            raise fail(f"{side} workload projection differs from the reviewed 109-route input")
    _validate_registry_pair(_projection_packages(before, "before"), _projection_packages(after, "after"))
    before_packages = _projection_packages(before, "before")
    after_packages = _projection_packages(after, "after")
    _validate_local_delta(before_packages, after_packages)
    return {
        "registry_package_count": REGISTRY_PACKAGE_COUNT,
        "registry_identities_checksums_dependencies_equal": True,
        "local_edge_differences": [dict(item) for item in LOCAL_EDGE_DELTAS],
    }


def _load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_bytes())
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise fail(f"{label} is invalid JSON") from error
    if not isinstance(value, dict):
        raise fail(f"{label} is not a JSON object")
    if canonical(value) != path.read_bytes():
        raise fail(f"{label} is not canonical JSON")
    return value


def _actual_artifact(row: object, label: str) -> dict[str, Any]:
    if not isinstance(row, dict) or not isinstance(row.get("path"), str):
        raise fail(f"{label} artifact is missing")
    actual = artifact(Path(row["path"]))
    if actual["bytes"] != row.get("bytes") or actual["sha256"] != row.get("sha256"):
        raise fail(f"{label} artifact identity is stale")
    return actual


def _endpoint(path: Path, label: str) -> dict[str, Any]:
    value = _load_json(path.resolve(), f"{label} endpoint")
    if value.get("schema") != "litchi_file_endpoint.v1":
        raise fail(f"{label} endpoint schema is invalid")
    commit = checked_commit(value.get("source_commit"), f"{label}.source_commit")
    rows: dict[str, dict[str, Any]] = {}
    for name in (
        "source", "lock", "probe", "fixture", "dependency_projection",
        "workspace_manifest", "package_manifest", "toolchain", "source_assembly",
    ):
        rows[name] = _actual_artifact(value.get(name), f"{label}.{name}")
    projection = _load_json(Path(rows["dependency_projection"]["path"]), f"{label}.projection")
    if rows["dependency_projection"]["sha256"] != value.get("dependency_projection_sha256"):
        raise fail(f"{label} projection hash is not bound to the endpoint")
    fixture = _load_json(Path(rows["fixture"]["path"]), f"{label}.fixture")
    routes = provenance.fixture_routes(fixture, f"{label}.fixture")
    assembly = _load_json(Path(rows["source_assembly"]["path"]), f"{label}.source_assembly")
    return {
        "path": path.resolve(),
        "value": value,
        "commit": commit,
        "rows": rows,
        "projection": projection,
        "fixture": fixture,
        "routes": routes,
        "assembly": assembly,
    }


def _check_endpoint_fixed(endpoint: dict[str, Any], side: str) -> None:
    expected = FIXED_INPUTS[side]
    expected_commit = expected["source_commit"]
    if endpoint["commit"] != expected_commit:
        raise fail(f"{side} source commit is not the reviewed full commit")
    for name in ("workspace_manifest", "package_manifest", "toolchain", "lock", "projection"):
        fixed = expected[name]
        actual = endpoint["rows"][name]
        if (actual["bytes"], actual["sha256"]) != (fixed["bytes"], fixed["sha256"]):
            raise fail(f"{side} {name} does not match the reviewed fixed pre/post image")
    for name, bytes_, sha in (("source", expected["source"]["bytes"], expected["source"]["sha256"]), ("probe", PROBE_BYTES, PROBE_SHA256), ("fixture", FIXTURE_BYTES, FIXTURE_SHA256)):
        actual = endpoint["rows"][name]
        if (actual["bytes"], actual["sha256"]) != (bytes_, sha):
            raise fail(f"{side} {name} is not the reviewed artifact")
    assembly = endpoint["assembly"]
    fixed_assembly = expected["source_assembly"]
    if (
        assembly.get("schema") != "litchi_w8_source_assembly.v1"
        or assembly.get("source_commit") != expected_commit
        or assembly.get("resolved_commit") != expected_commit
    ):
        raise fail(f"{side} source assembly is not bound to the reviewed commit")
    source_archive = assembly.get("source_archive")
    if not isinstance(source_archive, dict) or (source_archive.get("bytes"), source_archive.get("sha256")) != (endpoint["rows"]["source"]["bytes"], endpoint["rows"]["source"]["sha256"]):
        raise fail(f"{side} source assembly archive differs from the endpoint source")
    overlay = assembly.get("overlay")
    if not isinstance(overlay, dict) or overlay.get("kind") != fixed_assembly["overlay_kind"]:
        raise fail(f"{side} source assembly overlay kind is not reviewed")
    if fixed_assembly["overlay_kind"] == "compatibility":
        if overlay.get("base_commit") != fixed_assembly["overlay_base_commit"]:
            raise fail(f"{side} source assembly overlay base is not f2")
        patch = overlay.get("patch")
        if not isinstance(patch, dict) or patch.get("bytes") != fixed_assembly["patch_bytes"] or patch.get("sha256") != fixed_assembly["patch_sha256"]:
            raise fail(f"{side} source assembly adapter patch is not the reviewed consumer-only patch")


def _check_fields_fixed(endpoint: dict[str, Any], fields: dict[str, Any], side: str) -> None:
    """Check a runner-validated endpoint without hashing its source archive again."""
    expected = FIXED_INPUTS[side]
    if endpoint.get("schema") != "litchi_file_endpoint.v1":
        raise fail(f"{side} endpoint schema is invalid")
    if fields.get("source_commit") != expected["source_commit"]:
        raise fail(f"{side} source commit is not the reviewed full commit")
    if endpoint.get("source_commit") != fields.get("source_commit") or endpoint.get("dependency_projection_sha256") != fields.get("dependency_projection_sha256"):
        raise fail(f"{side} endpoint provenance is not bound to the runner fields")
    for name in ("source", "workspace_manifest", "package_manifest", "toolchain", "lock", "probe", "fixture"):
        actual = fields.get(name)
        fixed = expected.get(name)
        if fixed is None:
            fixed = {"bytes": PROBE_BYTES, "sha256": PROBE_SHA256} if name == "probe" else {"bytes": FIXTURE_BYTES, "sha256": FIXTURE_SHA256}
        if not isinstance(actual, dict) or (actual.get("bytes"), actual.get("sha256")) != (fixed["bytes"], fixed["sha256"]):
            raise fail(f"{side} {name} differs from the reviewed endpoint identity")
    if fields.get("dependency_projection_sha256") != expected["projection"]["sha256"]:
        raise fail(f"{side} dependency projection is not the reviewed fixed image")

    def endpoint_row(name: str) -> dict[str, Any]:
        row = endpoint.get(name)
        if not isinstance(row, dict) or not isinstance(row.get("path"), str):
            raise fail(f"{side} endpoint lacks {name}")
        if name == "dependency_projection":
            if (row.get("bytes"), row.get("sha256")) != (expected["projection"]["bytes"], expected["projection"]["sha256"]):
                raise fail(f"{side} endpoint {name} differs from the reviewed fixed image")
            return row
        if name == "source_assembly":
            return row
        field = fields.get(name, {})
        if not isinstance(field, dict) or (row.get("bytes"), row.get("sha256")) != (field.get("bytes"), field.get("sha256")):
            raise fail(f"{side} endpoint {name} identity differs from the runner fields")
        return row

    for name in ("source", "lock", "probe", "fixture", "workspace_manifest", "package_manifest", "toolchain"):
        endpoint_row(name)
    projection = endpoint_row("dependency_projection")
    if projection.get("sha256") != fields["dependency_projection_sha256"]:
        raise fail(f"{side} endpoint projection hash is stale")
    assembly_row = endpoint_row("source_assembly")
    assembly = _load_json(Path(assembly_row["path"]), f"{side}.source_assembly")
    if (
        assembly.get("schema") != "litchi_w8_source_assembly.v1"
        or assembly.get("source_commit") != expected["source_commit"]
        or assembly.get("resolved_commit") != expected["source_commit"]
    ):
        raise fail(f"{side} source assembly is not bound to the reviewed commit")
    source_archive = assembly.get("source_archive")
    if not isinstance(source_archive, dict) or (source_archive.get("bytes"), source_archive.get("sha256")) != (fields["source"]["bytes"], fields["source"]["sha256"]):
        raise fail(f"{side} source assembly archive differs from the endpoint source")
    overlay = assembly.get("overlay")
    fixed_assembly = expected["source_assembly"]
    if not isinstance(overlay, dict) or overlay.get("kind") != fixed_assembly["overlay_kind"]:
        raise fail(f"{side} source assembly overlay kind is not reviewed")
    if fixed_assembly["overlay_kind"] == "compatibility":
        if overlay.get("base_commit") != fixed_assembly["overlay_base_commit"]:
            raise fail(f"{side} source assembly overlay base is not f2")
        patch = overlay.get("patch")
        if not isinstance(patch, dict) or patch.get("bytes") != fixed_assembly["patch_bytes"] or patch.get("sha256") != fixed_assembly["patch_sha256"]:
            raise fail(f"{side} source assembly adapter patch is not the reviewed consumer-only patch")


def verify_pair(policy_path: Path, before_path: Path, after_path: Path) -> dict[str, Any]:
    """Verify real endpoint artifacts and return the canonical pair receipt."""
    policy = load_reviewed_policy(policy_path)
    before = _endpoint(before_path, "before")
    after = _endpoint(after_path, "after")
    _check_endpoint_fixed(before, "before")
    _check_endpoint_fixed(after, "after")
    if before["routes"] != after["routes"] or len(before["routes"]) != 109:
        raise fail("before/after route order is not the exact reviewed 109-route order")
    projection_delta = validate_projection_pair(before["projection"], after["projection"])
    receipt = {
        "schema": PAIR_SCHEMA,
        "policy": {
            "schema": policy["schema"],
            "version": policy["policy_version"],
            "sha256": REVIEWED_POLICY_SHA256,
        },
        "before": {
            "manifest": artifact(before_path),
            "source_commit": before["commit"],
            "source_sha256": before["rows"]["source"]["sha256"],
            "workspace_manifest_sha256": before["rows"]["workspace_manifest"]["sha256"],
            "package_manifest_sha256": before["rows"]["package_manifest"]["sha256"],
            "lock_sha256": before["rows"]["lock"]["sha256"],
            "projection_sha256": before["rows"]["dependency_projection"]["sha256"],
        },
        "after": {
            "manifest": artifact(after_path),
            "source_commit": after["commit"],
            "source_sha256": after["rows"]["source"]["sha256"],
            "workspace_manifest_sha256": after["rows"]["workspace_manifest"]["sha256"],
            "package_manifest_sha256": after["rows"]["package_manifest"]["sha256"],
            "lock_sha256": after["rows"]["lock"]["sha256"],
            "projection_sha256": after["rows"]["dependency_projection"]["sha256"],
        },
        "workload": {
            "probe_sha256": PROBE_SHA256,
            "fixture_sha256": FIXTURE_SHA256,
            "route_count": 109,
            "routes": before["routes"],
        },
        "build": {
            "target": TARGET,
            "profile": PROFILE,
            "features": list(FEATURES),
            "toolchain_sha256": TOOLCHAIN_SHA256,
        },
        "projection": projection_delta,
        "status": "verified_protocol_only",
    }
    return receipt


def write_pair_receipt(policy_path: Path, before_path: Path, after_path: Path, output: Path) -> dict[str, Any]:
    receipt = verify_pair(policy_path, before_path, after_path)
    output = output.resolve()
    if output.exists():
        raise fail(f"refusing to overwrite pair receipt: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(canonical(receipt))
    return {"path": str(output), "bytes": output.stat().st_size, "sha256": digest(output)}


def verify_pair_receipt(policy_path: Path, receipt_path: Path, before_fields: dict[str, Any], after_fields: dict[str, Any]) -> dict[str, Any]:
    """Verify a precomputed pair receipt against already-validated endpoints.

    The runner has just hashed the large source archives while checking each
    endpoint.  This path reuses those actual fields and therefore does not
    perform a second multi-gigabyte source hash during runner startup.
    """
    policy = load_reviewed_policy(policy_path)
    receipt = _load_json(receipt_path.resolve(), "original-to-final pair receipt")
    if receipt.get("schema") != PAIR_SCHEMA or receipt.get("policy") != {
        "schema": POLICY_SCHEMA,
        "version": POLICY_VERSION,
        "sha256": REVIEWED_POLICY_SHA256,
    }:
        raise fail("pair receipt is not bound to the reviewed policy version")
    for side, fields in (("before", before_fields), ("after", after_fields)):
        endpoint = _load_json(Path(fields["manifest"]["path"]), f"{side} endpoint")
        _check_fields_fixed(endpoint, fields, side)
        manifest = receipt[side].get("manifest") if isinstance(receipt.get(side), dict) else None
        if not isinstance(manifest, dict) or manifest.get("sha256") != fields["manifest"]["sha256"] or manifest.get("bytes") != fields["manifest"]["bytes"]:
            raise fail(f"{side} pair receipt is not bound to the endpoint manifest")
    # Re-read only the small projection JSONs and fixed endpoint manifests.
    before_value = _load_json(Path(before_fields["manifest"]["path"]), "before endpoint")
    after_value = _load_json(Path(after_fields["manifest"]["path"]), "after endpoint")
    before_projection = _load_json(Path(before_value["dependency_projection"]["path"]), "before projection")
    after_projection = _load_json(Path(after_value["dependency_projection"]["path"]), "after projection")
    delta = validate_projection_pair(before_projection, after_projection)
    before_routes = provenance.fixture_routes(_load_json(Path(before_value["fixture"]["path"]), "before fixture"))
    after_routes = provenance.fixture_routes(_load_json(Path(after_value["fixture"]["path"]), "after fixture"))
    if before_routes != after_routes or before_fields.get("routes") != before_routes or after_fields.get("routes") != after_routes:
        raise fail("before/after route order is not the exact reviewed 109-route order")
    if receipt.get("projection") != delta or receipt.get("workload", {}).get("route_count") != 109:
        raise fail("pair receipt projection proof is stale")
    if receipt.get("workload", {}).get("routes") != before_routes:
        raise fail("pair receipt route order is stale")
    return receipt


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n", 1)[0])
    pair = parser.add_subparsers(dest="command", required=True).add_parser(
        "pair", help="verify two endpoint manifests and write the fixed pair receipt"
    )
    pair.add_argument("--policy", required=True, type=Path)
    pair.add_argument("--before-manifest", required=True, type=Path)
    pair.add_argument("--after-manifest", required=True, type=Path)
    pair.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)
    if args.command == "pair":
        result = write_pair_receipt(args.policy, args.before_manifest, args.after_manifest, args.output)
        print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ProjectionPolicyError as error:
        print(f"w8 projection policy: {error}", file=sys.stderr)
        raise SystemExit(2)

