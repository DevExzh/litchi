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
already passed the complete existing endpoint validator before the runner
calls the in-memory verifier.  The standalone ``pair`` command invokes that
same validator for both endpoints before producing a pair receipt.  Pair
verification rehashes the source archives before timing starts so a stale or
missing source cannot be represented by metadata alone.
"""

from __future__ import annotations

import argparse
from collections.abc import Iterator, Mapping
import hashlib
import importlib.util
import json
import os
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
POLICY_VERSION = "0362yh-472b-v1"

# Filled after the reviewed policy file is assembled.  A fixed digest is
# deliberate: an arbitrary policy path or edited JSON cannot become an
# accepted graph rule merely by being passed on the command line.
REVIEWED_POLICY_SHA256 = "e56d549c733629ac73492a3511a45a1ed6a4cec97f1a4705ceea4ba9b67d97d2"

BASELINE_COMMIT = "f2bbbc8d07ab5da26c180f1ff74813d39f632d54"
FINAL_COMMIT = "472b6705475b073e7fe08f6e11ee17fb1e29f838"
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
    "odt", "opc", "oth", "pages", "ppt", "pptx", "rtf", "sign", "test-support",
    "xls", "xlsb", "xlsx", "yaml",
)

# The immutable f2 workspace predates the final `sign` facade feature.  Keep
# that side-specific all-features selection explicit: comparing both sides to
# the final list would silently require a production-manifest overlay in the
# historical endpoint, which is outside the reviewed consumer-only adapter.
BEFORE_FEATURES = (
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
            "litchi-sheet-text", "litchi-sign", "litchi-vba", "litchi-xls", "litchi-xlsb",
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
        "source": {"bytes": 6_634_147_840, "sha256": "44c11e933974c8fa3d6a67b46739bedeef866e60fbd4360ffe7f3bb8a4dc2274"},
        "workspace_manifest": {"path": "docs/unified-ops/results/0362yh/images/after/Cargo.toml", "bytes": 10_672, "sha256": "f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de"},
        "package_manifest": {"path": "docs/unified-ops/results/0362yh/images/after/litchi-ops.Cargo.toml", "bytes": 13_718, "sha256": "004d3b3cba8c035d209665a76c9cfe3932f12a9c198a4aea9fa314ef8ebb3a39"},
        "toolchain": {"path": "docs/unified-ops/results/0362yh/images/after/rust-toolchain.toml", "bytes": 86, "sha256": TOOLCHAIN_SHA256},
        "lock": {"path": "docs/unified-ops/results/0362yh/images/after/Cargo.lock", "bytes": 139_606, "sha256": "4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1"},
        "projection": {"path": "docs/unified-ops/results/0362yh/images/after/projection.json", "bytes": 146_621, "sha256": "2964946c802881909d5aaeea427fb32e85d0657f0987185ce21ff29e84c0d686"},
        "source_assembly": {"overlay_kind": "none"},
    },
}


class ProjectionPolicyError(RuntimeError):
    """A reviewed policy or endpoint failed closed."""


_VALIDATION_TOKEN = object()


class _FrozenMap(Mapping[str, Any]):
    """Recursively immutable mapping used by the endpoint capability."""

    __slots__ = ("_items",)

    def __init__(self, value: Mapping[str, Any]) -> None:
        object.__setattr__(self, "_items", tuple((key, _freeze(item)) for key, item in value.items()))

    def __getitem__(self, key: str) -> Any:
        for item_key, item in self._items:
            if item_key == key:
                return item
        raise KeyError(key)

    def __iter__(self) -> Iterator[str]:
        return (key for key, _item in self._items)

    def __len__(self) -> int:
        return len(self._items)

    def __setattr__(self, _name: str, _value: Any) -> None:
        raise TypeError("frozen endpoint mappings are immutable")

    def __delattr__(self, _name: str) -> None:
        raise TypeError("frozen endpoint mappings are immutable")


def _freeze(value: Any) -> Any:
    if isinstance(value, _FrozenMap):
        return value
    if isinstance(value, Mapping):
        return _FrozenMap(value)
    if isinstance(value, (list, tuple)):
        return tuple(_freeze(item) for item in value)
    return value


def _thaw(value: Any) -> Any:
    if isinstance(value, _FrozenMap):
        return {key: _thaw(item) for key, item in value.items()}
    if isinstance(value, tuple):
        return [_thaw(item) for item in value]
    return value


class _ValidatedEndpoint(Mapping[str, Any]):
    """Opaque fields returned only by the full W8 endpoint validator."""

    __slots__ = ("_fields", "_token")

    def __init__(self, fields: dict[str, Any], token: object) -> None:
        if token is not _VALIDATION_TOKEN:
            raise TypeError("validated endpoint objects are created by the endpoint validator")
        object.__setattr__(self, "_fields", _freeze(fields))
        object.__setattr__(self, "_token", token)

    def __getitem__(self, key: str) -> Any:
        return self._fields[key]

    def __iter__(self) -> Iterator[str]:
        return iter(self._fields)

    def __len__(self) -> int:
        return len(self._fields)

    def __setattr__(self, _name: str, _value: Any) -> None:
        raise TypeError("validated endpoint fields are immutable")

    def __delattr__(self, _name: str) -> None:
        raise TypeError("validated endpoint fields are immutable")

    def public(self) -> dict[str, Any]:
        """Return a detached copy for the qualification manifest only."""
        return _thaw(self._fields)


def _make_validated_endpoint(fields: dict[str, Any]) -> _ValidatedEndpoint:
    return _ValidatedEndpoint(fields, _VALIDATION_TOKEN)


def endpoint_public_fields(fields: object) -> dict[str, Any]:
    return _require_validated(fields, "endpoint").public()


def _runner_module() -> Any:
    path = ROOT / "docs/unified-ops/results/0362h/run.py"
    spec = importlib.util.spec_from_file_location("litchi_w8_qualification_runner", path)
    if spec is None or spec.loader is None:
        raise fail(f"cannot load the existing W8 endpoint validator: {path}")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as error:
        raise fail(f"cannot load the existing W8 endpoint validator: {error}") from error
    return module


def validate_endpoint_manifest(path: Path, binary: Path, label: str) -> _ValidatedEndpoint:
    """Run the complete runner endpoint validator and retain its opaque result."""
    try:
        fields = _runner_module().endpoint_manifest(path.resolve(), binary.resolve(), label)
    except Exception as error:
        if isinstance(error, ProjectionPolicyError):
            raise
        raise fail(f"{label} endpoint failed the complete W8 validator: {error}") from error
    if not isinstance(fields, dict) or not fields.get("manifest") or not fields.get("routes"):
        raise fail(f"{label} endpoint validator returned incomplete fields")
    return _make_validated_endpoint(fields)


def _validated_manifest(path: Path, label: str) -> _ValidatedEndpoint:
    value = _load_json(path.resolve(), f"{label} endpoint")
    binary = value.get("binary")
    if not isinstance(binary, dict) or not isinstance(binary.get("path"), str):
        raise fail(f"{label} endpoint has no binary path for complete validation")
    return validate_endpoint_manifest(path, Path(binary["path"]), label)


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
        "before_features": list(BEFORE_FEATURES),
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
    expected_features = {"before": BEFORE_FEATURES, "after": FEATURES}
    for side, projection in (("before", before), ("after", after)):
        fixed = _load_fixed_projection(side)
        if projection != fixed:
            raise fail(f"{side} projection does not match the reviewed fixed pre/post image")
        if projection.get("features") != {"package": "litchi-ops", "selected": list(expected_features[side]), "all_features": True}:
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


def _check_fields_fixed(endpoint: dict[str, Any], fields: Mapping[str, Any], side: str) -> None:
    """Check a runner-validated endpoint, including a fresh source hash."""
    expected = FIXED_INPUTS[side]
    if endpoint.get("schema") != "litchi_file_endpoint.v1":
        raise fail(f"{side} endpoint schema is invalid")
    manifest = fields.get("manifest")
    if not isinstance(manifest, Mapping) or not isinstance(manifest.get("path"), str):
        raise fail(f"{side} runner fields have no endpoint manifest identity")
    actual_manifest = artifact(Path(manifest["path"]))
    if (actual_manifest["bytes"], actual_manifest["sha256"]) != (manifest.get("bytes"), manifest.get("sha256")):
        raise fail(f"{side} endpoint manifest identity is stale")
    if fields.get("source_commit") != expected["source_commit"]:
        raise fail(f"{side} source commit is not the reviewed full commit")
    if endpoint.get("source_commit") != fields.get("source_commit") or endpoint.get("dependency_projection_sha256") != fields.get("dependency_projection_sha256"):
        raise fail(f"{side} endpoint provenance is not bound to the runner fields")
    for name in ("source", "workspace_manifest", "package_manifest", "toolchain", "lock", "probe", "fixture"):
        actual = fields.get(name)
        fixed = expected.get(name)
        if fixed is None:
            fixed = {"bytes": PROBE_BYTES, "sha256": PROBE_SHA256} if name == "probe" else {"bytes": FIXTURE_BYTES, "sha256": FIXTURE_SHA256}
        if not isinstance(actual, Mapping) or (actual.get("bytes"), actual.get("sha256")) != (fixed["bytes"], fixed["sha256"]):
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
            actual = artifact(Path(row["path"]))
            if (actual["bytes"], actual["sha256"]) != (row.get("bytes"), row.get("sha256")):
                raise fail(f"{side} dependency projection artifact identity is stale")
            return row
        if name == "source_assembly":
            actual = artifact(Path(row["path"]))
            if (actual["bytes"], actual["sha256"]) != (row.get("bytes"), row.get("sha256")):
                raise fail(f"{side} source assembly artifact identity is stale")
            return row
        if name in ("build_receipt", "validation"):
            actual = artifact(Path(row["path"]))
            if (actual["bytes"], actual["sha256"]) != (row.get("bytes"), row.get("sha256")):
                raise fail(f"{side} {name} artifact identity is stale")
            return row
        field = fields.get(name, {})
        if not isinstance(field, Mapping) or (row.get("bytes"), row.get("sha256")) != (field.get("bytes"), field.get("sha256")):
            raise fail(f"{side} endpoint {name} identity differs from the runner fields")
        actual = artifact(Path(row["path"]))
        if (actual["bytes"], actual["sha256"]) != (row.get("bytes"), row.get("sha256")):
            raise fail(f"{side} {name} artifact identity is stale")
        if name == "binary" and not os.access(Path(row["path"]), os.X_OK):
            raise fail(f"{side} endpoint binary is no longer executable")
        return row

    for name in ("source", "lock", "probe", "fixture", "dependency_projection", "workspace_manifest", "package_manifest", "toolchain", "binary", "build_receipt", "validation", "source_assembly"):
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
    try:
        if Path(source_archive["path"]).resolve() != Path(fields["source"]["path"]).resolve():
            raise fail(f"{side} source assembly archive path differs from the endpoint source")
    except (KeyError, OSError) as error:
        raise fail(f"{side} source assembly archive path is unavailable") from error
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
        try:
            actual_patch = artifact(Path(patch["path"]))
        except (KeyError, ProjectionPolicyError) as error:
            raise fail(f"{side} source assembly adapter patch artifact is unavailable") from error
        if actual_patch["bytes"] != patch["bytes"] or actual_patch["sha256"] != patch["sha256"]:
            raise fail(f"{side} source assembly adapter patch artifact identity is stale")


def _require_validated(fields: object, side: str) -> _ValidatedEndpoint:
    if not isinstance(fields, _ValidatedEndpoint) or fields._token is not _VALIDATION_TOKEN:
        raise fail(f"{side} endpoint was not returned by the complete W8 endpoint validator")
    return fields


def _receipt_from_validated(policy: dict[str, Any], before: object, after: object) -> dict[str, Any]:
    before_fields = _require_validated(before, "before")
    after_fields = _require_validated(after, "after")
    before_value = _load_json(Path(before_fields["manifest"]["path"]), "before endpoint")
    after_value = _load_json(Path(after_fields["manifest"]["path"]), "after endpoint")
    _check_fields_fixed(before_value, before_fields, "before")
    _check_fields_fixed(after_value, after_fields, "after")
    before_projection = _load_json(Path(before_value["dependency_projection"]["path"]), "before projection")
    after_projection = _load_json(Path(after_value["dependency_projection"]["path"]), "after projection")
    projection_delta = validate_projection_pair(before_projection, after_projection)
    before_fixture = _load_json(Path(before_value["fixture"]["path"]), "before fixture")
    after_fixture = _load_json(Path(after_value["fixture"]["path"]), "after fixture")
    before_routes = provenance.fixture_routes(before_fixture, "before fixture")
    after_routes = provenance.fixture_routes(after_fixture, "after fixture")
    if (
        before_routes != after_routes
        or before_routes != list(before_fields["routes"])
        or after_routes != list(after_fields["routes"])
        or len(before_routes) != 109
    ):
        raise fail("before/after route order is not the exact reviewed 109-route order")

    def endpoint_part(fields: _ValidatedEndpoint) -> dict[str, Any]:
        return {
            "manifest": dict(fields["manifest"]),
            "source_commit": fields["source_commit"],
            "source_sha256": fields["source"]["sha256"],
            "workspace_manifest_sha256": fields["workspace_manifest"]["sha256"],
            "package_manifest_sha256": fields["package_manifest"]["sha256"],
            "lock_sha256": fields["lock"]["sha256"],
            "projection_sha256": fields["dependency_projection_sha256"],
        }

    return {
        "schema": PAIR_SCHEMA,
        "policy": {
            "schema": policy["schema"],
            "version": policy["policy_version"],
            "sha256": REVIEWED_POLICY_SHA256,
        },
        "before": endpoint_part(before_fields),
        "after": endpoint_part(after_fields),
        "workload": {
            "probe_sha256": PROBE_SHA256,
            "fixture_sha256": FIXTURE_SHA256,
            "route_count": 109,
            "routes": before_routes,
        },
        "build": {
            "target": TARGET,
            "profile": PROFILE,
            "features": {
                "before": list(BEFORE_FEATURES),
                "after": list(FEATURES),
            },
            "toolchain_sha256": TOOLCHAIN_SHA256,
        },
        "projection": projection_delta,
        "status": "verified_protocol_only",
    }


def verify_pair(policy_path: Path, before_path: Path, after_path: Path) -> dict[str, Any]:
    """Run the complete endpoint validator and return the canonical pair receipt."""
    policy = load_reviewed_policy(policy_path)
    before = _validated_manifest(before_path, "before")
    after = _validated_manifest(after_path, "after")
    return _receipt_from_validated(policy, before, after)


def write_pair_receipt(policy_path: Path, before_path: Path, after_path: Path, output: Path) -> dict[str, Any]:
    receipt = verify_pair(policy_path, before_path, after_path)
    output = output.resolve()
    if output.exists():
        raise fail(f"refusing to overwrite pair receipt: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(canonical(receipt))
    return {"path": str(output), "bytes": output.stat().st_size, "sha256": digest(output)}


def verify_pair_receipt(policy_path: Path, receipt_path: Path, before_fields: object, after_fields: object) -> dict[str, Any]:
    """Verify a precomputed pair receipt against already-validated endpoints.

    The runner has just hashed the large source archives while checking each
    endpoint.  This path checks the endpoint again, including source archive
    hashes, before timing starts; the cost is outside all measured routes.
    """
    policy = load_reviewed_policy(policy_path)
    receipt = _load_json(receipt_path.resolve(), "original-to-final pair receipt")
    expected = _receipt_from_validated(policy, before_fields, after_fields)
    if receipt != expected:
        raise fail("pair receipt metadata is stale or does not bind every endpoint/build/workload field")
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
