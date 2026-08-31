#!/usr/bin/env python3
"""Run the non-iWork release gate without selecting the iWork workspace.

The workspace deliberately keeps the Apple iWork packages in the same Cargo
workspace as the Office and ODF packages.  This gate is the cheap, explicit
non-iWork slice: it derives the package and facade-feature sets from Cargo
metadata, checks every selected dependency tree, and then runs the requested
Cargo operation with the 45 ordinary packages plus the safe facade feature
closure.  Library tests and the other high-memory modes serialize package
roots; library tests clean each successful root's artifacts to keep its test
target from retaining every package's test binaries at once.

All Cargo invocations are argv lists.  In particular, feature names and
package exclusions never pass through a shell.  ``verify`` does not build
targets; it performs metadata and ``cargo tree`` checks and may resolve the
workspace or update the ignored root lockfile.  Other execution modes run one
command per bulk package root and separate facade commands using an isolated
target directory by default.  The facade's actual default-feature closure is
checked in its own invocation, in addition to each safe feature and the
combined explicit ``--no-default-features`` closure.
"""

from __future__ import annotations

import argparse
import json
import os
import posixpath
import platform
import re
import shlex
import signal
import stat
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass
from pathlib import Path, PureWindowsPath
from typing import Any, Callable, Iterable, Mapping, Sequence
from urllib.parse import unquote, urlsplit


ROOT = Path(__file__).resolve().parents[1]
FACADE_PACKAGE = "litchi"
PY_FACADE_PACKAGE = "litchi-py"
IWORK_LEAF_PACKAGES = frozenset(
    {
        "litchi-keynote",
        "litchi-numbers",
        "litchi-numbers-wire",
        "litchi-pages",
    }
)
IWORK_PACKAGE_PREFIX = "litchi-iwa"
EXPECTED_IWORK_PACKAGES = 17
EXPECTED_EXCLUDED_PACKAGES = 18
EXPECTED_BULK_PACKAGES = 45

# The release gate is intentionally audited against the current workspace
# topology, not merely package counts.  A rename, swap, or newly added member
# must update this inventory and its review before it can enter the gate.
EXPECTED_IWORK_PACKAGE_NAMES = frozenset(
    {
        "litchi-iwa",
        "litchi-iwa-archive",
        "litchi-iwa-cache",
        "litchi-iwa-common",
        "litchi-iwa-core",
        "litchi-iwa-detect",
        "litchi-iwa-graph",
        "litchi-iwa-index",
        "litchi-iwa-package",
        "litchi-iwa-protos",
        "litchi-iwa-structured",
        "litchi-iwa-text",
        "litchi-iwa-text-wire",
        "litchi-keynote",
        "litchi-numbers",
        "litchi-numbers-wire",
        "litchi-pages",
    }
)
EXPECTED_EXCLUDED_PACKAGE_NAMES = frozenset(
    {
        "litchi-iwa",
        "litchi-iwa-archive",
        "litchi-iwa-cache",
        "litchi-iwa-common",
        "litchi-iwa-core",
        "litchi-iwa-detect",
        "litchi-iwa-graph",
        "litchi-iwa-index",
        "litchi-iwa-package",
        "litchi-iwa-protos",
        "litchi-iwa-structured",
        "litchi-iwa-text",
        "litchi-iwa-text-wire",
        "litchi-keynote",
        "litchi-numbers",
        "litchi-numbers-wire",
        "litchi-pages",
        "litchi-py",
    }
)
EXPECTED_BULK_PACKAGE_NAMES = frozenset(
    {
        "litchi-biff",
        "litchi-cfb",
        "litchi-codepage",
        "litchi-core",
        "litchi-crypto",
        "litchi-doc",
        "litchi-docx",
        "litchi-drawingml",
        "litchi-eval",
        "litchi-fonts",
        "litchi-formula",
        "litchi-imgconv",
        "litchi-markdown",
        "litchi-odb",
        "litchi-odc",
        "litchi-odf",
        "litchi-odf-common",
        "litchi-odf-formula",
        "litchi-odg",
        "litchi-odi",
        "litchi-odm",
        "litchi-odp",
        "litchi-odraw",
        "litchi-ods",
        "litchi-odt",
        "litchi-ograph",
        "litchi-ole-common",
        "litchi-ooxml-common",
        "litchi-opc",
        "litchi-oth",
        "litchi-ppt",
        "litchi-pptx",
        "litchi-rtf",
        "litchi-sheet",
        "litchi-sign",
        "litchi-slide",
        "litchi-spreadsheet-drawing",
        "litchi-vba",
        "litchi-word",
        "litchi-xls",
        "litchi-xlsb",
        "litchi-xlsx",
        "soapberry-zip",
        "xml-minifier",
        "xml-minifier-macros",
    }
)
EXPECTED_WORKSPACE_PACKAGE_NAMES = frozenset(
    {
        "litchi",
        "litchi-biff",
        "litchi-cfb",
        "litchi-codepage",
        "litchi-core",
        "litchi-crypto",
        "litchi-doc",
        "litchi-docx",
        "litchi-drawingml",
        "litchi-eval",
        "litchi-fonts",
        "litchi-formula",
        "litchi-imgconv",
        "litchi-iwa",
        "litchi-iwa-archive",
        "litchi-iwa-cache",
        "litchi-iwa-common",
        "litchi-iwa-core",
        "litchi-iwa-detect",
        "litchi-iwa-graph",
        "litchi-iwa-index",
        "litchi-iwa-package",
        "litchi-iwa-protos",
        "litchi-iwa-structured",
        "litchi-iwa-text",
        "litchi-iwa-text-wire",
        "litchi-keynote",
        "litchi-markdown",
        "litchi-numbers",
        "litchi-numbers-wire",
        "litchi-odb",
        "litchi-odc",
        "litchi-odf",
        "litchi-odf-common",
        "litchi-odf-formula",
        "litchi-odg",
        "litchi-odi",
        "litchi-odm",
        "litchi-odp",
        "litchi-odraw",
        "litchi-ods",
        "litchi-odt",
        "litchi-ograph",
        "litchi-ole-common",
        "litchi-ooxml-common",
        "litchi-opc",
        "litchi-oth",
        "litchi-pages",
        "litchi-ppt",
        "litchi-pptx",
        "litchi-py",
        "litchi-rtf",
        "litchi-sheet",
        "litchi-sign",
        "litchi-slide",
        "litchi-spreadsheet-drawing",
        "litchi-vba",
        "litchi-word",
        "litchi-xls",
        "litchi-xlsb",
        "litchi-xlsx",
        "soapberry-zip",
        "xml-minifier",
        "xml-minifier-macros",
    }
)

# ``prost`` is the actual wire-format implementation used by the iWork
# protobuf crate.  Keep the family check broad so a newly split prost helper
# cannot silently enter a purportedly protobuf-free tree.
PROTOBUF_PACKAGE_RE = re.compile(r"^(?:prost(?:[-_]|$)|protobuf(?:[-_]|$))")
TREE_PACKAGE_RE = re.compile(
    r"^([A-Za-z0-9][A-Za-z0-9_-]*) v([^\s(]+|\*)(?:\s+(.*))?$"
)
TARGET_PLATFORM_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
WINDOWS_DRIVE_PATH_RE = re.compile(r"^[A-Za-z]:[\\/]")
WINDOWS_UNC_PATH_RE = re.compile(r"^(?:\\\\|//)[^\\/]+[\\/].+")
DEPENDENCY_ALIAS_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_-]*$")
PACKAGE_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]*$")
REGISTRY_SOURCE = "registry+https://github.com/rust-lang/crates.io-index"
TARGET_KINDS = frozenset(
    {
        "lib",
        "proc-macro",
        "cdylib",
        "bin",
        "example",
        "test",
        "bench",
        "custom-build",
    }
)
REPORT_VERSION = 1
RECORDED_ENV_KEYS = frozenset(
    {
        "CARGO_BUILD_JOBS",
        "CARGO_INCREMENTAL",
        "CARGO_PROFILE_DEV_DEBUG",
        "CARGO_PROFILE_TEST_DEBUG",
        "CARGO_TARGET_DIR",
        "RUSTDOCFLAGS",
        "RUSTFLAGS",
    }
)
RSS_POLL_INTERVAL_SECONDS = 0.02
RSS_SAMPLE_INTERVAL_MS = 20
RSS_MEASUREMENT = "sampled_summed_descendant_process_tree_high_water"
MAX_REPORT_ERROR_LENGTH = 4096
MAX_RECORDED_ENV_VALUE_LENGTH = 512
CAPTURE_CHUNK_BYTES = 64 * 1024
CAPTURE_OUTPUT_LIMIT_BYTES = 8 * 1024 * 1024
CAPTURE_POLL_INTERVAL_SECONDS = 0.02
PROCESS_TERM_GRACE_SECONDS = 1.0
TARGET_SCAN_MAX_ENTRIES = 100_000
TARGET_SCAN_MAX_DIRECTORIES = 20_000
TARGET_SCAN_MAX_DEPTH = 64

MODES = (
    "print",
    "verify",
    "check",
    "clippy",
    "doc",
    "lib-tests",
    "doc-tests",
    "deprecated",
)


class GateError(RuntimeError):
    """Raised for a topology, metadata, or Cargo gate failure."""


class CaptureLimitError(GateError):
    """Raised when a bounded Cargo probe output exceeds its byte limit."""

    def __init__(self, argv: Sequence[str], stream: str, limit_bytes: int) -> None:
        self.argv = tuple(argv)
        self.stream = stream
        self.limit_bytes = limit_bytes
        super().__init__(
            f"{shlex.join(self.argv)} {stream} output exceeded the "
            f"{limit_bytes}-byte capture limit"
        )


@dataclass(frozen=True)
class Dependency:
    """The workspace-relevant part of one Cargo dependency declaration."""

    name: str
    features: frozenset[str]
    uses_default_features: bool
    target: str | None
    source: str | None = None
    path: Path | None = None
    kind: str | None = None
    optional: bool = False
    rename: str | None = None
    req: str = "*"
    package_name: str | None = None

    @property
    def original_name(self) -> str:
        """Return the package name, independent of any dependency alias."""

        return self.package_name or self.name

    @property
    def alias(self) -> str:
        """Return the name used by Cargo feature references and Rust code."""

        return self.rename or self.name


@dataclass(frozen=True)
class Target:
    """The local identity of one Cargo target."""

    name: str
    kinds: frozenset[str]
    crate_types: frozenset[str]
    src_path: Path


@dataclass(frozen=True)
class Package:
    """A workspace package projected from ``cargo metadata``."""

    name: str
    features: Mapping[str, tuple[str, ...]]
    dependencies: tuple[Dependency, ...]
    package_id: str
    version: str
    source: None
    manifest_path: Path
    targets: tuple[Target, ...]


@dataclass(frozen=True)
class WorkspacePlan:
    """The derived package and facade feature closure for the gate."""

    packages: Mapping[str, Package]
    facade_features: Mapping[str, tuple[str, ...]]
    iwork_packages: frozenset[str]
    unsafe_facade_dependents: frozenset[str]
    excluded_packages: frozenset[str]
    bulk_packages: frozenset[str]
    safe_facade_features: frozenset[str]
    unsafe_facade_features: frozenset[str]


@dataclass(frozen=True)
class CommandSpec:
    """One Cargo command and its non-shell environment additions."""

    scope: str
    argv: tuple[str, ...]
    env: tuple[tuple[str, str], ...] = ()


@dataclass(frozen=True)
class TargetFootprint:
    """A logical regular-file inventory of the Cargo target directory."""

    status: str
    regular_file_bytes: int | None
    regular_file_count: int | None
    reason: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "regular_file_bytes": self.regular_file_bytes,
            "regular_file_count": self.regular_file_count,
            "reason": self.reason,
        }


@dataclass(frozen=True)
class ChildRss:
    """Cumulative sampled RSS for a child process tree."""

    high_water_bytes: int | None
    platform: str
    scope: str
    status: str = "unavailable"
    reason: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "measurement": RSS_MEASUREMENT,
            "high_water_bytes": self.high_water_bytes,
            "platform": self.platform,
            "scope": self.scope,
            "status": self.status,
            "reason": self.reason,
            "sample_interval_ms": RSS_SAMPLE_INTERVAL_MS,
        }


@dataclass(frozen=True)
class CommandExecution:
    """The correctness-relevant result of one non-shell child command."""

    returncode: int
    elapsed_ns: int
    child_rss: ChildRss


@dataclass(frozen=True)
class PhaseRecord:
    """One ordered execution phase in the optional JSON report."""

    mode: str
    index: int
    scope: str
    argv: tuple[str, ...]
    status: str
    returncode: int | None
    elapsed_ns: int
    target_before: TargetFootprint
    target_after: TargetFootprint
    child_rss: ChildRss
    env_keys: tuple[str, ...]
    cargo_env: tuple[tuple[str, str], ...] = ()

    def as_dict(self) -> dict[str, Any]:
        return {
            "mode": self.mode,
            "index": self.index,
            "scope": self.scope,
            "argv": list(self.argv),
            "status": self.status,
            "returncode": self.returncode,
            "elapsed_ns": self.elapsed_ns,
            "target_before": self.target_before.as_dict(),
            "target_after": self.target_after.as_dict(),
            "child_rss": self.child_rss.as_dict(),
            "env_keys": list(self.env_keys),
            "cargo_env": dict(self.cargo_env),
        }


class TreeEntries(dict[str, frozenset[Path | PureWindowsPath | None]]):
    """Cargo tree identities with versions retained alongside path values."""

    def __init__(
        self,
        paths: Mapping[str, frozenset[Path | PureWindowsPath | None]],
        versions: Mapping[str, frozenset[str]],
    ) -> None:
        super().__init__(paths)
        self.versions = versions


def _is_windows_absolute_path(value: str) -> bool:
    """Recognize Windows drive and UNC paths even when running on Unix."""

    return bool(
        WINDOWS_DRIVE_PATH_RE.match(value) or WINDOWS_UNC_PATH_RE.match(value)
    )


def _is_absolute_path_text(value: str) -> bool:
    """Return whether a path string is absolute on either supported host."""

    return Path(value).is_absolute() or _is_windows_absolute_path(value)


def _portable_canonical_path(value: Any, description: str = "path") -> str:
    """Normalize an absolute path without assuming the host path dialect.

    Cargo emits native filesystem paths in tree output and file URLs.  The
    gate itself resolves paths on its host, but this helper deliberately keeps
    Windows drive/UNC parsing platform-neutral so malformed identities can be
    tested on Linux as well.
    """

    if not isinstance(value, str):
        raise GateError(f"{description} is not a string")
    text = unquote(value).replace("\\", "/")
    if not _is_absolute_path_text(text):
        raise GateError(f"{description} is not an absolute path: {value!r}")
    if WINDOWS_DRIVE_PATH_RE.match(text):
        drive = text[0].upper()
        tail = posixpath.normpath(text[2:])
        if not tail.startswith("/"):
            tail = "/" + tail
        return drive + ":" + tail
    if text.startswith("//"):
        normalized = posixpath.normpath(text)
        return normalized if normalized.startswith("//") else "/" + normalized
    return posixpath.normpath(text)


def _file_uri_path(package_id: str) -> tuple[str, str]:
    """Extract a portable local path and fragment from a Cargo path ID."""

    if not isinstance(package_id, str):
        raise GateError("workspace package ID is not a string")
    try:
        parsed = urlsplit(package_id)
    except ValueError as error:
        raise GateError(f"workspace package ID is not a valid path URI: {package_id!r}") from error
    if parsed.scheme != "path+file" or parsed.query or parsed.username or parsed.password:
        raise GateError(f"workspace package ID is not a local path ID: {package_id!r}")
    if parsed.netloc:
        # URL parsing treats ``path+file://C:/dir`` as netloc ``C:`` and path
        # ``/dir``.  A non-drive netloc is a UNC server name instead.
        if re.fullmatch(r"[A-Za-z]:", parsed.netloc):
            raw_path = parsed.netloc + parsed.path
        else:
            raw_path = "//" + parsed.netloc + parsed.path
    else:
        raw_path = parsed.path
    raw_path = unquote(raw_path)
    # The canonical file URL spelling for a Windows drive is
    # ``file:///C:/...``; the leading URL slash is not part of the path.
    if re.match(r"^/[A-Za-z]:[\\/]", raw_path):
        raw_path = raw_path[1:]
    if not raw_path:
        raise GateError(f"workspace package ID has no local path: {package_id!r}")
    return _portable_canonical_path(raw_path, "workspace package ID path"), unquote(
        parsed.fragment
    )


def _resolved_absolute_path(value: Any, description: str) -> Path:
    """Resolve a metadata path while rejecting relative or malformed paths."""

    if not isinstance(value, str):
        raise GateError(f"{description} is not a string")
    path = Path(value)
    if not path.is_absolute() and not _is_windows_absolute_path(value):
        raise GateError(f"{description} is not an absolute path: {value!r}")
    if _is_windows_absolute_path(value) and os.name != "nt":
        raise GateError(
            f"{description} uses a Windows path on a non-Windows host: {value!r}"
        )
    try:
        return path.resolve()
    except OSError as error:
        raise GateError(f"{description} cannot be resolved: {value!r}: {error}") from error


def _package_id_path(package_id: str) -> Path:
    """Return the filesystem path encoded by a Cargo local package ID."""

    portable_path, _ = _file_uri_path(package_id)
    return _resolved_absolute_path(portable_path, "workspace package ID path")


def _package_id_suffix(package_id: str) -> str:
    """Return the decoded Cargo package ID suffix after ``#``."""

    _, suffix = _file_uri_path(package_id)
    if not suffix:
        raise GateError(f"workspace package ID has no package suffix: {package_id!r}")
    return suffix


def _metadata_target(raw_target: Mapping[str, Any], package_name: str, package_dir: Path) -> Target:
    """Validate one target's local source identity and target kind."""

    name = raw_target.get("name")
    raw_kinds = raw_target.get("kind")
    raw_crate_types = raw_target.get("crate_types")
    if not isinstance(name, str) or not name:
        raise GateError(f"package {package_name!r} has a malformed target name")
    if not isinstance(raw_kinds, list) or not raw_kinds or not all(
        isinstance(kind, str) for kind in raw_kinds
    ):
        raise GateError(f"package {package_name!r} target {name!r} has malformed kinds")
    if not isinstance(raw_crate_types, list) or not raw_crate_types or not all(
        isinstance(crate_type, str) for crate_type in raw_crate_types
    ):
        raise GateError(
            f"package {package_name!r} target {name!r} has malformed crate types"
        )
    kinds = frozenset(raw_kinds)
    unknown_kinds = kinds - TARGET_KINDS
    if unknown_kinds:
        raise GateError(
            f"package {package_name!r} target {name!r} has unknown kinds: "
            + ", ".join(sorted(unknown_kinds))
        )
    src_path = _resolved_absolute_path(
        raw_target.get("src_path"),
        f"package {package_name!r} target {name!r} source path",
    )
    try:
        src_path.relative_to(package_dir)
    except ValueError as error:
        raise GateError(
            f"package {package_name!r} target {name!r} source escapes its package: "
            f"{src_path}"
        ) from error
    if src_path == package_dir:
        raise GateError(
            f"package {package_name!r} target {name!r} source is its package directory"
        )
    return Target(name, kinds, frozenset(raw_crate_types), src_path)


def _metadata_dependency(
    raw_dependency: Mapping[str, Any], package_name: str
) -> Dependency:
    """Validate and project one Cargo metadata dependency declaration."""

    required_fields = {
        "name",
        "source",
        "req",
        "kind",
        "rename",
        "optional",
        "uses_default_features",
        "features",
        "target",
        "registry",
    }
    missing_fields = required_fields - set(raw_dependency)
    if missing_fields:
        raise GateError(
            f"package {package_name!r} dependency metadata is missing: "
            + ", ".join(sorted(missing_fields))
        )
    dependency_name = raw_dependency.get("name")
    dependency_features = raw_dependency.get("features")
    uses_default_features = raw_dependency.get("uses_default_features")
    dependency_target = raw_dependency.get("target")
    dependency_source = raw_dependency.get("source")
    dependency_req = raw_dependency.get("req")
    dependency_kind = raw_dependency.get("kind")
    dependency_optional = raw_dependency.get("optional")
    dependency_rename = raw_dependency.get("rename")
    dependency_registry = raw_dependency.get("registry")
    dependency_package = raw_dependency.get("package")

    if not isinstance(dependency_name, str) or not PACKAGE_NAME_RE.fullmatch(
        dependency_name
    ):
        raise GateError(f"package {package_name!r} has a dependency without a name")
    if not isinstance(dependency_req, str) or not dependency_req:
        raise GateError(
            f"package {package_name!r} dependency {dependency_name!r} has malformed req"
        )
    if not isinstance(dependency_features, list) or not all(
        isinstance(value, str) for value in dependency_features
    ):
        raise GateError(
            f"package {package_name!r} dependency {dependency_name!r} has invalid features"
        )
    if not isinstance(uses_default_features, bool):
        raise GateError(
            f"package {package_name!r} dependency {dependency_name!r} has malformed "
            "uses_default_features"
        )
    if dependency_target is not None and (
        not isinstance(dependency_target, str) or not dependency_target
    ):
        raise GateError(
            f"package {package_name!r} dependency {dependency_name!r} has malformed target"
        )
    if dependency_source is not None and not isinstance(dependency_source, str):
        raise GateError(
            f"package {package_name!r} dependency {dependency_name!r} has malformed source"
        )
    if dependency_kind not in {None, "dev", "build"}:
        raise GateError(
            f"package {package_name!r} dependency {dependency_name!r} has unknown kind "
            f"{dependency_kind!r}"
        )
    if not isinstance(dependency_optional, bool):
        raise GateError(
            f"package {package_name!r} dependency {dependency_name!r} has malformed optional"
        )
    if dependency_rename is not None and (
        not isinstance(dependency_rename, str)
        or not DEPENDENCY_ALIAS_RE.fullmatch(dependency_rename)
        or dependency_rename == dependency_name
    ):
        raise GateError(
            f"package {package_name!r} dependency {dependency_name!r} has malformed rename"
        )
    if dependency_registry is not None and not isinstance(dependency_registry, str):
        raise GateError(
            f"package {package_name!r} dependency {dependency_name!r} has malformed registry"
        )
    if dependency_package is not None and (
        not isinstance(dependency_package, str)
        or not PACKAGE_NAME_RE.fullmatch(dependency_package)
    ):
        raise GateError(
            f"package {package_name!r} dependency {dependency_name!r} has malformed package"
        )

    dependency_path: Path | None = None
    if "path" in raw_dependency:
        dependency_path = _resolved_absolute_path(
            raw_dependency.get("path"),
            f"package {package_name!r} dependency {dependency_name!r} path",
        )
        if dependency_source is not None:
            raise GateError(
                f"package {package_name!r} dependency {dependency_name!r} has both "
                "a local path and a source"
            )
        if dependency_registry is not None:
            raise GateError(
                f"package {package_name!r} local dependency {dependency_name!r} has a registry"
            )
    else:
        if dependency_source != REGISTRY_SOURCE:
            raise GateError(
                f"package {package_name!r} dependency {dependency_name!r} is not a "
                f"crates.io dependency: {dependency_source!r}"
            )
        if dependency_registry is not None:
            raise GateError(
                f"package {package_name!r} dependency {dependency_name!r} has an "
                "unexpected registry override"
            )

    return Dependency(
        name=dependency_name,
        features=frozenset(dependency_features),
        uses_default_features=uses_default_features,
        target=dependency_target,
        source=dependency_source,
        path=dependency_path,
        kind=dependency_kind,
        optional=dependency_optional,
        rename=dependency_rename,
        req=dependency_req,
        package_name=dependency_package,
    )


def _metadata_packages(metadata: Mapping[str, Any]) -> dict[str, Package]:
    """Project workspace members and validate their local Cargo identities."""

    raw_packages = metadata.get("packages")
    workspace_members = metadata.get("workspace_members")
    if not isinstance(raw_packages, list) or not isinstance(workspace_members, list):
        raise GateError("cargo metadata lacks packages or workspace_members")
    if metadata.get("version") != 1:
        raise GateError("cargo metadata must use format version 1")
    workspace_default_members = metadata.get("workspace_default_members")
    if not isinstance(workspace_default_members, list):
        raise GateError("cargo metadata lacks workspace_default_members")
    workspace_root = _resolved_absolute_path(
        metadata.get("workspace_root"), "cargo metadata workspace root"
    )
    if workspace_root != ROOT.resolve():
        raise GateError(
            f"cargo metadata workspace root {workspace_root} is not the gate root {ROOT}"
        )

    member_ids: list[str] = []
    for package_id in workspace_members:
        if not isinstance(package_id, str):
            raise GateError("cargo metadata has a non-string workspace member ID")
        member_ids.append(package_id)
    if len(set(member_ids)) != len(member_ids):
        raise GateError("cargo metadata repeats a workspace member ID")
    default_member_ids: list[str] = []
    for package_id in workspace_default_members:
        if not isinstance(package_id, str):
            raise GateError("cargo metadata has a non-string default member ID")
        default_member_ids.append(package_id)
    if len(set(default_member_ids)) != len(default_member_ids):
        raise GateError("cargo metadata repeats a workspace default member ID")
    if frozenset(default_member_ids) != frozenset(member_ids):
        raise GateError("workspace_default_members does not match workspace_members")

    by_id: dict[str, Mapping[str, Any]] = {}
    for raw_package in raw_packages:
        if not isinstance(raw_package, Mapping):
            raise GateError("cargo metadata contains a malformed package")
        package_id = raw_package.get("id")
        if not isinstance(package_id, str):
            raise GateError("cargo metadata package has no string id")
        if package_id in by_id:
            raise GateError(f"cargo metadata repeats package ID {package_id!r}")
        by_id[package_id] = raw_package
    if frozenset(by_id) != frozenset(member_ids):
        unexpected = sorted(set(by_id) - set(member_ids))
        missing = sorted(set(member_ids) - set(by_id))
        details: list[str] = []
        if unexpected:
            details.append("unexpected package IDs: " + ", ".join(unexpected))
        if missing:
            details.append("missing package IDs: " + ", ".join(missing))
        raise GateError("cargo metadata package/member identity mismatch (" + "; ".join(details) + ")")

    packages: dict[str, Package] = {}
    crates_root = ROOT / "crates"
    for package_id in member_ids:
        if package_id not in by_id:
            raise GateError(f"workspace member {package_id!r} is missing from metadata")
        raw_package = by_id[package_id]
        if raw_package.get("id") != package_id:
            raise GateError(f"workspace package ID key mismatch for {package_id!r}")
        if "source" not in raw_package or raw_package.get("source") is not None:
            raise GateError(
                f"workspace package {package_id!r} is not source-less/local"
            )
        name = raw_package.get("name")
        version = raw_package.get("version")
        raw_features = raw_package.get("features", {})
        raw_dependencies = raw_package.get("dependencies", [])
        raw_targets = raw_package.get("targets")
        if (
            not isinstance(name, str)
            or not name
            or not isinstance(version, str)
            or not version
            or not isinstance(raw_features, Mapping)
        ):
            raise GateError(f"package {name!r} has malformed feature metadata")
        if not isinstance(raw_dependencies, list):
            raise GateError(f"package {name!r} has malformed dependency metadata")
        if not isinstance(raw_targets, list) or not raw_targets:
            raise GateError(f"package {name!r} has no target metadata")

        manifest_path = _resolved_absolute_path(
            raw_package.get("manifest_path"), f"package {name!r} manifest path"
        )
        package_dir = manifest_path.parent
        if manifest_path.name != "Cargo.toml":
            raise GateError(f"package {name!r} manifest is not Cargo.toml")
        try:
            relative_package_dir = package_dir.relative_to(crates_root)
        except ValueError as error:
            raise GateError(
                f"package {name!r} manifest escapes the crates directory: {manifest_path}"
            ) from error
        if len(relative_package_dir.parts) != 1:
            raise GateError(f"package {name!r} is not a direct crates member")
        expected_directory = "pyo3-litchi" if name == PY_FACADE_PACKAGE else name
        if package_dir.name != expected_directory:
            raise GateError(
                f"package {name!r} manifest directory {package_dir.name!r} "
                f"does not match expected {expected_directory!r}"
            )
        id_path = _package_id_path(package_id)
        if id_path != package_dir:
            raise GateError(
                f"package {name!r} ID path {id_path} does not match manifest directory "
                f"{package_dir}"
            )
        id_suffix = _package_id_suffix(package_id)
        if id_suffix not in {version, f"{name}@{version}"}:
            raise GateError(
                f"package {name!r} ID suffix {id_suffix!r} does not match "
                f"name/version {name!r}/{version!r}"
            )

        features: dict[str, tuple[str, ...]] = {}
        for feature, values in raw_features.items():
            if not isinstance(feature, str) or not isinstance(values, list):
                raise GateError(f"package {name!r} has malformed feature {feature!r}")
            if not all(isinstance(value, str) for value in values):
                raise GateError(f"package {name!r} feature {feature!r} is not string-valued")
            features[feature] = tuple(values)

        dependencies: list[Dependency] = []
        for raw_dependency in raw_dependencies:
            if not isinstance(raw_dependency, Mapping):
                raise GateError(f"package {name!r} has a malformed dependency")
            dependencies.append(_metadata_dependency(raw_dependency, name))

        targets: list[Target] = []
        for raw_target in raw_targets:
            if not isinstance(raw_target, Mapping):
                raise GateError(f"package {name!r} has a malformed target")
            targets.append(_metadata_target(raw_target, name, package_dir))
        if name in packages:
            raise GateError(f"duplicate workspace package {name!r}")
        packages[name] = Package(
            name,
            features,
            tuple(dependencies),
            package_id,
            version,
            None,
            manifest_path,
            tuple(targets),
        )
    package_dirs = {
        package.manifest_path.parent: package.name for package in packages.values()
    }
    for package in packages.values():
        for dependency in package.dependencies:
            if dependency.path is None:
                continue
            dependency_dir = dependency.path.resolve()
            if dependency_dir not in package_dirs:
                raise GateError(
                    f"package {package.name!r} dependency {dependency.name!r} path "
                    f"{dependency_dir} is not a workspace package"
                )
            expected_name = package_dirs[dependency_dir]
            if dependency.original_name != expected_name:
                raise GateError(
                    f"package {package.name!r} dependency {dependency.name!r} path "
                    f"targets {expected_name!r}, not {dependency.original_name!r}"
                )
    return packages


def _iwork_packages(package_names: Iterable[str]) -> frozenset[str]:
    """Derive the iWork family from the accepted package naming topology."""

    return frozenset(
        name
        for name in package_names
        if name == IWORK_PACKAGE_PREFIX
        or name.startswith(f"{IWORK_PACKAGE_PREFIX}-")
        or name in IWORK_LEAF_PACKAGES
    )


def _require_exact_package_set(
    actual: Iterable[str], expected: frozenset[str], description: str
) -> None:
    """Reject an unreviewed package rename, swap, addition, or removal."""

    actual_set = frozenset(actual)
    if actual_set == expected:
        return
    unexpected = sorted(actual_set - expected)
    missing = sorted(expected - actual_set)
    details: list[str] = []
    if unexpected:
        details.append("unexpected: " + ", ".join(unexpected))
    if missing:
        details.append("missing: " + ", ".join(missing))
    raise GateError(f"{description} mismatch (" + "; ".join(details) + ")")


def _feature_references(
    feature_values: Mapping[str, Sequence[str]],
    dependency_aliases: Mapping[str, str] | None = None,
) -> tuple[dict[str, frozenset[str]], dict[str, frozenset[str]]]:
    """Split feature values into aliases and optional dependency references."""

    aliases: dict[str, set[str]] = {feature: set() for feature in feature_values}
    dependency_refs: dict[str, set[str]] = {feature: set() for feature in feature_values}
    feature_names = set(feature_values)
    dependency_aliases = dependency_aliases or {}

    def original_dependency_name(alias: str) -> str:
        return dependency_aliases.get(alias.rstrip("?"), alias.rstrip("?"))

    for feature, values in feature_values.items():
        for value in values:
            if value in feature_names:
                aliases[feature].add(value)
                continue
            if value.startswith("dep:"):
                dependency_refs[feature].add(
                    original_dependency_name(value.removeprefix("dep:"))
                )
                continue
            # A value such as litchi-pages/internal-iwork-source is a
            # feature forwarded to an optional dependency.  Its package name
            # is the part before the slash.
            dependency_alias, _, _ = value.partition("/")
            dependency_alias = dependency_alias.rstrip("?")
            if (
                dependency_alias.startswith("litchi-")
                or dependency_alias in dependency_aliases
                or "/" in value
            ):
                dependency_refs[feature].add(original_dependency_name(dependency_alias))
    return (
        {feature: frozenset(values) for feature, values in aliases.items()},
        {feature: frozenset(values) for feature, values in dependency_refs.items()},
    )


def _unsafe_features(
    feature_values: Mapping[str, Sequence[str]],
    iwork_packages: frozenset[str],
    dependency_aliases: Mapping[str, str] | None = None,
) -> frozenset[str]:
    """Return every facade feature whose expansion reaches an iWork package."""

    aliases, dependency_refs = _feature_references(feature_values, dependency_aliases)
    memo: dict[str, bool] = {}

    def reaches_iwork(feature: str, visiting: frozenset[str] = frozenset()) -> bool:
        if feature in memo:
            return memo[feature]
        if feature in visiting:
            raise GateError(f"cycle in litchi feature aliases at {feature!r}")
        result = bool(dependency_refs[feature] & iwork_packages)
        if not result:
            result = any(
                reaches_iwork(alias, visiting | {feature}) for alias in aliases[feature]
            )
        memo[feature] = result
        return result

    return frozenset(
        feature for feature in feature_values if reaches_iwork(feature)
    )


def _require_known_facade_dependency_references(
    feature_values: Mapping[str, Sequence[str]],
    dependency_aliases: Mapping[str, str],
) -> None:
    """Reject feature references that Cargo metadata cannot resolve."""

    _, dependency_references = _feature_references(feature_values, dependency_aliases)
    known_packages = frozenset(dependency_aliases.values())
    unknown = frozenset().union(*dependency_references.values()) - known_packages
    if unknown:
        raise GateError(
            "facade features reference unknown dependency packages: "
            + ", ".join(sorted(unknown))
        )


def _require_selected_root_targets(
    packages: Mapping[str, Package], selected: Iterable[str]
) -> None:
    """Require a library-like target for every root the gate invokes."""

    for name in sorted(selected):
        package = packages[name]
        library_targets = [
            target
            for target in package.targets
            if (
                "lib" in target.kinds
                and "lib" in target.crate_types
            )
            or (
                "proc-macro" in target.kinds
                and "proc-macro" in target.crate_types
            )
        ]
        if not library_targets:
            raise GateError(
                f"selected package root {name!r} has no local lib/proc-macro target"
            )


def _dependency_reaches_unsafe_facade(
    dependency: Dependency, unsafe_features: frozenset[str]
) -> bool:
    """Account for explicit features and Cargo's implicit default closure."""

    return bool(dependency.features & unsafe_features) or (
        dependency.uses_default_features and "default" in unsafe_features
    )


def _facade_dependency_aliases(dependencies: Iterable[Dependency]) -> dict[str, str]:
    """Map Cargo dependency names and rename aliases to original packages."""

    aliases: dict[str, str] = {}
    for dependency in dependencies:
        original = dependency.original_name
        for alias in {dependency.name, dependency.alias}:
            previous = aliases.setdefault(alias, original)
            if previous != original:
                raise GateError(
                    f"facade dependency alias {alias!r} names both {previous!r} and "
                    f"{original!r}"
                )
    return aliases


def _require_bulk_dependency_boundary(
    packages: Mapping[str, Package],
    bulk_packages: Iterable[str],
    excluded_packages: frozenset[str],
) -> None:
    """Reject every direct bulk edge to an excluded original package."""

    for package_name in sorted(bulk_packages):
        for dependency in packages[package_name].dependencies:
            if dependency.original_name not in excluded_packages:
                continue
            details = [f"original={dependency.original_name!r}"]
            if dependency.alias != dependency.original_name:
                details.append(f"alias={dependency.alias!r}")
            if dependency.target is not None:
                details.append(f"target={dependency.target!r}")
            if dependency.kind is not None:
                details.append(f"kind={dependency.kind!r}")
            details.append(f"optional={dependency.optional}")
            details.append(f"uses_default_features={dependency.uses_default_features}")
            raise GateError(
                f"bulk package {package_name!r} directly depends on excluded package "
                + f"{dependency.original_name!r} (" + ", ".join(details) + ")"
            )

def derive_plan(metadata: Mapping[str, Any]) -> WorkspacePlan:
    """Derive and validate the exact non-iWork package/feature plan."""

    packages = _metadata_packages(metadata)
    if FACADE_PACKAGE not in packages:
        raise GateError("workspace does not contain the litchi facade")
    _require_exact_package_set(
        packages,
        EXPECTED_WORKSPACE_PACKAGE_NAMES,
        "workspace package inventory",
    )
    facade = packages[FACADE_PACKAGE]
    if facade.features.get("default") != ():
        raise GateError(
            "facade default feature must remain explicitly empty; review its closure "
            "before changing the non-iWork gate"
        )
    iwork_packages = _iwork_packages(packages)
    _require_exact_package_set(
        iwork_packages,
        EXPECTED_IWORK_PACKAGE_NAMES,
        "iWork package inventory",
    )
    if len(iwork_packages) != EXPECTED_IWORK_PACKAGES:
        raise GateError(
            "expected "
            f"{EXPECTED_IWORK_PACKAGES} iWork packages, found {len(iwork_packages)}: "
            + ", ".join(sorted(iwork_packages))
        )

    dependency_aliases = _facade_dependency_aliases(facade.dependencies)
    _require_known_facade_dependency_references(facade.features, dependency_aliases)
    unsafe_features = _unsafe_features(
        facade.features, iwork_packages, dependency_aliases
    )
    safe_features = frozenset(set(facade.features) - unsafe_features)
    if safe_features | unsafe_features != frozenset(facade.features) or (
        safe_features & unsafe_features
    ):
        raise GateError("facade feature closure is not a disjoint partition")
    if not unsafe_features:
        raise GateError("facade feature closure does not identify any iWork feature")

    unsafe_dependents: set[str] = set()
    for package in packages.values():
        for dependency in package.dependencies:
            if dependency.name != FACADE_PACKAGE:
                continue
            unknown = dependency.features - set(facade.features)
            if unknown:
                raise GateError(
                    f"{package.name} requests unknown litchi features: "
                    + ", ".join(sorted(unknown))
                )
            if _dependency_reaches_unsafe_facade(dependency, unsafe_features):
                unsafe_dependents.add(package.name)

    _require_exact_package_set(
        unsafe_dependents,
        frozenset({PY_FACADE_PACKAGE}),
        "unsafe facade dependent inventory",
    )
    excluded_packages = frozenset(iwork_packages | unsafe_dependents)
    _require_exact_package_set(
        excluded_packages,
        EXPECTED_EXCLUDED_PACKAGE_NAMES,
        "excluded package inventory",
    )
    if len(excluded_packages) != EXPECTED_EXCLUDED_PACKAGES:
        raise GateError(
            "expected "
            f"{EXPECTED_EXCLUDED_PACKAGES} excluded packages, found "
            f"{len(excluded_packages)}: "
            + ", ".join(sorted(excluded_packages))
        )
    if FACADE_PACKAGE in excluded_packages:
        raise GateError("the litchi facade must be gated separately")

    bulk_packages = frozenset(set(packages) - set(excluded_packages) - {FACADE_PACKAGE})
    _require_exact_package_set(
        bulk_packages,
        EXPECTED_BULK_PACKAGE_NAMES,
        "bulk package inventory",
    )
    if len(bulk_packages) != EXPECTED_BULK_PACKAGES:
        raise GateError(
            "expected "
            f"{EXPECTED_BULK_PACKAGES} bulk packages, found {len(bulk_packages)}"
        )
    _require_bulk_dependency_boundary(packages, bulk_packages, excluded_packages)
    _require_selected_root_targets(
        packages, bulk_packages | frozenset({FACADE_PACKAGE})
    )
    return WorkspacePlan(
        packages=packages,
        facade_features=facade.features,
        iwork_packages=iwork_packages,
        unsafe_facade_dependents=frozenset(unsafe_dependents),
        excluded_packages=excluded_packages,
        bulk_packages=bulk_packages,
        safe_facade_features=safe_features,
        unsafe_facade_features=unsafe_features,
    )


def _wait_for_process(process: Any, timeout: float | None) -> bool:
    """Wait for a process when possible, reporting timeout/unavailable state."""

    try:
        if timeout is None:
            process.wait()
        else:
            process.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        return False
    except TypeError:
        # Small injected test doubles may not accept Popen's timeout keyword.
        process.wait()
    except OSError:
        return False
    return True


def _terminate_and_reap_process(
    process: Any, *, process_group: bool = False
) -> None:
    """Best-effort terminate/reap for an interrupted or capped child."""

    pid = getattr(process, "pid", None)
    group_signal_sent = False
    if process_group and os.name == "posix" and isinstance(pid, int) and pid > 0:
        try:
            os.killpg(pid, signal.SIGTERM)
            group_signal_sent = True
        except OSError:
            pass
    if not group_signal_sent:
        try:
            process.terminate()
        except (AttributeError, OSError):
            pass
    if _wait_for_process(process, PROCESS_TERM_GRACE_SECONDS):
        return

    killed = False
    if process_group and os.name == "posix" and isinstance(pid, int) and pid > 0:
        try:
            os.killpg(pid, signal.SIGKILL)
            killed = True
        except OSError:
            pass
    if not killed:
        try:
            process.kill()
        except (AttributeError, OSError):
            pass
    _wait_for_process(process, PROCESS_TERM_GRACE_SECONDS)


def _drain_capture_stream(
    stream: Any,
    stream_name: str,
    limit_bytes: int,
    output: bytearray,
    state: dict[str, str | None],
    state_lock: threading.Lock,
) -> None:
    """Drain one pipe while retaining at most the configured byte cap."""

    bytes_seen = 0
    try:
        while True:
            chunk = stream.read(CAPTURE_CHUNK_BYTES)
            if not chunk:
                break
            if isinstance(chunk, str):
                chunk = chunk.encode("utf-8", errors="replace")
            if not isinstance(chunk, (bytes, bytearray, memoryview)):
                raise TypeError("capture stream returned a non-byte chunk")
            chunk_bytes = bytes(chunk)
            if bytes_seen < limit_bytes:
                output.extend(chunk_bytes[: limit_bytes - bytes_seen])
            bytes_seen += len(chunk_bytes)
            if bytes_seen > limit_bytes:
                with state_lock:
                    state["exceeded_stream"] = state["exceeded_stream"] or stream_name
    except BaseException as error:
        with state_lock:
            state["read_error"] = state["read_error"] or (
                f"{type(error).__name__}: {error}"
            )
    finally:
        try:
            stream.close()
        except OSError:
            pass


def _run_capped_capture(
    argv: Sequence[str],
    *,
    cwd: Path = ROOT,
    env: Mapping[str, str] | None = None,
    limit_bytes: int = CAPTURE_OUTPUT_LIMIT_BYTES,
    popen_factory: Callable[..., Any] | None = None,
    sleep_fn: Callable[[float], None] | None = None,
) -> subprocess.CompletedProcess[str]:
    """Capture bounded decoded output while continuously draining both pipes."""

    if limit_bytes < 1:
        raise GateError("capture output limit must be positive")
    factory = popen_factory or subprocess.Popen
    popen_kwargs: dict[str, Any] = {
        "cwd": cwd,
        "stdout": subprocess.PIPE,
        "stderr": subprocess.PIPE,
    }
    if env is not None:
        popen_kwargs["env"] = dict(env)
    isolated_posix_session = popen_factory is None and os.name == "posix"
    if isolated_posix_session:
        popen_kwargs["start_new_session"] = True
    process = factory(list(argv), **popen_kwargs)
    outputs = {"stdout": bytearray(), "stderr": bytearray()}
    state: dict[str, str | None] = {
        "exceeded_stream": None,
        "read_error": None,
    }
    state_lock = threading.Lock()
    threads = [
        threading.Thread(
            target=_drain_capture_stream,
            args=(
                process.stdout,
                "stdout",
                limit_bytes,
                outputs["stdout"],
                state,
                state_lock,
            ),
            daemon=True,
        ),
        threading.Thread(
            target=_drain_capture_stream,
            args=(
                process.stderr,
                "stderr",
                limit_bytes,
                outputs["stderr"],
                state,
                state_lock,
            ),
            daemon=True,
        ),
    ]
    for thread in threads:
        thread.start()
    sleeper = sleep_fn or time.sleep
    try:
        while True:
            returncode = process.poll()
            with state_lock:
                exceeded_stream = state["exceeded_stream"]
            if exceeded_stream is not None:
                _terminate_and_reap_process(
                    process,
                    process_group=isolated_posix_session,
                )
                break
            if returncode is not None:
                break
            sleeper(CAPTURE_POLL_INTERVAL_SECONDS)
        returncode = process.wait()
    except KeyboardInterrupt:
        _terminate_and_reap_process(
            process,
            process_group=isolated_posix_session,
        )
        raise
    finally:
        for thread in threads:
            thread.join()
    with state_lock:
        exceeded_stream = state["exceeded_stream"]
        read_error = state["read_error"]
    if read_error is not None:
        raise GateError(
            f"{shlex.join(argv)} output capture failed: {read_error}"
        )
    if exceeded_stream is not None:
        raise CaptureLimitError(argv, exceeded_stream, limit_bytes)
    stdout = bytes(outputs["stdout"]).decode("utf-8", errors="replace")
    stderr = bytes(outputs["stderr"]).decode("utf-8", errors="replace")
    return subprocess.CompletedProcess(list(argv), returncode, stdout, stderr)


def _cargo_metadata(cargo: str) -> Mapping[str, Any]:
    """Read workspace metadata through Cargo without invoking a shell."""

    target_platform = _target_platform()
    result = _run_capped_capture(
        [
            cargo,
            "metadata",
            "--format-version",
            "1",
            "--filter-platform",
            target_platform,
            "--no-deps",
        ],
        cwd=ROOT,
    )
    if result.returncode:
        details = result.stderr.strip() or result.stdout.strip()
        raise GateError(f"cargo metadata failed ({result.returncode}): {details}")
    try:
        value = json.loads(result.stdout)
    except json.JSONDecodeError as error:
        raise GateError(f"cargo metadata returned invalid JSON: {error}") from error
    if not isinstance(value, Mapping):
        raise GateError("cargo metadata returned a non-object JSON value")
    return value


def _target_platform() -> str:
    """Return the target Cargo will use for the gate's host-only operations."""

    configured = os.environ.get("CARGO_BUILD_TARGET")
    if configured:
        if not TARGET_PLATFORM_RE.fullmatch(configured):
            raise GateError(
                "CARGO_BUILD_TARGET must be a target triple for metadata filtering, "
                f"got {configured!r}"
            )
        return configured
    rustc = os.environ.get("RUSTC", "rustc")
    result = _run_capped_capture([rustc, "-vV"], cwd=ROOT)
    if result.returncode:
        details = result.stderr.strip() or result.stdout.strip()
        raise GateError(f"rustc host target probe failed ({result.returncode}): {details}")
    for line in result.stdout.splitlines():
        key, separator, value = line.partition(":")
        if key == "host" and separator and TARGET_PLATFORM_RE.fullmatch(value.strip()):
            return value.strip()
    raise GateError("rustc -vV did not report a valid host target triple")


def _workspace_exclusions(plan: WorkspacePlan) -> list[str]:
    """Return Cargo's workspace selection and all 19 non-bulk exclusions."""

    exclusions = [FACADE_PACKAGE, *sorted(plan.excluded_packages)]
    return [item for name in exclusions for item in ("--exclude", name)]


def _facade_feature_args(plan: WorkspacePlan) -> list[str]:
    """Return one deterministic combined safe feature argument."""

    # The default feature is empty in this workspace and is represented by
    # --no-default-features.  All other safe features are compiled together so
    # each operation mode has one bounded facade invocation.
    features = sorted(plan.safe_facade_features - {"default"})
    return ["--features", ",".join(features)] if features else []


def _facade_tree_selection(
    plan: WorkspacePlan, feature: str | None = None
) -> list[str]:
    """Build one individual or combined safe facade tree selection."""

    selection = ["--package", FACADE_PACKAGE]
    if feature == "default":
        # Deliberately leave Cargo's default feature set enabled.  The default
        # closure is verified independently so a future non-empty default
        # cannot be hidden by the --no-default-features used by the other
        # isolated feature trees.
        return selection
    selection.append("--no-default-features")
    if feature is None:
        selection.extend(_facade_feature_args(plan))
    else:
        selection.extend(["--features", feature])
    return selection


def _bulk_root_specs(
    cargo: str,
    plan: WorkspacePlan,
    subcommand: str,
    common: Sequence[str],
    updates: tuple[tuple[str, str], ...],
) -> list[CommandSpec]:
    """Generate one all-feature command for every sorted bulk package root."""

    return [
        CommandSpec(
            f"bulk/{package}",
            (
                cargo,
                subcommand,
                "--package",
                package,
                "--all-features",
                *common,
            ),
            updates,
        )
        for package in sorted(plan.bulk_packages)
    ]


def command_specs(cargo: str, plan: WorkspacePlan, mode: str) -> tuple[CommandSpec, ...]:
    """Generate serialized bulk-root and facade argv lists for one mode."""

    if mode not in MODES or mode in {"print", "verify"}:
        raise GateError(f"{mode!r} does not generate Cargo run commands")
    if mode == "lib-tests":
        # A single workspace test invocation keeps every test binary alive
        # until the final link and can consume several GiB. Serialize each
        # exact bulk root, then remove only that package's test artifacts
        # before proceeding. Dependencies remain reusable across roots.
        specs: list[CommandSpec] = []
        for package in sorted(plan.bulk_packages):
            specs.append(
                CommandSpec(
                    f"bulk-test/{package}",
                    (
                        cargo,
                        "test",
                        "--package",
                        package,
                        "--all-features",
                        "--lib",
                        "--tests",
                        "--",
                        "--test-threads=1",
                    ),
                )
            )
            specs.append(
                CommandSpec(
                    f"bulk-clean/{package}",
                    (cargo, "clean", "--package", package),
                )
            )
        facade_default = (
            cargo,
            "test",
            "--package",
            FACADE_PACKAGE,
            "--lib",
            "--tests",
            "--",
            "--test-threads=1",
        )
        if "default" in plan.safe_facade_features:
            specs.append(CommandSpec("facade-default-feature", facade_default))
        facade = (
            cargo,
            "test",
            "--package",
            FACADE_PACKAGE,
            "--no-default-features",
            *_facade_feature_args(plan),
            "--lib",
            "--tests",
            "--",
            "--test-threads=1",
        )
        specs.append(CommandSpec("facade-safe-features", facade))
        specs.append(
            CommandSpec(
                "facade-clean",
                (cargo, "clean", "--package", FACADE_PACKAGE),
            )
        )
        return tuple(specs)
    # All features are intentional for bulk roots: each of the 45 packages
    # must prove that its non-default feature surface remains non-iWork.  Each
    # root is isolated so Cargo does not retain every high-memory target in a
    # single workspace invocation.  The facade is selected separately with
    # the exact safe closure below.
    facade_selection = [
        "--package",
        FACADE_PACKAGE,
        "--no-default-features",
        *_facade_feature_args(plan),
    ]
    if mode == "check":
        subcommand, common = "check", ["--lib", "--tests"]
    elif mode == "clippy":
        subcommand, common = "clippy", [
            "--lib",
            "--no-deps",
            "--",
            "-D",
            "warnings",
        ]
    elif mode == "doc":
        subcommand, common = "doc", ["--no-deps"]
    elif mode == "doc-tests":
        subcommand, common = "test", ["--doc", "--", "--test-threads=1"]
    else:
        subcommand, common = "check", ["--all-targets"]
    updates: tuple[tuple[str, str], ...] = ()
    if mode in {"doc", "doc-tests"}:
        updates = (("RUSTDOCFLAGS", "-D warnings"),)
    elif mode == "deprecated":
        updates = (("RUSTFLAGS", "-D deprecated"),)
    specs = _bulk_root_specs(cargo, plan, subcommand, common, updates)
    facade_default = (cargo, subcommand, "--package", FACADE_PACKAGE, *common)
    facade = (cargo, subcommand, *facade_selection, *common)
    if "default" in plan.safe_facade_features:
        specs.append(CommandSpec("facade-default-feature", facade_default, updates))
    specs.append(CommandSpec("facade-safe-features", facade, updates))
    return tuple(specs)


def parse_tree_package_names(output: str) -> frozenset[str]:
    """Parse package names from ``cargo tree --prefix none --format {p}`."""

    names: set[str] = set()
    for line in output.splitlines():
        match = TREE_PACKAGE_RE.match(line.strip())
        if match:
            names.add(match.group(1))
    return frozenset(names)


def _tree_local_path_text(suffix: str) -> str | None:
    """Extract an absolute local path from Cargo's package display suffix."""

    text = suffix.strip()
    if text.endswith(" (*)"):
        text = text[:-4].rstrip()
    if text.startswith("(proc-macro) "):
        text = text.removeprefix("(proc-macro) ").lstrip()
    if text.endswith(" (proc-macro)"):
        text = text.removesuffix(" (proc-macro)").rstrip()
    # The path is the only suffix token that can begin with a filesystem
    # absolute marker. Scan from that marker instead of assuming a fixed
    # number or order of parenthesized annotations.
    marker = re.compile(
        r"(?:^|(?<=[( ]))(?:/[^\s()]|[A-Za-z]:[\\/]|(?:\\\\|//))"
    )
    for match in marker.finditer(text):
        candidate = text[match.start() :].strip()
        if candidate.endswith(")"):
            candidate = candidate[:-1].rstrip()
        try:
            _portable_canonical_path(candidate, "cargo tree package path")
        except GateError:
            continue
        return candidate
    return None


def _tree_local_path(candidate: str, package_name: str) -> Path | PureWindowsPath:
    """Resolve a tree path on-host or preserve Windows paths on Unix tests."""

    portable_path = _portable_canonical_path(
        candidate, f"cargo tree package {package_name!r} path"
    )
    if _is_windows_absolute_path(candidate):
        if os.name == "nt":
            return _resolved_absolute_path(
                portable_path, f"cargo tree package {package_name!r} path"
            )
        return PureWindowsPath(portable_path)
    return _resolved_absolute_path(
        portable_path, f"cargo tree package {package_name!r} path"
    )


def _parse_tree_entries(
    output: str,
) -> TreeEntries:
    """Parse package names and local paths from Cargo tree output."""

    entries: dict[str, set[Path | PureWindowsPath | None]] = {}
    versions: dict[str, set[str]] = {}
    for line in output.splitlines():
        match = TREE_PACKAGE_RE.match(line.strip())
        if not match:
            continue
        name = match.group(1)
        version = match.group(2)
        suffix = (match.group(3) or "").strip()
        local_path: Path | PureWindowsPath | None = None
        candidate = _tree_local_path_text(suffix)
        if candidate is not None:
            local_path = _tree_local_path(candidate, name)
        if local_path is None and suffix in {"(*)", "(proc-macro) (*)"}:
            continue
        entries.setdefault(name, set()).add(local_path)
        versions.setdefault(name, set()).add(version)
    return TreeEntries(
        {name: frozenset(paths) for name, paths in entries.items()},
        {name: frozenset(values) for name, values in versions.items()},
    )


def _protobuf_packages(names: Iterable[str]) -> frozenset[str]:
    return frozenset(name for name in names if PROTOBUF_PACKAGE_RE.match(name))


def _tree_command(cargo: str, selection: Sequence[str]) -> list[str]:
    return [
        cargo,
        "tree",
        *selection,
        "--target",
        "all",
        # Include normal, build, and dev edges so an accidental protobuf build
        # dependency cannot hide behind the default normal-only tree.
        "--edges",
        "normal,build,dev",
        "--prefix",
        "none",
        "--format",
        "{p}",
    ]


def _capture_tree(
    cargo: str, selection: Sequence[str]
) -> dict[str, frozenset[Path | PureWindowsPath | None]]:
    argv = _tree_command(cargo, selection)
    result = _run_capped_capture(argv, cwd=ROOT)
    if result.returncode:
        details = result.stderr.strip() or result.stdout.strip()
        raise GateError(f"{shlex.join(argv)} failed ({result.returncode}): {details}")
    return _parse_tree_entries(result.stdout)


def _require_tree_roots(
    entries: Mapping[str, Any], required: Iterable[str], description: str
) -> None:
    """Reject successful but empty or partial Cargo tree output."""

    present = frozenset(entries)
    missing = frozenset(required) - present
    if missing:
        raise GateError(
            f"{description} is missing required package roots: "
            + ", ".join(sorted(missing))
        )


def _require_tree_identities(
    entries: Mapping[str, frozenset[Path | PureWindowsPath | None]],
    plan: WorkspacePlan,
    required: Iterable[str],
    description: str,
) -> None:
    """Require selected roots to be the canonical local workspace packages."""

    for name in sorted(required):
        expected = plan.packages[name].manifest_path.parent
        observed = entries.get(name, frozenset())
        observed_versions = getattr(entries, "versions", {}).get(name, frozenset())
        expected_version = plan.packages[name].version
        if observed_versions and observed_versions != {expected_version}:
            raise GateError(
                f"{description} has non-canonical versions for {name!r}: "
                + ", ".join(sorted(observed_versions))
            )
        if expected not in observed:
            raise GateError(
                f"{description} has no canonical local identity for {name!r} "
                f"at {expected}"
            )
        noncanonical = observed - {expected}
        if noncanonical:
            raise GateError(
                f"{description} has non-canonical identities for {name!r}: "
                + ", ".join(sorted(str(path) for path in noncanonical))
            )


def verify_dependency_trees(cargo: str, plan: WorkspacePlan) -> tuple[int, int]:
    """Verify the bulk all-features tree and the safe facade closure."""

    bulk_selection = ["--workspace", "--all-features", *_workspace_exclusions(plan)]
    bulk_entries = _capture_tree(cargo, bulk_selection)
    _require_tree_roots(bulk_entries, plan.bulk_packages, "bulk dependency tree")
    _require_tree_identities(
        bulk_entries, plan, plan.bulk_packages, "bulk dependency tree"
    )
    forbidden_bulk = (
        set(plan.excluded_packages)
        | {FACADE_PACKAGE}
        | set(_protobuf_packages(bulk_entries))
    )
    if forbidden_bulk & bulk_entries.keys():
        raise GateError(
            "bulk dependency tree pulls forbidden packages: "
            + ", ".join(sorted(forbidden_bulk & bulk_entries.keys()))
        )

    facade_trees = 0
    if "default" in plan.safe_facade_features:
        default_entries = _capture_tree(cargo, _facade_tree_selection(plan, "default"))
        _require_tree_roots(
            default_entries, {FACADE_PACKAGE}, "safe facade default feature tree"
        )
        _require_tree_identities(
            default_entries,
            plan,
            {FACADE_PACKAGE},
            "safe facade default feature tree",
        )
        default_forbidden = (
            set(plan.excluded_packages) | set(_protobuf_packages(default_entries))
        ) & default_entries.keys()
        if default_forbidden:
            raise GateError(
                "safe facade default feature tree pulls forbidden packages: "
                + ", ".join(sorted(default_forbidden))
            )
        facade_trees += 1
    for feature in sorted(plan.safe_facade_features - {"default"}):
        selection = _facade_tree_selection(plan, feature)
        entries = _capture_tree(cargo, selection)
        _require_tree_roots(
            entries, {FACADE_PACKAGE}, f"safe facade feature {feature!r} tree"
        )
        _require_tree_identities(
            entries, plan, {FACADE_PACKAGE}, f"safe facade feature {feature!r} tree"
        )
        forbidden = (
            set(plan.excluded_packages) | set(_protobuf_packages(entries))
        ) & entries.keys()
        if forbidden:
            raise GateError(
                f"safe facade feature {feature!r} pulls forbidden packages: "
                + ", ".join(sorted(forbidden))
            )
        facade_trees += 1
    combined_entries = _capture_tree(cargo, _facade_tree_selection(plan))
    _require_tree_roots(
        combined_entries, {FACADE_PACKAGE}, "combined safe facade feature tree"
    )
    _require_tree_identities(
        combined_entries,
        plan,
        {FACADE_PACKAGE},
        "combined safe facade feature tree",
    )
    combined_forbidden = (
        set(plan.excluded_packages) | set(_protobuf_packages(combined_entries))
    ) & combined_entries.keys()
    if combined_forbidden:
        raise GateError(
            "combined safe facade feature tree pulls forbidden packages: "
            + ", ".join(sorted(combined_forbidden))
        )
    return len(plan.bulk_packages), facade_trees


def _environment(target_dir: str | None) -> dict[str, str]:
    """Build a bounded Cargo environment for local and CI gate runs."""

    environment = os.environ.copy()
    if target_dir:
        candidate = Path(target_dir)
        if not candidate.is_absolute():
            candidate = ROOT / candidate
    else:
        candidate = ROOT / "target/non-iwork-gate"
    # These values are gate invariants.  Ambient values must not silently
    # widen the process fan-out or select a second artifact directory.
    environment["CARGO_TARGET_DIR"] = str(candidate)
    environment["CARGO_BUILD_JOBS"] = "1"
    environment["CARGO_INCREMENTAL"] = "0"
    environment["CARGO_PROFILE_DEV_DEBUG"] = "0"
    environment["CARGO_PROFILE_TEST_DEBUG"] = "0"
    return environment


def _updated_environment(
    base: Mapping[str, str], updates: Sequence[tuple[str, str]]
) -> dict[str, str]:
    environment = dict(base)
    encoded_equivalents = {
        "RUSTFLAGS": "CARGO_ENCODED_RUSTFLAGS",
        "RUSTDOCFLAGS": "CARGO_ENCODED_RUSTDOCFLAGS",
    }
    for key, value in updates:
        # These policies must be exact.  Cargo gives the encoded variables
        # precedence when both forms are present, so remove only the paired
        # flag variables rather than broadly scrubbing compiler wrappers.
        environment[key] = value
        encoded_key = encoded_equivalents.get(key)
        if encoded_key is not None:
            environment.pop(encoded_key, None)
    return environment


def _bounded_env_keys(environment: Mapping[str, str]) -> tuple[str, ...]:
    """Return only the non-secret, bounded environment key allow-list."""

    return tuple(sorted(key for key in RECORDED_ENV_KEYS if key in environment))


def _bounded_env_values(environment: Mapping[str, str]) -> tuple[tuple[str, str], ...]:
    """Return bounded values for the enforced Cargo environment allow-list."""

    values: list[tuple[str, str]] = []
    for key in sorted(RECORDED_ENV_KEYS):
        if key not in environment:
            continue
        value = environment[key]
        if len(value) > MAX_RECORDED_ENV_VALUE_LENGTH:
            value = value[:MAX_RECORDED_ENV_VALUE_LENGTH] + "..."
        values.append((key, value))
    return tuple(values)


def _unavailable_target_footprint(
    reason: str = "target directory is unavailable"
) -> TargetFootprint:
    return TargetFootprint("unavailable", None, None, reason)


def _target_footprint(
    target_dir: Path | None,
    *,
    max_entries: int = TARGET_SCAN_MAX_ENTRIES,
    max_directories: int = TARGET_SCAN_MAX_DIRECTORIES,
    max_depth: int = TARGET_SCAN_MAX_DEPTH,
) -> TargetFootprint:
    """Count regular files without following symlinks or exceeding scan caps."""

    if max_entries < 1 or max_directories < 1 or max_depth < 0:
        return TargetFootprint(
            "incomplete",
            None,
            None,
            "target scan limits must allow at least one entry and directory",
        )
    if target_dir is None:
        return _unavailable_target_footprint("target directory was not configured")
    try:
        root_stat = os.lstat(target_dir)
    except OSError:
        return _unavailable_target_footprint(
            "target directory is missing or unreadable"
        )
    if not stat.S_ISDIR(root_stat.st_mode):
        return _unavailable_target_footprint("target path is not a directory")

    total_bytes = 0
    total_count = 0
    entry_count = 0
    directory_count = 1
    complete = True
    incomplete_reason = "one or more target entries changed or could not be read"
    directories: list[tuple[Path, int]] = [(target_dir, 0)]
    while directories:
        directory, depth = directories.pop()
        try:
            # Recheck before opening each directory so a replacement symlink
            # is not deliberately traversed while a report is being made.
            directory_stat = os.lstat(directory)
            if not stat.S_ISDIR(directory_stat.st_mode):
                complete = False
                incomplete_reason = "a target directory changed during the scan"
                continue
            with os.scandir(directory) as entries:
                for entry in entries:
                    entry_count += 1
                    if entry_count > max_entries:
                        complete = False
                        incomplete_reason = "target scan entry cap reached"
                        directories.clear()
                        break
                    try:
                        entry_stat = os.lstat(entry.path)
                    except OSError:
                        complete = False
                        incomplete_reason = (
                            "one or more target entries changed or could not be read"
                        )
                        continue
                    if stat.S_ISREG(entry_stat.st_mode):
                        total_count += 1
                        total_bytes += entry_stat.st_size
                    elif stat.S_ISDIR(entry_stat.st_mode):
                        if depth >= max_depth:
                            complete = False
                            incomplete_reason = "target scan depth cap reached"
                            directories.clear()
                            break
                        if directory_count >= max_directories:
                            complete = False
                            incomplete_reason = "target scan directory cap reached"
                            directories.clear()
                            break
                        directory_count += 1
                        directories.append((Path(entry.path), depth + 1))
        except OSError:
            complete = False
            incomplete_reason = "one or more target directories could not be read"

    if not complete:
        return TargetFootprint("incomplete", None, None, incomplete_reason)
    return TargetFootprint(
        "complete",
        total_bytes,
        total_count,
        "regular directory entries counted with lstat; symlinks were not followed",
    )


def _linux_process_rss(pid: int) -> int | None:
    """Sample cumulative RSS for a Linux process and its descendants."""

    if not sys.platform.startswith("linux"):
        return None
    try:
        process_names = os.listdir("/proc")
    except OSError:
        return None

    processes: dict[int, tuple[int, int]] = {}
    for name in process_names:
        if not name.isdigit():
            continue
        process_id = int(name)
        parent_id: int | None = None
        rss_bytes: int | None = None
        try:
            with open(
                f"/proc/{name}/status",
                encoding="ascii",
                errors="replace",
            ) as status_file:
                for line in status_file:
                    if line.startswith("PPid:"):
                        fields = line.split()
                        if len(fields) == 2:
                            parent_id = int(fields[1])
                    elif line.startswith("VmRSS:"):
                        fields = line.split()
                        if len(fields) >= 2:
                            # Linux reports VmRSS in KiB.
                            rss_bytes = int(fields[1]) * 1024
        except (OSError, ValueError):
            continue
        if parent_id is not None and rss_bytes is not None:
            processes[process_id] = (parent_id, rss_bytes)

    if pid not in processes:
        return None
    descendants = {pid}
    changed = True
    while changed:
        changed = False
        for process_id, (parent_id, _) in processes.items():
            if process_id not in descendants and parent_id in descendants:
                descendants.add(process_id)
                changed = True
    return sum(processes[process_id][1] for process_id in descendants)


def _rss_configuration(
    enabled: bool,
    sampler: Callable[[int], int | None] | None,
    platform: str | None,
    scope: str | None,
) -> tuple[Callable[[int], int | None] | None, str, str]:
    """Select a truthful RSS source, including explicit unavailable states."""

    current_platform = platform or sys.platform or "unknown"
    if not enabled:
        return None, current_platform, scope or "not-collected"
    if sampler is not None:
        return sampler, platform or "injected", scope or "injected-process-tree"
    if sys.platform.startswith("linux"):
        return _linux_process_rss, "linux", "sampled-descendant-process-tree"
    return None, current_platform, scope or "unavailable"


class _ChildRssTracker:
    """Keep the maximum cumulative sampled child RSS without affecting a run."""

    def __init__(
        self,
        sampler: Callable[[int], int | None] | None,
        platform: str,
        scope: str,
    ) -> None:
        self.sampler = sampler
        self.platform = platform
        self.scope = scope
        self.high_water_bytes: int | None = None
        self.sample_count = 0
        self.sample_errors = 0

    def sample(self, pid: int) -> None:
        if self.sampler is None:
            return
        try:
            value = self.sampler(pid)
        except (OSError, ValueError):
            self.sample_errors += 1
            return
        if not isinstance(value, int) or value < 0:
            self.sample_errors += 1
            return
        self.sample_count += 1
        if self.high_water_bytes is None or value > self.high_water_bytes:
            self.high_water_bytes = value

    def result(self) -> ChildRss:
        if self.sampler is None:
            status = "not_collected" if self.scope == "not-collected" else "unavailable"
            reason = (
                "RSS collection was disabled"
                if status == "not_collected"
                else "RSS sampling is unavailable on this host"
            )
        elif self.high_water_bytes is None:
            status = "unavailable"
            reason = "no readable process-tree samples were observed"
        else:
            status = "partial" if self.sample_errors else "available"
            reason = (
                "one or more process-tree samples were unreadable"
                if self.sample_errors
                else "sampled process-tree RSS sums"
            )
        return ChildRss(
            self.high_water_bytes,
            self.platform,
            self.scope,
            status,
            reason,
        )


def _execute_command(
    argv: Sequence[str],
    environment: Mapping[str, str],
    *,
    clock: Callable[[], int] | None = None,
    popen_factory: Callable[..., Any] | None = None,
    rss_enabled: bool = False,
    rss_sampler: Callable[[int], int | None] | None = None,
    rss_platform: str | None = None,
    rss_scope: str | None = None,
    sleep_fn: Callable[[float], None] | None = None,
) -> CommandExecution:
    """Run one argv command while optionally sampling child-process RSS."""

    clock_fn = clock or time.monotonic_ns
    factory = popen_factory or subprocess.Popen
    sleeper = sleep_fn or time.sleep
    configured_sampler, configured_platform, configured_scope = _rss_configuration(
        rss_enabled, rss_sampler, rss_platform, rss_scope
    )
    tracker = _ChildRssTracker(
        configured_sampler,
        configured_platform,
        configured_scope,
    )
    started_ns = clock_fn()
    popen_kwargs: dict[str, Any] = {
        "cwd": ROOT,
        "env": dict(environment),
    }
    isolated_posix_session = popen_factory is None and os.name == "posix"
    if isolated_posix_session:
        popen_kwargs["start_new_session"] = True
    process = factory(list(argv), **popen_kwargs)
    try:
        if configured_sampler is None:
            returncode = process.wait()
        else:
            while True:
                tracker.sample(process.pid)
                returncode = process.poll()
                if returncode is not None:
                    break
                sleeper(RSS_POLL_INTERVAL_SECONDS)
    except KeyboardInterrupt:
        _terminate_and_reap_process(
            process,
            process_group=isolated_posix_session,
        )
        raise
    elapsed_ns = max(0, clock_fn() - started_ns)
    return CommandExecution(returncode, elapsed_ns, tracker.result())


class ExecutionRecorder:
    """Collect correctness-only execution telemetry for one gate mode."""

    def __init__(
        self,
        mode: str,
        environment: Mapping[str, str],
        *,
        clock: Callable[[], int] | None = None,
        target_scanner: Callable[[Path | None], TargetFootprint] | None = None,
        popen_factory: Callable[..., Any] | None = None,
        rss_sampler: Callable[[int], int | None] | None = None,
        rss_platform: str | None = None,
        rss_scope: str | None = None,
        sleep_fn: Callable[[float], None] | None = None,
    ) -> None:
        self.mode = mode
        self.environment = dict(environment)
        target_text = self.environment.get("CARGO_TARGET_DIR")
        self.target_dir = Path(target_text) if target_text else None
        self.clock = clock or time.monotonic_ns
        self.target_scanner = target_scanner or _target_footprint
        self.popen_factory = popen_factory
        self.rss_sampler, self.rss_platform, self.rss_scope = _rss_configuration(
            True, rss_sampler, rss_platform, rss_scope
        )
        self.sleep_fn = sleep_fn
        self.started_ns = self.clock()
        self.target_before = self._scan_target()
        self.target_after = self.target_before
        self.phases: list[PhaseRecord] = []
        self.child_rss_high_water_bytes: int | None = None
        self.outcome = "running"
        self.error: str | None = None
        self.elapsed_ns = 0

    def _scan_target(self) -> TargetFootprint:
        try:
            value = self.target_scanner(self.target_dir)
        except (OSError, ValueError, TypeError):
            return _unavailable_target_footprint()
        if not isinstance(value, TargetFootprint):
            return _unavailable_target_footprint()
        return value

    def run_phase(
        self,
        index: int,
        spec: CommandSpec,
        command_environment: Mapping[str, str],
    ) -> int:
        """Execute and record one phase, preserving command failures."""

        target_before = self._scan_target()
        started_ns = self.clock()
        try:
            execution = _execute_command(
                spec.argv,
                command_environment,
                clock=self.clock,
                popen_factory=self.popen_factory,
                rss_enabled=True,
                rss_sampler=self.rss_sampler,
                rss_platform=self.rss_platform,
                rss_scope=self.rss_scope,
                sleep_fn=self.sleep_fn,
            )
        except OSError:
            target_after = self._scan_target()
            elapsed_ns = max(0, self.clock() - started_ns)
            self.phases.append(
                PhaseRecord(
                    self.mode,
                    index,
                    spec.scope,
                    spec.argv,
                    "error",
                    None,
                    elapsed_ns,
                    target_before,
                    target_after,
                    ChildRss(
                        None,
                        self.rss_platform,
                        self.rss_scope,
                        "unavailable",
                        "child process could not be started",
                    ),
                    _bounded_env_keys(command_environment),
                    _bounded_env_values(command_environment),
                )
            )
            self.target_after = target_after
            raise
        except KeyboardInterrupt:
            target_after = self._scan_target()
            elapsed_ns = max(0, self.clock() - started_ns)
            self.phases.append(
                PhaseRecord(
                    self.mode,
                    index,
                    spec.scope,
                    spec.argv,
                    "interrupted",
                    None,
                    elapsed_ns,
                    target_before,
                    target_after,
                    ChildRss(
                        None,
                        self.rss_platform,
                        self.rss_scope,
                        "unavailable",
                        "KeyboardInterrupt interrupted RSS sampling",
                    ),
                    _bounded_env_keys(command_environment),
                    _bounded_env_values(command_environment),
                )
            )
            self.target_after = target_after
            raise
        target_after = self._scan_target()
        status = "passed" if execution.returncode == 0 else "failed"
        self.phases.append(
            PhaseRecord(
                self.mode,
                index,
                spec.scope,
                spec.argv,
                status,
                execution.returncode,
                execution.elapsed_ns,
                target_before,
                target_after,
                execution.child_rss,
                _bounded_env_keys(command_environment),
                _bounded_env_values(command_environment),
            )
        )
        if execution.child_rss.high_water_bytes is not None and (
            self.child_rss_high_water_bytes is None
            or execution.child_rss.high_water_bytes > self.child_rss_high_water_bytes
        ):
            self.child_rss_high_water_bytes = execution.child_rss.high_water_bytes
        self.target_after = target_after
        return execution.returncode

    def _aggregate_rss_status(self) -> tuple[str, str]:
        statuses = [phase.child_rss.status for phase in self.phases]
        if not statuses:
            return "unavailable", "no execution phases were sampled"
        if all(status == "available" for status in statuses):
            return "available", "all phases had readable process-tree RSS samples"
        if all(status in {"unavailable", "not_collected"} for status in statuses):
            return "unavailable", "RSS sampling was unavailable for all phases"
        return "partial", "one or more phase RSS samples were partial or unavailable"

    def finish(self, outcome: str, error: str | None) -> None:
        """Close the report state even when a phase stopped the mode early."""

        self.outcome = outcome
        self.error = error[:MAX_REPORT_ERROR_LENGTH] if error else None
        self.target_after = self._scan_target()
        self.elapsed_ns = max(0, self.clock() - self.started_ns)

    def as_dict(self) -> dict[str, Any]:
        rss_status, rss_reason = self._aggregate_rss_status()
        return {
            "version": REPORT_VERSION,
            "mode": self.mode,
            "claim_scope": "no performance claim",
            "outcome": self.outcome,
            "error": self.error,
            "elapsed_ns": self.elapsed_ns,
            "clock": {
                "name": "time.monotonic_ns",
                "unit": "ns",
                "monotonic": True,
            },
            "host": {
                "platform": sys.platform,
                "os_name": os.name,
                "python_implementation": platform.python_implementation(),
                "python_version": platform.python_version(),
                "machine": platform.machine(),
            },
            "target_dir": str(self.target_dir) if self.target_dir else None,
            "target_scan_limits": {
                "max_entries": TARGET_SCAN_MAX_ENTRIES,
                "max_directories": TARGET_SCAN_MAX_DIRECTORIES,
                "max_depth": TARGET_SCAN_MAX_DEPTH,
            },
            "env_keys": list(_bounded_env_keys(self.environment)),
            "cargo_env": dict(_bounded_env_values(self.environment)),
            "target_before": self.target_before.as_dict(),
            "target_after": self.target_after.as_dict(),
            "child_rss": {
                "measurement": RSS_MEASUREMENT,
                "high_water_bytes": self.child_rss_high_water_bytes,
                "platform": self.rss_platform,
                "scope": self.rss_scope,
                "status": rss_status,
                "reason": rss_reason,
                "sample_interval_ms": RSS_SAMPLE_INTERVAL_MS,
            },
            "cleanup": {
                "disposition": "retained_by_policy",
                "recursive": False,
                "failure_artifacts": "failing phase/root retained",
                "prior_successful_lib_test_roots": "may be cleaned",
                "package_clean_scopes": [
                    phase.scope
                    for phase in self.phases
                    if phase.scope.startswith("bulk-clean/")
                    or phase.scope == "facade-clean"
                ],
                "package_clean_commands": [
                    {
                        "scope": phase.scope,
                        "argv": list(phase.argv),
                    }
                    for phase in self.phases
                    if phase.scope.startswith("bulk-clean/")
                    or phase.scope == "facade-clean"
                ],
            },
            "feature_unification": {
                "gate_scope": "per-package-root --all-features",
                "aggregate_workspace_unification": "not_claimed",
                "aggregate_guard": "existing workspace CI",
            },
            "limitations": {
                "rss": (
                    "20 ms sampled sums of readable descendant VmRSS values; "
                    "short-lived processes and unreadable descendants may be missed; "
                    "this is not an OS or child-process peak-RSS measurement"
                ),
                "target_scan": (
                    "logical regular-file entries are counted with lstat without "
                    "following symlinks; concurrent changes or access errors produce "
                    "incomplete/unavailable status; traversal is bounded by explicit "
                    f"entry={TARGET_SCAN_MAX_ENTRIES}, "
                    f"directory={TARGET_SCAN_MAX_DIRECTORIES}, "
                    f"depth={TARGET_SCAN_MAX_DEPTH} caps"
                ),
                "interruption": (
                    "KeyboardInterrupt returns 130 and requests best-effort child or "
                    "process-group termination; SIGTERM handling is not guaranteed"
                ),
            },
            "phases": [phase.as_dict() for phase in self.phases],
        }


def _record_path(value: str | Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else ROOT / path


def _write_atomic_report(path_value: str | Path, report: Mapping[str, Any]) -> None:
    """Write a complete JSON report through same-directory replacement."""

    path = _record_path(path_value)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=path.parent,
            prefix=f".{path.name}.",
            suffix=".tmp",
            delete=False,
        ) as report_file:
            temporary_path = Path(report_file.name)
            json.dump(report, report_file, sort_keys=True, indent=2)
            report_file.write("\n")
            report_file.flush()
            os.fsync(report_file.fileno())
        os.replace(temporary_path, path)
        try:
            directory_flags = os.O_RDONLY | getattr(os, "O_DIRECTORY", 0)
            directory_fd = os.open(path.parent, directory_flags)
        except (AttributeError, OSError):
            pass
        else:
            try:
                os.fsync(directory_fd)
            except OSError:
                pass
            finally:
                os.close(directory_fd)
    finally:
        if temporary_path is not None:
            try:
                temporary_path.unlink()
            except FileNotFoundError:
                pass


def run_mode(
    cargo: str,
    plan: WorkspacePlan,
    mode: str,
    environment: Mapping[str, str],
    *,
    recorder: ExecutionRecorder | None = None,
) -> None:
    """Run one generated mode, stopping at the first failed Cargo command."""

    for index, spec in enumerate(command_specs(cargo, plan, mode), start=1):
        command_environment = _updated_environment(environment, spec.env)
        print(f"[{mode}/{spec.scope}] $ {shlex.join(spec.argv)}", flush=True)
        if recorder is None:
            execution = _execute_command(spec.argv, command_environment)
            returncode = execution.returncode
        else:
            returncode = recorder.run_phase(index, spec, command_environment)
        if returncode:
            raise GateError(
                f"{mode}/{spec.scope} failed with exit status {returncode}"
            )


def print_plan(plan: WorkspacePlan) -> None:
    """Print stable machine-readable counts followed by human-readable sets."""

    print(f"workspace_packages={len(plan.packages)}")
    print(f"excluded_packages={len(plan.excluded_packages)}")
    print(f"bulk_packages={len(plan.bulk_packages)}")
    print(f"facade_safe_features={len(plan.safe_facade_features)}")
    print(f"facade_unsafe_features={len(plan.unsafe_facade_features)}")
    print("excluded=" + ",".join(sorted(plan.excluded_packages)))
    print("bulk=" + ",".join(sorted(plan.bulk_packages)))
    print("facade-safe=" + ",".join(sorted(plan.safe_facade_features)))
    print("facade-unsafe=" + ",".join(sorted(plan.unsafe_facade_features)))
    print(
        "facade-safe-feature-argument="
        + ",".join(sorted(plan.safe_facade_features - {"default"}))
    )


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=MODES)
    parser.add_argument("--cargo", default="cargo", help="Cargo executable")
    parser.add_argument(
        "--target-dir",
        help="isolated Cargo target directory (default: target/non-iwork-gate)",
    )
    parser.add_argument(
        "--record-file",
        help="atomically write a correctness-only JSON execution report",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    execution_mode = args.mode not in {"print", "verify"}
    environment: Mapping[str, str] = {}
    recorder: ExecutionRecorder | None = None
    status = 0
    error_message: str | None = None
    try:
        if execution_mode:
            environment = _environment(args.target_dir)
            if args.record_file:
                recorder = ExecutionRecorder(args.mode, environment)
        plan = derive_plan(_cargo_metadata(args.cargo))
        if args.mode == "print":
            print_plan(plan)
        elif args.mode == "verify":
            print_plan(plan)
            bulk_count, facade_count = verify_dependency_trees(args.cargo, plan)
            print(f"verified_bulk_tree_roots={bulk_count}")
            print(f"verified_facade_safe_trees={facade_count}")
            print("verified_facade_combined_tree=1")
        else:
            run_mode(
                args.cargo,
                plan,
                args.mode,
                environment,
                recorder=recorder,
            )
    except GateError as error:
        status = 1
        error_message = f"non-iWork release gate: {error}"
        if recorder is not None:
            recorder.finish("failed", str(error))
    except OSError as error:
        status = 1
        error_message = f"non-iWork release gate: cannot execute Cargo: {error}"
        if recorder is not None:
            recorder.finish("failed", str(error))
    except KeyboardInterrupt:
        status = 130
        error_message = "non-iWork release gate: interrupted"
        if recorder is not None:
            recorder.finish("interrupted", "KeyboardInterrupt")
    else:
        if recorder is not None:
            recorder.finish("passed", None)

    if recorder is not None and args.record_file:
        try:
            _write_atomic_report(args.record_file, recorder.as_dict())
        except OSError as error:
            report_error = f"non-iWork release gate: cannot write execution report: {error}"
            if status == 0:
                status = 1
                error_message = report_error
            elif error_message:
                error_message = f"{error_message}; {report_error}"
            else:
                error_message = report_error
    if error_message:
        print(error_message, file=sys.stderr)
    return status


if __name__ == "__main__":  # pragma: no cover - exercised by the CLI
    raise SystemExit(main())
