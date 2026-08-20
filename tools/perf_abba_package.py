#!/usr/bin/env python3
"""Package strict ABBA JSON evidence into deterministic zstd artifacts.

The ABBA summarizer emits a small, strict schema whose ``report_identity``
section binds the canonical JSON hash of each of the four reports.  This tool
keeps the reports byte-for-byte unchanged, validates those bindings, and
stores each raw report as a deterministic single-threaded zstd frame.  The
summary is retained as ordinary JSON so it can be read without a decompressor.

Typical use from a clean checkout::

    python3 tools/perf_abba_package.py \
      --change 0238-perf-package \
      --output-dir target/perf/0238 \
      --summary target/perf/summary.json \
      --artifact a1=target/perf/control-a.json \
      --artifact b1=target/perf/candidate-a.json \
      --artifact b2=target/perf/candidate-b.json \
      --artifact a2=target/perf/control-b.json

The output directory receives ``summary.json`` (unless it is already the
same file), four ``*.json.zst`` files, and
``0238-perf-package-manifest.json``.  Output names can be overridden with
``--summary-name`` and ``--artifact-name ROLE=NAME``.  All output paths are
checked after symlink resolution and must remain below ``--output-dir``;
existing output files are never replaced.

Only the standard library is required.  The external ``zstd`` executable is
required for packaging and is invoked without a shell, with one compression
thread, an explicit input size, and an integrity checksum. No timestamps or
compressor process details enter the manifest; the canonical compressor path
is recorded alongside its version and file digest.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path, PurePath
from typing import Any, Iterable, Mapping, Sequence


MANIFEST_SCHEMA_VERSION = 1
MANIFEST_KIND = "litchi-perf-abba-artifacts"
SUMMARY_SCHEMA_VERSION = 1
SUMMARY_TOOL_NAME = "litchi-perf-abba-summary"
ABBA_ROLES = ("a1", "b1", "b2", "a2")
ZSTD_FORMAT = "zstd"
ZSTD_DEFAULT_LEVEL = 3
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
CHANGE_ID_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9._-]*\Z")
ROLE_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9._-]*\Z")


class ArtifactPackagingError(ValueError):
    """Raised when an artifact package cannot be safely produced."""


@dataclass(frozen=True)
class ArtifactSpec:
    """One raw JSON report and its ABBA leg."""

    role: str
    source: Path
    output_name: str | None = None


@dataclass(frozen=True)
class _RawJson:
    value: dict[str, Any]
    raw: bytes
    canonical: bytes
    raw_sha256: str
    canonical_sha256: str


@dataclass(frozen=True)
class _CompressedArtifact:
    role: str
    source: Path
    output_name: str
    raw: _RawJson
    compressed: bytes


@dataclass(frozen=True)
class _SummaryBindings:
    report_hashes: dict[str, str]
    harness_schema_version: int
    harness_tool: dict[str, Any]
    harness_configuration: dict[str, Any]
    environment_stable: dict[str, Any]
    environment_legs: dict[str, dict[str, Any]]
    control_revision: str
    candidate_revision: str
    result_keys: dict[tuple[str, str], str]


@dataclass(frozen=True)
class _ZstdIdentity:
    path: str
    version: str
    sha256: str
    bytes: int


def _reject_nonfinite(value: str) -> None:
    raise ArtifactPackagingError(f"JSON contains non-finite value {value!r}")


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ArtifactPackagingError(f"JSON contains duplicate object key {key!r}")
        result[key] = value
    return result


def _validate_json_tree(value: Any, location: str) -> None:
    """Reject values that canonical JSON cannot represent safely."""

    if value is None or isinstance(value, (bool, str, int)):
        return
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ArtifactPackagingError(f"{location} contains a non-finite number")
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _validate_json_tree(item, f"{location}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            if not isinstance(key, str):
                raise ArtifactPackagingError(f"{location} has a non-string object key")
            _validate_json_tree(item, f"{location}.{key}")
        return
    raise ArtifactPackagingError(
        f"{location} contains unsupported JSON value {type(value).__name__}"
    )


def canonical_json(value: Any, location: str = "value") -> bytes:
    """Return the strict-summary canonical JSON representation."""

    _validate_json_tree(value, location)
    try:
        return json.dumps(
            value,
            sort_keys=True,
            separators=(",", ":"),
            ensure_ascii=True,
            allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise ArtifactPackagingError(f"{location} is not canonical JSON: {error}") from error


def _load_json_bytes(raw: bytes, location: str) -> _RawJson:
    try:
        text = raw.decode("utf-8")
        value = json.loads(
            text,
            object_pairs_hook=_reject_duplicate_keys,
            parse_constant=_reject_nonfinite,
        )
    except UnicodeDecodeError as error:
        raise ArtifactPackagingError(f"{location} is not UTF-8 JSON: {error}") from error
    except json.JSONDecodeError as error:
        raise ArtifactPackagingError(f"{location} is not valid JSON: {error}") from error
    if not isinstance(value, dict):
        raise ArtifactPackagingError(f"{location} must contain a JSON object")
    canonical = canonical_json(value, location)
    return _RawJson(
        value=value,
        raw=raw,
        canonical=canonical,
        raw_sha256=hashlib.sha256(raw).hexdigest(),
        canonical_sha256=hashlib.sha256(canonical).hexdigest(),
    )


def load_json(path: Path) -> _RawJson:
    """Read and strictly parse one JSON file without changing its bytes."""

    try:
        raw = path.read_bytes()
    except OSError as error:
        raise ArtifactPackagingError(f"cannot read {path}: {error}") from error
    return _load_json_bytes(raw, str(path))


def _object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ArtifactPackagingError(f"{location} must be an object")
    return value


def _string(value: Any, location: str, *, nonempty: bool = False) -> str:
    if not isinstance(value, str) or (nonempty and not value):
        requirement = " non-empty" if nonempty else ""
        raise ArtifactPackagingError(f"{location} must be a{requirement} string")
    return value


def _positive_integer(value: Any, location: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ArtifactPackagingError(f"{location} must be a positive integer")
    return value


def _sha256(value: Any, location: str) -> str:
    result = _string(value, location)
    if SHA256_RE.fullmatch(result) is None:
        raise ArtifactPackagingError(f"{location} must be a lowercase SHA-256 digest")
    return result


def _summary_report_identities(summary: dict[str, Any]) -> _SummaryBindings:
    """Validate the strict summary envelope and return its report bindings."""

    if summary.get("schema_version") != SUMMARY_SCHEMA_VERSION:
        raise ArtifactPackagingError(
            f"summary.schema_version must be {SUMMARY_SCHEMA_VERSION}"
        )

    summary_tool = _object(summary.get("tool"), "summary.tool")
    if summary_tool.get("name") != SUMMARY_TOOL_NAME:
        raise ArtifactPackagingError(
            f"summary.tool.name must be {SUMMARY_TOOL_NAME!r}"
        )
    _string(summary_tool.get("version"), "summary.tool.version", nonempty=True)

    harness = _object(summary.get("harness_identity"), "summary.harness_identity")
    _positive_integer(harness.get("schema_version"), "summary.harness_identity.schema_version")
    harness_tool = _object(harness.get("tool"), "summary.harness_identity.tool")
    _string(harness_tool.get("name"), "summary.harness_identity.tool.name", nonempty=True)
    _string(
        harness_tool.get("version"),
        "summary.harness_identity.tool.version",
        nonempty=True,
    )
    harness_configuration = _object(
        harness.get("configuration"), "summary.harness_identity.configuration"
    )

    implementation = _object(
        summary.get("implementation_identity"), "summary.implementation_identity"
    )
    control = _object(implementation.get("control"), "summary.implementation_identity.control")
    candidate = _object(
        implementation.get("candidate"), "summary.implementation_identity.candidate"
    )
    control_revision = _string(
        control.get("git_revision"),
        "summary.implementation_identity.control.git_revision",
        nonempty=True,
    )
    candidate_revision = _string(
        candidate.get("git_revision"),
        "summary.implementation_identity.candidate.git_revision",
        nonempty=True,
    )
    if control_revision == candidate_revision:
        raise ArtifactPackagingError(
            "summary control and candidate git revisions must be distinct"
        )
    if implementation.get("distinct") is not True:
        raise ArtifactPackagingError("summary.implementation_identity.distinct must be true")
    if control.get("legs") != ["a1", "a2"]:
        raise ArtifactPackagingError(
            "summary.implementation_identity.control.legs must be ['a1', 'a2']"
        )
    if candidate.get("legs") != ["b1", "b2"]:
        raise ArtifactPackagingError(
            "summary.implementation_identity.candidate.legs must be ['b1', 'b2']"
        )

    report_identity = _object(summary.get("report_identity"), "summary.report_identity")
    if set(report_identity) != set(ABBA_ROLES):
        raise ArtifactPackagingError(
            "summary.report_identity must contain exactly a1, b1, b2, and a2"
        )
    identities: dict[str, str] = {}
    for role in ABBA_ROLES:
        item = _object(report_identity.get(role), f"summary.report_identity.{role}")
        identities[role] = _sha256(
            item.get("canonical_sha256"),
            f"summary.report_identity.{role}.canonical_sha256",
        )

    results = summary.get("results")
    if not isinstance(results, list) or not results:
        raise ArtifactPackagingError("summary.results must be a non-empty array")
    result_keys: dict[tuple[str, str], str] = {}
    for index, result in enumerate(results):
        result_object = _object(result, f"summary.results[{index}]")
        case = _string(result_object.get("case"), f"summary.results[{index}].case", nonempty=True)
        corpus = _object(result_object.get("corpus"), f"summary.results[{index}].corpus")
        shape = _string(
            result_object.get("shape"),
            f"summary.results[{index}].shape",
            nonempty=True,
        )
        corpus_identity = canonical_json(corpus, f"summary.results[{index}].corpus").decode(
            "utf-8"
        )
        if corpus.get("shape") != shape:
            raise ArtifactPackagingError(
                f"summary.results[{index}].shape disagrees with its corpus shape"
            )
        identity = _object(result_object.get("identity"), f"summary.results[{index}].identity")
        if identity.get("corpus") != corpus_identity:
            raise ArtifactPackagingError(
                f"summary.results[{index}].identity.corpus disagrees with corpus"
            )
        key = (case, corpus_identity)
        if key in result_keys:
            raise ArtifactPackagingError(
                f"summary.results contains duplicate case/corpus identity for {case!r}"
            )
        result_keys[key] = shape
    verification = _object(summary.get("verification"), "summary.verification")
    result_count = _positive_integer(
        verification.get("result_count"), "summary.verification.result_count"
    )
    if result_count != len(results):
        raise ArtifactPackagingError(
            "summary.verification.result_count disagrees with summary.results"
        )
    for field in (
        "tool_identity_verified",
        "configuration_identity_verified",
        "environment_stable_identity_verified",
        "environment_legs_recorded",
        "case_corpus_identity_verified",
        "statistics_recomputed_from_samples",
    ):
        if verification.get(field) is not True:
            raise ArtifactPackagingError(f"summary.verification.{field} must be true")

    environment = _object(summary.get("environment"), "summary.environment")
    environment_stable = _object(environment.get("stable"), "summary.environment.stable")
    legs = _object(environment.get("legs"), "summary.environment.legs")
    if set(legs) != set(ABBA_ROLES):
        raise ArtifactPackagingError("summary.environment.legs must contain all ABBA roles")
    for role in ABBA_ROLES:
        _object(legs[role], f"summary.environment.legs.{role}")

    protocol = _object(summary.get("protocol"), "summary.protocol")
    order = protocol.get("order")
    if order != ["a1_control", "b1_candidate", "b2_candidate", "a2_control"]:
        raise ArtifactPackagingError("summary.protocol.order is not the strict ABBA order")

    environment_legs = {
        role: dict(legs[role])
        for role in ABBA_ROLES
    }
    return _SummaryBindings(
        report_hashes=identities,
        harness_schema_version=harness["schema_version"],
        harness_tool=dict(harness_tool),
        harness_configuration=dict(harness_configuration),
        environment_stable=dict(environment_stable),
        environment_legs=environment_legs,
        control_revision=control_revision,
        candidate_revision=candidate_revision,
        result_keys=result_keys,
    )


def _validate_raw_report(
    raw_report: _RawJson,
    role: str,
    bindings: _SummaryBindings,
) -> None:
    """Cross-check every report identity used by the strict summary."""

    report = raw_report.value
    schema_version = report.get("schema_version")
    if schema_version != bindings.harness_schema_version:
        raise ArtifactPackagingError(
            f"{role} report schema_version disagrees with summary.harness_identity"
        )
    report_tool = _object(report.get("tool"), f"{role}.tool")
    if canonical_json(report_tool, f"{role}.tool") != canonical_json(
        bindings.harness_tool, "summary.harness_identity.tool"
    ):
        raise ArtifactPackagingError(
            f"{role} report tool identity disagrees with summary.harness_identity.tool"
        )
    report_configuration = _object(report.get("configuration"), f"{role}.configuration")
    if canonical_json(report_configuration, f"{role}.configuration") != canonical_json(
        bindings.harness_configuration, "summary.harness_identity.configuration"
    ):
        raise ArtifactPackagingError(
            f"{role} report configuration disagrees with summary.harness_identity.configuration"
        )

    report_environment = _object(report.get("environment"), f"{role}.environment")
    expected_environment = bindings.environment_legs[role]
    if canonical_json(report_environment, f"{role}.environment") != canonical_json(
        expected_environment, f"summary.environment.legs.{role}"
    ):
        raise ArtifactPackagingError(
            f"{role} report environment disagrees with summary.environment.legs"
        )
    expected_revision = (
        bindings.control_revision if role in ("a1", "a2") else bindings.candidate_revision
    )
    if report_environment.get("git_revision") != expected_revision:
        raise ArtifactPackagingError(
            f"{role} report git_revision does not match its summary implementation leg"
        )
    if report_environment.get("git_worktree_dirty") is not False:
        raise ArtifactPackagingError(f"{role} report git_worktree_dirty must be false")
    stable_environment = {
        key: value for key, value in report_environment.items() if key != "git_revision"
    }
    if canonical_json(stable_environment, f"{role}.environment") != canonical_json(
        bindings.environment_stable, "summary.environment.stable"
    ):
        raise ArtifactPackagingError(
            f"{role} report stable environment disagrees with summary.environment.stable"
        )

    report_results = report.get("results")
    if not isinstance(report_results, list) or not report_results:
        raise ArtifactPackagingError(f"{role}.results must be a non-empty array")
    observed_keys: dict[tuple[str, str], str] = {}
    for index, result in enumerate(report_results):
        result_object = _object(result, f"{role}.results[{index}]")
        case = _string(result_object.get("case"), f"{role}.results[{index}].case", nonempty=True)
        corpus = _object(result_object.get("corpus"), f"{role}.results[{index}].corpus")
        shape = _string(
            corpus.get("shape"), f"{role}.results[{index}].corpus.shape", nonempty=True
        )
        corpus_identity = canonical_json(corpus, f"{role}.results[{index}].corpus").decode(
            "utf-8"
        )
        key = (case, corpus_identity)
        if key in observed_keys:
            raise ArtifactPackagingError(
                f"{role}.results contains duplicate case/corpus identity for {case!r}"
            )
        observed_keys[key] = shape
    if observed_keys != bindings.result_keys:
        missing = sorted(set(bindings.result_keys) - set(observed_keys))
        extra = sorted(set(observed_keys) - set(bindings.result_keys))
        raise ArtifactPackagingError(
            f"{role}.results case/corpus/shape identity disagrees with summary "
            f"(missing={missing!r}, extra={extra!r})"
        )
    for key, shape in observed_keys.items():
        if bindings.result_keys[key] != shape:
            raise ArtifactPackagingError(
                f"{role}.results shape disagrees with summary for {key[0]!r}"
            )


def _normalize_role(role: str) -> str:
    role_text = _string(role, "artifact role", nonempty=True).strip().lower()
    if ROLE_RE.fullmatch(role_text) is None:
        raise ArtifactPackagingError(f"invalid artifact role {role!r}")
    compact = role_text.replace("-", "").replace("_", "").replace(".", "")
    aliases = {
        "a1": "a1",
        "a1control": "a1",
        "controla": "a1",
        "beforea": "a1",
        "b1": "b1",
        "b1candidate": "b1",
        "candidatea": "b1",
        "aftera": "b1",
        "b2": "b2",
        "b2candidate": "b2",
        "candidateb": "b2",
        "afterb": "b2",
        "a2": "a2",
        "a2control": "a2",
        "controlb": "a2",
        "beforeb": "a2",
    }
    normalized = aliases.get(compact)
    if normalized is None:
        raise ArtifactPackagingError(
            f"artifact role {role!r} must identify one of {', '.join(ABBA_ROLES)}"
        )
    return normalized


def _coerce_specs(
    artifacts: Mapping[str, Path] | Iterable[ArtifactSpec | tuple[str, Path]],
) -> tuple[ArtifactSpec, ...]:
    if isinstance(artifacts, Mapping):
        values: Iterable[ArtifactSpec | tuple[str, Path]] = (
            (role, path) for role, path in artifacts.items()
        )
    else:
        values = artifacts
    specs: list[ArtifactSpec] = []
    seen: set[str] = set()
    for item in values:
        if isinstance(item, ArtifactSpec):
            spec = item
        else:
            try:
                role, source = item
            except (TypeError, ValueError) as error:
                raise ArtifactPackagingError(
                    "artifacts must contain ArtifactSpec or (role, path) items"
                ) from error
            spec = ArtifactSpec(role=str(role), source=Path(source))
        role = _normalize_role(spec.role)
        if role in seen:
            raise ArtifactPackagingError(f"duplicate artifact role {role!r}")
        seen.add(role)
        source = Path(spec.source)
        if not source.name:
            raise ArtifactPackagingError(f"artifact {role} has an empty source name")
        output_name = spec.output_name
        if output_name is not None:
            output_name = _string(
                output_name, f"artifact {role} output name", nonempty=True
            )
        specs.append(ArtifactSpec(role=role, source=source, output_name=output_name))
    missing = [role for role in ABBA_ROLES if role not in seen]
    if missing:
        raise ArtifactPackagingError(
            f"missing ABBA artifact role(s): {', '.join(missing)}"
        )
    return tuple(sorted(specs, key=lambda spec: ABBA_ROLES.index(spec.role)))


def _path_exists(path: Path) -> bool:
    # Path.exists() is false for a broken symlink.  A broken symlink is still
    # an existing destination and must not be replaced by this tool.
    return os.path.lexists(path)


def _resolved_output_path(root: Path, relative_name: str, location: str) -> tuple[Path, str]:
    if not isinstance(relative_name, str) or not relative_name:
        raise ArtifactPackagingError(f"{location} must be a non-empty relative path")
    if "\\" in relative_name:
        raise ArtifactPackagingError(f"{location} must use '/' path separators")
    relative = PurePath(relative_name)
    if relative.is_absolute():
        raise ArtifactPackagingError(f"{location} must not be absolute")
    candidate = (root / Path(relative_name)).resolve(strict=False)
    try:
        candidate.relative_to(root)
    except ValueError as error:
        raise ArtifactPackagingError(f"{location} escapes output directory") from error
    normalized = candidate.relative_to(root).as_posix()
    if normalized in ("", "."):
        raise ArtifactPackagingError(f"{location} must identify a file")
    return candidate, normalized


def _default_artifact_name(source: Path) -> str:
    return f"{source.name}.zst"


def _validate_change_id(change_id: str) -> str:
    value = _string(change_id, "change id", nonempty=True)
    if CHANGE_ID_RE.fullmatch(value) is None:
        raise ArtifactPackagingError(
            "change id may contain only ASCII letters, digits, '.', '_' and '-'"
        )
    return value


def compress_zstd(
    raw: bytes,
    *,
    executable: str | os.PathLike[str] = "zstd",
    level: int = ZSTD_DEFAULT_LEVEL,
) -> bytes:
    """Compress raw bytes with deterministic, single-threaded zstd."""

    if isinstance(level, bool) or not isinstance(level, int) or not 1 <= level <= 22:
        raise ArtifactPackagingError("zstd compression level must be an integer from 1 to 22")
    if not isinstance(raw, bytes):
        raise ArtifactPackagingError("zstd input must be bytes")
    command = [
        os.fspath(executable),
        "--quiet",
        "--stdout",
        "--compress",
        f"-{level}",
        "--threads=1",
        "--check",
        f"--stream-size={len(raw)}",
    ]
    try:
        completed = subprocess.run(
            command,
            input=raw,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            env={**os.environ, "LC_ALL": "C"},
        )
    except OSError as error:
        raise ArtifactPackagingError(f"cannot execute zstd: {error}") from error
    if completed.returncode != 0:
        detail = completed.stderr.decode("utf-8", errors="replace").strip()
        if detail:
            raise ArtifactPackagingError(f"zstd compression failed: {detail}")
        raise ArtifactPackagingError(
            f"zstd compression failed with exit status {completed.returncode}"
        )
    return completed.stdout


def _resolve_zstd_identity(executable: str | os.PathLike[str]) -> _ZstdIdentity:
    requested = os.fspath(executable)
    found = shutil.which(requested)
    if found is None:
        raise ArtifactPackagingError(f"cannot find executable {requested!r}")
    path = Path(found).resolve(strict=False)
    try:
        path = path.resolve(strict=True)
        stat = path.stat()
        executable_bytes = path.read_bytes()
    except OSError as error:
        raise ArtifactPackagingError(f"cannot inspect zstd executable {path}: {error}") from error
    if not path.is_file() or not os.access(path, os.X_OK):
        raise ArtifactPackagingError(f"zstd executable is not an executable file: {path}")
    try:
        completed = subprocess.run(
            [os.fspath(path), "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            env={**os.environ, "LC_ALL": "C"},
        )
    except OSError as error:
        raise ArtifactPackagingError(f"cannot execute zstd version probe: {error}") from error
    if completed.returncode != 0:
        detail = completed.stderr.decode("utf-8", errors="replace").strip()
        raise ArtifactPackagingError(
            f"zstd version probe failed{': ' + detail if detail else ''}"
        )
    version_output = (
        completed.stdout.decode("utf-8", errors="replace")
        + completed.stderr.decode("utf-8", errors="replace")
    )
    version_lines = [line.strip() for line in version_output.splitlines() if line.strip()]
    if not version_lines:
        raise ArtifactPackagingError("zstd version probe returned no version")
    return _ZstdIdentity(
        path=os.fspath(path),
        version=version_lines[0],
        sha256=hashlib.sha256(executable_bytes).hexdigest(),
        bytes=stat.st_size,
    )


def _write_exclusive(path: Path, data: bytes, location: str) -> None:
    try:
        with path.open("xb") as handle:
            handle.write(data)
    except FileExistsError as error:
        raise ArtifactPackagingError(f"refusing to overwrite {location}: {path}") from error
    except OSError as error:
        raise ArtifactPackagingError(f"cannot write {location} {path}: {error}") from error


def _ensure_directory(path: Path) -> list[Path]:
    """Create a directory and return only directories created by this call."""

    path = Path(path)
    if _path_exists(path):
        if path.is_symlink() or not path.is_dir():
            raise ArtifactPackagingError(f"output parent is not a directory: {path}")
        return []
    missing: list[Path] = []
    current = path
    while not _path_exists(current):
        missing.append(current)
        parent = current.parent
        if parent == current:
            raise ArtifactPackagingError(f"cannot find parent for output directory: {path}")
        current = parent
    if current.is_symlink() or not current.is_dir():
        raise ArtifactPackagingError(f"output parent is not a directory: {current}")
    created: list[Path] = []
    try:
        for directory in reversed(missing):
            directory.mkdir()
            created.append(directory)
    except OSError as error:
        for directory in reversed(created):
            try:
                directory.rmdir()
            except OSError:
                pass
        raise ArtifactPackagingError(f"cannot create output directory {path}: {error}") from error
    return created


def _publish_exclusive(staged: Path, destination: Path, location: str) -> None:
    """Publish a staged file without replacing a destination.

    The staging hard link remains until the caller removes the whole private
    staging directory.  Keeping that cleanup in one place means a successful
    link is always tracked by the caller before any later cleanup can fail.
    """

    try:
        os.link(os.fspath(staged), os.fspath(destination))
    except FileExistsError as error:
        raise ArtifactPackagingError(
            f"refusing to overwrite {location}: {destination}"
        ) from error
    except OSError as error:
        raise ArtifactPackagingError(f"cannot publish {location} {destination}: {error}") from error


def _remove_empty_directories(directories: Iterable[Path]) -> None:
    for directory in reversed(tuple(directories)):
        try:
            directory.rmdir()
        except OSError:
            pass


def build_manifest(
    *,
    change_id: str,
    summary_identity: Mapping[str, Any],
    artifacts: Sequence[Mapping[str, Any]],
    compression_level: int = ZSTD_DEFAULT_LEVEL,
    zstd_identity: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Build the deterministic manifest object without writing it."""

    change = _validate_change_id(change_id)
    if isinstance(compression_level, bool) or not isinstance(compression_level, int):
        raise ArtifactPackagingError("compression level must be an integer")
    ordered_artifacts = sorted(
        (dict(item) for item in artifacts),
        key=lambda item: (ABBA_ROLES.index(item["role"]), item["path"]),
    )
    compression: dict[str, Any] = {
        "format": ZSTD_FORMAT,
        "level": compression_level,
        "threads": 1,
        "checksum": "XXH64",
        "content_size": True,
    }
    if zstd_identity is not None:
        compression["executable"] = dict(zstd_identity)
    return {
        "schema_version": MANIFEST_SCHEMA_VERSION,
        "manifest_kind": MANIFEST_KIND,
        "change_id": change,
        # Existing performance manifests call this field ``change``.  Keep
        # the spelling as a compatibility alias while ``change_id`` is the
        # unambiguous field used by this schema.
        "change": change,
        "compression": compression,
        "summary_identity": dict(summary_identity),
        "summary": dict(summary_identity),
        "artifacts": ordered_artifacts,
        "self_excluded": True,
    }


def _manifest_bytes(manifest: Mapping[str, Any]) -> bytes:
    try:
        return (
            json.dumps(
                manifest,
                sort_keys=True,
                indent=2,
                ensure_ascii=True,
                allow_nan=False,
            )
            + "\n"
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise ArtifactPackagingError(f"manifest is not deterministic JSON: {error}") from error


def package_artifacts(
    *,
    change_id: str,
    output_dir: Path,
    summary: Path,
    artifacts: Mapping[str, Path] | Iterable[ArtifactSpec | tuple[str, Path]],
    summary_name: str | None = None,
    manifest_name: str | None = None,
    artifact_names: Mapping[str, str] | None = None,
    zstd_executable: str | os.PathLike[str] = "zstd",
    compression_level: int = ZSTD_DEFAULT_LEVEL,
) -> dict[str, Any]:
    """Validate and package four strict ABBA reports.

    All input files are read and validated before the first destination is
    created.  Destination files are created exclusively; if a later write
    fails, files created by this invocation are removed and the existing
    destination is untouched.
    """

    change = _validate_change_id(change_id)
    specs = _coerce_specs(artifacts)
    if artifact_names is None:
        artifact_names = {}
    normalized_names: dict[str, str] = {}
    for role, name in artifact_names.items():
        normalized_role = _normalize_role(role)
        if normalized_role in normalized_names:
            raise ArtifactPackagingError(f"duplicate artifact-name role {role!r}")
        normalized_names[normalized_role] = _string(
            name, f"artifact-name {role}", nonempty=True
        )
    unknown_names = set(normalized_names) - set(ABBA_ROLES)
    if unknown_names:
        raise ArtifactPackagingError(
            f"artifact names contain unknown role(s): {', '.join(sorted(unknown_names))}"
        )

    root = Path(output_dir).expanduser().resolve(strict=False)
    if _path_exists(root) and not root.is_dir():
        raise ArtifactPackagingError(f"output directory is not a directory: {root}")
    raw_output_dir = Path(output_dir).expanduser()
    if _path_exists(raw_output_dir) and raw_output_dir.is_symlink():
        raise ArtifactPackagingError("output directory must not be a symlink")

    summary_path = Path(summary)
    summary_json = load_json(summary_path)
    summary_bindings = _summary_report_identities(summary_json.value)
    zstd_identity = _resolve_zstd_identity(zstd_executable)

    compressed: list[_CompressedArtifact] = []
    destination_names: list[tuple[str, Path, str]] = []
    seen_raw_hashes: dict[str, str] = {}
    seen_canonical_hashes: dict[str, str] = {}
    seen_source_paths: dict[Path, str] = {}
    for spec in specs:
        raw_report = load_json(spec.source)
        _validate_raw_report(raw_report, spec.role, summary_bindings)
        expected = summary_bindings.report_hashes[spec.role]
        if raw_report.canonical_sha256 != expected:
            raise ArtifactPackagingError(
                f"{spec.role} report canonical SHA-256 does not match summary.report_identity"
            )
        previous_role = seen_raw_hashes.get(raw_report.raw_sha256)
        if previous_role is not None:
            raise ArtifactPackagingError(
                f"raw report bytes are reused for roles {previous_role} and {spec.role}"
            )
        seen_raw_hashes[raw_report.raw_sha256] = spec.role
        previous_canonical_role = seen_canonical_hashes.get(raw_report.canonical_sha256)
        if previous_canonical_role is not None:
            raise ArtifactPackagingError(
                "canonical raw report identity is reused for roles "
                f"{previous_canonical_role} and {spec.role}"
            )
        seen_canonical_hashes[raw_report.canonical_sha256] = spec.role
        source_path = spec.source.resolve(strict=True)
        previous_path_role = seen_source_paths.get(source_path)
        if previous_path_role is not None:
            raise ArtifactPackagingError(
                f"raw report path is reused for roles {previous_path_role} and {spec.role}"
            )
        seen_source_paths[source_path] = spec.role
        requested_name = normalized_names.get(spec.role, spec.output_name)
        output_name = (
            _default_artifact_name(spec.source)
            if requested_name is None
            else requested_name
        )
        if not output_name.endswith(".zst"):
            output_name = f"{output_name}.zst"
        destination, relative = _resolved_output_path(
            root, output_name, f"artifact {spec.role} output name"
        )
        destination_names.append((spec.role, destination, relative))
        compressed.append(
            _CompressedArtifact(
                role=spec.role,
                source=spec.source,
                output_name=relative,
                raw=raw_report,
                compressed=compress_zstd(
                    raw_report.raw,
                    executable=zstd_identity.path,
                    level=compression_level,
                ),
            )
        )

    if len({destination for _, destination, _ in destination_names}) != len(destination_names):
        raise ArtifactPackagingError("artifact output names must be unique")

    requested_summary_name = summary_path.name if summary_name is None else summary_name
    summary_destination, summary_relative = _resolved_output_path(
        root, requested_summary_name, "summary output name"
    )
    requested_manifest_name = (
        f"{change}-manifest.json" if manifest_name is None else manifest_name
    )
    manifest_destination, manifest_relative = _resolved_output_path(
        root, requested_manifest_name, "manifest output name"
    )

    all_destinations = [destination for _, destination, _ in destination_names]
    all_destinations.extend((summary_destination, manifest_destination))
    if len(set(all_destinations)) != len(all_destinations):
        raise ArtifactPackagingError("artifact, summary, and manifest output names must be unique")
    summary_source_resolved = summary_path.resolve(strict=False)
    summary_copy_needed = summary_source_resolved != summary_destination
    for role, destination, _ in destination_names:
        if _path_exists(destination):
            raise ArtifactPackagingError(
                f"refusing to overwrite artifact {role} at {destination}"
            )
    if summary_copy_needed and _path_exists(summary_destination):
        raise ArtifactPackagingError(f"refusing to overwrite summary at {summary_destination}")
    if _path_exists(manifest_destination):
        raise ArtifactPackagingError(f"refusing to overwrite manifest at {manifest_destination}")

    summary_identity = {
        "path": summary_relative,
        "bytes": len(summary_json.raw),
        "sha256": summary_json.raw_sha256,
        "canonical_bytes": len(summary_json.canonical),
        "canonical_sha256": summary_json.canonical_sha256,
        "schema_version": summary_json.value["schema_version"],
        "tool": dict(_object(summary_json.value["tool"], "summary.tool")),
        "result_count": summary_json.value["verification"]["result_count"],
        "report_identity": dict(summary_bindings.report_hashes),
    }
    artifact_entries: list[dict[str, Any]] = []
    for item in compressed:
        artifact_entries.append(
            {
                "role": item.role,
                "path": item.output_name,
                "compression": ZSTD_FORMAT,
                "bytes": len(item.compressed),
                "sha256": hashlib.sha256(item.compressed).hexdigest(),
                "uncompressed_bytes": len(item.raw.raw),
                "uncompressed_sha256": item.raw.raw_sha256,
                "canonical_sha256": item.raw.canonical_sha256,
            }
        )
    manifest = build_manifest(
        change_id=change,
        summary_identity=summary_identity,
        artifacts=artifact_entries,
        compression_level=compression_level,
        zstd_identity={
            "path": zstd_identity.path,
            "version": zstd_identity.version,
            "sha256": zstd_identity.sha256,
            "bytes": zstd_identity.bytes,
        },
    )
    manifest["manifest_path"] = manifest_relative

    created_directories: list[Path] = []
    published: list[Path] = []
    stage_directory: Path | None = None
    try:
        created_directories.extend(_ensure_directory(root))
        try:
            stage_directory = Path(
                tempfile.mkdtemp(prefix=f".{change}.staging-", dir=root)
            )
        except OSError as error:
            raise ArtifactPackagingError(
                f"cannot create staging directory in {root}: {error}"
            ) from error

        staged_artifacts: list[tuple[_CompressedArtifact, Path, Path]] = []
        for index, item in enumerate(compressed):
            staged = stage_directory / f"artifact-{index:02d}.json.zst"
            _write_exclusive(staged, item.compressed, f"staged artifact {item.role}")
            staged_artifacts.append((item, staged, destination_names[index][1]))
        staged_summary: tuple[Path, Path] | None = None
        if summary_copy_needed:
            staged = stage_directory / "summary.json"
            _write_exclusive(staged, summary_json.raw, "staged summary")
            staged_summary = (staged, summary_destination)
        staged_manifest = stage_directory / "manifest.json"
        _write_exclusive(staged_manifest, _manifest_bytes(manifest), "staged manifest")

        for item, staged, destination in staged_artifacts:
            created_directories.extend(_ensure_directory(destination.parent))
            _publish_exclusive(staged, destination, f"artifact {item.role}")
            published.append(destination)
        if staged_summary is not None:
            staged, destination = staged_summary
            created_directories.extend(_ensure_directory(destination.parent))
            _publish_exclusive(staged, destination, "summary")
            published.append(destination)
        created_directories.extend(_ensure_directory(manifest_destination.parent))
        _publish_exclusive(staged_manifest, manifest_destination, "manifest")
        published.append(manifest_destination)
    except BaseException as error:
        for path in reversed(published):
            try:
                path.unlink()
            except OSError:
                pass
        if stage_directory is not None:
            shutil.rmtree(stage_directory, ignore_errors=True)
        _remove_empty_directories(created_directories)
        if isinstance(error, ArtifactPackagingError):
            raise
        if isinstance(error, OSError):
            raise ArtifactPackagingError(f"artifact publication failed: {error}") from error
        raise
    finally:
        if stage_directory is not None and stage_directory.exists():
            shutil.rmtree(stage_directory, ignore_errors=True)
    return manifest


def _parse_assignment(value: str, option: str) -> tuple[str, str]:
    role, separator, path = value.partition("=")
    if not separator or not role or not path:
        raise ArtifactPackagingError(f"{option} must use ROLE=PATH")
    return role, path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--change", "--change-id", dest="change_id", required=True)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--summary", required=True, type=Path)
    parser.add_argument(
        "--artifact",
        action="append",
        required=True,
        metavar="ROLE=PATH",
        help="raw JSON report; repeat once for a1, b1, b2, and a2",
    )
    parser.add_argument("--summary-name", metavar="NAME")
    parser.add_argument("--manifest-name", metavar="NAME")
    parser.add_argument(
        "--artifact-name",
        action="append",
        metavar="ROLE=NAME",
        help="override the output name for one compressed report",
    )
    parser.add_argument("--zstd", dest="zstd_executable", default="zstd")
    parser.add_argument("--compression-level", type=int, default=ZSTD_DEFAULT_LEVEL)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = build_parser().parse_args(argv)
    try:
        specs: list[ArtifactSpec] = []
        for item in arguments.artifact:
            role, path = _parse_assignment(item, "--artifact")
            specs.append(ArtifactSpec(role=role, source=Path(path)))
        artifact_names: dict[str, str] = {}
        for item in arguments.artifact_name or ():
            role, name = _parse_assignment(item, "--artifact-name")
            normalized_role = _normalize_role(role)
            if normalized_role in artifact_names:
                raise ArtifactPackagingError(f"duplicate --artifact-name role {role!r}")
            artifact_names[normalized_role] = name
        manifest = package_artifacts(
            change_id=arguments.change_id,
            output_dir=arguments.output_dir,
            summary=arguments.summary,
            artifacts=specs,
            summary_name=arguments.summary_name,
            manifest_name=arguments.manifest_name,
            artifact_names=artifact_names,
            zstd_executable=arguments.zstd_executable,
            compression_level=arguments.compression_level,
        )
        json.dump(manifest, sys.stdout, sort_keys=True, indent=2, allow_nan=False)
        sys.stdout.write("\n")
        return 0
    except (ArtifactPackagingError, OSError, ValueError) as error:
        print(f"perf-abba-package: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
