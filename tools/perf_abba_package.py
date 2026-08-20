#!/usr/bin/env python3
"""Package strict ABBA JSON evidence into deterministic zstd artifacts.

The ABBA summarizer emits a small, strict schema whose ``report_identity``
section binds the canonical JSON hash of each of the four reports. This tool
keeps the reports byte-for-byte unchanged, recomputes the complete summary
with that canonical implementation, and requires the supplied summary to
match it exactly before storing each raw report as a deterministic
single-threaded zstd frame. The summary is retained as ordinary JSON so it can
be read without a decompressor.

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
``--summary-name`` and ``--artifact-name ROLE=NAME``; every output name must
be a flat basename. The output directory is opened once with
``O_DIRECTORY|O_NOFOLLOW`` and all staging, publication, and cleanup use its
held descriptor, so a later pathname swap cannot redirect the package.
Existing output files are never replaced.

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
import secrets
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path, PurePath
from typing import Any, Iterable, Mapping, Sequence

if __package__:
    from . import perf_abba_summary
else:  # pragma: no cover - exercised by the direct CLI entry point
    import perf_abba_summary


MANIFEST_SCHEMA_VERSION = 1
MANIFEST_KIND = "litchi-perf-abba-artifacts"
ABBA_ROLES = ("a1", "b1", "b2", "a2")
ZSTD_FORMAT = "zstd"
ZSTD_DEFAULT_LEVEL = 3
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
class _ZstdIdentity:
    path: str
    version: str
    sha256: str
    bytes: int


@dataclass(frozen=True)
class _CreatedDirectory:
    parent_fd: int
    name: str
    fd: int


@dataclass(frozen=True)
class _OpenedDirectory:
    path: Path
    fd: int
    open_fds: tuple[int, ...]
    created: tuple[_CreatedDirectory, ...]


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


def _recompute_summary(raw_reports: Sequence[_RawJson]) -> dict[str, Any]:
    """Recompute the complete summary with the canonical summary implementation."""

    try:
        return perf_abba_summary.summarize_reports(
            [raw_report.value for raw_report in raw_reports]
        )
    except Exception as error:
        raise ArtifactPackagingError(
            f"cannot recompute canonical ABBA summary: {error}"
        ) from error


def _require_canonical_summary(
    supplied: _RawJson, raw_reports: Sequence[_RawJson]
) -> dict[str, Any]:
    """Require the supplied summary to equal the canonical recomputation exactly."""

    recomputed = _recompute_summary(raw_reports)
    recomputed_canonical = canonical_json(recomputed, "recomputed summary")
    if supplied.canonical != recomputed_canonical:
        raise ArtifactPackagingError(
            "supplied summary does not exactly match canonical ABBA recomputation"
        )
    return recomputed


def _object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ArtifactPackagingError(f"{location} must be an object")
    return value


def _string(value: Any, location: str, *, nonempty: bool = False) -> str:
    if not isinstance(value, str) or (nonempty and not value):
        requirement = " non-empty" if nonempty else ""
        raise ArtifactPackagingError(f"{location} must be a{requirement} string")
    return value


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


def _flat_basename(name: str, location: str) -> str:
    """Validate an output name that can be used with a held directory fd."""

    if not isinstance(name, str) or not name:
        raise ArtifactPackagingError(f"{location} must be a non-empty basename")
    path = PurePath(name)
    if (
        name in (".", "..")
        or "/" in name
        or "\\" in name
        or path.is_absolute()
        or path.name != name
    ):
        raise ArtifactPackagingError(
            f"{location} escapes or nests below the output directory; use a flat basename"
        )
    return name


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


def _directory_open_flags() -> int:
    required = ("O_DIRECTORY", "O_NOFOLLOW")
    if any(not hasattr(os, name) for name in required):
        raise ArtifactPackagingError(
            "safe artifact publication requires O_DIRECTORY and O_NOFOLLOW"
        )
    return (
        os.O_RDONLY
        | os.O_DIRECTORY
        | os.O_NOFOLLOW
        | getattr(os, "O_CLOEXEC", 0)
    )


def _close_fd(file_descriptor: int) -> None:
    if file_descriptor < 0:
        return
    try:
        os.close(file_descriptor)
    except OSError:
        pass


def _cleanup_created_directories(opened: _OpenedDirectory) -> None:
    """Remove only directories created while opening the output path."""

    for created in reversed(opened.created):
        _close_fd(created.fd)
        try:
            os.rmdir(created.name, dir_fd=created.parent_fd)
        except OSError:
            pass
    for file_descriptor in reversed(opened.open_fds):
        _close_fd(file_descriptor)


def _close_opened_directory(opened: _OpenedDirectory) -> None:
    for file_descriptor in reversed(opened.open_fds):
        _close_fd(file_descriptor)


def _open_output_directory(path: Path) -> _OpenedDirectory:
    """Open/create an output directory with every component no-followed."""

    flags = _directory_open_flags()
    absolute = Path(os.path.abspath(os.fspath(path)))
    file_descriptors: list[int] = []
    created: list[_CreatedDirectory] = []
    try:
        current_fd = os.open(os.sep, flags)
        file_descriptors.append(current_fd)
        for component in absolute.parts[1:]:
            created_by_call = False
            try:
                child_fd = os.open(component, flags, dir_fd=current_fd)
            except FileNotFoundError:
                try:
                    os.mkdir(component, mode=0o700, dir_fd=current_fd)
                    created_by_call = True
                except FileExistsError:
                    pass
                try:
                    child_fd = os.open(component, flags, dir_fd=current_fd)
                except OSError:
                    if created_by_call:
                        try:
                            os.rmdir(component, dir_fd=current_fd)
                        except OSError:
                            pass
                    raise
            if created_by_call:
                created.append(_CreatedDirectory(current_fd, component, child_fd))
            file_descriptors.append(child_fd)
            current_fd = child_fd
    except OSError as error:
        partial = _OpenedDirectory(
            absolute,
            current_fd if file_descriptors else -1,
            tuple(file_descriptors),
            tuple(created),
        )
        _cleanup_created_directories(partial)
        raise ArtifactPackagingError(
            f"cannot open output directory {absolute}: {error}"
        ) from error
    return _OpenedDirectory(absolute, current_fd, tuple(file_descriptors), tuple(created))


def _entry_exists(directory_fd: int, name: str) -> bool:
    try:
        os.stat(name, dir_fd=directory_fd, follow_symlinks=False)
    except FileNotFoundError:
        return False
    except OSError as error:
        raise ArtifactPackagingError(f"cannot inspect output entry {name}: {error}") from error
    return True


def _write_exclusive(
    directory_fd: int,
    name: str,
    data: bytes,
    location: str,
) -> None:
    """Write one exclusive file relative to a held directory descriptor."""

    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_NOFOLLOW", 0)
    file_descriptor = -1
    created = False
    success = False
    try:
        file_descriptor = os.open(name, flags, mode=0o600, dir_fd=directory_fd)
        created = True
        handle = os.fdopen(file_descriptor, "wb")
        file_descriptor = -1
        with handle:
            handle.write(data)
        success = True
    except FileExistsError as error:
        raise ArtifactPackagingError(f"refusing to overwrite {location}: {name}") from error
    except OSError as error:
        raise ArtifactPackagingError(f"cannot write {location} {name}: {error}") from error
    finally:
        _close_fd(file_descriptor)
        if created and not success:
            try:
                os.unlink(name, dir_fd=directory_fd)
            except OSError:
                pass


def _create_staging_directory(root_fd: int, change: str) -> tuple[str, int]:
    flags = _directory_open_flags()
    for _ in range(128):
        name = f".{change}.staging-{secrets.token_hex(16)}"
        try:
            os.mkdir(name, mode=0o700, dir_fd=root_fd)
        except FileExistsError:
            continue
        except OSError as error:
            raise ArtifactPackagingError(
                f"cannot create staging directory in output directory: {error}"
            ) from error
        try:
            return name, os.open(name, flags, dir_fd=root_fd)
        except OSError as error:
            try:
                os.rmdir(name, dir_fd=root_fd)
            except OSError:
                pass
            raise ArtifactPackagingError(
                f"cannot open staging directory in output directory: {error}"
            ) from error
    raise ArtifactPackagingError("cannot create a unique staging directory")


def _publish_exclusive(
    staged_fd: int,
    staged_name: str,
    root_fd: int,
    destination_name: str,
    location: str,
) -> None:
    """Publish one staged file using only held directory descriptors."""

    try:
        os.link(
            staged_name,
            destination_name,
            src_dir_fd=staged_fd,
            dst_dir_fd=root_fd,
            follow_symlinks=False,
        )
    except FileExistsError as error:
        raise ArtifactPackagingError(
            f"refusing to overwrite {location}: {destination_name}"
        ) from error
    except OSError as error:
        raise ArtifactPackagingError(
            f"cannot publish {location} {destination_name}: {error}"
        ) from error


def _remove_staging_directory(
    root_fd: int,
    stage_name: str,
    stage_fd: int,
    staged_names: Iterable[str],
) -> None:
    errors: list[OSError] = []
    for name in reversed(tuple(staged_names)):
        try:
            os.unlink(name, dir_fd=stage_fd)
        except FileNotFoundError:
            pass
        except OSError as error:
            errors.append(error)
    try:
        os.rmdir(stage_name, dir_fd=root_fd)
    except FileNotFoundError:
        pass
    except OSError as error:
        errors.append(error)
    if errors:
        raise ArtifactPackagingError(
            f"cannot remove staging directory {stage_name}: {errors[0]}"
        ) from errors[0]


def _remove_published(root_fd: int, names: Iterable[str]) -> None:
    for name in reversed(tuple(names)):
        try:
            os.unlink(name, dir_fd=root_fd)
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

    raw_output_dir = Path(output_dir).expanduser()
    root = Path(os.path.abspath(os.fspath(raw_output_dir)))
    if os.path.lexists(raw_output_dir) and raw_output_dir.is_symlink():
        raise ArtifactPackagingError("output directory must not be a symlink")

    summary_path = Path(summary)
    summary_json = load_json(summary_path)
    raw_reports = [load_json(spec.source) for spec in specs]
    recomputed_summary = _require_canonical_summary(summary_json, raw_reports)
    zstd_identity = _resolve_zstd_identity(zstd_executable)

    compressed: list[_CompressedArtifact] = []
    destination_names: list[tuple[str, str]] = []
    seen_raw_hashes: dict[str, str] = {}
    seen_canonical_hashes: dict[str, str] = {}
    seen_source_paths: dict[Path, str] = {}
    report_by_role = dict(zip((spec.role for spec in specs), raw_reports))
    for spec, raw_report in zip(specs, raw_reports):
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
        destination_name = _flat_basename(output_name, f"artifact {spec.role} output name")
        destination_names.append((spec.role, destination_name))
        compressed.append(
            _CompressedArtifact(
                role=spec.role,
                source=spec.source,
                output_name=destination_name,
                raw=raw_report,
                compressed=compress_zstd(
                    raw_report.raw,
                    executable=zstd_identity.path,
                    level=compression_level,
                ),
            )
        )

    if len({destination for _, destination in destination_names}) != len(destination_names):
        raise ArtifactPackagingError("artifact output names must be unique")

    requested_summary_name = summary_path.name if summary_name is None else summary_name
    summary_name = _flat_basename(requested_summary_name, "summary output name")
    requested_manifest_name = (
        f"{change}-manifest.json" if manifest_name is None else manifest_name
    )
    manifest_name = _flat_basename(requested_manifest_name, "manifest output name")

    all_destinations = [destination for _, destination in destination_names]
    all_destinations.extend((summary_name, manifest_name))
    if len(set(all_destinations)) != len(all_destinations):
        raise ArtifactPackagingError("artifact, summary, and manifest output names must be unique")

    summary_identity = {
        "path": summary_name,
        "bytes": len(summary_json.raw),
        "sha256": summary_json.raw_sha256,
        "canonical_bytes": len(summary_json.canonical),
        "canonical_sha256": summary_json.canonical_sha256,
        "schema_version": recomputed_summary["schema_version"],
        "tool": dict(_object(recomputed_summary["tool"], "summary.tool")),
        "result_count": recomputed_summary["verification"]["result_count"],
        "report_identity": {
            role: {"canonical_sha256": report_by_role[role].canonical_sha256}
            for role in ABBA_ROLES
        },
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
    manifest["manifest_path"] = manifest_name

    opened: _OpenedDirectory | None = None
    stage_name: str | None = None
    stage_fd = -1
    staged_names: list[str] = []
    published: list[str] = []
    try:
        opened = _open_output_directory(root)
        root_fd = opened.fd
        for role, destination_name in destination_names:
            if _entry_exists(root_fd, destination_name):
                raise ArtifactPackagingError(
                    f"refusing to overwrite artifact {role} at {destination_name}"
                )
        if _entry_exists(root_fd, manifest_name):
            raise ArtifactPackagingError(f"refusing to overwrite manifest at {manifest_name}")

        summary_copy_needed = True
        if _entry_exists(root_fd, summary_name):
            try:
                summary_stat = os.stat(summary_path, follow_symlinks=True)
                destination_stat = os.stat(
                    summary_name, dir_fd=root_fd, follow_symlinks=False
                )
            except OSError as error:
                raise ArtifactPackagingError(
                    f"cannot compare existing summary destination {summary_name}: {error}"
                ) from error
            if (
                summary_stat.st_dev == destination_stat.st_dev
                and summary_stat.st_ino == destination_stat.st_ino
            ):
                summary_copy_needed = False
            else:
                raise ArtifactPackagingError(
                    f"refusing to overwrite summary at {summary_name}"
                )

        stage_name, stage_fd = _create_staging_directory(root_fd, change)

        for index, item in enumerate(compressed):
            staged = f"artifact-{index:02d}.json.zst"
            staged_names.append(staged)
            _write_exclusive(stage_fd, staged, item.compressed, f"staged artifact {item.role}")
        if summary_copy_needed:
            staged = "summary.json"
            staged_names.append(staged)
            _write_exclusive(stage_fd, staged, summary_json.raw, "staged summary")
        staged_manifest = "manifest.json"
        staged_names.append(staged_manifest)
        _write_exclusive(stage_fd, staged_manifest, _manifest_bytes(manifest), "staged manifest")

        for index, (item, (_, destination_name)) in enumerate(
            zip(compressed, destination_names)
        ):
            staged = f"artifact-{index:02d}.json.zst"
            _publish_exclusive(
                stage_fd, staged, root_fd, destination_name, f"artifact {item.role}"
            )
            published.append(destination_name)
        if summary_copy_needed:
            _publish_exclusive(stage_fd, "summary.json", root_fd, summary_name, "summary")
            published.append(summary_name)
        _publish_exclusive(stage_fd, staged_manifest, root_fd, manifest_name, "manifest")
        published.append(manifest_name)

        _remove_staging_directory(root_fd, stage_name, stage_fd, staged_names)
        _close_fd(stage_fd)
        stage_fd = -1
        stage_name = None
    except BaseException as error:
        if opened is not None:
            _remove_published(opened.fd, published)
            if stage_name is not None and stage_fd >= 0:
                try:
                    _remove_staging_directory(
                        opened.fd, stage_name, stage_fd, staged_names
                    )
                except ArtifactPackagingError:
                    pass
                finally:
                    _close_fd(stage_fd)
                stage_fd = -1
            _cleanup_created_directories(opened)
            opened = None
        if isinstance(error, ArtifactPackagingError):
            raise
        if isinstance(error, OSError):
            raise ArtifactPackagingError(f"artifact publication failed: {error}") from error
        raise
    finally:
        if stage_fd >= 0:
            _close_fd(stage_fd)
        if opened is not None:
            _close_opened_directory(opened)
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
