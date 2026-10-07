#!/usr/bin/env python3
"""Archive and later clean the 0944 XLSX evidence root.

The 0944 export contains a complete source export at ``source`` and a smaller,
build-audited closure at ``source-build``.  The complete export is excluded
from the archive only after its receipt and manifest are hash-bound and the
selected closure has been re-audited against its manifest.  The selected
closure is retained byte-for-byte.

Archive, verification, and cleanup are gated by both ``--go`` and
``LITCHI_PERF_0944_ARCHIVE_GO=1``.  ``--plan`` is read-only.  The script never
follows symlinks and refuses to replace an existing artifact.  It does not
run until a separate binary descriptor supplied by the build owner binds the
final target-native binary by relative path, size, and SHA-256.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path
import shlex
import shutil
import subprocess
import sys
from typing import Any, Iterable, Iterator


SOURCE_ROOT = Path("/home/zhuhe/litchi-performance-0944").resolve()
WORK_ROOT = Path("/home/zhuhe/litchi-performance-0944-archive").resolve()
SCRIPT = Path(__file__).resolve()
ARCHIVE = WORK_ROOT / "litchi-performance-0944-v1.tar.zst"
MANIFEST = WORK_ROOT / "litchi-performance-0944-v1.files.jsonl"
MEMBERS = WORK_ROOT / "litchi-performance-0944-v1.members.list"
ARCHIVE_RECEIPT = WORK_ROOT / "litchi-performance-0944-v1.archive-receipt.json"
VERIFY_RECEIPT = WORK_ROOT / "litchi-performance-0944-v1.verify-receipt.json"
CLEANUP_RECEIPT = WORK_ROOT / "litchi-performance-0944-v1.cleanup-receipt.json"
EXTRACT_ROOT = WORK_ROOT / "litchi-performance-0944-v1.extract"
ARCHIVE_STDOUT = WORK_ROOT / "litchi-performance-0944-v1.archive.stdout.log"
ARCHIVE_STDERR = WORK_ROOT / "litchi-performance-0944-v1.archive.stderr.log"
VERIFY_STDOUT = WORK_ROOT / "litchi-performance-0944-v1.verify.stdout.log"
VERIFY_STDERR = WORK_ROOT / "litchi-performance-0944-v1.verify.stderr.log"
BINARY_DESCRIPTOR = WORK_ROOT / "binary-descriptor.json"
TAR = Path("/usr/bin/tar").resolve()
GO_ENV = "LITCHI_PERF_0944_ARCHIVE_GO"
GO_VALUE = "1"
ROOT_NAME = SOURCE_ROOT.name
PYCACHE_NAME = "__pycache__"
KNOWN_TARGET_CACHE_DIRS = frozenset({Path("target-native")})
REDUNDANT_TOP_LEVEL_DIRS = frozenset({Path("source")})
BOUND_TARGET_FILES: frozenset[Path] = frozenset()
TAR_ENVIRONMENT = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}


class ArchiveError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ArchiveError(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def sha256_path(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_descriptor(path: Path) -> dict[str, Any]:
    path = Path(path)
    require(not path.is_symlink(), f"symlink is not a custody file: {path}")
    path = path.resolve()
    require(path.is_file(), f"expected regular file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256_path(path)}


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing regular JSON file: {path}")
    try:
        with path.open("r", encoding="utf-8") as stream:
            return json.load(stream)
    except json.JSONDecodeError as error:
        raise ArchiveError(f"invalid JSON: {path}: {error}") from error


def write_json_exclusive(path: Path, value: Any) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    try:
        with path.open("xb") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise ArchiveError(f"refusing to replace existing artifact: {path}") from error


def write_bytes_exclusive(path: Path, data: bytes) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise ArchiveError(f"refusing to replace existing artifact: {path}") from error


def relative_to_source(path: Path) -> Path:
    path = path.resolve()
    try:
        return path.relative_to(SOURCE_ROOT)
    except ValueError as error:
        raise ArchiveError(f"path escaped source root: {path}") from error


def is_redundant_source(relative: Path) -> bool:
    return any(relative.parts[: len(root.parts)] == root.parts for root in REDUNDANT_TOP_LEVEL_DIRS)


def is_target_path(relative: Path) -> bool:
    return any(relative.parts[: len(cache.parts)] == cache.parts for cache in KNOWN_TARGET_CACHE_DIRS)


def has_bound_descendant(relative: Path) -> bool:
    return any(bound.parts[: len(relative.parts)] == relative.parts for bound in BOUND_TARGET_FILES)


def may_descend(relative_directory: Path) -> bool:
    if not relative_directory.parts:
        return True
    if PYCACHE_NAME in relative_directory.parts or is_redundant_source(relative_directory):
        return False
    if not is_target_path(relative_directory):
        return True
    return has_bound_descendant(relative_directory)


def selected_file(relative: Path) -> bool:
    if PYCACHE_NAME in relative.parts or is_redundant_source(relative):
        return False
    if is_target_path(relative):
        return relative in BOUND_TARGET_FILES
    return True


def iter_selected_files() -> Iterator[tuple[Path, Path]]:
    require(SOURCE_ROOT.is_dir() and not SOURCE_ROOT.is_symlink(), f"missing source root: {SOURCE_ROOT}")
    for current, directories, filenames in os.walk(SOURCE_ROOT, topdown=True, followlinks=False):
        current_path = Path(current)
        current_relative = relative_to_source(current_path)
        kept_directories: list[str] = []
        for name in sorted(directories):
            child = current_path / name
            child_relative = current_relative / name
            if child.is_symlink():
                raise ArchiveError(f"symlink encountered under source root: {child}")
            if may_descend(child_relative):
                kept_directories.append(name)
        directories[:] = kept_directories
        for name in sorted(filenames):
            path = current_path / name
            relative = current_relative / name
            if not selected_file(relative):
                continue
            require(not path.is_symlink(), f"symlink encountered under source root: {path}")
            require(path.is_file(), f"non-regular selected artifact: {path}")
            yield relative, path


def validate_relative(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label} path missing")
    relative = Path(value)
    require(not relative.is_absolute() and ".." not in relative.parts, f"{label} path escapes root: {value}")
    require(not any(part == "" for part in relative.parts), f"{label} path contains an empty component: {value}")
    return relative


def sha_sidecar(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing SHA sidecar: {path}")
    fields = path.read_text(encoding="utf-8").strip().split()
    require(fields and len(fields[0]) == 64 and all(c in "0123456789abcdef" for c in fields[0]),
            f"invalid SHA sidecar: {path}")
    return fields[0]


def manifest_scalar_prefix(path: Path, key: str) -> Any:
    """Read a scalar before the huge ``files`` array without loading it."""
    require(path.is_file() and not path.is_symlink(), f"missing manifest: {path}")
    prefix = path.open("rb").read(8192).decode("utf-8", errors="strict")
    marker = json.dumps(key) + ":"
    start = prefix.find(marker)
    require(start >= 0, f"manifest scalar missing: {path}: {key}")
    tail = prefix[start + len(marker):].lstrip()
    decoder = json.JSONDecoder()
    try:
        value, _ = decoder.raw_decode(tail)
    except json.JSONDecodeError as error:
        raise ArchiveError(f"manifest scalar is malformed: {path}: {key}") from error
    require(isinstance(value, (str, int, float, bool)) or value is None,
            f"manifest scalar is not scalar: {path}: {key}")
    return value


def manifest_records(value: Any, label: str) -> dict[str, dict[str, Any]]:
    require(isinstance(value, dict), f"{label} manifest is not an object")
    entries = value.get("files")
    require(isinstance(entries, list), f"{label} manifest files is not a list")
    result: dict[str, dict[str, Any]] = {}
    for index, item in enumerate(entries):
        require(isinstance(item, dict), f"{label} manifest record {index} is not an object")
        relative = validate_relative(item.get("path"), f"{label} manifest record {index}")
        key = relative.as_posix()
        require(key not in result, f"duplicate {label} manifest path: {key}")
        size = item.get("bytes")
        digest = item.get("sha256")
        require(type(size) is int and size >= 0, f"invalid {label} manifest size: {key}")
        require(isinstance(digest, str) and len(digest) == 64 and all(c in "0123456789abcdef" for c in digest),
                f"invalid {label} manifest digest: {key}")
        result[key] = {"path": key, "bytes": size, "sha256": digest}
    require(value.get("file_count") == len(result), f"{label} manifest file count mismatch")
    return result


def selected_source_audit() -> dict[str, Any]:
    """Bind every source-build file to the selected closure manifest."""
    selected_manifest_path = SOURCE_ROOT / "source-build-manifest.json"
    selected = read_json(selected_manifest_path)
    require(selected.get("schema") == "litchi-xlsx-profile-source-closure-v1", "selected source manifest schema changed")
    require(selected.get("source_root") == "source-build", "selected source manifest root changed")
    full_manifest = SOURCE_ROOT / "source-manifest.json"
    full_sha = sha256_path(full_manifest)
    require(selected.get("full_export_manifest") == "source-manifest.json", "selected manifest full-export binding changed")
    require(selected.get("full_export_manifest_sha256") == full_sha, "selected manifest full-export SHA changed")
    selected_manifest_sha = sha256_path(selected_manifest_path)
    require(sha_sidecar(SOURCE_ROOT / "source-build-manifest.sha256") == selected_manifest_sha,
            "selected source manifest sidecar changed")
    expected = manifest_records(selected, "selected source")
    actual: dict[str, dict[str, Any]] = {}
    selected_root = SOURCE_ROOT / "source-build"
    require(selected_root.is_dir() and not selected_root.is_symlink(), f"missing selected source root: {selected_root}")
    for current, directories, filenames in os.walk(selected_root, topdown=True, followlinks=False):
        current_path = Path(current)
        for name in sorted(directories):
            require(not (current_path / name).is_symlink(), f"selected source contains symlink: {current_path / name}")
        for name in sorted(filenames):
            path = current_path / name
            require(not path.is_symlink() and path.is_file(), f"selected source contains non-regular file: {path}")
            relative = path.relative_to(selected_root).as_posix()
            actual[relative] = {"path": relative, "bytes": path.stat().st_size, "sha256": sha256_path(path)}
    require(actual == expected, "selected source-build files differ from its bound manifest")
    selection = read_json(SOURCE_ROOT / "source-closure-selection.json")
    require(selection.get("schema") == "litchi-xlsx-profile-source-closure-selection-v1", "closure selection schema changed")
    require(selection.get("full_export_root") == "source" and selection.get("selected_root") == "source-build",
            "closure selection roots changed")
    require(selection.get("origin_commit") == selected.get("origin_commit"),
            "closure selection origin commit changed")
    if "origin_tree" in selection:
        require(selection.get("origin_tree") == selected.get("origin_tree"),
                "closure selection origin tree changed")
    return {
        "selected_manifest": file_descriptor(selected_manifest_path),
        "selected_manifest_sha_sidecar": file_descriptor(SOURCE_ROOT / "source-build-manifest.sha256"),
        "selected_manifest_sha256": selected_manifest_sha,
        "selected_files": len(actual),
        "selected_bytes": sum(item["bytes"] for item in actual.values()),
        "full_export_manifest_sha256": full_sha,
        "selection": file_descriptor(SOURCE_ROOT / "source-closure-selection.json"),
    }


def full_source_binding() -> dict[str, Any]:
    receipt_path = SOURCE_ROOT / "export-receipt.json"
    manifest_path = SOURCE_ROOT / "source-manifest.json"
    receipt = read_json(receipt_path)
    require(receipt.get("schema") == "litchi-xlsx-profile-source-export-v1", "full export receipt schema changed")
    require(receipt.get("source_root") == "source" and receipt.get("manifest") == "source-manifest.json",
            "full export receipt source binding changed")
    manifest_sha = sha256_path(manifest_path)
    require(receipt.get("manifest_sha256") == manifest_sha, "full export receipt manifest SHA changed")
    require(sha_sidecar(SOURCE_ROOT / "source-manifest.sha256") == manifest_sha, "full source manifest sidecar changed")
    require(manifest_scalar_prefix(manifest_path, "schema") == "litchi-source-manifest-v1", "full source manifest schema changed")
    require(manifest_scalar_prefix(manifest_path, "source_root") == "source", "full source manifest root changed")
    require(manifest_scalar_prefix(manifest_path, "file_count") == receipt.get("file_count"), "full source manifest count changed")
    require(manifest_scalar_prefix(manifest_path, "commit") == receipt.get("commit"), "full source manifest commit changed")
    require(manifest_scalar_prefix(manifest_path, "tree") == receipt.get("tree"), "full source manifest tree changed")
    return {
        "receipt": file_descriptor(receipt_path),
        "manifest": file_descriptor(manifest_path),
        "manifest_sha_sidecar": file_descriptor(SOURCE_ROOT / "source-manifest.sha256"),
        "file_count": receipt.get("file_count"),
        "commit": receipt.get("commit"),
        "tree": receipt.get("tree"),
    }


def validate_source_closure() -> dict[str, Any]:
    full = full_source_binding()
    selected = selected_source_audit()
    return {"full_export": full, "selected_closure": selected}


def load_binary_descriptor() -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    descriptor = read_json(BINARY_DESCRIPTOR)
    require(descriptor.get("schema") == "litchi.performance.0944.binary-descriptor-v1",
            "binary descriptor schema changed")
    require(descriptor.get("source_root") == str(SOURCE_ROOT), "binary descriptor source root changed")
    entries = descriptor.get("binaries")
    require(isinstance(entries, list) and entries, "binary descriptor has no binaries")
    actual: dict[str, dict[str, Any]] = {}
    for index, item in enumerate(entries):
        require(isinstance(item, dict), f"binary descriptor entry {index} is not an object")
        relative = validate_relative(item.get("path"), f"binary descriptor entry {index}")
        require(is_target_path(relative), f"bound binary is outside target-native: {relative}")
        require(relative not in actual, f"duplicate bound binary: {relative}")
        expected_size = item.get("bytes")
        expected_sha = item.get("sha256")
        require(type(expected_size) is int and expected_size > 0, f"invalid bound binary size: {relative}")
        require(isinstance(expected_sha, str) and len(expected_sha) == 64 and all(c in "0123456789abcdef" for c in expected_sha),
                f"invalid bound binary SHA: {relative}")
        path = SOURCE_ROOT / relative
        observed = file_descriptor(path)
        require(observed["bytes"] == expected_size, f"bound binary size changed: {relative}")
        require(observed["sha256"] == expected_sha, f"bound binary SHA changed: {relative}")
        actual[relative.as_posix()] = observed
    return descriptor, actual


def expected_binary_descriptors() -> dict[str, dict[str, Any]]:
    global BOUND_TARGET_FILES
    descriptor, actual = load_binary_descriptor()
    BOUND_TARGET_FILES = frozenset(Path(path) for path in actual)
    return {"descriptor": file_descriptor(BINARY_DESCRIPTOR), "binaries": actual, "build_receipt": descriptor.get("build_receipt")}


def collect_records() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    closure = validate_source_closure()
    binaries = expected_binary_descriptors()
    records: list[dict[str, Any]] = []
    seen: set[str] = set()
    for relative, path in iter_selected_files():
        key = relative.as_posix()
        require(key not in seen, f"duplicate selected path: {key}")
        seen.add(key)
        records.append({"path": key, "bytes": path.stat().st_size, "sha256": sha256_path(path)})
    records.sort(key=lambda item: item["path"])
    header = {
        "schema": "litchi.performance.0944.archive-file-manifest.v1",
        "source_root": str(SOURCE_ROOT),
        "root_name": ROOT_NAME,
        "generated_utc": utc_now(),
        "files": len(records),
        "bytes": sum(item["bytes"] for item in records),
        "source_closure": closure,
        "bound_binaries": binaries,
        "excluded": {
            "redundant_top_level_directories": sorted(path.as_posix() for path in REDUNDANT_TOP_LEVEL_DIRS),
            "directory_name": PYCACHE_NAME,
            "target_cache_policy": "exclude only target-native contents outside descriptor-bound binaries; preserve source-build and nested target-* names",
        },
    }
    return header, records


def write_manifest(header: dict[str, Any], records: list[dict[str, Any]]) -> None:
    try:
        with MANIFEST.open("x", encoding="utf-8") as stream:
            stream.write(json.dumps(header, sort_keys=True) + "\n")
            for record in records:
                stream.write(json.dumps(record, sort_keys=True) + "\n")
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise ArchiveError(f"refusing to replace existing artifact: {MANIFEST}") from error


def write_member_list(records: list[dict[str, Any]]) -> None:
    encoded = b"".join((ROOT_NAME + "/" + record["path"]).encode("utf-8") + b"\0" for record in records)
    write_bytes_exclusive(MEMBERS, encoded)


def read_manifest() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    require(MANIFEST.is_file() and not MANIFEST.is_symlink(), f"missing archive manifest: {MANIFEST}")
    records: list[dict[str, Any]] = []
    with MANIFEST.open("r", encoding="utf-8") as stream:
        first = stream.readline()
        require(first, "archive manifest is empty")
        header = json.loads(first)
        require(isinstance(header, dict) and header.get("schema") == "litchi.performance.0944.archive-file-manifest.v1",
                "archive manifest schema changed")
        require(header.get("source_root") == str(SOURCE_ROOT) and header.get("root_name") == ROOT_NAME,
                "archive manifest source binding changed")
        previous = ""
        for line_number, line in enumerate(stream, start=2):
            record = json.loads(line)
            require(isinstance(record, dict), f"archive manifest line {line_number} is not an object")
            relative = validate_relative(record.get("path"), f"archive manifest line {line_number}")
            key = relative.as_posix()
            require(key > previous, f"archive manifest paths are not strictly sorted at line {line_number}")
            previous = key
            require(not is_redundant_source(relative), f"archive manifest includes redundant source: {key}")
            require(not PYCACHE_NAME in relative.parts, f"archive manifest includes __pycache__: {key}")
            if is_target_path(relative):
                require(relative in BOUND_TARGET_FILES, f"archive manifest includes unbound target file: {key}")
            size = record.get("bytes")
            digest = record.get("sha256")
            require(type(size) is int and size >= 0, f"archive manifest size invalid at line {line_number}")
            require(isinstance(digest, str) and len(digest) == 64 and all(c in "0123456789abcdef" for c in digest),
                    f"archive manifest digest invalid at line {line_number}")
            records.append({"path": key, "bytes": size, "sha256": digest})
    require(header.get("files") == len(records), "archive manifest file count mismatch")
    require(header.get("bytes") == sum(item["bytes"] for item in records), "archive manifest byte total mismatch")
    return header, records


def tar_archive_command() -> list[str]:
    return [str(TAR), "--create", "--zstd", "--format=posix", "--numeric-owner", "--owner=0", "--group=0",
            "--mtime=@0", "--no-recursion", "--null", "--verbatim-files-from", "--directory", str(SOURCE_ROOT.parent),
            "--files-from", str(MEMBERS), "--file", str(ARCHIVE)]


def tar_extract_command() -> list[str]:
    return [str(TAR), "--zstd", "--extract", "--file", str(ARCHIVE), "--directory", str(EXTRACT_ROOT),
            "--no-same-owner", "--no-same-permissions"]


def run_logged(command: list[str], stdout_path: Path, stderr_path: Path, *, cwd: Path) -> int:
    stdout_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            completed = subprocess.run(command, cwd=str(cwd), env=dict(TAR_ENVIRONMENT), stdin=subprocess.DEVNULL,
                                       stdout=stdout, stderr=stderr, check=False)
    except FileExistsError as error:
        raise ArchiveError(f"refusing to replace existing command log: {error}") from error
    return completed.returncode


def script_command(*actions: str) -> str:
    return f"{GO_ENV}=1 /usr/bin/python3 -B {shlex.quote(str(SCRIPT))} --go " + " ".join(actions)


_GO_FLAG = False


def require_go() -> None:
    require(os.environ.get(GO_ENV) == GO_VALUE and _GO_FLAG, f"archive/cleanup is gated; require --go and {GO_ENV}=1")


def require_absent(paths: Iterable[Path]) -> None:
    for path in paths:
        require(not path.exists() and not path.is_symlink(), f"refusing to replace existing artifact: {path}")


def archive_mode() -> dict[str, Any]:
    require_go()
    require(TAR.is_file() and os.access(TAR, os.X_OK), f"tar with zstd support is unavailable: {TAR}")
    require_absent((ARCHIVE, MANIFEST, MEMBERS, ARCHIVE_RECEIPT, VERIFY_RECEIPT, CLEANUP_RECEIPT, EXTRACT_ROOT,
                    ARCHIVE_STDOUT, ARCHIVE_STDERR, VERIFY_STDOUT, VERIFY_STDERR))
    header, records = collect_records()
    write_manifest(header, records)
    write_member_list(records)
    command = tar_archive_command()
    return_code = run_logged(command, ARCHIVE_STDOUT, ARCHIVE_STDERR, cwd=SOURCE_ROOT.parent)
    require(return_code == 0, f"tar archive command failed with exit code {return_code}")
    require(ARCHIVE.is_file() and ARCHIVE.stat().st_size > 0, "tar archive was not produced")
    receipt = {
        "schema": "litchi.performance.0944.archive-receipt.v1", "status": "archived", "created_utc": utc_now(),
        "script": file_descriptor(SCRIPT), "source_root": str(SOURCE_ROOT), "manifest": file_descriptor(MANIFEST),
        "member_list": file_descriptor(MEMBERS), "archive": file_descriptor(ARCHIVE),
        "archive_stdout": file_descriptor(ARCHIVE_STDOUT), "archive_stderr": file_descriptor(ARCHIVE_STDERR),
        "files": header["files"], "bytes": header["bytes"], "source_closure": header["source_closure"],
        "bound_binaries": header["bound_binaries"], "tar_command": command, "tar_command_shell": shlex.join(command),
        "regeneration_commands": {"archive": script_command("--archive"), "verify": script_command("--verify"),
                                   "cleanup_after_verification": script_command("--cleanup")},
        "excluded_rules": header["excluded"],
        "cleanup_policy": {"preserve_bound_binaries": True, "remove_target_native_cache_components": True,
                            "remove_pycache_directories": True, "remove_redundant_top_level_source": True,
                            "remove_verification_extraction": str(EXTRACT_ROOT)},
    }
    write_json_exclusive(ARCHIVE_RECEIPT, receipt)
    return {"status": receipt["status"], "files": header["files"], "bytes": header["bytes"], "archive": receipt["archive"],
            "archive_receipt": str(ARCHIVE_RECEIPT)}


def extracted_file_records() -> list[dict[str, Any]]:
    root = EXTRACT_ROOT / ROOT_NAME
    require(root.is_dir() and not root.is_symlink(), f"archive extraction root missing: {root}")
    actual: list[dict[str, Any]] = []
    for current, directories, filenames in os.walk(root, topdown=True, followlinks=False):
        current_path = Path(current)
        for name in directories:
            require(not (current_path / name).is_symlink(), f"symlink in archive extraction: {current_path / name}")
        for name in sorted(filenames):
            path = current_path / name
            require(not path.is_symlink() and path.is_file(), f"non-regular archive extraction member: {path}")
            relative = path.relative_to(root)
            require(not is_redundant_source(relative), f"archive extraction includes redundant source: {relative}")
            actual.append({"path": ROOT_NAME + "/" + relative.as_posix(), "bytes": path.stat().st_size, "sha256": sha256_path(path)})
    actual.sort(key=lambda item: item["path"])
    return actual


def load_receipt(path: Path, schema: str, status: str) -> dict[str, Any]:
    value = read_json(path)
    require(isinstance(value, dict) and value.get("schema") == schema and value.get("status") == status,
            f"receipt is not {status}: {path}")
    return value


def verify_mode() -> dict[str, Any]:
    require_go()
    archive_receipt = load_receipt(ARCHIVE_RECEIPT, "litchi.performance.0944.archive-receipt.v1", "archived")
    closure = validate_source_closure()
    binaries = expected_binary_descriptors()
    header, records = read_manifest()
    require(archive_receipt.get("source_closure") == closure, "source closure binding differs from archive receipt")
    require(archive_receipt.get("bound_binaries") == binaries, "binary custody differs from archive receipt")
    require(file_descriptor(ARCHIVE) == archive_receipt["archive"], "archive changed since archive receipt")
    require(file_descriptor(MANIFEST) == archive_receipt["manifest"], "manifest changed since archive receipt")
    require_absent((EXTRACT_ROOT, VERIFY_RECEIPT, VERIFY_STDOUT, VERIFY_STDERR))
    EXTRACT_ROOT.mkdir(parents=True, exist_ok=False)
    command = tar_extract_command()
    return_code = run_logged(command, VERIFY_STDOUT, VERIFY_STDERR, cwd=SOURCE_ROOT.parent)
    require(return_code == 0, f"tar extraction command failed with exit code {return_code}")
    actual = extracted_file_records()
    expected = [{"path": ROOT_NAME + "/" + item["path"], "bytes": item["bytes"], "sha256": item["sha256"]} for item in records]
    require(actual == expected, "archive extraction file set, size, or digest differs from manifest")
    require(len(actual) == header["files"] and sum(item["bytes"] for item in actual) == header["bytes"],
            "archive extraction totals differ from manifest")
    receipt = {
        "schema": "litchi.performance.0944.archive-verify-receipt.v1", "status": "passed", "verified_utc": utc_now(),
        "script": file_descriptor(SCRIPT), "archive": file_descriptor(ARCHIVE),
        "archive_receipt": file_descriptor(ARCHIVE_RECEIPT), "manifest": file_descriptor(MANIFEST),
        "source_closure": closure, "bound_binaries": binaries, "extraction_root": str(EXTRACT_ROOT),
        "files": header["files"], "bytes": header["bytes"], "tar_command": command, "tar_command_shell": shlex.join(command),
        "verify_stdout": file_descriptor(VERIFY_STDOUT), "verify_stderr": file_descriptor(VERIFY_STDERR),
        "every_file_rehashed": True, "file_set_exact": True, "cleanup_authorized_by_this_receipt": False,
    }
    write_json_exclusive(VERIFY_RECEIPT, receipt)
    return {"status": receipt["status"], "files": receipt["files"], "bytes": receipt["bytes"], "verify_receipt": str(VERIFY_RECEIPT)}


def current_matches_archive(records: list[dict[str, Any]]) -> None:
    _header, current = collect_records()
    require(current == records, "selected evidence changed after archive verification")


def remove_path(path: Path) -> None:
    require(SOURCE_ROOT in path.parents, f"refusing removal outside source root: {path}")
    if path.is_symlink() or path.is_file():
        path.unlink()
    elif path.is_dir():
        shutil.rmtree(path)
    else:
        raise ArchiveError(f"cannot remove unknown path type: {path}")


def cleanup_target_native() -> list[str]:
    removed: list[str] = []
    has_allowlisted_descendant = lambda relative: any(bound.parts[: len(relative.parts)] == relative.parts for bound in BOUND_TARGET_FILES)
    for cache_relative in sorted(KNOWN_TARGET_CACHE_DIRS, key=Path.as_posix):
        cache_root = SOURCE_ROOT / cache_relative
        if not cache_root.exists() and not cache_root.is_symlink():
            continue
        require(cache_root.is_dir() and not cache_root.is_symlink(), f"known target cache is not a directory: {cache_root}")
        for current, directories, filenames in os.walk(cache_root, topdown=False, followlinks=False):
            current_path = Path(current)
            for name in sorted(filenames):
                path = current_path / name
                relative = path.relative_to(SOURCE_ROOT)
                if relative in BOUND_TARGET_FILES:
                    continue
                remove_path(path)
                removed.append(relative.as_posix())
            for name in sorted(directories):
                path = current_path / name
                relative = path.relative_to(SOURCE_ROOT)
                if not path.exists() and not path.is_symlink():
                    continue
                require(not path.is_symlink(), f"symlink target cache refused: {path}")
                if has_allowlisted_descendant(relative):
                    try:
                        path.rmdir()
                        removed.append(relative.as_posix())
                    except OSError:
                        pass
                else:
                    remove_path(path)
                    removed.append(relative.as_posix())
        if not has_allowlisted_descendant(cache_relative):
            cache_root.rmdir()
            removed.append(cache_relative.as_posix())
    return removed


def cleanup_pycache() -> list[str]:
    removed: list[str] = []
    for current, directories, _ in os.walk(SOURCE_ROOT, topdown=False, followlinks=False):
        current_path = Path(current)
        for name in sorted(directories):
            if name != PYCACHE_NAME:
                continue
            path = current_path / name
            require(not path.is_symlink(), f"symlink __pycache__ refused: {path}")
            if path.exists():
                relative = path.relative_to(SOURCE_ROOT)
                remove_path(path)
                removed.append(relative.as_posix())
    return removed


def cleanup_redundant_source() -> str | None:
    path = SOURCE_ROOT / "source"
    if not path.exists() and not path.is_symlink():
        return None
    require(path.is_dir() and not path.is_symlink(), f"redundant source export is not a directory: {path}")
    remove_path(path)
    return path.relative_to(SOURCE_ROOT).as_posix()


def assert_post_cleanup() -> dict[str, Any]:
    require(not (SOURCE_ROOT / "source").exists() and not (SOURCE_ROOT / "source").is_symlink(),
            "redundant top-level source remains after cleanup")
    for current, directories, filenames in os.walk(SOURCE_ROOT, topdown=True, followlinks=False):
        current_path = Path(current)
        for name in directories:
            path = current_path / name
            relative = path.relative_to(SOURCE_ROOT)
            require(name != PYCACHE_NAME, f"__pycache__ remains after cleanup: {path}")
            if is_target_path(relative):
                require(has_bound_descendant(relative), f"unallowlisted target-native directory remains: {path}")
                require(not path.is_symlink(), f"symlink remains after cleanup: {path}")
        for name in filenames:
            path = current_path / name
            relative = path.relative_to(SOURCE_ROOT)
            require(PYCACHE_NAME not in relative.parts, f"__pycache__ file remains: {path}")
            if is_target_path(relative):
                require(relative in BOUND_TARGET_FILES, f"unallowlisted target-native file remains: {relative}")
    closure = selected_source_audit()
    binaries = expected_binary_descriptors()
    return {"selected_source": closure, "bound_binaries": binaries}


def cleanup_mode() -> dict[str, Any]:
    require_go()
    archive_receipt = load_receipt(ARCHIVE_RECEIPT, "litchi.performance.0944.archive-receipt.v1", "archived")
    verify_receipt = load_receipt(VERIFY_RECEIPT, "litchi.performance.0944.archive-verify-receipt.v1", "passed")
    closure = validate_source_closure()
    binaries = expected_binary_descriptors()
    header, records = read_manifest()
    require(archive_receipt.get("source_closure") == closure and archive_receipt.get("bound_binaries") == binaries,
            "custody bindings changed before cleanup")
    require(file_descriptor(ARCHIVE) == archive_receipt["archive"], "archive changed before cleanup")
    require(file_descriptor(MANIFEST) == archive_receipt["manifest"], "manifest changed before cleanup")
    require(verify_receipt["archive"] == archive_receipt["archive"], "verify receipt archive binding changed")
    require(EXTRACT_ROOT.is_dir() and not EXTRACT_ROOT.is_symlink(), "verified extraction is missing")
    current_matches_archive(records)
    removed_targets = cleanup_target_native()
    removed_pycache = cleanup_pycache()
    removed_source = cleanup_redundant_source()
    shutil.rmtree(EXTRACT_ROOT)
    preserved = assert_post_cleanup()
    receipt = {
        "schema": "litchi.performance.0944.archive-cleanup-receipt.v1", "status": "completed", "cleaned_utc": utc_now(),
        "script": file_descriptor(SCRIPT), "source_root": str(SOURCE_ROOT), "archive": archive_receipt["archive"],
        "manifest": archive_receipt["manifest"], "verify_receipt": file_descriptor(VERIFY_RECEIPT),
        "files": header["files"], "bytes": header["bytes"], "removed_target_native_paths": removed_targets,
        "removed_pycache_paths": removed_pycache, "removed_redundant_top_level_source": removed_source,
        "removed_verification_extraction": str(EXTRACT_ROOT), "preserved_source_build": preserved["selected_source"],
        "preserved_bound_binaries": preserved["bound_binaries"], "checkout_touched": False,
        "other_performance_roots_touched": [],
    }
    write_json_exclusive(CLEANUP_RECEIPT, receipt)
    return {"status": receipt["status"], "removed_target_paths": len(removed_targets),
            "removed_pycache_paths": removed_pycache, "removed_source": removed_source,
            "cleanup_receipt": str(CLEANUP_RECEIPT)}


def plan_mode() -> dict[str, Any]:
    closure: dict[str, Any] | None = None
    closure_error: str | None = None
    try:
        closure = validate_source_closure()
    except (ArchiveError, OSError, json.JSONDecodeError) as error:
        closure_error = f"{type(error).__name__}: {error}"
    binary: dict[str, Any] | None = None
    binary_error: str | None = None
    try:
        binary = expected_binary_descriptors()
    except (ArchiveError, OSError, json.JSONDecodeError) as error:
        binary_error = f"{type(error).__name__}: {error}"
    return {
        "schema": "litchi.performance.0944.archive-plan.v1", "status": "prepared_not_started", "script": str(SCRIPT),
        "source_root": str(SOURCE_ROOT), "workspace": str(WORK_ROOT), "archive": str(ARCHIVE), "manifest": str(MANIFEST),
        "member_list": str(MEMBERS), "verification_extraction": str(EXTRACT_ROOT),
        "include": "all regular files under 0944 root except validated redundant top-level source, __pycache__, and target-native cache files outside descriptor-bound binaries",
        "exclude_redundant_source": {"path": "source", "requires": "full export receipt+manifest hash and selected source-build manifest audit"},
        "target_native_policy": "only explicit top-level target-native is a cache; preserve descriptor-bound release binary and remove other target-native contents after verification",
        "source_closure": closure if closure is not None else {"status": "pending", "error": closure_error},
        "binary_descriptor": binary if binary is not None else {"status": "pending", "path": str(BINARY_DESCRIPTOR), "error": binary_error},
        "go_gate": {"flag": "--go", "environment": f"{GO_ENV}=1"},
        "exact_regeneration_commands": {"archive": script_command("--archive"), "verify": script_command("--verify"), "cleanup": script_command("--cleanup")},
    }


def main(argv: list[str] | None = None) -> int:
    global _GO_FLAG
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--plan", action="store_true")
    modes.add_argument("--archive", action="store_true")
    modes.add_argument("--verify", action="store_true")
    modes.add_argument("--cleanup", action="store_true")
    parser.add_argument("--go", action="store_true", help="required with archive, verify, and cleanup")
    args = parser.parse_args(argv)
    _GO_FLAG = bool(args.go)
    try:
        if args.plan:
            result = plan_mode()
        elif args.archive:
            result = archive_mode()
        elif args.verify:
            result = verify_mode()
        else:
            result = cleanup_mode()
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (ArchiveError, OSError, json.JSONDecodeError, subprocess.SubprocessError) as error:
        print(f"0944 archive operation refused: {type(error).__name__}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
