#!/usr/bin/env python3
"""Archive and later clean the 0942 performance evidence root.

The archive and cleanup operations are deliberately gated twice: the caller
must pass ``--go`` and set ``LITCHI_PERF_0942_ARCHIVE_GO=1``.  ``--plan`` is
the only read-only mode.  The script never follows symlinks, never reads or
removes anything outside the 0942 root and its external archive workspace,
and refuses to replace an existing artifact.
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


SOURCE_ROOT = Path("/home/zhuhe/litchi-performance-0942").resolve()
WORK_ROOT = Path("/home/zhuhe/litchi-performance-0942-archive").resolve()
SCRIPT = Path(__file__).resolve()
ARCHIVE = WORK_ROOT / "litchi-performance-0942-v2.tar.zst"
MANIFEST = WORK_ROOT / "litchi-performance-0942-v2.files.jsonl"
MEMBERS = WORK_ROOT / "litchi-performance-0942-v2.members.list"
ARCHIVE_RECEIPT = WORK_ROOT / "litchi-performance-0942-v2.archive-receipt.json"
VERIFY_RECEIPT = WORK_ROOT / "litchi-performance-0942-v2.verify-receipt.json"
CLEANUP_RECEIPT = WORK_ROOT / "litchi-performance-0942-v2.cleanup-receipt.json"
EXTRACT_ROOT = WORK_ROOT / "litchi-performance-0942-v2.extract"
ARCHIVE_STDOUT = WORK_ROOT / "litchi-performance-0942-v2.archive.stdout.log"
ARCHIVE_STDERR = WORK_ROOT / "litchi-performance-0942-v2.archive.stderr.log"
VERIFY_STDOUT = WORK_ROOT / "litchi-performance-0942-v2.verify.stdout.log"
VERIFY_STDERR = WORK_ROOT / "litchi-performance-0942-v2.verify.stderr.log"
TAR = Path("/usr/bin/tar").resolve()
GO_ENV = "LITCHI_PERF_0942_ARCHIVE_GO"
GO_VALUE = "1"
ROOT_NAME = SOURCE_ROOT.name

ALLOWLISTED_TARGET_FILES = {
    Path("target-profile-v1/release/doc_retention_probe_0942_cfb_profile"),
    Path("target-profile-v2/release/doc_retention_probe_0942_cfb_profile"),
}
EXPECTED_TARGET_BINARIES = {
    Path("target-profile-v1/release/doc_retention_probe_0942_cfb_profile"): {
        "bytes": 125_803_872,
        "sha256": "838326af3130bc2aee6cbf1e9c68054e04321dec468274d173f431cceefe2530",
    },
    Path("target-profile-v2/release/doc_retention_probe_0942_cfb_profile"): {
        "bytes": 126_326_672,
        "sha256": "92f85ce6019405fa1b0d97ec796947cd7ec240d2b8143f9cf06356d97409e70e",
    },
}
PYCACHE_NAME = "__pycache__"
KNOWN_TARGET_CACHE_DIRS = frozenset({
    Path("target-profile-v1"),
    Path("target-profile-v2"),
})
TAR_ENVIRONMENT = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
}


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


def is_target_path(relative: Path) -> bool:
    """Return whether *relative* is inside a known top-level profile cache.

    Nested source files may legitimately use ``target-*`` names.  Filtering
    is therefore limited to the two explicitly bound top-level cache dirs.
    """
    return any(
        relative.parts[: len(cache.parts)] == cache.parts
        for cache in KNOWN_TARGET_CACHE_DIRS
    )


def may_descend(relative_directory: Path) -> bool:
    if not relative_directory.parts:
        return True
    if PYCACHE_NAME in relative_directory.parts:
        return False
    if not is_target_path(relative_directory):
        return True
    parts = relative_directory.parts
    return any(parts == allowed.parts[: len(parts)] for allowed in ALLOWLISTED_TARGET_FILES)


def selected_file(relative: Path) -> bool:
    if any(part == "__pycache__" for part in relative.parts):
        return False
    if is_target_path(relative):
        return relative in ALLOWLISTED_TARGET_FILES
    return True


def iter_selected_files() -> Iterator[tuple[Path, Path]]:
    require(SOURCE_ROOT.is_dir() and not SOURCE_ROOT.is_symlink(), f"missing source root: {SOURCE_ROOT}")
    for current, directories, filenames in os.walk(SOURCE_ROOT, topdown=True, followlinks=False):
        current_path = Path(current).resolve()
        current_relative = relative_to_source(current_path)
        kept_directories: list[str] = []
        for name in sorted(directories):
            child = current_path / name
            child_relative = current_relative / name
            if child.is_symlink():
                raise ArchiveError(f"symlink encountered under source root: {child}")
            if name == "__pycache__":
                continue
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


def expected_binary_descriptors() -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for relative, expected in sorted(EXPECTED_TARGET_BINARIES.items(), key=lambda item: item[0].as_posix()):
        path = SOURCE_ROOT / relative
        actual = file_descriptor(path)
        require(actual["bytes"] == expected["bytes"], f"bound profile binary size changed: {relative}")
        require(actual["sha256"] == expected["sha256"], f"bound profile binary hash changed: {relative}")
        result[relative.as_posix()] = actual
    return result


def collect_records() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    expected_binary_descriptors()
    records: list[dict[str, Any]] = []
    seen: set[str] = set()
    for relative, path in iter_selected_files():
        key = relative.as_posix()
        require(key not in seen, f"duplicate selected path: {key}")
        seen.add(key)
        records.append({
            "path": key,
            "bytes": path.stat().st_size,
            "sha256": sha256_path(path),
        })
    records.sort(key=lambda item: item["path"])
    header = {
        "schema": "litchi.performance.0942.archive-file-manifest.v2",
        "source_root": str(SOURCE_ROOT),
        "root_name": ROOT_NAME,
        "generated_utc": utc_now(),
        "files": len(records),
        "bytes": sum(item["bytes"] for item in records),
        "excluded": {
            "directory_name": "__pycache__",
            "target_cache_policy": "exclude only the known top-level target-profile-v1 and target-profile-v2 directories except the two bound profile binaries; preserve nested target-* names elsewhere",
            "target_binary_allowlist": sorted(item.as_posix() for item in ALLOWLISTED_TARGET_FILES),
        },
    }
    return header, records


def write_manifest(header: dict[str, Any], records: list[dict[str, Any]]) -> None:
    MANIFEST.parent.mkdir(parents=True, exist_ok=True)
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
    encoded = b"".join(
        (ROOT_NAME + "/" + record["path"]).encode("utf-8") + b"\0" for record in records
    )
    write_bytes_exclusive(MEMBERS, encoded)


def validate_relative_record(value: Any) -> str:
    require(isinstance(value, str) and value, "manifest file path missing")
    relative = Path(value)
    require(not relative.is_absolute() and ".." not in relative.parts, f"manifest path escapes root: {value}")
    require(not any(part == "__pycache__" for part in relative.parts), f"manifest includes __pycache__: {value}")
    if is_target_path(relative):
        require(relative in ALLOWLISTED_TARGET_FILES, f"manifest includes an unbound target file: {value}")
    return relative.as_posix()


def read_manifest() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    require(MANIFEST.is_file() and not MANIFEST.is_symlink(), f"missing archive manifest: {MANIFEST}")
    records: list[dict[str, Any]] = []
    with MANIFEST.open("r", encoding="utf-8") as stream:
        first = stream.readline()
        require(first, "archive manifest is empty")
        header = json.loads(first)
        require(isinstance(header, dict), "archive manifest header is not an object")
        require(header.get("schema") == "litchi.performance.0942.archive-file-manifest.v2",
                "archive manifest schema changed")
        require(header.get("source_root") == str(SOURCE_ROOT), "archive manifest source root changed")
        require(header.get("root_name") == ROOT_NAME, "archive manifest root name changed")
        previous = ""
        for line_number, line in enumerate(stream, start=2):
            record = json.loads(line)
            require(isinstance(record, dict), f"archive manifest line {line_number} is not an object")
            relative = validate_relative_record(record.get("path"))
            require(relative > previous, f"archive manifest paths are not strictly sorted at line {line_number}")
            previous = relative
            require(type(record.get("bytes")) is int and record["bytes"] >= 0,
                    f"archive manifest size invalid at line {line_number}")
            digest = record.get("sha256")
            require(isinstance(digest, str) and len(digest) == 64 and all(c in "0123456789abcdef" for c in digest),
                    f"archive manifest digest invalid at line {line_number}")
            records.append({"path": relative, "bytes": record["bytes"], "sha256": digest})
    require(header.get("files") == len(records), "archive manifest file count mismatch")
    require(header.get("bytes") == sum(item["bytes"] for item in records), "archive manifest byte count mismatch")
    return header, records


def tar_archive_command() -> list[str]:
    return [
        str(TAR), "--create", "--zstd", "--format=posix", "--numeric-owner", "--owner=0", "--group=0",
        "--mtime=@0", "--no-recursion", "--null", "--verbatim-files-from",
        "--directory", str(SOURCE_ROOT.parent), "--files-from", str(MEMBERS),
        "--file", str(ARCHIVE),
    ]


def tar_extract_command() -> list[str]:
    return [
        str(TAR), "--zstd", "--extract", "--file", str(ARCHIVE), "--directory", str(EXTRACT_ROOT),
        "--no-same-owner", "--no-same-permissions",
    ]


def run_logged(command: list[str], stdout_path: Path, stderr_path: Path, *, cwd: Path) -> int:
    stdout_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            completed = subprocess.run(
                command,
                cwd=str(cwd),
                env=dict(TAR_ENVIRONMENT),
                stdin=subprocess.DEVNULL,
                stdout=stdout,
                stderr=stderr,
                check=False,
            )
    except FileExistsError as error:
        raise ArchiveError(f"refusing to replace existing command log: {error}") from error
    return completed.returncode


def script_command(*actions: str) -> str:
    return f"{GO_ENV}=1 /usr/bin/python3 {shlex.quote(str(SCRIPT))} --go " + " ".join(actions)


def require_go() -> None:
    require(os.environ.get(GO_ENV) == GO_VALUE and _GO_FLAG,
            f"archive/cleanup is gated; require --go and {GO_ENV}=1")


_GO_FLAG = False


def require_absent(paths: Iterable[Path]) -> None:
    for path in paths:
        require(not path.exists() and not path.is_symlink(), f"refusing to replace existing artifact: {path}")


def archive_mode() -> dict[str, Any]:
    require_go()
    require(TAR.is_file() and os.access(TAR, os.X_OK), f"tar with zstd support is unavailable: {TAR}")
    require_absent((ARCHIVE, MANIFEST, MEMBERS, ARCHIVE_RECEIPT, VERIFY_RECEIPT, CLEANUP_RECEIPT,
                    EXTRACT_ROOT, ARCHIVE_STDOUT, ARCHIVE_STDERR, VERIFY_STDOUT, VERIFY_STDERR))
    header, records = collect_records()
    write_manifest(header, records)
    write_member_list(records)
    command = tar_archive_command()
    return_code = run_logged(command, ARCHIVE_STDOUT, ARCHIVE_STDERR, cwd=SOURCE_ROOT.parent)
    require(return_code == 0, f"tar archive command failed with exit code {return_code}")
    require(ARCHIVE.is_file() and ARCHIVE.stat().st_size > 0, "tar archive was not produced")
    receipt = {
        "schema": "litchi.performance.0942.archive-receipt.v2",
        "status": "archived",
        "created_utc": utc_now(),
        "script": file_descriptor(SCRIPT),
        "source_root": str(SOURCE_ROOT),
        "manifest": file_descriptor(MANIFEST),
        "member_list": file_descriptor(MEMBERS),
        "archive": file_descriptor(ARCHIVE),
        "archive_stdout": file_descriptor(ARCHIVE_STDOUT),
        "archive_stderr": file_descriptor(ARCHIVE_STDERR),
        "files": header["files"],
        "bytes": header["bytes"],
        "bound_profile_binaries": expected_binary_descriptors(),
        "tar_command": command,
        "tar_command_shell": shlex.join(command),
        "regeneration_commands": {
            "archive": script_command("--archive"),
            "verify": script_command("--verify"),
            "cleanup_after_verification": script_command("--cleanup"),
        },
        "excluded_rules": header["excluded"],
        "cleanup_policy": {
            "preserve_bound_profile_binaries": True,
            "remove_target_cache_components": True,
            "remove_pycache_directories": True,
            "remove_verification_extraction": str(EXTRACT_ROOT),
        },
    }
    write_json_exclusive(ARCHIVE_RECEIPT, receipt)
    return {"status": receipt["status"], "files": header["files"], "bytes": header["bytes"],
            "archive": receipt["archive"], "archive_receipt": str(ARCHIVE_RECEIPT)}


def extracted_file_records() -> list[dict[str, Any]]:
    root = EXTRACT_ROOT / ROOT_NAME
    require(not root.is_symlink() and root.is_dir(), f"archive extraction root missing: {root}")
    actual: list[dict[str, Any]] = []
    for current, directories, filenames in os.walk(root, topdown=True, followlinks=False):
        current_path = Path(current)
        for name in directories:
            require(not (current_path / name).is_symlink(), f"symlink in archive extraction: {current_path / name}")
        for name in sorted(filenames):
            path = current_path / name
            require(not path.is_symlink() and path.is_file(), f"non-regular archive extraction member: {path}")
            relative = path.relative_to(root).as_posix()
            actual.append({"path": ROOT_NAME + "/" + relative, "bytes": path.stat().st_size,
                           "sha256": sha256_path(path)})
    actual.sort(key=lambda item: item["path"])
    return actual


def load_receipt(path: Path, schema: str, status: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing receipt: {path}")
    with path.open("r", encoding="utf-8") as stream:
        value = json.load(stream)
    require(isinstance(value, dict) and value.get("schema") == schema and value.get("status") == status,
            f"receipt is not {status}: {path}")
    return value


def verify_mode() -> dict[str, Any]:
    require_go()
    archive_receipt = load_receipt(ARCHIVE_RECEIPT, "litchi.performance.0942.archive-receipt.v2", "archived")
    header, records = read_manifest()
    bound_binaries = expected_binary_descriptors()
    require(archive_receipt.get("bound_profile_binaries") == bound_binaries,
            "bound profile binary custody differs from archive receipt")
    record_paths = {item["path"] for item in records}
    require(all(relative.as_posix() in record_paths for relative in ALLOWLISTED_TARGET_FILES),
            "archive manifest omitted a bound profile release binary")
    require(file_descriptor(ARCHIVE) == archive_receipt["archive"], "archive changed since archive receipt")
    require(file_descriptor(MANIFEST) == archive_receipt["manifest"], "manifest changed since archive receipt")
    require_absent((EXTRACT_ROOT, VERIFY_RECEIPT, VERIFY_STDOUT, VERIFY_STDERR))
    EXTRACT_ROOT.mkdir(parents=True, exist_ok=False)
    command = tar_extract_command()
    return_code = run_logged(command, VERIFY_STDOUT, VERIFY_STDERR, cwd=SOURCE_ROOT.parent)
    require(return_code == 0, f"tar extraction command failed with exit code {return_code}")
    actual = extracted_file_records()
    expected = [{"path": ROOT_NAME + "/" + item["path"], "bytes": item["bytes"], "sha256": item["sha256"]}
                for item in records]
    require(actual == expected, "archive extraction file set, size, or digest differs from manifest")
    require(len(actual) == header["files"], "archive extraction file count differs from manifest")
    require(sum(item["bytes"] for item in actual) == header["bytes"],
            "archive extraction byte total differs from manifest")
    receipt = {
        "schema": "litchi.performance.0942.archive-verify-receipt.v2",
        "status": "passed",
        "verified_utc": utc_now(),
        "script": file_descriptor(SCRIPT),
        "archive": file_descriptor(ARCHIVE),
        "archive_receipt": file_descriptor(ARCHIVE_RECEIPT),
        "manifest": file_descriptor(MANIFEST),
        "extraction_root": str(EXTRACT_ROOT),
        "files": header["files"],
        "bytes": header["bytes"],
        "tar_command": command,
        "tar_command_shell": shlex.join(command),
        "verify_stdout": file_descriptor(VERIFY_STDOUT),
        "verify_stderr": file_descriptor(VERIFY_STDERR),
        "every_file_rehashed": True,
        "file_set_exact": True,
        "cleanup_authorized_by_this_receipt": False,
    }
    write_json_exclusive(VERIFY_RECEIPT, receipt)
    return {"status": receipt["status"], "files": receipt["files"], "bytes": receipt["bytes"],
            "verify_receipt": str(VERIFY_RECEIPT)}


def current_matches_archive(records: list[dict[str, Any]]) -> None:
    _header, current = collect_records()
    require(current == records, "0942 selected evidence changed after archive verification")


def remove_path(path: Path) -> None:
    require(SOURCE_ROOT in path.parents, f"refusing removal outside 0942 root: {path}")
    if path.is_symlink() or path.is_file():
        path.unlink()
    elif path.is_dir():
        shutil.rmtree(path)
    else:
        raise ArchiveError(f"cannot remove unknown path type: {path}")


def cleanup_target_caches() -> list[str]:
    removed: list[str] = []
    def has_allowlisted_descendant(relative: Path) -> bool:
        return any(
            allowed.parts[: len(relative.parts)] == relative.parts
            for allowed in ALLOWLISTED_TARGET_FILES
        )

    # Only the two known top-level profile caches are cleanup targets.  Do not
    # walk the whole source tree looking for a component named ``target-*``:
    # source corpora legitimately contain such filenames, including iWork
    # fuzz-seed artifacts that must remain in the archive.
    for cache_relative in sorted(KNOWN_TARGET_CACHE_DIRS, key=Path.as_posix):
        cache_root = SOURCE_ROOT / cache_relative
        if not cache_root.exists() and not cache_root.is_symlink():
            continue
        require(cache_root.is_dir() and not cache_root.is_symlink(),
                f"known target cache is not a directory: {cache_root}")
        for current, directories, filenames in os.walk(cache_root, topdown=False, followlinks=False):
            current_path = Path(current)
            for name in sorted(filenames):
                path = current_path / name
                relative = path.relative_to(SOURCE_ROOT)
                if relative in ALLOWLISTED_TARGET_FILES:
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
    return removed


def cleanup_pycache() -> list[str]:
    removed: list[str] = []
    for current, directories, _filenames in os.walk(SOURCE_ROOT, topdown=False, followlinks=False):
        current_path = Path(current)
        for name in sorted(directories):
            if name != PYCACHE_NAME:
                continue
            path = current_path / name
            relative = path.relative_to(SOURCE_ROOT)
            if not path.exists() and not path.is_symlink():
                continue
            require(not path.is_symlink(), f"symlink __pycache__ refused: {path}")
            remove_path(path)
            removed.append(relative.as_posix())
    return removed


def assert_post_cleanup() -> dict[str, dict[str, Any]]:
    descriptors = expected_binary_descriptors()
    for current, directories, filenames in os.walk(SOURCE_ROOT, topdown=True, followlinks=False):
        current_path = Path(current)
        for name in directories:
            path = current_path / name
            relative = path.relative_to(SOURCE_ROOT)
            require(name != PYCACHE_NAME, f"__pycache__ remains after cleanup: {path}")
            if is_target_path(relative):
                require(
                    any(allowed.parts[: len(relative.parts)] == relative.parts
                        for allowed in ALLOWLISTED_TARGET_FILES),
                    f"unallowlisted target cache remains: {path}",
                )
                require(not path.is_symlink(), f"symlink remains after cleanup: {path}")
        for name in filenames:
            path = current_path / name
            relative = path.relative_to(SOURCE_ROOT)
            require(PYCACHE_NAME not in relative.parts, f"__pycache__ file remains: {path}")
            if is_target_path(relative):
                require(relative in ALLOWLISTED_TARGET_FILES,
                        f"unallowlisted target file remains: {relative}")
    return descriptors


def cleanup_mode() -> dict[str, Any]:
    require_go()
    archive_receipt = load_receipt(ARCHIVE_RECEIPT, "litchi.performance.0942.archive-receipt.v2", "archived")
    verify_receipt = load_receipt(VERIFY_RECEIPT, "litchi.performance.0942.archive-verify-receipt.v2", "passed")
    header, records = read_manifest()
    bound_binaries = expected_binary_descriptors()
    require(archive_receipt.get("bound_profile_binaries") == bound_binaries,
            "bound profile binary custody differs from archive receipt")
    require(file_descriptor(ARCHIVE) == archive_receipt["archive"], "archive changed before cleanup")
    require(file_descriptor(MANIFEST) == archive_receipt["manifest"], "manifest changed before cleanup")
    require(verify_receipt["archive"] == archive_receipt["archive"], "verify receipt archive binding changed")
    require(EXTRACT_ROOT.is_dir() and not EXTRACT_ROOT.is_symlink(),
            "verified extraction is missing; cleanup requires completed verification")
    current_matches_archive(records)
    expected_binary_descriptors()
    removed_targets = cleanup_target_caches()
    removed_pycache = cleanup_pycache()
    require(EXTRACT_ROOT.parent == WORK_ROOT and EXTRACT_ROOT.name.endswith(".extract"),
            "verification extraction path changed")
    shutil.rmtree(EXTRACT_ROOT)
    preserved = assert_post_cleanup()
    receipt = {
        "schema": "litchi.performance.0942.archive-cleanup-receipt.v2",
        "status": "completed",
        "cleaned_utc": utc_now(),
        "script": file_descriptor(SCRIPT),
        "source_root": str(SOURCE_ROOT),
        "archive": archive_receipt["archive"],
        "manifest": archive_receipt["manifest"],
        "verify_receipt": file_descriptor(VERIFY_RECEIPT),
        "files": header["files"],
        "bytes": header["bytes"],
        "removed_target_cache_paths": removed_targets,
        "removed_pycache_paths": removed_pycache,
        "removed_verification_extraction": str(EXTRACT_ROOT),
        "preserved_bound_profile_binaries": preserved,
        "checkout_touched": False,
        "other_performance_roots_touched": [],
    }
    write_json_exclusive(CLEANUP_RECEIPT, receipt)
    return {"status": receipt["status"], "removed_target_paths": len(removed_targets),
            "removed_pycache_paths": removed_pycache, "cleanup_receipt": str(CLEANUP_RECEIPT)}


def plan_mode() -> dict[str, Any]:
    binaries: dict[str, Any] = {}
    for relative, expected in sorted(EXPECTED_TARGET_BINARIES.items(), key=lambda item: item[0].as_posix()):
        path = SOURCE_ROOT / relative
        binaries[relative.as_posix()] = {
            "path": str(path),
            "expected_bytes": expected["bytes"],
            "expected_sha256": expected["sha256"],
            "present": path.is_file() and not path.is_symlink(),
        }
    return {
        "schema": "litchi.performance.0942.archive-plan.v2",
        "status": "prepared_not_started",
        "script": str(SCRIPT),
        "source_root": str(SOURCE_ROOT),
        "workspace": str(WORK_ROOT),
        "archive": str(ARCHIVE),
        "manifest": str(MANIFEST),
        "member_list": str(MEMBERS),
        "verification_extraction": str(EXTRACT_ROOT),
        "include": "all regular files under 0942 root, except __pycache__ and only the known top-level target-profile-v1/target-profile-v2 cache contents outside the two bound binaries",
        "target_binary_allowlist": binaries,
        "exclude_directory_name": "__pycache__",
        "cleanup_target_policy": "remove only known top-level target-profile-v1/target-profile-v2 cache contents after archive verification, preserving the two allowlisted profile binaries and every nested target-* source name elsewhere",
        "cleanup_pycache_policy": "remove every __pycache__ directory after archive verification",
        "cleanup_other_roots": [],
        "go_gate": {"flag": "--go", "environment": f"{GO_ENV}=1"},
        "exact_regeneration_commands": {
            "archive": script_command("--archive"),
            "verify": script_command("--verify"),
            "cleanup": script_command("--cleanup"),
        },
    }


def main(argv: list[str] | None = None) -> int:
    global _GO_FLAG
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--plan", action="store_true")
    modes.add_argument("--archive", action="store_true")
    modes.add_argument("--verify", action="store_true")
    modes.add_argument("--cleanup", action="store_true")
    parser.add_argument("--go", action="store_true", help="required with the archive, verify, and cleanup modes")
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
        print(f"0942 archive operation refused: {type(error).__name__}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
