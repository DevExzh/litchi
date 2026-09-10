#!/usr/bin/env python3
"""Bounded cleanup for the 0497 tail append experiment.

The build and capture evidence lives below this directory while disposable
Cargo/source scratch lives below ``/home/zhuhe/.cache/litchi-goal-0497``.  The
driver authenticates both phase build receipts and the ASan fuzz build/run
bundle before removing anything.  It only removes the named children owned by
this experiment, keeps the four benchmark executables plus the retained fuzz
executable, and records an exclusive ``cleanup.json`` receipt.

This module is intentionally usable as a read-only planning library.  A
caller should run ``--dry-run`` and review the resulting inventory before
calling the destructive mode.  All destructive paths are checked again after
planning, and process references, symlinks, special files, and incomplete
terminal receipts fail closed.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import shutil
import stat
import sys
from typing import Any, Iterable, Mapping


sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parent
TEMP = Path("/home/zhuhe/.cache/litchi-goal-0497")
TARGET = TEMP / "target"
TARGET_DIR = TARGET
FUZZ_TEMP = Path("/home/zhuhe/.cache/litchi-goal-0497/fuzz-target")
FUZZ_TEMP_NAME = "fuzz-target"
FUZZ_EVIDENCE_NAME = "fuzz"
FUZZ_BINARY_NAME = "source_backed_tail_append_stream"
FUZZ_TARGET_TRIPLE = "x86_64-unknown-linux-gnu"
FUZZ_RETAINED = TEMP / "retained" / "fuzz" / FUZZ_BINARY_NAME
FUZZ_RUN_SEEDS = (497, 498)
EARLY_TARGET_CLEANUP = "early-target-cleanup.json"
EARLY_FUZZ_CLEANUP = "early-fuzz-target-cleanup.json"
PROTECTED_WORKTREE = Path("/home/zhuhe/code/litchi-spec-gaps")
SCHEMA = "docx-tail-append-0497-cleanup-v1"
VERIFY_SCHEMA = "docx-tail-append-0497-cleanup-verification-v1"
VERSION = 1
SEAL_NAME = "cleanup.json"
BUILD_ROLES = ("normal", "allocator")
PHASES = ("before", "after")
BINARY_NAME = "docx_replayable_tail_append"
OWNED_CHILDREN = frozenset(
    {"before", "after", "target", "tmp", "projections", "runs", "publication-profiles",
     FUZZ_TEMP_NAME}
)
BUILD_RECEIPTS = {
    "before": ROOT / "build-before-baseline1.json",
    "after": ROOT / "build-after-candidate1.json",
}
ROOT_PROCESS_SCRIPTS = frozenset(
    {
        "cleanup.py",
        "build.py",
        "gate.py",
        "gates.py",
        "measure.py",
        "profile.py",
        "fuzz.py",
        "seal.py",
        "test_cleanup.py",
        "test_measure.py",
        "test_profile.py",
    }
)
PROTECTED_DAEMONS = frozenset({"systemd", "(sd-pam)", "sshd-session"})
_DELETED_SUFFIX = " (deleted)"
_BLOCK_SIZE = 512


class CleanupError(RuntimeError):
    """A custody or process-safety precondition failed."""


def fail(message: str) -> None:
    raise CleanupError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _lstat(path: Path, label: str) -> os.stat_result:
    try:
        return path.lstat()
    except OSError as error:
        fail(f"{label}: cannot stat {path}: {error}")
    raise AssertionError("unreachable")


def _no_symlink_components(path: Path, label: str) -> None:
    absolute = Path(os.path.abspath(path))
    current = Path(absolute.anchor)
    for component in absolute.parts[1:]:
        current /= component
        try:
            state = current.lstat()
        except FileNotFoundError:
            continue
        except OSError as error:
            fail(f"{label}: cannot inspect {current}: {error}")
        require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink component: {current}")


def _regular(path: Path, label: str, *, owner_uid: int | None = None,
             allow_missing: bool = False) -> os.stat_result | None:
    if allow_missing and not path.exists() and not path.is_symlink():
        return None
    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is forbidden: {path}")
    require(stat.S_ISREG(state.st_mode), f"{label}: regular file required: {path}")
    if owner_uid is not None:
        require(state.st_uid == owner_uid, f"{label}: unexpected owner: {path}")
    return state


def _directory(path: Path, label: str, *, owner_uid: int | None = None,
               allow_missing: bool = False) -> os.stat_result | None:
    if allow_missing and not path.exists() and not path.is_symlink():
        return None
    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is forbidden: {path}")
    require(stat.S_ISDIR(state.st_mode), f"{label}: directory required: {path}")
    if owner_uid is not None:
        require(state.st_uid == owner_uid, f"{label}: unexpected owner: {path}")
    return state


def _canonical_directory(path: Path, label: str, *, owner_uid: int | None = None,
                         allow_missing: bool = False) -> Path:
    _no_symlink_components(path, label)
    if allow_missing and not path.exists() and not path.is_symlink():
        return path.absolute()
    _directory(path, label, owner_uid=owner_uid)
    try:
        resolved = path.resolve(strict=True)
    except OSError as error:
        fail(f"{label}: cannot resolve {path}: {error}")
    require(resolved == path.absolute(), f"{label}: path is not canonical: {path}")
    return resolved


def _path_inside(path: Path, root: Path) -> bool:
    try:
        path.resolve(strict=False).relative_to(root.resolve(strict=False))
    except ValueError:
        return False
    return True


def _relative(path: Path, root: Path, label: str) -> str:
    try:
        return path.resolve(strict=False).relative_to(root.resolve(strict=False)).as_posix()
    except ValueError as error:
        fail(f"{label}: path escapes {root}: {error}")
    raise AssertionError("unreachable")


def _sha256(path: Path) -> str:
    _regular(path, f"hash {path}")
    try:
        with path.open("rb") as stream:
            return hashlib.file_digest(stream, "sha256").hexdigest()
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    raise AssertionError("unreachable")


def _read_json(path: Path, label: str) -> Any:
    _no_symlink_components(path, label)
    _regular(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        fail(f"{label}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def _descriptor(path: Path, label: str, *, root: Path | None = None) -> dict[str, Any]:
    _no_symlink_components(path, label)
    state = _regular(path, label)
    assert state is not None
    record: dict[str, Any] = {
        "path": str(path.resolve(strict=True)),
        "bytes": state.st_size,
        "sha256": _sha256(path),
    }
    if root is not None:
        record["relative_path"] = _relative(path, root, label)
    return record


def _valid_digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and len(value) == 64
            and all(char in "0123456789abcdef" for char in value),
            f"{label}: SHA-256 is malformed")
    return value


def _safe_token(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and len(value) <= 128
            and all(char.isalnum() or char in "_-" for char in value),
            f"{label}: unsafe token")
    return value


def _path_descriptor(value: Any, label: str) -> tuple[Path, int, str]:
    require(isinstance(value, Mapping), f"{label}: descriptor is missing")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and raw_path.startswith("/"),
            f"{label}: absolute path is required")
    size = value.get("bytes")
    require(type(size) is int and size >= 0, f"{label}: byte length is malformed")
    digest = _valid_digest(value.get("sha256"), f"{label}.sha256")
    return Path(raw_path), size, digest


def _validate_descriptor(value: Any, path: Path, label: str, *, present: bool,
                         executable: bool = False) -> dict[str, Any]:
    recorded_path, size, digest = _path_descriptor(value, label)
    _no_symlink_components(path, label)
    require(recorded_path.resolve(strict=False) == path.resolve(strict=False),
            f"{label}: path binding differs")
    if present:
        state = _regular(path, label)
        assert state is not None
        require(state.st_size == size and _sha256(path) == digest,
                f"{label}: content changed")
        if executable:
            require(bool(state.st_mode & 0o111) and os.access(path, os.X_OK),
                    f"{label}: executable bit is missing")
    return {"path": str(path.resolve(strict=False)), "bytes": size,
            "sha256": digest}


def _assert_disjoint(root: Path, temp: Path, target: Path,
                     fuzz_temp: Path | None = None) -> None:
    # 0497 deliberately keeps Cargo's target directory as the owned
    # ``temporary_root/target`` child.  Only evidence must be disjoint from
    # that scope; the target/temporary overlap is the bounded layout itself.
    pairs = ((root, temp, "evidence and temporary roots overlap"),
             (root, target, "evidence and target roots overlap"))
    for left, right, message in pairs:
        require(not _path_inside(left, right) and not _path_inside(right, left), message)
    if fuzz_temp is not None:
        require(fuzz_temp.resolve(strict=False) == (temp / FUZZ_TEMP_NAME).resolve(strict=False),
                "fuzz temporary root is outside the owned benchmark temporary root")
    protected = PROTECTED_WORKTREE.resolve(strict=False)
    for path in (root, temp, target):
        require(not _path_inside(path, protected) and not _path_inside(protected, path),
                "protected worktree overlaps cleanup scope")


def _allocated_bytes(state: os.stat_result) -> int:
    blocks = getattr(state, "st_blocks", None)
    if blocks is None:
        return ((state.st_size + _BLOCK_SIZE - 1) // _BLOCK_SIZE) * _BLOCK_SIZE
    require(type(blocks) is int and blocks >= 0, "invalid allocated block count")
    return blocks * _BLOCK_SIZE


@dataclass(frozen=True)
class _TreeStats:
    path: Path
    kind: str
    files: int
    directories: int
    logical_bytes: int
    allocated_bytes: int
    identity: tuple[int, int]
    fingerprint: str


def _tree_stats(path: Path, label: str, *, owner_uid: int | None = None,
                root_device: int | None = None) -> _TreeStats:
    root_state = _lstat(path, label)
    require(not stat.S_ISLNK(root_state.st_mode), f"{label}: symlink is forbidden: {path}")
    require(stat.S_ISREG(root_state.st_mode) or stat.S_ISDIR(root_state.st_mode),
            f"{label}: only regular files and directories are allowed: {path}")
    if root_device is None:
        root_device = root_state.st_dev
    require(root_state.st_dev == root_device, f"{label}: device boundary: {path}")
    seen: set[tuple[int, int]] = set()
    fingerprint = hashlib.sha256()

    def visit(current: Path, relative: str) -> tuple[int, int, int, int]:
        state = _lstat(current, f"{label}/{relative or '.'}")
        require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is forbidden: {current}")
        require(stat.S_ISREG(state.st_mode) or stat.S_ISDIR(state.st_mode),
                f"{label}: special file is forbidden: {current}")
        require(state.st_dev == root_device, f"{label}: device boundary: {current}")
        if owner_uid is not None:
            require(state.st_uid == owner_uid, f"{label}: unexpected owner: {current}")
        kind = "file" if stat.S_ISREG(state.st_mode) else "directory"
        allocated = _allocated_bytes(state)
        identity = (state.st_dev, state.st_ino)
        fingerprint.update("\0".join(str(value) for value in (
            relative, kind, state.st_size if kind == "file" else 0,
            allocated, state.st_mode, state.st_mtime_ns, state.st_dev, state.st_ino,
        )).encode())
        fingerprint.update(b"\n")
        unique_allocated = allocated if identity not in seen else 0
        seen.add(identity)
        if kind == "file":
            return 1, 0, state.st_size, unique_allocated
        files = directories = logical = 0
        allocated_total = unique_allocated
        try:
            children = sorted(current.iterdir(), key=lambda item: item.name)
        except OSError as error:
            fail(f"{label}: cannot enumerate {current}: {error}")
        for child in children:
            child_files, child_dirs, child_logical, child_allocated = visit(
                child, f"{relative}/{child.name}" if relative else child.name
            )
            files += child_files
            directories += child_dirs
            logical += child_logical
            allocated_total += child_allocated
        return files, directories + 1, logical, allocated_total

    files, directories, logical, allocated = visit(path, "")
    return _TreeStats(
        path=path,
        kind="directory" if stat.S_ISDIR(root_state.st_mode) else "file",
        files=files,
        directories=directories,
        logical_bytes=logical,
        allocated_bytes=allocated,
        identity=(root_state.st_dev, root_state.st_ino),
        fingerprint=fingerprint.hexdigest(),
    )


def _stats_record(stats: _TreeStats, *, root: Path, label: str) -> dict[str, Any]:
    return {
        "path": str(stats.path.resolve(strict=False)),
        "relative_path": _relative(stats.path, root, label),
        "kind": stats.kind,
        "files": stats.files,
        "directories": stats.directories,
        "logical_bytes": stats.logical_bytes,
        "allocated_bytes": stats.allocated_bytes,
        "device": stats.identity[0],
        "inode": stats.identity[1],
        "fingerprint": stats.fingerprint,
    }


def _disk(path: Path) -> dict[str, int]:
    try:
        value = shutil.disk_usage(path)
    except OSError as error:
        fail(f"cannot inspect free space at {path}: {error}")
    return {"total_bytes": value.total, "used_bytes": value.used, "free_bytes": value.free}


def _early_target_binding(root: Path, target: Path, *, required: bool) -> dict[str, Any] | None:
    """Authenticate a target removed before final cleanup planning.

    The formal capture may remove a completed Cargo target to recover disk
    space before the final evidence cleanup.  An absent target is acceptable
    only when this receipt explains the removal and proves that no build was
    still using it.
    """

    path = root / EARLY_TARGET_CLEANUP
    if not path.exists() and not path.is_symlink():
        require(not required, f"early target cleanup receipt is missing: {path}")
        return None
    value = _read_json(path, "early target cleanup")
    require(isinstance(value, Mapping)
            and set(value) == {"path", "reason", "allocated_bytes", "free_before",
                                "free_after", "completed_ns", "active_build_refs"},
            "early target cleanup schema differs")
    require(value.get("path") == str(target.resolve(strict=False)),
            "early target cleanup path differs")
    require(isinstance(value.get("reason"), str) and value["reason"].strip(),
            "early target cleanup reason is missing")
    for key in ("allocated_bytes", "free_before", "free_after", "completed_ns"):
        require(type(value.get(key)) is int and value[key] >= 0,
                f"early target cleanup {key} is malformed")
    require(value["allocated_bytes"] > 0 and value["completed_ns"] > 0,
            "early target cleanup counters are empty")
    require(value.get("active_build_refs") == [],
            "early target cleanup has active build references")
    return {
        "descriptor": _descriptor(path, "early target cleanup", root=root),
        "path": value["path"],
        "reason": value["reason"],
        "allocated_bytes": value["allocated_bytes"],
        "free_before": value["free_before"],
        "free_after": value["free_after"],
        "completed_ns": value["completed_ns"],
        "active_build_refs": [],
    }


def _early_fuzz_binding(root: Path, fuzz_temp: Path, *, required: bool) -> dict[str, Any] | None:
    """Authenticate the separately removed ASan Cargo target."""

    path = root / EARLY_FUZZ_CLEANUP
    if not path.exists() and not path.is_symlink():
        require(not required, f"early fuzz target cleanup receipt is missing: {path}")
        return None
    value = _read_json(path, "early fuzz target cleanup")
    require(isinstance(value, Mapping)
            and set(value) == {"path", "allocated_bytes", "completed_ns", "reason"},
            "early fuzz target cleanup schema differs")
    target = fuzz_temp / "target"
    require(value.get("path") == str(target.resolve(strict=False)),
            "early fuzz target cleanup path differs")
    require(isinstance(value.get("reason"), str) and value["reason"].strip(),
            "early fuzz target cleanup reason is missing")
    require(type(value.get("allocated_bytes")) is int and value["allocated_bytes"] > 0
            and type(value.get("completed_ns")) is int and value["completed_ns"] > 0,
            "early fuzz target cleanup counters are malformed")
    return {
        "descriptor": _descriptor(path, "early fuzz target cleanup", root=root),
        "path": value["path"],
        "allocated_bytes": value["allocated_bytes"],
        "completed_ns": value["completed_ns"],
        "reason": value["reason"],
    }


def _coerce_build_receipts(build_receipts: Mapping[str, Any] | None = None,
                           before_build: Path | None = None,
                           after_build: Path | None = None) -> dict[str, Path]:
    values: dict[str, Any] = dict(BUILD_RECEIPTS)
    if build_receipts is not None:
        require(isinstance(build_receipts, Mapping), "build receipt mapping is malformed")
        for phase in PHASES:
            if phase in build_receipts:
                values[phase] = build_receipts[phase]
            elif f"{phase}-build" in build_receipts:
                values[phase] = build_receipts[f"{phase}-build"]
    if before_build is not None:
        values["before"] = before_build
    if after_build is not None:
        values["after"] = after_build
    result: dict[str, Path] = {}
    for phase in PHASES:
        require(isinstance(values[phase], (str, Path)),
                f"{phase} build receipt path is malformed")
        result[phase] = Path(values[phase])
    return result


def _receipt_path(path: Path, root: Path, label: str, *, allow_root: bool = False) -> Path:
    _no_symlink_components(path, label)
    resolved = path.resolve(strict=False)
    allowed = _path_inside(resolved, root / "builds")
    if allow_root:
        allowed = allowed or resolved.parent == root.resolve(strict=False)
    require(allowed, f"{label}: path is outside the evidence builds directory")
    return resolved


def _manifest_key(key: Any, label: str) -> PurePosixPath:
    require(isinstance(key, str) and key and not key.startswith("/"),
            f"{label}: manifest key is malformed")
    path = PurePosixPath(key)
    require(".." not in path.parts and not path.is_absolute(),
            f"{label}: manifest path escapes the phase clone")
    return path


def _validate_source_manifest(path: Path, phase_root: Path, label: str, *, present: bool) -> dict[str, Any]:
    value = _read_json(path, label)
    require(isinstance(value, Mapping) and value, f"{label}: source manifest is empty")
    normalized: dict[str, Any] = {}
    for raw_name, raw_descriptor in sorted(value.items(), key=lambda item: str(item[0])):
        name = _manifest_key(raw_name, f"{label} entry")
        expected = phase_root.joinpath(*name.parts)
        descriptor = _validate_descriptor(raw_descriptor, expected,
                                          f"{label} {name}", present=present)
        normalized[name.as_posix()] = descriptor
    return normalized


def _validate_gate(path: Path, phase: str, root: Path, temp: Path,
                   source_manifest: Mapping[str, Any], label: str) -> dict[str, Any]:
    value = _read_json(path, label)
    require(isinstance(value, Mapping), f"{label}: terminal gate is not an object")
    require(type(value.get("exit_code")) is int and value.get("exit_code") == 0,
            f"{label}: build gate did not pass")
    require(value.get("source_unchanged") is True, f"{label}: source changed")
    require(value.get("cwd") == str((temp / phase).resolve(strict=False)),
            f"{label}: source cwd differs")
    recorded_manifest = value.get("source_manifest")
    manifest_path, _, _ = _path_descriptor(recorded_manifest, f"{label}.source_manifest")
    require(manifest_path.resolve(strict=False) == Path(str(source_manifest["path"])).resolve(strict=False),
            f"{label}: source manifest binding differs")
    driver = value.get("driver")
    driver_path, _, _ = _path_descriptor(driver, f"{label}.driver")
    require(driver_path.resolve(strict=False) == (root / "build.py").resolve(strict=False),
            f"{label}: build driver differs")
    _validate_descriptor(driver, root / "build.py", f"{label}.driver", present=True)
    for field in ("stdout", "stderr"):
        if field in value:
            output_path, _, _ = _path_descriptor(value[field], f"{label}.{field}")
            require(_path_inside(output_path, root / "builds"),
                    f"{label}.{field}: output escapes evidence")
            _validate_descriptor(value[field], output_path, f"{label}.{field}", present=True)
    started = path.with_suffix(".started.json")
    _regular(started, f"{label} started receipt")
    _read_json(started, f"{label} started receipt")
    terminal = _descriptor(path, label, root=root)
    terminal["exit_code"] = value["exit_code"]
    terminal["source_unchanged"] = True
    return terminal


def _validate_build_phase(phase: str, root: Path, temp: Path, target: Path,
                          receipt_path: Path, *, target_present: bool) -> dict[str, Any]:
    path = _receipt_path(receipt_path, root, f"{phase} build receipt", allow_root=True)
    value = _read_json(path, f"{phase} build receipt")
    expected_keys = {f"{phase}/{role}" for role in BUILD_ROLES}
    require(isinstance(value, Mapping) and set(value) == expected_keys,
            f"{phase} build receipt keys differ")
    records: dict[str, Any] = {}
    source_identity: dict[str, str] | None = None
    for role in BUILD_ROLES:
        key = f"{phase}/{role}"
        entry = value[key]
        require(isinstance(entry, Mapping)
                and set(entry) == {"binary", "gate", "git_revision", "source_manifest"},
                f"{key} build receipt fields differ")
        revision = entry.get("git_revision")
        require(isinstance(revision, str) and len(revision) == 40
                and all(char in "0123456789abcdef" for char in revision),
                f"{key}: git revision is malformed")
        source_meta = entry["source_manifest"]
        source_path, _, _ = _path_descriptor(source_meta, f"{key}.source_manifest")
        source_path = _receipt_path(source_path, root, f"{key}.source_manifest")
        normalized_source = _validate_descriptor(source_meta, source_path,
                                                 f"{key}.source_manifest", present=True)
        source_manifest = _validate_source_manifest(
            source_path, temp / phase, f"{key}.source_manifest", present=target_present
        )
        if source_identity is None:
            source_identity = normalized_source
        else:
            require(normalized_source == source_identity,
                    f"{phase}: normal and allocator source manifests differ")
        gate_meta = entry["gate"]
        gate_path, _, _ = _path_descriptor(gate_meta, f"{key}.gate")
        gate_path = _receipt_path(gate_path, root, f"{key}.gate")
        gate = _validate_descriptor(gate_meta, gate_path, f"{key}.gate", present=True)
        gate_terminal = _validate_gate(gate_path, phase, root, temp, normalized_source,
                                       f"{key}.gate")
        retained = temp / "retained" / phase / role / BINARY_NAME
        binary = _validate_descriptor(entry["binary"], retained, f"{key}.binary",
                                      present=True, executable=True)
        records[key] = {
            "phase": phase,
            "role": role,
            "receipt_path": path,
            "receipt": _descriptor(path, f"{phase} build receipt", root=root),
            "git_revision": revision,
            "source_manifest": normalized_source,
            "source_manifest_path": source_path,
            "gate": gate,
            "gate_path": gate_path,
            "gate_terminal": gate_terminal,
            "binary": binary | {"executable": True},
        }
    return {
        "phase": phase,
        "path": path,
        "receipt": _descriptor(path, f"{phase} build receipt", root=root),
        "records": records,
    }


def _validate_retained_layout(temp: Path, builds: Mapping[str, Any]) -> None:
    retained = temp / "retained"
    _directory(retained, "retained root", owner_uid=os.getuid())
    require({child.name for child in retained.iterdir()} == set(PHASES) | {"fuzz"},
            "retained root has unexpected children")
    for phase in PHASES:
        phase_root = retained / phase
        _directory(phase_root, f"retained {phase} root", owner_uid=os.getuid())
        require({child.name for child in phase_root.iterdir()} == set(BUILD_ROLES),
                f"retained {phase} root has unexpected children")
        for role in BUILD_ROLES:
            role_root = phase_root / role
            _directory(role_root, f"retained {phase}/{role} root", owner_uid=os.getuid())
            require({child.name for child in role_root.iterdir()} == {BINARY_NAME},
                    f"retained {phase}/{role} has unexpected children")
            binary = role_root / BINARY_NAME
            state = _regular(binary, f"retained {phase}/{role} binary", owner_uid=os.getuid())
            assert state is not None
            require(bool(state.st_mode & 0o111) and os.access(binary, os.X_OK),
                    f"retained {phase}/{role} binary is not executable")
            key = f"{phase}/{role}"
            require(builds[key]["binary"]["path"] == str(binary.resolve(strict=False)),
                    f"retained {key} path differs")
    fuzz_root = retained / "fuzz"
    _directory(fuzz_root, "retained fuzz root", owner_uid=os.getuid())
    require({child.name for child in fuzz_root.iterdir()} == {FUZZ_BINARY_NAME},
            "retained fuzz root has unexpected children")
    fuzz_binary = fuzz_root / FUZZ_BINARY_NAME
    fuzz_state = _regular(fuzz_binary, "retained fuzz binary", owner_uid=os.getuid())
    assert fuzz_state is not None
    require(bool(fuzz_state.st_mode & 0o111) and os.access(fuzz_binary, os.X_OK),
            "retained fuzz binary is not executable")


def _terminal_summary(root: Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """Return stable terminal evidence and private records used for custody."""
    _directory(root, "evidence root")
    started_paths: list[tuple[Path, Path]] = []
    try:
        entries = sorted(root.rglob("*"), key=lambda item: str(item))
    except OSError as error:
        fail(f"evidence root: cannot enumerate terminal receipts: {error}")
    for item in entries:
        if item.is_symlink():
            fail(f"evidence root contains symlink: {item}")
        if not item.is_file():
            continue
        if item.name.endswith(".started.json"):
            started_paths.append((item, item.with_name(item.name[:-len(".started.json")] + ".json")))
        elif item.name == "started.json":
            started_paths.append((item, item.with_name("terminal.json")))
    require(started_paths, "no started terminal receipts were found")
    unfinished: list[str] = []
    public: list[dict[str, Any]] = []
    records: list[dict[str, Any]] = []
    seen: set[Path] = set()
    for started, terminal in started_paths:
        if terminal in seen:
            continue
        seen.add(terminal)
        _regular(started, f"terminal start {started}")
        if terminal.is_symlink() or not terminal.exists():
            unfinished.append(str(started))
            continue
        _regular(terminal, f"terminal receipt {terminal}")
        value = _read_json(terminal, f"terminal receipt {terminal}")
        require(isinstance(value, Mapping), f"terminal receipt {terminal}: object required")
        exit_code = value.get("exit_code")
        status = value.get("status")
        require(type(exit_code) is int or isinstance(status, str),
                f"terminal receipt {terminal}: status is malformed")
        cleanup = value.get("cleanup")
        cleanup_status = cleanup.get("status") if isinstance(cleanup, Mapping) else None
        cleanup_remaining = cleanup.get("remaining") if isinstance(cleanup, Mapping) else None
        record = {
            "path": str(terminal.resolve(strict=True)),
            "descriptor": _descriptor(terminal, f"terminal receipt {terminal}", root=root),
            "status": status,
            "exit_code": exit_code,
            "source_unchanged": value.get("source_unchanged"),
            "cleanup_status": cleanup_status,
            "cleanup_remaining": cleanup_remaining,
        }
        public.append(record)
        records.append({"path": terminal, "value": value, "summary": record})
    require(not unfinished, f"terminal jobs are not complete: {unfinished}")
    public.sort(key=lambda item: item["path"])
    records.sort(key=lambda item: str(item["path"]))
    return {"count": len(public), "receipts": public}, records


def _public_terminal(summary: Mapping[str, Any]) -> dict[str, Any]:
    return {"count": summary["count"], "receipts": summary["receipts"]}


def _scratch_paths(records: Iterable[Mapping[str, Any]], temp: Path) -> tuple[list[Path], list[Path]]:
    failed: list[Path] = []
    successful: list[Path] = []
    fields = ("root", "run_root", "private_root", "tmpdir", "replay_dir")
    for record in records:
        value = record["value"]
        cleanup = value.get("cleanup")
        paths: list[Path] = []
        clean = False
        if isinstance(cleanup, Mapping):
            for field in fields:
                raw = cleanup.get(field)
                if isinstance(raw, str) and raw.startswith("/"):
                    candidate = Path(raw).resolve(strict=False)
                    if _path_inside(candidate, temp):
                        paths.append(candidate)
            remaining = cleanup.get("remaining")
            if isinstance(remaining, list):
                for raw in remaining:
                    if isinstance(raw, str) and raw.startswith("/"):
                        candidate = Path(raw).resolve(strict=False)
                        if _path_inside(candidate, temp):
                            paths.append(candidate)
            clean = cleanup.get("status") == "pass" and cleanup.get("remaining") == []
        # A failed terminal that names the target or shared scratch directly
        # has no private cleanup receipt to establish custody.  Keep that
        # subtree until the caller archives it.  Phase clones are handled by
        # their authenticated source manifests, so a failed validation cwd in
        # ``before`` or ``after`` does not unnecessarily block source cleanup.
        if not clean and (value.get("status") not in (None, "pass")
                          or value.get("exit_code") not in (None, 0)):
            for field in ("cwd", "private_root", "tmpdir", "replay_dir"):
                raw = value.get(field)
                if isinstance(raw, str) and raw.startswith("/"):
                    candidate = Path(raw).resolve(strict=False)
                    if _path_inside(candidate, temp):
                        relative = candidate.relative_to(temp)
                        first = relative.parts[0] if relative.parts else ""
                        if first in {"target", "tmp", "projections", "runs", "publication-profiles"}:
                            paths.append(candidate)
            environment = value.get("environment")
            if isinstance(environment, Mapping):
                for raw in environment.values():
                    if isinstance(raw, str) and raw.startswith("/"):
                        candidate = Path(raw).resolve(strict=False)
                        if _path_inside(candidate, temp):
                            relative = candidate.relative_to(temp)
                            first = relative.parts[0] if relative.parts else ""
                            if first in {"target", "tmp", "projections", "runs", "publication-profiles"}:
                                paths.append(candidate)
        for path in paths:
            (successful if clean else failed).append(path)
    return failed, successful


def _candidate_paths(temp: Path, target: Path, records: Iterable[Mapping[str, Any]],
                     *, target_present: bool) -> tuple[list[Path], list[dict[str, Any]]]:
    _directory(temp, "temporary root", owner_uid=os.getuid())
    require(target.resolve(strict=False) == (temp / "target").resolve(strict=False),
            "Cargo target is outside the owned 0497 temporary root")
    _no_symlink_components(target, "Cargo target")
    if target_present:
        _directory(target, "Cargo target", owner_uid=os.getuid())
    else:
        require(not target.exists() and not target.is_symlink(),
                "Cargo target presence changed while planning")
    retained = temp / "retained"
    _directory(retained, "retained root", owner_uid=os.getuid())
    unexpected = sorted(child.name for child in temp.iterdir()
                        if child.name != "retained" and child.name not in OWNED_CHILDREN)
    require(not unexpected, f"temporary root contains unowned paths: {unexpected}")
    failed, successful = _scratch_paths(records, temp)
    candidates: list[Path] = []
    preserved: list[dict[str, Any]] = []
    for name in sorted(OWNED_CHILDREN):
        if name == FUZZ_TEMP_NAME:
            continue
        path = temp / name
        if not path.exists() and not path.is_symlink():
            # The target is required while planning; optional scratch roots are
            # simply absent after a previous successful private cleanup.
            if name == "target":
                require(not target_present, f"owned Cargo target is missing: {path}")
            continue
        _tree_stats(path, f"temporary candidate {name}", owner_uid=os.getuid(),
                    root_device=temp.lstat().st_dev)
        matching_failed = [item for item in failed if _path_inside(item, path)]
        if name in {"runs", "publication-profiles"}:
            children = sorted(path.iterdir(), key=lambda item: item.name)
            if not children:
                candidates.append(path)
                continue
            removable: list[Path] = []
            for child in children:
                child_failed = [item for item in matching_failed if _path_inside(item, child)]
                if child_failed:
                    preserved.append({"path": str(child.resolve(strict=False)),
                                      "reason": "failed private cleanup terminal"})
                    continue
                child_state = _tree_stats(child, f"private scratch {child}",
                                          owner_uid=os.getuid(), root_device=path.lstat().st_dev)
                if child_state.files == 0:
                    removable.append(child)
                else:
                    successful_match = any(_path_inside(child, item) or _path_inside(item, child)
                                          for item in successful)
                    if successful_match:
                        removable.append(child)
                    else:
                        preserved.append({"path": str(child.resolve(strict=False)),
                                          "reason": "private scratch has no successful custody"})
            # Once every private child has authenticated custody, remove the
            # owned parent as one candidate too.  Otherwise the children are
            # removed and the empty parent survives the postcondition check.
            if len(removable) == len(children):
                candidates.append(path)
            else:
                candidates.extend(removable)
            continue
        if matching_failed:
            preserved.append({"path": str(path.resolve(strict=False)),
                              "reason": "failed private cleanup terminal"})
            continue
        candidates.append(path)
    if target_present:
        require(target in candidates, "Cargo target was not selected as a cleanup candidate")
    # Every candidate is inventoried here so an accidental special file or
    # mount boundary fails before any destructive operation is attempted.
    for candidate in candidates:
        _tree_stats(candidate, f"cleanup candidate {candidate}", owner_uid=os.getuid(),
                    root_device=temp.lstat().st_dev)
    return candidates, preserved


def _proc_path(raw: str) -> Path | None:
    if not raw or not raw.startswith("/"):
        return None
    if raw.endswith(_DELETED_SUFFIX):
        raw = raw[:-len(_DELETED_SUFFIX)]
    try:
        return Path(raw).resolve(strict=False)
    except OSError:
        return None


def _proc_link(path: Path) -> Path | None:
    try:
        raw = os.readlink(path)
    except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
        return None
    except OSError as error:
        fail(f"cannot inspect process link {path}: {error}")
    return _proc_path(os.fsdecode(raw))


def _proc_bytes(path: Path) -> bytes:
    try:
        return path.read_bytes()
    except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
        return b""
    except OSError as error:
        fail(f"cannot inspect process file {path}: {error}")
    return b""


def _proc_optional_bytes(path: Path) -> tuple[bytes, str | None]:
    """Read a foreign process metadata file without hiding its visibility state."""

    try:
        return path.read_bytes(), None
    except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
        return b"", "vanished"
    except OSError as error:
        return b"", str(error)


def _proc_is_kernel_thread(process: Path) -> bool:
    """Return whether a proc entry is a Linux kernel thread.

    Kernel threads expose an unreadable ``cwd``/``exe`` in many containers,
    but cannot hold a user-space path.  Their status file provides an explicit
    marker, so they can be recorded and excluded without weakening checks for
    ordinary processes.
    """

    status = _proc_bytes(process / "status")
    return any(line.strip() == b"Kthread:\t1" or line.strip() == b"Kthread: 1"
               for line in status.splitlines())


def _foreign_process_observation(process: Path, pid: int, uid: int, comm: str,
                                 evidence_root: Path, temp: Path, target: Path,
                                 fuzz_temp: Path | None) -> dict[str, Any]:
    """Observe only safe metadata when a foreign proc entry hides its links."""

    command_raw, command_error = _proc_optional_bytes(process / "cmdline")
    cgroup_raw, cgroup_error = _proc_optional_bytes(process / "cgroup")
    command = [os.fsdecode(item) for item in command_raw.split(b"\0") if item]
    command_text = " ".join(command)
    cgroup = os.fsdecode(cgroup_raw).strip()
    searchable = "\n".join((command_text, cgroup))
    searchable_paths = [str(temp), str(target)]
    if fuzz_temp is not None:
        searchable_paths.append(str(fuzz_temp))
    workspace_reference = any(token in searchable for token in searchable_paths)
    script_reference = any(
        str(evidence_root / script) in searchable for script in ROOT_PROCESS_SCRIPTS
    ) or str(evidence_root) in searchable
    return {
        "pid": pid,
        "uid": uid,
        "comm": comm,
        # Command-line contents are inspected only transiently.  Keep bounded
        # evidence that the file was checked without persisting service
        # arguments, which can contain credentials or other unrelated data.
        "command_bytes": len(command_raw),
        "command_sha256": hashlib.sha256(command_raw).hexdigest(),
        "cgroup": cgroup or None,
        "command_error": command_error,
        "cgroup_error": cgroup_error,
        "unobservable_references": ["cwd", "exe", "fd"],
        "references_observable": False,
        "script_reference": script_reference,
        "workspace_reference": workspace_reference,
        "explicit_experiment_reference": script_reference or workspace_reference,
    }


def _process_audit(deletion_roots: Iterable[Path], evidence_root: Path,
                   temp: Path, target: Path, *, proc_root: Path = Path("/proc"),
                   self_pid: int | None = None,
                   fuzz_temp: Path | None = None) -> dict[str, Any]:
    """Audit accessible jobs and record the foreign-process visibility limit.

    Same-UID process links remain fail-closed because they may be this
    experiment's drivers.  A foreign process whose proc links are hidden is
    checked through readable command/cgroup metadata for explicit 0497
    references and retained as an unobservable-reference limitation; this
    routine does not claim host-wide quiescence.
    """

    _directory(proc_root, "process information root")
    deletion = tuple(Path(item).resolve(strict=False) for item in deletion_roots)
    self_pid = os.getpid() if self_pid is None else self_pid
    uid_scope = os.getuid()
    busy: list[dict[str, Any]] = []
    drivers: list[dict[str, Any]] = []
    vanished = 0
    scanned = 0
    excluded_other_uid: list[dict[str, Any]] = []
    excluded_daemons: list[dict[str, Any]] = []
    excluded_kernel_threads: list[dict[str, Any]] = []
    foreign_unobservable: list[dict[str, Any]] = []
    try:
        entries = sorted(proc_root.iterdir(), key=lambda item: item.name)
    except OSError as error:
        fail(f"process information root: cannot enumerate: {error}")
    for process in entries:
        if not process.name.isdecimal():
            continue
        pid = int(process.name)
        if pid == self_pid:
            continue
        try:
            state = process.lstat()
        except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
            vanished += 1
            continue
        except OSError as error:
            fail(f"/proc/{pid}: cannot identify process owner: {error}")
        require(stat.S_ISDIR(state.st_mode) and not stat.S_ISLNK(state.st_mode),
                f"/proc/{pid}: process entry is unsafe")
        if state.st_uid != uid_scope:
            # Ownership does not make a process harmless: a different-user
            # cargo, gate, or capture process can still hold a cwd, executable,
            # or fd in this experiment's scratch.  Keep the uid in the audit
            # record and inspect its links as well.  If links are inaccessible,
            # inspect readable command/cgroup metadata before recording the
            # remaining visibility limitation.
            excluded_other_uid.append({"pid": pid, "uid": state.st_uid,
                                       "reason": "audited despite different uid"})
        name = _proc_bytes(process / "comm").decode(errors="replace").strip()
        if _proc_is_kernel_thread(process):
            excluded_kernel_threads.append({"pid": pid, "comm": name,
                                            "reason": "kernel thread cannot hold user-space paths"})
            continue
        try:
            cwd = _proc_link(process / "cwd")
            exe = _proc_link(process / "exe")
            try:
                fd_entries = sorted((process / "fd").iterdir(), key=lambda item: item.name)
            except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
                vanished += 1
                continue
            fd_paths = [(fd.name, _proc_link(fd)) for fd in fd_entries]
            command = [os.fsdecode(item) for item in _proc_bytes(process / "cmdline").split(b"\0") if item]
        except CleanupError as error:
            if state.st_uid != uid_scope:
                observation = _foreign_process_observation(
                    process, pid, state.st_uid, name, evidence_root, temp, target, fuzz_temp
                )
                if observation["explicit_experiment_reference"]:
                    drivers.append(observation)
                else:
                    foreign_unobservable.append(observation)
                continue
            if name in PROTECTED_DAEMONS:
                excluded_daemons.append({"pid": pid, "comm": name})
                continue
            raise error
        scanned += 1
        references: list[tuple[str, Path]] = []
        if cwd is not None:
            references.append(("cwd", cwd))
        if exe is not None:
            references.append(("exe", exe))
        references.extend((f"fd:{fd}", path) for fd, path in fd_paths if path is not None)
        for kind, path in references:
            for candidate in deletion:
                if _path_inside(path, candidate):
                    busy.append({"pid": pid, "reference": kind, "path": str(path),
                                 "candidate": str(candidate)})
                    break
        command_text = " ".join(command)
        script_match = False
        for argument in command:
            if Path(argument).name not in ROOT_PROCESS_SCRIPTS:
                continue
            script = _proc_path(argument) if argument.startswith("/") else (
                _proc_path(str(cwd / argument)) if cwd is not None else None
            )
            if script is not None and _path_inside(script, evidence_root):
                script_match = True
                break
        workspace_match = (str(temp) in command_text or str(target) in command_text
                           or (fuzz_temp is not None and str(fuzz_temp) in command_text))
        executable_match = exe is not None and any(_path_inside(exe, candidate) for candidate in deletion)
        root_executable_match = exe is not None and _path_inside(exe, evidence_root)
        if script_match or workspace_match or executable_match or root_executable_match:
            drivers.append({"pid": pid, "exe": None if exe is None else str(exe),
                            "command": command_text, "script_reference": script_match,
                            "workspace_reference": workspace_match,
                            "executable_reference": executable_match,
                            "root_executable_reference": root_executable_match})
    return {
        "proc_root": str(proc_root.resolve(strict=False)),
        "self_pid": self_pid,
        "uid_scope": uid_scope,
        "scanned_processes": scanned,
        "vanished_processes": vanished,
        "excluded_other_uid_processes": excluded_other_uid,
        "excluded_session_daemons": excluded_daemons,
        "excluded_kernel_threads": excluded_kernel_threads,
        "foreign_unobservable_processes": foreign_unobservable,
        "candidate_references": busy,
        "root_gate_capture_processes": drivers,
        "safe": not busy and not drivers,
    }


def _process_failure(value: Mapping[str, Any]) -> str:
    if value.get("candidate_references"):
        return f"active process references cleanup candidate: {value['candidate_references']}"
    if value.get("root_gate_capture_processes"):
        return f"live 0497 gate/capture process remains: {value['root_gate_capture_processes']}"
    return "process audit failed"


def _removed_stats(candidates: Iterable[Path], temp: Path,
                   fuzz_temp: Path | None = None) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for path in candidates:
        record_root = temp
        if fuzz_temp is not None and _path_inside(path, fuzz_temp):
            record_root = fuzz_temp.parent
        stats = _tree_stats(path, f"cleanup candidate {path}", owner_uid=os.getuid(),
                            root_device=record_root.lstat().st_dev)
        records.append(_stats_record(stats, root=record_root, label="cleanup candidate"))
    return records


def _totals(records: Iterable[Mapping[str, Any]]) -> dict[str, int]:
    values = list(records)
    return {
        "paths": len(values),
        "files": sum(int(value["files"]) for value in values),
        "directories": sum(int(value["directories"]) for value in values),
        "logical_bytes": sum(int(value["logical_bytes"]) for value in values),
        "allocated_bytes": sum(int(value["allocated_bytes"]) for value in values),
    }


def _retained(builds: Mapping[str, Any], fuzz: Mapping[str, Any] | None = None) -> list[dict[str, Any]]:
    values: list[dict[str, Any]] = []
    for phase in PHASES:
        for role in BUILD_ROLES:
            key = f"{phase}/{role}"
            build = builds[key]
            values.append({
                "phase": phase,
                "role": role,
                "path": build["binary"]["path"],
                "bytes": build["binary"]["bytes"],
                "sha256": build["binary"]["sha256"],
                "executable": True,
                "build_receipt": build["receipt"],
            })
    if fuzz is not None:
        values.append({
            "phase": "fuzz",
            "role": "source_backed_tail_append_stream",
            "path": fuzz["binary"]["path"],
            "bytes": fuzz["binary"]["bytes"],
            "sha256": fuzz["binary"]["sha256"],
            "executable": True,
            "build_receipt": fuzz["build"],
            "verify_receipt": fuzz["verify"],
            "run_receipts": [item["run"] for item in fuzz["runs"]],
        })
    return values


def _revalidate_retained(builds: Mapping[str, Any]) -> None:
    for key, build in builds.items():
        path = Path(build["binary"]["path"])
        _validate_descriptor(build["binary"], path, f"{key} retained binary",
                             present=True, executable=True)


def _same_stats(path: Path, expected: Mapping[str, Any], temp: Path, label: str) -> None:
    stats = _tree_stats(path, label, owner_uid=os.getuid(), root_device=int(expected["device"]))
    actual = _stats_record(stats, root=temp, label=label)
    for field in ("kind", "files", "directories", "logical_bytes", "allocated_bytes",
                  "device", "inode", "fingerprint"):
        require(actual.get(field) == expected.get(field), f"{label}: changed after planning ({field})")


def _remove_tree(path: Path, label: str, *, root_device: int) -> None:
    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is forbidden")
    require(stat.S_ISREG(state.st_mode) or stat.S_ISDIR(state.st_mode),
            f"{label}: special path is forbidden")
    require(state.st_uid == os.getuid() and state.st_dev == root_device,
            f"{label}: ownership/device changed")
    if stat.S_ISREG(state.st_mode):
        try:
            path.unlink()
        except OSError as error:
            fail(f"{label}: cannot remove file: {error}")
        return
    try:
        children = sorted(path.iterdir(), key=lambda item: item.name, reverse=True)
    except OSError as error:
        fail(f"{label}: cannot enumerate {path}: {error}")
    for child in children:
        _remove_tree(child, f"{label}/{child.name}", root_device=root_device)
    try:
        path.rmdir()
    except OSError as error:
        fail(f"{label}: cannot remove directory: {error}")


def _absent(path: Path, label: str) -> None:
    _no_symlink_components(path, label)
    require(not path.exists() and not path.is_symlink(), f"{label}: path remains: {path}")


def _validate_scope(root: Path, temp: Path, target: Path,
                    fuzz_temp: Path | None = None) -> tuple[Path, Path, Path, Path | None]:
    evidence = _canonical_directory(root, "evidence root")
    temporary = _canonical_directory(temp, "temporary root", owner_uid=os.getuid())
    build_target = _canonical_directory(target, "Cargo target", owner_uid=os.getuid(),
                                        allow_missing=True)
    fuzz = None
    if fuzz_temp is not None:
        fuzz = _canonical_directory(fuzz_temp, "fuzz temporary root", owner_uid=os.getuid())
    _assert_disjoint(evidence, temporary, build_target, fuzz)
    return evidence, temporary, build_target, fuzz


def _validate_builds(root: Path, temp: Path, target: Path,
                     build_receipts: Mapping[str, Path], *, target_present: bool) -> dict[str, Any]:
    values: dict[str, Any] = {}
    phase_builds: dict[str, Any] = {}
    for phase in PHASES:
        phase_build = _validate_build_phase(phase, root, temp, target,
                                            build_receipts[phase], target_present=target_present)
        phase_builds[phase] = phase_build
        values.update(phase_build["records"])
    require(set(values) == {f"{phase}/{role}" for phase in PHASES for role in BUILD_ROLES},
            "build receipt roles are incomplete")
    _validate_retained_layout(temp, values)
    return values


def _fuzz_meta(value: Any, path: Path, label: str, *, present: bool) -> dict[str, Any]:
    """Authenticate a fuzz receipt's compact ``bytes``/``sha256`` binding."""
    require(isinstance(value, Mapping), f"{label}: metadata is malformed")
    size = value.get("bytes")
    digest = _valid_digest(value.get("sha256"), f"{label}.sha256")
    require(type(size) is int and size >= 0, f"{label}.bytes is malformed")
    recorded = value.get("path")
    if recorded is not None:
        require(isinstance(recorded, str) and Path(recorded).resolve(strict=False)
                == path.resolve(strict=False), f"{label}: path binding differs")
    if present:
        state = _regular(path, label)
        assert state is not None
        require(state.st_size == size and _sha256(path) == digest,
                f"{label}: content changed")
    return {"path": str(path.resolve(strict=False)), "bytes": size, "sha256": digest}


def _fuzz_inventory(directory: Path, expected: Any, label: str, *, present: bool) -> dict[str, Any]:
    require(isinstance(expected, Mapping), f"{label}: inventory is malformed")
    normalized: dict[str, Any] = {}
    if not present:
        for raw_name, raw_meta in expected.items():
            name = _manifest_key(raw_name, f"{label} entry")
            require(isinstance(raw_meta, Mapping), f"{label} {name}: metadata is malformed")
            size = raw_meta.get("bytes")
            digest = _valid_digest(raw_meta.get("sha256"), f"{label} {name}.sha256")
            require(type(size) is int and size >= 0, f"{label} {name}.bytes is malformed")
            normalized[name.as_posix()] = {"bytes": size, "sha256": digest}
        return normalized
    _directory(directory, label)
    actual_names: set[str] = set()
    for path in sorted(directory.rglob("*"), key=lambda item: str(item)):
        if path.is_symlink():
            fail(f"{label}: symlink is forbidden: {path}")
        if path.is_file():
            name = path.relative_to(directory).as_posix()
            actual_names.add(name)
            normalized[name] = {"bytes": path.stat().st_size, "sha256": _sha256(path)}
        elif not path.is_dir():
            fail(f"{label}: special path is forbidden: {path}")
    expected_normalized = _fuzz_inventory(directory, expected, label, present=False)
    require(normalized == expected_normalized, f"{label}: inventory differs")
    return normalized


def _validate_fuzz_receipts(root: Path, temp: Path, *, present: bool,
                            origin_present: bool | None = None,
                            retained_root: Path | None = None) -> dict[str, Any]:
    """Authenticate the ASan build, both run terminals, and evidence copies."""
    if origin_present is None:
        origin_present = present
    evidence = root / FUZZ_EVIDENCE_NAME
    _directory(evidence, "fuzz evidence root")
    driver_path = (root / "fuzz.py").resolve(strict=False)
    _regular(driver_path, "fuzz driver")
    prepared_path = evidence / "prepared.json"
    build_path = evidence / "build.json"
    verify_path = evidence / "verify.json"
    prepared = _read_json(prepared_path, "fuzz prepared receipt")
    require(isinstance(prepared, Mapping)
            and prepared.get("schema") == "docx-stream-fuzz-prepared-0497-v1",
            "fuzz prepared receipt schema differs")
    _fuzz_meta(prepared.get("driver"), driver_path, "fuzz prepared driver", present=True)
    source_binding = prepared.get("source_binding")
    require(isinstance(source_binding, Mapping), "fuzz prepared source binding is malformed")

    paths = prepared.get("paths")
    require(isinstance(paths, Mapping), "fuzz prepared paths are malformed")
    expected_paths = {
        "temp_root": temp,
        "package": temp / "fuzz",
        "cargo_target": temp / "target",
        "run_root": temp / "runs",
        "tmpdir": temp / "tmp",
    }
    for key, expected in expected_paths.items():
        require(paths.get(key) == str(expected.resolve(strict=False)),
                f"fuzz prepared {key} path differs")
    require(prepared.get("target") == FUZZ_BINARY_NAME,
            "fuzz prepared target differs")
    require(prepared.get("run_seeds") == list(FUZZ_RUN_SEEDS),
            "fuzz prepared run seeds differ")

    evidence_paths = prepared.get("evidence_paths")
    evidence_files = prepared.get("evidence_files")
    require(isinstance(evidence_paths, Mapping) and isinstance(evidence_files, Mapping),
            "fuzz evidence-side bindings are malformed")
    for key, filename in (("target_source", "target-source.rs"),
                          ("manifest", "Cargo.toml"), ("lock", "Cargo.lock")):
        path = evidence / filename
        require(evidence_paths.get(key) == str(path),
                f"fuzz evidence {key} path differs")
        _fuzz_meta(evidence_files.get(key), path, f"fuzz evidence {key}", present=True)

    target_source = prepared.get("target_source")
    manifest = prepared.get("manifest")
    lock = prepared.get("lock")
    require(isinstance(target_source, Mapping) and isinstance(manifest, Mapping)
            and isinstance(lock, Mapping), "fuzz prepared source artifacts are malformed")
    package = temp / "fuzz"
    package_paths = {
        "target_source": package / f"{FUZZ_BINARY_NAME}.rs",
        "manifest": package / "Cargo.toml",
        "lock": package / "Cargo.lock",
    }
    if present:
        _directory(temp, "fuzz temporary root", owner_uid=os.getuid())
        _directory(package, "fuzz package", owner_uid=os.getuid())
        _fuzz_meta(target_source.get("copied"), package_paths["target_source"],
                   "fuzz copied target", present=True)
        _fuzz_meta(manifest.get("package"), package_paths["manifest"],
                   "fuzz package manifest", present=True)
        _fuzz_meta(lock.get("package"), package_paths["lock"],
                   "fuzz package lock", present=True)
        _fuzz_inventory(package / "corpus-start", prepared.get("corpus_start"),
                        "fuzz starting corpus", present=True)
    else:
        for key, path in package_paths.items():
            record = target_source.get("copied") if key == "target_source" else (
                manifest.get("package") if key == "manifest" else lock.get("package")
            )
            _fuzz_meta(record, path, f"fuzz package {key}", present=False)
        _fuzz_inventory(package / "corpus-start", prepared.get("corpus_start"),
                        "fuzz starting corpus", present=False)

    retained_0485 = prepared.get("retained_0485")
    retained_evidence = evidence / "retained-0485"
    require(isinstance(retained_0485, Mapping), "fuzz retained-0485 binding is malformed")
    _directory(retained_evidence, "fuzz evidence retained-0485")
    for field, relative in (("driver", "fuzz-stream.py"),
                            ("generator", "generator.json"),
                            ("manifest", "seed-manifest.json")):
        _fuzz_meta(retained_0485.get(field), retained_evidence / relative,
                   f"fuzz retained-0485 {field}", present=True)
    seed_inventory = retained_0485.get("seed_inventory")
    _fuzz_inventory(retained_evidence / "seeds", seed_inventory,
                    "fuzz retained-0485 seeds", present=True)

    build = _read_json(build_path, "fuzz build receipt")
    require(isinstance(build, Mapping)
            and build.get("schema") == "docx-stream-fuzz-build-0497-v1",
            "fuzz build receipt schema differs")
    _fuzz_meta(build.get("driver"), driver_path, "fuzz build driver", present=True)
    require(build.get("source_binding_before") == source_binding
            and build.get("source_binding_after") == source_binding,
            "fuzz build source bindings differ from preparation")
    require(build.get("prepared_sha256") == _sha256(prepared_path),
            "fuzz build prepared binding differs")
    binary = build.get("binary")
    require(isinstance(binary, Mapping), "fuzz build binary binding is malformed")
    origin = temp / "target" / FUZZ_TARGET_TRIPLE / "release" / FUZZ_BINARY_NAME
    retained_base = temp.parent / "retained" if retained_root is None else retained_root
    retained = retained_base / "fuzz" / FUZZ_BINARY_NAME
    require(binary.get("origin") == str(origin), "fuzz build origin path differs")
    require(binary.get("retained") == str(retained), "fuzz build retained path differs")
    retained_meta = _fuzz_meta(binary, retained, "fuzz retained binary", present=True)
    _fuzz_meta({"path": str(origin), "bytes": binary.get("bytes"),
                "sha256": binary.get("sha256")}, origin,
               "fuzz origin binary", present=origin_present)

    verify = _read_json(verify_path, "fuzz verification receipt")
    require(isinstance(verify, Mapping)
            and verify.get("schema") == "docx-stream-fuzz-verify-0497-v1"
            and verify.get("passed") is True,
            "fuzz verification did not pass")
    require(verify.get("prepared_sha256") == _sha256(prepared_path)
            and verify.get("build_sha256") == _sha256(build_path),
            "fuzz verification evidence bindings differ")
    require(verify.get("source_binding") == source_binding,
            "fuzz verification source binding differs")
    require(verify.get("run_seeds") == list(FUZZ_RUN_SEEDS),
            "fuzz verification run seeds differ")
    verify_binary = verify.get("binary")
    require(isinstance(verify_binary, Mapping)
            and verify_binary.get("path") == str(retained),
            "fuzz verification binary path differs")
    _fuzz_meta(verify_binary, retained, "fuzz verification binary", present=True)

    runs: list[dict[str, Any]] = []
    for seed in FUZZ_RUN_SEEDS:
        label = f"run-{seed}"
        run_path = evidence / label / "run.json"
        run = _read_json(run_path, f"fuzz {label} receipt")
        require(isinstance(run, Mapping)
                and run.get("schema") == "docx-stream-fuzz-run-0497-v1"
                and run.get("label") == label and run.get("seed") == seed
                and run.get("exit_code") == 0,
                f"fuzz {label} terminal did not pass")
        require(run.get("source_binding_before") == source_binding
                and run.get("source_binding_after") == source_binding
                and run.get("source_binding_before") == run.get("source_binding_after"),
                f"fuzz {label} source bindings differ")
        run_binary = run.get("binary")
        require(isinstance(run_binary, Mapping)
                and run_binary.get("path") == str(retained),
                f"fuzz {label} binary path differs")
        _fuzz_meta(run_binary, retained, f"fuzz {label} binary", present=True)
        terminal = run.get("terminal")
        require(isinstance(terminal, Mapping), f"fuzz {label} terminal binding is malformed")
        for field in ("stdout", "stderr"):
            terminal_path = evidence / label / f"terminal.{field}"
            require(terminal.get(field) == str(terminal_path),
                    f"fuzz {label} terminal {field} path differs")
            _fuzz_meta(terminal.get(f"{field}_meta"), terminal_path,
                       f"fuzz {label} terminal {field}", present=True)
        retained_paths = run.get("retained")
        require(isinstance(retained_paths, Mapping),
                f"fuzz {label} retained evidence binding is malformed")
        starting = evidence / label / "starting-corpus"
        post = evidence / label / "post-corpus"
        artifacts = evidence / label / "artifacts"
        require(retained_paths.get("starting_corpus") == str(starting)
                and retained_paths.get("post_corpus") == str(post)
                and retained_paths.get("artifacts") == str(artifacts),
                f"fuzz {label} retained evidence paths differ")
        starting_inventory = _fuzz_inventory(starting, run.get("corpus_before"),
                                              f"fuzz {label} starting corpus", present=True)
        post_inventory = _fuzz_inventory(post, run.get("corpus_after"),
                                          f"fuzz {label} post corpus", present=True)
        artifact_inventory = _fuzz_inventory(artifacts,
                                             retained_paths.get("artifacts_inventory"),
                                             f"fuzz {label} artifacts", present=True)
        require(retained_paths.get("post_corpus_inventory") == post_inventory
                and run.get("corpus_before") == starting_inventory
                and run.get("corpus_after") == post_inventory,
                f"fuzz {label} evidence inventories differ")
        runs.append({"label": label, "seed": seed, "exit_code": 0,
                     "run": _descriptor(run_path, f"fuzz {label} receipt", root=root),
                     "starting_files": len(starting_inventory),
                     "post_files": len(post_inventory),
                     "artifact_files": len(artifact_inventory)})

    verify_runs = verify.get("runs")
    require(isinstance(verify_runs, list), "fuzz verification run inventory is malformed")
    expected_run_keys = {(item["label"], item["seed"]) for item in runs}
    actual_run_keys = {(item.get("label"), item.get("seed")) for item in verify_runs
                       if isinstance(item, Mapping)}
    require(actual_run_keys == expected_run_keys,
            "fuzz verification run inventory differs")
    return {
        "evidence": str(evidence.resolve(strict=False)),
        "prepared": _descriptor(prepared_path, "fuzz prepared receipt", root=root),
        "build": _descriptor(build_path, "fuzz build receipt", root=root),
        "verify": _descriptor(verify_path, "fuzz verification receipt", root=root),
        "runs": runs,
        "source_binding": source_binding,
        "binary": retained_meta | {"executable": True},
    }


def _build_receipt_summary(builds: Mapping[str, Any]) -> dict[str, Any]:
    return {
        phase: {
            "path": builds[f"{phase}/normal"]["receipt"]["path"],
            "bytes": builds[f"{phase}/normal"]["receipt"]["bytes"],
            "sha256": builds[f"{phase}/normal"]["receipt"]["sha256"],
        }
        for phase in PHASES
    }


def _revalidate_plan(plan: Mapping[str, Any]) -> dict[str, Any]:
    root, temp, target, fuzz_temp = _validate_scope(
        Path(str(plan["root"])), Path(str(plan["temp"])), Path(str(plan["target"])),
        Path(str(plan["fuzz_temp"])),
    )
    receipt_path = root / SEAL_NAME
    require(not receipt_path.exists() and not receipt_path.is_symlink(),
            f"refusing to replace existing cleanup receipt: {receipt_path}")
    build_receipts = {phase: Path(str(plan["build_receipts"][phase])) for phase in PHASES}
    builds = _validate_builds(root, temp, target, build_receipts, target_present=True)
    target_present = target.exists()
    require(target_present is bool(plan.get("target_present")),
            "Cargo target presence changed after planning")
    early_target = _early_target_binding(root, target, required=not target_present)
    require(early_target == plan.get("early_target_cleanup"),
            "early target cleanup binding changed after planning")
    fuzz_origin = fuzz_temp / "target"
    fuzz_origin_present = fuzz_origin.exists()
    require(fuzz_origin_present is bool(plan.get("fuzz_target_present")),
            "fuzz Cargo target presence changed after planning")
    early_fuzz = _early_fuzz_binding(root, fuzz_temp, required=not fuzz_origin_present)
    require(early_fuzz == plan.get("early_fuzz_cleanup"),
            "early fuzz target cleanup binding changed after planning")
    fuzz = _validate_fuzz_receipts(root, fuzz_temp, present=True,
                                   origin_present=fuzz_origin_present,
                                   retained_root=temp / "retained")
    terminal, terminal_records = _terminal_summary(root)
    require(_build_receipt_summary(builds) == _build_receipt_summary(plan["builds"]),
            "build receipt inventory changed after planning")
    require(fuzz == plan["fuzz"], "fuzz evidence inventory changed after planning")
    require(_public_terminal(terminal) == _public_terminal(plan["terminal"]),
            "terminal evidence changed after planning")
    main_candidates, preserved = _candidate_paths(temp, target, terminal_records,
                                                  target_present=target_present)
    candidates = [*main_candidates, fuzz_temp]
    expected_candidates = {str(Path(str(path)).resolve(strict=False)) for path in plan["candidates"]}
    require({str(path.resolve(strict=False)) for path in candidates} == expected_candidates,
            "cleanup candidate set changed")
    expected_preserved = {str(item["path"]) for item in plan.get("preserved_failed", [])}
    require({str(item["path"]) for item in preserved} == expected_preserved,
            "preserved failed-artifact set changed")
    return {"root": root, "temp": temp, "target": target, "builds": builds,
            "fuzz_temp": fuzz_temp, "fuzz": fuzz,
            "target_present": target_present,
            "fuzz_target_present": fuzz_origin_present,
            "early_target_cleanup": early_target,
            "early_fuzz_cleanup": early_fuzz,
            "terminal": terminal, "terminal_records": terminal_records,
            "candidates": candidates, "preserved_failed": preserved,
            "build_receipts": build_receipts}


def _write_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to replace receipt: {path}")
    try:
        with path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
    except OSError as error:
        fail(f"cannot write receipt {path}: {error}")


def plan_cleanup(*, root: Path = ROOT, temp: Path = TEMP, target: Path = TARGET,
                 fuzz_temp: Path = FUZZ_TEMP,
                 proc_root: Path = Path("/proc"), self_pid: int | None = None,
                 build_receipts: Mapping[str, Any] | None = None,
                 before_build: Path | None = None, after_build: Path | None = None) -> dict[str, Any]:
    evidence, temporary, build_target, fuzz_temporary = _validate_scope(
        Path(root), Path(temp), Path(target), Path(fuzz_temp)
    )
    receipt_path = evidence / SEAL_NAME
    require(not receipt_path.exists() and not receipt_path.is_symlink(),
            f"refusing to replace existing cleanup receipt: {receipt_path}")
    receipts = _coerce_build_receipts(build_receipts, before_build, after_build)
    builds = _validate_builds(evidence, temporary, build_target, receipts, target_present=True)
    target_present = build_target.exists()
    early_target = _early_target_binding(evidence, build_target, required=not target_present)
    fuzz_origin = fuzz_temporary / "target"
    fuzz_origin_present = fuzz_origin.exists()
    early_fuzz = _early_fuzz_binding(evidence, fuzz_temporary, required=not fuzz_origin_present)
    fuzz = _validate_fuzz_receipts(evidence, fuzz_temporary, present=True,
                                   origin_present=fuzz_origin_present,
                                   retained_root=temporary / "retained")
    terminal, terminal_records = _terminal_summary(evidence)
    candidates, preserved = _candidate_paths(temporary, build_target, terminal_records,
                                              target_present=target_present)
    fuzz_candidate = fuzz_temporary
    removed = _removed_stats([*candidates, fuzz_candidate], temporary, fuzz_temporary)
    deletion_roots = [*candidates, fuzz_candidate]
    process = _process_audit(deletion_roots, evidence, temporary, build_target,
                             proc_root=proc_root, self_pid=self_pid,
                             fuzz_temp=fuzz_temporary)
    require(process["safe"], _process_failure(process))
    return {
        "root": evidence,
        "temp": temporary,
        "target": build_target,
        "fuzz_temp": fuzz_temporary,
        "fuzz": fuzz,
        "build_receipts": receipts,
        "builds": builds,
        "terminal": terminal,
        "terminal_records": terminal_records,
        "candidates": [*candidates, fuzz_candidate],
        "target_candidate": build_target if target_present else None,
        "target_present": target_present,
        "fuzz_target_present": fuzz_origin_present,
        "early_target_cleanup": early_target,
        "early_fuzz_cleanup": early_fuzz,
        "preserved_failed": preserved,
        "removed_stats": removed,
        "process_before": process,
        "proc_root": Path(proc_root),
        "self_pid": os.getpid() if self_pid is None else self_pid,
        "disk_before": _disk(temporary),
    }


def _plan_cleanup(**kwargs: Any) -> dict[str, Any]:
    """Compatibility entry point used by the other bounded cleanup drivers."""
    return plan_cleanup(**kwargs)


def execute_cleanup(plan: Mapping[str, Any]) -> dict[str, Any]:
    authenticated = _revalidate_plan(plan)
    require(not authenticated["preserved_failed"],
            "failed scratch artifacts require archived custody before cleanup: "
            f"{authenticated['preserved_failed']}")
    root = authenticated["root"]
    temp = authenticated["temp"]
    target = authenticated["target"]
    fuzz_temp = authenticated["fuzz_temp"]
    receipt_path = root / SEAL_NAME
    before = _process_audit(authenticated["candidates"], root, temp, target,
                            proc_root=Path(str(plan["proc_root"])),
                            self_pid=int(plan["self_pid"]), fuzz_temp=fuzz_temp)
    require(before["safe"], _process_failure(before))
    expected_by_path = {str(Path(item["path"]).resolve(strict=False)): item
                        for item in plan["removed_stats"]}
    require(len(expected_by_path) == len(plan["removed_stats"]),
            "cleanup inventory has duplicate paths")
    require({str(path.resolve(strict=False)) for path in authenticated["candidates"]}
            == set(expected_by_path), "cleanup inventory changed")
    for path in authenticated["candidates"]:
        stats_root = fuzz_temp.parent if _path_inside(path, fuzz_temp) else temp
        _same_stats(path, expected_by_path[str(path.resolve(strict=False))], stats_root,
                    f"cleanup candidate {path}")
    for path in authenticated["candidates"]:
        record = expected_by_path[str(path.resolve(strict=False))]
        _remove_tree(path, f"remove {path}", root_device=int(record["device"]))
        _absent(path, f"removed cleanup candidate {path}")
    _revalidate_retained(authenticated["builds"])
    _validate_fuzz_receipts(root, fuzz_temp, present=False, origin_present=False,
                            retained_root=temp / "retained")
    require({child.name for child in temp.iterdir()} == {"retained"},
            "unexpected temporary paths remain after cleanup")
    _absent(target, "Cargo target after cleanup")
    _absent(fuzz_temp, "removed fuzz temporary root")
    after = _process_audit(authenticated["candidates"], root, temp, target,
                           proc_root=Path(str(plan["proc_root"])),
                           self_pid=int(plan["self_pid"]), fuzz_temp=fuzz_temp)
    require(after["safe"], _process_failure(after))
    terminal, _ = _terminal_summary(root)
    receipt = {
        "schema": SCHEMA,
        "version": VERSION,
        "status": "pass",
        "completed_utc": _now(),
        "driver": _descriptor(Path(__file__).resolve(), "cleanup driver"),
        "scope": {
            "evidence_root": str(root),
            "temporary_root": str(temp),
            "target_root": str(target),
            "fuzz_temporary_root": str(fuzz_temp),
            "retained_subtree": str((temp / "retained").resolve(strict=True)),
            "retained_binaries": 5,
            "owned_cleanup_children": sorted(OWNED_CHILDREN),
            "root_evidence_preserved": True,
            "protected_worktree_touched": False,
        },
        "build_receipts": _build_receipt_summary(authenticated["builds"]),
        "fuzz": authenticated["fuzz"],
        "target_cleanup": {
            "mode": "final" if authenticated["target_present"] else "early",
            "early_receipt": authenticated["early_target_cleanup"],
        },
        "fuzz_target_cleanup": {
            "mode": "final" if authenticated["fuzz_target_present"] else "early",
            "early_receipt": authenticated["early_fuzz_cleanup"],
        },
        "terminal_validation": _public_terminal(terminal),
        "retained_binaries": _retained(authenticated["builds"], authenticated["fuzz"]),
        "removed": list(plan["removed_stats"]),
        "removed_paths": list(plan["removed_stats"]),
        "removed_totals": _totals(plan["removed_stats"]),
        "preserved_failed": [],
        "temporary_scratch_remaining": ["retained"],
        "remaining": [],
        "disk": {"before": plan.get("disk_before", _disk(temp)), "after": _disk(temp)},
        "process_audit": {"before": before, "after": after},
    }
    _write_exclusive(receipt_path, receipt)
    return receipt


def _verify_removed(records: Any, *, temp: Path, target: Path, fuzz_temp: Path,
                    target_precleaned: bool = False) -> None:
    require(isinstance(records, list) and records, "cleanup removed inventory is missing")
    seen: set[str] = set()
    target_seen = False
    fuzz_seen = False
    for index, record in enumerate(records):
        require(isinstance(record, Mapping), f"cleanup.removed_paths[{index}] is malformed")
        raw = record.get("path")
        require(isinstance(raw, str) and raw, f"cleanup.removed_paths[{index}].path is missing")
        path = Path(raw)
        key = str(path.resolve(strict=False))
        require(key not in seen, f"cleanup.removed_paths[{index}] is duplicated")
        seen.add(key)
        if key == str(target.resolve(strict=False)):
            target_seen = True
        elif key == str(fuzz_temp.resolve(strict=False)):
            fuzz_seen = True
        else:
            require((_path_inside(path, temp) and not _path_inside(path, temp / "retained"))
                    or _path_inside(path, fuzz_temp),
                    f"cleanup.removed_paths[{index}] escapes owned scratch")
        _absent(path, f"cleanup.removed_paths[{index}]")
        for field in ("files", "directories", "logical_bytes", "allocated_bytes",
                      "device", "inode"):
            require(type(record.get(field)) is int and record[field] >= 0,
                    f"cleanup.removed_paths[{index}].{field} is malformed")
        digest = record.get("fingerprint")
        require(isinstance(digest, str) and len(digest) == 64
                and all(char in "0123456789abcdef" for char in digest),
                f"cleanup.removed_paths[{index}].fingerprint is malformed")
    if target_precleaned:
        require(not target_seen, "cleanup removed inventory unexpectedly includes pre-cleaned target")
    else:
        require(target_seen, "cleanup removed inventory omits Cargo target")
    require(fuzz_seen, "cleanup removed inventory omits fuzz temporary root")


def _validate_recovery_provenance(receipt: Mapping[str, Any], *, root: Path,
                                  temp: Path, target: Path, fuzz_temp: Path) -> None:
    """Authenticate the one bounded partial-cleanup recovery shape.

    A failed postcondition may leave only an empty owned parent after its
    authenticated children have been removed.  Recovery is accepted only when
    the immutable attempt receipt records the same inventory and the observed
    pre-repair state; ordinary cleanup receipts omit this optional binding.
    """

    recovery = receipt.get("recovery")
    if recovery is None:
        return
    require(isinstance(recovery, Mapping)
            and set(recovery) == {"partial_attempt", "postcondition_repaired", "note"},
            "cleanup recovery provenance is malformed")
    require(recovery.get("postcondition_repaired")
            == "removed authenticated empty publication-profiles parent with rmdir",
            "cleanup recovery repair is not the bounded empty-parent repair")
    require(isinstance(recovery.get("note"), str) and recovery["note"].strip(),
            "cleanup recovery note is missing")
    attempt_path = root / "cleanup-attempt1.json"
    expected_descriptor = _descriptor(attempt_path, "cleanup partial attempt", root=root)
    require(recovery.get("partial_attempt") == expected_descriptor,
            "cleanup recovery attempt descriptor differs")
    attempt = _read_json(attempt_path, "cleanup partial attempt")
    require(isinstance(attempt, Mapping)
            and attempt.get("schema") == "docx-tail-append-0497-cleanup-attempt-v1"
            and attempt.get("version") == 1 and attempt.get("status") == "partial",
            "cleanup partial attempt schema/status differs")
    require(attempt.get("exit_code") == 1
            and attempt.get("stderr") == "cleanup.py: error: unexpected temporary paths remain after cleanup\n",
            "cleanup partial attempt failure differs")
    scope = attempt.get("scope")
    require(scope == {
        "evidence_root": str(root),
        "temporary_root": str(temp),
        "target_root": str(target),
        "fuzz_temporary_root": str(fuzz_temp),
    }, "cleanup partial attempt scope differs")
    plan = attempt.get("plan")
    require(isinstance(plan, Mapping)
            and plan.get("removed_paths") == receipt.get("removed_paths")
            and plan.get("removed_totals") == receipt.get("removed_totals"),
            "cleanup recovery inventory differs from the partial attempt")
    postcondition = attempt.get("postcondition")
    require(postcondition == {
        "cleanup_receipt_written": False,
        "error": "unexpected temporary paths remain after cleanup",
        "expected_temporary_children": ["retained"],
        "fuzz_temporary_present": False,
        "observed_publication_profiles_children": [],
        "observed_temporary_children": ["publication-profiles", "retained"],
        "target_present": False,
    }, "cleanup partial attempt postcondition differs")
    require(attempt.get("removed_paths_observed") == receipt.get("removed_paths"),
            "cleanup partial attempt removed inventory differs")


def verify(*, root: Path = ROOT, temp: Path = TEMP, target: Path = TARGET,
           fuzz_temp: Path = FUZZ_TEMP,
           proc_root: Path = Path("/proc"),
           build_receipts: Mapping[str, Any] | None = None,
           before_build: Path | None = None, after_build: Path | None = None) -> dict[str, Any]:
    evidence = _canonical_directory(Path(root), "evidence root")
    temporary = _canonical_directory(Path(temp), "temporary root", owner_uid=os.getuid())
    target_path = Path(target)
    fuzz_path = Path(fuzz_temp)
    _no_symlink_components(target_path, "Cargo target")
    _no_symlink_components(fuzz_path, "fuzz temporary root")
    _assert_disjoint(evidence, temporary, target_path.resolve(strict=False),
                     fuzz_path.resolve(strict=False))
    _absent(target_path, "Cargo target")
    _absent(fuzz_path, "fuzz temporary root")
    receipt_path = evidence / SEAL_NAME
    receipt = _read_json(receipt_path, "cleanup receipt")
    require(isinstance(receipt, Mapping) and receipt.get("schema") == SCHEMA
            and receipt.get("version") == VERSION and receipt.get("status") == "pass",
            "cleanup receipt schema/status differs")
    require(receipt.get("scope", {}).get("evidence_root") == str(evidence)
            and receipt.get("scope", {}).get("temporary_root") == str(temporary)
            and receipt.get("scope", {}).get("target_root") == str(target_path.resolve(strict=False))
            and receipt.get("scope", {}).get("fuzz_temporary_root") == str(fuzz_path.resolve(strict=False)),
            "cleanup scope binding differs")
    receipts = _coerce_build_receipts(build_receipts, before_build, after_build)
    builds = _validate_builds(evidence, temporary, target_path, receipts, target_present=False)
    fuzz = _validate_fuzz_receipts(evidence, fuzz_path, present=False, origin_present=False,
                                   retained_root=temporary / "retained")
    _validate_retained_layout(temporary, builds)
    require({child.name for child in temporary.iterdir()} == {"retained"},
            "temporary root retains disposable paths")
    require(receipt.get("build_receipts") == _build_receipt_summary(builds),
            "build receipt inventory changed")
    terminal, _ = _terminal_summary(evidence)
    require(receipt.get("terminal_validation") == _public_terminal(terminal),
            "terminal validation changed")
    require(receipt.get("fuzz") == fuzz,
            "fuzz evidence inventory changed")
    require(receipt.get("retained_binaries") == _retained(builds, fuzz),
            "retained binary inventory changed")
    target_cleanup = receipt.get("target_cleanup")
    require(isinstance(target_cleanup, Mapping)
            and set(target_cleanup) == {"mode", "early_receipt"},
            "cleanup target custody binding is missing")
    target_mode = target_cleanup.get("mode")
    require(target_mode in {"early", "final"}, "cleanup target custody mode differs")
    if target_mode == "early":
        expected_target_cleanup = _early_target_binding(evidence, target_path, required=True)
        require(not target_path.exists() and target_cleanup.get("early_receipt") == expected_target_cleanup,
                "cleanup early target custody differs")
    else:
        require(target_cleanup.get("early_receipt") is None,
                "cleanup final target custody unexpectedly has an early receipt")

    fuzz_target_cleanup = receipt.get("fuzz_target_cleanup")
    require(isinstance(fuzz_target_cleanup, Mapping)
            and set(fuzz_target_cleanup) == {"mode", "early_receipt"},
            "cleanup fuzz target custody binding is missing")
    fuzz_target_mode = fuzz_target_cleanup.get("mode")
    require(fuzz_target_mode in {"early", "final"},
            "cleanup fuzz target custody mode differs")
    if fuzz_target_mode == "early":
        expected_fuzz_cleanup = _early_fuzz_binding(evidence, fuzz_path, required=True)
        require(fuzz_target_cleanup.get("early_receipt") == expected_fuzz_cleanup,
                "cleanup early fuzz target custody differs")
    else:
        require(fuzz_target_cleanup.get("early_receipt") is None,
                "cleanup final fuzz target custody unexpectedly has an early receipt")

    _verify_removed(receipt.get("removed_paths"), temp=temporary, target=target_path,
                    fuzz_temp=fuzz_path, target_precleaned=target_mode == "early")
    _validate_recovery_provenance(receipt, root=evidence, temp=temporary,
                                  target=target_path, fuzz_temp=fuzz_path)
    require(receipt.get("removed") == receipt.get("removed_paths"),
            "removed inventory aliases differ")
    require(receipt.get("removed_totals") == _totals(receipt["removed_paths"]),
            "removed totals changed")
    require(receipt.get("temporary_scratch_remaining") == ["retained"]
            and receipt.get("remaining") == []
            and receipt.get("preserved_failed") == [],
            "cleanup remaining binding differs")
    require(receipt.get("scope") == {
        "evidence_root": str(evidence),
        "temporary_root": str(temporary),
        "target_root": str(target_path.resolve(strict=False)),
        "fuzz_temporary_root": str(fuzz_path.resolve(strict=False)),
        "retained_subtree": str((temporary / "retained").resolve(strict=True)),
        "retained_binaries": 5,
        "owned_cleanup_children": sorted(OWNED_CHILDREN),
        "root_evidence_preserved": True,
        "protected_worktree_touched": False,
    }, "cleanup scope differs")
    require(receipt.get("driver") == _descriptor(Path(__file__).resolve(), "cleanup driver"),
            "cleanup driver changed")
    recorded_process = receipt.get("process_audit")
    require(isinstance(recorded_process, Mapping), "cleanup process audit is missing")
    for phase in ("before", "after"):
        audit = recorded_process.get(phase)
        require(isinstance(audit, Mapping) and audit.get("safe") is True
                and audit.get("candidate_references") == []
                and audit.get("root_gate_capture_processes") == [],
                f"cleanup process audit {phase} was unsafe")
    current = _process_audit([], evidence, temporary, target_path, proc_root=proc_root,
                             fuzz_temp=fuzz_path)
    require(current["safe"], _process_failure(current))
    return {
        "schema": VERIFY_SCHEMA,
        "version": VERSION,
        "status": "pass",
        "cleanup_receipt": _descriptor(receipt_path, "cleanup receipt", root=evidence),
        "retained_binaries": _retained(builds, fuzz),
        "build_target_removed": True,
        "remaining": [],
        "process_audit": current,
    }


def _print_plan(plan: Mapping[str, Any]) -> None:
    print(json.dumps({
        "schema": SCHEMA,
        "status": "ready",
        "dry_run": True,
        "retained_binaries": _retained(plan["builds"], plan["fuzz"]),
        "removed_paths": plan["removed_stats"],
        "removed_totals": _totals(plan["removed_stats"]),
        "target_cleanup": {
            "mode": "final" if plan["target_present"] else "early",
            "early_receipt": plan["early_target_cleanup"],
        },
        "fuzz_target_cleanup": {
            "mode": "final" if plan["fuzz_target_present"] else "early",
            "early_receipt": plan["early_fuzz_cleanup"],
        },
        "preserved_failed": plan["preserved_failed"],
        "process_audit": plan["process_before"],
    }, indent=2, sort_keys=True))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--dry-run", action="store_true")
    modes.add_argument("--verify", action="store_true")
    parser.add_argument("--root", type=Path, default=ROOT, help=argparse.SUPPRESS)
    parser.add_argument("--temp-root", type=Path, default=TEMP, help=argparse.SUPPRESS)
    parser.add_argument("--target-root", type=Path, default=TARGET, help=argparse.SUPPRESS)
    parser.add_argument("--fuzz-temp-root", type=Path, default=FUZZ_TEMP, help=argparse.SUPPRESS)
    parser.add_argument("--proc-root", type=Path, default=Path("/proc"), help=argparse.SUPPRESS)
    parser.add_argument("--before-build", type=Path, default=None, help=argparse.SUPPRESS)
    parser.add_argument("--after-build", type=Path, default=None, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    try:
        configured = (args.root.resolve(), args.temp_root.resolve(), args.target_root.resolve(),
                      args.fuzz_temp_root.resolve())
        fixed = (ROOT.resolve(), TEMP.resolve(), TARGET.resolve(), FUZZ_TEMP.resolve())
        if not args.dry_run and not args.verify:
            require(configured == fixed, "destructive mode is restricted to fixed 0497 roots")
        kwargs = {
            "root": args.root,
            "temp": args.temp_root,
            "target": args.target_root,
            "fuzz_temp": args.fuzz_temp_root,
            "proc_root": args.proc_root,
            "before_build": args.before_build,
            "after_build": args.after_build,
        }
        if args.verify:
            print(json.dumps(verify(**kwargs), indent=2, sort_keys=True))
            return 0
        plan = plan_cleanup(**kwargs)
        if args.dry_run:
            _print_plan(plan)
            return 0
        receipt = execute_cleanup(plan)
        print(json.dumps({"schema": SCHEMA, "status": receipt["status"],
                          "removed_totals": receipt["removed_totals"]}, sort_keys=True))
        return 0
    except (CleanupError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"cleanup.py: error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
