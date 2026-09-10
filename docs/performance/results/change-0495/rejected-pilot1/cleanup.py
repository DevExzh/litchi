#!/usr/bin/env python3
"""Safely clean the disposable 0495 build and capture roots.

The evidence directory is never a cleanup target.  The owned temporary root
keeps exactly the four authenticated before/after executables under one
retained attempt; every other child and the owned Cargo target are removed
only after a read-only inventory, ownership/device checks, and a process audit.
Symlinks, special files, active references, and stale gate/capture processes
fail closed.  ``--dry-run`` and ``--verify`` never mutate a path.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat
import sys
from typing import Any, Iterable, Mapping


sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parent
TEMP = Path("/home/zhuhe/.cache/litchi-goal-0495")
TARGET = Path("/home/zhuhe/.cache/litchi-build-0495")
# Keep the support.py spelling available to callers that use the cleanup
# driver as a library.
TARGET_DIR = TARGET
PROTECTED_WORKTREE = Path("/home/zhuhe/code/litchi-spec-gaps")
SCHEMA = "docx-edit-provider-cleanup-v1"
VERIFY_SCHEMA = "docx-edit-provider-cleanup-verification-v1"
VERSION = 1
SEAL_NAME = "cleanup.json"
BUILD_SCHEMA = "docx-provider-managed-build-v1"
BUILD_ROLES = {
    "before-normal": ("before", "normal", "litchi-perf-baseline"),
    "before-allocator": ("before", "allocator", "litchi-perf-baseline-alloc"),
    "after-normal": ("after", "normal", "litchi-perf-baseline"),
    "after-allocator": ("after", "allocator", "litchi-perf-baseline-alloc"),
}
ROOT_PROCESS_SCRIPTS = frozenset({
    "cleanup.py", "gate.py", "measure.py", "retain_build.py", "verify.py",
    "profile.py", "test_cleanup.py", "test_measure.py", "test_profile.py",
    "test_verify.py",
})
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


def _sha256(path: Path) -> str:
    try:
        with path.open("rb") as stream:
            return hashlib.file_digest(stream, "sha256").hexdigest()
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
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
    if allow_missing and not path.exists():
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
        files = directories = logical = allocated_total = unique_allocated
        # The root directory contributes one directory, but not a logical file.
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
        "kind": stats.kind, "files": stats.files, "directories": stats.directories,
        "logical_bytes": stats.logical_bytes, "allocated_bytes": stats.allocated_bytes,
        "device": stats.identity[0], "inode": stats.identity[1],
        "fingerprint": stats.fingerprint,
    }


def _assert_disjoint(evidence: Path, temp: Path, target: Path) -> None:
    require(not _path_inside(evidence, temp) and not _path_inside(temp, evidence),
            "evidence and temporary roots overlap")
    require(not _path_inside(evidence, target) and not _path_inside(target, evidence),
            "evidence and build roots overlap")
    require(not _path_inside(temp, target) and not _path_inside(target, temp),
            "temporary and build roots overlap")
    protected = PROTECTED_WORKTREE.resolve(strict=False)
    require(not _path_inside(evidence, protected) and not _path_inside(protected, evidence)
            and not _path_inside(temp, protected) and not _path_inside(protected, temp)
            and not _path_inside(target, protected) and not _path_inside(protected, target),
            "protected worktree overlaps cleanup scope")


def _read_json(path: Path, label: str) -> Any:
    _regular(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        fail(f"{label}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def _descriptor(path: Path, label: str, *, root: Path | None = None) -> dict[str, Any]:
    _no_symlink_components(path, label)
    _regular(path, label)
    resolved = path.resolve(strict=True)
    record: dict[str, Any] = {
        "path": str(resolved), "bytes": resolved.stat().st_size,
        "sha256": _sha256(resolved),
    }
    if root is not None:
        record["relative_path"] = _relative(resolved, root, label)
    return record


def _safe_attempt(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and len(value) <= 96
            and all(char.isalnum() or char in "_-" for char in value),
            f"{label}: unsafe attempt name")
    return value


def _safe_descriptor(value: Any, label: str, path: Path, *, present: bool) -> dict[str, Any]:
    require(isinstance(value, Mapping), f"{label}: descriptor missing")
    require(value.get("executable") is True, f"{label}: executable binding failed")
    size, digest = value.get("bytes"), value.get("sha256")
    require(type(size) is int and size >= 0, f"{label}: byte length malformed")
    require(isinstance(digest, str) and len(digest) == 64
            and all(char in "0123456789abcdef" for char in digest),
            f"{label}: SHA-256 malformed")
    _no_symlink_components(path, label)
    if present:
        state = _regular(path, label)
        assert state is not None
        require(state.st_size == size and _sha256(path) == digest,
                f"{label}: content changed")
        require(os.access(path, os.X_OK) and bool(state.st_mode & 0o111),
                f"{label}: executable bit missing")
    return {"path": str(path.resolve(strict=False)), "bytes": size,
            "sha256": digest, "executable": True}


def _inside(path: Path, root: Path, label: str) -> None:
    require(_path_inside(path, root) and path.resolve(strict=False) != root.resolve(strict=False),
            f"{label}: path escapes owned root")


def _validate_build(key: str, root: Path, temp: Path, target: Path, *,
                    target_present: bool, expected_attempt: str | None = None) -> dict[str, Any]:
    phase, role, binary_name = BUILD_ROLES[key]
    path = root / f"build-{phase}-{role}.json"
    value = _read_json(path, f"{key} build receipt")
    required = {
        "attempt", "binary", "command", "copied_utc", "environment", "gate", "phase",
        "original_binary", "role", "schema", "source_after", "source_before",
        "source_unchanged", "version", "git_revision", "retainer_sha256",
    }
    require(isinstance(value, Mapping) and set(value) == required,
            f"{key} build receipt fields differ")
    require(value["schema"] == BUILD_SCHEMA and value["version"] == VERSION
            and value["phase"] == phase and value["role"] == role
            and value["source_unchanged"] is True
            and value["source_before"] == value["source_after"],
            f"{key} build identity differs")
    attempt = _safe_attempt(value["attempt"], f"{key} build receipt")
    if expected_attempt is not None:
        require(attempt == expected_attempt, f"{key}: retained attempts differ")
    require(isinstance(value["command"], list) and value["command"],
            f"{key}: build command is missing")
    require(value["command"][-1] == binary_name, f"{key}: build binary name differs")
    retained = temp / "retained" / attempt / phase / role / binary_name
    binary_path = Path(str(value["binary"].get("path"))) if isinstance(value["binary"], Mapping) else Path(".")
    require(binary_path.resolve(strict=False) == retained.resolve(strict=False),
            f"{key}: retained binary path differs")
    binary = _safe_descriptor(value["binary"], f"{key} retained binary", retained, present=True)
    original_path = Path(str(value["original_binary"].get("path"))) if isinstance(value["original_binary"], Mapping) else Path(".")
    _inside(original_path, target, f"{key} original binary")
    # Cargo uses one release path for a role across phases.  The after build
    # can therefore overwrite the before receipt's original path before
    # cleanup starts.  The retained phase/role copy is the immutable artifact
    # we authenticate; the original descriptor remains historical custody and
    # must not be restated against the later Cargo target contents.
    original = _safe_descriptor(value["original_binary"], f"{key} original binary",
                                original_path, present=False)
    require(binary["bytes"] == original["bytes"] and binary["sha256"] == original["sha256"],
            f"{key}: retained/original binaries differ")
    return {
        "key": key, "phase": phase, "role": role, "name": binary_name,
        "attempt": attempt, "path": path, "receipt": _descriptor(path, f"{key} build receipt", root=root),
        "binary": binary, "original_binary": original,
    }


def _validate_retained_layout(temp: Path, builds: Mapping[str, Mapping[str, Any]]) -> str:
    attempts = {str(value["attempt"]) for value in builds.values()}
    require(len(attempts) == 1, "retained build attempts differ")
    attempt = next(iter(attempts))
    retained = temp / "retained"
    _directory(retained, "retained root", owner_uid=os.getuid())
    attempt_root = retained / attempt
    _directory(attempt_root, "retained attempt", owner_uid=os.getuid())
    require({item.name for item in attempt_root.iterdir()} == {"before", "after"},
            "retained attempt has unexpected phases")
    for phase in ("before", "after"):
        phase_root = attempt_root / phase
        _directory(phase_root, f"retained {phase} root", owner_uid=os.getuid())
        require({item.name for item in phase_root.iterdir()} == {"normal", "allocator"},
                f"retained {phase} has unexpected roles")
        for role in ("normal", "allocator"):
            key = f"{phase}-{role}"
            role_root = phase_root / role
            _directory(role_root, f"retained {key} root", owner_uid=os.getuid())
            name = BUILD_ROLES[key][2]
            require({item.name for item in role_root.iterdir()} == {name},
                    f"retained {key} has unexpected children")
            path = role_root / name
            state = _regular(path, f"retained {key} binary", owner_uid=os.getuid())
            assert state is not None
            require(os.access(path, os.X_OK) and bool(state.st_mode & 0o111),
                    f"retained {key} binary is not executable")
            require(Path(builds[key]["binary"]["path"]).resolve(strict=False) == path.resolve(),
                    f"retained {key} path differs")
    return attempt


def _validate_terminal_receipts(root: Path) -> dict[str, Any]:
    validation = root / "validation"
    _directory(validation, "validation directory")
    receipts: list[dict[str, Any]] = []
    unfinished: list[str] = []
    for started in sorted(validation.glob("*.started.json")):
        _regular(started, f"validation start {started.name}")
        terminal = validation / f"{started.name[:-len('.started.json')]}.json"
        if terminal.is_symlink() or not terminal.exists():
            unfinished.append(started.name)
            continue
        _regular(terminal, f"validation terminal {terminal.name}")
        value = _read_json(terminal, f"validation terminal {terminal.name}")
        require(isinstance(value, Mapping) and type(value.get("exit_code")) is int,
                f"validation terminal {terminal.name}: malformed result")
        receipts.append(_descriptor(terminal, f"validation terminal {terminal.name}", root=root))
    require(not unfinished, f"validation jobs are not terminal: {unfinished}")
    require(receipts, "no terminal validation receipts were found")
    return {"count": len(receipts), "receipts": receipts}


def _candidate_paths(temp: Path, target: Path, attempt: str) -> tuple[list[Path], Path]:
    temporary_state = _directory(temp, "temporary root", owner_uid=os.getuid())
    target_state = _directory(target, "Cargo target", owner_uid=os.getuid())
    assert temporary_state is not None and target_state is not None
    retained = temp / "retained"
    _directory(retained, "retained root", owner_uid=os.getuid())
    _directory(retained / attempt, "retained attempt", owner_uid=os.getuid())
    candidates: list[Path] = []
    device = temp.lstat().st_dev
    for child in sorted(temp.iterdir(), key=lambda item: item.name):
        if child.name == "retained":
            continue
        _tree_stats(child, f"temporary candidate {child.name}", owner_uid=os.getuid(), root_device=device)
        candidates.append(child)
    for child in sorted(retained.iterdir(), key=lambda item: item.name):
        if child.name == attempt:
            continue
        _tree_stats(child, f"retained candidate {child.name}", owner_uid=os.getuid(), root_device=retained.lstat().st_dev)
        candidates.append(child)
    _tree_stats(target, "Cargo target", owner_uid=os.getuid(), root_device=target.lstat().st_dev)
    return candidates, target


def _disk(path: Path) -> dict[str, int]:
    try:
        value = shutil.disk_usage(path)
    except OSError as error:
        fail(f"cannot inspect free space at {path}: {error}")
    return {"total_bytes": value.total, "used_bytes": value.used, "free_bytes": value.free}


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


def _process_audit(deletion_roots: Iterable[Path], evidence_root: Path,
                   temp: Path, target: Path, *, proc_root: Path = Path("/proc"),
                   self_pid: int | None = None) -> dict[str, Any]:
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
            excluded_other_uid.append({"pid": pid, "uid": state.st_uid})
            continue
        name = _proc_bytes(process / "comm").decode(errors="replace").strip()
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
            if name in PROTECTED_DAEMONS:
                excluded_daemons.append({"pid": pid, "comm": name,
                                         "reason": "protected session daemon; descendants remain audited"})
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
        workspace_match = str(temp) in command_text or str(target) in command_text
        executable_match = exe is not None and any(_path_inside(exe, candidate) for candidate in deletion)
        if script_match or workspace_match or executable_match:
            drivers.append({"pid": pid, "exe": None if exe is None else str(exe),
                            "command": command_text, "script_reference": script_match,
                            "workspace_reference": workspace_match,
                            "executable_reference": executable_match})
    return {
        "proc_root": str(proc_root.resolve(strict=False)), "self_pid": self_pid,
        "uid_scope": uid_scope, "scanned_processes": scanned,
        "vanished_processes": vanished,
        "excluded_other_uid_processes": excluded_other_uid,
        "excluded_session_daemons": excluded_daemons,
        "candidate_references": busy, "root_gate_capture_processes": drivers,
        "safe": not busy and not drivers,
    }


def _process_failure(value: Mapping[str, Any]) -> str:
    if value.get("candidate_references"):
        return f"active process references cleanup candidate: {value['candidate_references']}"
    if value.get("root_gate_capture_processes"):
        return f"live 0495 gate/capture process remains: {value['root_gate_capture_processes']}"
    return "process audit failed"


def _validate_capture_custody(root: Path, temp: Path, target: Path) -> None:
    fixed = (ROOT.resolve(), TEMP.resolve(), TARGET.resolve())
    configured = (Path(root).resolve(), Path(temp).resolve(), Path(target).resolve())
    matches = [left == right for left, right in zip(configured, fixed)]
    if configured != fixed and any(matches):
        fail("production cleanup custody requires all fixed roots")
    if configured != fixed:
        return
    try:
        import measure
        measure.load_builds()
    except (RuntimeError, OSError, TypeError, ValueError, KeyError) as error:
        fail(f"capture build/gate custody failed: {error}")


def _plan_cleanup(*, root: Path = ROOT, temp: Path = TEMP, target: Path = TARGET,
                  proc_root: Path = Path("/proc"), self_pid: int | None = None) -> dict[str, Any]:
    _validate_capture_custody(root, temp, target)
    uid = os.getuid()
    evidence = _canonical_directory(Path(root), "evidence root")
    temporary = _canonical_directory(Path(temp), "temporary root", owner_uid=uid)
    build_target = _canonical_directory(Path(target), "Cargo target", owner_uid=uid)
    _assert_disjoint(evidence, temporary, build_target)
    cleanup_path = evidence / SEAL_NAME
    require(not cleanup_path.exists() and not cleanup_path.is_symlink(),
            f"refusing to replace existing cleanup receipt: {cleanup_path}")
    terminal = _validate_terminal_receipts(evidence)
    builds = {key: _validate_build(key, evidence, temporary, build_target, target_present=True)
              for key in BUILD_ROLES}
    attempt = _validate_retained_layout(temporary, builds)
    candidates, target_candidate = _candidate_paths(temporary, build_target, attempt)
    # Inventory every candidate against the device of its owned parent.  The
    # explicit loop keeps device identity visible in the receipt.
    removed_stats = []
    for path in candidates:
        device = temporary.lstat().st_dev if path.parent == temporary else (temporary / "retained").lstat().st_dev
        removed_stats.append(_stats_record(
            _tree_stats(path, f"cleanup candidate {path}", owner_uid=uid, root_device=device),
            root=temporary, label="cleanup candidate"))
    removed_stats.append(_stats_record(
        _tree_stats(target_candidate, "Cargo target", owner_uid=uid, root_device=build_target.lstat().st_dev),
        root=build_target.parent, label="Cargo target"))
    process = _process_audit([*candidates, target_candidate], evidence, temporary,
                             build_target, proc_root=proc_root, self_pid=self_pid)
    require(process["safe"], _process_failure(process))
    return {"root": evidence, "temp": temporary, "target": build_target,
            "attempt": attempt, "terminal": terminal, "builds": builds,
            "candidates": candidates, "target_candidate": target_candidate,
            "removed_stats": removed_stats, "process_before": process,
            "proc_root": proc_root, "self_pid": self_pid if self_pid is not None else os.getpid(),
            "disk_before": _disk(temporary)}


def plan_cleanup(**kwargs: Any) -> dict[str, Any]:
    return _plan_cleanup(**kwargs)


def _same_stats(path: Path, expected: Mapping[str, Any], *, root: Path, label: str) -> None:
    stats = _tree_stats(path, label, owner_uid=os.getuid(), root_device=int(expected["device"]))
    actual = _stats_record(stats, root=root, label=label)
    for key in ("kind", "files", "directories", "logical_bytes", "allocated_bytes",
                "device", "inode", "fingerprint"):
        require(actual.get(key) == expected.get(key), f"{label}: changed after planning ({key})")


def _remove_tree(path: Path, label: str, *, root_device: int) -> None:
    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: refusing symlink")
    require(stat.S_ISREG(state.st_mode) or stat.S_ISDIR(state.st_mode),
            f"{label}: refusing special path")
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
        fail(f"{label}: cannot enumerate directory: {error}")
    for child in children:
        _remove_tree(child, f"{label}/{child.name}", root_device=root_device)
    try:
        path.rmdir()
    except OSError as error:
        fail(f"{label}: cannot remove directory: {error}")


def _absent(path: Path, label: str) -> None:
    _no_symlink_components(path, label)
    require(not path.exists() and not path.is_symlink(), f"{label}: path remains: {path}")


def _revalidate_plan(plan: Mapping[str, Any]) -> dict[str, Any]:
    root = _canonical_directory(Path(str(plan["root"])), "evidence root")
    temp = _canonical_directory(Path(str(plan["temp"])), "temporary root", owner_uid=os.getuid())
    target = _canonical_directory(Path(str(plan["target"])), "Cargo target", owner_uid=os.getuid())
    _assert_disjoint(root, temp, target)
    attempt = _safe_attempt(plan.get("attempt"), "cleanup plan")
    builds = {key: _validate_build(key, root, temp, target, target_present=True,
                                   expected_attempt=attempt) for key in BUILD_ROLES}
    require(_validate_retained_layout(temp, builds) == attempt, "retained layout changed")
    candidates, target_candidate = _candidate_paths(temp, target, attempt)
    expected = {str(Path(item).resolve(strict=False)) for item in plan["candidates"]}
    require({str(path.resolve(strict=False)) for path in candidates} == expected,
            "cleanup candidate set changed")
    require(str(target_candidate.resolve(strict=False)) == str(Path(str(plan["target_candidate"])).resolve(strict=False)),
            "target candidate changed")
    return {"root": root, "temp": temp, "target": target, "attempt": attempt,
            "builds": builds, "candidates": candidates, "target_candidate": target_candidate}


def _retained(builds: Mapping[str, Mapping[str, Any]]) -> list[dict[str, Any]]:
    return [builds[key]["binary"] | {"key": key, "build_receipt": builds[key]["receipt"]}
            for key in BUILD_ROLES]


def _totals(records: Iterable[Mapping[str, Any]]) -> dict[str, int]:
    values = list(records)
    return {"paths": len(values), "files": sum(int(v["files"]) for v in values),
            "directories": sum(int(v["directories"]) for v in values),
            "logical_bytes": sum(int(v["logical_bytes"]) for v in values),
            "allocated_bytes": sum(int(v["allocated_bytes"]) for v in values)}


def _revalidate_retained(builds: Mapping[str, Mapping[str, Any]]) -> None:
    for key, build in builds.items():
        path = Path(str(build["binary"]["path"]))
        state = _regular(path, f"{key} retained binary")
        assert state is not None
        require(state.st_size == build["binary"]["bytes"]
                and _sha256(path) == build["binary"]["sha256"]
                and os.access(path, os.X_OK), f"{key} retained binary changed")


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


def execute_cleanup(plan: Mapping[str, Any]) -> dict[str, Any]:
    _validate_capture_custody(Path(str(plan["root"])), Path(str(plan["temp"])), Path(str(plan["target"])))
    authenticated = _revalidate_plan(plan)
    root, temp, target = authenticated["root"], authenticated["temp"], authenticated["target"]
    receipt_path = root / SEAL_NAME
    require(not receipt_path.exists() and not receipt_path.is_symlink(),
            f"refusing to replace receipt: {receipt_path}")
    before = _process_audit([*authenticated["candidates"], authenticated["target_candidate"]],
                            root, temp, target, proc_root=Path(str(plan["proc_root"])),
                            self_pid=int(plan["self_pid"]))
    require(before["safe"], _process_failure(before))
    expected_by_path = {str(Path(item["path"]).resolve(strict=False)): item for item in plan["removed_stats"]}
    require(len(expected_by_path) == len(plan["removed_stats"]), "cleanup inventory has duplicate paths")
    paths = [*authenticated["candidates"], authenticated["target_candidate"]]
    require({str(path.resolve(strict=False)) for path in paths} == set(expected_by_path),
            "cleanup inventory changed")
    for path in paths:
        expected = expected_by_path[str(path.resolve(strict=False))]
        root_for_record = temp if _path_inside(path, temp) else target.parent
        _same_stats(path, expected, root=root_for_record, label=f"cleanup candidate {path}")
    for path in paths:
        record = expected_by_path[str(path.resolve(strict=False))]
        _remove_tree(path, f"remove {path}", root_device=int(record["device"]))
        _absent(path, f"removed cleanup candidate {path}")
    _revalidate_retained(authenticated["builds"])
    require({item.name for item in temp.iterdir()} == {"retained"},
            "unexpected temporary paths remain")
    after = _process_audit(paths, root, temp, target,
                           proc_root=Path(str(plan["proc_root"])), self_pid=int(plan["self_pid"]))
    require(after["safe"], _process_failure(after))
    receipt = {
        "schema": SCHEMA, "version": VERSION, "status": "pass", "completed_utc": _now(),
        "driver": _descriptor(Path(__file__).resolve(), "cleanup driver"),
        "target": str(target), "temporary_root": str(temp),
        "scope": {"evidence_root": str(root), "temporary_root": str(temp),
                  "build_target": str(target), "retained_attempt": authenticated["attempt"],
                  "retained_subtree": str((temp / "retained" / authenticated["attempt"]).resolve()),
                  "retained_executables": 4, "root_evidence_preserved": True,
                  "protected_worktree_touched": False},
        "terminal_validation": _validate_terminal_receipts(root),
        "retained_attempt": authenticated["attempt"], "retained_binaries": _retained(authenticated["builds"]),
        "removed": list(plan["removed_stats"]), "removed_paths": list(plan["removed_stats"]),
        "removed_totals": _totals(plan["removed_stats"]), "target_removed": True,
        "temporary_scratch_remaining": ["retained", authenticated["attempt"]],
        "disk": {"before": plan.get("disk_before", _disk(temp)), "after": _disk(temp)},
        "process_audit": {"before": before, "after": after},
    }
    _write_exclusive(receipt_path, receipt)
    return receipt


def _verify_removed(records: Any, *, temp: Path, target: Path, attempt: str) -> dict[str, Any]:
    require(isinstance(records, list) and records, "cleanup removed inventory is missing")
    seen: set[str] = set()
    target_key = str(target.resolve(strict=False))
    target_record: Mapping[str, Any] | None = None
    for index, record in enumerate(records):
        require(isinstance(record, Mapping), f"cleanup.removed[{index}] is malformed")
        raw = record.get("path")
        require(isinstance(raw, str) and raw, f"cleanup.removed[{index}].path missing")
        path = Path(raw)
        _no_symlink_components(path, f"cleanup.removed[{index}]")
        key = str(path.resolve(strict=False))
        require(key not in seen, f"cleanup.removed[{index}] is duplicated")
        seen.add(key)
        if key == target_key:
            target_record = record
        else:
            require(_path_inside(path, temp) and not _path_inside(path, temp / "retained" / attempt),
                    f"cleanup.removed[{index}] escapes disposable scratch")
        _absent(path, f"cleanup.removed[{index}]")
        for field in ("files", "directories", "logical_bytes", "allocated_bytes", "device", "inode"):
            require(type(record.get(field)) is int and record[field] >= 0,
                    f"cleanup.removed[{index}].{field} is malformed")
        fingerprint = record.get("fingerprint")
        require(isinstance(fingerprint, str) and len(fingerprint) == 64
                and all(char in "0123456789abcdef" for char in fingerprint),
                f"cleanup.removed[{index}].fingerprint is malformed")
    require(target_record is not None, "cleanup removed inventory omits Cargo target")
    return {"records": records, "target": target_record}


def verify(*, root: Path = ROOT, temp: Path = TEMP, target: Path = TARGET,
           proc_root: Path = Path("/proc")) -> dict[str, Any]:
    evidence = _canonical_directory(Path(root), "evidence root")
    temporary = _canonical_directory(Path(temp), "temporary root", owner_uid=os.getuid())
    target_path = Path(target)
    _no_symlink_components(target_path, "Cargo target")
    _assert_disjoint(evidence, temporary, target_path.resolve(strict=False))
    _absent(target_path, "Cargo target")
    receipt_path = evidence / SEAL_NAME
    receipt = _read_json(receipt_path, "cleanup receipt")
    require(isinstance(receipt, Mapping) and receipt.get("schema") == SCHEMA
            and receipt.get("version") == VERSION and receipt.get("status") == "pass",
            "cleanup receipt schema/status differs")
    require(receipt.get("target") == str(target_path.resolve(strict=False))
            and receipt.get("temporary_root") == str(temporary),
            "cleanup roots differ")
    attempt = _safe_attempt(receipt.get("retained_attempt"), "cleanup receipt")
    builds = {key: _validate_build(key, evidence, temporary, target_path,
                                   target_present=False, expected_attempt=attempt)
              for key in BUILD_ROLES}
    require(_validate_retained_layout(temporary, builds) == attempt,
            "retained layout differs")
    require({item.name for item in temporary.iterdir()} == {"retained"},
            "temporary root retains disposable paths")
    require({item.name for item in (temporary / "retained").iterdir()} == {attempt},
            "retained root contains an unexpected attempt")
    expected_scope = {"evidence_root": str(evidence), "temporary_root": str(temporary),
                      "build_target": str(target_path.resolve(strict=False)),
                      "retained_attempt": attempt,
                      "retained_subtree": str((temporary / "retained" / attempt).resolve(strict=True)),
                      "retained_executables": 4, "root_evidence_preserved": True,
                      "protected_worktree_touched": False}
    require(receipt.get("scope") == expected_scope, "cleanup scope binding differs")
    require(receipt.get("retained_binaries") == _retained(builds),
            "retained executable inventory differs")
    terminal = _validate_terminal_receipts(evidence)
    require(receipt.get("terminal_validation") == terminal, "terminal validation changed")
    removed = _verify_removed(receipt.get("removed"), temp=temporary, target=target_path, attempt=attempt)
    require(receipt.get("removed_paths") == removed["records"]
            and receipt.get("removed_totals") == _totals(removed["records"]),
            "removed inventory changed")
    require(receipt.get("target_removed") is True
            and receipt.get("temporary_scratch_remaining") == ["retained", attempt],
            "remaining scratch binding differs")
    driver = receipt.get("driver")
    require(isinstance(driver, Mapping) and driver == _descriptor(Path(__file__).resolve(), "cleanup driver"),
            "cleanup driver changed")
    current = _process_audit([Path(str(item["path"])) for item in removed["records"]],
                             evidence, temporary, target_path, proc_root=proc_root)
    require(current["safe"], _process_failure(current))
    process = receipt.get("process_audit")
    require(isinstance(process, Mapping), "cleanup process audit missing")
    for phase in ("before", "after"):
        value = process.get(phase)
        require(isinstance(value, Mapping) and value.get("safe") is True
                and value.get("candidate_references") == []
                and value.get("root_gate_capture_processes") == [],
                f"cleanup process audit {phase} was unsafe")
    return {"schema": VERIFY_SCHEMA, "version": VERSION, "status": "pass",
            "cleanup_receipt": _descriptor(receipt_path, "cleanup receipt", root=evidence),
            "retained_binaries": _retained(builds), "retained_attempt": attempt,
            "build_target_removed": True, "remaining": [], "process_audit": current}


def _print_plan(plan: Mapping[str, Any]) -> None:
    print(json.dumps({"schema": SCHEMA, "status": "ready", "dry_run": True,
                      "retained_attempt": plan["attempt"],
                      "retained_binaries": _retained(plan["builds"]),
                      "removed": plan["removed_stats"],
                      "removed_totals": _totals(plan["removed_stats"]),
                      "process_audit": plan["process_before"]}, indent=2, sort_keys=True))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--dry-run", action="store_true")
    modes.add_argument("--verify", action="store_true")
    parser.add_argument("--root", type=Path, default=ROOT, help=argparse.SUPPRESS)
    parser.add_argument("--temp-root", type=Path, default=TEMP, help=argparse.SUPPRESS)
    parser.add_argument("--target-root", type=Path, default=TARGET, help=argparse.SUPPRESS)
    parser.add_argument("--proc-root", type=Path, default=Path("/proc"), help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    try:
        configured = (args.root.resolve(), args.temp_root.resolve(), args.target_root.resolve())
        fixed = (ROOT.resolve(), TEMP.resolve(), TARGET.resolve())
        if not args.dry_run and not args.verify:
            require(configured == fixed, "destructive mode is restricted to fixed 0495 roots")
        if args.verify:
            print(json.dumps(verify(root=args.root, temp=args.temp_root, target=args.target_root,
                                    proc_root=args.proc_root), indent=2, sort_keys=True))
            return 0
        plan = _plan_cleanup(root=args.root, temp=args.temp_root, target=args.target_root,
                             proc_root=args.proc_root)
        if args.dry_run:
            _print_plan(plan)
            return 0
        receipt = execute_cleanup(plan)
        print(json.dumps({"schema": SCHEMA, "status": receipt["status"],
                          "retained_attempt": receipt["retained_attempt"],
                          "removed_totals": receipt["removed_totals"]}, sort_keys=True))
        return 0
    except (CleanupError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"cleanup.py: error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
