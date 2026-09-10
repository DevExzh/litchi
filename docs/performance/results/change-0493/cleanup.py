#!/usr/bin/env python3
"""Authenticate and clean the disposable 0493 benchmark caches.

The build receipts identify the two release executables which must survive the
cleanup.  Their attempt name is read from the receipts, so this driver does
not depend on a particular ``finalN`` naming convention.  All other children
of the owned scratch root and the owned Cargo target are candidates.  The
driver inventories candidates without following links, rejects special files
and device boundaries, audits current processes, and writes an exclusive
receipt only after a successful cleanup.  ``--dry-run`` and ``--verify`` are
read-only.  The evidence directory and ``litchi-spec-gaps`` are never targets.
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
TEMP = Path("/home/zhuhe/.cache/litchi-goal-0493")
TARGET = Path("/home/zhuhe/.cache/litchi-build-0493")
BUILD_SCHEMA = "docx-provider-lifecycle-build-v1"
SCHEMA = "docx-managed-read-ahead-cleanup-v1"
VERSION = 1
SEAL_NAME = "cleanup.json"
BUILD_ROLES = {
    "normal": "litchi-perf-baseline",
    "allocator": "litchi-perf-baseline-alloc",
}
ROOT_PROCESS_SCRIPTS = frozenset(
    {
        "cleanup.py",
        "docx_managed_read_ahead.py",
        "gate.py",
        "measure.py",
        "retain_build.py",
        "test_cleanup.py",
        "test_measure.py",
        "verify_bundle.py",
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


def _sha256(path: Path) -> str:
    try:
        with path.open("rb") as stream:
            return hashlib.file_digest(stream, "sha256").hexdigest()
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    raise AssertionError("unreachable")


def _no_symlink_components(path: Path, label: str) -> None:
    """Reject a symlink in every existing component of an absolute path."""

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


def _regular(path: Path, label: str, *, owner_uid: int | None = None) -> os.stat_result:
    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is not allowed: {path}")
    require(stat.S_ISREG(state.st_mode), f"{label}: regular file required: {path}")
    if owner_uid is not None:
        require(state.st_uid == owner_uid, f"{label}: unexpected owner: {path}")
    return state


def _directory(path: Path, label: str, *, owner_uid: int | None = None) -> os.stat_result:
    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is not allowed: {path}")
    require(stat.S_ISDIR(state.st_mode), f"{label}: directory required: {path}")
    if owner_uid is not None:
        require(state.st_uid == owner_uid, f"{label}: unexpected owner: {path}")
    return state


def _canonical_directory(path: Path, label: str, *, owner_uid: int | None = None) -> Path:
    _no_symlink_components(path, label)
    _directory(path, label, owner_uid=owner_uid)
    try:
        resolved = path.resolve(strict=True)
    except OSError as error:
        fail(f"{label}: cannot resolve {path}: {error}")
    require(resolved == path.absolute(), f"{label}: path is not canonical: {path}")
    _directory(resolved, label, owner_uid=owner_uid)
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
        fail(f"{label}: {path} escapes {root}: {error}")
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


def _tree_stats(
    path: Path,
    label: str,
    *,
    owner_uid: int | None = None,
    root_device: int | None = None,
) -> _TreeStats:
    """Inventory a regular tree without following links, mounts, or devices."""

    root_state = _lstat(path, label)
    require(not stat.S_ISLNK(root_state.st_mode), f"{label}: symlink is not allowed: {path}")
    require(
        stat.S_ISREG(root_state.st_mode) or stat.S_ISDIR(root_state.st_mode),
        f"{label}: only regular files and directories are allowed: {path}",
    )
    if root_device is None:
        root_device = root_state.st_dev
    require(root_state.st_dev == root_device, f"{label}: device boundary: {path}")
    seen: set[tuple[int, int]] = set()
    fingerprint = hashlib.sha256()

    def visit(current: Path, relative: str) -> tuple[int, int, int, int]:
        state = _lstat(current, f"{label}/{relative or '.'}")
        require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is not allowed: {current}")
        require(
            stat.S_ISREG(state.st_mode) or stat.S_ISDIR(state.st_mode),
            f"{label}: special file is not allowed: {current}",
        )
        require(state.st_dev == root_device, f"{label}: device boundary: {current}")
        if owner_uid is not None:
            require(state.st_uid == owner_uid, f"{label}: unexpected owner: {current}")
        kind = "file" if stat.S_ISREG(state.st_mode) else "directory"
        allocated = _allocated_bytes(state)
        identity = (state.st_dev, state.st_ino)
        fingerprint.update(
            "\0".join(
                str(value)
                for value in (
                    relative,
                    kind,
                    state.st_size if kind == "file" else 0,
                    allocated,
                    state.st_mode,
                    state.st_mtime_ns,
                    state.st_dev,
                    state.st_ino,
                )
            ).encode("utf-8")
        )
        fingerprint.update(b"\n")
        unique_allocated = 0
        if identity not in seen:
            seen.add(identity)
            unique_allocated = allocated
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


def _assert_disjoint(evidence: Path, temp: Path, target: Path) -> None:
    require(not _path_inside(evidence, temp) and not _path_inside(temp, evidence), "evidence and temporary roots overlap")
    require(not _path_inside(evidence, target) and not _path_inside(target, evidence), "evidence and build roots overlap")
    require(not _path_inside(temp, target) and not _path_inside(target, temp), "temporary and build roots overlap")


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
    record: dict[str, Any] = {"path": str(resolved), "bytes": resolved.stat().st_size, "sha256": _sha256(resolved)}
    if root is not None:
        record["relative_path"] = _relative(resolved, root, label)
    return record


def _path_from_receipt(value: Any, label: str) -> Path:
    require(isinstance(value, Mapping), f"{label}: descriptor missing")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path: missing")
    path = Path(raw)
    require(path.is_absolute(), f"{label}.path: absolute path required")
    _no_symlink_components(path, label)
    return path


def _safe_attempt(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label}: attempt missing")
    require(len(value) <= 96 and all(char.isalnum() or char in "_-" for char in value), f"{label}: unsafe attempt name")
    require(value not in {".", ".."}, f"{label}: unsafe attempt name")
    return value


def _binary_binding(value: Any, label: str, path: Path, *, present: bool) -> dict[str, Any]:
    require(isinstance(value, Mapping), f"{label}: descriptor missing")
    require(value.get("executable") is True, f"{label}: executable binding failed")
    size = value.get("bytes")
    digest = value.get("sha256")
    require(type(size) is int and size >= 0, f"{label}: byte length malformed")
    require(
        isinstance(digest, str)
        and len(digest) == 64
        and all(char in "0123456789abcdef" for char in digest),
        f"{label}: SHA-256 malformed",
    )
    if present:
        _no_symlink_components(path, label)
        state = _regular(path, label)
        require({"bytes": state.st_size, "sha256": _sha256(path)} == {"bytes": size, "sha256": digest}, f"{label}: content changed")
        require(bool(state.st_mode & 0o111) and os.access(path, os.X_OK), f"{label}: executable bit missing")
    return {"path": str(path.resolve(strict=False)), "bytes": size, "sha256": digest, "executable": True}


def _original_binding(value: Any, label: str, target: Path, *, present: bool) -> dict[str, Any]:
    path = _path_from_receipt(value, label)
    resolved = path.resolve(strict=False)
    require(_path_inside(resolved, target), f"{label}: path escapes build target")
    require(resolved != target.resolve(strict=False), f"{label}: path points at target root")
    return _binary_binding(value, label, resolved, present=present)


def _validate_build(
    role: str,
    root: Path,
    temp: Path,
    target: Path,
    *,
    target_present: bool,
    expected_attempt: str | None = None,
) -> dict[str, Any]:
    receipt_path = root / f"build-{role}.json"
    value = _read_json(receipt_path, f"{role} build receipt")
    require(isinstance(value, Mapping), f"{role} build receipt: object required")
    require(value.get("schema") == BUILD_SCHEMA and value.get("version") == VERSION, f"{role} build receipt: schema mismatch")
    require(value.get("role") == role, f"{role} build receipt: role mismatch")
    attempt = _safe_attempt(value.get("attempt"), f"{role} build receipt")
    if expected_attempt is not None:
        require(attempt == expected_attempt, f"{role} build receipt: attempts differ")
    require(value.get("source_unchanged") is True and value.get("source_before") == value.get("source_after"), f"{role} build receipt: source custody failed")

    binary_value = value.get("binary")
    binary_path = _path_from_receipt(binary_value, f"{role} retained binary")
    expected_path = temp / attempt / role / BUILD_ROLES[role]
    _no_symlink_components(expected_path, f"{role} retained binary destination")
    require(binary_path.resolve(strict=False) == expected_path.absolute(), f"{role} retained binary path differs")
    binary = _binary_binding(binary_value, f"{role} retained binary", expected_path, present=True)

    original = _original_binding(value.get("original_binary"), f"{role} original Cargo binary", target, present=target_present)
    return {
        "role": role,
        "attempt": attempt,
        "build_receipt": _descriptor(receipt_path, f"{role} build receipt", root=root),
        "binary": binary,
        "original_binary": original,
    }


def _validate_retained_layout(temp: Path, builds: Mapping[str, Mapping[str, Any]]) -> str:
    attempts = {str(builds[role]["attempt"]) for role in BUILD_ROLES}
    require(len(attempts) == 1, "retained build attempts differ")
    attempt = next(iter(attempts))
    retained_root = temp / attempt
    temp_state = _directory(temp, "temporary root", owner_uid=os.getuid())
    _tree_stats(retained_root, "retained build root", owner_uid=os.getuid(), root_device=temp_state.st_dev)
    _directory(retained_root, "retained build root", owner_uid=os.getuid())
    require({child.name for child in retained_root.iterdir()} == set(BUILD_ROLES), "retained root has unexpected children")
    for role in BUILD_ROLES:
        role_root = retained_root / role
        _directory(role_root, f"retained {role} directory", owner_uid=os.getuid())
        name = BUILD_ROLES[role]
        require({child.name for child in role_root.iterdir()} == {name}, f"retained {role} directory has unexpected children")
        binary = role_root / name
        state = _regular(binary, f"retained {role} binary", owner_uid=os.getuid())
        require(bool(state.st_mode & 0o111) and os.access(binary, os.X_OK), f"retained {role} binary is not executable")
        require(Path(str(builds[role]["binary"]["path"])).resolve(strict=False) == binary.absolute(), f"retained {role} path differs")
    return attempt


def _validate_terminal_receipts(root: Path) -> dict[str, Any]:
    validation = root / "validation"
    _directory(validation, "validation receipt directory")
    receipts: list[dict[str, Any]] = []
    unfinished: list[str] = []
    for started in sorted(validation.glob("*.started.json"), key=lambda item: item.name):
        _regular(started, f"validation start receipt {started.name}")
        result = validation / f"{started.name[:-len('.started.json')]}.json"
        if result.is_symlink() or not result.exists():
            unfinished.append(started.name)
            continue
        value = _read_json(result, f"validation receipt {result.name}")
        require(isinstance(value, Mapping), f"validation receipt {result.name}: object required")
        require(type(value.get("exit_code")) is int, f"validation receipt {result.name}: exit_code missing")
        require(isinstance(value.get("finished_utc"), str) and value["finished_utc"], f"validation receipt {result.name}: finished_utc missing")
        receipts.append(_descriptor(result, f"validation receipt {result.name}", root=root))
    require(not unfinished, f"validation jobs are not terminal: {unfinished}")
    require(receipts, "no terminal validation receipts were found")
    return {"count": len(receipts), "receipts": receipts}


def _candidate_paths(temp: Path, target: Path, attempt: str) -> tuple[list[Path], Path]:
    _directory(temp, "temporary root", owner_uid=os.getuid())
    _directory(target, "Cargo build target", owner_uid=os.getuid())
    retained = temp / attempt
    _directory(retained, "retained build root", owner_uid=os.getuid())
    temp_device = temp.lstat().st_dev
    candidates: list[Path] = []
    for child in sorted(temp.iterdir(), key=lambda item: item.name):
        if child.name == attempt:
            continue
        _tree_stats(child, f"temporary candidate {child.name}", owner_uid=os.getuid(), root_device=temp_device)
        candidates.append(child)
    _tree_stats(target, "Cargo build target", owner_uid=os.getuid(), root_device=target.lstat().st_dev)
    return candidates, target


def _disk(path: Path) -> dict[str, int]:
    try:
        usage = shutil.disk_usage(path)
    except OSError as error:
        fail(f"cannot inspect free space at {path}: {error}")
    return {"total_bytes": usage.total, "used_bytes": usage.used, "free_bytes": usage.free}


def _proc_path(raw: str) -> Path | None:
    if not raw or not raw.startswith("/"):
        return None
    if raw.endswith(_DELETED_SUFFIX):
        raw = raw[: -len(_DELETED_SUFFIX)]
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
        raise CleanupError(f"cannot inspect process link {path}: {error}") from error
    return _proc_path(os.fsdecode(raw))


def _proc_bytes(path: Path) -> bytes:
    try:
        return path.read_bytes()
    except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
        return b""
    except OSError as error:
        raise CleanupError(f"cannot inspect process file {path}: {error}") from error


def _process_audit(
    deletion_roots: Iterable[Path],
    evidence_root: Path,
    temp: Path,
    target: Path,
    *,
    proc_root: Path = Path("/proc"),
    self_pid: int | None = None,
) -> dict[str, Any]:
    """Audit cwd, executable, descriptors, and active 0493 drivers."""

    _directory(proc_root, "process information root")
    deletion = tuple(path.resolve(strict=False) for path in deletion_roots)
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
        require(stat.S_ISDIR(state.st_mode) and not stat.S_ISLNK(state.st_mode), f"/proc/{pid}: process entry is not a directory")
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
                excluded_daemons.append({"pid": pid, "comm": name, "reason": "protected session daemon; descendants remain audited"})
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
                    busy.append({"pid": pid, "reference": kind, "path": str(path), "candidate": str(candidate)})
                    break

        command_text = " ".join(command)
        script_match = False
        for argument in command:
            if Path(argument).name not in ROOT_PROCESS_SCRIPTS:
                continue
            if argument.startswith("/"):
                script = _proc_path(argument)
            elif cwd is not None:
                script = _proc_path(str(cwd / argument))
            else:
                script = None
            if script is not None and _path_inside(script, evidence_root):
                script_match = True
                break
        workspace_match = str(temp) in command_text or str(target) in command_text
        executable_match = exe is not None and any(_path_inside(exe, candidate) for candidate in deletion)
        if script_match or workspace_match or executable_match:
            drivers.append(
                {
                    "pid": pid,
                    "exe": None if exe is None else str(exe),
                    "command": command_text,
                    "script_reference": script_match,
                    "workspace_reference": workspace_match,
                    "executable_reference": executable_match,
                }
            )

    return {
        "proc_root": str(proc_root.resolve(strict=False)),
        "self_pid": self_pid,
        "uid_scope": uid_scope,
        "scanned_processes": scanned,
        "vanished_processes": vanished,
        "excluded_other_uid_processes": excluded_other_uid,
        "excluded_session_daemons": excluded_daemons,
        "candidate_references": busy,
        "root_gate_capture_processes": drivers,
        "safe": not busy and not drivers,
    }


def _process_failure(value: Mapping[str, Any]) -> str:
    if value.get("candidate_references"):
        return f"active process references cleanup candidate: {value['candidate_references']}"
    if value.get("root_gate_capture_processes"):
        return f"live 0493 gate/capture process remains: {value['root_gate_capture_processes']}"
    return "process audit failed"


def _validate_capture_custody(root: Path, temp: Path, target: Path) -> None:
    """Authenticate production build/gate custody before destructive cleanup.

    Isolated fixture roots use the local synthetic build validator. Any use of
    a production root requires the complete fixed scope and capture validator.
    """
    configured = (root.resolve(), temp.resolve(), target.resolve())
    fixed = (ROOT.resolve(), TEMP.resolve(), TARGET.resolve())
    if not any(value == expected for value, expected in zip(configured, fixed)):
        return
    require(configured == fixed, "production cleanup custody requires all fixed roots")
    import measure
    try:
        measure.load_builds()
    except (RuntimeError, OSError, TypeError, ValueError, KeyError) as error:
        fail(f"capture build/gate custody failed: {error}")


def _plan_cleanup(
    *,
    root: Path = ROOT,
    temp: Path = TEMP,
    target: Path = TARGET,
    proc_root: Path = Path("/proc"),
    self_pid: int | None = None,
) -> dict[str, Any]:
    _validate_capture_custody(root, temp, target)
    uid = os.getuid()
    evidence = _canonical_directory(root, "evidence root")
    temporary = _canonical_directory(temp, "temporary root", owner_uid=uid)
    build_target = _canonical_directory(target, "Cargo build target", owner_uid=uid)
    _assert_disjoint(evidence, temporary, build_target)
    cleanup_path = evidence / SEAL_NAME
    require(not cleanup_path.exists() and not cleanup_path.is_symlink(), f"refusing to replace existing receipt: {cleanup_path}")
    terminal = _validate_terminal_receipts(evidence)
    builds = {
        role: _validate_build(role, evidence, temporary, build_target, target_present=True)
        for role in BUILD_ROLES
    }
    attempt = _validate_retained_layout(temporary, builds)
    candidates, target_candidate = _candidate_paths(temporary, build_target, attempt)
    removed_stats = [
        _stats_record(_tree_stats(path, f"temporary candidate {path.name}", owner_uid=uid, root_device=temporary.lstat().st_dev), root=temporary, label="temporary candidate")
        for path in candidates
    ]
    removed_stats.append(
        _stats_record(_tree_stats(target_candidate, "Cargo build target", owner_uid=uid, root_device=build_target.lstat().st_dev), root=build_target.parent, label="Cargo build target")
    )
    process = _process_audit([*candidates, target_candidate], evidence, temporary, build_target, proc_root=proc_root, self_pid=self_pid)
    require(process["safe"], _process_failure(process))
    return {
        "root": evidence,
        "temp": temporary,
        "target": build_target,
        "attempt": attempt,
        "terminal": terminal,
        "builds": builds,
        "candidates": candidates,
        "target_candidate": target_candidate,
        "removed_stats": removed_stats,
        "process_before": process,
        "proc_root": proc_root,
        "self_pid": self_pid if self_pid is not None else os.getpid(),
        "disk_before": _disk(temporary),
    }


def plan_cleanup(**kwargs: Any) -> dict[str, Any]:
    """Run read-only custody checks and return an authenticated plan."""

    return _plan_cleanup(**kwargs)


def _same_stats(path: Path, expected: Mapping[str, Any], *, root: Path, label: str) -> None:
    stats = _tree_stats(path, label, owner_uid=os.getuid(), root_device=int(expected["device"]))
    actual = _stats_record(stats, root=root, label=label)
    for key in ("kind", "files", "directories", "logical_bytes", "allocated_bytes", "device", "inode", "fingerprint"):
        require(actual.get(key) == expected.get(key), f"{label}: changed after planning ({key})")


def _remove_tree(path: Path, label: str, *, root_device: int) -> None:
    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: refusing symlink: {path}")
    require(stat.S_ISREG(state.st_mode) or stat.S_ISDIR(state.st_mode), f"{label}: refusing special path: {path}")
    require(state.st_uid == os.getuid(), f"{label}: ownership changed: {path}")
    require(state.st_dev == root_device, f"{label}: device changed: {path}")
    if stat.S_ISREG(state.st_mode):
        try:
            path.unlink()
        except OSError as error:
            fail(f"{label}: cannot remove {path}: {error}")
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
        fail(f"{label}: cannot remove directory {path}: {error}")


def _absent(path: Path, label: str) -> None:
    _no_symlink_components(path, label)
    require(not path.exists() and not path.is_symlink(), f"{label}: path remains: {path}")


def _revalidate_plan(plan: Mapping[str, Any]) -> dict[str, Any]:
    root = _canonical_directory(Path(str(plan["root"])), "evidence root")
    temp = _canonical_directory(Path(str(plan["temp"])), "temporary root", owner_uid=os.getuid())
    target = _canonical_directory(Path(str(plan["target"])), "Cargo build target", owner_uid=os.getuid())
    _assert_disjoint(root, temp, target)
    attempt = _safe_attempt(plan.get("attempt"), "cleanup plan")
    builds = {
        role: _validate_build(role, root, temp, target, target_present=True, expected_attempt=attempt)
        for role in BUILD_ROLES
    }
    require(_validate_retained_layout(temp, builds) == attempt, "retained layout changed")
    candidates, target_candidate = _candidate_paths(temp, target, attempt)
    expected_candidates = {str(Path(path).resolve(strict=False)) for path in plan["candidates"]}
    actual_candidates = {str(path.resolve(strict=False)) for path in candidates}
    require(actual_candidates == expected_candidates, "temporary candidate set changed after planning")
    require(str(target_candidate.resolve(strict=False)) == str(Path(str(plan["target_candidate"])).resolve(strict=False)), "target candidate changed")
    return {"root": root, "temp": temp, "target": target, "attempt": attempt, "builds": builds, "candidates": candidates, "target_candidate": target_candidate}


def _retained(builds: Mapping[str, Mapping[str, Any]]) -> list[dict[str, Any]]:
    return [
        builds[role]["binary"] | {"role": role, "build_receipt": builds[role]["build_receipt"]}
        for role in BUILD_ROLES
    ]


def _totals(records: Iterable[Mapping[str, Any]]) -> dict[str, int]:
    values = list(records)
    return {
        "paths": len(values),
        "files": sum(int(value["files"]) for value in values),
        "directories": sum(int(value["directories"]) for value in values),
        "logical_bytes": sum(int(value["logical_bytes"]) for value in values),
        "allocated_bytes": sum(int(value["allocated_bytes"]) for value in values),
    }


def _revalidate_retained(builds: Mapping[str, Mapping[str, Any]]) -> None:
    for role, build in builds.items():
        path = Path(str(build["binary"]["path"]))
        actual = {"bytes": _regular(path, f"{role} retained binary").st_size, "sha256": _sha256(path)}
        require(actual == {"bytes": build["binary"]["bytes"], "sha256": build["binary"]["sha256"]}, f"{role} retained binary changed")
        require(bool(path.stat().st_mode & 0o111) and os.access(path, os.X_OK), f"{role} retained binary lost executable bit")


def _write_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to replace existing receipt: {path}")
    try:
        with path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
    except OSError as error:
        fail(f"cannot write cleanup receipt {path}: {error}")


def execute_cleanup(plan: Mapping[str, Any]) -> dict[str, Any]:
    """Execute an authenticated plan and write the cleanup receipt."""

    _validate_capture_custody(Path(str(plan["root"])), Path(str(plan["temp"])), Path(str(plan["target"])))
    authenticated = _revalidate_plan(plan)
    root = authenticated["root"]
    temp = authenticated["temp"]
    target = authenticated["target"]
    cleanup_path = root / SEAL_NAME
    require(not cleanup_path.exists() and not cleanup_path.is_symlink(), f"refusing to replace existing receipt: {cleanup_path}")
    before = _process_audit([*authenticated["candidates"], authenticated["target_candidate"]], root, temp, target, proc_root=Path(str(plan["proc_root"])), self_pid=int(plan["self_pid"]))
    require(before["safe"], _process_failure(before))

    expected_by_path = {str(Path(item["path"]).resolve(strict=False)): item for item in plan["removed_stats"]}
    require(len(expected_by_path) == len(plan["removed_stats"]), "cleanup inventory has duplicate paths")
    paths = [*authenticated["candidates"], authenticated["target_candidate"]]
    require({str(path.resolve(strict=False)) for path in paths} == set(expected_by_path), "cleanup inventory changed")
    for path in paths:
        record = expected_by_path[str(path.resolve(strict=False))]
        _same_stats(path, record, root=temp if path in authenticated["candidates"] else target.parent, label=f"cleanup candidate {path}")

    for path in paths:
        _remove_tree(path, f"remove {path}", root_device=int(expected_by_path[str(path.resolve(strict=False))]["device"]))
        _absent(path, f"removed cleanup candidate {path}")
    _revalidate_retained(authenticated["builds"])
    require({child.name for child in temp.iterdir()} == {authenticated["attempt"]}, "unexpected temporary paths remain")
    after = _process_audit(paths, root, temp, target, proc_root=Path(str(plan["proc_root"])), self_pid=int(plan["self_pid"]))
    require(after["safe"], _process_failure(after))

    removed = list(plan["removed_stats"])
    receipt: dict[str, Any] = {
        "schema": SCHEMA,
        "version": VERSION,
        "status": "pass",
        "completed_utc": _now(),
        "driver": _descriptor(Path(__file__).resolve(), "cleanup driver"),
        "target": str(target),
        "temporary_root": str(temp),
        "scope": {
            "evidence_root": str(root),
            "temporary_root": str(temp),
            "build_target": str(target),
            "retained_attempt": authenticated["attempt"],
            "retained_subtree": str((temp / authenticated["attempt"]).resolve(strict=True)),
            "root_evidence_preserved": True,
            "iwork_tree_touched": False,
        },
        "terminal_validation": _validate_terminal_receipts(root),
        "retained_attempt": authenticated["attempt"],
        "retained_binaries": _retained(authenticated["builds"]),
        "retained": {
            role: {
                "path": authenticated["builds"][role]["binary"]["path"],
                "bytes": authenticated["builds"][role]["binary"]["bytes"],
                "sha256": authenticated["builds"][role]["binary"]["sha256"],
            }
            for role in BUILD_ROLES
        },
        "removed": removed,
        "removed_paths": removed,
        "removed_totals": _totals(removed),
        "target_removed": True,
        "temporary_scratch_remaining": [authenticated["attempt"]],
        "disk": {"before": plan.get("disk_before", _disk(temp)), "after": _disk(temp)},
        "process_audit": {"before": before, "after": after},
    }
    _write_exclusive(cleanup_path, receipt)
    return receipt


def _verify_removed(records: Any, *, temp: Path, target: Path, attempt: str) -> dict[str, Any]:
    require(isinstance(records, list) and records, "cleanup removed inventory is missing")
    seen: set[str] = set()
    target_key = str(target.resolve(strict=False))
    target_record: Mapping[str, Any] | None = None
    for index, record in enumerate(records):
        require(isinstance(record, Mapping), f"cleanup.removed[{index}]: object required")
        raw = record.get("path")
        require(isinstance(raw, str) and raw, f"cleanup.removed[{index}].path missing")
        path = Path(raw)
        _no_symlink_components(path, f"cleanup.removed[{index}]")
        key = str(path.resolve(strict=False))
        require(key not in seen, f"cleanup.removed[{index}]: duplicate path")
        seen.add(key)
        if key == target_key:
            target_record = record
        else:
            require(_path_inside(path, temp), f"cleanup.removed[{index}]: path escapes temporary root")
            require(not _path_inside(path, temp / attempt), f"cleanup.removed[{index}]: retained path was removed")
            require(key != str(temp.resolve(strict=False)), f"cleanup.removed[{index}]: temporary root itself was removed")
        _absent(path, f"cleanup.removed[{index}]")
        for field in ("files", "directories", "logical_bytes", "allocated_bytes", "device", "inode"):
            require(type(record.get(field)) is int and record[field] >= 0, f"cleanup.removed[{index}].{field} malformed")
        fingerprint = record.get("fingerprint")
        require(isinstance(fingerprint, str) and len(fingerprint) == 64 and all(char in "0123456789abcdef" for char in fingerprint), f"cleanup.removed[{index}].fingerprint malformed")
    require(target_record is not None, "cleanup removed inventory omits exact build target")
    return {"records": records, "target": target_record}


def verify(
    *,
    root: Path = ROOT,
    temp: Path = TEMP,
    target: Path = TARGET,
    proc_root: Path = Path("/proc"),
) -> dict[str, Any]:
    """Verify a completed cleanup without changing any path."""

    evidence = _canonical_directory(root, "evidence root")
    temporary = _canonical_directory(temp, "temporary root", owner_uid=os.getuid())
    target_path = Path(target)
    _no_symlink_components(target_path, "Cargo build target")
    _assert_disjoint(evidence, temporary, target_path.resolve(strict=False))
    _absent(target_path, "Cargo build target")
    receipt_path = evidence / SEAL_NAME
    receipt = _read_json(receipt_path, "cleanup receipt")
    require(isinstance(receipt, Mapping), "cleanup receipt: object required")
    require(receipt.get("schema") == SCHEMA and receipt.get("version") == VERSION and receipt.get("status") == "pass", "cleanup receipt: schema/status mismatch")
    require(receipt.get("target") == str(target_path.resolve(strict=False)) and receipt.get("temporary_root") == str(temporary), "cleanup root binding differs")
    attempt = _safe_attempt(receipt.get("retained_attempt"), "cleanup receipt")
    builds = {
        role: _validate_build(role, evidence, temporary, target_path, target_present=False, expected_attempt=attempt)
        for role in BUILD_ROLES
    }
    require(_validate_retained_layout(temporary, builds) == attempt, "retained layout differs")
    require({child.name for child in temporary.iterdir()} == {attempt}, "temporary root retains disposable paths")
    expected_scope = {
        "evidence_root": str(evidence),
        "temporary_root": str(temporary),
        "build_target": str(target_path.resolve(strict=False)),
        "retained_attempt": attempt,
        "retained_subtree": str((temporary / attempt).resolve(strict=True)),
        "root_evidence_preserved": True,
        "iwork_tree_touched": False,
    }
    require(receipt.get("scope") == expected_scope, "cleanup scope binding differs")
    require(receipt.get("retained_binaries") == _retained(builds), "retained binary inventory differs")
    require(receipt.get("retained") == {
        role: {
            "path": builds[role]["binary"]["path"],
            "bytes": builds[role]["binary"]["bytes"],
            "sha256": builds[role]["binary"]["sha256"],
        }
        for role in BUILD_ROLES
    }, "retained compatibility inventory differs")
    terminal = _validate_terminal_receipts(evidence)
    require(receipt.get("terminal_validation") == terminal, "terminal validation inventory differs")
    removed = _verify_removed(receipt.get("removed"), temp=temporary, target=target_path, attempt=attempt)
    require(receipt.get("removed_paths") == removed["records"], "removed path inventory differs")
    require(receipt.get("removed_totals") == _totals(removed["records"]), "removed totals differ")
    require(receipt.get("target_removed") is True and receipt.get("temporary_scratch_remaining") == [attempt], "remaining-root binding differs")
    driver = receipt.get("driver")
    require(isinstance(driver, Mapping) and dict(driver) == _descriptor(Path(__file__).resolve(), "cleanup driver"), "cleanup driver changed")
    removed_paths = [Path(str(item["path"])) for item in removed["records"]]
    current = _process_audit([*removed_paths, target_path], evidence, temporary, target_path, proc_root=proc_root)
    require(current["safe"], _process_failure(current))
    process = receipt.get("process_audit")
    require(isinstance(process, Mapping), "cleanup process audit missing")
    for phase in ("before", "after"):
        value = process.get(phase)
        require(isinstance(value, Mapping) and value.get("safe") is True, f"cleanup process audit {phase} was unsafe")
        require(value.get("candidate_references") == [] and value.get("root_gate_capture_processes") == [], f"cleanup process audit {phase} retained a live process")
    return {
        "schema": "docx-managed-read-ahead-cleanup-verification-v1",
        "version": VERSION,
        "status": "pass",
        "cleanup_receipt": _descriptor(receipt_path, "cleanup receipt", root=evidence),
        "retained_binaries": _retained(builds),
        "retained_attempt": attempt,
        "build_target_removed": True,
        "remaining": [],
        "process_audit": current,
    }


def _print_plan(plan: Mapping[str, Any]) -> None:
    print(json.dumps({
        "schema": SCHEMA,
        "status": "ready",
        "dry_run": True,
        "retained_attempt": plan["attempt"],
        "retained_binaries": _retained(plan["builds"]),
        "removed": plan["removed_stats"],
        "removed_totals": _totals(plan["removed_stats"]),
        "process_audit": plan["process_before"],
    }, indent=2, sort_keys=True))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--dry-run", action="store_true", help="validate and print a plan without mutation")
    modes.add_argument("--verify", action="store_true", help="verify a completed cleanup without mutation")
    parser.add_argument("--root", type=Path, default=ROOT, help=argparse.SUPPRESS)
    parser.add_argument("--temp-root", type=Path, default=TEMP, help=argparse.SUPPRESS)
    parser.add_argument("--target-root", type=Path, default=TARGET, help=argparse.SUPPRESS)
    parser.add_argument("--proc-root", type=Path, default=Path("/proc"), help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    try:
        configured = (args.root.resolve(), args.temp_root.resolve(), args.target_root.resolve())
        fixed = (ROOT.resolve(), TEMP.resolve(), TARGET.resolve())
        if not args.dry_run and not args.verify:
            require(configured == fixed, "destructive mode is restricted to the fixed 0493 roots")
        if args.verify:
            print(json.dumps(verify(root=args.root, temp=args.temp_root, target=args.target_root, proc_root=args.proc_root), indent=2, sort_keys=True))
            return 0
        plan = _plan_cleanup(root=args.root, temp=args.temp_root, target=args.target_root, proc_root=args.proc_root)
        if args.dry_run:
            _print_plan(plan)
            return 0
        receipt = execute_cleanup(plan)
        print(json.dumps({"schema": SCHEMA, "status": receipt["status"], "retained_attempt": receipt["retained_attempt"], "removed_totals": receipt["removed_totals"]}, sort_keys=True))
        return 0
    except (CleanupError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"cleanup.py: error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
