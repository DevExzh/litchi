#!/usr/bin/env python3
"""Bounded cleanup for the 0491 benchmark workspace.

The evidence directory is immutable custody.  This command only removes the
two explicitly owned cache roots used by the run:

* disposable children of ``/home/zhuhe/.cache/litchi-goal-0491``; and
* ``/home/zhuhe/.cache/litchi-build-0491``.

The ``final5/{normal,allocator}`` executables are authenticated from the build
receipts and are retained.  Every other path is checked without following
links, and a live process referring to a deletion candidate by cwd, executable
or file descriptor aborts the operation. The audit covers the current user;
other users and protected session daemons are recorded as excluded.  A live gate/capture process is also
an abort condition even when it has not opened a candidate yet.  The default
mode mutates only those exact cache roots; ``--dry-run`` performs all custody
checks and prints the plan without mutating either cache or the evidence tree.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import stat
import sys
from typing import Any, Iterable, Mapping


sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parent
TEMP = Path("/home/zhuhe/.cache/litchi-goal-0491")
TARGET = Path("/home/zhuhe/.cache/litchi-build-0491")
BUILD_SCHEMA = "docx-provider-lifecycle-build-v1"
SCHEMA = "docx-office-0491-cleanup-v1"
VERSION = 1
SEAL_NAME = "cleanup.json"
FINAL_ATTEMPT = "final5"
BUILD_ROLES = {
    "normal": "litchi-perf-baseline",
    "allocator": "litchi-perf-baseline-alloc",
}

# These are the scripts which can launch or supervise a benchmark.  A Python
# process has the interpreter as /proc/PID/exe, so command-line inspection is
# needed in addition to checking the executable path.  The cleanup process is
# excluded by PID while it scans itself.
ROOT_PROCESS_SCRIPTS = frozenset(
    {
        "gate.py",
        "provider_matrix.py",
        "cold_matrix.py",
        "profile_providers.py",
        "record_environment.py",
        "retain_build.py",
        "test_provider_matrix.py",
        "test_cold_matrix.py",
        "test_profile_providers.py",
    }
)
_DELETED_SUFFIX = " (deleted)"
_BLOCK_SIZE = 512


class CleanupError(RuntimeError):
    """A cleanup custody or safety precondition failed."""


def fail(message: str) -> None:
    raise CleanupError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _sha256(path: Path) -> str:
    try:
        with path.open("rb") as stream:
            return hashlib.file_digest(stream, "sha256").hexdigest()
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    raise AssertionError("unreachable")


def _lstat(path: Path, label: str) -> os.stat_result:
    try:
        state = path.lstat()
    except OSError as error:
        fail(f"{label}: cannot stat {path}: {error}")
    return state


def _regular_file(path: Path, label: str) -> os.stat_result:
    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is not allowed: {path}")
    require(stat.S_ISREG(state.st_mode), f"{label}: regular file is required: {path}")
    return state


def _directory(path: Path, label: str) -> os.stat_result:
    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is not allowed: {path}")
    require(stat.S_ISDIR(state.st_mode), f"{label}: directory is required: {path}")
    return state


def _no_symlink_components(path: Path, label: str) -> None:
    """Reject a symlink anywhere in an absolute path's existing prefix."""

    absolute = Path(os.path.abspath(path))
    current = Path(absolute.anchor)
    for component in absolute.parts[1:]:
        current /= component
        try:
            state = current.lstat()
        except FileNotFoundError:
            # Callers separately authenticate the leaf.  A missing suffix is
            # harmless here and will be diagnosed by _lstat at the boundary.
            continue
        except OSError as error:
            fail(f"{label}: cannot inspect path component {current}: {error}")
        require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink path component: {current}")


def _canonical_directory(path: Path, label: str) -> Path:
    _no_symlink_components(path, label)
    _directory(path, label)
    try:
        resolved = path.resolve(strict=True)
    except OSError as error:
        fail(f"{label}: cannot resolve {path}: {error}")
    require(resolved.is_dir() and not resolved.is_symlink(), f"{label}: resolved path is not a directory: {resolved}")
    _no_symlink_components(resolved, label)
    return resolved


def _regular_meta(path: Path, label: str) -> dict[str, int | str]:
    state = _regular_file(path, label)
    return {"bytes": state.st_size, "sha256": _sha256(path)}


def _read_json(path: Path, label: str) -> Any:
    _regular_file(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        fail(f"{label}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def _descriptor(path: Path, label: str, *, root: Path | None = None) -> dict[str, Any]:
    _regular_file(path, label)
    resolved = path.resolve(strict=True)
    result: dict[str, Any] = {"path": str(resolved), **_regular_meta(resolved, label)}
    if root is not None:
        try:
            result["relative_path"] = resolved.relative_to(root).as_posix()
        except ValueError as error:
            fail(f"{label}: path is outside {root}: {resolved} ({error})")
    return result


def _relative(path: Path, root: Path, label: str) -> str:
    try:
        return path.resolve(strict=False).relative_to(root).as_posix()
    except ValueError as error:
        fail(f"{label}: {path} is outside {root}: {error}")
    raise AssertionError("unreachable")


def _allocated_bytes(state: os.stat_result) -> int:
    blocks = getattr(state, "st_blocks", None)
    if blocks is None:
        return ((state.st_size + _BLOCK_SIZE - 1) // _BLOCK_SIZE) * _BLOCK_SIZE
    require(type(blocks) is int and blocks >= 0, "filesystem returned an invalid allocated-block count")
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


def _tree_stats(path: Path, label: str, *, root_device: int | None = None) -> _TreeStats:
    """Inventory a regular tree without following links or mount points."""

    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is not allowed: {path}")
    require(stat.S_ISREG(state.st_mode) or stat.S_ISDIR(state.st_mode),
            f"{label}: only regular files and directories are allowed: {path}")
    if root_device is None:
        root_device = state.st_dev
    require(state.st_dev == root_device, f"{label}: mount/device boundary is not allowed: {path}")

    fingerprint = hashlib.sha256()
    seen: set[tuple[int, int]] = set()

    def visit(current: Path, relative: str) -> tuple[int, int, int, int]:
        current_state = _lstat(current, f"{label}/{relative or '.'}")
        require(not stat.S_ISLNK(current_state.st_mode),
                f"{label}: symlink is not allowed: {current}")
        require(stat.S_ISREG(current_state.st_mode) or stat.S_ISDIR(current_state.st_mode),
                f"{label}: non-regular path is not allowed: {current}")
        require(current_state.st_dev == root_device,
                f"{label}: mount/device boundary is not allowed: {current}")
        identity = (current_state.st_dev, current_state.st_ino)
        kind = "file" if stat.S_ISREG(current_state.st_mode) else "directory"
        allocated = _allocated_bytes(current_state)
        fingerprint.update(
            "\0".join(
                str(item)
                for item in (
                    relative,
                    kind,
                    current_state.st_size if kind == "file" else 0,
                    allocated,
                    current_state.st_mode,
                    current_state.st_mtime_ns,
                    current_state.st_dev,
                    current_state.st_ino,
                )
            ).encode("utf-8")
        )
        fingerprint.update(b"\n")
        unique_allocated = 0
        if identity not in seen:
            seen.add(identity)
            unique_allocated = allocated
        if stat.S_ISREG(current_state.st_mode):
            return 1, 0, current_state.st_size, unique_allocated

        files = directories = logical = allocated = 0
        allocated += unique_allocated
        try:
            children = sorted(current.iterdir(), key=lambda item: item.name)
        except OSError as error:
            fail(f"{label}: cannot enumerate {current}: {error}")
        for child in children:
            child_files, child_directories, child_logical, child_allocated = visit(
                child, f"{relative}/{child.name}" if relative else child.name
            )
            files += child_files
            directories += child_directories
            logical += child_logical
            allocated += child_allocated
        return files, directories + 1, logical, allocated

    files, directories, logical, allocated = visit(path, "")
    return _TreeStats(
        path=path,
        kind="directory" if stat.S_ISDIR(state.st_mode) else "file",
        files=files,
        directories=directories,
        logical_bytes=logical,
        allocated_bytes=allocated,
        identity=(state.st_dev, state.st_ino),
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
        "fingerprint": stats.fingerprint,
    }


def _path_inside(path: Path, root: Path) -> bool:
    try:
        path.resolve(strict=False).relative_to(root.resolve(strict=False))
    except ValueError:
        return False
    return True


def _assert_disjoint_roots(root: Path, temp: Path, target: Path) -> None:
    require(root not in (temp, target), "evidence root cannot be a cleanup root")
    require(not _path_inside(root, temp) and not _path_inside(temp, root),
            "evidence root and temporary root overlap")
    require(not _path_inside(root, target) and not _path_inside(target, root),
            "evidence root and build target overlap")
    require(not _path_inside(temp, target) and not _path_inside(target, temp),
            "temporary root and build target overlap")


def _validate_terminal_receipts(root: Path) -> dict[str, Any]:
    validation = root / "validation"
    _directory(validation, "validation receipt directory")
    receipts: list[dict[str, Any]] = []
    unfinished: list[str] = []
    for started in sorted(validation.glob("*.started.json"), key=lambda item: item.name):
        _regular_file(started, f"validation start receipt {started.name}")
        result = validation / f"{started.name.removesuffix('.started.json')}.json"
        if result.is_symlink() or not result.exists():
            unfinished.append(started.name)
            continue
        value = _read_json(result, f"validation receipt {result.name}")
        require(isinstance(value, Mapping), f"validation receipt {result.name}: expected an object")
        require(type(value.get("exit_code")) is int,
                f"validation receipt {result.name}: exit_code is missing")
        require(isinstance(value.get("finished_utc"), str) and value["finished_utc"],
                f"validation receipt {result.name}: finished_utc is missing")
        receipts.append(_descriptor(result, f"validation receipt {result.name}", root=root))
    require(not unfinished, f"validation jobs are not terminal: {unfinished}")
    return {"count": len(receipts), "receipts": receipts}


def _path_from_receipt(value: Any, label: str, *, base: Path | None = None) -> Path:
    require(isinstance(value, Mapping), f"{label}: descriptor is missing")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path: missing")
    path = Path(raw)
    if not path.is_absolute():
        require(base is not None, f"{label}.path: relative path has no base")
        path = base / path
    _no_symlink_components(path, label)
    return path


def _validate_build(
    role: str,
    root: Path,
    temp: Path,
    target: Path,
    *,
    target_present: bool = True,
) -> dict[str, Any]:
    receipt_path = root / f"build-{role}.json"
    value = _read_json(receipt_path, f"{role} build receipt")
    require(isinstance(value, Mapping), f"{role} build receipt: expected an object")
    require(value.get("schema") == BUILD_SCHEMA,
            f"{role} build receipt: schema differs")
    require(value.get("version") == VERSION and value.get("role") == role,
            f"{role} build receipt: role or version differs")
    require(value.get("attempt") == FINAL_ATTEMPT,
            f"{role} build receipt: retained binary is not from {FINAL_ATTEMPT}")
    require(value.get("source_unchanged") is True and value.get("source_before") == value.get("source_after"),
            f"{role} build receipt: source custody failed")

    binary_value = value.get("binary")
    binary_path = _path_from_receipt(binary_value, f"{role} retained binary")
    expected_binary = temp / FINAL_ATTEMPT / role / BUILD_ROLES[role]
    _no_symlink_components(expected_binary, f"{role} retained binary destination")
    require(binary_path.resolve(strict=True) == expected_binary.resolve(strict=True),
            f"{role} retained binary path is not the exact final5 path")
    actual_state = _regular_file(binary_path, f"{role} retained binary")
    actual = {"bytes": actual_state.st_size, "sha256": _sha256(binary_path)}
    require(type(binary_value.get("bytes")) is int and binary_value["bytes"] == actual["bytes"],
            f"{role} retained binary byte length changed")
    require(binary_value.get("sha256") == actual["sha256"],
            f"{role} retained binary SHA-256 changed")
    require(binary_value.get("executable") is True and os.access(binary_path, os.X_OK),
            f"{role} retained binary executable binding failed")

    final_root = temp / FINAL_ATTEMPT
    _directory(final_root, "final5 retained root")
    role_root = final_root / role
    _directory(role_root, f"final5 {role} retained directory")
    expected_children = {BUILD_ROLES[role]}
    actual_children = {child.name for child in role_root.iterdir()}
    require(actual_children == expected_children,
            f"final5 {role} directory contains unexpected paths: {sorted(actual_children - expected_children)}")
    original_value = value.get("original_binary")
    if target_present:
        original = _validate_original_binary(value, role, target)
    else:
        # The target is deliberately absent after cleanup.  Retain the
        # authenticated historical binding without trying to stat a removed
        # Cargo artifact.
        original_path = _path_from_receipt(original_value, f"{role} original Cargo binary")
        try:
            original_resolved = original_path.resolve(strict=False)
            original_resolved.relative_to(target.resolve(strict=False))
        except ValueError as error:
            fail(f"{role} original Cargo binary is outside the owned build target: {error}")
        require(original_resolved != target.resolve(strict=False),
                f"{role} original Cargo binary points at the target root")
        require(original_resolved.name == BUILD_ROLES[role],
                f"{role} original Cargo binary name differs")
        require(isinstance(original_value, Mapping),
                f"{role} original Cargo binary descriptor is missing")
        require(type(original_value.get("bytes")) is int and original_value["bytes"] >= 0,
                f"{role} original Cargo binary byte length is malformed")
        original_sha = original_value.get("sha256")
        require(isinstance(original_sha, str) and len(original_sha) == 64 and
                all(character in "0123456789abcdef" for character in original_sha),
                f"{role} original Cargo binary SHA-256 is malformed")
        require(original_value.get("executable") is True,
                f"{role} original Cargo binary executable binding failed")
        original = {
            "path": str(original_resolved),
            "bytes": original_value["bytes"],
            "sha256": original_sha,
            "executable": True,
        }
    return {
        "role": role,
        "build_receipt": _descriptor(receipt_path, f"{role} build receipt", root=root),
        "binary": {
            "path": str(binary_path.resolve()),
            **actual,
            "executable": True,
        },
        "original_binary": original,
    }


def _validate_original_binary(value: Mapping[str, Any], role: str, target: Path) -> dict[str, Any]:
    original_value = value.get("original_binary")
    path = _path_from_receipt(original_value, f"{role} original Cargo binary")
    try:
        resolved = path.resolve(strict=True)
        resolved.relative_to(target.resolve(strict=True))
    except (OSError, ValueError) as error:
        fail(f"{role} original Cargo binary is outside the owned build target: {path} ({error})")
    require(resolved != target.resolve(strict=True),
            f"{role} original Cargo binary points at the target root")
    require(resolved.name == BUILD_ROLES[role],
            f"{role} original Cargo binary name differs")
    state = _regular_file(resolved, f"{role} original Cargo binary")
    actual = {"bytes": state.st_size, "sha256": _sha256(resolved)}
    require(type(original_value.get("bytes")) is int and original_value["bytes"] == actual["bytes"],
            f"{role} original Cargo binary byte length changed")
    require(original_value.get("sha256") == actual["sha256"],
            f"{role} original Cargo binary SHA-256 changed")
    require(original_value.get("executable") is True and os.access(resolved, os.X_OK),
            f"{role} original Cargo binary executable binding failed")
    return {"path": str(resolved), **actual, "executable": True}


def _validate_retained_layout(temp: Path, builds: Mapping[str, Mapping[str, Any]]) -> None:
    final_root = temp / FINAL_ATTEMPT
    actual_final_children = {child.name for child in final_root.iterdir()}
    require(actual_final_children == set(BUILD_ROLES),
            f"final5 retained root contains unexpected paths: {sorted(actual_final_children - set(BUILD_ROLES))}")
    for role in BUILD_ROLES:
        expected = Path(str(builds[role]["binary"]["path"])).resolve(strict=True)
        require(expected.parent == (final_root / role).resolve(strict=True),
                f"{role} retained binary parent differs from final5 role directory")


def _proc_path(raw: str) -> Path | None:
    if not raw or not raw.startswith("/"):
        return None
    if raw.endswith(_DELETED_SUFFIX):
        raw = raw[: -len(_DELETED_SUFFIX)]
    try:
        return Path(raw).resolve(strict=False)
    except OSError:
        return None


def _read_proc_link(path: Path, label: str) -> Path | None:
    try:
        raw = os.readlink(path)
    except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
        return None
    except PermissionError as error:
        fail(f"{label}: cannot inspect process link: {error}")
    except OSError as error:
        fail(f"{label}: cannot inspect process link: {error}")
    return _proc_path(os.fsdecode(raw))


def _read_proc_cmdline(path: Path, label: str) -> list[str]:
    try:
        data = path.read_bytes()
    except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
        return []
    except PermissionError as error:
        fail(f"{label}: cannot inspect process command line: {error}")
    except OSError as error:
        fail(f"{label}: cannot inspect process command line: {error}")
    return [os.fsdecode(item) for item in data.split(b"\0") if item]


def _process_audit(
    deletion_roots: Iterable[Path],
    executable_roots: Iterable[Path],
    evidence_root: Path,
    *,
    proc_root: Path = Path("/proc"),
    self_pid: int | None = None,
) -> dict[str, Any]:
    """Inspect live Linux processes without requiring psutil.

    A process disappearing between directory enumeration and link inspection
    is harmless and recorded as a race.  Permission failures are fatal: a
    cleanup that cannot see an owned-user process must not delete its possible workspace.
    Other users and protected session daemons are outside this user-owned
    benchmark-job audit and explicitly recorded; their inspectable descendants
    are still checked. This is not a claim of visibility into every host fd.
    """

    _directory(proc_root, "process information root")
    deletion = tuple(path.resolve(strict=False) for path in deletion_roots)
    executable = tuple(path.resolve(strict=False) for path in executable_roots)
    self_pid = os.getpid() if self_pid is None else self_pid
    busy: list[dict[str, Any]] = []
    gate_processes: list[dict[str, Any]] = []
    vanished = 0
    scanned = 0
    excluded_other_uid = []
    excluded_session_daemons = []
    uid_scope = os.getuid()
    try:
        entries = sorted(proc_root.iterdir(), key=lambda item: item.name)
    except OSError as error:
        fail(f"process information root: cannot enumerate: {error}")

    script_names = set(ROOT_PROCESS_SCRIPTS)
    for process_dir in entries:
        if not process_dir.name.isdecimal():
            continue
        try:
            pid = int(process_dir.name)
        except ValueError:
            continue
        if pid == self_pid:
            continue
        try:
            process_uid = process_dir.stat().st_uid
        except (FileNotFoundError, ProcessLookupError):
            vanished += 1
            continue
        if process_uid != uid_scope:
            excluded_other_uid.append({"pid": pid, "uid": process_uid})
            continue
        if not os.access(process_dir / "fd", os.R_OK | os.X_OK):
            try:
                comm = (process_dir / "comm").read_text().strip()
            except (FileNotFoundError, ProcessLookupError):
                vanished += 1
                continue
            if comm in {"systemd", "(sd-pam)", "sshd-session"}:
                excluded_session_daemons.append({"pid": pid, "comm": comm,
                    "reason": "protected session daemon; process links not inspectable; descendants audited separately"})
                continue
        scanned += 1
        try:
            cwd = _read_proc_link(process_dir / "cwd", f"/proc/{pid}/cwd")
            exe = _read_proc_link(process_dir / "exe", f"/proc/{pid}/exe")
        except CleanupError as error:
            if isinstance(error.__context__, PermissionError):
                comm = (process_dir / "comm").read_text().strip()
                if comm in {"systemd", "(sd-pam)", "sshd-session"}:
                    excluded_session_daemons.append({"pid": pid, "comm": comm,
                        "reason": "protected session daemon; process links not inspectable; descendants audited separately"})
                    continue
            raise
        try:
            fd_entries = sorted((process_dir / "fd").iterdir(), key=lambda item: item.name)
        except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
            vanished += 1
            continue
        except PermissionError as error:
            fail(f"/proc/{pid}/fd: cannot enumerate: {error}")
        except OSError as error:
            fail(f"/proc/{pid}/fd: cannot enumerate: {error}")
        fd_paths: list[tuple[str, Path]] = []
        for fd in fd_entries:
            target = _read_proc_link(fd, f"/proc/{pid}/fd/{fd.name}")
            if target is not None:
                fd_paths.append((fd.name, target))

        references: list[tuple[str, Path]] = []
        if cwd is not None:
            references.append(("cwd", cwd))
        if exe is not None:
            references.append(("exe", exe))
        references.extend((f"fd:{fd}", target) for fd, target in fd_paths)
        for kind, target in references:
            for candidate in deletion:
                if _path_inside(target, candidate):
                    busy.append(
                        {
                            "pid": pid,
                            "reference": kind,
                            "path": str(target),
                            "candidate": str(candidate),
                        }
                    )
                    break

        command = _read_proc_cmdline(process_dir / "cmdline", f"/proc/{pid}/cmdline")
        command_text = " ".join(command)
        script_match = False
        for argument in command:
            if Path(argument).name not in script_names:
                continue
            if argument.startswith("/"):
                candidate_script = _proc_path(argument)
                script_match = candidate_script is not None and _path_inside(
                    candidate_script, evidence_root
                )
            elif cwd is not None:
                # gate.py launches scripts using paths relative to the
                # repository checkout. Resolve against the process cwd before
                # declaring the argument to be a root script.
                candidate_script = _proc_path(str(cwd / argument))
                script_match = candidate_script is not None and _path_inside(
                    candidate_script, evidence_root
                )
            if script_match:
                break
        executable_match = exe is not None and any(_path_inside(exe, candidate) for candidate in executable)
        if script_match or executable_match:
            gate_processes.append(
                {
                    "pid": pid,
                    "exe": None if exe is None else str(exe),
                    "command": command_text,
                    "script_reference": script_match,
                    "executable_reference": executable_match,
                }
            )

    return {
        "proc_root": str(proc_root.resolve(strict=False)),
        "self_pid": self_pid,
        "scanned_processes": scanned,
        "uid_scope": uid_scope,
        "excluded_other_uid_processes": excluded_other_uid,
        "excluded_session_daemons": excluded_session_daemons,
        "vanished_processes": vanished,
        "candidate_references": busy,
        "root_gate_capture_processes": gate_processes,
        "safe": not busy and not gate_processes,
    }


def _candidate_paths(temp: Path, target: Path) -> tuple[list[Path], Path]:
    _directory(temp, "temporary root")
    _directory(target, "Cargo build target")
    final_root = temp / FINAL_ATTEMPT
    _directory(final_root, "final5 retained root")
    candidates: list[Path] = []
    try:
        children = sorted(temp.iterdir(), key=lambda item: item.name)
    except OSError as error:
        fail(f"temporary root: cannot enumerate: {error}")
    for child in children:
        if child.name == FINAL_ATTEMPT:
            continue
        # Authenticate each candidate before it can enter the deletion plan;
        # this catches a link at the top level as well as links below it.
        _tree_stats(child, f"temporary candidate {child.name}")
        candidates.append(child)
    _tree_stats(target, "Cargo build target")
    return candidates, target


def _plan_cleanup(
    *,
    root: Path = ROOT,
    temp: Path = TEMP,
    target: Path = TARGET,
    proc_root: Path = Path("/proc"),
    self_pid: int | None = None,
) -> dict[str, Any]:
    evidence_root = _canonical_directory(root, "evidence root")
    temporary_root = _canonical_directory(temp, "temporary root")
    target_root = _canonical_directory(target, "Cargo build target")
    _assert_disjoint_roots(evidence_root, temporary_root, target_root)
    require(temporary_root.stat().st_uid == os.getuid() and target_root.stat().st_uid == os.getuid(),
            "cleanup cache roots must belong to the audited user")
    cleanup_path = evidence_root / SEAL_NAME
    require(not cleanup_path.exists() and not cleanup_path.is_symlink(),
            f"refusing to replace existing cleanup receipt: {cleanup_path}")

    terminal = _validate_terminal_receipts(evidence_root)
    builds = {
        role: _validate_build(role, evidence_root, temporary_root, target_root)
        for role in BUILD_ROLES
    }
    _validate_retained_layout(temporary_root, builds)
    candidates, target_candidate = _candidate_paths(temporary_root, target_root)
    candidate_stats = [
        _tree_stats(path, f"temporary candidate {path.name}")
        for path in candidates
    ]
    target_stats = _tree_stats(target_candidate, "Cargo build target")
    removed_stats = [
        _stats_record(stats, root=temporary_root, label="temporary candidate")
        for stats in candidate_stats
    ]
    removed_stats.append(_stats_record(target_stats, root=target_root.parent, label="Cargo build target"))
    process = _process_audit(
        [*candidates, target_candidate],
        [temporary_root, target_root],
        evidence_root,
        proc_root=proc_root,
        self_pid=self_pid,
    )
    require(process["safe"], _process_failure(process))
    return {
        "root": evidence_root,
        "temp": temporary_root,
        "target": target_root,
        "terminal": terminal,
        "builds": builds,
        "candidates": candidates,
        "target_candidate": target_candidate,
        "removed_stats": removed_stats,
        "process_before": process,
        "proc_root": proc_root,
        "self_pid": self_pid,
    }


def _process_failure(process: Mapping[str, Any]) -> str:
    candidate = process.get("candidate_references", [])
    gate = process.get("root_gate_capture_processes", [])
    if candidate:
        return f"active process references cleanup candidate: {candidate}"
    if gate:
        return f"live root gate/capture process remains: {gate}"
    return "process audit failed"


def _same_stats(path: Path, expected: Mapping[str, Any], label: str, *, root: Path) -> _TreeStats:
    actual = _tree_stats(path, label)
    current = _stats_record(actual, root=root, label=label)
    for key in ("kind", "files", "directories", "logical_bytes", "allocated_bytes", "fingerprint"):
        require(current[key] == expected[key], f"{label}: changed after planning ({key})")
    return actual


def _remove_tree(path: Path, label: str) -> None:
    """Remove an authenticated regular tree without ever following links."""

    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: refusing to remove symlink: {path}")
    require(stat.S_ISREG(state.st_mode) or stat.S_ISDIR(state.st_mode),
            f"{label}: refusing to remove non-regular path: {path}")
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
        _remove_tree(child, f"{label}/{child.name}")
    try:
        path.rmdir()
    except OSError as error:
        fail(f"{label}: cannot remove directory {path}: {error}")


def _revalidate_retained(builds: Mapping[str, Mapping[str, Any]]) -> None:
    for role, record in builds.items():
        binary = record["binary"]
        path = Path(str(binary["path"]))
        actual = _regular_meta(path, f"{role} retained binary")
        require(actual == {"bytes": binary["bytes"], "sha256": binary["sha256"]},
                f"{role} retained binary changed during cleanup")
        require(os.access(path, os.X_OK), f"{role} retained binary lost executable permission")


def _totals(records: Iterable[Mapping[str, Any]]) -> dict[str, int]:
    values = list(records)
    return {
        "paths": len(values),
        "files": sum(int(value["files"]) for value in values),
        "directories": sum(int(value["directories"]) for value in values),
        "logical_bytes": sum(int(value["logical_bytes"]) for value in values),
        "allocated_bytes": sum(int(value["allocated_bytes"]) for value in values),
    }


def _receipt(plan: Mapping[str, Any], after: Mapping[str, Any]) -> dict[str, Any]:
    root = Path(plan["root"])
    driver = _descriptor(Path(__file__).resolve(), "cleanup driver")
    removed = list(plan["removed_stats"])
    return {
        "schema": SCHEMA,
        "version": VERSION,
        "status": "pass",
        "completed_utc": _now(),
        "driver": driver,
        "scope": {
            "evidence_root": str(root),
            "temporary_root": str(plan["temp"]),
            "build_target": str(plan["target"]),
            "retained_subtree": str((Path(plan["temp"]) / FINAL_ATTEMPT).resolve()),
            "root_evidence_preserved": True,
            "iwork_tree_touched": False,
        },
        "terminal_validation": plan["terminal"],
        "retained_binaries": [
            plan["builds"][role]["binary"] | {
                "role": role,
                "build_receipt": plan["builds"][role]["build_receipt"],
            }
            for role in BUILD_ROLES
        ],
        "removed": removed,
        "removed_paths": removed,
        "removed_totals": _totals(removed),
        "target_removed": True,
        "temporary_scratch_remaining": [FINAL_ATTEMPT],
        "process_audit": {
            "before": plan["process_before"],
            "after": after,
        },
    }


def _write_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    require(not path.exists() and not path.is_symlink(),
            f"refusing to replace existing cleanup receipt: {path}")
    try:
        with path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
    except OSError as error:
        fail(f"cannot write cleanup receipt {path}: {error}")


def execute_cleanup(plan: Mapping[str, Any]) -> dict[str, Any]:
    """Execute a previously authenticated plan and write cleanup.json."""

    root = Path(plan["root"])
    temp = Path(plan["temp"])
    target = Path(plan["target"])
    cleanup_path = root / SEAL_NAME
    require(not cleanup_path.exists() and not cleanup_path.is_symlink(),
            f"refusing to replace existing cleanup receipt: {cleanup_path}")
    _revalidate_retained(plan["builds"])
    before = _process_audit(
        [*plan["candidates"], plan["target_candidate"]],
        [temp, target],
        root,
        proc_root=Path(plan["proc_root"]),
        self_pid=plan["self_pid"],
    )
    require(before["safe"], _process_failure(before))

    expected_by_path = {
        str(Path(item["path"]).resolve(strict=False)): item
        for item in plan["removed_stats"]
    }
    for path in [*plan["candidates"], plan["target_candidate"]]:
        expected = expected_by_path[str(Path(path).resolve(strict=False))]
        _same_stats(
            path,
            expected,
            f"cleanup candidate {path}",
            root=temp if path in plan["candidates"] else target.parent,
        )

    for path in [*plan["candidates"], plan["target_candidate"]]:
        _remove_tree(path, f"remove {path}")
        require(not path.exists() and not path.is_symlink(),
                f"cleanup candidate remains after removal: {path}")

    _revalidate_retained(plan["builds"])
    after = _process_audit(
        [*plan["candidates"], plan["target_candidate"]],
        [temp, target],
        root,
        proc_root=Path(plan["proc_root"]),
        self_pid=plan["self_pid"],
    )
    require(after["safe"], _process_failure(after))
    receipt = _receipt({**plan, "process_before": before}, after)
    _write_exclusive(cleanup_path, receipt)
    return receipt


def _absent(path: Path, label: str) -> None:
    _no_symlink_components(path, label)
    require(not path.exists() and not path.is_symlink(), f"{label}: path remains: {path}")


def _verify_removed_records(
    records: Any,
    *,
    temp: Path,
    target: Path,
) -> dict[str, Any]:
    require(isinstance(records, list) and records,
            "cleanup receipt removed inventory is missing")
    seen: set[str] = set()
    target_path = str(target.resolve(strict=False))
    target_record: Mapping[str, Any] | None = None
    for index, record in enumerate(records):
        require(isinstance(record, Mapping), f"cleanup.removed[{index}]: expected an object")
        raw_path = record.get("path")
        require(isinstance(raw_path, str) and raw_path,
                f"cleanup.removed[{index}].path: missing")
        path = Path(raw_path)
        _no_symlink_components(path, f"cleanup.removed[{index}]")
        resolved = path.resolve(strict=False)
        canonical = str(resolved)
        require(canonical not in seen, f"cleanup.removed[{index}]: duplicate path")
        seen.add(canonical)
        if canonical == target_path:
            target_record = record
        else:
            require(_path_inside(resolved, temp),
                    f"cleanup.removed[{index}]: path escapes owned roots: {path}")
            require(not _path_inside(resolved, temp / FINAL_ATTEMPT),
                    f"cleanup.removed[{index}]: retained final5 path is listed for deletion: {path}")
        _absent(path, f"cleanup.removed[{index}]")
        for key in ("files", "directories", "logical_bytes", "allocated_bytes"):
            require(type(record.get(key)) is int and record[key] >= 0,
                    f"cleanup.removed[{index}].{key}: malformed")
        fingerprint = record.get("fingerprint")
        require(isinstance(fingerprint, str) and len(fingerprint) == 64 and
                all(character in "0123456789abcdef" for character in fingerprint),
                f"cleanup.removed[{index}].fingerprint: malformed")
    require(target_record is not None,
            "cleanup.removed inventory does not include the exact build target")
    return {"records": records, "target": target_record}


def verify(
    *,
    root: Path = ROOT,
    temp: Path = TEMP,
    target: Path = TARGET,
    proc_root: Path = Path("/proc"),
) -> dict[str, Any]:
    """Recheck a completed cleanup for the evidence seal.

    This is read-only.  It requires the Cargo target and every disposable temp
    child to be gone, then authenticates the final5 binaries, cleanup receipt,
    terminal validation inventory, and a fresh process audit.  The returned
    ``remaining`` list is the canonical empty assertion used by ``seal.py``.
    """

    evidence_root = _canonical_directory(root, "evidence root")
    temporary_root = _canonical_directory(temp, "temporary root")
    target_path = Path(target)
    _no_symlink_components(target_path, "Cargo build target")
    _assert_disjoint_roots(evidence_root, temporary_root, target_path.resolve(strict=False))
    _absent(target_path, "Cargo build target")

    cleanup_path = evidence_root / SEAL_NAME
    receipt = _read_json(cleanup_path, "cleanup receipt")
    require(isinstance(receipt, Mapping), "cleanup receipt: expected an object")
    require(receipt.get("schema") == SCHEMA and receipt.get("version") == VERSION,
            "cleanup receipt: schema or version differs")
    require(receipt.get("status") == "pass", "cleanup receipt: status is not pass")

    builds = {
        role: _validate_build(
            role,
            evidence_root,
            temporary_root,
            target_path,
            target_present=False,
        )
        for role in BUILD_ROLES
    }
    _validate_retained_layout(temporary_root, builds)
    require(
        {child.name for child in temporary_root.iterdir()} == {FINAL_ATTEMPT},
        "temporary root retains disposable paths",
    )

    expected_scope = {
        "evidence_root": str(evidence_root),
        "temporary_root": str(temporary_root),
        "build_target": str(target_path.resolve(strict=False)),
        "retained_subtree": str((temporary_root / FINAL_ATTEMPT).resolve(strict=True)),
        "root_evidence_preserved": True,
        "iwork_tree_touched": False,
    }
    require(receipt.get("scope") == expected_scope,
            "cleanup receipt: cleanup scope binding differs")
    expected_retained = [
        builds[role]["binary"] | {
            "role": role,
            "build_receipt": builds[role]["build_receipt"],
        }
        for role in BUILD_ROLES
    ]
    require(receipt.get("retained_binaries") == expected_retained,
            "cleanup receipt: retained binary inventory differs")
    terminal = _validate_terminal_receipts(evidence_root)
    require(receipt.get("terminal_validation") == terminal,
            "cleanup receipt: terminal validation inventory differs")
    removed = _verify_removed_records(receipt.get("removed"), temp=temporary_root, target=target_path)
    require(receipt.get("removed_paths") == removed["records"],
            "cleanup receipt: removed path inventory differs")
    totals = receipt.get("removed_totals")
    require(totals == _totals(removed["records"]),
            "cleanup receipt: removed totals differ from inventory")
    require(receipt.get("target_removed") is True and
            receipt.get("temporary_scratch_remaining") == [FINAL_ATTEMPT],
            "cleanup receipt: remaining-root binding differs")

    driver = receipt.get("driver")
    require(isinstance(driver, Mapping) and driver.get("path") == str(Path(__file__).resolve()),
            "cleanup receipt: driver binding differs")
    current_driver = _descriptor(Path(__file__).resolve(), "cleanup driver")
    require(dict(driver) == current_driver,
            "cleanup receipt: cleanup driver changed")

    current_process = _process_audit(
        [Path(str(record["path"])) for record in removed["records"]],
        [temporary_root, target_path],
        evidence_root,
        proc_root=proc_root,
    )
    require(current_process["safe"], _process_failure(current_process))
    process_receipt = receipt.get("process_audit")
    require(isinstance(process_receipt, Mapping),
            "cleanup receipt: process audit is missing")
    for phase in ("before", "after"):
        phase_record = process_receipt.get(phase)
        require(isinstance(phase_record, Mapping) and phase_record.get("safe") is True,
                f"cleanup receipt: process audit {phase} was not safe")
        require(phase_record.get("candidate_references") == [] and
                phase_record.get("root_gate_capture_processes") == [],
                f"cleanup receipt: process audit {phase} retained a live reference")

    return {
        "schema": "docx-office-0491-cleanup-verification-v1",
        "version": VERSION,
        "status": "pass",
        "cleanup_receipt": _descriptor(cleanup_path, "cleanup receipt", root=evidence_root),
        "retained_binaries": expected_retained,
        "build_target_removed": True,
        "temporary_root": str(temporary_root),
        "remaining": [],
        "process_audit": current_process,
    }


def plan_cleanup(**kwargs: Any) -> dict[str, Any]:
    """Public read-only planning entry point used by isolated tests."""

    return _plan_cleanup(**kwargs)


def _print_plan(plan: Mapping[str, Any]) -> None:
    print(
        json.dumps(
            {
                "schema": SCHEMA,
                "status": "ready",
                "dry_run": True,
                "removed": plan["removed_stats"],
                "removed_totals": _totals(plan["removed_stats"]),
                "retained_binaries": [plan["builds"][role]["binary"] for role in BUILD_ROLES],
                "process_audit": plan["process_before"],
            },
            indent=2,
            sort_keys=True,
        )
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--dry-run", action="store_true", help="validate and print the plan without deleting or writing a receipt")
    modes.add_argument("--verify", action="store_true", help="verify an already completed cleanup without mutating anything")
    parser.add_argument("--root", type=Path, default=ROOT, help=argparse.SUPPRESS)
    parser.add_argument("--temp-root", type=Path, default=TEMP, help=argparse.SUPPRESS)
    parser.add_argument("--target-root", type=Path, default=TARGET, help=argparse.SUPPRESS)
    parser.add_argument("--proc-root", type=Path, default=Path("/proc"), help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    try:
        configured = (args.root.resolve(), args.temp_root.resolve(), args.target_root.resolve())
        fixed = (ROOT.resolve(), TEMP.resolve(), TARGET.resolve())
        if not args.dry_run and not args.verify:
            require(configured == fixed,
                    "destructive mode is restricted to the fixed 0491 evidence and cache roots")
        if args.verify:
            result = verify(
                root=args.root,
                temp=args.temp_root,
                target=args.target_root,
                proc_root=args.proc_root,
            )
            print(json.dumps(result, indent=2, sort_keys=True))
            return 0
        plan = _plan_cleanup(
            root=args.root,
            temp=args.temp_root,
            target=args.target_root,
            proc_root=args.proc_root,
        )
        if args.dry_run:
            _print_plan(plan)
            return 0
        receipt = execute_cleanup(plan)
        print(json.dumps({"schema": SCHEMA, "status": receipt["status"], "removed": receipt["removed_totals"]}, sort_keys=True))
        return 0
    except (CleanupError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"cleanup.py: error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
