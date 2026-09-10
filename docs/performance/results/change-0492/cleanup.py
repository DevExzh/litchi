#!/usr/bin/env python3
"""Bounded cleanup for the 0492 read-ahead evidence run.

Only the explicitly owned Cargo target and disposable children of the 0492
cache are eligible for removal.  The authenticated ``final2`` binaries are
retained.  Every candidate is inventoried without following links, checked
for ownership and device boundaries, rechecked immediately before removal,
and removed without a recursive shell command.  Dry-run and verify modes are
read-only.  The evidence tree and ``litchi-spec-gaps`` are never cleanup
targets.
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
TEMP = Path("/home/zhuhe/.cache/litchi-goal-0492")
TARGET = Path("/home/zhuhe/.cache/litchi-build-0492")
FINAL_ATTEMPT = "final2"
BUILD_SCHEMA = "docx-provider-lifecycle-build-v1"
SCHEMA = "docx-read-ahead-cleanup-v1"
VERSION = 1
BUILD_ROLES = {
    "normal": "litchi-perf-baseline",
    "allocator": "litchi-perf-baseline-alloc",
}
PROTECTED_DAEMONS = frozenset({"systemd", "(sd-pam)", "sshd-session"})
ROOT_PROCESS_SCRIPTS = frozenset(
    {
        "cleanup.py",
        "gate.py",
        "measure.py",
        "retain_build.py",
        "test_measure.py",
        "verify_bundle.py",
    }
)
DELETED_SUFFIX = " (deleted)"
BLOCK_SIZE = 512


class CleanupError(RuntimeError):
    """A custody or safety precondition failed."""


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
        return path.lstat()
    except OSError as error:
        fail(f"{label}: cannot stat {path}: {error}")
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
        require(state.st_uid == owner_uid, f"{label}: unexpected owner for {path}")
    return state


def _directory(path: Path, label: str, *, owner_uid: int | None = None) -> os.stat_result:
    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is not allowed: {path}")
    require(stat.S_ISDIR(state.st_mode), f"{label}: directory required: {path}")
    if owner_uid is not None:
        require(state.st_uid == owner_uid, f"{label}: unexpected owner for {path}")
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
        return ((state.st_size + BLOCK_SIZE - 1) // BLOCK_SIZE) * BLOCK_SIZE
    require(type(blocks) is int and blocks >= 0, "invalid allocated block count")
    return blocks * BLOCK_SIZE


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


def _tree_stats(path: Path, label: str, *, owner_uid: int | None = None) -> _TreeStats:
    """Inventory a regular tree without following links, devices, or mounts."""

    root_state = _lstat(path, label)
    require(not stat.S_ISLNK(root_state.st_mode), f"{label}: symlink is not allowed: {path}")
    require(
        stat.S_ISREG(root_state.st_mode) or stat.S_ISDIR(root_state.st_mode),
        f"{label}: only regular files and directories are allowed: {path}",
    )
    root_device = root_state.st_dev
    seen: set[tuple[int, int]] = set()
    fingerprint = hashlib.sha256()

    def visit(current: Path, relative: str) -> tuple[int, int, int, int]:
        state = _lstat(current, f"{label}/{relative or '.'}")
        require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink is not allowed: {current}")
        require(
            stat.S_ISREG(state.st_mode) or stat.S_ISDIR(state.st_mode),
            f"{label}: non-regular path is not allowed: {current}",
        )
        require(state.st_dev == root_device, f"{label}: device boundary at {current}")
        if owner_uid is not None:
            require(state.st_uid == owner_uid, f"{label}: unexpected owner at {current}")
        kind = "file" if stat.S_ISREG(state.st_mode) else "directory"
        allocated = _allocated_bytes(state)
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
            ).encode()
        )
        fingerprint.update(b"\n")
        identity = (state.st_dev, state.st_ino)
        unique_allocated = 0
        if identity not in seen:
            seen.add(identity)
            unique_allocated = allocated
        if kind == "file":
            return 1, 0, state.st_size, unique_allocated

        files = directories = logical = allocated_total = 0
        allocated_total += unique_allocated
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
        "fingerprint": stats.fingerprint,
    }


def _meta(path: Path, label: str, *, owner_uid: int | None = None) -> dict[str, int | str]:
    state = _regular(path, label, owner_uid=owner_uid)
    return {"bytes": state.st_size, "sha256": _sha256(path)}


def _descriptor(path: Path, label: str, *, root: Path | None = None) -> dict[str, Any]:
    _regular(path, label)
    resolved = path.resolve(strict=True)
    value: dict[str, Any] = {"path": str(resolved), **_meta(resolved, label)}
    if root is not None:
        value["relative_path"] = _relative(resolved, root, label)
    return value


def _read_json(path: Path, label: str) -> Any:
    _regular(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        fail(f"{label}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def _path_descriptor(value: Any, label: str, *, base: Path | None = None) -> Path:
    require(isinstance(value, Mapping), f"{label}: descriptor missing")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path: missing")
    path = Path(raw)
    if not path.is_absolute():
        require(base is not None, f"{label}.path: relative path has no base")
        path = base / path
    require(path.is_absolute(), f"{label}.path: absolute path required")
    _no_symlink_components(path, label)
    return path


def _validate_terminal_receipts(root: Path) -> dict[str, Any]:
    validation = root / "validation"
    _directory(validation, "validation receipt directory")
    receipts: list[dict[str, Any]] = []
    unfinished: list[str] = []
    for started in sorted(validation.glob("*.started.json"), key=lambda item: item.name):
        _regular(started, f"validation start receipt {started.name}")
        result = validation / f"{started.name.removesuffix('.started.json')}.json"
        if result.is_symlink() or not result.exists():
            unfinished.append(started.name)
            continue
        value = _read_json(result, f"validation receipt {result.name}")
        require(isinstance(value, Mapping), f"validation receipt {result.name}: object required")
        require(type(value.get("exit_code")) is int, f"validation receipt {result.name}: exit_code missing")
        require(
            isinstance(value.get("finished_utc"), str) and value["finished_utc"],
            f"validation receipt {result.name}: finished_utc missing",
        )
        receipts.append(_descriptor(result, f"validation receipt {result.name}", root=root))
    require(not unfinished, f"validation jobs are not terminal: {unfinished}")
    require(receipts, "no terminal validation receipts were found")
    return {"count": len(receipts), "receipts": receipts}


def _binary_descriptor(value: Any, label: str, path: Path, *, present: bool) -> dict[str, Any]:
    require(isinstance(value, Mapping), f"{label}: descriptor missing")
    require(value.get("executable") is True, f"{label}: executable binding failed")
    expected_bytes = value.get("bytes")
    expected_sha = value.get("sha256")
    require(type(expected_bytes) is int and expected_bytes >= 0, f"{label}: byte length malformed")
    require(
        isinstance(expected_sha, str)
        and len(expected_sha) == 64
        and all(character in "0123456789abcdef" for character in expected_sha),
        f"{label}: SHA-256 malformed",
    )
    if present:
        actual = _meta(path, label)
        require(actual == {"bytes": expected_bytes, "sha256": expected_sha}, f"{label}: content changed")
        require(os.access(path, os.X_OK), f"{label}: executable bit missing")
    return {"path": str(path.resolve(strict=False)), "bytes": expected_bytes, "sha256": expected_sha}


def _validate_build(
    role: str,
    root: Path,
    temp: Path,
    target: Path,
    *,
    target_present: bool,
) -> dict[str, Any]:
    receipt_path = root / f"build-{role}.json"
    value = _read_json(receipt_path, f"{role} build receipt")
    require(isinstance(value, Mapping), f"{role} build receipt: object required")
    require(value.get("schema") == BUILD_SCHEMA and value.get("version") == VERSION, f"{role} build receipt: schema mismatch")
    require(value.get("role") == role and value.get("attempt") == FINAL_ATTEMPT, f"{role} build receipt: final2 binding missing")
    require(value.get("source_unchanged") is True and value.get("source_before") == value.get("source_after"), f"{role} build receipt: source custody failed")

    expected_binary = temp / FINAL_ATTEMPT / role / BUILD_ROLES[role]
    binary_path = _path_descriptor(value.get("binary"), f"{role} retained binary")
    require(binary_path.resolve(strict=False) == expected_binary.resolve(strict=False), f"{role} retained binary path differs")
    binary = _binary_descriptor(value.get("binary"), f"{role} retained binary", binary_path, present=True)

    original_value = value.get("original_binary")
    original_path = _path_descriptor(original_value, f"{role} original Cargo binary")
    require(_path_inside(original_path, target) and original_path.resolve(strict=False) != target.resolve(strict=False), f"{role} original binary escapes target")
    require(original_path.name == BUILD_ROLES[role], f"{role} original binary name differs")
    if target_present:
        original = _binary_descriptor(original_value, f"{role} original Cargo binary", original_path, present=True)
    else:
        original = _binary_descriptor(original_value, f"{role} original Cargo binary", original_path, present=False)
    return {
        "role": role,
        "receipt": _descriptor(receipt_path, f"{role} build receipt", root=root),
        "binary": binary,
        "original_binary": original,
    }


def _validate_retained_layout(temp: Path, builds: Mapping[str, Mapping[str, Any]]) -> None:
    final_root = temp / FINAL_ATTEMPT
    _directory(final_root, "final2 retained root")
    require({child.name for child in final_root.iterdir()} == set(BUILD_ROLES), "final2 retained root has unexpected children")
    for role in BUILD_ROLES:
        role_root = final_root / role
        _directory(role_root, f"final2 {role} directory")
        expected_name = BUILD_ROLES[role]
        require({child.name for child in role_root.iterdir()} == {expected_name}, f"final2 {role} directory has unexpected children")
        expected = Path(str(builds[role]["binary"]["path"]))
        require(expected.resolve(strict=False) == (role_root / expected_name).resolve(strict=False), f"{role} retained path differs")


def _legacy_final1_custody(temp: Path, builds: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    legacy_root = temp / "final1"
    if not legacy_root.exists():
        return {"status": "absent", "source_attempt": "final1", "entries": []}
    _tree_stats(legacy_root, "legacy final1", owner_uid=os.getuid())
    entries: list[dict[str, Any]] = []
    for role in BUILD_ROLES:
        source = legacy_root / role / BUILD_ROLES[role]
        _regular(source, f"legacy final1 {role}", owner_uid=os.getuid())
        actual = _meta(source, f"legacy final1 {role}", owner_uid=os.getuid())
        expected = builds[role]["binary"]
        require(actual == {"bytes": expected["bytes"], "sha256": expected["sha256"]}, f"legacy final1 {role} differs from authenticated final2")
        entries.append({"role": role, "path": str(source.resolve()), **actual, "same_sha256": True})
    return {"status": "authenticated", "source_attempt": "final1", "entries": entries}


def _candidate_paths(temp: Path, target: Path) -> tuple[list[Path], Path]:
    _directory(temp, "temporary root")
    _directory(target, "Cargo build target")
    _directory(temp / FINAL_ATTEMPT, "final2 retained root")
    candidates: list[Path] = []
    for child in sorted(temp.iterdir(), key=lambda item: item.name):
        if child.name == FINAL_ATTEMPT:
            continue
        _tree_stats(child, f"temporary candidate {child.name}", owner_uid=os.getuid())
        candidates.append(child)
    _tree_stats(target, "Cargo build target", owner_uid=os.getuid())
    return candidates, target


def _proc_path(raw: str) -> Path | None:
    if not raw or not raw.startswith("/"):
        return None
    if raw.endswith(DELETED_SUFFIX):
        raw = raw[: -len(DELETED_SUFFIX)]
    try:
        return Path(raw).resolve(strict=False)
    except OSError:
        return None


def _read_proc_link(path: Path) -> Path | None:
    try:
        raw = os.readlink(path)
    except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
        return None
    return _proc_path(os.fsdecode(raw))


def _read_proc_bytes(path: Path) -> bytes:
    try:
        return path.read_bytes()
    except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
        return b""


def _comm(process: Path) -> str:
    data = _read_proc_bytes(process / "comm")
    return data.decode(errors="replace").strip()


def _process_audit(
    deletion_roots: Iterable[Path],
    evidence_root: Path,
    temp: Path,
    target: Path,
    *,
    proc_root: Path = Path("/proc"),
    self_pid: int | None = None,
) -> dict[str, Any]:
    """Inspect current-user processes before deleting owned caches."""

    _directory(proc_root, "process information root")
    deletion = tuple(path.resolve(strict=False) for path in deletion_roots)
    self_pid = os.getpid() if self_pid is None else self_pid
    uid_scope = os.getuid()
    busy: list[dict[str, Any]] = []
    gates: list[dict[str, Any]] = []
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
            state = process.stat()
        except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
            vanished += 1
            continue
        except PermissionError as error:
            fail(f"/proc/{pid}: cannot identify process owner: {error}")
        if state.st_uid != uid_scope:
            excluded_other_uid.append({"pid": pid, "uid": state.st_uid})
            continue
        name = _comm(process)
        try:
            cwd = _read_proc_link(process / "cwd")
            exe = _read_proc_link(process / "exe")
            fd_dir = process / "fd"
            try:
                fds = sorted(fd_dir.iterdir(), key=lambda item: item.name)
            except (FileNotFoundError, ProcessLookupError, NotADirectoryError):
                vanished += 1
                continue
            fd_paths = [(fd.name, _read_proc_link(fd)) for fd in fds]
            command = [os.fsdecode(item) for item in _read_proc_bytes(process / "cmdline").split(b"\0") if item]
        except PermissionError as error:
            if name in PROTECTED_DAEMONS:
                excluded_daemons.append({"pid": pid, "comm": name, "reason": "protected session daemon; descendants remain in audit scope"})
                continue
            fail(f"/proc/{pid}: cannot inspect current-user process: {error}")
        scanned += 1

        references: list[tuple[str, Path]] = []
        if cwd is not None:
            references.append(("cwd", cwd))
        if exe is not None:
            references.append(("exe", exe))
        references.extend((f"fd:{fd}", path) for fd, path in fd_paths if path is not None)
        for reference, path in references:
            for candidate in deletion:
                if _path_inside(path, candidate):
                    busy.append({"pid": pid, "reference": reference, "path": str(path), "candidate": str(candidate)})
                    break

        command_text = " ".join(command)
        script_reference = False
        for argument in command:
            if Path(argument).name not in ROOT_PROCESS_SCRIPTS:
                continue
            candidate_script = _proc_path(argument) if argument.startswith("/") else (None if cwd is None else _proc_path(str(cwd / argument)))
            if candidate_script is not None and _path_inside(candidate_script, evidence_root):
                script_reference = True
                break
        workspace_reference = any(token in command_text for token in (str(temp), str(target)))
        executable_reference = exe is not None and any(_path_inside(exe, candidate) for candidate in deletion)
        if script_reference or workspace_reference or executable_reference:
            gates.append({
                "pid": pid,
                "exe": None if exe is None else str(exe),
                "command": command_text,
                "script_reference": script_reference,
                "workspace_reference": workspace_reference,
                "executable_reference": executable_reference,
            })

    return {
        "proc_root": str(proc_root.resolve(strict=False)),
        "self_pid": self_pid,
        "uid_scope": uid_scope,
        "scanned_processes": scanned,
        "vanished_processes": vanished,
        "excluded_other_uid_processes": excluded_other_uid,
        "excluded_session_daemons": excluded_daemons,
        "candidate_references": busy,
        "root_gate_capture_processes": gates,
        "safe": not busy and not gates,
    }


def _process_failure(value: Mapping[str, Any]) -> str:
    if value.get("candidate_references"):
        return f"active process references cleanup candidate: {value['candidate_references']}"
    if value.get("root_gate_capture_processes"):
        return f"live gate or capture process remains: {value['root_gate_capture_processes']}"
    return "process audit failed"


def _assert_disjoint(evidence: Path, temp: Path, target: Path) -> None:
    require(not _path_inside(evidence, temp) and not _path_inside(temp, evidence), "evidence and temporary roots overlap")
    require(not _path_inside(evidence, target) and not _path_inside(target, evidence), "evidence and build roots overlap")
    require(not _path_inside(temp, target) and not _path_inside(target, temp), "temporary and build roots overlap")


def _disk(path: Path) -> dict[str, int]:
    try:
        usage = shutil.disk_usage(path)
    except OSError as error:
        fail(f"cannot inspect free space at {path}: {error}")
    return {"total_bytes": usage.total, "used_bytes": usage.used, "free_bytes": usage.free}


def _same_stats(path: Path, expected: Mapping[str, Any], *, root: Path, label: str) -> None:
    actual = _stats_record(_tree_stats(path, label, owner_uid=os.getuid()), root=root, label=label)
    for key in ("kind", "files", "directories", "logical_bytes", "allocated_bytes", "fingerprint"):
        require(actual[key] == expected[key], f"{label}: changed after planning ({key})")


def _remove_tree(path: Path, label: str) -> None:
    state = _lstat(path, label)
    require(not stat.S_ISLNK(state.st_mode), f"{label}: refusing symlink {path}")
    require(stat.S_ISREG(state.st_mode) or stat.S_ISDIR(state.st_mode), f"{label}: refusing special path {path}")
    require(state.st_uid == os.getuid(), f"{label}: ownership changed for {path}")
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


def _absent(path: Path, label: str) -> None:
    _no_symlink_components(path, label)
    require(not path.exists() and not path.is_symlink(), f"{label}: path remains: {path}")


def _retained(builds: Mapping[str, Mapping[str, Any]]) -> dict[str, dict[str, Any]]:
    return {
        role: {
            "path": str(builds[role]["binary"]["path"]),
            "bytes": builds[role]["binary"]["bytes"],
            "sha256": builds[role]["binary"]["sha256"],
        }
        for role in BUILD_ROLES
    }


def _revalidate_retained(builds: Mapping[str, Mapping[str, Any]]) -> None:
    for role, build in builds.items():
        path = Path(str(build["binary"]["path"]))
        require(_meta(path, f"{role} retained binary") == {"bytes": build["binary"]["bytes"], "sha256": build["binary"]["sha256"]}, f"{role} retained binary changed")
        require(os.access(path, os.X_OK), f"{role} retained binary lost executable bit")


def _totals(records: Iterable[Mapping[str, Any]]) -> dict[str, int]:
    values = list(records)
    return {
        "paths": len(values),
        "files": sum(int(value["files"]) for value in values),
        "directories": sum(int(value["directories"]) for value in values),
        "logical_bytes": sum(int(value["logical_bytes"]) for value in values),
        "allocated_bytes": sum(int(value["allocated_bytes"]) for value in values),
    }


def plan_cleanup(
    *,
    root: Path = ROOT,
    temp: Path = TEMP,
    target: Path = TARGET,
    proc_root: Path = Path("/proc"),
    self_pid: int | None = None,
) -> dict[str, Any]:
    """Perform all read-only preconditions and return an authenticated plan."""

    uid = os.getuid()
    evidence = _canonical_directory(root, "evidence root")
    temporary = _canonical_directory(temp, "temporary root", owner_uid=uid)
    build_target = _canonical_directory(target, "Cargo build target", owner_uid=uid)
    _assert_disjoint(evidence, temporary, build_target)
    cleanup_path = evidence / "cleanup.json"
    require(not cleanup_path.exists() and not cleanup_path.is_symlink(), f"refusing to replace existing receipt: {cleanup_path}")
    terminal = _validate_terminal_receipts(evidence)
    builds = {role: _validate_build(role, evidence, temporary, build_target, target_present=True) for role in BUILD_ROLES}
    _validate_retained_layout(temporary, builds)
    custody = _legacy_final1_custody(temporary, builds)
    candidates, target_candidate = _candidate_paths(temporary, build_target)
    candidate_stats = [_stats_record(_tree_stats(path, f"temporary candidate {path.name}", owner_uid=uid), root=temporary, label="temporary candidate") for path in candidates]
    target_stats = _stats_record(_tree_stats(target_candidate, "Cargo build target", owner_uid=uid), root=build_target.parent, label="Cargo build target")
    removed_stats = [*candidate_stats, target_stats]
    process = _process_audit([*candidates, target_candidate], evidence, temporary, build_target, proc_root=proc_root, self_pid=self_pid)
    require(process["safe"], _process_failure(process))
    return {
        "root": evidence,
        "temp": temporary,
        "target": build_target,
        "terminal": terminal,
        "builds": builds,
        "custody": custody,
        "candidates": candidates,
        "target_candidate": target_candidate,
        "removed_stats": removed_stats,
        "process_before": process,
        "proc_root": proc_root,
        "self_pid": self_pid,
        "disk_before": _disk(temporary),
    }


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
    """Execute a previously authenticated plan and write cleanup.json."""

    root = Path(plan["root"])
    temp = Path(plan["temp"])
    target = Path(plan["target"])
    candidates = [Path(path) for path in plan["candidates"]]
    target_candidate = Path(plan["target_candidate"])
    builds = plan["builds"]
    _revalidate_retained(builds)
    before = _process_audit([*candidates, target_candidate], root, temp, target, proc_root=Path(plan["proc_root"]), self_pid=plan["self_pid"])
    require(before["safe"], _process_failure(before))
    expected = {str(Path(item["path"]).resolve(strict=False)): item for item in plan["removed_stats"]}
    for path in [*candidates, target_candidate]:
        key = str(path.resolve(strict=False))
        require(key in expected, f"missing authenticated inventory for {path}")
        _same_stats(path, expected[key], root=temp if path in candidates else target.parent, label=f"cleanup candidate {path}")
    for path in [*candidates, target_candidate]:
        _remove_tree(path, f"remove {path}")
        _absent(path, f"removed cleanup candidate {path}")
    _revalidate_retained(builds)
    require({child.name for child in temp.iterdir()} == {FINAL_ATTEMPT}, "unexpected temporary paths remain")
    after = _process_audit([*candidates, target_candidate], root, temp, target, proc_root=Path(plan["proc_root"]), self_pid=plan["self_pid"])
    require(after["safe"], _process_failure(after))
    retained = _retained(builds)
    receipt: dict[str, Any] = {
        "schema": SCHEMA,
        "version": VERSION,
        "status": "pass",
        "completed_utc": _now(),
        "target": str(target),
        "temporary_root": str(temp),
        "retained_attempt": FINAL_ATTEMPT,
        "retained": retained,
        "removed": list(plan["removed_stats"]),
        "removed_totals": _totals(plan["removed_stats"]),
        "terminal_validation": plan["terminal"],
        "custody": plan["custody"],
        "disk": {"before": plan["disk_before"], "after": _disk(temp)},
        "process_audit": {"before": before, "after": after},
        "driver": _descriptor(Path(__file__).resolve(), "cleanup driver"),
    }
    _write_exclusive(root / "cleanup.json", receipt)
    return receipt


def _verify_removed(records: Any, *, temp: Path, target: Path) -> None:
    require(isinstance(records, list) and records, "cleanup removed inventory is missing")
    seen: set[str] = set()
    target_key = str(target.resolve(strict=False))
    found_target = False
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
            found_target = True
        else:
            require(_path_inside(path, temp), f"cleanup.removed[{index}]: path escapes temporary root")
            require(not _path_inside(path, temp / FINAL_ATTEMPT), f"cleanup.removed[{index}]: retained path was removed")
        _absent(path, f"cleanup.removed[{index}]")
        for field in ("files", "directories", "logical_bytes", "allocated_bytes"):
            require(type(record.get(field)) is int and record[field] >= 0, f"cleanup.removed[{index}].{field} malformed")
        fingerprint = record.get("fingerprint")
        require(isinstance(fingerprint, str) and len(fingerprint) == 64 and all(character in "0123456789abcdef" for character in fingerprint), f"cleanup.removed[{index}].fingerprint malformed")
    require(found_target, "cleanup removed inventory omits the exact build target")


def verify(
    *,
    root: Path = ROOT,
    temp: Path = TEMP,
    target: Path = TARGET,
    proc_root: Path = Path("/proc"),
) -> dict[str, Any]:
    """Verify a completed cleanup without changing any path."""

    uid = os.getuid()
    evidence = _canonical_directory(root, "evidence root")
    temporary = _canonical_directory(temp, "temporary root", owner_uid=uid)
    build_target = Path(target)
    _no_symlink_components(build_target, "Cargo build target")
    _assert_disjoint(evidence, temporary, build_target.resolve(strict=False))
    _absent(build_target, "Cargo build target")
    cleanup_path = evidence / "cleanup.json"
    receipt = _read_json(cleanup_path, "cleanup receipt")
    require(isinstance(receipt, Mapping), "cleanup receipt: object required")
    require(receipt.get("schema") == SCHEMA and receipt.get("version") == VERSION and receipt.get("status") == "pass", "cleanup receipt: schema/status mismatch")
    builds = {role: _validate_build(role, evidence, temporary, build_target, target_present=False) for role in BUILD_ROLES}
    _validate_retained_layout(temporary, builds)
    require({child.name for child in temporary.iterdir()} == {FINAL_ATTEMPT}, "temporary root retains disposable paths")
    expected_retained = _retained(builds)
    require(receipt.get("target") == str(build_target), "cleanup target binding differs")
    require(receipt.get("retained") == expected_retained, "cleanup retained binary inventory differs")
    terminal = _validate_terminal_receipts(evidence)
    require(receipt.get("terminal_validation") == terminal, "cleanup terminal validation inventory differs")
    _verify_removed(receipt.get("removed"), temp=temporary, target=build_target)
    require(receipt.get("removed_totals") == _totals(receipt["removed"]), "cleanup removed totals differ")
    driver = receipt.get("driver")
    require(isinstance(driver, Mapping), "cleanup driver binding missing")
    require(dict(driver) == _descriptor(Path(__file__).resolve(), "cleanup driver"), "cleanup driver changed")
    removed_paths = [Path(str(item["path"])) for item in receipt["removed"]]
    current = _process_audit([*removed_paths, build_target], evidence, temporary, build_target, proc_root=proc_root)
    require(current["safe"], _process_failure(current))
    process = receipt.get("process_audit")
    require(isinstance(process, Mapping), "cleanup process audit missing")
    for phase in ("before", "after"):
        value = process.get(phase)
        require(isinstance(value, Mapping) and value.get("safe") is True, f"cleanup process audit {phase} was unsafe")
        require(value.get("candidate_references") == [] and value.get("root_gate_capture_processes") == [], f"cleanup process audit {phase} retained a live process")
    return {
        "schema": "docx-read-ahead-cleanup-verification-v1",
        "version": VERSION,
        "status": "pass",
        "cleanup_receipt": _descriptor(cleanup_path, "cleanup receipt", root=evidence),
        "retained": expected_retained,
        "target_removed": True,
        "remaining": [],
        "process_audit": current,
    }


def _print_plan(plan: Mapping[str, Any]) -> None:
    print(
        json.dumps(
            {
                "schema": SCHEMA,
                "status": "ready",
                "dry_run": True,
                "target": str(plan["target"]),
                "retained": _retained(plan["builds"]),
                "removed": plan["removed_stats"],
                "removed_totals": _totals(plan["removed_stats"]),
                "process_audit": plan["process_before"],
            },
            indent=2,
            sort_keys=True,
        )
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--dry-run", action="store_true", help="validate and print a deletion plan")
    modes.add_argument("--verify", action="store_true", help="verify a completed cleanup")
    parser.add_argument("--root", type=Path, default=ROOT, help=argparse.SUPPRESS)
    parser.add_argument("--temp-root", type=Path, default=TEMP, help=argparse.SUPPRESS)
    parser.add_argument("--target-root", type=Path, default=TARGET, help=argparse.SUPPRESS)
    parser.add_argument("--proc-root", type=Path, default=Path("/proc"), help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    try:
        configured = (args.root.resolve(), args.temp_root.resolve(), args.target_root.resolve())
        fixed = (ROOT.resolve(), TEMP.resolve(), TARGET.resolve())
        if not args.dry_run and not args.verify:
            require(configured == fixed, "destructive mode is restricted to the fixed 0492 roots")
        if args.verify:
            print(json.dumps(verify(root=args.root, temp=args.temp_root, target=args.target_root, proc_root=args.proc_root), indent=2, sort_keys=True))
            return 0
        plan = plan_cleanup(root=args.root, temp=args.temp_root, target=args.target_root, proc_root=args.proc_root)
        if args.dry_run:
            _print_plan(plan)
            return 0
        receipt = execute_cleanup(plan)
        print(json.dumps({"schema": SCHEMA, "status": receipt["status"], "target": receipt["target"], "removed_totals": receipt["removed_totals"]}, sort_keys=True))
        return 0
    except (CleanupError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"cleanup.py: error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
