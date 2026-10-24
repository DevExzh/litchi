#!/usr/bin/env python3
"""Fail-closed quiet-host ABBA driver for the OPC managed read-ahead lane.

This driver is deliberately separate from the build/correctness prefix and
from the historical all-in-one shell recipe.  It owns only latency captures:
the source pair, probe and binaries are supplied as immutable inputs.  Every
child is run in its own process group and CPU affinity, with a bounded address
space, output streams, report, target and deadline.  Full ``/proc`` snapshots
are retained before, during and after every child.  A busy host, a surviving
descendant, a changed source, malformed report, or any resource violation
fails closed.

The normal binary's allocation field is intentionally not interpreted.  This
is a latency-only lane; allocator qualification requires a separate binary
and protocol.
"""

from __future__ import annotations

import argparse
import errno
import fcntl
import hashlib
import json
import math
import os
from pathlib import Path
import resource
import signal
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Iterable, Mapping


SCHEMA = "litchi.opc.managed-read-ahead.timing.v1"
REPORT_SCHEMA = "docx_managed_read_ahead_v1"
ARCHIVE_BYTES = 16_793_036
ARCHIVE_SHA256 = "a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4"
TEXT_BYTES = 10_000
TEXT_SHA256 = "ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af"
EXACT_RANGES = [
    (16_793_014, 22),
    (16_791_709, 46),
    (16_791_709, 1_305),
    (0, 30),
    (49, 363),
    (412, 16),
    (412, 16),
    (428, 30),
    (469, 234),
    (703, 16),
    (703, 16),
    (2_907, 30),
    (2_965, 324),
    (3_289, 16),
    (3_289, 16),
    (1_420, 30),
    (1_467, 1_424),
    (2_891, 16),
    (2_891, 16),
]
FORWARD_RANGES = [(16_793_014, 22), (16_791_709, 1_327), (0, 4_096)]
HEAVY_NAMES = frozenset(
    {
        "cargo",
        "cargo-fuzz",
        "rustc",
        "rustdoc",
        "rustfmt",
        "clippy-driver",
        "valgrind",
        "perf",
        "heaptrack",
    }
)
DEFAULT_MAX_OUTPUT_BYTES = 4 * 1024 * 1024
DEFAULT_MAX_REPORT_BYTES = 4 * 1024 * 1024
DEFAULT_MAX_RSS_BYTES = 20_000_000_000
DEFAULT_MAX_TARGET_BYTES = 20_000_000_000
DEFAULT_MAX_HOST_BYTES = 16 * 1024 * 1024
DEFAULT_MAX_EVENT_BYTES = 1 * 1024 * 1024
DEFAULT_MAX_RECEIPT_BYTES = 1 * 1024 * 1024
DEFAULT_MAX_PRERUN_LOAD = 1.0
DEFAULT_TIMEOUT_SECONDS = 120.0
DEFAULT_POLL_SECONDS = 0.1
ABBA_ORDER = ("a1", "b1", "b2", "a2")
CPU_TICK_TOLERANCE = 1


class TimingError(RuntimeError):
    """A source, host, child, output, or report contract failed."""


def guarded_process_name(value: str) -> bool:
    """Identify build-heavy or litchi-owned process names from readable text."""
    lowered = value.lower()
    return Path(value).name.lower() in HEAVY_NAMES or "litchi" in lowered


def restricted_metadata_is_unrelated(comm: str, cmdline: list[str], kernel_thread: bool) -> bool:
    """Whether readable process text proves restricted exe/cwd is irrelevant."""
    if kernel_thread:
        return True
    return bool(cmdline) and not any(guarded_process_name(value) for value in (comm, *cmdline))


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise TimingError(f"expected regular file, not symlink: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def write_json_exclusive(path: Path, value: Mapping[str, Any], maximum: int | None = None) -> int:
    if path.exists() or path.is_symlink():
        raise TimingError(f"refusing to overwrite output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()
    if maximum is not None and len(data) > maximum:
        raise TimingError(f"JSON output exceeds cap: {path} ({len(data)} > {maximum})")
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())
    return len(data)


def reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise TimingError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def read_json(path: Path, max_bytes: int) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise TimingError(f"missing or symlinked JSON file: {path}")
    if path.stat().st_size > max_bytes:
        raise TimingError(f"JSON file exceeds cap: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_duplicate_keys)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise TimingError(f"invalid JSON {path}: {error}") from error
    if not isinstance(value, dict):
        raise TimingError(f"JSON root must be an object: {path}")
    return value


def require(condition: bool, message: str) -> None:
    if not condition:
        raise TimingError(message)


def require_equal(actual: Any, expected: Any, location: str) -> None:
    if actual != expected:
        raise TimingError(f"{location}: expected {expected!r}, got {actual!r}")


def range_dicts(ranges: Iterable[tuple[int, int]]) -> list[dict[str, int]]:
    return [{"offset": offset, "requested": length, "returned": length} for offset, length in ranges]


def directory_size(path: Path) -> int:
    if path.is_symlink():
        raise TimingError(f"target is a symlink: {path}")
    if not path.exists():
        return 0
    if path.is_file():
        return path.stat().st_size
    if not path.is_dir():
        raise TimingError(f"target is not a regular file or directory: {path}")
    total = 0
    pending = [path]
    while pending:
        current = pending.pop()
        try:
            entries = list(os.scandir(current))
        except OSError as error:
            raise TimingError(f"cannot inspect target directory {current}: {error}") from error
        for entry in entries:
            try:
                if entry.is_symlink():
                    raise TimingError(f"target contains a symlink: {entry.path}")
                if entry.is_dir(follow_symlinks=False):
                    pending.append(Path(entry.path))
                elif entry.is_file(follow_symlinks=False):
                    total += entry.stat(follow_symlinks=False).st_size
                else:
                    raise TimingError(f"target contains a non-regular entry: {entry.path}")
            except FileNotFoundError as error:
                raise TimingError(f"target entry disappeared during inspection: {entry.path}") from error
            except OSError as error:
                raise TimingError(f"cannot inspect target entry {entry.path}: {error}") from error
    return total


def target_identity(path: Path, maximum: int) -> dict[str, Any]:
    """Hash target structure, metadata, and file contents for immutability."""
    if path.is_symlink():
        raise TimingError(f"target is a symlink: {path}")
    if not path.exists():
        missing = hashlib.sha256(b"missing\n").hexdigest()
        return {"exists": False, "bytes": 0, "entries": 0, "sha256": missing}
    digest = hashlib.sha256()
    total = 0
    entries = 0

    def record(kind: str, relative: str, metadata: os.stat_result, content: str | None = None) -> None:
        digest.update(
            (
                json.dumps(
                    {
                        "kind": kind,
                        "path": relative,
                        "dev": metadata.st_dev,
                        "ino": metadata.st_ino,
                        "mode": metadata.st_mode,
                        "size": metadata.st_size,
                        "mtime_ns": metadata.st_mtime_ns,
                        "ctime_ns": metadata.st_ctime_ns,
                        "content": content,
                    },
                    sort_keys=True,
                    separators=(",", ":"),
                )
                + "\n"
            ).encode()
        )

    def visit(current: Path, relative: Path) -> None:
        nonlocal total, entries
        try:
            current_stat = current.stat(follow_symlinks=False)
            if not current.is_dir():
                raise TimingError(f"target entry is not a directory: {current}")
            record("directory", relative.as_posix(), current_stat)
            entries += 1
            children = sorted(os.scandir(current), key=lambda entry: entry.name)
        except OSError as error:
            raise TimingError(f"cannot inspect target directory {current}: {error}") from error
        for child in children:
            child_path = Path(child.path)
            child_relative = relative / child.name
            try:
                if child.is_symlink():
                    raise TimingError(f"target contains a symlink: {child_path}")
                if child.is_dir(follow_symlinks=False):
                    visit(child_path, child_relative)
                    continue
                if not child.is_file(follow_symlinks=False):
                    raise TimingError(f"target contains a non-regular entry: {child_path}")
                before_stat = child.stat(follow_symlinks=False)
                if before_stat.st_size > maximum:
                    raise TimingError(f"target exceeds cap: {path} ({before_stat.st_size} > {maximum})")
                child_digest, child_bytes = sha256_file(child_path)
                after_stat = child.stat(follow_symlinks=False)
                before_identity = (
                    before_stat.st_dev,
                    before_stat.st_ino,
                    before_stat.st_mode,
                    before_stat.st_size,
                    before_stat.st_mtime_ns,
                    before_stat.st_ctime_ns,
                )
                after_identity = (
                    after_stat.st_dev,
                    after_stat.st_ino,
                    after_stat.st_mode,
                    after_stat.st_size,
                    after_stat.st_mtime_ns,
                    after_stat.st_ctime_ns,
                )
                if before_identity != after_identity:
                    raise TimingError(f"target file changed while fingerprinting: {child_path}")
                if child_bytes != before_stat.st_size:
                    raise TimingError(f"target file size changed while fingerprinting: {child_path}")
                record("file", child_relative.as_posix(), after_stat, child_digest)
                total += child_bytes
                entries += 1
                if total > maximum:
                    raise TimingError(f"target exceeds cap: {path} ({total} > {maximum})")
            except FileNotFoundError as error:
                raise TimingError(f"target entry disappeared during fingerprinting: {child_path}") from error
            except OSError as error:
                raise TimingError(f"cannot fingerprint target entry {child_path}: {error}") from error

    if path.is_dir():
        visit(path, Path("."))
    else:
        before_stat = path.stat(follow_symlinks=False)
        if before_stat.st_size > maximum:
            raise TimingError(f"target exceeds cap: {path} ({before_stat.st_size} > {maximum})")
        file_digest, file_bytes = sha256_file(path)
        after_stat = path.stat(follow_symlinks=False)
        before_identity = (
            before_stat.st_dev,
            before_stat.st_ino,
            before_stat.st_mode,
            before_stat.st_size,
            before_stat.st_mtime_ns,
            before_stat.st_ctime_ns,
        )
        after_identity = (
            after_stat.st_dev,
            after_stat.st_ino,
            after_stat.st_mode,
            after_stat.st_size,
            after_stat.st_mtime_ns,
            after_stat.st_ctime_ns,
        )
        if before_identity != after_identity:
            raise TimingError(f"target file changed while fingerprinting: {path}")
        record("file", path.name, after_stat, file_digest)
        total = file_bytes
        entries = 1
    return {"exists": True, "bytes": total, "entries": entries, "sha256": digest.hexdigest()}


def _proc_stat_fields(raw: str) -> list[str]:
    right = raw.rsplit(") ", 1)
    if len(right) != 2:
        raise ValueError("malformed /proc stat")
    return right[1].split()


def read_proc_identity(pid: int) -> dict[str, Any] | None:
    """Read only identity fields for cleanup when the full monitor is broken."""
    proc = Path("/proc") / str(pid)
    try:
        fields = _proc_stat_fields((proc / "stat").read_text(encoding="utf-8", errors="replace"))
    except FileNotFoundError:
        return None
    except (PermissionError, OSError, ValueError) as error:
        raise TimingError(f"cannot inspect cleanup identity /proc/{pid}: {error}") from error
    try:
        return {
            "pid": pid,
            "ppid": int(fields[1]),
            "pgrp": int(fields[2]),
            "session": int(fields[3]),
            "starttime": int(fields[19]),
            "state": fields[0],
        }
    except (IndexError, ValueError) as error:
        raise TimingError(f"malformed cleanup identity /proc/{pid}/stat: {error}") from error


def read_proc(pid: int) -> dict[str, Any] | None:
    proc = Path("/proc") / str(pid)
    try:
        comm = (proc / "comm").read_text(encoding="utf-8", errors="replace").strip()
        cmdline = [part.decode("utf-8", "replace") for part in (proc / "cmdline").read_bytes().split(b"\0") if part]
        fields = _proc_stat_fields((proc / "stat").read_text(encoding="utf-8", errors="replace"))
        status = (proc / "status").read_text(encoding="utf-8", errors="replace")
    except FileNotFoundError:
        return None
    except (PermissionError, OSError, ValueError) as error:
        raise TimingError(f"cannot inspect /proc/{pid}: {error}") from error
    kernel_thread = not cmdline and len(fields) > 3 and fields[2] == "0" and fields[3] == "0"
    metadata_restricted = False

    # exe/cwd are restricted for some otherwise unrelated host processes
    # (notably PID 1 in a user namespace).  comm plus readable cmdline proves
    # those processes are both outside the guarded heavy set and not another
    # litchi process; retain explicit restricted markers instead of weakening
    # RSS/stat identity checks.  A user process without readable cmdline is
    # not provably unrelated and remains fail-closed.
    try:
        exe = os.readlink(proc / "exe")
    except FileNotFoundError:
        return None
    except PermissionError as error:
        if restricted_metadata_is_unrelated(comm, cmdline, kernel_thread):
            metadata_restricted = True
            exe = "<restricted>"
        else:
            raise TimingError(f"cannot prove restricted /proc/{pid}/exe is unrelated") from error
    except OSError as error:
        if error.errno == errno.EACCES:
            if not restricted_metadata_is_unrelated(comm, cmdline, kernel_thread):
                raise TimingError(f"cannot prove restricted /proc/{pid}/exe is unrelated") from error
            metadata_restricted = True
            exe = "<restricted>"
        else:
            raise TimingError(f"cannot inspect /proc/{pid}/exe: {error}") from error
    try:
        cwd = os.readlink(proc / "cwd")
    except FileNotFoundError:
        return None
    except PermissionError as error:
        if not restricted_metadata_is_unrelated(comm, cmdline, kernel_thread):
            raise TimingError(f"cannot prove restricted /proc/{pid}/cwd is unrelated") from error
        metadata_restricted = True
        cwd = "<restricted>"
    except OSError as error:
        if error.errno == errno.EACCES:
            if not restricted_metadata_is_unrelated(comm, cmdline, kernel_thread):
                raise TimingError(f"cannot prove restricted /proc/{pid}/cwd is unrelated") from error
            metadata_restricted = True
            cwd = "<restricted>"
        else:
            raise TimingError(f"cannot inspect /proc/{pid}/cwd: {error}") from error
    rss_lines = [line for line in status.splitlines() if line.startswith("VmRSS:")]
    if not rss_lines:
        # Kernel threads have no address space and zombies can disappear
        # between the stat/status reads; neither has a meaningful VmRSS.
        # Missing RSS for any other process is an unexpected accounting error.
        stat_state = fields[0] if fields else ""
        stat_pgrp = fields[2] if len(fields) > 2 else "-1"
        stat_session = fields[3] if len(fields) > 3 else "-1"
        if stat_state == "Z" or kernel_thread:
            rss = 0
        else:
            raise TimingError(f"cannot read RSS for /proc/{pid}: VmRSS is absent")
    else:
        try:
            rss = int(rss_lines[0].split()[1]) * os.sysconf("SC_PAGE_SIZE")
        except (IndexError, ValueError, OSError) as error:
            raise TimingError(f"cannot read RSS for /proc/{pid}: {error}") from error
    try:
        return {
            "pid": pid,
            "ppid": int(fields[1]),
            "pgrp": int(fields[2]),
            "session": int(fields[3]),
            "starttime": int(fields[19]),
            "state": fields[0],
            "utime": int(fields[11]),
            "stime": int(fields[12]),
            "psr": int(fields[36]),
            "rss_bytes": rss,
            "comm": comm,
            "exe": exe,
            "cwd": cwd,
            "cmdline": cmdline,
            "metadata_restricted": metadata_restricted,
        }
    except (IndexError, ValueError) as error:
        raise TimingError(f"malformed /proc/{pid}/stat: {error}") from error


def proc_snapshot() -> dict[int, dict[str, Any]]:
    result: dict[int, dict[str, Any]] = {}
    try:
        entries = list(Path("/proc").iterdir())
    except OSError as error:
        raise TimingError(f"cannot enumerate /proc: {error}") from error
    for item in entries:
        if not item.name.isdigit():
            continue
        entry = read_proc(int(item.name))
        if entry is not None:
            result[entry["pid"]] = entry
    return result


def load_average() -> list[float]:
    try:
        values = [float(value) for value in os.getloadavg()]
    except (AttributeError, OSError, ValueError) as error:
        raise TimingError(f"cannot read load average: {error}") from error
    if not values or any(not math.isfinite(value) or value < 0 for value in values):
        raise TimingError(f"invalid load average: {values!r}")
    return values


def cpu_counters(cpu: int) -> tuple[int, int]:
    """Return (total_ticks, busy_ticks) for one CPU over an interval."""
    try:
        lines = Path("/proc/stat").read_text(encoding="ascii").splitlines()
    except (OSError, UnicodeError) as error:
        raise TimingError(f"cannot read /proc/stat: {error}") from error
    label = f"cpu{cpu}"
    for line in lines:
        fields = line.split()
        if not fields or fields[0] != label:
            continue
        try:
            values = [int(value) for value in fields[1:9]]
        except ValueError as error:
            raise TimingError(f"invalid {label} counters: {error}") from error
        if len(values) < 8 or any(value < 0 for value in values):
            raise TimingError(f"incomplete or negative {label} counters")
        total = sum(values)
        busy = values[0] + values[1] + values[2] + values[5] + values[6] + values[7]
        return total, busy
    raise TimingError(f"/proc/stat has no counters for selected CPU {cpu}")


def descendants(snapshot: Mapping[int, Mapping[str, Any]], root: int) -> set[int]:
    owned = {root}
    changed = True
    while changed:
        changed = False
        for pid, entry in snapshot.items():
            if pid not in owned and int(entry.get("ppid", -1)) in owned:
                owned.add(pid)
                changed = True
    # A detached descendant can be reparented while retaining the process
    # group.  Keep it owned for cleanup and do not mistake it for host work.
    owned.update(pid for pid, entry in snapshot.items() if int(entry.get("pgrp", -1)) == root)
    return owned


def refresh_owned(
    snapshot: Mapping[int, Mapping[str, Any]],
    known: Mapping[int, int],
    root_pid: int,
    root_pgrp: int,
    root_session: int,
) -> dict[int, int]:
    """Keep ownership across reparenting while rejecting PID reuse."""
    owned: dict[int, int] = {}
    for pid, expected_start in known.items():
        entry = snapshot.get(pid)
        if entry is None:
            continue
        actual_start = int(entry["starttime"])
        if actual_start != expected_start:
            raise TimingError(f"owned child PID {pid} changed identity")
        owned[pid] = expected_start
    changed = True
    while changed:
        changed = False
        for pid, entry in snapshot.items():
            if pid in owned:
                continue
            if (
                int(entry.get("ppid", -1)) in owned
                or int(entry.get("pgrp", -1)) == root_pgrp
                or int(entry.get("session", -1)) == root_session
            ):
                owned[pid] = int(entry["starttime"])
                changed = True
    if root_pid in snapshot and root_pid not in owned:
        owned[root_pid] = int(snapshot[root_pid]["starttime"])
    return owned


def discover_cleanup_owned(
    known: Mapping[int, int],
    runner_pid: int,
    preexisting_pids: set[int],
    root_pgrp: int | None = None,
    root_session: int | None = None,
) -> dict[int, int]:
    """Discover newly adopted descendants without the full monitor snapshot."""
    try:
        proc_entries = list(Path("/proc").iterdir())
    except OSError as error:
        raise TimingError(f"cannot enumerate /proc for cleanup: {error}") from error
    identities: dict[int, dict[str, Any]] = {}
    for item in proc_entries:
        if not item.name.isdigit():
            continue
        pid = int(item.name)
        if pid in preexisting_pids and pid not in known:
            continue
        entry = read_proc_identity(pid)
        if entry is not None:
            identities[pid] = entry
    owned = dict(known)
    changed = True
    while changed:
        changed = False
        for pid, entry in identities.items():
            if pid in owned or pid in preexisting_pids:
                continue
            if (
                int(entry["ppid"]) in owned
                or int(entry["ppid"]) == runner_pid
                or (root_pgrp is not None and int(entry["pgrp"]) == root_pgrp)
                or (root_session is not None and int(entry["session"]) == root_session)
            ):
                owned[pid] = int(entry["starttime"])
                changed = True
    return owned


def independent_live_owned(owned: Mapping[int, int], *, root_pid: int | None = None) -> dict[int, int]:
    """Check owned identities using only /proc stat, not the full monitor."""
    result: dict[int, int] = {}
    for pid, expected_start in owned.items():
        entry = read_proc_identity(pid)
        if entry is None:
            continue
        actual_start = int(entry["starttime"])
        if actual_start != expected_start:
            raise TimingError(f"refusing to treat reused owned PID {pid} as a child")
        if entry.get("state") == "Z" and pid != root_pid:
            try:
                waited, _ = os.waitpid(pid, os.WNOHANG)
            except ChildProcessError:
                result[pid] = expected_start
                continue
            except ProcessLookupError:
                continue
            except OSError as error:
                raise TimingError(f"cannot reap owned zombie PID {pid}: {error}") from error
            if waited == pid:
                continue
        result[pid] = expected_start
    return result


def is_heavy(entry: Mapping[str, Any]) -> bool:
    names = [str(entry.get("exe", "")), str(entry.get("comm", ""))]
    if any(guarded_process_name(value) for value in names):
        return True
    command = entry.get("cmdline", [])
    return bool(command) and any(guarded_process_name(str(value)) for value in command)


def snapshot_json(snapshot: Mapping[int, Mapping[str, Any]], load: list[float]) -> dict[str, Any]:
    return {
        "utc": utc_now(),
        "loadavg": load,
        "processes": [snapshot[pid] for pid in sorted(snapshot)],
    }


def quiet_violations(
    previous: Mapping[int, Mapping[str, Any]],
    current: Mapping[int, Mapping[str, Any]],
    *,
    owned: Mapping[int, int],
    load_before: float,
    max_load_delta: float,
    previous_cpu: tuple[int, int],
    current_cpu: tuple[int, int],
    check_load_delta: bool,
    absolute_load_limit: float | None,
) -> list[str]:
    # Account for the complete selected-CPU interval from /proc/stat and
    # subtract ticks charged to the monitor and the owned process identities.
    # This deliberately does not use instantaneous PSR as a migration proxy.
    owned_ids = set(owned)
    owned_ids.add(os.getpid())
    violations: list[str] = []
    heavy = [
        entry
        for pid, entry in current.items()
        if pid not in owned_ids and is_heavy(entry)
    ]
    if heavy:
        violations.append("unrelated heavy process: " + repr(heavy))
    previous_total, previous_busy = previous_cpu
    current_total, current_busy = current_cpu
    if current_total < previous_total or current_busy < previous_busy:
        raise TimingError("selected CPU counters moved backwards")
    selected_busy_delta = current_busy - previous_busy
    owned_ticks = 0
    for pid in owned_ids:
        old = previous.get(pid)
        new = current.get(pid)
        old_ticks = 0 if old is None else int(old.get("utime", 0)) + int(old.get("stime", 0))
        # If an owned child exited between samples, its final CPU interval is
        # bounded by the last observed value.  Do not charge that disappearance
        # as unowned host work; newly appearing unowned PIDs remain in the full
        # current snapshot and are still accounted against /proc/stat.
        new_ticks = old_ticks if new is None and old is not None else (
            0 if new is None else int(new.get("utime", 0)) + int(new.get("stime", 0))
        )
        if new_ticks >= old_ticks:
            owned_ticks += new_ticks - old_ticks
    outside_busy = selected_busy_delta - owned_ticks
    if outside_busy > CPU_TICK_TOLERANCE:
        violations.append(
            "selected CPU interval had unowned busy ticks: "
            f"busy_delta={selected_busy_delta}, owned_ticks={owned_ticks}"
        )
    load = load_average()
    if absolute_load_limit is not None and load[0] > absolute_load_limit:
        violations.append(f"load1 exceeded absolute quiet bound: {load[0]} > {absolute_load_limit}")
    if check_load_delta and load[0] > load_before + max_load_delta:
        violations.append(f"load1 increased beyond quiet bound: before={load_before}, now={load[0]}")
    return violations


def process_rss(snapshot: Mapping[int, Mapping[str, Any]], owned: Mapping[int, int]) -> int:
    total = 0
    for pid, expected_start in owned.items():
        entry = snapshot.get(pid)
        if entry is None:
            continue
        if int(entry["starttime"]) != expected_start:
            raise TimingError(f"owned child PID {pid} changed identity while reading RSS")
        total += int(entry["rss_bytes"])
    return total


def process_group_exists(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        raise TimingError(f"cannot inspect process group {pgid}: permission denied")
    except OSError as error:
        raise TimingError(f"cannot inspect process group {pgid}: {error}") from error
    return True


def set_child_subreaper(enabled: bool) -> int:
    """Adopt detached descendants so cleanup remains attributable to this run."""
    try:
        import ctypes

        prctl = ctypes.CDLL(None, use_errno=True).prctl
        prctl.argtypes = [ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong]
        prctl.restype = ctypes.c_int
        previous = ctypes.c_int()
        if prctl(37, ctypes.addressof(previous), 0, 0, 0) != 0:  # PR_GET_CHILD_SUBREAPER
            raise OSError(ctypes.get_errno(), "PR_GET_CHILD_SUBREAPER failed")
        if prctl(36, 1 if enabled else 0, 0, 0, 0) != 0:  # PR_SET_CHILD_SUBREAPER
            raise OSError(ctypes.get_errno(), "PR_SET_CHILD_SUBREAPER failed")
        return previous.value
    except (AttributeError, OSError) as error:
        raise TimingError(f"cannot configure child subreaper: {error}") from error


def live_owned(owned: Mapping[int, int], *, root_pid: int | None = None) -> dict[int, int]:
    snapshot = proc_snapshot()
    result: dict[int, int] = {}
    for pid, expected_start in owned.items():
        entry = snapshot.get(pid)
        if entry is None:
            continue
        actual_start = int(entry["starttime"])
        if actual_start != expected_start:
            raise TimingError(f"refusing to treat reused owned PID {pid} as a child")
        if entry.get("state") == "Z" and pid != root_pid:
            # A detached descendant can be adopted by this runner's
            # subreaper as a zombie.  Reap it explicitly; treating a zombie
            # as a running descendant would make cleanup fail forever.
            try:
                waited, _ = os.waitpid(pid, os.WNOHANG)
            except ChildProcessError:
                # It is still present but not waitable by this process: fail
                # closed rather than claiming ownership was cleaned up.
                result[pid] = expected_start
                continue
            except ProcessLookupError:
                continue
            except OSError as error:
                raise TimingError(f"cannot reap owned zombie PID {pid}: {error}") from error
            if waited == pid:
                continue
        result[pid] = expected_start
    return result


def terminate_group(
    process: subprocess.Popen[bytes],
    owned: Mapping[int, int],
    pgid: int,
    *,
    runner_pid: int | None = None,
    preexisting_pids: set[int] | None = None,
    grace_seconds: float = 5.0,
) -> bool:
    cleanup_owned = dict(owned)

    def refresh_cleanup_owned() -> None:
        if runner_pid is None or preexisting_pids is None:
            return
        cleanup_owned.update(
            discover_cleanup_owned(
                cleanup_owned,
                runner_pid,
                preexisting_pids,
                root_pgrp=pgid,
                root_session=pgid,
            )
        )

    def signal_owned(kind: signal.Signals) -> None:
        refresh_cleanup_owned()
        for pid, expected_start in cleanup_owned.items():
            entry = read_proc_identity(pid)
            if entry is None:
                continue
            if int(entry["starttime"]) != expected_start:
                raise TimingError(f"refusing to signal reused owned PID {pid}")
            try:
                os.kill(pid, kind)
            except ProcessLookupError:
                continue
            except OSError as error:
                raise TimingError(f"cannot signal owned PID {pid}: {error}") from error

    forced = False
    try:
        os.killpg(pgid, signal.SIGTERM)
    except ProcessLookupError:
        pass
    except OSError as error:
        raise TimingError(f"cannot terminate process group {pgid}: {error}") from error
    signal_owned(signal.SIGTERM)
    deadline = time.monotonic() + grace_seconds
    while time.monotonic() < deadline:
        refresh_cleanup_owned()
        if process.poll() is not None and not process_group_exists(pgid) and not independent_live_owned(cleanup_owned, root_pid=process.pid):
            return forced
        time.sleep(0.05)
    forced = True
    refresh_cleanup_owned()
    try:
        os.killpg(pgid, signal.SIGKILL)
    except ProcessLookupError:
        pass
    except OSError as error:
        raise TimingError(f"cannot kill process group {pgid}: {error}") from error
    signal_owned(signal.SIGKILL)
    try:
        process.wait(timeout=grace_seconds)
    except subprocess.TimeoutExpired as error:
        raise TimingError(f"child did not terminate after process-group cleanup: {process.pid}") from error
    deadline = time.monotonic() + grace_seconds
    while time.monotonic() < deadline:
        refresh_cleanup_owned()
        if not process_group_exists(pgid) and not independent_live_owned(cleanup_owned, root_pid=process.pid):
            return forced
        time.sleep(0.05)
    refresh_cleanup_owned()
    if process_group_exists(pgid) or independent_live_owned(cleanup_owned, root_pid=process.pid):
        raise TimingError(f"descendant process group survived cleanup: {process.pid}")
    return forced


def ensure_child_reaped(
    process: subprocess.Popen[bytes],
    owned: Mapping[int, int],
    pgid: int | None,
    forced_cleanup: bool,
    *,
    runner_pid: int | None = None,
    preexisting_pids: set[int] | None = None,
) -> bool:
    """Always reap the direct child, even when monitoring raised an error."""
    cleanup_owned = dict(owned)
    if runner_pid is not None and preexisting_pids is not None:
        cleanup_owned.update(discover_cleanup_owned(cleanup_owned, runner_pid, preexisting_pids, pgid, pgid))
        if isinstance(owned, dict):
            owned.update(cleanup_owned)
    cleanup_error: BaseException | None = None
    try:
        group_alive = pgid is not None and process_group_exists(pgid)
        owned_alive = bool(independent_live_owned(cleanup_owned, root_pid=process.pid)) if pgid is not None else False
        if process.poll() is None or group_alive or owned_alive:
            if pgid is not None:
                forced_cleanup = terminate_group(
                    process,
                    cleanup_owned,
                    pgid,
                    runner_pid=runner_pid,
                    preexisting_pids=preexisting_pids,
                ) or forced_cleanup
            else:
                process.terminate()
    except BaseException as error:
        cleanup_error = error
        # Cleanup itself must not prevent the direct wait/reap.  Try the
        # process group and direct child independently, then surface the
        # original cleanup failure after wait below.
        if pgid is not None:
            try:
                os.killpg(pgid, signal.SIGKILL)
            except ProcessLookupError:
                pass
            except OSError:
                pass
        try:
            process.kill()
        except ProcessLookupError:
            pass
        except OSError:
            pass
        forced_cleanup = True
    finally:
        try:
            process.wait(timeout=5.0)
        except subprocess.TimeoutExpired as error:
            if pgid is not None:
                try:
                    os.killpg(pgid, signal.SIGKILL)
                except ProcessLookupError:
                    pass
                except OSError:
                    pass
            try:
                process.kill()
            except ProcessLookupError:
                pass
            except OSError:
                pass
            try:
                process.wait(timeout=5.0)
            except BaseException as reap_error:
                if cleanup_error is None:
                    cleanup_error = reap_error
        except BaseException as reap_error:
            if cleanup_error is None:
                cleanup_error = reap_error
    if cleanup_error is not None:
        raise cleanup_error
    return forced_cleanup


class BoundedCapture(threading.Thread):
    def __init__(self, pipe: Any, path: Path, maximum: int):
        super().__init__(daemon=True)
        self.pipe = pipe
        self.path = path
        self.maximum = maximum
        self.bytes = 0
        self.overflow = False
        self.error: str | None = None

    def run(self) -> None:
        try:
            with self.path.open("xb") as output:
                while True:
                    chunk = self.pipe.read(64 * 1024)
                    if not chunk:
                        break
                    remaining = self.maximum - self.bytes
                    if remaining > 0:
                        output.write(chunk[:remaining])
                        self.bytes += min(len(chunk), remaining)
                    if len(chunk) > remaining:
                        self.overflow = True
                output.flush()
                os.fsync(output.fileno())
        except BaseException as error:  # retain the failure for the parent loop
            self.error = f"{type(error).__name__}: {error}"


def host_snapshot_descriptor(
    path: Path,
    snapshot: Mapping[int, Mapping[str, Any]],
    load: list[float],
    maximum: int,
) -> dict[str, Any]:
    bytes_written = write_json_exclusive(path, snapshot_json(snapshot, load), maximum)
    digest, _ = sha256_file(path)
    return {"path": str(path), "bytes": bytes_written, "sha256": digest}


def validate_report(path: Path, endpoint: str, revision: str, policy: str, warmup: int, samples: int, max_bytes: int) -> dict[str, Any]:
    report = read_json(path, max_bytes)
    require_equal(report.get("schema"), REPORT_SCHEMA, "report.schema")
    require_equal(report.get("version"), 1, "report.version")
    require_equal(report.get("requested_source_revision"), revision, "report.requested_source_revision")
    require_equal(report.get("source_bytes"), ARCHIVE_BYTES, "report.source_bytes")
    require_equal(report.get("source_sha256"), ARCHIVE_SHA256, "report.source_sha256")
    require_equal(report.get("expected_text_bytes"), TEXT_BYTES, "report.expected_text_bytes")
    require_equal(report.get("expected_text_sha256"), TEXT_SHA256, "report.expected_text_sha256")
    require_equal(report.get("expected_archive_members"), 20, "report.expected_archive_members")
    require_equal(report.get("warmup"), warmup, "report.warmup")
    require_equal(report.get("samples"), samples, "report.samples")
    provider = report.get("provider")
    require(isinstance(provider, dict), "report.provider must be an object")
    require_equal(provider.get("name"), "PptxRangeSource", "report.provider.name")
    require_equal(provider.get("max_range_bytes"), 65_536, "report.provider.max_range_bytes")
    require_equal(provider.get("delay_us"), 0, "report.provider.delay_us")
    require_equal(provider.get("transfer_bytes_per_second"), None, "report.provider.transfer_bytes_per_second")
    corpus = report.get("corpus")
    require(isinstance(corpus, dict), "report.corpus must be an object")
    require_equal(corpus.get("archive_bytes"), ARCHIVE_BYTES, "report.corpus.archive_bytes")
    require_equal(corpus.get("archive_sha256"), ARCHIVE_SHA256, "report.corpus.archive_sha256")
    require_equal(corpus.get("target_payload_bytes"), TEXT_BYTES, "report.corpus.target_payload_bytes")
    require_equal(corpus.get("target_payload_sha256"), TEXT_SHA256, "report.corpus.target_payload_sha256")
    media_ranges = report.get("media_ranges")
    require(isinstance(media_ranges, list) and len(media_ranges) == 8, "report.media_ranges must contain eight members")
    expected = FORWARD_RANGES if policy == "forward" else EXACT_RANGES
    policy_record = report.get("policy")
    require(isinstance(policy_record, dict), "report.policy must be an object")
    require_equal(policy_record.get("name"), "forward_start" if policy == "forward" else "exact", "report.policy.name")
    require_equal(policy_record.get("enabled"), policy == "forward", "report.policy.enabled")
    require_equal(policy_record.get("configured_window_bytes"), 4_096 if policy == "forward" else None, "report.policy.configured_window_bytes")
    rows = report.get("rows")
    require(isinstance(rows, list) and len(rows) == samples, "report.rows length does not match samples")
    for row_index, row in enumerate(rows):
        require(isinstance(row, dict), f"report.rows[{row_index}] must be an object")
        require_equal(row.get("sample_index"), warmup + row_index, f"rows[{row_index}].sample_index")
        require(isinstance(row.get("elapsed_ns"), int) and row["elapsed_ns"] > 0, f"rows[{row_index}].elapsed_ns")
        require_equal(row.get("actual_text_verified"), True, f"rows[{row_index}].actual_text_verified")
        require_equal(row.get("actual_text_bytes"), TEXT_BYTES, f"rows[{row_index}].actual_text_bytes")
        require_equal(row.get("actual_text_sha256"), TEXT_SHA256, f"rows[{row_index}].actual_text_sha256")
        require_equal(row.get("source_version_unchanged"), True, f"rows[{row_index}].source_version_unchanged")
        require_equal(row.get("policy"), policy_record, f"rows[{row_index}].policy")
        cache = row.get("cache")
        require(isinstance(cache, dict), f"rows[{row_index}].cache")
        require_equal(cache.get("open_successful_loads"), 0, f"rows[{row_index}].cache.open_successful_loads")
        require_equal(cache.get("successful_loads"), 1, f"rows[{row_index}].cache.successful_loads")
        require_equal(cache.get("failed_loads"), 0, f"rows[{row_index}].cache.failed_loads")
        require_equal(cache.get("budget_managed"), True, f"rows[{row_index}].cache.budget_managed")
        physical = row.get("physical")
        require(isinstance(physical, dict), f"rows[{row_index}].physical")
        require_equal(physical.get("calls"), len(expected), f"rows[{row_index}].physical.calls")
        expected_bytes = sum(length for _, length in expected)
        require_equal(physical.get("requested_bytes"), expected_bytes, f"rows[{row_index}].physical.requested_bytes")
        require_equal(physical.get("returned_bytes"), expected_bytes, f"rows[{row_index}].physical.returned_bytes")
        require_equal(physical.get("short_reads"), 0, f"rows[{row_index}].physical.short_reads")
        require_equal(physical.get("ranges"), range_dicts(expected), f"rows[{row_index}].physical.ranges")
        source_read = row.get("source_read")
        if policy == "forward":
            require(isinstance(source_read, dict), f"rows[{row_index}].source_read")
            for key, value in {"enabled": True, "configured_window_bytes": 4_096, "requests": 19, "hits": 16, "misses": 3, "fills": 3, "requested_bytes": 3_966, "returned_bytes": 5_445}.items():
                require_equal(source_read.get(key), value, f"rows[{row_index}].source_read.{key}")
        else:
            require_equal(source_read, None, f"rows[{row_index}].source_read")
        budget = row.get("budget")
        require(isinstance(budget, dict), f"rows[{row_index}].budget")
        require_equal(budget.get("memory_after_drop"), 0, f"rows[{row_index}].budget.memory_after_drop")
        require_equal(budget.get("objects_after_drop"), 0, f"rows[{row_index}].budget.objects_after_drop")
        require_equal(budget.get("managed"), True, f"rows[{row_index}].budget.managed")
        require_equal(budget.get("input_bytes_match_physical_returned"), True, f"rows[{row_index}].budget.input_bytes_match_physical_returned")
        allocation = row.get("allocation")
        require(isinstance(allocation, dict), f"rows[{row_index}].allocation")
        require(allocation.get("sample") in (None, {}), f"rows[{row_index}].allocation.sample must be unavailable in latency-only mode")
    return {
        "endpoint": endpoint,
        "revision": revision,
        "policy": policy,
        "samples": samples,
        "warmup": warmup,
        "report_sha256": sha256_file(path)[0],
        "report_bytes": path.stat().st_size,
    }


def safe_target_check(path: Path, maximum: int) -> int:
    size = directory_size(path)
    if size > maximum:
        raise TimingError(f"target exceeds cap: {path} ({size} > {maximum})")
    return size


def verify_source(root: Path, revision: str, tree: str) -> dict[str, str]:
    """Verify that a timing endpoint is the exact clean source checkout."""
    if not root.is_dir() or root.is_symlink():
        raise TimingError(f"source checkout is missing or symlinked: {root}")

    def git(*arguments: str) -> str:
        completed = subprocess.run(
            ["git", "-C", str(root), *arguments],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )
        if completed.returncode != 0:
            detail = completed.stderr.strip() or completed.stdout.strip()
            raise TimingError(f"git {' '.join(arguments)} failed for {root}: {detail}")
        return completed.stdout.strip()

    actual_revision = git("rev-parse", "HEAD")
    actual_tree = git("rev-parse", "HEAD^{tree}")
    status = git("status", "--porcelain=v1", "--untracked-files=all")
    require_equal(actual_revision, revision, f"{root} HEAD")
    require_equal(actual_tree, tree, f"{root} HEAD^{{tree}}")
    require_equal(status, "", f"{root} status")
    return {"path": str(root), "revision": actual_revision, "tree": actual_tree}


def verify_binary(path: Path, expected: Mapping[str, Any], location: str) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file() or not os.access(path, os.X_OK):
        raise TimingError(f"{location} is missing, symlinked, or non-executable: {path}")
    digest, size = sha256_file(path)
    require_equal(digest, expected.get("sha256"), f"{location}.sha256")
    require_equal(size, expected.get("bytes"), f"{location}.bytes")
    return {"path": str(path), "sha256": digest, "bytes": size}


def acquire_cpu_lock(path: Path) -> Any:
    """Hold the protocol's external CPU lock for the complete ABBA sequence."""
    if path.is_symlink():
        raise TimingError(f"CPU lock is symlinked: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    handle = path.open("a+b")
    try:
        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except (BlockingIOError, OSError) as error:
        handle.close()
        raise TimingError(f"CPU lock is busy: {path}: {error}") from error
    return handle


def run_child(
    *,
    output: Path,
    endpoint: str,
    binary: Path,
    binary_identity: Mapping[str, Any],
    revision: str,
    source_tree: str,
    policy: str,
    label: str,
    source_cwd: Path,
    target: Path,
    tmpdir: Path,
    cpu: int,
    samples: int,
    warmup: int,
    timeout_seconds: float,
    poll_seconds: float,
    max_rss: int,
    max_target: int,
    max_output: int,
    max_report: int,
    max_host: int,
    max_event: int,
    max_load_delta: float,
    max_prerun_load: float,
) -> dict[str, Any]:
    report_path = output / "reports" / policy / f"{label}.json"
    stdout_path = output / "logs" / f"{policy}-{label}.stdout"
    stderr_path = output / "logs" / f"{policy}-{label}.stderr"
    event_path = output / "events" / f"{policy}-{label}.json"
    host_before_path = output / "host" / f"{policy}-{label}-before.json"
    host_probe_path = output / "host" / f"{policy}-{label}-probe.json"
    host_after_path = output / "host" / f"{policy}-{label}-after.json"
    during_path = output / "host" / f"{policy}-{label}-during.jsonl"
    for path in (
        report_path,
        stdout_path,
        stderr_path,
        event_path,
        host_before_path,
        host_probe_path,
        host_after_path,
        during_path,
    ):
        if path.exists() or path.is_symlink():
            raise TimingError(f"output path is not fresh: {path}")
    source_before = verify_source(source_cwd, revision, source_tree)
    binary_before = verify_binary(binary, binary_identity, f"{endpoint} binary before {policy}/{label}")
    if cpu not in os.sched_getaffinity(0):
        raise TimingError(f"selected CPU is not allowed: {cpu}")
    if tmpdir.is_symlink():
        raise TimingError(f"temporary directory is symlinked: {tmpdir}")
    tmpdir.mkdir(parents=True, exist_ok=True)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    stdout_path.parent.mkdir(parents=True, exist_ok=True)
    event_path.parent.mkdir(parents=True, exist_ok=True)
    host_before_path.parent.mkdir(parents=True, exist_ok=True)
    before = proc_snapshot()
    preexisting_pids = set(before)
    before_load = load_average()
    before_cpu = cpu_counters(cpu)
    before_host = host_snapshot_descriptor(host_before_path, before, before_load, max_host)
    time.sleep(poll_seconds)
    quiet_probe = proc_snapshot()
    probe_load = load_average()
    probe_cpu = cpu_counters(cpu)
    probe_host = host_snapshot_descriptor(host_probe_path, quiet_probe, probe_load, max_host)
    if before_load[0] > max_prerun_load or probe_load[0] > max_prerun_load:
        raise TimingError(
            f"quiet-host preflight rejected stable load for {endpoint}/{policy}/{label}: "
            f"before={before_load[0]}, probe={probe_load[0]}, limit={max_prerun_load}"
        )
    pre_violations = quiet_violations(
        before,
        quiet_probe,
        owned={},
        load_before=before_load[0],
        max_load_delta=max_load_delta,
        previous_cpu=before_cpu,
        current_cpu=probe_cpu,
        check_load_delta=True,
        absolute_load_limit=max_prerun_load,
    )
    if pre_violations:
        raise TimingError(f"quiet-host preflight rejected {endpoint}/{policy}/{label}: {pre_violations}")
    target_before_identity = target_identity(target, max_target)
    target_before = int(target_before_identity["bytes"])
    env = os.environ.copy()
    env.pop("LITCHI_CAPPED", None)
    env.pop("LITCHI_BLESS", None)
    env.pop("RUSTFLAGS", None)
    env["TMPDIR"] = str(tmpdir)
    command = [
        str(binary),
        "docx-managed-read-ahead",
        "--policy",
        "forward-start" if policy == "forward" else "exact",
    ]
    if policy == "forward":
        command.extend(["--window-bytes", "4096"])
    command.extend(
        [
            "--max-range",
            "65536",
            "--delay-us",
            "0",
            "--samples",
            str(samples),
            "--warmup",
            str(warmup),
            "--source-revision",
            revision,
            "--output",
            str(report_path),
        ]
    )

    parent_pid = os.getpid()

    def preexec() -> None:
        try:
            if os.getppid() != parent_pid:
                os._exit(127)
            os.setsid()
            import ctypes

            prctl = ctypes.CDLL(None, use_errno=True).prctl
            prctl.argtypes = [ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong]
            prctl.restype = ctypes.c_int
            if prctl(1, signal.SIGKILL, 0, 0, 0) != 0:  # PR_SET_PDEATHSIG
                os._exit(127)
            os.sched_setaffinity(0, {cpu})
            resource.setrlimit(resource.RLIMIT_AS, (max_rss, max_rss))
            resource.setrlimit(resource.RLIMIT_CORE, (0, 0))
        except BaseException:
            os._exit(127)

    process: subprocess.Popen[bytes] | None = None
    pgid: int | None = None
    owned: dict[int, int] = {}
    owned_seen: dict[int, int] = {}
    stdout_capture: BoundedCapture | None = None
    stderr_capture: BoundedCapture | None = None
    violations: list[str] = []
    timeout = False
    forced_cleanup = False
    peak_rss = 0
    started_utc: str | None = None
    target_after = target_before
    target_after_identity = target_before_identity
    during_bytes = 0
    previous = quiet_probe
    previous_cpu = probe_cpu
    with during_path.open("x", encoding="utf-8") as during:
        try:
            started_utc = utc_now()
            binary_before = verify_binary(binary, binary_identity, f"{endpoint} binary before spawn {policy}/{label}")
            process = subprocess.Popen(
                command,
                cwd=str(source_cwd),
                env=env,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                start_new_session=False,
                preexec_fn=preexec,
            )
            require(process.stdout is not None and process.stderr is not None, "child pipes were not created")
            try:
                pgid = os.getpgid(process.pid)
            except OSError as error:
                pgid = process.pid
                raise TimingError(f"cannot inspect child process group: {error}") from error
            require(pgid == process.pid, f"child did not create an independent process group: {pgid}")
            root_entry = read_proc(process.pid)
            require(root_entry is not None, "child disappeared before process-group identity was checked")
            require_equal(root_entry.get("pgrp"), process.pid, "child process group")
            require_equal(root_entry.get("session"), process.pid, "child session")
            owned = {process.pid: int(root_entry["starttime"])}
            owned_seen.update(owned)
            stdout_capture = BoundedCapture(process.stdout, stdout_path, max_output)
            stderr_capture = BoundedCapture(process.stderr, stderr_path, max_output)
            stdout_capture.start()
            stderr_capture.start()
            started = time.monotonic()
            while process.poll() is None:
                time.sleep(poll_seconds)
                elapsed = time.monotonic() - started
                if elapsed > timeout_seconds:
                    timeout = True
                    violations.append(f"child exceeded timeout: {timeout_seconds}s")
                    break
                current = proc_snapshot()
                current_load = load_average()
                current_cpu = cpu_counters(cpu)
                require(pgid is not None, "child process group identity was not retained")
                owned = refresh_owned(current, owned, process.pid, pgid, process.pid)
                # A detached descendant adopted by this runner's subreaper
                # has neither the child's pgrp nor session.  It is still
                # attributable when it is a newly adopted direct child.
                for pid, entry in current.items():
                    if pid not in preexisting_pids and int(entry.get("ppid", -1)) == parent_pid:
                        owned[pid] = int(entry["starttime"])
                owned_seen.update(owned)
                accounting_current = dict(current)
                for pid in owned:
                    if pid not in accounting_current and pid in previous:
                        accounting_current[pid] = previous[pid]
                line = (json.dumps(snapshot_json(current, current_load), sort_keys=True) + "\n").encode()
                if during_bytes + len(line) > max_host:
                    violations.append(f"during host snapshot exceeded cap: {max_host}")
                    break
                during.write(line.decode("utf-8"))
                during.flush()
                during_bytes += len(line)
                violations.extend(
                    quiet_violations(
                        previous,
                        accounting_current,
                        owned=owned,
                        load_before=before_load[0],
                        max_load_delta=max_load_delta,
                        previous_cpu=previous_cpu,
                        current_cpu=current_cpu,
                        check_load_delta=False,
                        absolute_load_limit=None,
                    )
                )
                previous = accounting_current
                previous_cpu = current_cpu
                peak_rss = max(peak_rss, process_rss(current, owned))
                if peak_rss > max_rss:
                    violations.append(f"RSS exceeded cap: {peak_rss} > {max_rss}")
                target_after = safe_target_check(target, max_target)
                if target_after > target_before:
                    violations.append(f"target grew during child: {target_before} -> {target_after}")
                if report_path.exists() and report_path.stat().st_size > max_report:
                    violations.append(f"report exceeded cap: {report_path}")
                if stdout_capture.overflow or stderr_capture.overflow:
                    violations.append("stdout/stderr exceeded output cap")
                if violations:
                    break
            if time.monotonic() - started > timeout_seconds and not timeout:
                timeout = True
                violations.append(f"child exceeded timeout: {timeout_seconds}s")
            if process.poll() is None or violations or timeout:
                require(pgid is not None, "child process group identity was not retained")
                forced_cleanup = terminate_group(
                    process,
                    owned,
                    pgid,
                    runner_pid=parent_pid,
                    preexisting_pids=preexisting_pids,
                )
            else:
                process.wait(timeout=5)
        finally:
            if process is not None:
                forced_cleanup = ensure_child_reaped(
                    process,
                    owned,
                    pgid,
                    forced_cleanup,
                    runner_pid=parent_pid,
                    preexisting_pids=preexisting_pids,
                )
    if stdout_capture is not None:
        stdout_capture.join(timeout=5)
        if stdout_capture.is_alive():
            violations.append("stdout capture thread did not join")
    if stderr_capture is not None:
        stderr_capture.join(timeout=5)
        if stderr_capture.is_alive():
            violations.append("stderr capture thread did not join")
    if stdout_capture is not None and stdout_capture.error:
        violations.append(stdout_capture.error)
    if stderr_capture is not None and stderr_capture.error:
        violations.append(stderr_capture.error)
    require(process is not None, "child was not started")
    if pgid is None:
        raise TimingError("child process group identity was not retained")
    after = proc_snapshot()
    after_load = load_average()
    after_cpu = cpu_counters(cpu)
    owned = refresh_owned(after, owned, process.pid, pgid, process.pid)
    for pid, entry in after.items():
        if pid not in preexisting_pids and int(entry.get("ppid", -1)) == parent_pid:
            owned[pid] = int(entry["starttime"])
    owned_seen.update(owned)
    live = live_owned(owned, root_pid=process.pid)
    if live or process_group_exists(pgid):
        forced_cleanup = terminate_group(
            process,
            owned,
            pgid,
            runner_pid=parent_pid,
            preexisting_pids=preexisting_pids,
        ) or forced_cleanup
        violations.append(f"child ownership survived normal exit: {sorted(live)}")
        after = proc_snapshot()
        after_load = load_average()
        after_cpu = cpu_counters(cpu)
        owned = refresh_owned(after, owned, process.pid, pgid, process.pid)
        owned_seen.update(owned)
        live = live_owned(owned, root_pid=process.pid)
        if live or process_group_exists(pgid):
            raise TimingError(f"child ownership survived cleanup: {sorted(live)}")
    source_after: dict[str, str] | None = None
    try:
        source_after = verify_source(source_cwd, revision, source_tree)
    except TimingError as error:
        violations.append(f"source changed during child: {error}")
    post_violations = quiet_violations(
        quiet_probe,
        after,
        owned=owned_seen,
        load_before=before_load[0],
        max_load_delta=max_load_delta,
        previous_cpu=previous_cpu,
        current_cpu=after_cpu,
        check_load_delta=True,
        absolute_load_limit=None,
    )
    violations.extend(post_violations)
    descendants_left = sorted(pid for pid in live if pid != process.pid)
    if descendants_left:
        violations.append(f"descendants survived child exit: {descendants_left}")
    if forced_cleanup:
        violations.append("forced process-group cleanup was required")
    target_after_identity = target_identity(target, max_target)
    target_after = int(target_after_identity["bytes"])
    if target_after != target_before or target_after_identity["sha256"] != target_before_identity["sha256"]:
        violations.append(
            "target changed during child: "
            f"bytes {target_before} -> {target_after}, "
            f"sha256 {target_before_identity['sha256']} -> {target_after_identity['sha256']}"
        )
    try:
        binary_after = verify_binary(binary, binary_identity, f"{endpoint} binary after {policy}/{label}")
    except TimingError as error:
        binary_after = None
        violations.append(f"binary changed during child: {error}")
    validation: dict[str, Any] | None = None
    if process.returncode == 0 and not violations and not timeout:
        validation = validate_report(report_path, endpoint, revision, policy, warmup, samples, max_report)
    elif process.returncode != 0:
        violations.append(f"child exited with status {process.returncode}")
    stdout_identity = {"path": str(stdout_path), "bytes": stdout_capture.bytes if stdout_capture else 0, "sha256": sha256_file(stdout_path)[0]}
    stderr_identity = {"path": str(stderr_path), "bytes": stderr_capture.bytes if stderr_capture else 0, "sha256": sha256_file(stderr_path)[0]}
    during_identity = {"path": str(during_path), "bytes": during_path.stat().st_size, "sha256": sha256_file(during_path)[0]}
    report_identity = {
        "path": str(report_path),
        "bytes": report_path.stat().st_size if report_path.exists() else 0,
        "sha256": sha256_file(report_path)[0] if report_path.exists() else None,
    }
    event = {
        "schema": SCHEMA,
        "endpoint": endpoint,
        "policy": policy,
        "label": label,
        "argv": command,
        "cwd": str(source_cwd),
        "revision": revision,
        "source": {"before": source_before, "after": source_after, "expected_tree": source_tree},
        "binary": {"expected": dict(binary_identity), "before": binary_before, "after": binary_after},
        "started_utc": started_utc,
        "ended_utc": utc_now(),
        "returncode": process.returncode,
        "timeout": timeout,
        "forced_cleanup": forced_cleanup,
        "cpu": cpu,
        "rss_peak_bytes": peak_rss,
        "rss_cap_bytes": max_rss,
        "target": {
            "path": str(target),
            "bytes_before": target_before,
            "bytes_after": target_after,
            "cap_bytes": max_target,
            "identity_before": target_before_identity,
            "identity_after": target_after_identity,
        },
        "stdout": {**stdout_identity, "overflow": stdout_capture.overflow if stdout_capture else True},
        "stderr": {**stderr_identity, "overflow": stderr_capture.overflow if stderr_capture else True},
        "report": report_identity,
        "host": {
            "before": before_host,
            "probe": probe_host,
            "during_jsonl": during_identity,
            "after": host_snapshot_descriptor(host_after_path, after, after_load, max_host),
        },
        "violations": violations,
        "validation": validation,
    }
    write_json_exclusive(event_path, event, max_event)
    if violations:
        raise TimingError(f"{endpoint}/{policy}/{label} rejected: {violations}")
    require(validation is not None, f"{endpoint}/{policy}/{label} has no validation result")
    return event


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-binary", required=True, type=Path)
    parser.add_argument("--candidate-binary", required=True, type=Path)
    parser.add_argument("--baseline-revision", required=True)
    parser.add_argument("--candidate-revision", required=True)
    parser.add_argument("--baseline-source", required=True, type=Path)
    parser.add_argument("--candidate-source", required=True, type=Path)
    parser.add_argument("--baseline-tree", required=True)
    parser.add_argument("--candidate-tree", required=True)
    parser.add_argument("--target-baseline", required=True, type=Path)
    parser.add_argument("--target-candidate", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--tmpdir", required=True, type=Path)
    parser.add_argument("--cpu-lock", required=True, type=Path)
    parser.add_argument("--cpu", required=True, type=int)
    parser.add_argument("--samples", type=int, default=30)
    parser.add_argument("--warmup", type=int, default=3)
    parser.add_argument("--timeout-seconds", type=float, default=DEFAULT_TIMEOUT_SECONDS)
    parser.add_argument("--poll-seconds", type=float, default=DEFAULT_POLL_SECONDS)
    parser.add_argument("--max-rss-bytes", type=int, default=DEFAULT_MAX_RSS_BYTES)
    parser.add_argument("--max-target-bytes", type=int, default=DEFAULT_MAX_TARGET_BYTES)
    parser.add_argument("--max-output-bytes", type=int, default=DEFAULT_MAX_OUTPUT_BYTES)
    parser.add_argument("--max-report-bytes", type=int, default=DEFAULT_MAX_REPORT_BYTES)
    parser.add_argument("--max-host-bytes", type=int, default=DEFAULT_MAX_HOST_BYTES)
    parser.add_argument("--max-event-bytes", type=int, default=DEFAULT_MAX_EVENT_BYTES)
    parser.add_argument("--max-receipt-bytes", type=int, default=DEFAULT_MAX_RECEIPT_BYTES)
    parser.add_argument("--max-load-delta", type=float, default=1.0)
    parser.add_argument("--max-prerun-load", type=float, default=DEFAULT_MAX_PRERUN_LOAD)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        require(
            len(args.baseline_revision) == 40
            and len(args.candidate_revision) == 40
            and len(args.baseline_tree) == 40
            and len(args.candidate_tree) == 40,
            "revisions and trees must be 40-character hashes",
        )
        require(args.samples > 0 and args.warmup >= 0, "samples/warmup are invalid")
        require(args.timeout_seconds > 0 and args.poll_seconds > 0, "timeouts are invalid")
        require(args.max_rss_bytes > 0 and args.max_target_bytes > 0, "resource caps are invalid")
        require(
            args.max_output_bytes > 0
            and args.max_report_bytes > 0
            and args.max_host_bytes > 0
            and args.max_event_bytes > 0
            and args.max_receipt_bytes > 0,
            "output caps are invalid",
        )
        require(math.isfinite(args.max_load_delta) and args.max_load_delta >= 0, "max-load-delta is invalid")
        require(math.isfinite(args.max_prerun_load) and args.max_prerun_load > 0, "max-prerun-load is invalid")
        require(args.output_dir.is_absolute() and not args.output_dir.exists() and not args.output_dir.is_symlink(), "output-dir must be a fresh absent absolute path")
        sources = {"baseline": args.baseline_source.resolve(), "candidate": args.candidate_source.resolve()}
        trees = {"baseline": args.baseline_tree, "candidate": args.candidate_tree}
        revisions = {"baseline": args.baseline_revision, "candidate": args.candidate_revision}
        source_identities = {endpoint: verify_source(sources[endpoint], revisions[endpoint], trees[endpoint]) for endpoint in ("baseline", "candidate")}
        args.output_dir.mkdir(parents=True)
        for directory in (args.target_baseline, args.target_candidate):
            require(not directory.is_symlink(), f"target is symlinked: {directory}")
            safe_target_check(directory, args.max_target_bytes)
        require(not args.tmpdir.is_symlink(), f"temporary directory is symlinked: {args.tmpdir}")
        args.tmpdir.mkdir(parents=True, exist_ok=True)
        args.output_dir = args.output_dir.resolve()
        args.tmpdir = args.tmpdir.resolve()
        for binary in (args.baseline_binary, args.candidate_binary):
            require(not binary.is_symlink(), f"binary is symlinked: {binary}")
        binaries = {"baseline": args.baseline_binary.resolve(), "candidate": args.candidate_binary.resolve()}
        binary_identities: dict[str, dict[str, Any]] = {}
        for endpoint, binary in binaries.items():
            if binary.is_symlink() or not binary.is_file() or not os.access(binary, os.X_OK):
                raise TimingError(f"{endpoint} binary is missing, symlinked, or non-executable: {binary}")
            digest, size = sha256_file(binary)
            binary_identities[endpoint] = {"path": str(binary), "sha256": digest, "bytes": size}
        targets = {"baseline": args.target_baseline.resolve(), "candidate": args.target_candidate.resolve()}
        events: list[dict[str, Any]] = []
        previous_affinity = os.sched_getaffinity(0)
        require(args.cpu in previous_affinity, f"selected CPU is not allowed: {args.cpu}")
        previous_subreaper = set_child_subreaper(True)
        cpu_lock: Any | None = None
        try:
            os.sched_setaffinity(0, {args.cpu})
            cpu_lock = acquire_cpu_lock(args.cpu_lock.resolve())
            for policy in ("forward", "exact"):
                for label in ABBA_ORDER:
                    endpoint = "baseline" if label.startswith("a") else "candidate"
                    events.append(
                        run_child(
                            output=args.output_dir,
                            endpoint=endpoint,
                            binary=binaries[endpoint],
                            binary_identity=binary_identities[endpoint],
                            revision=revisions[endpoint],
                            source_tree=trees[endpoint],
                            policy=policy,
                            label=label,
                            source_cwd=sources[endpoint],
                            target=targets[endpoint],
                            tmpdir=args.tmpdir,
                            cpu=args.cpu,
                            samples=args.samples,
                            warmup=args.warmup,
                            timeout_seconds=args.timeout_seconds,
                            poll_seconds=args.poll_seconds,
                            max_rss=args.max_rss_bytes,
                            max_target=args.max_target_bytes,
                            max_output=args.max_output_bytes,
                            max_report=args.max_report_bytes,
                            max_host=args.max_host_bytes,
                            max_event=args.max_event_bytes,
                            max_load_delta=args.max_load_delta,
                            max_prerun_load=args.max_prerun_load,
                        )
                    )
        finally:
            if cpu_lock is not None:
                fcntl.flock(cpu_lock.fileno(), fcntl.LOCK_UN)
                cpu_lock.close()
            os.sched_setaffinity(0, previous_affinity)
            set_child_subreaper(bool(previous_subreaper))
        final_binaries = {
            endpoint: verify_binary(binary, binary_identities[endpoint], f"{endpoint} binary after ABBA")
            for endpoint, binary in binaries.items()
        }
        receipt = {
            "schema": SCHEMA,
            "classification": "latency-only capture; allocator is not applicable",
            "sources": source_identities,
            "source_revisions": revisions,
            "source_trees": trees,
            "binaries": final_binaries,
            "cpu": args.cpu,
            "cpu_lock": str(args.cpu_lock.resolve()),
            "abba_order": list(ABBA_ORDER),
            "policies": ["forward", "exact"],
            "samples": args.samples,
            "warmup": args.warmup,
            "resource_caps": {
                "rss_bytes": args.max_rss_bytes,
                "target_bytes": args.max_target_bytes,
                "output_bytes": args.max_output_bytes,
                "report_bytes": args.max_report_bytes,
                "host_bytes": args.max_host_bytes,
                "event_bytes": args.max_event_bytes,
                "receipt_bytes": args.max_receipt_bytes,
                "timeout_seconds": args.timeout_seconds,
                "max_load_delta": args.max_load_delta,
                "max_prerun_load": args.max_prerun_load,
            },
            "events": [event["endpoint"] + "/" + event["policy"] + "/" + event["label"] for event in events],
            "qualification_status": "capture_complete_pending_analysis",
        }
        write_json_exclusive(args.output_dir / "receipt.json", receipt, args.max_receipt_bytes)
        return 0
    except (OSError, TimingError, ValueError) as error:
        print(f"opc read-ahead timing rejected: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
