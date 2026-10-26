#!/usr/bin/env python3
"""Run the workspace gate with a bounded, resumable, non-destructive policy.

This runner is deliberately outside the repository.  It records the exact
source revision, command lines, environment, log sizes and SHA-256 digests.
It never invokes the historical target-headroom cleanup.  The only optional
artifact release is performed by safe_chunked_tests.py, which moves
successfully completed Cargo-reported debug test/example executables to an
external hashed quarantine after their child process has exited.
"""
from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
from pathlib import Path
import shlex
import signal
import shutil
import subprocess
import sys
import threading
import time
import traceback
import tomllib

ROOT_DEFAULT = Path("/home/zhuhe/code/litchi-worktrees/unified-ops-api")
COORD = Path("/home/zhuhe/code/litchi-worktrees/.coord")
SAFE_CHUNKED = Path(__file__).with_name("safe_chunked_tests.py")
MAX_TARGET = 20_000_000_000
LIVE_TARGET = 19_000_000_000
START_TARGET = 18_000_000_000
MIN_FREE = 12_000_000_000
SAFE_CHUNK_SIZE = 2
BUILDER_EXECUTABLES = frozenset({"cargo", "rustc", "rustdoc"})
POST_EXIT_POLL_SECONDS = 0.05
POST_EXIT_STABLE_SAMPLES = 2
POST_EXIT_TIMEOUT_SECONDS = 2.0
PRE_EXIT_POLL_SECONDS = 0.01
PROCESS_KILL_GRACE_SECONDS = 1.0
TARGET_MEASURE_RETRIES = 5
TARGET_MEASURE_DELAY_SECONDS = 0.02
PR_SET_CHILD_SUBREAPER = 36
_SUBREAPER_ENABLED = False
_SUBREAPER_ERROR: str | None = None


def sha256(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def git(root: Path, *args: str) -> str:
    return subprocess.check_output(["git", "-C", str(root), *args], text=True).strip()


def source_state(root: Path) -> dict:
    return {
        "commit": git(root, "rev-parse", "HEAD"),
        "status": git(root, "status", "--porcelain"),
        "branch": git(root, "branch", "--show-current"),
        "cargo_locks": lock_hashes(root),
    }


def lock_hashes(root: Path) -> dict[str, str]:
    """Hash locks belonging to commands this gate can actually invoke.

    The repository also contains many historical standalone benchmark locks
    under docs/performance/results. They are not Cargo workspaces used by this
    gate; scanning them on every step would add noise and substantial I/O.
    """
    relative_paths = (
        "Cargo.lock",
        "tools/perf-baseline/Cargo.lock",
        "tools/odf-raw-codec-consumer/Cargo.lock",
        "tools/native-resave/Cargo.lock",
        "crates/litchi-core/fuzz/Cargo.lock",
        "crates/litchi-ops/fuzz/Cargo.lock",
        "crates/litchi/fuzz/Cargo.lock",
        "crates/litchi-rtf/fuzz/Cargo.lock",
    )
    result: dict[str, str] = {}
    for relative in relative_paths:
        path = root / relative
        if path.is_file():
            result[relative] = sha256(path)
    return result


def process_table() -> dict[int, dict]:
    info: dict[int, dict] = {}
    for entry in Path("/proc").glob("[0-9]*"):
        try:
            pid = int(entry.name)
            command = (entry / "cmdline").read_bytes().replace(b"\0", b" ").decode(errors="replace").strip()
        except (FileNotFoundError, ProcessLookupError, PermissionError, ValueError):
            continue
        try:
            stat = (entry / "stat").read_text()
            close = stat.rfind(")")
            fields = stat[close + 2 :].split()
            ppid = int(fields[1])
            pgrp = int(fields[2])
            session = int(fields[3])
            start_time = int(fields[19])
        except (FileNotFoundError, ProcessLookupError, PermissionError, ValueError, IndexError):
            continue
        # Keep zombies in the identity table: an adopted escaped child has no
        # /proc/$pid/exe after it exits, but the runner still has to reap it.
        try:
            executable = Path(os.readlink(entry / "exe")).name
        except (FileNotFoundError, ProcessLookupError, PermissionError):
            executable = ""
        try:
            cwd = str(Path(os.readlink(entry / "cwd")))
        except (FileNotFoundError, ProcessLookupError, PermissionError):
            cwd = ""
        info[pid] = {
            "pid": pid,
            "ppid": ppid,
            "pgrp": pgrp,
            "session": session,
            "start_time": start_time,
            "command": command,
            "executable": executable,
            "cwd": cwd,
        }
    return info


def process_key(row: dict) -> tuple[int, int | None]:
    return (int(row["pid"]), row.get("start_time"))


def process_output(row: dict) -> dict:
    return {
        key: row.get(key)
        for key in ("pid", "ppid", "pgrp", "session", "start_time", "executable", "cwd", "command")
    }


def enable_child_subreaper() -> None:
    """Make this runner the reaper for descendants orphaned by a step.

    ``start_new_session`` gives each step a private process group, but a
    helper that calls ``setsid`` escapes that group. Linux subreaper adoption
    keeps such a helper in this runner's observable process tree without
    signalling or reaping any process that the runner did not start.
    """
    global _SUBREAPER_ENABLED, _SUBREAPER_ERROR
    if _SUBREAPER_ENABLED:
        return
    if sys.platform != "linux":
        raise RuntimeError("safe gate requires Linux child-subreaper support")
    try:
        libc = ctypes.CDLL(None, use_errno=True)
        result = libc.prctl(PR_SET_CHILD_SUBREAPER, 1, 0, 0, 0)
        if result != 0:
            error = ctypes.get_errno()
            raise OSError(error, os.strerror(error))
    except (OSError, AttributeError) as error:
        _SUBREAPER_ERROR = repr(error)
        raise RuntimeError(f"cannot enable child subreaper: {error}") from error
    _SUBREAPER_ENABLED = True
    _SUBREAPER_ERROR = None


def ensure_child_subreaper() -> None:
    """Give the hermetic Python row ownership of its test grandchildren."""
    libc = ctypes.CDLL(None, use_errno=True)
    if libc.prctl(PR_SET_CHILD_SUBREAPER, 1, 0, 0, 0) != 0:
        error = ctypes.get_errno()
        raise OSError(error, os.strerror(error))


def builders(exclude: set[int], root: Path, target: Path) -> list[dict]:
    rows = []
    info = process_table()
    info = {pid: row for pid, row in info.items() if pid not in exclude}
    root_related = set()
    root_text = str(root)
    target_text = str(target)
    litchi_prefixes = (
        "/home/zhuhe/code/litchi",
        "/home/zhuhe/code/litchi-worktrees/.coord",
    )

    def litchi_path(text: str) -> bool:
        return any(text == prefix or text.startswith(prefix + "/") or prefix in text for prefix in litchi_prefixes)

    for pid, row in info.items():
        if row["executable"] not in BUILDER_EXECUTABLES:
            continue
        if (
            litchi_path(row["cwd"])
            or litchi_path(row["command"])
            or row["cwd"] == root_text
            or row["cwd"].startswith(root_text + "/")
            or root_text in row["command"]
            or target_text in row["command"]
        ):
            root_related.add(pid)
    changed = True
    while changed:
        changed = False
        for pid, row in info.items():
            if pid in root_related or row["executable"] not in BUILDER_EXECUTABLES:
                continue
            if row["ppid"] in root_related:
                root_related.add(pid)
                changed = True
    for pid in sorted(root_related):
        rows.append({key: info[pid][key] for key in ("pid", "ppid", "pgrp", "executable", "cwd", "command")})
    return rows


def _path_related(row: dict, root: Path, target: Path) -> bool:
    root_text = str(root)
    target_text = str(target)
    return any(
        value == root_text
        or value.startswith(root_text + "/")
        or value == target_text
        or value.startswith(target_text + "/")
        or root_text in value
        or target_text in value
        for value in (row.get("cwd", ""), row.get("command", ""))
    )


def _descendant_pids(info: dict[int, dict], roots: set[int]) -> set[int]:
    descendants = set(roots)
    changed = True
    while changed:
        changed = False
        for pid, row in info.items():
            if pid not in descendants and row["ppid"] in descendants:
                descendants.add(pid)
                changed = True
    return descendants


def _runner_ancestry(info: dict[int, dict], runner_pid: int) -> tuple[set[int], set[tuple[int, int | None]]]:
    """Return the runner and its pre-existing parent chain.

    Ownership is allowed to follow children adopted by the runner, but it
    must never infer ownership of the runner itself (or of a coordinator
    above it) from a process-group match.  Keep both PID and ``(PID,
    start_time)`` identities: the PID set is a conservative no-signal guard
    if an ancestor disappears, while the identity set retains the usual PID
    reuse protection for rows that remain observable.
    """
    pids = {runner_pid}
    identities: set[tuple[int, int | None]] = set()
    current = runner_pid
    visited: set[int] = set()
    while current not in visited:
        visited.add(current)
        row = info.get(current)
        if row is None:
            break
        pids.add(current)
        identities.add(process_key(row))
        parent = int(row.get("ppid", 0))
        if parent <= 0:
            break
        current = parent
    return pids, identities


def _protected_runner_process(ownership: dict, row: dict) -> bool:
    """Whether a row belongs to this runner or one of its ancestors."""
    pid = int(row["pid"])
    return (
        pid in ownership.get("protected_runner_pids", set())
        or process_key(row) in ownership.get("protected_runner_identities", set())
    )


def new_ownership(
    process_pid: int,
    root: Path,
    target: Path,
    prelaunch: dict[int, dict],
    launch: dict[int, dict],
) -> dict:
    """Create identities used to follow a step across reparenting and setsid."""
    root_row = launch.get(process_pid)
    runner_pid = os.getpid()
    ancestry_info = dict(prelaunch)
    ancestry_info.update(launch)
    protected_runner_pids, protected_runner_identities = _runner_ancestry(ancestry_info, runner_pid)
    return {
        "process_pid": process_pid,
        "runner_pid": runner_pid,
        "protected_runner_pids": protected_runner_pids,
        "protected_runner_identities": protected_runner_identities,
        "root_identity": process_key(root_row) if root_row else (process_pid, None),
        "root_pgrps": {root_row["pgrp"]} if root_row else {process_pid},
        "known": {},
        "pgrps": set(),
        "lock": threading.RLock(),
        "preexisting_runner_children": {
            process_key(row)
            for row in prelaunch.values()
            if row["ppid"] == os.getpid()
        },
        "root": str(root),
        "target": str(target),
    }


def _capture_owned(ownership: dict, info: dict[int, dict]) -> None:
    lock = ownership.get("lock")
    if lock is None:
        return _capture_owned_unlocked(ownership, info)
    with lock:
        _capture_owned_unlocked(ownership, info)


def _capture_owned_unlocked(ownership: dict, info: dict[int, dict]) -> None:
    """Capture live identities before a fast child can disappear or reparent."""
    process_pid = ownership["process_pid"]
    root_row = info.get(process_pid)
    root_matches = root_row and process_key(root_row) == ownership["root_identity"]
    if root_matches:
        ownership["root_pgrps"].add(root_row["pgrp"])
    roots = {process_pid} if root_matches else set()
    descendants = _descendant_pids(info, roots)
    runner_pid = ownership["runner_pid"]
    for pid, row in info.items():
        identity = process_key(row)
        if _protected_runner_process(ownership, row):
            continue
        adopted = (
            row["ppid"] == runner_pid
            and identity not in ownership["preexisting_runner_children"]
        )
        if pid in descendants or row["pgrp"] in ownership["root_pgrps"] or adopted:
            ownership["known"][identity] = process_output(row)
            ownership["pgrps"].add(row["pgrp"])

    # A captured helper may fork after its original parent has exited. Follow
    # that helper's current descendants as well, even after it changes session.
    known_pids = {identity[0] for identity in ownership["known"]}
    changed = True
    while changed:
        changed = False
        for pid, row in info.items():
            identity = process_key(row)
            if _protected_runner_process(ownership, row):
                continue
            parent = info.get(row["ppid"])
            parent_known = parent is not None and process_key(parent) in ownership["known"]
            if parent_known and identity not in ownership["known"]:
                ownership["known"][identity] = process_output(row)
                ownership["pgrps"].add(row["pgrp"])
                known_pids.add(pid)
                changed = True


def owned_processes(
    process_pid: int,
    info: dict[int, dict] | None = None,
    ownership: dict | None = None,
) -> list[dict]:
    """Return live processes owned by a step, including reparented helpers."""
    info = process_table() if info is None else info
    if ownership is None:
        descendants = _descendant_pids(info, {process_pid})
        return [process_output(row) for pid, row in info.items() if pid in descendants or row["pgrp"] == process_pid]
    _capture_owned(ownership, info)
    lock = ownership.get("lock")
    if lock is None:
        known = set(ownership["known"])
    else:
        with lock:
            known = set(ownership["known"])
    return [
        process_output(row)
        for row in info.values()
        if not _protected_runner_process(ownership, row)
        and (process_key(row) in known or row["pgrp"] in ownership["pgrps"])
    ]


def owned_builders(
    process_pid: int,
    root: Path,
    target: Path,
    info: dict[int, dict] | None = None,
    ownership: dict | None = None,
) -> list[dict]:
    """Return builders still owned by a completed step or its source/target."""
    info = process_table() if info is None else info
    owned = {process_key(row) for row in owned_processes(process_pid, info, ownership)}
    rows = []
    for row in info.values():
        if row["executable"] not in BUILDER_EXECUTABLES:
            continue
        if ownership is not None and _protected_runner_process(ownership, row):
            continue
        if process_key(row) in owned or _path_related(row, root, target):
            rows.append(process_output(row))
    return rows


def _owned_identity_rows(process_pid: int, ownership: dict | None) -> list[dict]:
    """Return only current rows whose PID and start time were captured."""
    rows = owned_processes(process_pid, ownership=ownership)
    if ownership is None:
        return rows
    lock = ownership.get("lock")
    if lock is None:
        known = set(ownership["known"])
    else:
        with lock:
            known = set(ownership["known"])
    return [row for row in rows if process_key(row) in known]


def _signal_rows(rows: list[dict], signal_number: int) -> None:
    for row in rows:
        pid = int(row["pid"])
        if pid == os.getpid():
            continue
        try:
            os.kill(pid, signal_number)
        except (ProcessLookupError, PermissionError):
            continue


def _reap_rows(rows: list[dict], process_pid: int) -> None:
    for row in rows:
        pid = int(row["pid"])
        # Popen owns the top-level wait; reaping it here would race the
        # caller's process.wait(). Adopted descendants are safe to reap.
        if pid in {os.getpid(), process_pid}:
            continue
        try:
            os.waitpid(pid, os.WNOHANG)
        except (ChildProcessError, ProcessLookupError, PermissionError):
            continue


def _terminate_owned_tree(process_pid: int, ownership: dict | None, reason: str) -> dict:
    """Terminate and reap only identities owned by one gate step.

    Individual PID signals are intentional here: a helper that calls
    ``setsid`` is outside the original process group, while the captured
    ``(pid,start_time)`` identities prevent a recycled PID from being killed.
    The loop also captures children forked during cleanup before escalating to
    SIGKILL.
    """
    term_sent = False
    kill_sent = False
    seen: dict[tuple[int, int | None], dict] = {}
    deadline = time.monotonic() + PROCESS_KILL_GRACE_SECONDS
    while True:
        rows = _owned_identity_rows(process_pid, ownership)
        for row in rows:
            seen[process_key(row)] = process_output(row)
        _reap_rows(rows, process_pid)
        if not rows:
            return {
                "reason": reason,
                "signal": "SIGKILL" if kill_sent else ("SIGTERM" if term_sent else None),
                "seen": list(seen.values()),
                "remaining": [],
            }
        if not term_sent:
            _signal_rows(rows, signal.SIGTERM)
            term_sent = True
        if time.monotonic() >= deadline:
            break
        time.sleep(PRE_EXIT_POLL_SECONDS)

    rows = _owned_identity_rows(process_pid, ownership)
    for row in rows:
        seen[process_key(row)] = process_output(row)
    _signal_rows(rows, signal.SIGKILL)
    kill_sent = True
    deadline = time.monotonic() + PROCESS_KILL_GRACE_SECONDS
    while time.monotonic() < deadline:
        rows = _owned_identity_rows(process_pid, ownership)
        for row in rows:
            seen[process_key(row)] = process_output(row)
        _reap_rows(rows, process_pid)
        if not rows:
            return {
                "reason": reason,
                "signal": "SIGKILL",
                "seen": list(seen.values()),
                "remaining": [],
            }
        time.sleep(PRE_EXIT_POLL_SECONDS)
    rows = _owned_identity_rows(process_pid, ownership)
    for row in rows:
        seen[process_key(row)] = process_output(row)
    _reap_rows(rows, process_pid)
    return {
        "reason": reason,
        "signal": "SIGKILL",
        "seen": list(seen.values()),
        "remaining": rows,
    }


def post_exit_guard(process_pid: int, root: Path, target: Path, ownership: dict | None = None) -> dict:
    """Require owned builders to exit and the selected target to stabilize."""
    deadline = time.monotonic() + POST_EXIT_TIMEOUT_SECONDS
    previous_size: int | None = None
    stable_samples = 0
    samples = []
    while True:
        info = process_table()
        owned = owned_processes(process_pid, info, ownership)
        _reap_rows(owned, process_pid)
        if owned:
            # Re-query after WNOHANG so a naturally exited adopted child is
            # not mistaken for a live lifetime leak.
            info = process_table()
            owned = owned_processes(process_pid, info, ownership)
        leaked = owned_builders(process_pid, root, target, info, ownership)
        size = target_bytes(target)
        samples.append({"at": time.time(), "target_bytes": size, "owned": owned, "builders": leaked})
        if leaked:
            cleanup = _terminate_owned_tree(process_pid, ownership, "owned-builder-after-exit")
            raise RuntimeError(f"owned builder remained after child exit: {leaked}; cleanup={cleanup}")
        if owned:
            stable_samples = 0
            if time.monotonic() >= deadline:
                cleanup = _terminate_owned_tree(process_pid, ownership, "owned-child-timeout")
                raise RuntimeError(f"owned child remained after child exit: {owned}; cleanup={cleanup}")
            time.sleep(POST_EXIT_POLL_SECONDS)
            continue
        if previous_size is not None and size == previous_size:
            stable_samples += 1
        else:
            stable_samples = 0
        if stable_samples >= POST_EXIT_STABLE_SAMPLES:
            return {"target_bytes": size, "samples": samples}
        if time.monotonic() >= deadline:
            cleanup = _terminate_owned_tree(process_pid, ownership, "target-stability-timeout")
            raise RuntimeError(f"target did not stabilize after child exit: {samples[-1]}; cleanup={cleanup}")
        previous_size = size
        time.sleep(POST_EXIT_POLL_SECONDS)


def target_bytes(target: Path) -> int:
    if not target.exists():
        return 0
    last_error: str | None = None
    for attempt in range(TARGET_MEASURE_RETRIES):
        result = subprocess.run(
            ["du", "-sb", str(target)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )
        fields = result.stdout.split()
        if result.returncode == 0 and fields:
            try:
                return int(fields[0])
            except ValueError:
                pass
        last_error = f"du exit={result.returncode}: {result.stderr.strip()[-500:]}"
        if attempt + 1 < TARGET_MEASURE_RETRIES:
            time.sleep(TARGET_MEASURE_DELAY_SECONDS)
    raise RuntimeError(f"target measurement failed after {TARGET_MEASURE_RETRIES} attempts: {last_error}")


def _terminate_process_group(process: subprocess.Popen, ownership: dict | None = None) -> dict | None:
    """Stop the isolated gate command before it reaches the hard target cap."""
    if ownership is not None:
        return _terminate_owned_tree(process.pid, ownership, "target-live-overflow")
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return None
    deadline = time.monotonic() + PROCESS_KILL_GRACE_SECONDS
    while process.poll() is None and time.monotonic() < deadline:
        time.sleep(PRE_EXIT_POLL_SECONDS)
    if process.poll() is None:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
    return None


def _watch_target(
    process: subprocess.Popen,
    target: Path,
    stop: threading.Event,
    overflow: list[dict],
    ownership: dict | None = None,
) -> dict:
    """Enforce the live cap for every gate command, not just Cargo chunks."""
    try:
        maximum = target_bytes(target)
    except Exception as error:
        detail = {
            "reason": "target measurement failed",
            "error": repr(error),
            "samples": 0,
            "last_sampled_bytes": None,
        }
        overflow.append(detail)
        cleanup = _terminate_process_group(process, ownership)
        if cleanup is not None:
            detail["cleanup"] = cleanup
        return {"max_sampled_bytes": None, "samples": 0, "live_target": LIVE_TARGET, "measurement_error": repr(error)}
    samples = 0
    exited_at: float | None = None
    while not stop.wait(PRE_EXIT_POLL_SECONDS):
        if process.poll() is not None:
            if ownership is None:
                break
            owned = _owned_identity_rows(process.pid, ownership)
            if not owned:
                if exited_at is None:
                    exited_at = time.monotonic()
                if time.monotonic() - exited_at >= POST_EXIT_TIMEOUT_SECONDS:
                    break
            else:
                exited_at = None
        try:
            size = target_bytes(target)
        except Exception as error:
            detail = {
                "reason": "target measurement failed",
                "error": repr(error),
                "samples": samples,
                "last_sampled_bytes": maximum,
            }
            overflow.append(detail)
            cleanup = _terminate_process_group(process, ownership)
            if cleanup is not None:
                detail["cleanup"] = cleanup
            break
        maximum = max(maximum, size)
        samples += 1
        if size >= LIVE_TARGET:
            detail = {
                "target_bytes": size,
                "live_target": LIVE_TARGET,
                "max_sampled_bytes": maximum,
                "last_sampled_bytes": size,
                "samples": samples,
            }
            overflow.append(detail)
            cleanup = _terminate_process_group(process, ownership)
            if cleanup is not None:
                detail["cleanup"] = cleanup
            break
    return {"max_sampled_bytes": maximum, "samples": samples, "live_target": LIVE_TARGET}


def _run_target_guard(
    process: subprocess.Popen,
    target: Path,
    stop: threading.Event,
    overflow: list[dict],
    errors: list[dict],
    ownership: dict,
    stats: dict,
) -> None:
    """Convert any watcher exception into a recorded, fail-closed cleanup."""
    try:
        stats.update(_watch_target(process, target, stop, overflow, ownership))
    except BaseException as error:
        detail = {
            "reason": "target guard exception",
            "error": repr(error),
            "traceback": traceback.format_exc(),
        }
        try:
            cleanup = _terminate_process_group(process, ownership)
            if cleanup is not None:
                detail["cleanup"] = cleanup
        except BaseException as cleanup_error:
            detail["cleanup_error"] = repr(cleanup_error)
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except (ProcessLookupError, PermissionError):
                pass
        errors.append(detail)


def preflight(root: Path, target: Path, source: dict, isolated: bool = False) -> dict:
    current = source_state(root)
    if current["commit"] != source["commit"] or current["status"]:
        raise RuntimeError(f"source changed or became dirty: {current}")
    size = target_bytes(target)
    free = shutil.disk_usage(target.parent).free
    if size > START_TARGET:
        raise RuntimeError(f"target requires headroom before command: {size} bytes > {START_TARGET}")
    if free < MIN_FREE:
        raise RuntimeError(f"free disk below 12 GiB: {free}")
    active = builders({os.getpid(), os.getppid()}, root, target)
    if isolated and active:
        raise RuntimeError(f"isolated command has unrelated build processes: {active}")
    return {"started": time.time(), "target_bytes": size, "free_bytes": free, "builders": active}


def cargo_env(root: Path, target: Path, state_dir: Path) -> dict[str, str]:
    env = os.environ.copy()
    env.pop("LITCHI_CAPPED", None)
    env.update({
        "PATH": f"{Path(__file__).resolve().with_name('bin')}:{COORD / 'bin'}:{env.get('PATH', '')}",
        "CARGO_TARGET_DIR": str(target),
        "LITCHI_GATE_TARGET_ROOT": str(target),
        "CARGO_NET_OFFLINE": "true",
        "LITCHI_MEM_MAX": "20G",
        "CARGO_BUILD_JOBS": "1",
        "RUST_TEST_THREADS": "1",
        "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_DEV_DEBUG": "0",
        "CARGO_PROFILE_TEST_DEBUG": "0",
        "CARGO_TERM_COLOR": "never",
        "TMPDIR": str(state_dir / "tmp"),
    })
    Path(env["TMPDIR"]).mkdir(parents=True, exist_ok=True)
    return env


def direct(
    name: str,
    argv: list[str],
    *,
    isolated: bool = False,
    env: dict[str, str] | None = None,
    pass_fds: tuple[int, ...] = (),
    target_guard: bool = True,
) -> dict:
    step = {"name": name, "kind": "argv", "argv": argv, "isolated": isolated, "env": env or {}}
    if not target_guard:
        # Python tooling tests are hermetic and must not inherit the Cargo
        # target override.  The runner still monitors the shared target path;
        # only Cargo-producing steps receive CARGO_TARGET_DIR in their env.
        step["target_guard"] = False
    if pass_fds:
        step["pass_fds"] = list(pass_fds)
    return step


def shell(name: str, command: str, *, isolated: bool = False, env: dict[str, str] | None = None) -> dict:
    return {"name": name, "kind": "shell", "command": command, "isolated": isolated, "env": env or {}}


PPT_AUXILIARY_NAMES = (
    "ppt-scoped-build",
    "ppt-scoped-correctness-suite",
    "ppt-scoped-correctness",
    "ppt-scoped-correctness-validate",
)
PPT_CANDIDATE_LOCK_SHA256 = (
    "4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1"
)


def feature_names(root: Path) -> tuple[list[str], list[str]]:
    with (root / "crates/litchi-ops/Cargo.toml").open("rb") as stream:
        ops = tomllib.load(stream).get("features", {})
    all_features = [name for name in ops if name not in {"default", "encryption"}]
    formats = [name for name in ("xlsx", "pptx", "docx", "xls", "xlsb", "doc", "ppt", "rtf", "odt", "ods", "odp", "oth", "odm", "odg", "odc", "odi", "odb", "odf-formula", "pages", "keynote", "numbers") if name in ops]
    return all_features, formats


def make_ppt_steps(root: Path, state_dir: Path, target: Path) -> list[dict]:
    """Return the opt-in PPT qualification steps.

    These steps are intentionally separate from make_steps(): an ordinary
    full gate must retain its historical plan and count.  The coordinator
    selects this four-step auxiliary plan explicitly with --only or --start-at.
    """
    ppt_fixture = root / "test-data/unified-ops/0294/video.ppt"
    ppt_host_guard = root / "docs/unified-ops/results/0362g/ppt_host_process_guard.py"
    ppt_lock_guard = root / "docs/unified-ops/results/0362g/validate_ppt_candidate_lock.py"
    ppt_correctness_validator = root / "docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py"
    ppt_correctness_binary = target / "release/examples/scoped_patch_correctness"
    ppt_correctness_receipt = state_dir / "ppt-scoped-correctness.json"
    lock = root / "Cargo.lock"
    guarded_cargo = [
        "python3",
        str(ppt_lock_guard),
        str(lock),
        PPT_CANDIDATE_LOCK_SHA256,
        "--",
        "python3",
        str(ppt_host_guard),
        "--",
    ]
    correctness_commands = (
        "set -eu; "
        "cargo test --offline --locked -p litchi-ppt --all-features --lib scoped::tests -- --test-threads=1; "
        "cargo test --offline --locked -p litchi-ppt --all-features --lib -- --test-threads=1; "
        "cargo test --offline --locked -p litchi-cfb --all-features --lib -- --test-threads=1; "
        "cargo test --offline --locked -p litchi-ppt --all-features --tests -- --test-threads=1; "
        "cargo test --offline --locked -p litchi-cfb --all-features --tests -- --test-threads=1; "
        "cargo clippy --offline --locked -p litchi-ppt --all-features --all-targets -- -D warnings; "
        "cargo clippy --offline --locked -p litchi-cfb --all-features --all-targets -- -D warnings"
    )
    return [
        direct(
            "ppt-scoped-build",
            [
                *guarded_cargo,
                "cargo",
                "build",
                "--offline",
                "--locked",
                "--release",
                "-p",
                "litchi-ppt",
                "--example",
                "scoped_patch_probe",
                "--example",
                "scoped_patch_correctness",
            ],
            isolated=True,
        ),
        direct(
            "ppt-scoped-correctness-suite",
            [
                *guarded_cargo,
                "sh",
                "-c",
                correctness_commands,
            ],
            isolated=True,
            env={"RUST_TEST_THREADS": "1"},
        ),
        direct(
            "ppt-scoped-correctness",
            [
                "python3",
                str(ppt_host_guard),
                "--",
                str(ppt_correctness_binary),
                str(ppt_fixture),
                str(ppt_correctness_receipt),
            ],
            isolated=True,
        ),
        direct(
            "ppt-scoped-correctness-validate",
            [
                "python3",
                str(ppt_correctness_validator),
                str(ppt_correctness_receipt),
                str(ppt_fixture),
            ],
            isolated=True,
        ),
    ]


def make_steps(root: Path, state_dir: Path, target: Path) -> list[dict]:
    all_features, format_features = feature_names(root)
    leaf = ["litchi-docx", "litchi-pptx", "litchi-xlsx", "litchi-xlsb", "litchi-doc", "litchi-xls", "litchi-ppt", "litchi-rtf", "litchi-odt", "litchi-odp", "litchi-opc", "xml-minifier", "litchi-drawingml", "litchi-spreadsheet-drawing", "litchi-keynote", "litchi-pages", "litchi-numbers", "litchi-numbers-wire", "litchi-iwa-text-wire"]
    leaf_args = [item for package in leaf for item in ("-p", package)]
    steps: list[dict] = [
        direct("fmt", ["cargo", "fmt", "--all", "--check"]),
        direct("clippy-all-features", ["cargo", "clippy", "-q", "-p", "litchi-cli", "-p", "litchi-ops", "-p", "litchi", "-p", "litchi-py", "--all-targets", "--all-features", "--", "-D", "warnings"]),
        direct("clippy-ops-no-encryption", ["cargo", "clippy", "-q", "-p", "litchi-ops", "--features", ",".join(all_features), "--all-targets", "--", "-D", "warnings"]),
    ]
    steps += [direct(f"clippy-ops-{feature}", ["cargo", "clippy", "-q", "-p", "litchi-ops", "--no-default-features", *(["--features", feature] if feature else []), "--all-targets", "--", "-D", "warnings"]) for feature in [""] + format_features]
    safe = [
        "python3",
        str(SAFE_CHUNKED),
        "--chunk-size",
        str(SAFE_CHUNK_SIZE),
        "--quarantine-dir",
        str(state_dir / "artifact-quarantine"),
    ]
    steps += [
        direct("tests-ops", [*safe, "-p", "litchi-ops", "--all-features", "--log-dir", str(state_dir / "ops-tests")]),
        direct("tests-facade", [*safe, "-p", "litchi", "--all-features", "--log-dir", str(state_dir / "facade-tests")]),
        direct("tests-core", [*safe, "-p", "litchi-core", "--all-features", "--log-dir", str(state_dir / "core-tests")]),
        direct("clippy-leaves", ["cargo", "clippy", "-q", *leaf_args, "--all-targets", "--all-features", "--", "-D", "warnings"]),
        direct("clippy-ods", ["cargo", "clippy", "-q", "-p", "litchi-ods", "--all-targets", "--all-features", "--", "-D", "warnings"]),
        direct("tests-leaves", [*safe, *leaf_args, "--all-features", "--log-dir", str(state_dir / "leaf-tests")]),
        direct("tests-ods-full", [*safe, "-p", "litchi-ods", "--all-features", "--log-dir", str(state_dir / "ods-tests")]),
        direct("clippy-workspace", ["cargo", "clippy", "-q", "--workspace", "--all-targets", "--all-features", "--", "-D", "warnings"]),
        direct("tests-workspace-excluding-ods-ops", [*safe, "--workspace", "--exclude", "litchi-ods", "--exclude", "litchi-ops", "--all-features", "--log-dir", str(state_dir / "workspace-tests")]),
        direct("wsdoc", ["cargo", "doc", "-q", "--workspace", "--no-deps", "--all-features"], isolated=True, env={"RUSTDOCFLAGS": "-Dwarnings"}),
        direct("ppt-alone", ["cargo", "test", "-q", "-p", "litchi-ppt", "--all-features", "--lib"]),
        direct("rustdoc-core-ops-facade", ["cargo", "doc", "-q", "--no-deps", "--all-features", "-p", "litchi-core", "-p", "litchi-ops", "-p", "litchi"], isolated=True, env={"RUSTDOCFLAGS": "-Dwarnings"}),
        direct("cli-tests", [*safe, "-p", "litchi-cli", "--all-features", "--log-dir", str(state_dir / "cli-tests")]),
        shell("python-bindings", "bash crates/pyo3-litchi/tests/run_tests.sh"),
        direct("feature-matrix", ["python3", "tools/check_feature_matrix.py"]),
        direct("feature-matrix-clippy", ["python3", "tools/check_feature_matrix.py", "--only", "facade,facade-aggregate", "--all-targets", "--clippy"]),
        shell("fuzz", "set -e; export CARGO_TARGET_DIR=\"$CARGO_TARGET_DIR/fuzz\"; cargo check -q --manifest-path crates/litchi-core/fuzz/Cargo.toml --bins; cargo check -q --manifest-path crates/litchi-rtf/fuzz/Cargo.toml --bins; cargo check -q --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bins; cargo check -q --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bins --features apply; cargo check -q --manifest-path crates/litchi/fuzz/Cargo.toml --bins; cargo run -q --release --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bin smoke_batch_payload; cargo run -q --release --manifest-path crates/litchi-ops/fuzz/Cargo.toml --features apply --bin smoke_batch_payload"),
        shell("goldens", f"bash {COORD / 'scripts/check-goldens.sh'}"),
        shell("golden-clean", "test -z \"$(git status --short | grep -v '^[AMDR] ' || true)\""),
        shell("perf-harness-check", "cd tools/perf-baseline && cargo check --locked --all-targets"),
        shell("native-resave-check", "cd tools/native-resave && cargo check --locked --all-targets"),
        direct("api-glossary", ["python3", "tools/check_api_glossary.py", "--target-dir", str(target)], env={"RUSTC_BOOTSTRAP": "1"}),
        direct("iwork-leaf-api", ["python3", "-m", "unittest", "tools.test_check_iwork_leaf_public_api"]),
        direct("tooling-unit", ["python3", "-m", "unittest", "tools.test_check_api_glossary", "tools.test_api_debt_inventory", "tools.test_validate_api_disposition", "tools.test_final_api_inventory", "tools.test_raw_api_coverage", "tools.test_prepare_raw_api_census", "tools.test_w8_file_provenance", "tools.test_w8_original_baseline", "tools.test_w8_projection_policy", "tools.test_w8_quiet_abba", "tools.test_officecli_corpus", "tools.test_officecli_capability_delta", "tools.test_no_3rdparty_references", "tools.test_check_crate_boundaries", "tools.test_non_iwork_gate", "tools.test_unified_bench_matrix", "tools.test_check_iwork_leaf_public_api", "tools.test_native_qualification_runner", "tools.test_legacy_native_qualification", "tools.test_prepare_xlsx_raw_inputs", "tools.test_xlsx_raw_native_qualification", "tools.test_prepare_ole_owner_inputs", "tools.test_ole_owner_paired_runner", "tools.test_validate_ole_owner_run", "tools.test_ppt_lazy_build_receipt_adapter", "tools.test_ppt_lazy_cold_first_commit", "tools.test_ppt_lazy_native_qualification", "tools.test_prepare_ppt_lazy_native_inputs", "tools.test_drawingml_compile_driver", "tools.test_drawingml_raw_qualification", "tools.test_drawingml_raw_policy_v2", "tools.test_signed_native_policy_v2", "tools.test_current_native_policy_v2", "docs/unified-ops/results/0362ul29/tooling/test_safe_chunked_tests.py", "docs/unified-ops/results/0362ul29/tooling/test_run_safe_gate_ppt.py", "tools.test_ink_safe_runner", "docs/unified-ops/results/0362ul24/test_prepare_signed_native_consumer.py", "tools.test_opc_read_ahead_timing_guard", "tools.test_xls_f3_qualification_abba", "docs/unified-ops/results/0362ul24/test_validate_signed_native.py"], target_guard=False),
        direct("crate-boundaries", ["python3", "tools/check_crate_boundaries.py"]),
        direct("non-iwork-gate", ["python3", "tools/non_iwork_gate.py", "verify"]),
        direct("bench-matrix", ["python3", "tools/unified_bench_matrix.py", "--check"]),
        shell("corpus-dispositions", "python3 tools/officecli_capability_delta.py --check && python3 tools/officecli_corpus.py --check"),
    ]
    return steps


def plan_steps(
    root: Path,
    state_dir: Path,
    target: Path,
    requested_auxiliary: set[str] | None = None,
) -> list[dict]:
    """Compose the normal plan and explicitly requested auxiliary steps."""
    steps = make_steps(root, state_dir, target)
    if requested_auxiliary and requested_auxiliary.intersection(PPT_AUXILIARY_NAMES):
        steps.extend(make_ppt_steps(root, state_dir, target))
    return steps


def attempt_paths(state_dir: Path, index: int, name: str, attempt: str) -> tuple[Path, Path]:
    stem = f"{index:02d}-{name}-attempt-{attempt}"
    return state_dir / f"{stem}.log", state_dir / f"{stem}-cargo-targets.tsv"


def new_attempt() -> str:
    return f"{time.time_ns()}-{os.getpid()}"


def capture_step_identities(process: subprocess.Popen, ownership: dict) -> None:
    """Poll while the top-level child lives so fast forks are not missed."""
    try:
        while True:
            _capture_owned(ownership, process_table())
            if process.poll() is not None:
                _capture_owned(ownership, process_table())
                return
            time.sleep(PRE_EXIT_POLL_SECONDS)
    except BaseException as error:
        ownership["monitor_error"] = repr(error)


def run_step(step: dict, root: Path, target: Path, state_dir: Path, source: dict, index: int, total: int, attempt: str = "manual") -> dict:
    pre = preflight(root, target, source, bool(step.get("isolated")))
    started = time.time()
    env = cargo_env(root, target, state_dir)
    target_guard = step.get("target_guard", True)
    if not target_guard:
        env.pop("CARGO_TARGET_DIR", None)
        env.pop("LITCHI_GATE_TARGET_ROOT", None)
    env.update(step.get("env", {}))
    env["LITCHI_GATE_ATTEMPT"] = attempt
    log, trace = attempt_paths(state_dir, index, step["name"], attempt)
    log.write_text("", encoding="utf-8")
    trace.write_text("", encoding="utf-8")
    env["LITCHI_GATE_TARGET_TRACE"] = str(trace)
    prelaunch = process_table()
    enable_child_subreaper()
    if step["kind"] == "argv":
        command = step["argv"]
        process = subprocess.Popen(command, cwd=root, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, start_new_session=True, pass_fds=tuple(step.get("pass_fds", ())), preexec_fn=None if target_guard else ensure_child_subreaper)
    else:
        command = ["bash", "-c", step["command"]]
        process = subprocess.Popen(command, cwd=root, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, start_new_session=True, pass_fds=tuple(step.get("pass_fds", ())), preexec_fn=None if target_guard else ensure_child_subreaper)
    ownership = new_ownership(process.pid, root, target, prelaunch, process_table())
    monitor = threading.Thread(target=capture_step_identities, args=(process, ownership), name=f"gate-owner-{process.pid}")
    monitor.start()
    live_stop = threading.Event()
    live_overflow: list[dict] = []
    live_errors: list[dict] = []
    live_guard_stats: dict = {}
    live_guard = threading.Thread(
        target=lambda: _run_target_guard(process, target, live_stop, live_overflow, live_errors, ownership, live_guard_stats),
        name=f"gate-target-guard-{process.pid}",
        daemon=True,
    )
    live_guard.start()
    with log.open("w", encoding="utf-8") as stream:
        assert process.stdout is not None
        for line in process.stdout:
            stream.write(line)
            stream.flush()
    status = process.wait()
    live_stop.set()
    live_guard.join()
    monitor.join()
    if ownership.get("monitor_error"):
        raise RuntimeError(f"process identity monitor failed: {ownership['monitor_error']}")
    post_exit = post_exit_guard(process.pid, root, target, ownership)
    size = post_exit["target_bytes"]
    after_source = source_state(root)
    actual_targets = []
    if trace.exists():
        for line in trace.read_text(encoding="utf-8").splitlines():
            fields = line.split("\t", 2)
            if len(fields) == 3:
                actual_targets.append({"pid": fields[0], "cargo_command": fields[1], "target": fields[2]})
    actual_target_paths = sorted({row["target"] for row in actual_targets if row["target"]})
    actual_target_sizes = {}
    for path_text in actual_target_paths:
        path = Path(path_text).resolve()
        if not path.is_relative_to(target.resolve()):
            status = 1
            actual_target_sizes[path_text] = {"bytes": None, "error": "target escaped gate root"}
        else:
            actual_target_sizes[path_text] = {"bytes": target_bytes(path)}
    row = {
        "index": index,
        "total": total,
        "name": step["name"],
        "attempt": attempt,
        "command": command,
        "environment": {key: env[key] for key in ("CARGO_TARGET_DIR", "CARGO_NET_OFFLINE", "LITCHI_MEM_MAX", "CARGO_BUILD_JOBS", "RUST_TEST_THREADS", "CARGO_INCREMENTAL", "RUSTDOCFLAGS", "LITCHI_GATE_ATTEMPT") if key in env},
        "exit": status,
        "seconds": time.time() - started,
        "log": str(log),
        "log_bytes": log.stat().st_size,
        "log_sha256": sha256(log),
        "target_bytes_after": size,
        "post_exit": post_exit,
        "source_after": after_source,
        "cargo_locks_before": source["cargo_locks"],
        "cargo_locks_after": after_source["cargo_locks"],
        "cargo_target_trace": actual_targets,
        "cargo_target_sizes": actual_target_sizes,
        "cargo_target_trace_file": str(trace),
        "cargo_target_trace_bytes": trace.stat().st_size,
        "cargo_target_trace_sha256": sha256(trace),
        "preflight": pre,
        "target_live_guard": live_guard_stats,
        "target_live_guard_errors": live_errors,
    }
    if live_errors:
        row["resource_failure"] = {
            "reason": "target live guard raised an unexpected exception",
            "errors": live_errors,
        }
        status = 1
    if live_overflow:
        row["resource_failure"] = {"reason": "target live guard stopped command before hard cap", **live_overflow[0]}
        if live_errors:
            row["resource_failure"]["guard_errors"] = live_errors
        status = 1
    if size > MAX_TARGET:
        row["resource_failure"] = f"target exceeded 20 GiB: {size}"
        status = 1
    free_after = shutil.disk_usage(target.parent).free
    row["free_bytes_after"] = free_after
    if free_after < MIN_FREE:
        row["resource_failure"] = f"free disk below 12 GB after command: {free_after}"
        status = 1
    if any(value.get("bytes", 0) > MAX_TARGET for value in actual_target_sizes.values() if value.get("bytes") is not None):
        row["resource_failure"] = "a nested Cargo target exceeded 20 GB"
        status = 1
    if after_source != source:
        row["provenance_failure"] = "source commit, tracked status, branch, or Cargo.lock hash changed"
        status = 1
    row["effective_exit"] = status
    if status:
        row["status"] = "failed"
    else:
        row["status"] = "passed"
    return row


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT_DEFAULT)
    parser.add_argument("--target", type=Path)
    parser.add_argument("--state", type=Path, required=True)
    parser.add_argument("--resume", action="store_true")
    parser.add_argument("--retry-failed", action="store_true")
    parser.add_argument("--only", help="comma-separated step names to run (partial gate)")
    parser.add_argument("--start-at", help="start at this step name and continue (partial gate)")
    args = parser.parse_args()
    root = args.root.resolve()
    target = (args.target or root / "target").resolve()
    if target.name != "target" or not target.parent.is_dir():
        raise SystemExit(f"target must be an existing worktree target directory, got {target}")
    state_dir = args.state.resolve()
    state_dir.mkdir(parents=True, exist_ok=True)
    source = source_state(root)
    if source["status"]:
        raise SystemExit(f"source worktree must be clean: {source['status']!r}")
    only = None if args.only is None else [name for name in args.only.split(",") if name]
    requested_auxiliary = set(only or ())
    if args.start_at is not None:
        requested_auxiliary.add(args.start_at)
    all_steps = plan_steps(root, state_dir, target, requested_auxiliary)
    by_name = {step["name"]: index for index, step in enumerate(all_steps)}
    if len(by_name) != len(all_steps):
        raise SystemExit("gate plan contains duplicate step names")
    if only is not None:
        unknown = sorted(set(only) - set(by_name))
        if unknown:
            raise SystemExit(f"unknown --only step(s): {', '.join(unknown)}")
        selected_indexes = [by_name[name] for name in only]
    else:
        selected_indexes = list(range(len(all_steps)))
    if args.start_at is not None:
        if args.start_at not in by_name:
            raise SystemExit(f"unknown --start-at step: {args.start_at}")
        start_index = by_name[args.start_at]
        selected_indexes = [index for index in selected_indexes if index >= start_index]
    if not selected_indexes:
        raise SystemExit("selection contains no steps")
    steps = [(index, all_steps[index]) for index in selected_indexes]
    selection = {
        "only": only,
        "start_at": args.start_at,
        "indexes": selected_indexes,
        "full_step_count": len(all_steps),
        "partial": len(selected_indexes) != len(all_steps),
    }
    runner_files = {
        str(path): sha256(path)
        for path in (Path(__file__).resolve(), SAFE_CHUNKED.resolve(), Path(__file__).resolve().with_name("bin") / "cargo")
    }
    policy = {
        "runner_files": runner_files,
        "cargo_wrapper": str(Path(__file__).resolve().with_name("bin") / "cargo"),
        "offline": True,
        "locked": True,
        "memory_max": "20G",
        "build_jobs": 1,
        "test_threads": 1,
        "target_guard_before": START_TARGET,
        "target_live_guard": LIVE_TARGET,
        "target_guard_after": MAX_TARGET,
        "free_guard": MIN_FREE,
        "safe_chunk_size": SAFE_CHUNK_SIZE,
        "safe_quarantine_root": str(state_dir / "artifact-quarantine"),
    }
    plan_path = state_dir / "plan.json"
    state_path = state_dir / "state.json"
    if state_path.exists():
        if not args.resume:
            raise SystemExit(f"state exists; use --resume: {state_path}")
        state = json.loads(state_path.read_text(encoding="utf-8"))
        if state.get("source") != source or state.get("root") != str(root) or state.get("target") != str(target) or state.get("selection") != selection or state.get("policy") != policy:
            raise SystemExit("resume source/root/target does not match recorded run")
        rows = state.get("rows", [])
    else:
        rows = []
        plan_path.write_text(json.dumps({"source": source, "root": str(root), "target": str(target), "policy": policy, "selection": selection, "steps": all_steps}, indent=2) + "\n", encoding="utf-8")
        state = {"source": source, "root": str(root), "target": str(target), "policy": policy, "selection": selection, "started": time.time(), "rows": rows}
        state_path.write_text(json.dumps(state, indent=2) + "\n", encoding="utf-8")
    completed = {row["index"]: row for row in rows if row.get("status") == "passed"}
    failed = {row["index"] for row in rows if row.get("status") == "failed"}
    for index, step in steps:
        if index in completed:
            continue
        if index in failed and not args.retry_failed:
            raise SystemExit(f"step {index} previously failed; use --retry-failed")
        attempt = new_attempt()
        try:
            row = run_step(step, root, target, state_dir, source, index, len(all_steps), attempt)
        except Exception as error:  # retain a resumable resource/provenance failure
            log, _ = attempt_paths(state_dir, index, step["name"], attempt)
            # Preserve any child output already written before a resource or
            # provenance exception. Store the runner failure separately.
            if not log.exists():
                log.write_text("", encoding="utf-8")
            failure_log = log.with_suffix(".blocked.log")
            failure_log.write_text(
                f"gate step blocked before or during launch: {error!r}\n",
                encoding="utf-8",
            )
            row = {
                "index": index,
                "total": len(all_steps),
                "name": step["name"],
                "attempt": attempt,
                "status": "blocked",
                "error": repr(error),
                "started": time.time(),
                "ended": time.time(),
                "log": str(log),
                "log_bytes": log.stat().st_size,
                "log_sha256": sha256(log),
                "failure_log": str(failure_log),
                "failure_log_bytes": failure_log.stat().st_size,
                "failure_log_sha256": sha256(failure_log),
            }
        rows = [item for item in rows if item.get("index") != index] + [row]
        rows.sort(key=lambda item: item["index"])
        state.update({"rows": rows, "updated": time.time()})
        state_path.write_text(json.dumps(state, indent=2) + "\n", encoding="utf-8")
        print(json.dumps({key: value for key, value in row.items() if key not in {"preflight"}}), flush=True)
        if row["status"] != "passed":
            return 1
    final_source = source_state(root)
    if final_source != source:
        raise SystemExit(f"source changed after final step: {final_source}")
    state.update({"finished": time.time(), "status": "partial-passed" if selection["partial"] else "passed"})
    state_path.write_text(json.dumps(state, indent=2) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
