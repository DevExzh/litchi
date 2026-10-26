#!/usr/bin/env python3
"""Bounded replacement for the historical chunked Cargo test helper.

The coordinator gate's helper invokes target-headroom.py, which deletes the
whole debug tree above 16 GiB.  This copy keeps the useful package/target
chunking and moves executable artifacts explicitly reported by a completed
Cargo command to an external, hashed quarantine.  It never removes a cache
or a release artifact, and a live target guard stops growth before the hard
20 GB cap.
"""
from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
from pathlib import Path
import shutil
import signal
import stat
import subprocess
import sys
import threading
import time
import traceback

MAX_TARGET = 20_000_000_000
LIVE_TARGET = 19_000_000_000
START_TARGET = 18_000_000_000
MIN_FREE = 12_000_000_000
TARGET_POLL_SECONDS = 0.1
TARGET_MEASURE_RETRIES = 5
TARGET_MEASURE_DELAY_SECONDS = 0.02
PROCESS_KILL_GRACE_SECONDS = 1.0
DESCENDANT_WAIT_SECONDS = 2.0
PR_SET_CHILD_SUBREAPER = 36
_SUBREAPER_ENABLED = False


class TargetOverflow(RuntimeError):
    """The live target guard stopped a command before the hard cap."""

    def __init__(self, detail: dict):
        self.detail = detail
        super().__init__(f"target live guard stopped command: {json.dumps(detail, sort_keys=True)}")


def enable_child_subreaper() -> None:
    """Keep descendants that call ``setsid`` observable by this helper."""
    global _SUBREAPER_ENABLED
    if _SUBREAPER_ENABLED:
        return
    if sys.platform != "linux":
        raise RuntimeError("safe chunked tests require Linux child-subreaper support")
    try:
        libc = ctypes.CDLL(None, use_errno=True)
        result = libc.prctl(PR_SET_CHILD_SUBREAPER, 1, 0, 0, 0)
        if result != 0:
            error = ctypes.get_errno()
            raise OSError(error, os.strerror(error))
    except (OSError, AttributeError) as error:
        raise RuntimeError(f"cannot enable child subreaper: {error}") from error
    _SUBREAPER_ENABLED = True


def plan(metadata: dict, selected: set[str], features: list[str], chunk_size: int):
    commands: list[tuple[str, list[str]]] = []
    for package in metadata["packages"]:
        if package["name"] not in selected:
            continue
        base = ["-p", package["name"], *features]
        targets = package["targets"]
        libraries = [t for t in targets if set(t["kind"]) & {"lib", "rlib", "dylib", "cdylib", "staticlib", "proc-macro"}]
        if any(t["test"] for t in libraries):
            commands.append(("test", [*base, "--lib"]))
        for kind, switch in (("bin", "--bin"), ("test", "--test"), ("example", "--example")):
            runnable = [t for t in targets if kind in t["kind"] and t["test"]]
            for start in range(0, len(runnable), chunk_size):
                selectors = [item for t in runnable[start:start + chunk_size] for item in (switch, t["name"])]
                commands.append(("test", [*base, *selectors]))
        examples = [t for t in targets if "example" in t["kind"] and not t["test"]]
        for start in range(0, len(examples), chunk_size):
            selectors = [item for t in examples[start:start + chunk_size] for item in ("--example", t["name"])]
            commands.append(("build", [*base, "--profile", "test", *selectors]))
        if any(t["doctest"] for t in libraries):
            commands.append(("test", [*base, "--doc"]))
    return commands


def digest(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


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


def _terminate_process_group(process: subprocess.Popen) -> None:
    """Stop only the isolated Cargo command and all of its children."""
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return
    deadline = time.monotonic() + PROCESS_KILL_GRACE_SECONDS
    while process.poll() is None and time.monotonic() < deadline:
        time.sleep(0.02)
    if process.poll() is None:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass


def _process_group_members(pgrp: int) -> list[int]:
    return [pid for pid, row in _process_table().items() if row["pgrp"] == pgrp]


def _process_table() -> dict[int, dict[str, int]]:
    """Read process ancestry without requiring ``/proc/$pid/exe`` to exist.

    A just-exited adopted child is a zombie until this helper reaps it, and
    Linux removes its ``exe`` link while it is a zombie.  The stat fields are
    sufficient to distinguish the process identity and its ancestry.
    """
    result: dict[int, dict[str, int]] = {}
    for entry in Path("/proc").glob("[0-9]*"):
        try:
            pid = int(entry.name)
            text = (entry / "stat").read_text()
            close = text.rfind(")")
            fields = text[close + 2 :].split()
            result[pid] = {
                "pid": pid,
                "ppid": int(fields[1]),
                "pgrp": int(fields[2]),
                "start_time": int(fields[19]),
            }
        except (FileNotFoundError, ProcessLookupError, PermissionError, ValueError, IndexError):
            continue
    return result


def _owned_descendants(
    process: subprocess.Popen,
    baseline: set[tuple[int, int]],
) -> tuple[dict[int, dict[str, int]], set[int]]:
    """Find new group members and adopted descendants of one Cargo command."""
    table = _process_table()
    runner_pid = os.getpid()
    new = {
        pid: row
        for pid, row in table.items()
        if pid != runner_pid and (pid, row["start_time"]) not in baseline
    }
    roots = {
        pid
        for pid, row in new.items()
        if row["ppid"] == process.pid or row["ppid"] == runner_pid
    }
    owned = set(roots)
    changed = True
    while changed:
        changed = False
        for pid, row in new.items():
            if pid not in owned and row["ppid"] in owned:
                owned.add(pid)
                changed = True
    # A process started with start_new_session=True has pgrp == pid.  Keep
    # the group check as well because it catches a fast fork before ancestry
    # has been reparented, while the roots above catch a child that called
    # setsid() and was adopted by this subreaper.
    owned.update(pid for pid, row in new.items() if row["pgrp"] == process.pid)
    return table, owned


def _reap_owned(table: dict[int, dict[str, int]], owned: set[int]) -> None:
    for pid in sorted(owned):
        if pid == os.getpid():
            continue
        try:
            os.waitpid(pid, os.WNOHANG)
        except (ChildProcessError, ProcessLookupError):
            continue


def _signal_owned(process: subprocess.Popen, pids: set[int], signal_number: int) -> None:
    try:
        os.killpg(process.pid, signal_number)
    except ProcessLookupError:
        pass
    for pid in sorted(pids):
        if pid == os.getpid():
            continue
        try:
            os.kill(pid, signal_number)
        except ProcessLookupError:
            continue


def _wait_for_descendants(process: subprocess.Popen, baseline: set[tuple[int, int]]) -> tuple[list[int], bool]:
    """Do not retire artifacts while a group or adopted descendant lives.

    The second return value records that the timeout path had to signal a
    descendant.  Even when SIGTERM succeeds, that is not a successful Cargo
    lifetime: callers must retain artifacts and produce a blocked receipt.
    """
    deadline = time.monotonic() + DESCENDANT_WAIT_SECONDS
    while time.monotonic() < deadline:
        table, owned = _owned_descendants(process, baseline)
        _reap_owned(table, owned)
        table, owned = _owned_descendants(process, baseline)
        if not owned:
            return [], False
        time.sleep(0.02)

    table, owned = _owned_descendants(process, baseline)
    _signal_owned(process, owned, signal.SIGTERM)
    deadline = time.monotonic() + PROCESS_KILL_GRACE_SECONDS
    while time.monotonic() < deadline:
        table, owned = _owned_descendants(process, baseline)
        _reap_owned(table, owned)
        table, owned = _owned_descendants(process, baseline)
        if not owned:
            return [], True
        time.sleep(0.02)
    table, owned = _owned_descendants(process, baseline)
    _signal_owned(process, owned, signal.SIGKILL)
    time.sleep(0.05)
    table, owned = _owned_descendants(process, baseline)
    _reap_owned(table, owned)
    _, owned = _owned_descendants(process, baseline)
    return sorted(owned), True


def _watch_target(
    process: subprocess.Popen,
    target: Path,
    stop: threading.Event,
    overflow: list[dict],
    baseline: set[tuple[int, int]] | None = None,
) -> dict:
    """Sample target growth and stop a command before the hard cap."""
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
        if baseline is not None:
            _, owned = _owned_descendants(process, baseline)
            _signal_owned(process, owned, signal.SIGTERM)
        _terminate_process_group(process)
        if baseline is not None:
            remaining, forced = _wait_for_descendants(process, baseline)
            detail["cleanup"] = {"remaining_descendants": remaining, "forced_cleanup": forced}
        return {"max_sampled_bytes": None, "samples": 0, "live_target": LIVE_TARGET, "measurement_error": repr(error)}
    samples = 0
    while not stop.wait(TARGET_POLL_SECONDS):
        if process.poll() is not None:
            break
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
            if baseline is not None:
                _, owned = _owned_descendants(process, baseline)
                _signal_owned(process, owned, signal.SIGTERM)
            _terminate_process_group(process)
            if baseline is not None:
                remaining, forced = _wait_for_descendants(process, baseline)
                detail["cleanup"] = {"remaining_descendants": remaining, "forced_cleanup": forced}
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
            if baseline is not None:
                _, owned = _owned_descendants(process, baseline)
                _signal_owned(process, owned, signal.SIGTERM)
            _terminate_process_group(process)
            if baseline is not None:
                remaining, forced = _wait_for_descendants(process, baseline)
                detail["cleanup"] = {"remaining_descendants": remaining, "forced_cleanup": forced}
            break
    return {"max_sampled_bytes": maximum, "samples": samples, "live_target": LIVE_TARGET}


def _run_target_guard(
    process: subprocess.Popen,
    target: Path,
    stop: threading.Event,
    overflow: list[dict],
    errors: list[dict],
    baseline: set[tuple[int, int]],
    stats: dict,
) -> None:
    """Convert any watcher exception into a recorded, fail-closed cleanup."""
    try:
        stats.update(_watch_target(process, target, stop, overflow, baseline))
    except BaseException as error:
        detail = {
            "reason": "target guard exception",
            "error": repr(error),
            "traceback": traceback.format_exc(),
        }
        try:
            _, owned = _owned_descendants(process, baseline)
            _signal_owned(process, owned, signal.SIGTERM)
            _terminate_process_group(process)
            remaining, forced = _wait_for_descendants(process, baseline)
            detail["cleanup"] = {"remaining_descendants": remaining, "forced_cleanup": forced}
        except BaseException as cleanup_error:
            detail["cleanup_error"] = repr(cleanup_error)
        errors.append(detail)


def _append_manifest(path: Path, row: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(row, sort_keys=True) + "\n")
        stream.flush()
        os.fsync(stream.fileno())


def _is_elf(path: Path) -> bool:
    try:
        with path.open("rb") as stream:
            return stream.read(4) == b"\x7fELF"
    except (FileNotFoundError, OSError):
        return False


def _private_elf(path: Path, debug: Path) -> bool:
    if (
        path.is_symlink()
        or not path.is_file()
        or path.suffix in {".so", ".dylib", ".dll"}
        or not os.access(path, os.X_OK)
        or not _is_elf(path)
    ):
        return False
    try:
        relative = path.resolve().relative_to(debug)
    except ValueError:
        return False
    return len(relative.parts) == 2 and relative.parts[0] in {"deps", "examples"}


def _sidecars(path: Path) -> set[Path]:
    """Return only Cargo sidecars whose stem exactly matches one ELF."""
    candidates = {
        path.with_name(path.name + ".d"),
        path.with_name(path.name + ".rmeta"),
        path.with_suffix(".d"),
        path.with_suffix(".rmeta"),
    }
    return {
        candidate
        for candidate in candidates
        if candidate.is_file()
        and not candidate.is_symlink()
        and candidate.suffix in {".d", ".rmeta"}
    }


def _quarantine_namespace(scope: str) -> str:
    """Map an invocation scope to a path-safe, stable quarantine namespace."""
    if not isinstance(scope, str) or not scope:
        raise RuntimeError("quarantine scope must be a non-empty string")
    return hashlib.sha256(scope.encode("utf-8")).hexdigest()[:24]


def _retire_artifacts(
    artifacts: set[Path],
    target: Path,
    quarantine: Path,
    manifest: Path,
    command: list[str],
    chunk_index: int,
    *,
    scan_private: bool = True,
    quarantine_scope: str = "direct",
) -> list[dict]:
    """Move only private Cargo ELF outputs and their exact sidecars."""
    debug = (target / "debug").resolve()
    namespace = _quarantine_namespace(quarantine_scope)
    destination_root = quarantine / namespace / f"{chunk_index:04d}"
    reported = {Path(path).resolve() for path in artifacts}
    selected_elfs: set[Path] = set()
    for raw_path in artifacts:
        path = Path(raw_path)
        if _private_elf(path, debug):
            selected_elfs.add(path)
    if scan_private:
        deps = debug / "deps"
        examples = debug / "examples"
        for private_dir in (deps, examples):
            if private_dir.is_dir():
                selected_elfs.update(
                    path for path in private_dir.iterdir() if _private_elf(path, debug)
                )
    selected = set(selected_elfs)
    for path in selected_elfs:
        selected.update(_sidecars(path))
    retired: list[dict] = []
    for raw_path in sorted(selected):
        path = Path(raw_path)
        if path.is_symlink():
            continue
        path = path.resolve()
        if not path.is_file():
            continue
        if path.suffix in {".d", ".rmeta"}:
            if not any(path in _sidecars(elf) for elf in selected_elfs):
                continue
        elif not _private_elf(path, debug):
            continue
        source_stat = path.stat()
        source_digest = digest(path)
        relative = path.relative_to(target.resolve())
        destination = (destination_root / relative).resolve()
        if destination == quarantine or not destination.is_relative_to(quarantine.resolve()):
            raise RuntimeError(f"quarantine destination escaped quarantine root: {destination}")
        mode = stat.S_IMODE(source_stat.st_mode)
        reused = False
        if destination.exists() or destination.is_symlink():
            if destination.is_symlink() or not destination.is_file():
                raise RuntimeError(f"quarantine destination already exists and is not a regular file: {destination}")
            destination_stat = destination.stat()
            destination_digest = digest(destination)
            if (
                destination_digest != source_digest
                or destination_stat.st_size != source_stat.st_size
                or stat.S_IMODE(destination_stat.st_mode) != mode
            ):
                raise RuntimeError(f"quarantine destination differs from source artifact: {destination}")
            # A previous invocation may have completed the move before its
            # manifest append.  Preserve that verified artifact and remove
            # only the identical source, recording this reuse explicitly.
            path.unlink()
            reused = True
        else:
            destination.parent.mkdir(parents=True, exist_ok=True)
            # The gate places quarantine beside target on the same filesystem,
            # so this is atomic and preserves the original inode/mode without
            # a temporary duplicate inside target.
            os.replace(path, destination)
            destination_stat = destination.stat()
            destination_digest = digest(destination)
            if (
                destination_digest != source_digest
                or destination_stat.st_size != source_stat.st_size
                or stat.S_IMODE(destination_stat.st_mode) != mode
            ):
                raise RuntimeError(f"quarantined artifact changed: {path}")
        row = {
            "source": str(path),
            "relative_source": str(relative),
            "quarantine": str(destination),
            "bytes": source_stat.st_size,
            "sha256": source_digest,
            "mode": f"{mode:04o}",
            "chunk_index": chunk_index,
            "quarantine_scope": quarantine_scope,
            "quarantine_namespace": namespace,
            "reused": reused,
            "reason": "cargo-reported" if path in reported else "completed-target-private",
            "command": command,
        }
        _append_manifest(manifest, row)
        retired.append(row)
    return retired


def execute(
    command: list[str],
    target: Path,
    log: Path,
    quarantine: Path,
    chunk_index: int,
    *,
    quarantine_scope: str = "direct",
) -> tuple[int, list[dict], dict]:
    started = time.time()
    size = target_bytes(target)
    free = shutil.disk_usage(target.parent).free
    if size > START_TARGET:
        raise RuntimeError(f"target requires headroom before command: {size} bytes > {START_TARGET}")
    if free < MIN_FREE:
        raise RuntimeError(f"free disk below 12 GiB: {free}")
    manifest = quarantine / "manifest.jsonl"
    command += ["--message-format", "json-render-diagnostics"]
    artifacts: set[Path] = set()
    enable_child_subreaper()
    baseline = {
        (pid, row["start_time"])
        for pid, row in _process_table().items()
    }
    guard_stop = threading.Event()
    overflow: list[dict] = []
    guard_errors: list[dict] = []
    guard_stats: dict = {}
    with log.open("w", encoding="utf-8") as stream:
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            start_new_session=True,
        )
        guard = threading.Thread(
            target=lambda: _run_target_guard(process, target, guard_stop, overflow, guard_errors, baseline, guard_stats),
            name=f"target-guard-{process.pid}",
            daemon=True,
        )
        guard.start()
        assert process.stdout is not None
        for line in process.stdout:
            stream.write(line)
            stream.flush()
            try:
                message = json.loads(line)
            except (TypeError, ValueError):
                continue
            if not isinstance(message, dict) or message.get("reason") != "compiler-artifact":
                continue
            target_kinds = message.get("target", {}).get("kind", [])
            if "custom-build" in target_kinds:
                continue
            executable = message.get("executable")
            if executable:
                artifacts.add(Path(executable))
            if executable or any(kind in {"bin", "example", "test"} for kind in target_kinds):
                for filename in message.get("filenames", []):
                    path = Path(filename)
                    if path.is_file() and (os.access(path, os.X_OK) or path.suffix in {".d", ".rmeta"}):
                        artifacts.add(path)
        status = process.wait()
        process.stdout.close()
        guard_stop.set()
        guard.join()
        descendants, forced_cleanup = _wait_for_descendants(process, baseline)
        if guard_errors:
            stream.write(json.dumps({"safe_chunked_guard_failure": guard_errors[0]}, sort_keys=True) + "\n")
            stream.flush()
        if descendants or forced_cleanup:
            detail = {
                "process_group": process.pid,
                "remaining_descendants": descendants,
                "forced_cleanup": forced_cleanup,
                "command_exit": status,
            }
            stream.write(json.dumps({"safe_chunked_descendant_guard": detail}, sort_keys=True) + "\n")
            stream.flush()
            raise RuntimeError(f"Cargo descendant lifetime required forced cleanup: {detail}")
        if guard_errors:
            raise RuntimeError(f"Cargo target guard failed: {guard_errors[0]}")
        if overflow:
            detail = {**overflow[0], "command_exit": status}
            stream.write(json.dumps({"safe_chunked_live_guard": detail}, sort_keys=True) + "\n")
            stream.flush()
            raise TargetOverflow(detail)

    removed: list[dict] = []
    # Only successful completed commands may release the test executables they
    # reported.  The target/debug boundary prevents a release binary or a
    # cache from being moved accidentally. This happens after process.wait(),
    # so no test/example caller can still need the executable.
    if status == 0:
        removed = _retire_artifacts(
            artifacts,
            target,
            quarantine,
            manifest,
            command,
            chunk_index,
            quarantine_scope=quarantine_scope,
        )
    after = target_bytes(target)
    if after > MAX_TARGET:
        raise RuntimeError(f"target exceeded 20 GB after retention: {after}")
    ended = time.time()
    return status, removed, {
        "started": started,
        "ended": ended,
        "seconds": ended - started,
        "target_bytes_before": size,
        "target_bytes_after": after,
        "free_bytes_before": free,
        "free_bytes_after": shutil.disk_usage(target.parent).free,
        "live_guard": guard_stats,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("-p", "--package", action="append", default=[])
    parser.add_argument("--workspace", action="store_true")
    parser.add_argument("--exclude", action="append", default=[])
    parser.add_argument("--all-features", action="store_true")
    parser.add_argument("--features")
    parser.add_argument("--no-default-features", action="store_true")
    parser.add_argument("--chunk-size", type=int, default=2)
    parser.add_argument("--plan", action="store_true")
    parser.add_argument("--log-dir", type=Path)
    parser.add_argument("--quarantine-dir", type=Path)
    args = parser.parse_args()
    if args.chunk_size < 1:
        parser.error("--chunk-size must be positive")
    metadata = json.loads(subprocess.check_output(["cargo", "metadata", "--format-version", "1", "--no-deps"], text=True))
    workspace = set(metadata["workspace_members"])
    names = {p["name"] for p in metadata["packages"] if p["id"] in workspace}
    selected = (names if args.workspace else set(args.package)) - set(args.exclude)
    if not selected or selected - names:
        parser.error("select existing workspace packages")
    features: list[str] = []
    if args.all_features:
        features.append("--all-features")
    if args.no_default_features:
        features.append("--no-default-features")
    if args.features:
        features += ["--features", args.features]
    commands = plan(metadata, selected, features, args.chunk_size)
    if args.plan:
        for verb, arguments in commands:
            print(" ".join(["cargo", verb, *arguments, "--no-fail-fast" if verb == "test" else ""]))
        return 0
    target = Path(os.environ.get("CARGO_TARGET_DIR", metadata["target_directory"])).resolve()
    log_dir = (args.log_dir or Path("safe-chunked-logs")).resolve()
    log_dir.mkdir(parents=True, exist_ok=True)
    quarantine = (args.quarantine_dir or target.parent / "artifact-quarantine").resolve()
    if quarantine == target or quarantine.is_relative_to(target):
        parser.error("--quarantine-dir must be outside CARGO_TARGET_DIR")
    quarantine.mkdir(parents=True, exist_ok=True)
    attempt = os.environ.get("LITCHI_GATE_ATTEMPT", "")
    quarantine_scope = f"log-dir:{log_dir}\nattempt:{attempt}"
    quarantine_namespace = _quarantine_namespace(quarantine_scope)
    manifest = quarantine / "manifest.jsonl"
    failures = []
    rows = []
    for index, (verb, arguments) in enumerate(commands):
        command = ["cargo", verb, *arguments]
        if verb == "test":
            command.append("--no-fail-fast")
        log = log_dir / f"{index:04d}.log"
        started = time.time()
        try:
            status, removed, measurement = execute(
                command,
                target,
                log,
                quarantine,
                index,
                quarantine_scope=quarantine_scope,
            )
        except Exception as error:
            if not log.exists():
                log.write_text(f"safe chunked gate blocked: {error!r}\n", encoding="utf-8")
            row = {
                "index": index,
                "command": command,
                "status": "blocked",
                "error": repr(error),
                "started": started,
                "ended": time.time(),
                "seconds": time.time() - started,
                "log": str(log),
                "log_bytes": log.stat().st_size,
                "log_sha256": digest(log),
                "quarantine_scope": quarantine_scope,
                "quarantine_namespace": quarantine_namespace,
            }
            rows.append(row)
            print(json.dumps(row), flush=True)
            failures.append(row)
            break
        row = {
            "index": index,
            "command": command,
            "exit": status,
            "status": "passed" if status == 0 else "failed",
            "log": str(log),
            "log_bytes": log.stat().st_size,
            "log_sha256": digest(log),
            "pruned": removed,
            "quarantine_scope": quarantine_scope,
            "quarantine_namespace": quarantine_namespace,
            "quarantine_manifest": str(manifest),
            **measurement,
        }
        rows.append(row)
        print(json.dumps(row), flush=True)
        if status:
            failures.append(row)
            # A failed chunk is a boundary: its outputs may still be needed
            # for diagnosis or a controlled retry.  Never let a later
            # successful private scan retire those artifacts as if this
            # command had completed.
            break
    (log_dir / "checks.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
