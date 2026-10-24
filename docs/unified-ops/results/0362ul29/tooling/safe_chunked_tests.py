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

MAX_TARGET = 20_000_000_000
LIVE_TARGET = 19_000_000_000
START_TARGET = 18_000_000_000
MIN_FREE = 12_000_000_000
TARGET_POLL_SECONDS = 0.1
PROCESS_KILL_GRACE_SECONDS = 1.0
DESCENDANT_WAIT_SECONDS = 2.0


class TargetOverflow(RuntimeError):
    """The live target guard stopped a command before the hard cap."""

    def __init__(self, detail: dict):
        self.detail = detail
        super().__init__(f"target live guard stopped command: {json.dumps(detail, sort_keys=True)}")


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
    return int(subprocess.check_output(["du", "-sb", str(target)], text=True).split()[0])


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
    members: list[int] = []
    for entry in Path("/proc").glob("[0-9]*"):
        try:
            pid = int(entry.name)
            text = (entry / "stat").read_text()
            close = text.rfind(")")
            fields = text[close + 2 :].split()
            if int(fields[2]) == pgrp:
                members.append(pid)
        except (FileNotFoundError, ProcessLookupError, ValueError, IndexError):
            continue
    return members


def _wait_for_descendants(process: subprocess.Popen) -> list[int]:
    """Do not retire artifacts while a child of the completed Cargo group lives."""
    deadline = time.monotonic() + DESCENDANT_WAIT_SECONDS
    while time.monotonic() < deadline:
        members = _process_group_members(process.pid)
        if not members:
            return []
        time.sleep(0.02)
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return []
    deadline = time.monotonic() + PROCESS_KILL_GRACE_SECONDS
    while time.monotonic() < deadline:
        members = _process_group_members(process.pid)
        if not members:
            return []
        time.sleep(0.02)
    try:
        os.killpg(process.pid, signal.SIGKILL)
    except ProcessLookupError:
        return []
    time.sleep(0.05)
    return _process_group_members(process.pid)


def _watch_target(
    process: subprocess.Popen,
    target: Path,
    stop: threading.Event,
    overflow: list[dict],
) -> dict:
    """Sample target growth and stop a command before the hard cap."""
    maximum = target_bytes(target)
    samples = 0
    while not stop.wait(TARGET_POLL_SECONDS):
        if process.poll() is not None:
            break
        size = target_bytes(target)
        maximum = max(maximum, size)
        samples += 1
        if size >= LIVE_TARGET:
            detail = {
                "target_bytes": size,
                "live_target": LIVE_TARGET,
                "max_sampled_bytes": maximum,
                "samples": samples,
            }
            overflow.append(detail)
            _terminate_process_group(process)
            break
    return {"max_sampled_bytes": maximum, "samples": samples, "live_target": LIVE_TARGET}


def _append_manifest(path: Path, row: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(row, sort_keys=True) + "\n")
        stream.flush()
        os.fsync(stream.fileno())


def _retire_artifacts(
    artifacts: set[Path],
    target: Path,
    quarantine: Path,
    manifest: Path,
    command: list[str],
    chunk_index: int,
    *,
    scan_private: bool = True,
) -> list[dict]:
    """Move completed command binaries and their private sidecars, retaining evidence."""
    debug = (target / "debug").resolve()
    destination_root = quarantine / f"{chunk_index:04d}"
    reported = {Path(path).resolve() for path in artifacts}
    if scan_private:
        deps = debug / "deps"
        examples = debug / "examples"
        if deps.is_dir():
            for path in deps.iterdir():
                if path.is_file() and not path.is_symlink() and os.access(path, os.X_OK) and path.suffix not in {".so", ".dylib", ".dll"}:
                    artifacts.add(path)
                    sidecar = path.with_name(path.name + ".d")
                    if sidecar.is_file():
                        artifacts.add(sidecar)
        if examples.is_dir():
            for path in examples.iterdir():
                if path.is_file() and not path.is_symlink() and (os.access(path, os.X_OK) or path.suffix in {".d", ".rmeta"}):
                    artifacts.add(path)
    retired: list[dict] = []
    for raw_path in sorted(artifacts):
        path = Path(raw_path)
        if path.is_symlink():
            continue
        path = path.resolve()
        if not path.is_file() or path.suffix in {".so", ".dylib", ".dll"}:
            continue
        if not path.is_relative_to(debug) or not (os.access(path, os.X_OK) or path.suffix in {".d", ".rmeta"}):
            continue
        source_stat = path.stat()
        source_digest = digest(path)
        relative = path.relative_to(target.resolve())
        destination = (destination_root / relative).resolve()
        if destination == quarantine or not destination.is_relative_to(quarantine.resolve()):
            raise RuntimeError(f"quarantine destination escaped quarantine root: {destination}")
        if destination.exists():
            raise RuntimeError(f"quarantine destination already exists: {destination}")
        destination.parent.mkdir(parents=True, exist_ok=True)
        # The gate places quarantine beside target on the same filesystem, so
        # this is atomic and preserves the original inode/mode without a
        # temporary duplicate inside target.
        os.replace(path, destination)
        destination_stat = destination.stat()
        destination_digest = digest(destination)
        mode = stat.S_IMODE(source_stat.st_mode)
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
    guard_stop = threading.Event()
    overflow: list[dict] = []
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
            target=lambda: guard_stats.update(_watch_target(process, target, guard_stop, overflow)),
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
        descendants = _wait_for_descendants(process)
        if descendants:
            detail = {"process_group": process.pid, "remaining_descendants": descendants, "command_exit": status}
            stream.write(json.dumps({"safe_chunked_descendant_guard": detail}, sort_keys=True) + "\n")
            stream.flush()
            raise RuntimeError(f"Cargo descendants remained after command: {descendants}")
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
        removed = _retire_artifacts(artifacts, target, quarantine, manifest, command, chunk_index)
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
            status, removed, measurement = execute(command, target, log, quarantine, index)
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
            "quarantine_manifest": str(manifest),
            **measurement,
        }
        rows.append(row)
        print(json.dumps(row), flush=True)
        if status:
            failures.append(row)
    (log_dir / "checks.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
