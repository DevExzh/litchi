#!/usr/bin/env python3
"""Build the 0948 candidate native binary with durable custody receipts.

This runner owns one build lane only.  It validates the frozen source export,
binds every build input before starting Cargo, starts Cargo in a fresh process
group, and writes terminal evidence after the child has been reaped and the
group has been cleared.  Receipts are create-only: a rerun in a packet that
already contains any receipt or log is refused.
"""

from __future__ import annotations

import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import signal
import subprocess
import sys
import time
from typing import Any


PACKET = Path(__file__).resolve().parents[1]
SOURCE = PACKET / "candidate-build-v1"
SOURCE_MANIFEST = PACKET / "candidate-build-manifest-v1.json"
FREEZE_RECEIPT = PACKET / "receipts" / "candidate-freeze-v1.json"
RUNNER = Path(__file__).resolve()
BUILD = PACKET / "builds" / "candidate-native-v1"
TARGET = PACKET / "target-candidate-native-v1"
TMP = PACKET / "tmp" / "candidate-native-v1"

DEADLINE_SECONDS = 1_800

HOME = Path.home()
CARGO_HOME = HOME / ".cargo"
CARGO = CARGO_HOME / "bin" / "cargo"
MANIFEST = SOURCE / "tools" / "perf-baseline" / "Cargo.toml"
LOCK = SOURCE / "tools" / "perf-baseline" / "Cargo.lock"
CARGO_CONFIG = SOURCE / ".cargo" / "config.toml"
BINARY = TARGET / "release" / "litchi-perf-baseline"

BINDINGS = BUILD / "bindings.json"
START = BUILD / "start.json"
LAUNCH = BUILD / "launch.json"
TERMINAL = BUILD / "terminal.json"
BINARY_RECEIPT = BUILD / "binary.json"
RESULT = BUILD / "result.json"
STDOUT = BUILD / "build.stdout.log"
STDERR = BUILD / "build.stderr.log"

ARGV = [
    str(CARGO),
    "build",
    "--release",
    "--locked",
    "--offline",
    "--manifest-path",
    str(MANIFEST),
    "--bin",
    "litchi-perf-baseline",
    "-j2",
]

ENV = {
    "PATH": f"{CARGO_HOME / 'bin'}:/usr/bin:/bin",
    "HOME": str(HOME),
    "CARGO_HOME": str(CARGO_HOME),
    "RUSTUP_HOME": str(HOME / ".rustup"),
    "CARGO_TARGET_DIR": str(TARGET),
    "CARGO_NET_OFFLINE": "true",
    "CARGO_INCREMENTAL": "0",
    "CARGO_BUILD_JOBS": "2",
    "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1",
    "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "CARGO_PROFILE_RELEASE_LTO": "fat",
    "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3",
    "RUSTUP_TOOLCHAIN": "1.95.0",
    "LANG": "C",
    "LC_ALL": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
    "TMPDIR": str(TMP),
    "TEMP": str(TMP),
    "TMP": str(TMP),
}
UNSET_ENV = [
    "RUSTFLAGS",
    "CARGO_ENCODED_RUSTFLAGS",
    "CARGO_BUILD_RUSTC_WRAPPER",
    "RUSTC_WRAPPER",
]


class BuildError(RuntimeError):
    """A binding, custody, launch, or build failure."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise BuildError(message)


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def strict_json(path: Path) -> Any:
    def duplicate_key(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise BuildError(f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result

    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=duplicate_key)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise BuildError(f"cannot read JSON {path}: {error}") from error


def sha256_file(path: Path) -> tuple[int, str]:
    digest = hashlib.sha256()
    total = 0
    try:
        with path.open("rb") as source:
            for chunk in iter(lambda: source.read(1024 * 1024), b""):
                digest.update(chunk)
                total += len(chunk)
    except OSError as error:
        raise BuildError(f"cannot hash {path}: {error}") from error
    return total, digest.hexdigest()


def packet_path(path: Path) -> str:
    try:
        return path.resolve().relative_to(PACKET.resolve()).as_posix()
    except ValueError as error:
        raise BuildError(f"path is outside packet: {path}") from error


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    require(not path.is_symlink(), f"bound artifact must not be a symlink: {path}")
    require(path.is_file(), f"bound artifact is not a regular file: {path}")
    size, digest = sha256_file(path)
    mode = path.stat().st_mode & 0o777
    return {
        "path": packet_path(path),
        "bytes": size,
        "mode": mode,
        "executable": bool(mode & 0o111),
        "sha256": digest,
    }


def executable_descriptor(path: Path) -> dict[str, Any]:
    """Bind an executable launcher while retaining a symlink's public path."""

    path = path.absolute()
    require(path.exists() and path.is_file() and os.access(path, os.X_OK),
            f"bound executable is not runnable: {path}")
    resolved = path.resolve()
    require(resolved.is_file() and os.access(resolved, os.X_OK),
            f"resolved executable is not runnable: {resolved}")
    size, digest = sha256_file(resolved)
    mode = resolved.stat().st_mode & 0o777
    return {
        "path": str(path),
        "resolved_path": str(resolved),
        "bytes": size,
        "mode": mode,
        "executable": bool(mode & 0o111),
        "sha256": digest,
    }


def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite retained artifact {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False) + "\n").encode()
    try:
        with path.open("xb") as output:
            output.write(encoded)
            output.flush()
            os.fsync(output.fileno())
    except FileExistsError as error:
        raise BuildError(f"refusing to overwrite retained artifact {path}") from error
    directory = os.open(path.parent, os.O_RDONLY | os.O_DIRECTORY)
    try:
        os.fsync(directory)
    finally:
        os.close(directory)


def file_list(root: Path) -> list[str]:
    names: list[str] = []
    for directory, directories, files in os.walk(root, topdown=True, followlinks=False):
        directory_path = Path(directory)
        for name in directories:
            path = directory_path / name
            require(not path.is_symlink(), f"source directory is a symlink: {path}")
        for name in files:
            path = directory_path / name
            require(not path.is_symlink(), f"source file is a symlink: {path}")
            require(path.is_file(), f"source entry is not a regular file: {path}")
            names.append(path.relative_to(root).as_posix())
    return sorted(names)


def manifest_entries(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    value = manifest.get("files")
    if isinstance(value, list):
        entries = value
    elif isinstance(value, dict):
        entries = []
        for name, item in value.items():
            require(isinstance(item, dict), f"source manifest entry is not an object: {name}")
            item = dict(item)
            item.setdefault("path", name)
            entries.append(item)
    else:
        raise BuildError("source manifest files must be a list or mapping")
    return entries


def freeze_manifest_sha(freeze: dict[str, Any]) -> str | None:
    for key in ("manifest", "candidate_manifest", "source_manifest"):
        value = freeze.get(key)
        if isinstance(value, dict) and isinstance(value.get("sha256"), str):
            return value["sha256"]
        if isinstance(value, str) and len(value) == 64:
            return value
    value = freeze.get("manifest_sha256")
    return value if isinstance(value, str) else None


def validate_source() -> dict[str, Any]:
    require(SOURCE.is_dir(), f"missing frozen source root: {SOURCE}")
    require(SOURCE_MANIFEST.is_file(), f"missing source manifest: {SOURCE_MANIFEST}")
    manifest_size, manifest_sha = sha256_file(SOURCE_MANIFEST)
    manifest = strict_json(SOURCE_MANIFEST)
    require(isinstance(manifest, dict), "source manifest must be an object")
    require(isinstance(manifest.get("schema"), str), "source manifest schema is missing")
    require(manifest.get("source_root") == SOURCE.name, "source root binding mismatch")
    revision = (manifest.get("base_commit") or manifest.get("commit") or
                manifest.get("source_commit") or manifest.get("revision"))
    require(isinstance(revision, str) and revision, "source revision binding is missing")
    file_count = manifest.get("file_count")
    source_bytes = manifest.get("bytes", manifest.get("source_bytes"))
    require(isinstance(file_count, int) and file_count > 0, "source file count binding is missing")
    require(isinstance(source_bytes, int) and source_bytes >= 0, "source byte binding is missing")
    expected: dict[str, dict[str, Any]] = {}
    for item in manifest_entries(manifest):
        require(isinstance(item, dict), "source manifest entry is not an object")
        name = item.get("path")
        require(isinstance(name, str) and name and not Path(name).is_absolute(),
                "source manifest path is invalid")
        normalized = Path(name)
        require(".." not in normalized.parts and normalized.as_posix() == name,
                f"source manifest path is not normalized: {name}")
        require(name not in expected, f"duplicate source manifest path: {name}")
        require(isinstance(item.get("bytes"), int) and item["bytes"] >= 0,
                f"source manifest bytes missing: {name}")
        require(isinstance(item.get("sha256"), str) and len(item["sha256"]) == 64,
                f"source manifest SHA missing: {name}")
        expected[name] = item
    require(len(expected) == file_count, "source manifest entry count mismatch")
    actual_names = file_list(SOURCE)
    require(actual_names == sorted(expected), "source tree differs from frozen file set")
    actual_bytes = 0
    for name in actual_names:
        path = SOURCE / name
        size, digest = sha256_file(path)
        item = expected[name]
        require(size == item["bytes"], f"source bytes differ: {name}")
        require(digest == item["sha256"], f"source SHA differs: {name}")
        actual_bytes += size
    require(actual_bytes == source_bytes, "source byte total differs from frozen manifest")
    for path in (MANIFEST, LOCK, CARGO_CONFIG):
        require(path.is_file() and not path.is_symlink(), f"missing bound build input: {path}")
    require(CARGO.is_file() and os.access(CARGO, os.X_OK),
            f"missing bound executable: {CARGO}")
    require(FREEZE_RECEIPT.is_file(), f"missing candidate freeze receipt: {FREEZE_RECEIPT}")
    freeze_size, freeze_sha = sha256_file(FREEZE_RECEIPT)
    freeze = strict_json(FREEZE_RECEIPT)
    require(isinstance(freeze, dict), "candidate freeze receipt must be an object")
    status = freeze.get("status")
    require(isinstance(status, str) and "frozen" in status.lower(),
            "candidate freeze is not immutable")
    if "source_root" in freeze:
        require(freeze["source_root"] == SOURCE.name, "candidate freeze source root mismatch")
    require(freeze_manifest_sha(freeze) == manifest_sha,
            "candidate freeze manifest binding mismatch")
    if "file_count" in freeze:
        require(freeze["file_count"] == file_count, "candidate freeze file count mismatch")
    if "source_bytes" in freeze:
        require(freeze["source_bytes"] == source_bytes, "candidate freeze byte count mismatch")
    changes = manifest.get("changes", [])
    require(isinstance(changes, list), "source changes binding must be a list")
    if "changes" in freeze:
        require(freeze["changes"] == changes, "candidate freeze change binding mismatch")
    return {
        "manifest": descriptor(SOURCE_MANIFEST),
        "manifest_schema": manifest["schema"],
        "manifest_size": manifest_size,
        "freeze_receipt": {**descriptor(FREEZE_RECEIPT), "bytes_verified": freeze_size},
        "source_root": SOURCE.name,
        "source_commit": revision,
        "file_count": file_count,
        "bytes": source_bytes,
        "changes": changes,
        "cargo_manifest": descriptor(MANIFEST),
        "cargo_lock": descriptor(LOCK),
        "cargo_config": descriptor(CARGO_CONFIG),
    }


def validate_clean_packet() -> None:
    for path in (BUILD, TARGET, TMP):
        require(not path.exists(), f"refusing to reuse existing build path: {path}")
    BUILD.parent.mkdir(parents=True, exist_ok=True)
    TMP.parent.mkdir(parents=True, exist_ok=True)
    BUILD.mkdir()
    TARGET.mkdir()
    TMP.mkdir()


def wait_for_group_clear(pgid: int | None, seconds: float = 3.0) -> bool:
    if pgid is None:
        return True
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        try:
            os.killpg(pgid, 0)
        except ProcessLookupError:
            return True
        except PermissionError:
            return False
        time.sleep(0.05)
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return True
    except PermissionError:
        return False
    return False


def stop_group(pgid: int | None, proc: subprocess.Popen[bytes]) -> dict[str, Any]:
    actions: list[str] = []
    if pgid is not None:
        try:
            os.killpg(pgid, signal.SIGTERM)
            actions.append("SIGTERM")
        except ProcessLookupError:
            pass
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        if pgid is not None:
            try:
                os.killpg(pgid, signal.SIGKILL)
                actions.append("SIGKILL")
            except ProcessLookupError:
                pass
        proc.wait()
    clear = wait_for_group_clear(pgid)
    if not clear and pgid is not None:
        try:
            os.killpg(pgid, signal.SIGKILL)
            actions.append("SIGKILL-after-wait")
        except ProcessLookupError:
            pass
        clear = wait_for_group_clear(pgid)
    return {"actions": actions, "process_group_clear": clear, "reaped": True}


def log_descriptor(path: Path) -> dict[str, Any] | None:
    return descriptor(path) if path.is_file() else None


def main() -> int:
    # Every artifact path is reserved before Cargo starts; no run can silently
    # replace a prior build or reuse a prior target cache.
    source = validate_source()
    validate_clean_packet()
    runner = descriptor(RUNNER)
    cargo = executable_descriptor(CARGO)
    for key in UNSET_ENV:
        require(key not in ENV, f"forbidden environment key is set: {key}")
    require(ENV["CARGO_TARGET_DIR"] == str(TARGET), "target binding mismatch")

    common = {
        "schema": "litchi.performance.0948.candidate-native-build-v1",
        "source": source,
        "runner": runner,
        "cargo": cargo,
        "argv": ARGV,
        "cwd": str(SOURCE),
        "cwd_relative": packet_path(SOURCE),
        "environment": dict(sorted(ENV.items())),
        "unset_environment": sorted(UNSET_ENV),
        "target_relative": packet_path(TARGET),
        "binary_relative": packet_path(BINARY),
        "deadline_seconds": DEADLINE_SECONDS,
        "start_new_session": True,
        "process_group": True,
        "terminal_before_binary_parse": True,
    }
    write_once(BINDINGS, {
        **common,
        "receipt_kind": "prelaunch-bindings",
        "bound_utc": utc_now(),
        "bindings_path": packet_path(BINDINGS),
    })

    started = utc_now()
    write_once(START, {
        **common,
        "receipt_kind": "start",
        "started_utc": started,
        "bindings": descriptor(BINDINGS),
        "stdout_path": packet_path(STDOUT),
        "stderr_path": packet_path(STDERR),
        "launch_receipt": packet_path(LAUNCH),
        "terminal_receipt": packet_path(TERMINAL),
        "binary_receipt": packet_path(BINARY_RECEIPT),
        "result_receipt": packet_path(RESULT),
    })

    stdout_handle = None
    stderr_handle = None
    process: subprocess.Popen[bytes] | None = None
    pid: int | None = None
    pgid: int | None = None
    launch_error: str | None = None
    launch_utc = utc_now()
    timed_out = False
    returncode: int | None = None
    cleanup = {"actions": [], "process_group_clear": True, "reaped": False}
    try:
        stdout_handle = STDOUT.open("xb")
        stderr_handle = STDERR.open("xb")
        try:
            process = subprocess.Popen(
                ARGV,
                cwd=SOURCE,
                env=dict(ENV),
                stdin=subprocess.DEVNULL,
                stdout=stdout_handle,
                stderr=stderr_handle,
                start_new_session=True,
            )
            pid = process.pid
            try:
                pgid = os.getpgid(pid)
            except ProcessLookupError:
                # A fast failure still has the PID as its new-session group ID.
                pgid = pid
        except OSError as error:
            launch_error = f"{type(error).__name__}: {error}"
    except OSError as error:
        launch_error = f"{type(error).__name__}: {error}"

    if process is not None:
        write_once(LAUNCH, {
            **common,
            "receipt_kind": "launch",
            "launched_utc": launch_utc,
            "status": "launched",
            "pid": pid,
            "process_group_id": pgid,
            "start_receipt": packet_path(START),
        })
        print(json.dumps({
            "status": "launched",
            "pid": pid,
            "process_group_id": pgid,
            "start": str(START),
            "terminal": str(TERMINAL),
            "binary": str(BINARY),
        }, sort_keys=True), flush=True)
        deadline = time.monotonic() + DEADLINE_SECONDS
        while True:
            returncode = process.poll()
            if returncode is not None:
                break
            if time.monotonic() >= deadline:
                timed_out = True
                cleanup = stop_group(pgid, process)
                returncode = process.returncode
                break
            time.sleep(0.25)
        if not cleanup["reaped"]:
            returncode = process.wait()
            cleanup = {
                "actions": [],
                "process_group_clear": wait_for_group_clear(pgid),
                "reaped": True,
            }
    else:
        write_once(LAUNCH, {
            **common,
            "receipt_kind": "launch",
            "launched_utc": launch_utc,
            "status": "launch_failed",
            "pid": None,
            "process_group_id": None,
            "error": launch_error,
            "start_receipt": packet_path(START),
        })

    if stdout_handle is not None:
        stdout_handle.close()
    if stderr_handle is not None:
        stderr_handle.close()

    terminal_utc = utc_now()
    terminal = {
        **common,
        "receipt_kind": "terminal",
        "started_utc": started,
        "launched_utc": launch_utc,
        "terminal_utc": terminal_utc,
        "status": "launch_failed" if process is None else "terminal",
        "pid": pid,
        "process_group_id": pgid,
        "returncode": returncode,
        "timed_out": timed_out,
        "launch_error": launch_error,
        "cleanup": cleanup,
        "start_receipt": packet_path(START),
        "launch_receipt": packet_path(LAUNCH),
        "stdout": log_descriptor(STDOUT),
        "stderr": log_descriptor(STDERR),
        # The terminal is durable before any binary digest is computed.
        "binary": {"path": packet_path(BINARY), "exists": BINARY.is_file()},
    }
    write_once(TERMINAL, terminal)
    print(json.dumps({
        "status": terminal["status"],
        "returncode": returncode,
        "timed_out": timed_out,
        "terminal": str(TERMINAL),
    }, sort_keys=True), flush=True)

    binary_info: dict[str, Any]
    if BINARY.is_file() and not BINARY.is_symlink():
        binary_info = {
            "path": packet_path(BINARY),
            "bytes": BINARY.stat().st_size,
            "sha256": sha256_file(BINARY)[1],
            "mode": BINARY.stat().st_mode & 0o777,
            "executable": bool(BINARY.stat().st_mode & 0o111),
            "terminal_receipt": packet_path(TERMINAL),
            "terminal_sha256": sha256_file(TERMINAL)[1],
            "status": "present",
        }
    else:
        binary_info = {
            "path": packet_path(BINARY),
            "status": "missing",
            "terminal_receipt": packet_path(TERMINAL),
            "terminal_sha256": sha256_file(TERMINAL)[1],
        }
    write_once(BINARY_RECEIPT, {
        "schema": "litchi.performance.0948.candidate-native-binary-receipt-v1",
        "build_terminal": descriptor(TERMINAL),
        "source_commit": source["source_commit"],
        "binary": binary_info,
    })
    result_status = (
        process is not None and returncode == 0 and not timed_out
        and cleanup["process_group_clear"] and binary_info["status"] == "present"
    )
    write_once(RESULT, {
        "schema": "litchi.performance.0948.candidate-native-build-result-v1",
        "status": "passed" if result_status else "failed",
        "terminal": descriptor(TERMINAL),
        "binary_receipt": descriptor(BINARY_RECEIPT),
        "source_commit": source["source_commit"],
        "binary": binary_info,
    })
    return 0 if result_status else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except BuildError as error:
        print(f"build wrapper refused: {error}", file=sys.stderr)
        raise SystemExit(2)
