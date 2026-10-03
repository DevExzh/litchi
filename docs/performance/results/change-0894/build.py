#!/usr/bin/env python3
"""Build the 0894 baseline binaries from the frozen source archive.

The builder owns a marked source root and a marked Cargo target root under the
workspace.  It never reads the live source tree for compilation, except for
the source/archive and lock receipts supplied by ``source.py``.  Native and
allocator builds use the same pinned release settings as the retained 0891
profiles and leave complete command, source, binary, depfile, and ELF build-ID
receipts for the protocol agent.

All outputs are create-only.  A failed role leaves its start/role/failure
receipts and logs in place; callers must preserve that packet and use a fresh
attempt for a retry.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import stat
import subprocess
import tarfile
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
ARCHIVE = PACKET / "source.tar.gz"
MANIFEST = PACKET / "source.json"
START = PACKET / "build-start.json"
BUILD = PACKET / "build.json"
SOURCE = ROOT / "target/litchi-0894-baseline-src"
TARGET = ROOT / "target/litchi-0894-baseline-build"
BINARIES = PACKET / "binaries"
ROOT_LOCK = ROOT / "Cargo.lock"
MARKER = ".owner-litchi-0894-baseline"
SOURCE_MARKER_TEXT = "litchi 0894 frozen baseline source\n"
TARGET_MARKER_TEXT = "litchi 0894 frozen baseline build target\n"
REVISION = "506aa2edae6666d9c76ae6827c2f5941a8e47896"
GUARD_CONTRACT = "post-endpoint-per-worksheet-tolerance-v1"
INHERITED_ENVIRONMENT = ("PATH", "HOME", "CARGO_HOME", "RUSTUP_HOME")
FORBIDDEN_OVERRIDES = ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_WRAPPER", "RUSTC_WORKSPACE_WRAPPER")
BUILD_FLAGS = {
    "toolchain": "1.95.0",
    "profile": "release",
    "opt_level": 3,
    "debug": 1,
    "lto": "fat",
    "codegen_units": 1,
    "jobs": 2,
}
ROLES = (
    ("native", "litchi-perf-baseline", ()),
    ("allocator", "litchi-perf-baseline-alloc", ("allocator-metrics",)),
)


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    info = path.stat()
    if path.is_symlink() or not stat.S_ISREG(info.st_mode):
        raise RuntimeError(f"expected regular non-symlink file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": sha256(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def write_once(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def load_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def archive_members() -> dict[str, dict[str, Any]]:
    members: dict[str, dict[str, Any]] = {}
    with tarfile.open(ARCHIVE, "r:gz") as handle:
        for member in handle:
            path = Path(member.name)
            require(not path.is_absolute() and ".." not in path.parts, f"unsafe archive path: {member.name}")
            if member.isdir():
                continue
            require(member.isfile(), f"archive contains non-regular member: {member.name}")
            require(member.name not in members, f"duplicate archive member: {member.name}")
            source = handle.extractfile(member)
            require(source is not None, f"cannot read archive member: {member.name}")
            digest = hashlib.sha256()
            size = 0
            with source:
                for block in iter(lambda: source.read(1024 * 1024), b""):
                    digest.update(block)
                    size += len(block)
            members[member.name] = {
                "bytes": size,
                "sha256": digest.hexdigest(),
                "mode": stat.S_IMODE(member.mode),
            }
    return members


def source_snapshot() -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for path in SOURCE.rglob("*"):
        if path == SOURCE / MARKER:
            continue
        if path.is_dir():
            continue
        require(not path.is_symlink() and path.is_file(), f"source contains unexpected non-regular file: {path}")
        info = path.stat()
        result[path.relative_to(SOURCE).as_posix()] = {
            "bytes": info.st_size,
            "sha256": sha256(path),
            "mode": stat.S_IMODE(info.st_mode),
        }
    return dict(sorted(result.items()))


def source_receipt(label: str, environment: dict[str, str], root_lock: dict[str, Any]) -> dict[str, Any]:
    snapshot = source_snapshot()
    return {
        "schema": f"litchi.xlsx-source-{label}-0894.v1",
        "status": "pass",
        "created_utc": now(),
        "source_root": str(SOURCE.resolve()),
        "source_marker": {"path": str((SOURCE / MARKER).resolve()), "text": SOURCE_MARKER_TEXT},
        "file_count": len(snapshot),
        "files": snapshot,
        "root_generated_lock": root_lock,
        "environment": {key: environment[key] for key in ("TMPDIR", "TEMP", "TMP")},
    }


def build_environment() -> dict[str, str]:
    unexpected = [key for key in FORBIDDEN_OVERRIDES if os.environ.get(key)]
    require(not unexpected, "unexpected compiler override in environment: " + ", ".join(unexpected))
    environment = {key: os.environ[key] for key in INHERITED_ENVIRONMENT if key in os.environ}
    private_tmp = TARGET / "tmp"
    private_tmp.mkdir(parents=True, exist_ok=True)
    environment.update(
        {
            "LANG": "C",
            "LC_ALL": "C",
            "TZ": "UTC",
            "RUSTUP_TOOLCHAIN": "1.95.0",
            "TMPDIR": str(private_tmp.resolve()),
            "TEMP": str(private_tmp.resolve()),
            "TMP": str(private_tmp.resolve()),
            "CARGO_TARGET_DIR": str(TARGET.resolve()),
            "CARGO_BUILD_JOBS": "2",
            "CARGO_NET_OFFLINE": "true",
            "CARGO_INCREMENTAL": "0",
            "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3",
            "CARGO_PROFILE_RELEASE_DEBUG": "1",
            "CARGO_PROFILE_RELEASE_LTO": "fat",
            "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1",
        }
    )
    return environment


def elf_build_id(path: Path, environment: dict[str, str]) -> str:
    result = subprocess.run(
        ["readelf", "-n", str(path)],
        cwd=SOURCE,
        env=environment,
        check=True,
        capture_output=True,
        text=True,
    )
    match = re.search(r"Build ID:\s*([0-9a-fA-F]+)", result.stdout)
    require(match is not None, f"ELF build ID is missing: {path}")
    return match.group(1).lower()


def validate_source(manifest: dict[str, Any]) -> dict[str, dict[str, Any]]:
    require(manifest.get("schema") == "litchi.xlsx-source-snapshot-0894.v1", "source manifest schema mismatch")
    require(manifest.get("status") == "pass", "source manifest is not passing")
    require(manifest.get("revision") == REVISION, "source manifest revision mismatch")
    require(manifest.get("file_count") == 10_771, "source manifest file count mismatch")
    files = manifest.get("files")
    require(isinstance(files, dict) and len(files) == manifest["file_count"], "source file census is incomplete")
    archive = manifest.get("archive")
    require(isinstance(archive, dict), "source archive descriptor is missing")
    require(archive.get("path") == str(ARCHIVE.resolve()), "source archive path mismatch")
    require(archive.get("bytes") == ARCHIVE.stat().st_size and archive.get("sha256") == sha256(ARCHIVE), "source archive identity mismatch")
    standalone = manifest.get("standalone_lock")
    require(isinstance(standalone, dict) and standalone.get("path") == "tools/perf-baseline/Cargo.lock", "standalone lock identity is missing")
    require(standalone.get("sha256") == files["tools/perf-baseline/Cargo.lock"]["sha256"], "standalone lock does not match source census")
    root_lock = manifest.get("root_generated_lock")
    require(isinstance(root_lock, dict), "root generated lock identity is missing")
    current_root_lock = descriptor(ROOT_LOCK)
    require(root_lock.get("path") == current_root_lock["path"], "root generated lock path changed")
    require(root_lock.get("bytes") == current_root_lock["bytes"] and root_lock.get("sha256") == current_root_lock["sha256"], "root generated lock changed")
    return files


def extract_and_check(files: dict[str, dict[str, Any]]) -> None:
    members = archive_members()
    expected = set(files)
    require(set(members) == expected, f"archive census mismatch: expected {len(expected)}, got {len(members)}")
    for name, expected_record in files.items():
        observed = members[name]
        for key in ("bytes", "sha256", "mode"):
            require(observed[key] == expected_record[key], f"archive {key} mismatch: {name}")
    with tarfile.open(ARCHIVE, "r:gz") as handle:
        handle.extractall(SOURCE, filter="data")
    observed_files = source_snapshot()
    actual = set(observed_files)
    require(actual == expected, f"extracted source census mismatch: expected {len(expected)}, got {len(actual)}")
    for name, expected_record in files.items():
        observed = observed_files[name]
        for key in ("bytes", "sha256", "mode"):
            require(observed[key] == expected_record[key], f"source {key} mismatch: {name}")


def command(role: str, binary: str, features: tuple[str, ...]) -> list[str]:
    result = [
        "cargo",
        "build",
        "--release",
        "--locked",
        "--offline",
        "--manifest-path",
        "tools/perf-baseline/Cargo.toml",
        "--bin",
        binary,
        "-j",
        "2",
    ]
    if features:
        result.extend(["--features", ",".join(features)])
    return result


def main() -> int:
    for path in (START, BUILD, SOURCE, TARGET, BINARIES):
        require(not path.exists(), f"owned 0894 path already exists: {path}")
    require(ARCHIVE.is_file() and MANIFEST.is_file(), "run source.py before build.py")
    manifest = load_json(MANIFEST)
    files = validate_source(manifest)

    SOURCE.mkdir(parents=True)
    TARGET.mkdir(parents=True)
    (SOURCE / MARKER).write_text(SOURCE_MARKER_TEXT, encoding="utf-8")
    (TARGET / MARKER).write_text(TARGET_MARKER_TEXT, encoding="utf-8")
    extract_and_check(files)
    environment = build_environment()
    root_lock_before = descriptor(ROOT_LOCK)
    before = source_receipt("before", environment, root_lock_before)
    write_once(PACKET / "source-before.json", before)

    start = {
        "schema": "litchi.xlsx-baseline-build-start-0894.v1",
        "status": "building",
        "started_utc": now(),
        "revision": REVISION,
        "base_commit": REVISION,
        "baseline_commit": REVISION,
        "source_revision": REVISION,
        "profile": "release",
        "guard_contract": GUARD_CONTRACT,
        "timed_harness_guard_contract": GUARD_CONTRACT,
        "flags": BUILD_FLAGS,
        "argv": ["cargo", "build", "--release", "--locked", "--offline"],
        "source_manifest": descriptor(MANIFEST),
        "source_archive": descriptor(ARCHIVE),
        "source_root": str(SOURCE.resolve()),
        "target_root": str(TARGET.resolve()),
        "source_marker": {"path": str((SOURCE / MARKER).resolve()), "text": SOURCE_MARKER_TEXT},
        "target_marker": {"path": str((TARGET / MARKER).resolve()), "text": TARGET_MARKER_TEXT},
        "environment": environment,
        "compiler": subprocess.check_output(["rustc", "--version", "--verbose"], cwd=SOURCE, env=environment, text=True),
        "cargo": subprocess.check_output(["cargo", "--version"], cwd=SOURCE, env=environment, text=True).strip(),
        "root_generated_lock": root_lock_before,
        "standalone_lock": manifest["standalone_lock"],
        "binaries": {},
    }
    write_once(START, start)
    BINARIES.mkdir()

    failure: dict[str, Any] | None = None
    try:
        for role, binary, features in ROLES:
            role_start_path = PACKET / f"{role}-build-start.json"
            role_receipt_path = PACKET / f"{role}-build.json"
            log_path = PACKET / f"{role}-build.log"
            require(not role_start_path.exists() and not role_receipt_path.exists() and not log_path.exists(), f"{role} receipt/log already exists")
            argv = command(role, binary, features)
            role_start = {
                "schema": f"litchi.xlsx-{role}-build-start-0894.v1",
                "status": "building",
                "role": role,
                "binary": binary,
                "features": list(features),
                "argv": argv,
                "cwd": str(SOURCE.resolve()),
                "environment": environment,
                "started_utc": now(),
            }
            write_once(role_start_path, role_start)
            with log_path.open("x", encoding="utf-8") as handle:
                completed = subprocess.run(argv, cwd=SOURCE, env=environment, stdout=handle, stderr=subprocess.STDOUT, check=False)
            role_receipt = {
                "schema": f"litchi.xlsx-{role}-build-0894.v1",
                "status": "pass" if completed.returncode == 0 else "fail",
                "role": role,
                "binary": binary,
                "features": list(features),
                "argv": argv,
                "cwd": str(SOURCE.resolve()),
                "environment": environment,
                "started_utc": role_start["started_utc"],
                "finished_utc": now(),
                "exit_code": completed.returncode,
                "log": descriptor(log_path),
            }
            write_once(role_receipt_path, role_receipt)
            if completed.returncode:
                failure = {"role": role, "exit_code": completed.returncode, "receipt": descriptor(role_receipt_path)}
                raise RuntimeError(f"{role} release build failed with exit code {completed.returncode}")

            executable = TARGET / "release" / binary
            depfile = TARGET / "release" / f"{binary}.d"
            require(executable.is_file() and executable.stat().st_mode & stat.S_IXUSR, f"built executable missing: {executable}")
            require(depfile.is_file(), f"built depfile missing: {depfile}")
            retained = BINARIES / f"baseline-{role}"
            shutil.copy2(executable, retained)
            retained.chmod(retained.stat().st_mode | stat.S_IXUSR)
            retained_depfile = PACKET / f"{role}.d"
            shutil.copy2(depfile, retained_depfile)
            start["binaries"][role] = {
                "role": role,
                "name": binary,
                "features": list(features),
                "path": str(retained.resolve()),
                **descriptor(retained),
                "executable": True,
                "elf_build_id": elf_build_id(retained, environment),
                "build_receipt": descriptor(role_receipt_path),
                "depfile": descriptor(retained_depfile),
            }

        root_lock_after = descriptor(ROOT_LOCK)
        require(root_lock_after == root_lock_before, "root Cargo.lock changed during release builds")
        after = source_receipt("after", environment, root_lock_after)
        require(after["files"] == before["files"], "frozen source changed during release builds")
        write_once(PACKET / "source-after.json", after)
        start["status"] = "pass"
        start["finished_utc"] = now()
        start["source_before"] = descriptor(PACKET / "source-before.json")
        start["source_after"] = descriptor(PACKET / "source-after.json")
        final = {
            **start,
            "schema": "litchi.xlsx-baseline-build-0894.v1",
            "status": "pass",
            "build_start": descriptor(START),
            "source_rechecked_files": len(before["files"]),
            "performance_claim": None,
        }
        write_once(BUILD, final)
        print("baseline native and allocator binaries retained", flush=True)
        return 0
    except Exception as exc:
        root_lock_after = descriptor(ROOT_LOCK)
        after_path = PACKET / "source-after.json"
        if not after_path.exists():
            try:
                write_once(after_path, source_receipt("after-failure", environment, root_lock_after))
            except Exception:
                pass
        failure_receipt = {
            "schema": "litchi.xlsx-baseline-build-failure-0894.v1",
            "status": "fail",
            "finished_utc": now(),
            "error": str(exc),
            "role": failure.get("role") if failure else None,
            "exit_code": failure.get("exit_code") if failure else None,
            "source_before": descriptor(PACKET / "source-before.json") if (PACKET / "source-before.json").is_file() else None,
            "source_after": descriptor(after_path) if after_path.is_file() else None,
        }
        write_once(PACKET / "build-failure.json", failure_receipt)
        raise


if __name__ == "__main__":
    raise SystemExit(main())
