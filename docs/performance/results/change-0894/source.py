#!/usr/bin/env python3
"""Freeze the exact 0894 baseline source snapshot.

The source snapshot is made from the committed tree, rather than from the
working directory.  In particular, ignored or untracked files cannot enter
the archive.  The manifest keeps the complete Git tree census and the two
lock identities used by the build protocol: the tracked standalone
``tools/perf-baseline/Cargo.lock`` and the ignored root ``Cargo.lock`` that is
observed for provenance only.

This helper is create-only.  It writes a start receipt before creating the
archive and never replaces a receipt or a partial archive.  A failed attempt
therefore remains visible and the next attempt must use a fresh packet or an
explicitly preserved failure directory.
"""

from __future__ import annotations

import datetime as dt
import gzip
import hashlib
import json
from pathlib import Path
import shutil
import stat
import subprocess
import tarfile
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
REVISION = "506aa2edae6666d9c76ae6827c2f5941a8e47896"
SELECTION = (
    ".cargo",
    "Cargo.toml",
    "README.md",
    "clippy.toml",
    "deny.toml",
    "rust-toolchain.toml",
    "rustfmt.toml",
    "crates",
    "test-data",
    "tools/perf-baseline",
)
STANDALONE_LOCK = "tools/perf-baseline/Cargo.lock"
ROOT_LOCK = ROOT / "Cargo.lock"
UNTRACKED_BENCH = "crates/litchi/examples/bench_word_parse.rs"
GUARD_CONTRACT = "post-endpoint-per-worksheet-tolerance-v1"
ARCHIVE = PACKET / "source.tar.gz"
MANIFEST = PACKET / "source.json"
START = PACKET / "source-start.json"


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


def git(*args: str) -> str:
    return subprocess.check_output(["git", *args], cwd=ROOT, text=True).strip()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def git_entries() -> dict[str, dict[str, str]]:
    """Return the exact regular-file census selected from ``REVISION``."""

    raw = subprocess.check_output(
        ["git", "ls-tree", "--full-tree", "-r", "-z", REVISION, "--", *SELECTION],
        cwd=ROOT,
    )
    entries: dict[str, dict[str, str]] = {}
    for record in raw.split(b"\0"):
        if not record:
            continue
        header, name_bytes = record.split(b"\t", 1)
        mode, kind, oid = header.decode("ascii").split(" ")
        name = name_bytes.decode("utf-8")
        require(kind == "blob", f"selected Git entry is not a blob: {name}")
        require(name not in entries, f"duplicate selected Git path: {name}")
        entries[name] = {"mode": mode, "type": kind, "oid": oid}
    require(entries, "selected Git census is empty")
    return entries


def untracked_selected() -> list[str]:
    raw = git("ls-files", "--others", "--exclude-standard", "--", *SELECTION)
    return [line for line in raw.splitlines() if line]


def archive_members() -> dict[str, dict[str, Any]]:
    """Read every regular archive member and calculate its content identity."""

    members: dict[str, dict[str, Any]] = {}
    with tarfile.open(ARCHIVE, "r:gz") as handle:
        for member in handle:
            name = member.name
            path = Path(name)
            require(not path.is_absolute() and ".." not in path.parts, f"unsafe archive member: {name}")
            if member.isdir():
                continue
            require(member.isfile(), f"archive contains non-regular member: {name}")
            require(name not in members, f"duplicate archive member: {name}")
            source = handle.extractfile(member)
            require(source is not None, f"cannot read archive member: {name}")
            digest = hashlib.sha256()
            size = 0
            with source:
                for block in iter(lambda: source.read(1024 * 1024), b""):
                    digest.update(block)
                    size += len(block)
            members[name] = {
                "bytes": size,
                "sha256": digest.hexdigest(),
                "mode": stat.S_IMODE(member.mode),
            }
    return members


def untracked_descriptors(paths: list[str]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for name in paths:
        path = ROOT / name
        if not path.exists() or path.is_symlink() or not path.is_file():
            result.append({"path": name, "present": False})
            continue
        item = descriptor(path)
        item["path"] = name
        result.append(item)
    return result


def main() -> int:
    require(not START.exists(), f"source start receipt already exists: {START}")
    require(not ARCHIVE.exists(), f"source archive already exists: {ARCHIVE}")
    require(not MANIFEST.exists(), f"source manifest already exists: {MANIFEST}")
    require(git("rev-parse", "HEAD") == REVISION, "working HEAD is not the sealed 0894 baseline revision")
    require(git("rev-parse", REVISION) == REVISION, "sealed 0894 revision does not resolve exactly")
    require(ROOT_LOCK.exists(), f"ignored root Cargo.lock is required for explicit provenance: {ROOT_LOCK}")
    require((ROOT / STANDALONE_LOCK).is_file(), "tracked standalone Cargo.lock is missing")

    git_census = git_entries()
    selected_untracked = untracked_selected()
    root_lock_before = descriptor(ROOT_LOCK)
    start = {
        "schema": "litchi.xlsx-source-freeze-start-0894.v1",
        "status": "building",
        "started_utc": now(),
        "revision": REVISION,
        "observed_head": git("rev-parse", "HEAD"),
        "selection": list(SELECTION),
        "git_file_count": len(git_census),
        "untracked_selected": selected_untracked,
        "untracked_benchmark_excluded": UNTRACKED_BENCH not in git_census,
        "root_generated_lock": root_lock_before,
        "standalone_lock_path": STANDALONE_LOCK,
        "argv": ["python3", "-B", str(Path(__file__).resolve())],
        "archive_argv": ["git", "-c", "tar.umask=0022", "archive", "--format=tar", REVISION, "--", *SELECTION],
    }
    write_once(START, start)

    # Stream Git's deterministic tar output through a zero-mtime gzip member.
    # The archive is opened exclusively so a failed or interrupted attempt is
    # retained and cannot be silently replaced.
    with ARCHIVE.open("xb") as raw:
        with gzip.GzipFile(fileobj=raw, mode="wb", mtime=0) as compressed:
            process = subprocess.Popen(
                ["git", "-c", "tar.umask=0022", "archive", "--format=tar", REVISION, "--", *SELECTION],
                cwd=ROOT,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            require(process.stdout is not None and process.stderr is not None, "git archive pipes were not created")
            # Do not pass GzipFile as Popen's stdout: GzipFile exposes the
            # underlying descriptor and a child would bypass compression.
            shutil.copyfileobj(process.stdout, compressed, length=1024 * 1024)
            process.stdout.close()
            stderr = process.stderr.read()
            process.stderr.close()
            returncode = process.wait()
        if returncode:
            raise RuntimeError(
                "git archive failed with exit code "
                f"{returncode}: {stderr.decode('utf-8', 'replace').strip()}"
            )

    members = archive_members()
    require(set(members) == set(git_census), "archive regular-file census differs from Git tree census")
    files: dict[str, dict[str, Any]] = {}
    for name in sorted(git_census):
        record = git_census[name]
        member = members[name]
        expected_mode = int(record["mode"], 8)
        require(member["mode"] == stat.S_IMODE(expected_mode), f"archive mode mismatch: {name}")
        files[name] = {
            **member,
            "git_mode": record["mode"],
            "git_type": record["type"],
            "git_oid": record["oid"],
        }

    require(STANDALONE_LOCK in files, "standalone Cargo.lock was not included in selected archive")
    require(UNTRACKED_BENCH not in members, "untracked benchmark entered the Git archive")
    root_lock_after = descriptor(ROOT_LOCK)
    require(root_lock_after == root_lock_before, "root Cargo.lock changed while freezing the source archive")

    manifest = {
        "schema": "litchi.xlsx-source-snapshot-0894.v1",
        "status": "pass",
        "revision": REVISION,
        "baseline_commit": REVISION,
        "source_revision": REVISION,
        "timed_harness_guard_contract": GUARD_CONTRACT,
        "observed_head": git("rev-parse", "HEAD"),
        "selection": list(SELECTION),
        "file_count": len(files),
        "selected_file_count": len(files),
        "source_file_count": len(files),
        "live_files_hash_checked": len(files),
        "git_file_count": len(git_census),
        "files": files,
        "archive": descriptor(ARCHIVE),
        "standalone_lock": {
            "path": STANDALONE_LOCK,
            "source": "selected Git archive",
            **files[STANDALONE_LOCK],
        },
        "root_generated_lock": {
            "source": "ignored working-tree provenance; excluded from archive and standalone build source",
            **root_lock_after,
        },
        "untracked_selected": untracked_descriptors(selected_untracked),
        "untracked_benchmark_excluded": {
            "path": UNTRACKED_BENCH,
            "in_git_census": UNTRACKED_BENCH in git_census,
            "in_archive": UNTRACKED_BENCH in members,
        },
        "production_changes": False,
        "profiling_performed": False,
        "performance_claim": None,
        "created_utc": now(),
    }
    write_once(MANIFEST, manifest)
    print(f"sealed {len(files)} files at {REVISION}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
