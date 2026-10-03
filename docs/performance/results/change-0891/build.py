#!/usr/bin/env python3
"""Build the sealed 0891 x14ac candidate over the retained 0885 source.

The candidate is intentionally assembled from the retained source archive and
the frozen overlay.  This keeps the release build independent of unrelated
working-tree changes and makes the source census replayable.  The helper is
create-only: it refuses existing owned roots and receipts, never edits the
worktree, and leaves a failed build receipt/log in place for review.
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
BASELINE = PACKET.parent / "change-0885"
BASELINE_REVISION = "335a40350bf06f49691745c14d310fd7a3ca9117"
CURRENT_BASE = "3348479699f248223184c7bdbc005879b0153307"
ARCHIVE = BASELINE / "source.tar.gz"
MANIFEST = BASELINE / "source.json"
FREEZE = PACKET / "candidate-freeze-v3.json"
PATCH = PACKET / "candidate-v3.patch"
SOURCE = Path("/tmp/litchi-0891-source")
TARGET = ROOT / "target/litchi-0891-build"
ARCHIVE_SHA256 = "bdf6e1c815fbb33796cc779134d27a311cc25ba5043bb4f5e5a73fc4d5643bee"
CANDIDATE_PATHS = (
    "crates/litchi-xlsx/src/raw/worksheet/x14ac.rs",
    "crates/litchi-xlsx/src/raw/worksheet/x14ac/reuse_tests.rs",
)
CANDIDATE_ROOT = "crates/litchi-xlsx/src/raw/worksheet/x14ac"
DIAGNOSTIC_PATHS = frozenset(
    {
        "tools/perf-baseline/src/bin/xml_attribute_bounds.rs",
        "tools/perf-baseline/src/bin/xml_attribute_bounds/boundary_tests.rs",
        "tools/perf-baseline/src/bin/xml_attribute_bounds/input_tests.rs",
        "tools/perf-baseline/src/bin/xml_attribute_bounds/qualification_tests.rs",
    }
)
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
    if not stat.S_ISREG(info.st_mode) or path.is_symlink():
        raise ValueError(f"expected regular non-symlink file: {path}")
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


def patch_paths(raw: bytes) -> set[str]:
    paths: set[str] = set()
    for line in raw.decode("utf-8").splitlines():
        match = re.fullmatch(r"diff --git a/(.+) b/(.+)", line)
        if match is None:
            continue
        left, right = match.groups()
        name = right if left == "/dev/null" else left
        require(left == right or left == "/dev/null", f"candidate patch has asymmetric paths: {left!r} {right!r}")
        path = Path(name)
        require(not path.is_absolute() and ".." not in path.parts, f"candidate patch has unsafe path: {name}")
        paths.add(name)
    require(paths == set(CANDIDATE_PATHS), "candidate patch paths differ from the two frozen x14ac files")
    return paths


def load_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def load_freeze(manifest: dict[str, Any]) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    require(FREEZE.is_file() and PATCH.is_file(), "candidate-freeze-v3.json and candidate-v3.patch are required")
    freeze = load_json(FREEZE)
    require(freeze.get("schema") == "litchi.xlsx-x14ac-candidate-freeze-0891.v1", "candidate freeze schema mismatch")
    predecessor = freeze.get("supersedes")
    require(isinstance(predecessor, dict), "candidate freeze v3 predecessor descriptor is missing")
    predecessor_path = Path(str(predecessor.get("path", "")))
    if not predecessor_path.is_absolute():
        predecessor_path = ROOT / predecessor_path
    predecessor_path = predecessor_path.resolve()
    predecessor_file = PACKET / "candidate-freeze-v2.json"
    require(predecessor_path == predecessor_file.resolve(), "candidate freeze v3 predecessor path mismatch")
    require(predecessor_file.is_file() and predecessor.get("bytes") == predecessor_file.stat().st_size and predecessor.get("sha256") == sha256(predecessor_file), "candidate freeze v2 predecessor identity mismatch")
    require(freeze.get("base_commit") == CURRENT_BASE and git("rev-parse", "HEAD") == CURRENT_BASE, "candidate freeze is not bound to HEAD 3348479699")
    paths = freeze.get("candidate_paths")
    require(isinstance(paths, list) and set(paths) == set(CANDIDATE_PATHS) and len(paths) == len(CANDIDATE_PATHS), "candidate path inventory mismatch")
    entries = freeze.get("files")
    require(isinstance(entries, list) and len(entries) == len(CANDIDATE_PATHS), "candidate file inventory is incomplete")
    files: dict[str, dict[str, Any]] = {}
    for entry in entries:
        require(isinstance(entry, dict) and isinstance(entry.get("path"), str), "invalid candidate file descriptor")
        name = entry["path"]
        require(name in CANDIDATE_PATHS and name not in files, f"unexpected or duplicate candidate file: {name}")
        require(isinstance(entry.get("bytes"), int) and entry["bytes"] >= 0, f"candidate byte count missing: {name}")
        require(isinstance(entry.get("sha256"), str) and re.fullmatch(r"[0-9a-f]{64}", entry["sha256"]) is not None, f"candidate hash missing: {name}")
        current = ROOT / name
        require(current.is_file() and not current.is_symlink(), f"candidate overlay is missing: {name}")
        require(entry["bytes"] == current.stat().st_size and entry["sha256"] == sha256(current), f"candidate overlay changed after freeze: {name}")
        files[name] = entry
    require(set(files) == set(CANDIDATE_PATHS), "candidate file and path inventories differ")
    require(patch_paths(PATCH.read_bytes()) == set(files), "candidate patch and freeze paths differ")
    patch_record = freeze.get("patch")
    require(isinstance(patch_record, dict), "candidate patch descriptor is missing")
    patch_path = Path(str(patch_record.get("path", "")))
    if not patch_path.is_absolute():
        patch_path = ROOT / patch_path
    require(patch_path.resolve() == PATCH.resolve(), "candidate patch path is not packet-local")
    require(patch_record.get("bytes") == PATCH.stat().st_size and patch_record.get("sha256") == sha256(PATCH), "candidate patch changed after freeze")
    require(freeze.get("immutable") is True or freeze.get("create_only") is True, "candidate freeze must declare immutable/create-only custody")

    baseline_files = manifest.get("files")
    require(isinstance(baseline_files, dict) and len(baseline_files) == 10_765, "retained source manifest file count mismatch")
    tracked = {
        path
        for path in git("diff", "--name-only", BASELINE_REVISION, "--", *SELECTION).splitlines()
        if path
    }
    untracked = {
        path
        for path in git("ls-files", "--others", "--exclude-standard", "--", CANDIDATE_ROOT).splitlines()
        if path
    }
    allowed = set(CANDIDATE_PATHS) | DIAGNOSTIC_PATHS
    require(not (tracked - allowed), "selected source differs outside candidate/known diagnostic paths: " + ", ".join(sorted(tracked - allowed)))
    require(not (untracked - set(CANDIDATE_PATHS)), "untracked x14ac source is outside the frozen inventory: " + ", ".join(sorted(untracked - set(CANDIDATE_PATHS))))
    return freeze, files


def source_environment(target: Path) -> dict[str, str]:
    env = {key: os.environ[key] for key in ("PATH", "HOME", "CARGO_HOME", "RUSTUP_HOME") if key in os.environ}
    env.update(
        {
            "LANG": "C",
            "LC_ALL": "C",
            "TZ": "UTC",
            "CARGO_TARGET_DIR": str(target),
            "CARGO_BUILD_JOBS": "2",
            "CARGO_NET_OFFLINE": "true",
            "CARGO_INCREMENTAL": "0",
            "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3",
            "CARGO_PROFILE_RELEASE_DEBUG": "1",
            "CARGO_PROFILE_RELEASE_LTO": "fat",
            "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1",
        }
    )
    return env


def main() -> int:
    for path in (SOURCE, TARGET, PACKET / "build-start.json", PACKET / "build.json", PACKET / "binaries"):
        require(not path.exists(), f"owned 0891 path already exists: {path}")
    require(ARCHIVE.is_file() and MANIFEST.is_file(), "retained 0885 source archive/manifest is missing")
    require(sha256(ARCHIVE) == ARCHIVE_SHA256, "retained 0885 source archive hash changed")
    manifest = load_json(MANIFEST)
    require(manifest.get("revision") == BASELINE_REVISION and manifest.get("file_count") == 10_765, "retained 0885 manifest identity mismatch")
    freeze, frozen_files = load_freeze(manifest)

    SOURCE.mkdir()
    TARGET.mkdir()
    (SOURCE / ".owner-0891").write_text("x14ac candidate 0891\n", encoding="utf-8")
    (TARGET / ".owner-0891").write_text("x14ac candidate 0891 target\n", encoding="utf-8")
    with tarfile.open(ARCHIVE, "r:*") as archive:
        archive.extractall(SOURCE, filter="data")
    for name in CANDIDATE_PATHS:
        destination = SOURCE / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / name, destination)

    baseline_names = set(manifest["files"])
    expected_names = baseline_names | set(CANDIDATE_PATHS)
    actual_names = {
        path.relative_to(SOURCE).as_posix()
        for path in SOURCE.rglob("*")
        if path.is_file() and path.name != ".owner-0891"
    }
    require(actual_names == expected_names, f"candidate source census mismatch: expected {len(expected_names)}, got {len(actual_names)}")
    expected_hashes = {name: frozen_files.get(name, manifest["files"].get(name))["sha256"] for name in expected_names}
    for name, expected in expected_hashes.items():
        require(sha256(SOURCE / name) == expected, f"candidate source hash mismatch: {name}")

    env = source_environment(TARGET)
    build = {
        "schema": "litchi.xlsx-x14ac-candidate-build-0891.v1",
        "status": "building",
        "base_commit": CURRENT_BASE,
        "baseline_revision": BASELINE_REVISION,
        "source_archive": descriptor(ARCHIVE),
        "source_manifest": descriptor(MANIFEST),
        "candidate_freeze": descriptor(FREEZE),
        "candidate_patch": descriptor(PATCH),
        "source_root": str(SOURCE),
        "target_root": str(TARGET),
        "source_marker": {"path": str(SOURCE / ".owner-0891"), "text": "x14ac candidate 0891\n"},
        "target_marker": {"path": str(TARGET / ".owner-0891"), "text": "x14ac candidate 0891 target\n"},
        "source_file_count": len(expected_names),
        "source_overlay_files": list(CANDIDATE_PATHS),
        "environment": env,
        "compiler": subprocess.check_output(["rustc", "--version", "--verbose"], cwd=SOURCE, env=env, text=True),
        "cargo": subprocess.check_output(["cargo", "--version"], cwd=SOURCE, env=env, text=True).strip(),
        "binaries": {},
    }
    write_once(PACKET / "build-start.json", build)
    (PACKET / "binaries").mkdir()
    for role, binary_name in (("native", "litchi-perf-baseline"), ("allocator", "litchi-perf-baseline-alloc")):
        argv = [
            "cargo", "build", "--release", "--locked", "--offline",
            "--manifest-path", "tools/perf-baseline/Cargo.toml",
            "--bin", binary_name, "-j", "2",
        ]
        if role == "allocator":
            argv += ["--features", "allocator-metrics"]
        baseline_receipt = load_json(BASELINE / f"{role}-build.json")
        require(argv == baseline_receipt.get("argv"), f"candidate {role} build argv differs from retained 0885")
        start = {"schema": "litchi.xlsx-x14ac-build-attempt-0891.v1", "role": role, "argv": argv, "cwd": str(SOURCE), "environment": env, "started_utc": now()}
        write_once(PACKET / f"{role}-build-start.json", start)
        log = PACKET / f"{role}-build.log"
        with log.open("x", encoding="utf-8") as handle:
            code = subprocess.run(argv, cwd=SOURCE, env=env, stdout=handle, stderr=subprocess.STDOUT, check=False).returncode
        receipt = dict(start)
        receipt.update({"status": "pass" if code == 0 else "fail", "exit_code": code, "finished_utc": now(), "log": descriptor(log)})
        write_once(PACKET / f"{role}-build.json", receipt)
        if code != 0:
            raise SystemExit(code)
        executable = TARGET / "release" / binary_name
        depfile = TARGET / "release" / f"{binary_name}.d"
        require(executable.is_file() and depfile.is_file(), f"{role} release outputs are missing")
        retained = PACKET / "binaries" / role
        shutil.copy2(executable, retained)
        shutil.copy2(depfile, PACKET / f"{role}.d")
        build["binaries"][role] = {
            **descriptor(retained),
            "features": [] if role == "native" else ["allocator-metrics"],
            "build_receipt": descriptor(PACKET / f"{role}-build.json"),
            "depfile": descriptor(PACKET / f"{role}.d"),
        }

    for name, expected in expected_hashes.items():
        require(sha256(SOURCE / name) == expected, f"candidate source changed during build: {name}")
    build.update({"status": "pass", "source_rechecked_files": len(expected_names)})
    write_once(PACKET / "build.json", build)
    print("0891 candidate release binaries retained")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
