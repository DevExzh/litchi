#!/usr/bin/env python3
"""Freeze and build the 0894 selected-metadata candidate.

The baseline archive is immutable.  ``freeze`` copies the five reviewed
candidate files into a packet-local overlay, records their hashes and a
baseline diff, and refuses to replace any evidence.  ``build`` reconstructs a
new source root from that archive plus the frozen overlay and builds native
and allocator release binaries with the exact baseline configuration.

This helper deliberately has no workload or profiler mode.  The protocol
adapter consumes the resulting candidate-build receipt later.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
from pathlib import Path
import shutil
import stat
import subprocess
import tarfile
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
ARCHIVE = PACKET / "source.tar.gz"
MANIFEST = PACKET / "source.json"
FREEZE_START = PACKET / "candidate-freeze-start.json"
FREEZE = PACKET / "candidate-freeze.json"
PATCH = PACKET / "candidate.patch"
OVERLAY = PACKET / "candidate-overlay"
SOURCE = ROOT / "target/litchi-0894-candidate-src"
TARGET = ROOT / "target/litchi-0894-candidate-build"
BINARIES = PACKET / "binaries"
BUILD_START = PACKET / "candidate-build-start.json"
BUILD = PACKET / "candidate-build.json"
SOURCE_BEFORE = PACKET / "candidate-source-before.json"
SOURCE_AFTER = PACKET / "candidate-source-after.json"
ROOT_LOCK = ROOT / "Cargo.lock"
REVISION = "506aa2edae6666d9c76ae6827c2f5941a8e47896"
GUARD_CONTRACT = "post-endpoint-per-worksheet-tolerance-v1"
BASELINE_FILE_COUNT = 10_771
MARKER = ".owner-litchi-0894-candidate"
SOURCE_MARKER_TEXT = "litchi 0894 frozen candidate source\n"
TARGET_MARKER_TEXT = "litchi 0894 frozen candidate build target\n"
CANDIDATE_PATHS = (
    "crates/litchi-xlsx/src/raw/worksheet/selected.rs",
    "crates/litchi-xlsx/src/raw/worksheet/selected/producer_metadata.rs",
    "crates/litchi-xlsx/src/raw/worksheet/selected/producer_metadata_tests.rs",
    "tools/perf-baseline/src/producer_shape/selected_route_tests.rs",
    "crates/litchi-xlsx/src/raw/worksheet/tests.rs",
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
UNRELATED_ALLOWED = frozenset({"crates/litchi/examples/bench_word_parse.rs"})
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
    require(not path.is_symlink(), f"expected regular non-symlink file: {path}")
    path = path.resolve()
    info = path.stat()
    require(stat.S_ISREG(info.st_mode), f"expected regular file: {path}")
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


def git(*args: str) -> str:
    return subprocess.check_output(["git", *args], cwd=ROOT, text=True).strip()


def read_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def load_baseline() -> dict[str, Any]:
    require(ARCHIVE.is_file() and MANIFEST.is_file(), "baseline source archive/manifest is missing")
    manifest = read_json(MANIFEST)
    require(manifest.get("schema") == "litchi.xlsx-source-snapshot-0894.v1", "baseline source schema mismatch")
    require(manifest.get("status") == "pass", "baseline source receipt does not pass")
    require(manifest.get("revision") == REVISION and manifest.get("baseline_commit") == REVISION, "baseline source revision mismatch")
    require(manifest.get("file_count") == BASELINE_FILE_COUNT, "baseline source file count mismatch")
    archive = manifest.get("archive")
    require(isinstance(archive, dict), "baseline source archive descriptor is missing")
    require(archive.get("path") == str(ARCHIVE.resolve()), "baseline archive path mismatch")
    require(archive.get("bytes") == ARCHIVE.stat().st_size and archive.get("sha256") == sha256(ARCHIVE), "baseline archive changed")
    standalone = manifest.get("standalone_lock")
    require(isinstance(standalone, dict) and standalone.get("path") == "tools/perf-baseline/Cargo.lock", "standalone lock identity is missing")
    require(standalone.get("sha256") == manifest.get("files", {}).get("tools/perf-baseline/Cargo.lock", {}).get("sha256"), "standalone lock does not match source census")
    root_lock = manifest.get("root_generated_lock")
    require(isinstance(root_lock, dict), "root generated lock identity is missing")
    current_root_lock = descriptor(ROOT_LOCK)
    require(root_lock.get("path") == current_root_lock["path"], "root generated lock path changed")
    require(root_lock.get("bytes") == current_root_lock["bytes"] and root_lock.get("sha256") == current_root_lock["sha256"], "root generated lock changed")
    return manifest


def changed_paths() -> tuple[set[str], set[str]]:
    tracked = {path for path in git("diff", "--name-only", REVISION, "--", *CANDIDATE_PATHS).splitlines() if path}
    untracked = {path for path in git("ls-files", "--others", "--exclude-standard", "--", *CANDIDATE_PATHS).splitlines() if path}
    require(tracked | untracked <= set(CANDIDATE_PATHS), "candidate path census has an unexpected path")
    require(tracked | untracked == set(CANDIDATE_PATHS), "all five candidate paths must be changed or added")
    all_tracked = {path for path in git("diff", "--name-only", REVISION, "--", *SELECTION).splitlines() if path}
    require(all_tracked <= set(CANDIDATE_PATHS), "tracked selected source changed outside the frozen candidate paths")
    selected_untracked = {path for path in git("ls-files", "--others", "--exclude-standard", "--", *SELECTION).splitlines() if path}
    require(selected_untracked - set(CANDIDATE_PATHS) <= UNRELATED_ALLOWED, "untracked selected source is outside candidate/allowed benchmark")
    return tracked, untracked


def source_path_descriptor(name: str) -> dict[str, Any]:
    path = ROOT / name
    require(path.is_file() and not path.is_symlink(), f"candidate source file is missing: {name}")
    item = descriptor(path)
    item["path"] = name
    return item


def overlay_descriptor(name: str) -> dict[str, Any]:
    path = OVERLAY / name
    item = descriptor(path)
    item["path"] = name
    return item


def patch_bytes() -> bytes:
    tracked, untracked = changed_paths()
    tracked_paths = [name for name in CANDIDATE_PATHS if name in tracked]
    result = bytearray()
    if tracked_paths:
        result.extend(
            subprocess.check_output(["git", "diff", "--binary", "--full-index", REVISION, "--", *tracked_paths], cwd=ROOT)
        )
    for name in CANDIDATE_PATHS:
        if name not in untracked:
            continue
        process = subprocess.run(
            ["git", "diff", "--no-index", "--binary", "--", "/dev/null", name],
            cwd=ROOT,
            capture_output=True,
            check=False,
        )
        require(process.returncode in (0, 1), f"cannot create untracked candidate diff: {name}")
        result.extend(process.stdout)
    require(result, "candidate patch is empty")
    return bytes(result)


def freeze() -> int:
    for path in (FREEZE_START, FREEZE, PATCH, OVERLAY):
        require(not path.exists(), f"candidate freeze output already exists: {path}")
    require(git("rev-parse", "HEAD") == REVISION, "candidate freeze requires exact 0894 baseline HEAD")
    manifest = load_baseline()
    tracked, untracked = changed_paths()
    current_root_lock = descriptor(ROOT_LOCK)
    start = {
        "schema": "litchi.xlsx-selected-metadata-candidate-freeze-start-0894.v1",
        "status": "building",
        "started_utc": now(),
        "base_commit": REVISION,
        "candidate_paths": list(CANDIDATE_PATHS),
        "tracked_changed": sorted(tracked),
        "untracked_added": sorted(untracked),
        "root_generated_lock": current_root_lock,
        "source_manifest": descriptor(MANIFEST),
        "source_archive": descriptor(ARCHIVE),
    }
    write_once(FREEZE_START, start)
    OVERLAY.mkdir()
    files: dict[str, dict[str, Any]] = {}
    for name in CANDIDATE_PATHS:
        source = ROOT / name
        target = OVERLAY / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        current = source_path_descriptor(name)
        copied = overlay_descriptor(name)
        require(current["bytes"] == copied["bytes"] and current["sha256"] == copied["sha256"], f"overlay copy changed: {name}")
        baseline = manifest["files"].get(name)
        files[name] = {
            "baseline": baseline,
            "current": current,
            "overlay": copied,
            "replacement": baseline is not None,
        }

    patch = patch_bytes()
    # Re-read the live files after constructing the patch.  The overlay and
    # diff must describe one stable worktree observation, even when another
    # process is editing a candidate file concurrently.
    for name in CANDIDATE_PATHS:
        current = source_path_descriptor(name)
        frozen = files[name]["current"]
        require(current["bytes"] == frozen["bytes"] and current["sha256"] == frozen["sha256"], f"candidate changed while freezing: {name}")
        copied = overlay_descriptor(name)
        require(copied["bytes"] == frozen["bytes"] and copied["sha256"] == frozen["sha256"], f"overlay changed while freezing: {name}")
    with PATCH.open("xb") as handle:
        handle.write(patch)
    freeze_receipt = {
        "schema": "litchi.xlsx-selected-metadata-candidate-freeze-0894.v1",
        "status": "frozen",
        "frozen_utc": now(),
        "base_commit": REVISION,
        "baseline_commit": REVISION,
        "timed_harness_guard_contract": GUARD_CONTRACT,
        "source_archive": descriptor(ARCHIVE),
        "source_manifest": descriptor(MANIFEST),
        "candidate_paths": list(CANDIDATE_PATHS),
        "files": files,
        "replacement_files": sum(1 for item in files.values() if item["replacement"]),
        "added_files": sum(1 for item in files.values() if not item["replacement"]),
        "baseline_file_count": BASELINE_FILE_COUNT,
        "candidate_file_count": BASELINE_FILE_COUNT + sum(1 for item in files.values() if not item["replacement"]),
        "candidate_patch": descriptor(PATCH),
        "candidate_overlay_root": str(OVERLAY.resolve()),
        "root_generated_lock": current_root_lock,
        "production_changes": True,
        "performance_claim": None,
    }
    write_once(FREEZE, freeze_receipt)
    print(f"candidate freeze sealed for {len(CANDIDATE_PATHS)} files", flush=True)
    return 0


def load_build_helpers() -> Any:
    path = PACKET / "build.py"
    spec = importlib.util.spec_from_file_location("litchi_0894_build_helper", path)
    require(spec is not None and spec.loader is not None, "cannot load baseline build helper")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.SOURCE = SOURCE
    module.TARGET = TARGET
    module.MARKER = MARKER
    module.SOURCE_MARKER_TEXT = SOURCE_MARKER_TEXT
    module.TARGET_MARKER_TEXT = TARGET_MARKER_TEXT
    module.BINARIES = BINARIES
    return module


def candidate_files(manifest: dict[str, Any], freeze_receipt: dict[str, Any]) -> dict[str, dict[str, Any]]:
    base_files = manifest["files"]
    result = dict(base_files)
    for name, item in freeze_receipt["files"].items():
        result[name] = item["current"]
    return result


def validate_freeze(manifest: dict[str, Any]) -> dict[str, Any]:
    require(FREEZE.is_file() and PATCH.is_file() and OVERLAY.is_dir(), "candidate freeze outputs are missing")
    receipt = read_json(FREEZE)
    require(receipt.get("schema") == "litchi.xlsx-selected-metadata-candidate-freeze-0894.v1", "candidate freeze schema mismatch")
    require(receipt.get("status") == "frozen" and receipt.get("base_commit") == REVISION, "candidate freeze does not pass")
    require(receipt.get("candidate_paths") == list(CANDIDATE_PATHS), "candidate path order changed")
    require(receipt.get("candidate_patch", {}).get("sha256") == sha256(PATCH), "candidate patch changed")
    require(receipt.get("candidate_patch", {}).get("bytes") == PATCH.stat().st_size, "candidate patch size changed")
    require(receipt.get("source_archive", {}).get("sha256") == manifest["archive"]["sha256"], "candidate freeze archive mismatch")
    require(receipt.get("source_manifest", {}).get("sha256") == sha256(MANIFEST), "candidate freeze manifest changed")
    for name in CANDIDATE_PATHS:
        item = receipt.get("files", {}).get(name)
        require(isinstance(item, dict), f"candidate freeze file missing: {name}")
        overlay = OVERLAY / name
        actual = descriptor(overlay)
        expected = item.get("overlay", {})
        require(expected.get("bytes") == actual["bytes"] and expected.get("sha256") == actual["sha256"], f"candidate overlay changed: {name}")
        current = item.get("current", {})
        require(current.get("bytes") == expected.get("bytes") and current.get("sha256") == expected.get("sha256"), f"candidate freeze identity mismatch: {name}")
    return receipt


def archive_members() -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    with tarfile.open(ARCHIVE, "r:gz") as handle:
        for member in handle:
            path = Path(member.name)
            require(not path.is_absolute() and ".." not in path.parts, f"unsafe baseline archive path: {member.name}")
            if member.isdir():
                continue
            require(member.isfile(), f"baseline archive contains non-regular member: {member.name}")
            require(member.name not in result, f"duplicate baseline archive member: {member.name}")
            source = handle.extractfile(member)
            require(source is not None, f"cannot read baseline archive member: {member.name}")
            digest = hashlib.sha256()
            size = 0
            with source:
                for block in iter(lambda: source.read(1024 * 1024), b""):
                    digest.update(block)
                    size += len(block)
            result[member.name] = {"bytes": size, "sha256": digest.hexdigest(), "mode": stat.S_IMODE(member.mode)}
    return result


def prepare_source(manifest: dict[str, Any], freeze_receipt: dict[str, Any], helper: Any) -> dict[str, dict[str, Any]]:
    members = archive_members()
    require(set(members) == set(manifest["files"]), "baseline archive census changed")
    for name, expected in manifest["files"].items():
        require(all(members[name][key] == expected[key] for key in ("bytes", "sha256", "mode")), f"baseline archive identity changed: {name}")
    with tarfile.open(ARCHIVE, "r:gz") as handle:
        handle.extractall(SOURCE, filter="data")
    for name in CANDIDATE_PATHS:
        target = SOURCE / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(OVERLAY / name, target)
    expected = candidate_files(manifest, freeze_receipt)
    observed = helper.source_snapshot()
    require(set(observed) == set(expected), f"candidate source census mismatch: expected {len(expected)}, got {len(observed)}")
    for name, record in expected.items():
        require(all(observed[name][key] == record[key] for key in ("bytes", "sha256", "mode")), f"candidate source identity mismatch: {name}")
    return expected


def build_environment(helper: Any) -> dict[str, str]:
    # helper.build_environment enforces the whitelist, private TMPDIR and
    # compiler override rejection after its SOURCE/TARGET globals are bound.
    return helper.build_environment()


def run_build(manifest: dict[str, Any], freeze_receipt: dict[str, Any]) -> int:
    outputs = (BUILD_START, BUILD, SOURCE, TARGET, SOURCE_BEFORE, SOURCE_AFTER)
    for path in outputs:
        require(not path.exists(), f"candidate build output already exists: {path}")
    require(BINARIES.is_dir(), "baseline binaries directory must exist before candidate build")
    for name in ("candidate-native", "candidate-allocator"):
        require(not (BINARIES / name).exists(), f"candidate binary already exists: {BINARIES / name}")
    require(git("rev-parse", "HEAD") == REVISION, "candidate build requires exact baseline HEAD")
    helper = load_build_helpers()
    SOURCE.mkdir(parents=True)
    TARGET.mkdir(parents=True)
    (SOURCE / MARKER).write_text(SOURCE_MARKER_TEXT, encoding="utf-8")
    (TARGET / MARKER).write_text(TARGET_MARKER_TEXT, encoding="utf-8")
    expected = prepare_source(manifest, freeze_receipt, helper)
    environment = build_environment(helper)
    root_lock_before = descriptor(ROOT_LOCK)
    before = {
        "schema": "litchi.xlsx-candidate-source-before-0894.v1",
        "status": "pass",
        "created_utc": now(),
        "source_root": str(SOURCE.resolve()),
        "file_count": len(expected),
        "files": helper.source_snapshot(),
        "root_generated_lock": root_lock_before,
        "environment": {key: environment[key] for key in ("TMPDIR", "TEMP", "TMP")},
    }
    write_once(SOURCE_BEFORE, before)
    start = {
        "schema": "litchi.xlsx-selected-metadata-candidate-build-start-0894.v1",
        "status": "building",
        "started_utc": now(),
        "base_commit": REVISION,
        "baseline_commit": REVISION,
        "source_revision": REVISION,
        "profile": "release",
        "guard_contract": GUARD_CONTRACT,
        "timed_harness_guard_contract": GUARD_CONTRACT,
        "flags": BUILD_FLAGS,
        "source_manifest": descriptor(MANIFEST),
        "source_archive": descriptor(ARCHIVE),
        "candidate_freeze": descriptor(FREEZE),
        "candidate_patch": descriptor(PATCH),
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
    write_once(BUILD_START, start)
    failure: dict[str, Any] | None = None
    try:
        for role, binary, features in ROLES:
            role_start_path = PACKET / f"candidate-{role}-build-start.json"
            role_receipt_path = PACKET / f"candidate-{role}-build.json"
            log_path = PACKET / f"candidate-{role}-build.log"
            require(not role_start_path.exists() and not role_receipt_path.exists() and not log_path.exists(), f"candidate {role} output already exists")
            argv = helper.command(role, binary, features)
            role_start = {
                "schema": f"litchi.xlsx-candidate-{role}-build-start-0894.v1",
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
                "schema": f"litchi.xlsx-candidate-{role}-build-0894.v1",
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
                failure = {"role": role, "exit_code": completed.returncode}
                raise RuntimeError(f"candidate {role} release build failed with exit code {completed.returncode}")
            executable = TARGET / "release" / binary
            depfile = TARGET / "release" / f"{binary}.d"
            require(executable.is_file() and executable.stat().st_mode & stat.S_IXUSR, f"candidate executable missing: {executable}")
            require(depfile.is_file(), f"candidate depfile missing: {depfile}")
            retained = BINARIES / f"candidate-{role}"
            shutil.copy2(executable, retained)
            retained.chmod(retained.stat().st_mode | stat.S_IXUSR)
            retained_depfile = PACKET / f"candidate-{role}.d"
            shutil.copy2(depfile, retained_depfile)
            start["binaries"][role] = {
                "role": role,
                "name": binary,
                "features": list(features),
                **helper.descriptor(retained),
                "executable": True,
                "profile": "release",
                "elf_build_id": helper.elf_build_id(retained, environment),
                "build_receipt": helper.descriptor(role_receipt_path),
                "depfile": helper.descriptor(retained_depfile),
            }
        root_lock_after = descriptor(ROOT_LOCK)
        require(root_lock_after == root_lock_before, "root Cargo.lock changed during candidate release builds")
        after_files = helper.source_snapshot()
        require(after_files == before["files"], "candidate frozen source changed during release builds")
        after = {
            "schema": "litchi.xlsx-candidate-source-after-0894.v1",
            "status": "pass",
            "created_utc": now(),
            "source_root": str(SOURCE.resolve()),
            "file_count": len(after_files),
            "files": after_files,
            "root_generated_lock": root_lock_after,
            "environment": {key: environment[key] for key in ("TMPDIR", "TEMP", "TMP")},
        }
        write_once(SOURCE_AFTER, after)
        start.update({
            "status": "pass",
            "finished_utc": now(),
            "source_before": descriptor(SOURCE_BEFORE),
            "source_after": descriptor(SOURCE_AFTER),
        })
        final = {
            **start,
            "schema": "litchi.xlsx-selected-metadata-candidate-build-0894.v1",
            "status": "pass",
            "build_start": descriptor(BUILD_START),
            "candidate_overlay": descriptor(PATCH),
            "source_rechecked_files": len(after_files),
            "performance_claim": None,
        }
        write_once(BUILD, final)
        print("candidate native and allocator binaries retained", flush=True)
        return 0
    except Exception as exc:
        failure_receipt = {
            "schema": "litchi.xlsx-selected-metadata-candidate-build-failure-0894.v1",
            "status": "fail",
            "finished_utc": now(),
            "error": str(exc),
            "role": failure.get("role") if failure else None,
            "exit_code": failure.get("exit_code") if failure else None,
            "source_before": descriptor(SOURCE_BEFORE) if SOURCE_BEFORE.is_file() else None,
        }
        failure_path = PACKET / "candidate-build-failure.json"
        if not failure_path.exists():
            write_once(failure_path, failure_receipt)
        raise


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("freeze", "build"))
    args = parser.parse_args()
    if args.action == "freeze":
        return freeze()
    manifest = load_baseline()
    freeze_receipt = validate_freeze(manifest)
    return run_build(manifest, freeze_receipt)


if __name__ == "__main__":
    raise SystemExit(main())
