#!/usr/bin/env python3
"""Prepare and remove the isolated ZIP fuzz-smoke workspace for change 0431.

This controller deliberately does not invoke Cargo or a fuzz target.  ``prepare``
copies only the ZIP crate source and the retained deterministic seed corpus into
an exclusive task directory.  The root agent owns the serialized Rust build and
libFuzzer invocation.  ``cleanup`` verifies the copied inputs, retains the
generated Cargo.lock bytes in the evidence bundle, records artifact hashes, and
removes only that exact task directory.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
CHECKS = ROOT / "checks"
TASK_ROOT = Path("/tmp/litchi-goal-0431-zip-fuzz-refined")
CORPUS_ROOT = TASK_ROOT / "corpus"
CUSTODY_PATH = CHECKS / "fuzz-refined-source-custody.json"
CLEANUP_PATH = CHECKS / "fuzz-refined-cleanup.json"
LOCK_EVIDENCE_PATH = CHECKS / "soapberry-zip-fuzz-refined-Cargo.lock.txt"
DRIVER_PATH = ROOT / "fuzz-control-refined.py"

EXPECTED_LOCK_RELATIVE = Path("crate/fuzz/Cargo.lock")

# These are the exact 0416 seed bytes.  Keeping the expected values here makes
# a later prepare fail closed if a retained seed was changed or replaced.
SEEDS = (
    ("many-small-central-local-signed.zip", 44130, "8352e19b15374d816f3845e8dae1f772cf083369e721dacaed7ec8ea852d803c"),
    ("many-small-central-local-zip32-tail-signed.zip", 44054, "8ee78ee40493d5f468b4f8520cb9de153ec3f365993ff0305864a2c6468929be"),
    ("many-small-local-only-signed.zip", 38934, "e77578623954dc19cf2c0966511596ce196178438a225e798afa5532b12e934c"),
    ("many-small-zip32-signed.zip", 31766, "7b8ba022a2ab76def4ae1758091859de3b5c1551664c84d27c85b800e407300a"),
    ("opc-local-only-signed.zip", 1348, "6c534b529d6546bac4f9c7d1e9a0a18991751101d0e6b1b1488547b1f8d2d51a"),
    ("zip32-signed-store-deflate.zip", 373, "e81bdc5cd5cbf6a1ae911abe8bca73c7a88947e99ceb01b0aa295d675f2ec4be"),
    ("zip64-central-local-signed.zip", 545, "d1e6f408641468355c2d0750c99094dec4f9c37135348b9e66452f642d9cf119"),
    ("zip64-central-local-zip32-tail-signed.zip", 469, "6b181748e6d19f30ee1a135c5be3895621314961ae9435bc05d98da47a75a953"),
    ("zip64-local-only-empty-signed.zip", 160, "ac285a6dfdb9b614f4fabd7fc1de04dc43141f792c66710eb7d5250cc60341ec"),
    ("zip64-local-only-empty-unsigned.zip", 156, "6c04fea48bd3f3daafd5931ac697365a0a3aa551258b14b3e4f3c96aa528b038"),
    ("zip64-local-only-seekable-no-descriptor.zip", 417, "d1366df6e357f7d5edea8127fd8c21f9ab4312c8413f3a171b48955b3c17561e"),
    ("zip64-local-only-signed.zip", 429, "f2312a2a777c66b381ee07ceadc27c34cc1135e9849d59b159fa1babeb56db69"),
    ("zip64-local-only-unsigned-crc-marker.zip", 424, "1cf1b52754f238ef3c864e2bf3340027aadc7e793ca308b18482a036ee12d670"),
)


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def file_record(path: Path, relative_to: Path | None = None) -> dict[str, object]:
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"expected a regular file: {path}")
    record: dict[str, object] = {
        "bytes": path.stat().st_size,
        "sha256": sha256(path),
    }
    if relative_to is not None:
        record["path"] = path.relative_to(relative_to).as_posix()
    return record


def assert_exact_task_root() -> None:
    expected = Path("/tmp/litchi-goal-0431-zip-fuzz-refined")
    if TASK_ROOT != expected or TASK_ROOT.parent != Path("/tmp"):
        raise RuntimeError(f"task root constant is not exact: {TASK_ROOT}")
    if TASK_ROOT.is_symlink():
        raise RuntimeError(f"task root must not be a symlink: {TASK_ROOT}")


def task_file(relative: str | Path) -> Path:
    relative_path = Path(relative)
    if relative_path.is_absolute() or ".." in relative_path.parts:
        raise RuntimeError(f"task-relative path escapes the task root: {relative}")
    return TASK_ROOT / relative_path


def write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def source_file_specs() -> list[tuple[Path, Path]]:
    """Return (repository-relative source, task-relative copy) pairs."""
    specs: list[tuple[Path, Path]] = [
        (Path("crates/soapberry-zip/Cargo.toml"), Path("crate/Cargo.toml")),
        (Path("crates/soapberry-zip/fuzz/Cargo.toml"), Path("crate/fuzz/Cargo.toml")),
    ]
    for source_root, copy_root in [
        (Path("crates/soapberry-zip/src"), Path("crate/src")),
        (Path("crates/soapberry-zip/fuzz/fuzz_targets"), Path("crate/fuzz/fuzz_targets")),
    ]:
        absolute_root = REPO / source_root
        if absolute_root.is_symlink() or not absolute_root.is_dir():
            raise RuntimeError(f"expected a regular source directory: {absolute_root}")
        for path in sorted(absolute_root.rglob("*")):
            if path.is_symlink() or not path.is_file():
                if path.is_symlink():
                    raise RuntimeError(f"source tree contains a symlink: {path}")
                continue
            specs.append((source_root / path.relative_to(absolute_root), copy_root / path.relative_to(absolute_root)))
    return sorted(specs, key=lambda pair: pair[0].as_posix())


def copy_exact(source: Path, destination: Path) -> None:
    if destination.exists() or destination.is_symlink():
        raise RuntimeError(f"destination already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)


def expected_seed_rows() -> list[tuple[Path, int, str]]:
    return [
        (Path("docs/performance/results/change-0416/corpus") / name, size, digest)
        for name, size, digest in SEEDS
    ]


def verify_expected_seed(source: Path, expected_bytes: int, expected_sha256: str) -> dict[str, object]:
    actual = file_record(source, REPO)
    if actual["bytes"] != expected_bytes or actual["sha256"] != expected_sha256:
        raise RuntimeError(
            f"retained seed changed: {source} "
            f"({actual['bytes']}/{actual['sha256']}, expected {expected_bytes}/{expected_sha256})"
        )
    return actual


def git_revision() -> str:
    return subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO, text=True).strip()


def driver_record() -> dict[str, object]:
    return {
        "path": DRIVER_PATH.relative_to(ROOT).as_posix(),
        "bytes": DRIVER_PATH.stat().st_size,
        "sha256": sha256(DRIVER_PATH),
    }


def prepare() -> int:
    assert_exact_task_root()
    if TASK_ROOT.exists():
        raise RuntimeError(f"exclusive task directory already exists: {TASK_ROOT}")
    for evidence in (CUSTODY_PATH, CLEANUP_PATH, LOCK_EVIDENCE_PATH):
        if evidence.exists() or evidence.is_symlink():
            raise RuntimeError(f"evidence path already exists: {evidence}")

    specs = source_file_specs()
    for source_relative, _ in specs:
        source = REPO / source_relative
        if source.is_symlink() or not source.is_file():
            raise RuntimeError(f"expected a regular source file: {source}")

    TASK_ROOT.mkdir(mode=0o700)
    source_files: list[dict[str, object]] = []
    for source_relative, copy_relative in specs:
        source = REPO / source_relative
        destination = task_file(copy_relative)
        copy_exact(source, destination)
        source_info = file_record(source, REPO)
        copy_info = file_record(destination, TASK_ROOT)
        if source_info["bytes"] != copy_info["bytes"] or source_info["sha256"] != copy_info["sha256"]:
            raise RuntimeError(f"source copy mismatch: {source} -> {destination}")
        source_files.append({
            "source": source_info,
            "copy": copy_info,
        })

    seed_files: list[dict[str, object]] = []
    for source_relative, expected_bytes, expected_sha256 in expected_seed_rows():
        source = REPO / source_relative
        source_info = verify_expected_seed(source, expected_bytes, expected_sha256)
        destination = task_file(Path("corpus") / source.name)
        copy_exact(source, destination)
        copy_info = file_record(destination, TASK_ROOT)
        if source_info["bytes"] != copy_info["bytes"] or source_info["sha256"] != copy_info["sha256"]:
            raise RuntimeError(f"seed copy mismatch: {source} -> {destination}")
        seed_files.append({
            "source": source_info,
            "copy": copy_info,
            "expected": {"bytes": expected_bytes, "sha256": expected_sha256},
        })

    artifacts = [
        {"kind": "source-copy", **row["copy"]} for row in source_files
    ] + [
        {"kind": "seed-copy", **row["copy"]} for row in seed_files
    ]
    custody = {
        "schema": 1,
        "change": 431,
        "status": "prepared",
        "prepared_utc": now(),
        "task_root": str(TASK_ROOT),
        "crate_manifest": "crate/fuzz/Cargo.toml",
        "fuzz_target": "parse_zip",
        "source_revision": git_revision(),
        "driver": driver_record(),
        "source_files": source_files,
        "seed_files": seed_files,
        "expected_generated_lock": EXPECTED_LOCK_RELATIVE.as_posix(),
        "cargo_lock_present_before_root_build": False,
        "artifact_hashes": artifacts,
    }
    write_json(CUSTODY_PATH, custody)
    print(json.dumps({
        "status": "pass",
        "action": "prepare",
        "task_root": str(TASK_ROOT),
        "crate_manifest": str(task_file("crate/fuzz/Cargo.toml")),
        "corpus": str(CORPUS_ROOT),
        "source_files": len(source_files),
        "seed_files": len(seed_files),
        "custody": str(CUSTODY_PATH.relative_to(ROOT)),
    }, sort_keys=True))
    return 0


def expect_recorded_file(path: Path, recorded: dict[str, object], label: str) -> dict[str, object]:
    actual = file_record(path)
    expected = {"bytes": recorded["bytes"], "sha256": recorded["sha256"]}
    if {"bytes": actual["bytes"], "sha256": actual["sha256"]} != expected:
        raise RuntimeError(f"{label} changed: {path}")
    return actual


def verify_baseline_files(custody: dict[str, object]) -> list[dict[str, object]]:
    artifacts: list[dict[str, object]] = []
    for kind, key in [("source-copy", "source_files"), ("seed-copy", "seed_files")]:
        rows = custody[key]
        if not isinstance(rows, list):
            raise RuntimeError(f"invalid custody rows: {key}")
        for row in rows:
            if not isinstance(row, dict) or not isinstance(row.get("source"), dict) or not isinstance(row.get("copy"), dict):
                raise RuntimeError(f"invalid custody record: {key}")
            source_info = row["source"]
            copy_info = row["copy"]
            source = REPO / str(source_info["path"])
            copied = task_file(str(copy_info["path"]))
            expect_recorded_file(source, source_info, f"source original {source}")
            expect_recorded_file(copied, copy_info, f"task copy {copied}")
            if source_info["bytes"] != copy_info["bytes"] or source_info["sha256"] != copy_info["sha256"]:
                raise RuntimeError(f"custody source/copy hashes disagree: {source}")
            artifacts.append({"kind": kind, **copy_info})
    return artifacts


def generated_artifacts(baseline_paths: set[str], lock_relative: str) -> list[dict[str, object]]:
    artifacts: list[dict[str, object]] = []
    for path in sorted(TASK_ROOT.rglob("*")):
        if path.is_symlink():
            raise RuntimeError(f"task directory contains an unexpected symlink: {path}")
        if not path.is_file():
            continue
        relative = path.relative_to(TASK_ROOT).as_posix()
        if relative in baseline_paths or relative == lock_relative:
            continue
        info = file_record(path, TASK_ROOT)
        artifacts.append({"kind": "generated", **info})
    return artifacts


def cleanup() -> int:
    assert_exact_task_root()
    if not CUSTODY_PATH.is_file() or CUSTODY_PATH.is_symlink():
        raise RuntimeError(f"missing custody evidence: {CUSTODY_PATH}")
    if CLEANUP_PATH.exists() or CLEANUP_PATH.is_symlink():
        raise RuntimeError(f"cleanup evidence already exists: {CLEANUP_PATH}")
    if LOCK_EVIDENCE_PATH.exists() or LOCK_EVIDENCE_PATH.is_symlink():
        raise RuntimeError(f"Cargo.lock evidence already exists: {LOCK_EVIDENCE_PATH}")
    custody = json.loads(CUSTODY_PATH.read_text())
    if custody.get("status") != "prepared":
        raise RuntimeError(f"unexpected custody status: {custody.get('status')!r}")
    if custody.get("task_root") != str(TASK_ROOT):
        raise RuntimeError("custody task root does not match the exact cleanup target")
    if custody.get("source_revision") != git_revision():
        raise RuntimeError("repository revision changed while the isolated fuzz task was active")
    driver = custody.get("driver")
    if not isinstance(driver, dict) or driver.get("sha256") != sha256(DRIVER_PATH):
        raise RuntimeError("fuzz controller changed after prepare")
    if not TASK_ROOT.is_dir():
        raise RuntimeError(f"missing isolated fuzz task directory: {TASK_ROOT}")

    baseline_artifacts = verify_baseline_files(custody)
    baseline_paths = {str(row["path"]) for row in baseline_artifacts}
    lock_relative = str(custody.get("expected_generated_lock"))
    if lock_relative != EXPECTED_LOCK_RELATIVE.as_posix():
        raise RuntimeError("unexpected generated Cargo.lock path")
    lock_path = task_file(lock_relative)
    if not lock_path.is_file() or lock_path.is_symlink():
        raise RuntimeError(f"generated Cargo.lock is missing: {lock_path}")
    lock_info = file_record(lock_path, TASK_ROOT)
    generated = generated_artifacts(baseline_paths, lock_relative)

    copy_exact(lock_path, LOCK_EVIDENCE_PATH)
    evidence_lock_info = file_record(LOCK_EVIDENCE_PATH, ROOT)
    if lock_info["bytes"] != evidence_lock_info["bytes"] or lock_info["sha256"] != evidence_lock_info["sha256"]:
        raise RuntimeError("retained Cargo.lock evidence does not match the generated lock")
    lock_record = {
        "generated": lock_info,
        "evidence": evidence_lock_info,
    }
    artifacts = baseline_artifacts + [{"kind": "generated-cargo-lock", **lock_info}] + generated

    custody["status"] = "cleanup_pending"
    custody["cleanup_prepared_utc"] = now()
    custody["cargo_lock"] = lock_record
    custody["generated_artifacts"] = generated
    custody["artifact_hashes"] = artifacts
    write_json(CUSTODY_PATH, custody)

    assert_exact_task_root()
    if not TASK_ROOT.is_dir() or TASK_ROOT.is_symlink():
        raise RuntimeError("exact task directory changed before cleanup")
    shutil.rmtree(TASK_ROOT)
    if TASK_ROOT.exists():
        raise RuntimeError(f"exact task directory remains after cleanup: {TASK_ROOT}")

    custody["status"] = "cleaned"
    custody["cleaned_utc"] = now()
    custody["cleanup"] = {
        "removed_directory": str(TASK_ROOT),
        "temporary_directory_absent": True,
        "baseline_copies_verified": True,
        "source_originals_verified": True,
        "cargo_lock_bytes_retained": True,
    }
    write_json(CUSTODY_PATH, custody)

    cleanup_record = {
        "schema": 1,
        "change": 431,
        "status": "pass",
        "cleaned_utc": custody["cleaned_utc"],
        "task_root": str(TASK_ROOT),
        "removed_directory": str(TASK_ROOT),
        "temporary_directory_absent": not TASK_ROOT.exists(),
        "baseline_copies_verified": True,
        "source_originals_verified": True,
        "cargo_lock": lock_record,
        "artifact_hashes": artifacts,
        "custody": {
            "path": str(CUSTODY_PATH.relative_to(ROOT)),
            "sha256": sha256(CUSTODY_PATH),
        },
        "driver": driver_record(),
    }
    write_json(CLEANUP_PATH, cleanup_record)
    print(json.dumps({
        "status": "pass",
        "action": "cleanup",
        "task_root": str(TASK_ROOT),
        "temporary_directory_absent": not TASK_ROOT.exists(),
        "cargo_lock_evidence": str(LOCK_EVIDENCE_PATH.relative_to(ROOT)),
        "custody": str(CUSTODY_PATH.relative_to(ROOT)),
        "cleanup": str(CLEANUP_PATH.relative_to(ROOT)),
    }, sort_keys=True))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("prepare", "cleanup"))
    args = parser.parse_args()
    try:
        return prepare() if args.command == "prepare" else cleanup()
    except (OSError, RuntimeError, subprocess.CalledProcessError, json.JSONDecodeError) as error:
        print(f"fuzz-control: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
