#!/usr/bin/env python3
"""Prepare and run the bounded ZIP/XML ASAN fuzz lanes for change 0457.

The runner owns only an isolated /tmp source tree.  It refuses an existing
staging or evidence tree, refuses source-snapshot drift, invokes the immutable
change-level check.py once for each lock/build/smoke stage, and never retries a
failed stage.  The shared Cargo target directory is intentionally preserved.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[4]
PARENT_CHECK = ROOT.parent / "check.py"
STAGE_ROOT = Path("/tmp/litchi-goal-0457")
TARGET_DIR = REPO / "target" / "fuzz-asan"
TARGET_TRIPLE = "x86_64-unknown-linux-gnu"
ASAN_FLAGS = (
    "-C passes=sancov-module "
    "-C llvm-args=-sanitizer-coverage-level=4 "
    "-C llvm-args=-sanitizer-coverage-inline-8bit-counters "
    "-C llvm-args=-sanitizer-coverage-pc-table "
    "-C llvm-args=-sanitizer-coverage-trace-compares "
    "-Z sanitizer=address --cfg fuzzing"
)
RUNS = 1000
SEED = 457
TIMEOUT = 10


TARGETS: dict[str, dict[str, Any]] = {
    "zip": {
        "manifest_snapshot": "source/soapberry-zip-Cargo.toml",
        "source_snapshot": "source/parse_zip.rs",
        "source_name": "parse_zip.rs",
        "binary_name": "parse_zip",
        "corpora": ("seeds/zip", "seeds/native"),
        "max_len": 1 << 20,
    },
    "xml": {
        "manifest_snapshot": "source/litchi-odf-common-Cargo.toml",
        "source_snapshot": "source/scan_xml.rs",
        "source_name": "scan_xml.rs",
        "binary_name": "scan_xml",
        "corpora": ("seeds/xml",),
        "max_len": 64 << 10,
    },
}


def sha_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def file_record(path: Path, relative_to: Path) -> dict[str, Any]:
    return {
        "path": path.relative_to(relative_to).as_posix(),
        "bytes": path.stat().st_size,
        "sha256": sha_file(path),
    }


def refuse_existing(path: Path) -> None:
    if path.exists():
        raise RuntimeError(f"refusing to overwrite existing path: {path}")


def adapted_manifest(source: Path, dependency: str, crate_target: Path) -> bytes:
    text = source.read_text()
    old = f'{dependency} = {{ path = ".." }}'
    new = f'{dependency} = {{ path = "{crate_target}" }}'
    if text.count(old) != 1:
        raise RuntimeError(f"unexpected dependency declaration in {source}")
    return text.replace(old, new).encode()


def load_manifest(name: str) -> dict[str, Any]:
    return json.loads((ROOT / name).read_text())


def verify_snapshots() -> None:
    source_manifest = load_manifest("source-manifest.json")
    seed_manifest = load_manifest("seed-manifest.json")
    if source_manifest["change"] != 457 or seed_manifest["change"] != 457:
        raise RuntimeError("snapshot change number mismatch")

    for item in source_manifest["source_rows"]:
        path = ROOT / item["path"]
        if not path.is_file():
            raise RuntimeError(f"missing source snapshot: {path}")
        actual = file_record(path, ROOT)
        if actual["bytes"] != item["bytes"] or actual["sha256"] != item["sha256"]:
            raise RuntimeError(f"source snapshot hash drift: {path}")

    current_zip = REPO / "crates/soapberry-zip/fuzz/Cargo.toml"
    expected_zip = adapted_manifest(
        current_zip,
        "soapberry-zip",
        REPO / "crates/soapberry-zip",
    )
    if expected_zip != (ROOT / "source/soapberry-zip-Cargo.toml").read_bytes():
        raise RuntimeError("soapberry-zip fuzz manifest drifted from the snapshot")
    current_xml = REPO / "crates/litchi-odf-common/fuzz/Cargo.toml"
    expected_xml = adapted_manifest(
        current_xml,
        "litchi-odf-common",
        REPO / "crates/litchi-odf-common",
    )
    if expected_xml != (ROOT / "source/litchi-odf-common-Cargo.toml").read_bytes():
        raise RuntimeError("litchi-odf-common fuzz manifest drifted from the snapshot")

    current_sources = {
        "source/parse_zip.rs": REPO / "crates/soapberry-zip/fuzz/fuzz_targets/parse_zip.rs",
        "source/scan_xml.rs": REPO / "crates/litchi-odf-common/fuzz/fuzz_targets/scan_xml.rs",
    }
    source_rows = {item["path"]: item for item in source_manifest["source_rows"]}
    for snapshot, current in current_sources.items():
        if sha_file(current) != source_rows[snapshot]["sha256"]:
            raise RuntimeError(f"fuzz target drifted from the archived snapshot: {current}")

    for item in seed_manifest["seed_rows"]:
        path = ROOT / item["path"]
        if not path.is_file():
            raise RuntimeError(f"missing seed: {path}")
        actual = file_record(path, ROOT)
        if actual["bytes"] != item["bytes"] or actual["sha256"] != item["sha256"]:
            raise RuntimeError(f"seed hash drift: {path}")


def stage_target(kind: str, config: dict[str, Any]) -> Path:
    stage = STAGE_ROOT / f"fuzz-{kind}"
    refuse_existing(stage)
    corpus = stage / "corpus"
    fuzz_targets = stage / "fuzz_targets"
    fuzz_targets.mkdir(parents=True)
    corpus.mkdir()
    shutil.copy2(ROOT / config["source_snapshot"], fuzz_targets / config["source_name"])
    shutil.copy2(ROOT / config["manifest_snapshot"], stage / "Cargo.toml")
    for corpus_snapshot in config["corpora"]:
        for seed in sorted((ROOT / corpus_snapshot).iterdir()):
            if not seed.is_file():
                continue
            destination = corpus / seed.name
            if destination.exists():
                raise RuntimeError(f"duplicate staged seed name: {destination}")
            shutil.copy2(seed, destination)
    return stage


def archive_pre_run(stages: dict[str, Path]) -> Path:
    archive_root = ROOT / "artifacts"
    refuse_existing(archive_root)
    archive_root.mkdir()
    pre_run = archive_root / "pre-run"
    pre_run.mkdir()
    for target_name in TARGETS:
        (pre_run / target_name).mkdir()

    for target_name, target in TARGETS.items():
        target_stage = stages[target_name]
        target_archive = pre_run / target_name
        source_archive = target_archive / "source"
        source_archive.mkdir()
        shutil.copy2(
            ROOT / target["source_snapshot"],
            source_archive / target["source_name"],
        )
        shutil.copy2(
            ROOT / target["manifest_snapshot"],
            source_archive / "Cargo.toml",
        )
        corpus_archive = target_archive / "corpus"
        corpus_archive.mkdir()
        for seed in sorted((target_stage / "corpus").iterdir()):
            shutil.copy2(seed, corpus_archive / seed.name)

    for snapshot in ("source-manifest.json", "seed-manifest.json", "SHA256SUMS"):
        shutil.copy2(ROOT / snapshot, pre_run / snapshot)

    inventory = {
        "schema": 1,
        "change": 457,
        "stage_root": str(STAGE_ROOT),
        "target_dir": str(TARGET_DIR),
        "parent_check": file_record(PARENT_CHECK, PARENT_CHECK.parent),
        "files": [],
    }
    for path in sorted((pre_run).rglob("*")):
        if path.is_file():
            inventory["files"].append(file_record(path, pre_run))
    (pre_run / "inventory-before-lock.json").write_text(
        json.dumps(inventory, indent=2) + "\n"
    )
    return pre_run


def stage_lock_archive(pre_run: Path, kind: str, stage: Path) -> None:
    target_archive = pre_run / kind
    lock = stage / "Cargo.lock"
    if not lock.is_file():
        raise RuntimeError(f"lockfile was not produced: {lock}")
    shutil.copy2(lock, target_archive / "Cargo.lock")
    inventory = json.loads((pre_run / "inventory-before-lock.json").read_text())
    inventory["files"] = [
        file_record(path, pre_run)
        for path in sorted(pre_run.rglob("*"))
        if path.is_file() and path.name != "inventory-before-lock.json"
    ]
    inventory["lock_target"] = kind
    (target_archive / "inventory-before-build.json").write_text(
        json.dumps(inventory, indent=2) + "\n"
    )


def parent_receipt_exists(tag: str) -> bool:
    checks = ROOT.parent / "checks"
    return any(
        (checks / suffix).exists()
        for suffix in (f"{tag}.json", f"{tag}.log", f"{tag}.log.gz")
    )


def run_stage(tag: str, argv: list[str]) -> None:
    if parent_receipt_exists(tag):
        raise RuntimeError(f"refusing to reuse parent check tag: {tag}")
    parent_hash = sha_file(PARENT_CHECK)
    subprocess.run(
        [sys.executable, "-B", str(PARENT_CHECK), "--tag", tag, "--", *argv],
        cwd=REPO,
        check=True,
    )
    if sha_file(PARENT_CHECK) != parent_hash:
        raise RuntimeError("parent check.py changed during fuzz staging")


def command_set(kind: str, stage: Path, config: dict[str, Any]) -> tuple[list[str], list[str], list[str]]:
    manifest = stage / "Cargo.toml"
    target_env = f"CARGO_TARGET_DIR={TARGET_DIR}"
    lock = [
        "env",
        target_env,
        "cargo",
        "generate-lockfile",
        "--offline",
        "--manifest-path",
        str(manifest),
    ]
    build = [
        "env",
        target_env,
        "RUSTC_BOOTSTRAP=1",
        f"RUSTFLAGS={ASAN_FLAGS}",
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        str(manifest),
        "--target",
        TARGET_TRIPLE,
        "--bin",
        config["binary_name"],
    ]
    binary = TARGET_DIR / TARGET_TRIPLE / "release" / config["binary_name"]
    smoke = [
        str(binary),
        str(stage / "corpus"),
        "-runs=1000",
        f"-seed={SEED}",
        f"-max_len={config['max_len']}",
        f"-timeout={TIMEOUT}",
        f"-artifact_prefix={stage}/",
    ]
    return lock, build, smoke


def archive_post_run(pre_run: Path, stages: dict[str, Path]) -> None:
    post_run = pre_run.parent / "post-run"
    post_run.mkdir()
    inventory: dict[str, Any] = {
        "schema": 1,
        "change": 457,
        "runs": RUNS,
        "seed": SEED,
        "timeout_seconds": TIMEOUT,
        "target_dir": str(TARGET_DIR),
        "files": [],
    }
    for kind, config in TARGETS.items():
        target_post = post_run / kind
        target_post.mkdir()
        binary = TARGET_DIR / TARGET_TRIPLE / "release" / config["binary_name"]
        if not binary.is_file():
            raise RuntimeError(f"fuzz binary missing after run: {binary}")
        shutil.copy2(binary, target_post / binary.name)
        checks = ROOT.parent / "checks"
        for stage_name in ("lock", "build", "smoke"):
            tag = f"fuzz-0457-{kind}-{stage_name}"
            for suffix in ("json", "log"):
                source = checks / f"{tag}.{suffix}"
                if source.is_file():
                    shutil.copy2(source, target_post / source.name)
        corpus = stages[kind] / "corpus"
        corpus_inventory = []
        for path in sorted(corpus.iterdir()):
            if path.is_file():
                corpus_inventory.append(file_record(path, corpus))
        (target_post / "corpus-inventory.json").write_text(
            json.dumps({"kind": kind, "files": corpus_inventory}, indent=2) + "\n"
        )
        for path in sorted(target_post.rglob("*")):
            if path.is_file():
                inventory["files"].append(file_record(path, post_run))
    (post_run / "inventory.json").write_text(json.dumps(inventory, indent=2) + "\n")


def main() -> int:
    verify_snapshots()
    refuse_existing(STAGE_ROOT / "fuzz-zip")
    refuse_existing(STAGE_ROOT / "fuzz-xml")
    refuse_existing(ROOT / "artifacts")
    for kind in TARGETS:
        for phase in ("lock", "build", "smoke"):
            if parent_receipt_exists(f"fuzz-0457-{kind}-{phase}"):
                raise RuntimeError(f"existing parent check evidence for {kind}/{phase}")

    STAGE_ROOT.mkdir(exist_ok=True)
    stages = {kind: stage_target(kind, config) for kind, config in TARGETS.items()}
    pre_run = archive_pre_run(stages)
    for kind, config in TARGETS.items():
        lock, build, smoke = command_set(kind, stages[kind], config)
        run_stage(f"fuzz-0457-{kind}-lock", lock)
        stage_lock_archive(pre_run, kind, stages[kind])
        run_stage(f"fuzz-0457-{kind}-build", build)
        run_stage(f"fuzz-0457-{kind}-smoke", smoke)
    archive_post_run(pre_run, stages)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
