#!/usr/bin/env python3
"""Prepare, build, run, and verify the bounded 0497 DOCX fuzz lane.

The helper follows the retained 0485 stream-fuzz protocol, but its Cargo
package is standalone.  It copies the unchanged target into a dedicated
temporary package, points the package at the authenticated 0497 ``after``
checkout, and uses a private Cargo target directory.  The four commands are
deliberately independent so the coordinator can put the build and two fuzz
runs under the normal gate/CPU-lock machinery::

    python3 -B docs/performance/results/change-0497/fuzz.py prepare
    python3 -B docs/performance/results/change-0497/fuzz.py build
    python3 -B docs/performance/results/change-0497/fuzz.py run
    python3 -B docs/performance/results/change-0497/fuzz.py verify

The script does not delete old evidence or clean the temporary tree.  A
failed command leaves its inputs and raw terminal output available for review.
Only the coordinator invokes ``build`` and ``run``; merely importing or
checking this file never invokes Cargo or the fuzz binary.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any, Iterable, NoReturn


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
EVIDENCE = ROOT / "fuzz"

# Keep every generated build/run file below this private subtree.  In
# particular, this must never be the repository's target/ directory: the
# normal Rust lanes and this sanitizer lane have different compiler flags.
TEMP_ROOT = Path("/home/zhuhe/.cache/litchi-goal-0497/fuzz-target")
PACKAGE = TEMP_ROOT / "fuzz"
CARGO_TARGET = TEMP_ROOT / "target"
RUN_ROOT = TEMP_ROOT / "runs"
TMPDIR = TEMP_ROOT / "tmp"

AFTER_ROOT = Path("/home/zhuhe/.cache/litchi-goal-0497/after")
AFTER_CRATE = AFTER_ROOT / "crates/litchi-docx"
AFTER_TARGET = AFTER_CRATE / "fuzz/fuzz_targets/source_backed_tail_append_stream.rs"

TARGET_SOURCE = REPO / "crates/litchi-docx/fuzz/fuzz_targets/source_backed_tail_append_stream.rs"
PRIMARY_LOCK = REPO / "crates/litchi-docx/fuzz/Cargo.lock"
RETAINED_0485 = REPO / "docs/performance/results/change-0485/fuzz-stream/candidate1"
RETAINED_DRIVER = REPO / "docs/performance/results/change-0485/fuzz-stream.py"
RETAINED_SEEDS = RETAINED_0485 / "seeds"
RETAINED_GENERATOR = RETAINED_0485 / "generator.json"
RETAINED_MANIFEST = RETAINED_0485 / "seed-manifest.json"

BINARY_NAME = "source_backed_tail_append_stream"
RETAINED_BINARY = Path("/home/zhuhe/.cache/litchi-goal-0497/retained/fuzz") / BINARY_NAME
TARGET_TRIPLE = "x86_64-unknown-linux-gnu"
INPUT_CAP = 65_536
FUZZ_RUNS = 10_000
RUN_SEEDS = (497, 498)
TOOLCHAIN = "1.98.1"

# This is the exact sanitizer/coverage string recorded by the accepted 0485
# build receipt.  Keep it as one string because RUSTFLAGS is itself one
# environment value, and record it verbatim in build.json.
ASAN_FLAGS = (
    "-C passes=sancov-module -C llvm-args=-sanitizer-coverage-level=4 "
    "-C llvm-args=-sanitizer-coverage-inline-8bit-counters "
    "-C llvm-args=-sanitizer-coverage-pc-table "
    "-C llvm-args=-sanitizer-coverage-trace-compares -Z sanitizer=address --cfg fuzzing"
)

BASE_ENV = {
    "RUSTUP_TOOLCHAIN": TOOLCHAIN,
    "CARGO_BUILD_JOBS": "4",
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "DEBUGINFOD_URLS": "",
    "LC_ALL": "C",
    "PYTHONDONTWRITEBYTECODE": "1",
}


def fail(message: str) -> NoReturn:
    raise RuntimeError(message)


def now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def sha(path: Path) -> str:
    if path.is_symlink() or not path.is_file():
        fail(f"expected a regular file: {path}")
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def meta(path: Path) -> dict[str, int | str]:
    if path.is_symlink() or not path.is_file():
        fail(f"expected a regular file: {path}")
    return {"bytes": path.stat().st_size, "sha256": sha(path)}


def inventory(directory: Path) -> dict[str, dict[str, int | str]]:
    if not directory.is_dir() or directory.is_symlink():
        fail(f"expected a regular directory: {directory}")
    result: dict[str, dict[str, int | str]] = {}
    for path in sorted(directory.rglob("*")):
        if path.is_symlink():
            fail(f"symlinks are not accepted in fuzz custody: {path}")
        if path.is_file():
            result[path.relative_to(directory).as_posix()] = meta(path)
        elif not path.is_dir():
            fail(f"special file is not accepted in fuzz custody: {path}")
    return result


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() or path.is_symlink():
        fail(f"refusing to replace retained evidence: {path}")
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
        stream.write("\n")


def read_json(path: Path) -> Any:
    if path.is_symlink() or not path.is_file():
        fail(f"missing JSON receipt: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"), parse_constant=reject_constant)
    except (json.JSONDecodeError, ValueError) as exc:
        fail(f"invalid JSON receipt {path}: {exc}")


def reject_constant(token: str) -> NoReturn:
    raise ValueError(f"non-finite JSON constant {token}")


def validate_timestamp(value: Any, path: str) -> None:
    if not isinstance(value, str):
        fail(f"{path}: timestamp is not text")
    try:
        parsed = _datetime.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        fail(f"{path}: malformed timestamp: {exc}")
    if parsed.tzinfo is None:
        fail(f"{path}: timestamp has no timezone")
    if parsed.utcoffset() != _datetime.timedelta(0):
        fail(f"{path}: timestamp is not UTC")


def write_or_check_verification(path: Path, payload: dict[str, Any]) -> bool:
    """Create the final receipt once, then make later verification read-only."""

    if not path.exists() and not path.is_symlink():
        write_json(path, payload)
        return False
    if path.is_symlink() or not path.is_file():
        fail(f"verification receipt is not a regular file: {path}")
    existing = read_json(path)
    if not isinstance(existing, dict):
        fail("verification receipt is not an object")
    if existing.get("schema") != payload.get("schema"):
        fail("verification receipt schema differs")
    validate_timestamp(existing.get("verified_utc"), "verify.json.verified_utc")
    existing_stable = dict(existing)
    existing_stable.pop("verified_utc", None)
    expected_stable = dict(payload)
    expected_stable.pop("verified_utc", None)
    if existing_stable != expected_stable:
        fail("retained verification receipt differs from the current verification")
    return True


def copy_file(source: Path, destination: Path) -> None:
    source_meta = meta(source)
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists() or destination.is_symlink():
        fail(f"refusing to replace retained file: {destination}")
    shutil.copyfile(source, destination)
    if meta(destination) != source_meta:
        fail(f"copied file differs: {source} -> {destination}")


def copy_directory(source: Path, destination: Path) -> None:
    if not source.is_dir() or source.is_symlink():
        fail(f"expected source directory: {source}")
    if destination.exists() or destination.is_symlink():
        fail(f"refusing to replace retained directory: {destination}")
    destination.mkdir(parents=True)
    for child in sorted(source.iterdir()):
        if child.is_symlink():
            fail(f"source corpus contains a symlink: {child}")
        target = destination / child.name
        if child.is_dir():
            copy_directory(child, target)
        elif child.is_file():
            copy_file(child, target)
        else:
            fail(f"source corpus contains a special file: {child}")


def copy_corpus(source: Path, destination: Path) -> None:
    copy_directory(source, destination)


def canonical_digest(value: Any) -> str:
    encoded = json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False).encode()
    return hashlib.sha256(encoded).hexdigest()


def git_output(directory: Path, *arguments: str) -> str:
    try:
        output = subprocess.check_output(
            ["git", "-C", str(directory), *arguments],
            cwd=REPO,
            stderr=subprocess.STDOUT,
        )
    except (OSError, subprocess.CalledProcessError) as exc:
        fail(f"git identity query failed for {directory}: {exc}")
    return output.decode("utf-8").strip()


def after_tree_binding() -> dict[str, Any]:
    """Hash the standalone dependency and its workspace lock identity."""

    if not AFTER_ROOT.is_dir() or AFTER_ROOT.is_symlink():
        fail(f"after checkout is missing: {AFTER_ROOT}")
    if not AFTER_CRATE.is_dir() or AFTER_CRATE.is_symlink():
        fail(f"after DOCX crate is missing: {AFTER_CRATE}")
    # The coordinator may have staged the production candidate in this
    # isolated checkout without committing it yet.  Bind that exact status
    # and the full DOCX crate inventory instead of silently assuming HEAD is
    # the build input; any subsequent edit then fails the custody comparison.
    status = git_output(AFTER_ROOT, "status", "--porcelain")

    files: dict[str, dict[str, int | str]] = {}
    for path in sorted(AFTER_CRATE.rglob("*")):
        if path.is_symlink():
            fail(f"after DOCX crate contains a symlink: {path}")
        if path.is_file():
            files[path.relative_to(AFTER_ROOT).as_posix()] = meta(path)
        elif not path.is_dir():
            fail(f"after DOCX crate contains a special file: {path}")
    return {
        "git_head": git_output(AFTER_ROOT, "rev-parse", "HEAD"),
        "git_status": status,
        "crate_files": len(files),
        "crate_inventory_sha256": canonical_digest(files),
        "workspace_manifest": meta(AFTER_ROOT / "Cargo.toml"),
        "workspace_lock": meta(AFTER_ROOT / "Cargo.lock"),
        "crate_inventory": files,
    }


def current_git_head() -> str:
    return git_output(REPO, "rev-parse", "HEAD")


def source_binding() -> dict[str, Any]:
    binding = {
        "current_git_head": current_git_head(),
        "target_source_current": meta(TARGET_SOURCE),
        "target_source_after": meta(AFTER_TARGET),
        "after": after_tree_binding(),
    }
    if binding["target_source_current"] != binding["target_source_after"]:
        fail("current and after fuzz target sources differ")
    return binding


def authenticate_retained_0485(root: Path = RETAINED_0485) -> dict[str, Any]:
    """Authenticate the immutable 0485 90-seed source before copying it."""

    seeds_root = root / "seeds"
    generator_path = root / "generator.json"
    manifest_path = root / "seed-manifest.json"
    driver_path = root / "fuzz-stream.py"
    # The historical bundle keeps its driver beside candidate1; the
    # evidence-side copy keeps it inside retained-0485.
    if not driver_path.is_file() and root == RETAINED_0485:
        driver_path = root.parent.parent / "fuzz-stream.py"

    for path in (driver_path, generator_path, manifest_path):
        if path.is_symlink() or not path.is_file():
            fail(f"retained 0485 custody input is missing: {path}")
    if not seeds_root.is_dir() or seeds_root.is_symlink():
        fail(f"retained 0485 seed directory is missing: {seeds_root}")

    generator = read_json(generator_path)
    manifest = read_json(manifest_path)
    if not isinstance(generator, dict) or not isinstance(manifest, dict):
        fail("retained 0485 generator or seed manifest is not an object")
    if generator.get("schema") != "litchi-docx-stream-fuzz-generator-v1":
        fail("retained 0485 generator schema differs")
    if generator.get("target") != BINARY_NAME:
        fail("retained 0485 generator target differs")
    if generator.get("input_cap_bytes") != INPUT_CAP:
        fail("retained 0485 input cap differs")
    if generator.get("seed_count") != 90 or generator.get("raw_input_sha256_verified") is not True:
        fail("retained 0485 generator does not authenticate exactly 90 raw seeds")
    if generator.get("generator_sha256") != sha(driver_path):
        fail("retained 0485 generator is not bound to its historical driver")
    if len(manifest) != 90:
        fail("retained 0485 seed manifest does not contain exactly 90 entries")

    actual = sorted(
        path.name
        for path in seeds_root.iterdir()
        if path.is_file() and not path.is_symlink()
    )
    for path in seeds_root.iterdir():
        if path.is_symlink() or not path.is_file():
            fail(f"retained 0485 seed directory contains a non-regular entry: {path}")
    if sorted(manifest) != actual or len(actual) != 90:
        fail("retained 0485 seed directory differs from its authenticated manifest")

    records: dict[str, dict[str, int | str]] = {}
    for name in actual:
        path = seeds_root / name
        record = manifest.get(name)
        if not isinstance(record, dict):
            fail(f"retained 0485 seed record is not an object: {name}")
        expected_bytes = record.get("bytes")
        expected_sha = record.get("sha256")
        if not isinstance(expected_bytes, int) or isinstance(expected_bytes, bool):
            fail(f"retained 0485 seed byte count is invalid: {name}")
        if not isinstance(expected_sha, str) or len(expected_sha) != 64:
            fail(f"retained 0485 seed hash is invalid: {name}")
        actual_meta = meta(path)
        if actual_meta != {"bytes": expected_bytes, "sha256": expected_sha}:
            fail(f"retained 0485 seed bytes differ from its manifest: {name}")
        records[name] = actual_meta

    return {
        "driver": meta(driver_path),
        "generator": meta(generator_path),
        "manifest": meta(manifest_path),
        "generator_sha256": generator["generator_sha256"],
        "seed_count": len(records),
        "seed_inventory": records,
    }


def validate_lock(path: Path) -> dict[str, Any]:
    """Validate the checked-in primary fuzz lock without regenerating it."""

    if path.is_symlink() or not path.is_file():
        fail(f"primary fuzz lock is missing: {path}")
    try:
        import tomllib

        lock = tomllib.loads(path.read_text(encoding="utf-8"))
    except (ImportError, OSError, ValueError) as exc:
        fail(f"cannot parse primary fuzz lock {path}: {exc}")
    packages = lock.get("package")
    if not isinstance(packages, list):
        fail("primary fuzz lock has no package table")
    package_names = [item.get("name") for item in packages if isinstance(item, dict)]
    if "litchi-docx-fuzz" not in package_names or "litchi-docx" not in package_names:
        fail("primary fuzz lock is not a standalone DOCX fuzz lock")
    return {
        "version": lock.get("version"),
        "package_count": len(packages),
        "package_names_sha256": canonical_digest(sorted(package_names)),
    }


def manifest_text() -> str:
    after_dependency = json.dumps(str(AFTER_CRATE))
    return """[package]
name = "litchi-docx-fuzz"
version = "0.0.0"
edition = "2024"
publish = false

[dependencies]
libfuzzer-sys = "0.4"
litchi-docx = { path = %s }

[[bin]]
name = "source_backed_tail_append_stream"
path = "source_backed_tail_append_stream.rs"
test = false
doc = false
bench = false

[profile.release]
debug = 1
codegen-units = 1
lto = "thin"

[workspace]
""" % after_dependency


def static_receipt() -> dict[str, Any]:
    receipt = read_json(EVIDENCE / "prepared.json")
    if receipt.get("schema") != "docx-stream-fuzz-prepared-0497-v1":
        fail("prepared receipt schema differs")
    return receipt


def verify_static(receipt: dict[str, Any], *, allow_cleaned: bool = False) -> None:
    if EVIDENCE.is_symlink() or not EVIDENCE.is_dir():
        fail(f"fuzz evidence directory is missing or symlinked: {EVIDENCE}")
    if meta(Path(__file__)) != receipt["driver"]:
        fail("fuzz helper changed after preparation")
    expected = receipt["source_binding"]
    if AFTER_ROOT.is_dir() and not AFTER_ROOT.is_symlink():
        if source_binding() != expected:
            fail("source or after dependency changed after fuzz preparation")
    elif not allow_cleaned:
        fail("after checkout is missing before fuzz verification")

    paths = receipt["paths"]
    package = Path(paths["package"])
    if package != PACKAGE:
        fail("prepared package path differs from the fixed private fuzz path")
    if TARGET_SOURCE.is_file() and not TARGET_SOURCE.is_symlink():
        if meta(TARGET_SOURCE) != receipt["target_source"]["current"]:
            fail("current target source changed after fuzz preparation")
    elif not allow_cleaned:
        fail("current target source is missing before fuzz verification")
    if AFTER_TARGET.is_file() and not AFTER_TARGET.is_symlink():
        if meta(AFTER_TARGET) != receipt["target_source"]["after"]:
            fail("after target source changed after preparation")
    elif not allow_cleaned:
        fail("after target source is missing before fuzz verification")
    if package.is_dir() and not package.is_symlink():
        if meta(package / "source_backed_tail_append_stream.rs") != receipt["target_source"]["copied"]:
            fail("standalone copied fuzz target changed after preparation")
        if meta(package / "Cargo.toml") != receipt["manifest"]["package"]:
            fail("standalone fuzz manifest changed after preparation")
        if meta(package / "Cargo.lock") != receipt["lock"]["package"]:
            fail("standalone fuzz lock changed after preparation")
        if inventory(package / "corpus-start") != receipt["corpus_start"]:
            fail("standalone starting corpus changed after preparation")
    elif not allow_cleaned:
        fail("standalone fuzz package is missing before fuzz verification")

    evidence_paths = receipt["evidence_paths"]
    for key, expected_meta in receipt["evidence_files"].items():
        path = Path(evidence_paths[key])
        if meta(path) != expected_meta:
            fail(f"evidence custody file changed: {path}")

    retained_root = RETAINED_0485 if RETAINED_SEEDS.is_dir() else EVIDENCE / "retained-0485"
    retained = authenticate_retained_0485(retained_root)
    if retained != receipt["retained_0485"]:
        fail("retained 0485 source corpus changed after preparation")
    evidence_retained = EVIDENCE / "retained-0485"
    if not evidence_retained.is_dir() or evidence_retained.is_symlink():
        fail("evidence-side retained 0485 corpus is missing")
    if authenticate_retained_0485(evidence_retained) != receipt["retained_0485"]:
        fail("evidence-side retained 0485 source corpus changed")


def env_for_build() -> dict[str, str]:
    environment = dict(os.environ)
    environment.update(BASE_ENV)
    environment.update(
        {
            "RUSTC_BOOTSTRAP": "1",
            "RUSTFLAGS": ASAN_FLAGS,
            "CARGO_TARGET_DIR": str(CARGO_TARGET),
            "TMPDIR": str(TMPDIR),
        }
    )
    return environment


def env_for_run() -> dict[str, str]:
    environment = dict(os.environ)
    environment.update(BASE_ENV)
    environment["TMPDIR"] = str(TMPDIR)
    return environment


def build_display(argv: list[str], environment: dict[str, str]) -> list[str]:
    keys = (
        "CARGO_TARGET_DIR",
        "TMPDIR",
        "RUSTC_BOOTSTRAP",
        "RUSTFLAGS",
        "RUSTUP_TOOLCHAIN",
        "CARGO_BUILD_JOBS",
        "CARGO_INCREMENTAL",
        "CARGO_PROFILE_RELEASE_DEBUG",
        "DEBUGINFOD_URLS",
        "LC_ALL",
    )
    return ["env", *(f"{key}={environment[key]}" for key in keys), *argv]


def prepare() -> None:
    if TEMP_ROOT.exists() or TEMP_ROOT.is_symlink():
        fail(f"refusing to replace existing private fuzz tree: {TEMP_ROOT}")
    prepared_path = EVIDENCE / "prepared.json"
    if EVIDENCE.is_symlink():
        fail(f"refusing symlinked fuzz evidence directory: {EVIDENCE}")
    if prepared_path.exists() or prepared_path.is_symlink():
        fail(f"refusing to replace retained fuzz evidence: {prepared_path}")

    retained = authenticate_retained_0485()
    lock_info = validate_lock(PRIMARY_LOCK)
    source = source_binding()

    TEMP_ROOT.mkdir(parents=True)
    PACKAGE.mkdir()
    RUN_ROOT.mkdir()
    TMPDIR.mkdir()
    EVIDENCE.mkdir(parents=True, exist_ok=True)

    copied_target = PACKAGE / "source_backed_tail_append_stream.rs"
    copy_file(TARGET_SOURCE, copied_target)
    manifest = PACKAGE / "Cargo.toml"
    with manifest.open("x", encoding="utf-8") as stream:
        stream.write(manifest_text())
    lock = PACKAGE / "Cargo.lock"
    copy_file(PRIMARY_LOCK, lock)

    corpus_start = PACKAGE / "corpus-start"
    copy_corpus(RETAINED_SEEDS, corpus_start)

    # Preserve an evidence-side immutable copy.  This keeps the accepted
    # 0485 source and its two authentication records reviewable after the
    # temporary package is cleaned, without modifying the old bundle.
    retained_evidence = EVIDENCE / "retained-0485"
    retained_evidence.mkdir()
    copy_file(RETAINED_DRIVER, retained_evidence / "fuzz-stream.py")
    copy_file(RETAINED_GENERATOR, retained_evidence / "generator.json")
    copy_file(RETAINED_MANIFEST, retained_evidence / "seed-manifest.json")
    copy_corpus(RETAINED_SEEDS, retained_evidence / "seeds")

    evidence_files: dict[str, dict[str, int | str]] = {}
    evidence_paths: dict[str, str] = {}
    for key, source_path, name in (
        ("target_source", copied_target, "target-source.rs"),
        ("manifest", manifest, "Cargo.toml"),
        ("lock", lock, "Cargo.lock"),
    ):
        destination = EVIDENCE / name
        copy_file(source_path, destination)
        evidence_paths[key] = str(destination)
        evidence_files[key] = meta(destination)

    corpus_inventory = inventory(corpus_start)
    write_json(
        prepared_path,
        {
            "schema": "docx-stream-fuzz-prepared-0497-v1",
            "prepared_utc": now(),
            "driver": meta(Path(__file__)),
            "target": BINARY_NAME,
            "target_triple": TARGET_TRIPLE,
            "input_cap_bytes": INPUT_CAP,
            "fuzz_runs": FUZZ_RUNS,
            "run_seeds": list(RUN_SEEDS),
            "toolchain": TOOLCHAIN,
            "asan_flags": ASAN_FLAGS,
            "paths": {
                "temp_root": str(TEMP_ROOT),
                "package": str(PACKAGE),
                "cargo_target": str(CARGO_TARGET),
                "run_root": str(RUN_ROOT),
                "tmpdir": str(TMPDIR),
                "after_root": str(AFTER_ROOT),
            },
            "source_binding": source,
            "target_source": {
                "current": meta(TARGET_SOURCE),
                "after": meta(AFTER_TARGET),
                "copied": meta(copied_target),
            },
            "manifest": {"package": meta(manifest), "text_sha256": sha(manifest)},
            "lock": {
                "primary": meta(PRIMARY_LOCK),
                "package": meta(lock),
                "validation": lock_info,
            },
            "retained_0485": retained,
            "corpus_start": corpus_inventory,
            "evidence_paths": evidence_paths,
            "evidence_files": evidence_files,
            "evidence_retained_0485": str(retained_evidence),
        },
    )
    print(json.dumps({"prepared": str(prepared_path), "package": str(PACKAGE)}, sort_keys=True))


def build() -> None:
    receipt = static_receipt()
    verify_static(receipt)
    build_json = EVIDENCE / "build.json"
    if build_json.exists() or build_json.is_symlink():
        fail(f"refusing to replace retained build evidence: {build_json}")

    argv = [
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        str(PACKAGE / "Cargo.toml"),
        "--target",
        TARGET_TRIPLE,
        "--bin",
        BINARY_NAME,
    ]
    environment = env_for_build()
    display_argv = build_display(argv, environment)
    source_before = source_binding()
    print(json.dumps(display_argv), flush=True)
    started = now()
    subprocess.run(argv, cwd=REPO, env=environment, check=True)
    source_after = source_binding()
    if source_after != source_before:
        fail("fuzz build changed source or after dependency inputs")

    origin = CARGO_TARGET / TARGET_TRIPLE / "release" / BINARY_NAME
    if origin.is_symlink() or not origin.is_file():
        fail(f"Cargo did not produce the expected sanitizer binary: {origin}")
    retained_binary = RETAINED_BINARY
    copy_file(origin, retained_binary)
    retained_binary.chmod(0o755)
    binary_meta = meta(retained_binary)
    if meta(origin) != binary_meta:
        fail("retained fuzz binary differs from Cargo output")

    write_json(
        build_json,
        {
            "schema": "docx-stream-fuzz-build-0497-v1",
            "driver": meta(Path(__file__)),
            "argv": argv,
            "display_argv": display_argv,
            "cwd": str(REPO),
            "environment": {
                key: environment[key]
                for key in (
                    "CARGO_TARGET_DIR",
                    "TMPDIR",
                    "RUSTC_BOOTSTRAP",
                    "RUSTFLAGS",
                    "RUSTUP_TOOLCHAIN",
                    "CARGO_BUILD_JOBS",
                    "CARGO_INCREMENTAL",
                    "CARGO_PROFILE_RELEASE_DEBUG",
                    "DEBUGINFOD_URLS",
                    "LC_ALL",
                )
            },
            "started_utc": started,
            "finished_utc": now(),
            "source_binding_before": source_before,
            "source_binding_after": source_after,
            "prepared_sha256": sha(EVIDENCE / "prepared.json"),
            "binary": {
                "origin": str(origin),
                "retained": str(retained_binary),
                **binary_meta,
            },
        },
    )


def run() -> None:
    receipt = static_receipt()
    verify_static(receipt)
    build_receipt = read_json(EVIDENCE / "build.json")
    if build_receipt.get("schema") != "docx-stream-fuzz-build-0497-v1":
        fail("build receipt schema differs")
    if build_receipt.get("driver") != meta(Path(__file__)):
        fail("build receipt helper custody differs")
    binary = Path(build_receipt["binary"]["retained"])
    if binary != RETAINED_BINARY:
        fail("build receipt retained binary path differs")
    expected_binary = {
        "bytes": build_receipt["binary"]["bytes"],
        "sha256": build_receipt["binary"]["sha256"],
    }
    if meta(binary) != expected_binary:
        fail("retained fuzz binary changed before run")
    overall_exit = 0
    for seed in RUN_SEEDS:
        label = f"run-{seed}"
        work = RUN_ROOT / label
        evidence = EVIDENCE / label
        if work.exists() or work.is_symlink() or evidence.exists() or evidence.is_symlink():
            fail(f"refusing to replace retained fuzz run: {label}")
        work.mkdir(parents=True)
        corpus = work / "corpus"
        copy_corpus(PACKAGE / "corpus-start", corpus)
        artifacts = work / "artifacts"
        artifacts.mkdir()
        evidence.mkdir()
        copy_corpus(corpus, evidence / "starting-corpus")
        stdout_path = evidence / "terminal.stdout"
        stderr_path = evidence / "terminal.stderr"

        argv = [
            str(binary),
            str(corpus),
            f"-runs={FUZZ_RUNS}",
            f"-seed={seed}",
            f"-max_len={INPUT_CAP}",
            "-timeout=10",
            f"-artifact_prefix={artifacts}/",
        ]
        source_before = source_binding()
        corpus_before = inventory(corpus)
        started = now()
        print(json.dumps(argv), flush=True)
        with stdout_path.open("x", encoding="utf-8") as stdout, stderr_path.open(
            "x", encoding="utf-8"
        ) as stderr:
            try:
                result = subprocess.run(
                    argv,
                    cwd=REPO,
                    env=env_for_run(),
                    stdout=stdout,
                    stderr=stderr,
                )
                exit_code = result.returncode
            except OSError as exc:
                stderr.write(f"{type(exc).__name__}: {exc}\n")
                exit_code = 127
        source_after = source_binding()
        corpus_after = inventory(corpus)
        copy_corpus(corpus, evidence / "post-corpus")
        copy_corpus(artifacts, evidence / "artifacts")
        run_receipt = {
            "schema": "docx-stream-fuzz-run-0497-v1",
            "label": label,
            "seed": seed,
            "argv": argv,
            "cwd": str(REPO),
            "environment": {
                key: env_for_run()[key]
                for key in (
                    "RUSTUP_TOOLCHAIN",
                    "CARGO_BUILD_JOBS",
                    "CARGO_INCREMENTAL",
                    "CARGO_PROFILE_RELEASE_DEBUG",
                    "DEBUGINFOD_URLS",
                    "LC_ALL",
                    "TMPDIR",
                )
            },
            "started_utc": started,
            "finished_utc": now(),
            "exit_code": exit_code,
            "binary": {"path": str(binary), **meta(binary)},
            "source_binding_before": source_before,
            "source_binding_after": source_after,
            "corpus_before": corpus_before,
            "corpus_after": corpus_after,
            "terminal": {
                "stdout": str(stdout_path),
                "stderr": str(stderr_path),
                "stdout_meta": meta(stdout_path),
                "stderr_meta": meta(stderr_path),
            },
            "retained": {
                "starting_corpus": str(evidence / "starting-corpus"),
                "post_corpus": str(evidence / "post-corpus"),
                "artifacts": str(evidence / "artifacts"),
                "post_corpus_inventory": inventory(evidence / "post-corpus"),
                "artifacts_inventory": inventory(evidence / "artifacts"),
            },
        }
        write_json(evidence / "run.json", run_receipt)
        if exit_code and not overall_exit:
            overall_exit = exit_code

    if overall_exit:
        raise SystemExit(overall_exit)


def verify() -> None:
    receipt = static_receipt()
    # The final verifier may run after cleanup has removed the standalone
    # package, Cargo target, and per-run working copies.  Evidence-side seed
    # records and the retained executable remain sufficient for this pass.
    verify_static(receipt, allow_cleaned=True)
    build_receipt = read_json(EVIDENCE / "build.json")
    if build_receipt.get("schema") != "docx-stream-fuzz-build-0497-v1":
        fail("build receipt schema differs")
    if build_receipt.get("driver") != meta(Path(__file__)):
        fail("build receipt helper custody differs")
    binary = Path(build_receipt["binary"]["retained"])
    if binary != RETAINED_BINARY:
        fail("build receipt retained binary path differs")
    expected_binary = {
        "bytes": build_receipt["binary"]["bytes"],
        "sha256": build_receipt["binary"]["sha256"],
    }
    if meta(binary) != expected_binary:
        fail("retained fuzz binary differs from build receipt")

    results: list[dict[str, Any]] = []
    for seed in RUN_SEEDS:
        label = f"run-{seed}"
        run_path = EVIDENCE / label / "run.json"
        run_receipt = read_json(run_path)
        if run_receipt.get("schema") != "docx-stream-fuzz-run-0497-v1":
            fail(f"run receipt schema differs: {label}")
        if run_receipt.get("label") != label or run_receipt.get("seed") != seed:
            fail(f"run identity differs: {label}")
        if run_receipt.get("exit_code") != 0:
            fail(f"fuzz run did not exit successfully: {label}")
        expected_argv = [
            str(binary),
            str(RUN_ROOT / label / "corpus"),
            f"-runs={FUZZ_RUNS}",
            f"-seed={seed}",
            f"-max_len={INPUT_CAP}",
            "-timeout=10",
            f"-artifact_prefix={RUN_ROOT / label / 'artifacts'}/",
        ]
        if run_receipt.get("argv") != expected_argv:
            fail(f"fuzz command vector differs: {label}")
        if run_receipt.get("source_binding_before") != run_receipt.get("source_binding_after"):
            fail(f"source binding changed during fuzz run: {label}")
        if run_receipt.get("binary", {}).get("path") != str(binary):
            fail(f"run binary path differs: {label}")
        if run_receipt.get("binary", {}).get("bytes") != expected_binary["bytes"]:
            fail(f"binary size changed during fuzz run: {label}")
        if run_receipt.get("binary", {}).get("sha256") != expected_binary["sha256"]:
            fail(f"binary hash changed during fuzz run: {label}")

        evidence = EVIDENCE / label
        terminal = run_receipt.get("terminal", {})
        stdout_path = evidence / "terminal.stdout"
        stderr_path = evidence / "terminal.stderr"
        if terminal.get("stdout") != str(stdout_path) or terminal.get("stderr") != str(stderr_path):
            fail(f"terminal paths differ: {label}")
        if terminal.get("stdout_meta") != meta(stdout_path):
            fail(f"stdout custody differs: {label}")
        if terminal.get("stderr_meta") != meta(stderr_path):
            fail(f"stderr custody differs: {label}")
        terminal_text = (stdout_path.read_text(errors="replace") + stderr_path.read_text(errors="replace"))
        forbidden = (
            "AddressSanitizer",
            "UndefinedBehaviorSanitizer",
            "LeakSanitizer",
            "runtime error:",
        )
        if any(marker in terminal_text for marker in forbidden):
            fail(f"sanitizer failure marker found in raw terminal: {label}")

        starting = inventory(evidence / "starting-corpus")
        post = inventory(evidence / "post-corpus")
        expected_start = receipt["corpus_start"]
        if starting != expected_start:
            fail(f"starting corpus differs from authenticated 0485 copy: {label}")
        if any(post.get(name) != record for name, record in expected_start.items()):
            fail(f"post-run corpus lost an authenticated starting seed: {label}")
        retained = run_receipt.get("retained", {})
        if retained.get("starting_corpus") != str(evidence / "starting-corpus"):
            fail(f"retained starting-corpus path differs: {label}")
        if retained.get("post_corpus") != str(evidence / "post-corpus"):
            fail(f"retained post-corpus path differs: {label}")
        if retained.get("artifacts") != str(evidence / "artifacts"):
            fail(f"retained artifacts path differs: {label}")
        if run_receipt.get("corpus_before") != starting or run_receipt.get("corpus_after") != post:
            fail(f"run corpus inventory differs from retained files: {label}")
        if retained.get("post_corpus_inventory") != post:
            fail(f"post-run corpus receipt differs: {label}")
        if retained.get("artifacts_inventory") != inventory(evidence / "artifacts"):
            fail(f"artifact receipt differs from retained files: {label}")
        results.append(
            {
                "label": label,
                "seed": seed,
                "exit_code": 0,
                "starting_seed_count": len(starting),
                "post_seed_count": len(post),
                "terminal": {
                    "stdout_path": str(stdout_path),
                    "stderr_path": str(stderr_path),
                    "stdout_meta": terminal["stdout_meta"],
                    "stderr_meta": terminal["stderr_meta"],
                },
            }
        )

    verify_path = EVIDENCE / "verify.json"
    verification = {
        "schema": "docx-stream-fuzz-verify-0497-v1",
        "verified_utc": now(),
        "passed": True,
        "target": BINARY_NAME,
        "target_triple": TARGET_TRIPLE,
        "input_cap_bytes": INPUT_CAP,
        "fuzz_runs": FUZZ_RUNS,
        "run_seeds": list(RUN_SEEDS),
        "binary": {"path": str(binary), **expected_binary},
        "prepared_sha256": sha(EVIDENCE / "prepared.json"),
        "build_sha256": sha(EVIDENCE / "build.json"),
        "source_binding": (
            source_binding()
            if AFTER_ROOT.is_dir() and not AFTER_ROOT.is_symlink()
            else receipt["source_binding"]
        ),
        "runs": results,
    }
    reused = write_or_check_verification(verify_path, verification)
    print(
        json.dumps(
            {"verified": str(verify_path), "runs": len(results), "reused": reused},
            sort_keys=True,
        )
    )


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("prepare", "build", "run", "verify"))
    options = parser.parse_args(list(argv) if argv is not None else None)
    try:
        {"prepare": prepare, "build": build, "run": run, "verify": verify}[options.command]()
    except (RuntimeError, OSError, subprocess.CalledProcessError) as exc:
        print(f"fuzz helper: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
