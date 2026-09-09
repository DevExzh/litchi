#!/usr/bin/env python3
"""Prepare, build, and retain the bounded DOCX stream fuzz lane.

The target accepts a small recipe before a DOCX archive.  This driver prefixes
the accepted 0483 plain, Strict, default-namespace, and final-section seeds
with valid and intentionally malformed event recipes.  The archive bytes are
copied unchanged, and every derived input is recorded with its raw SHA-256.

The three commands are deliberately independent so the coordinator can run
each through ``gate.py``:

    python3 fuzz-stream.py prepare accepted
    python3 fuzz-stream.py build accepted
    python3 fuzz-stream.py smoke accepted

This helper does not take the CPU lock itself.  The outer 0485 gate owns that
lock and captures the command receipts.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any

from support import ENV, REPO, TEMP, meta, now, read, sha, snapshot, write


ROOT = Path(__file__).resolve().parent
FUZZ = ROOT / "fuzz-stream"
ATTEMPT = sys.argv[2] if len(sys.argv) > 2 else "accepted"
if not ATTEMPT or "/" in ATTEMPT or ATTEMPT in {".", ".."}:
    raise SystemExit("attempt must be a non-empty path-safe token")
DATA = FUZZ / ATTEMPT
WORK = TEMP / ("fuzz-docx-stream" if ATTEMPT == "initial" else f"fuzz-docx-stream-{ATTEMPT}")
TARGET = "x86_64-unknown-linux-gnu"
BINARY_NAME = "source_backed_tail_append_stream"
INPUT_CAP = 64 * 1024

ASAN_FLAGS = (
    "-C passes=sancov-module -C llvm-args=-sanitizer-coverage-level=4 "
    "-C llvm-args=-sanitizer-coverage-inline-8bit-counters "
    "-C llvm-args=-sanitizer-coverage-pc-table "
    "-C llvm-args=-sanitizer-coverage-trace-compares -Z sanitizer=address --cfg fuzzing"
)

# These are small source-backed documents already authenticated by the sealed
# 0483 corpus.  The first four provide the positive namespace/section matrix;
# the empty-paragraph case adds a valid empty authored paragraph destination.
SOURCE_SEEDS = (
    "plain-stored.docx",
    "plain-deflate.docx",
    "strict-stored.docx",
    "strict-deflate.docx",
    "default-namespace-stored.docx",
    "default-namespace-deflate.docx",
    "section-stored.docx",
    "section-deflate.docx",
    "empty-paragraph-stored.docx",
)

# The first byte selects the source-backed target's event case modulo eight.
# Empty recipe means the target's default valid one-paragraph stream.  The
# non-empty valid recipes exercise borrowed chunks and XML escaping; the other
# recipes are expected typed refusals.
RECIPES = (
    ("valid-empty", b"", True),
    ("valid-escaped-chunks", bytes((0, 0, 2, 3, 4, 5, 6, 7, 8, 9)), True),
    ("valid-many-paragraphs", bytes((0, 3, 1, 2, 3, 4, 5, 6, 7, 8, 9)), True),
    ("text-before-start", b"\x01", False),
    ("nested-start", b"\x02", False),
    ("truncated-paragraph", b"\x03", False),
    ("unmatched-end", b"\x04", False),
    ("invalid-xml-character", b"\x05", False),
    ("changed-replay", bytes((6, 0, 2, 3, 4)), False),
    ("empty-stream", b"\x07", False),
)


def inventory(directory: Path) -> dict[str, dict[str, int | str]]:
    return {
        path.relative_to(directory).as_posix(): meta(path)
        for path in sorted(directory.rglob("*"))
        if path.is_file()
    }


def source_manifest() -> tuple[Path, dict[str, Any]]:
    source_root = ROOT.parent / "change-0483" / "fuzz"
    seeds = source_root / "seeds"
    manifest_path = source_root / "seed-manifest-v2.json"
    if not seeds.is_dir() or not manifest_path.is_file():
        raise RuntimeError("0483 fuzz seed custody inputs are missing")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    for name in SOURCE_SEEDS:
        record = manifest.get(name)
        path = seeds / name
        if not isinstance(record, dict) or not path.is_file():
            raise RuntimeError(f"0483 source seed is missing: {name}")
        expected = record.get("sha256")
        if not isinstance(expected, str) or sha(path) != expected:
            raise RuntimeError(f"0483 source seed hash differs: {name}")
    return seeds, manifest


def make_seed_records() -> tuple[dict[str, bytes], dict[str, dict[str, Any]]]:
    source_dir, source_manifest_data = source_manifest()
    inputs: dict[str, bytes] = {}
    records: dict[str, dict[str, Any]] = {}
    for source_name in SOURCE_SEEDS:
        source = (source_dir / source_name).read_bytes()
        source_record = source_manifest_data[source_name]
        stem = source_name.removesuffix(".docx")
        for recipe_name, recipe, expected_success in RECIPES:
            name = f"{recipe_name}-{stem}.bin"
            data = recipe + source
            if len(data) > INPUT_CAP:
                raise RuntimeError(f"generated fuzz seed exceeds input cap: {name}")
            inputs[name] = data
            records[name] = {
                "bytes": len(data),
                "sha256": hashlib.sha256(data).hexdigest(),
                "input_cap_bytes": INPUT_CAP,
                "recipe_hex": recipe.hex(),
                "recipe_name": recipe_name,
                "expected_prepare_success": expected_success,
                "source_seed": source_name,
                "source_bytes": source_record["bytes"],
                "source_sha256": source_record["sha256"],
                "source_main_xml_sha256": source_record["main_xml_sha256"],
            }
    return inputs, records


def build_manifest() -> str:
    original = REPO / "crates/litchi-docx/fuzz/Cargo.toml"
    target = REPO / "crates/litchi-docx/fuzz/fuzz_targets/source_backed_tail_append_stream.rs"
    manifest = original.read_text(encoding="utf-8")
    manifest = manifest.replace(
        'path = ".."',
        "path = " + json.dumps(str(REPO / "crates/litchi-docx")),
    )
    manifest = manifest.replace(
        'path = "fuzz_targets/source_backed_tail_append_stream.rs"',
        "path = " + json.dumps(str(target)),
    )
    manifest = manifest.replace(
        'path = "fuzz_targets/tail_append.rs"',
        "path = " + json.dumps(str(REPO / "crates/litchi-docx/fuzz/fuzz_targets/tail_append.rs")),
    )
    manifest = manifest.replace(
        'path = "fuzz_targets/parse_docx.rs"',
        "path = " + json.dumps(str(REPO / "crates/litchi-docx/fuzz/fuzz_targets/parse_docx.rs")),
    )
    return manifest


def prepare() -> None:
    if WORK.exists() or DATA.exists():
        raise RuntimeError(f"refusing to replace existing fuzz attempt: {WORK} or {DATA}")
    inputs, records = make_seed_records()
    WORK.mkdir(parents=True)
    DATA.mkdir(parents=True)
    seeds = DATA / "seeds"
    seeds.mkdir()
    corpus = WORK / "corpus"
    corpus.mkdir()
    for name, data in inputs.items():
        for destination in (seeds / name, corpus / name):
            with destination.open("xb") as stream:
                stream.write(data)

    manifest_text = json.dumps(records, indent=2, sort_keys=True) + "\n"
    manifest_path = DATA / "seed-manifest.json"
    with manifest_path.open("x", encoding="utf-8") as stream:
        stream.write(manifest_text)

    generator = {
        "schema": "litchi-docx-stream-fuzz-generator-v1",
        "target": BINARY_NAME,
        "input_cap_bytes": INPUT_CAP,
        "source_seed_change": "0483",
        "source_seed_manifest_sha256": sha(
            ROOT.parent / "change-0483" / "fuzz" / "seed-manifest-v2.json"
        ),
        "generator_sha256": sha(Path(__file__)),
        "source_seeds": list(SOURCE_SEEDS),
        "recipes": [
            {"name": name, "recipe_hex": recipe.hex(), "expected_prepare_success": expected}
            for name, recipe, expected in RECIPES
        ],
        "seed_count": len(records),
        "raw_input_sha256_verified": True,
    }
    with (DATA / "generator.json").open("x", encoding="utf-8") as stream:
        json.dump(generator, stream, indent=2, sort_keys=True)
        stream.write("\n")

    manifest = build_manifest()
    build_inputs = WORK / "build-inputs"
    build_inputs.mkdir()
    with (WORK / "Cargo.toml").open("x", encoding="utf-8") as stream:
        stream.write(manifest)
    with (build_inputs / "Cargo.toml").open("x", encoding="utf-8") as stream:
        stream.write(manifest)
    subprocess.run(
        [
            "cargo",
            "generate-lockfile",
            "--offline",
            "--manifest-path",
            str(WORK / "Cargo.toml"),
        ],
        cwd=REPO,
        env=ENV,
        check=True,
    )
    shutil.copyfile(WORK / "Cargo.lock", build_inputs / "Cargo.lock.txt")

    prepared = {
        "schema": "docx-stream-fuzz-prepared-v1",
        "prepared_utc": now(),
        "target_source": meta(REPO / "crates/litchi-docx/fuzz/fuzz_targets/source_backed_tail_append_stream.rs"),
        "manifest": meta(WORK / "Cargo.toml"),
        "lock": meta(WORK / "Cargo.lock"),
        "generator": meta(DATA / "generator.json"),
        "seed_manifest": meta(manifest_path),
        "seed_records": records,
        "corpus_before": inventory(corpus),
        "source_snapshot": snapshot(),
    }
    write(DATA / "prepared.json", prepared)


def verify_inputs() -> dict[str, Any]:
    prepared = read(DATA / "prepared.json")
    target = REPO / "crates/litchi-docx/fuzz/fuzz_targets/source_backed_tail_append_stream.rs"
    if meta(target) != prepared["target_source"]:
        raise RuntimeError("fuzz target source changed after prepare")
    if meta(WORK / "Cargo.toml") != prepared["manifest"]:
        raise RuntimeError("temporary fuzz manifest changed after prepare")
    if meta(WORK / "Cargo.lock") != prepared["lock"]:
        raise RuntimeError("temporary fuzz lock changed after prepare")
    if inventory(DATA / "seeds") != {
        name: {"bytes": record["bytes"], "sha256": record["sha256"]}
        for name, record in prepared["seed_records"].items()
    }:
        raise RuntimeError("retained seed bytes differ from prepared raw hashes")
    if inventory(WORK / "corpus") != prepared["corpus_before"]:
        raise RuntimeError("fuzz corpus changed before build/smoke")
    return prepared


def build() -> None:
    prepared = verify_inputs()
    before = snapshot()
    argv = [
        "env",
        f"CARGO_TARGET_DIR={REPO / 'target/fuzz-asan'}",
        "RUSTC_BOOTSTRAP=1",
        f"RUSTFLAGS={ASAN_FLAGS}",
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        str(WORK / "Cargo.toml"),
        "--target",
        TARGET,
        "--bin",
        BINARY_NAME,
    ]
    print(json.dumps(argv), flush=True)
    started = now()
    subprocess.run(argv, cwd=REPO, env=ENV, check=True)
    after = snapshot()
    if after != before:
        raise RuntimeError("fuzz build changed source inputs")
    origin = REPO / "target/fuzz-asan" / TARGET / "release" / BINARY_NAME
    destination = WORK / BINARY_NAME
    with origin.open("rb") as source, destination.open("xb") as output:
        shutil.copyfileobj(source, output)
    destination.chmod(0o755)
    if meta(origin) != meta(destination):
        raise RuntimeError("retained fuzz binary differs from Cargo output")
    write(
        DATA / "build.json",
        {
            "schema": "docx-stream-fuzz-build-v1",
            "argv": argv,
            "cwd": str(REPO),
            "started_utc": started,
            "finished_utc": now(),
            "source_snapshot_before": before,
            "source_snapshot_after": after,
            "inputs": prepared,
            "binary": {"path": str(destination), **meta(destination)},
        },
    )


def positive_check(prepared: dict[str, Any], binary: Path) -> dict[str, Any]:
    """Run every fixed expected-positive seed before corpus mutation.

    The environment variable is consumed only by the fuzz target.  It turns
    an otherwise useful skip-on-unadmissible fuzz callback into a strict
    success assertion for these immutable, known-valid inputs.
    """

    root = DATA / "positive-check"
    if root.exists():
        raise RuntimeError("refusing to replace retained positive-check artifacts")
    root.mkdir()
    strict_env = dict(ENV, LITCHI_DOCX_STREAM_REQUIRE_SUCCESS="1")
    cases: list[dict[str, Any]] = []
    overall_exit = 0
    positive_records = sorted(
        (
            name,
            record,
        )
        for name, record in prepared["seed_records"].items()
        if record["expected_prepare_success"] is True
    )
    for index, (name, record) in enumerate(positive_records, start=1):
        case_root = root / f"{index:03d}-{name}"
        case_root.mkdir()
        artifacts = case_root / "artifacts"
        artifacts.mkdir()
        stdout_path = case_root / "stdout"
        stderr_path = case_root / "stderr"
        input_path = DATA / "seeds" / name
        argv = [
            str(binary),
            str(input_path),
            "-runs=1",
            f"-seed={4_840 + index}",
            f"-max_len={INPUT_CAP}",
            "-timeout=10",
            f"-artifact_prefix={artifacts}/",
        ]
        started = now()
        with stdout_path.open("x", encoding="utf-8") as stdout, stderr_path.open(
            "x", encoding="utf-8"
        ) as stderr:
            result = subprocess.run(
                argv,
                cwd=REPO,
                env=strict_env,
                stdout=stdout,
                stderr=stderr,
            )
        if result.returncode and not overall_exit:
            overall_exit = result.returncode
        cases.append(
            {
                "name": name,
                "input": {"path": str(input_path), **meta(input_path)},
                "source_seed": record["source_seed"],
                "recipe_name": record["recipe_name"],
                "recipe_hex": record["recipe_hex"],
                "argv": argv,
                "cwd": str(REPO),
                "started_utc": started,
                "finished_utc": now(),
                "exit_code": result.returncode,
                "artifacts": inventory(case_root),
            }
        )
    receipt = {
        "schema": "docx-stream-fuzz-positive-v1",
        "environment_variable": "LITCHI_DOCX_STREAM_REQUIRE_SUCCESS",
        "input_cap_bytes": INPUT_CAP,
        "case_count": len(cases),
        "exit_code": overall_exit,
        "cases": cases,
    }
    write(DATA / "positive.json", receipt)
    if overall_exit:
        raise SystemExit(overall_exit)
    return receipt


def smoke() -> None:
    prepared = verify_inputs()
    build_record = read(DATA / "build.json")
    binary = Path(build_record["binary"]["path"])
    expected_binary = {key: build_record["binary"][key] for key in ("bytes", "sha256")}
    if meta(binary) != expected_binary:
        raise RuntimeError("retained fuzz binary changed before smoke")
    positive = positive_check(prepared, binary)
    corpus = WORK / "corpus"
    retained = DATA / "post-run"
    if retained.exists():
        raise RuntimeError("refusing to replace retained fuzz smoke artifacts")
    retained.mkdir()
    run_records: list[dict[str, Any]] = []
    overall_exit = 0
    for run_number, seed in enumerate((484, 485), start=1):
        artifacts = WORK / "artifacts" / f"run-{run_number}"
        artifacts.mkdir(parents=True)
        before = inventory(corpus)
        argv = [
            str(binary),
            str(corpus),
            "-runs=10000",
            f"-seed={seed}",
            f"-max_len={INPUT_CAP}",
            "-timeout=10",
            f"-artifact_prefix={artifacts}/",
        ]
        started = now()
        print(json.dumps(argv), flush=True)
        result = subprocess.run(argv, cwd=REPO, env=ENV)
        if result.returncode and not overall_exit:
            overall_exit = result.returncode
        retained_run = retained / f"run-{run_number}"
        retained_run.mkdir()
        shutil.copytree(corpus, retained_run / "corpus")
        shutil.copytree(artifacts, retained_run / "artifacts")
        run_records.append(
            {
                "argv": argv,
                "cwd": str(REPO),
                "started_utc": started,
                "finished_utc": now(),
                "exit_code": result.returncode,
                "corpus_before": before,
                "retained": inventory(retained_run),
            }
        )
    write(
        DATA / "smoke.json",
        {
            "schema": "docx-stream-fuzz-smoke-v1",
            "positive_check": positive,
            "runs": run_records,
            "binary": {"path": str(binary), **meta(binary)},
            "input_cap_bytes": INPUT_CAP,
            "corpus_after": inventory(corpus),
        },
    )
    raise SystemExit(overall_exit)


if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] not in {"prepare", "build", "smoke"}:
        raise SystemExit("usage: fuzz-stream.py {prepare|build|smoke} [ATTEMPT]")
    {"prepare": prepare, "build": build, "smoke": smoke}[sys.argv[1]]()
