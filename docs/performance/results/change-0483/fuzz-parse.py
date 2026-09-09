#!/usr/bin/env python3
"""Retain an isolated ASan/libFuzzer check for the existing parse_docx target.

This helper is deliberately separate from the bounded tail-append fuzz plan.
It consumes the accepted tail fuzz workspace's manifest and lock in place,
copies only the 62 already-authenticated seeds into a separate temporary
corpus, and never takes the CPU lock itself.  Run each subcommand through a
separate ``gate.py`` invocation so that the gate retains this process's
stdout/stderr and source custody.
"""

from __future__ import annotations

import json
from pathlib import Path
import shutil
import subprocess
import sys
import zipfile

from common import ENV, ROOT, REPO, TEMP, meta, now, read, sha, snapshot, write


# ``FUZZ`` is intentionally derived here rather than imported from fuzz.py:
# importing that driver parses its command-line arguments and couples this
# independent check to the tail runner's lifecycle.
FUZZ = ROOT / "fuzz"
TARGET = "x86_64-unknown-linux-gnu"
FLAGS = (
    "-C passes=sancov-module -C llvm-args=-sanitizer-coverage-level=4 "
    "-C llvm-args=-sanitizer-coverage-inline-8bit-counters "
    "-C llvm-args=-sanitizer-coverage-pc-table "
    "-C llvm-args=-sanitizer-coverage-trace-compares -Z sanitizer=address --cfg fuzzing"
)
SEED_ROOT = FUZZ / "seeds"
SEED_MANIFEST = FUZZ / "seed-manifest-v2.json"
GENERATOR = FUZZ / "generator-v2.json"
GENERATOR_SCRIPT = ROOT / "fuzz-seeds.py"
PARSE_TARGET = REPO / "crates/litchi-docx/fuzz/fuzz_targets/parse_docx.rs"
TAIL_TARGET = REPO / "crates/litchi-docx/fuzz/fuzz_targets/tail_append.rs"


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def inventory(directory: Path) -> dict[str, dict[str, int | str]]:
    require(directory.is_dir(), f"missing directory: {directory}")
    return {
        path.relative_to(directory).as_posix(): meta(path)
        for path in sorted(directory.rglob("*"))
        if path.is_file()
    }


def file_ref(path: Path) -> dict[str, int | str]:
    require(path.is_file(), f"missing file: {path}")
    return {"path": str(path), **meta(path)}


def attempt_paths(attempt: str) -> tuple[Path, Path, Path]:
    require(attempt and "/" not in attempt and attempt not in {".", ".."},
            "attempt must be a non-empty path-safe token")
    data = FUZZ if attempt == "initial" else FUZZ / attempt
    tail_work = TEMP / "fuzz-docx" if attempt == "initial" else TEMP / f"fuzz-docx-{attempt}"
    parse_work = TEMP / "fuzz-parse" if attempt == "initial" else TEMP / f"fuzz-parse-{attempt}"
    return data, tail_work, parse_work


def _seed_bundle() -> tuple[dict[str, dict[str, int | str]], dict[str, object]]:
    """Authenticate the v2 manifest, generator, ZIP CRCs, and main XML."""

    require(SEED_MANIFEST.is_file(), f"missing seed manifest: {SEED_MANIFEST}")
    require(GENERATOR.is_file(), f"missing generator receipt: {GENERATOR}")
    manifest = read(SEED_MANIFEST)
    generator = read(GENERATOR)
    require(isinstance(manifest, dict), "v2 seed manifest must be an object")
    require(isinstance(generator, dict), "v2 generator receipt must be an object")
    require(len(manifest) == 62, f"expected 62 seeds, got {len(manifest)}")
    require(generator.get("schema") == "litchi-docx-fuzz-generator-v2",
            "unexpected v2 generator schema")
    require(generator.get("version") == 2, "unexpected generator version")
    require(generator.get("seed_count") == 62 and generator.get("seeds") == 62,
            "generator does not bind 62 seeds")
    require(generator.get("settings_seed_count") == 32,
            "generator does not bind the 32 settings seeds")
    require(generator.get("zip_crc_and_main_hashes_verified") is True,
            "generator did not record ZIP/main hash verification")
    require(generator.get("seed_manifest_sha256") == sha(SEED_MANIFEST),
            "generator does not bind the current seed manifest")
    require(generator.get("generator_sha256") == sha(GENERATOR_SCRIPT),
            "generator does not bind the current generator driver")

    actual: dict[str, dict[str, int | str]] = {}
    for name, expected in sorted(manifest.items()):
        require(isinstance(name, str) and name and "/" not in name and "\\" not in name,
                f"unsafe seed name: {name!r}")
        path = SEED_ROOT / name
        require(path.is_file(), f"manifest seed is missing: {path}")
        observed = meta(path)
        require(observed == {key: expected[key] for key in ("bytes", "sha256")},
                f"seed bytes/hash changed: {name}")
        try:
            with zipfile.ZipFile(path) as archive:
                require(archive.testzip() is None, f"ZIP CRC failed: {name}")
                main_xml = archive.read("word/document.xml")
        except (KeyError, zipfile.BadZipFile, zipfile.LargeZipFile) as error:
            raise RuntimeError(f"invalid DOCX seed {name}: {error}") from error
        require(sha_bytes(main_xml) == expected.get("main_xml_sha256"),
                f"main XML hash changed: {name}")
        actual[name] = observed

    require(inventory(SEED_ROOT) == actual, "seed directory has an unmanifested file")
    return actual, {
        "manifest": file_ref(SEED_MANIFEST),
        "generator": file_ref(GENERATOR),
        "generator_script": file_ref(GENERATOR_SCRIPT),
        "count": len(actual),
        "manifest_schema": "flat-v2-seed-records",
        "zip_crc_and_main_hashes_verified": True,
    }


def sha_bytes(value: bytes) -> str:
    import hashlib

    return hashlib.sha256(value).hexdigest()


def _tail_inputs(data: Path, tail_work: Path, source: dict[str, object]) -> dict[str, object]:
    """Bind the accepted tail workspace and its retained build-input copies."""

    prepared_path = data / "prepared.json"
    build_path = data / "build.json"
    retained_manifest = data / "build-inputs" / "Cargo.toml"
    retained_lock = data / "build-inputs" / "Cargo.lock.txt"
    tail_manifest = tail_work / "Cargo.toml"
    tail_lock = tail_work / "Cargo.lock"
    for path in (prepared_path, build_path, retained_manifest, retained_lock,
                 tail_manifest, tail_lock, PARSE_TARGET, TAIL_TARGET):
        require(path.is_file(), f"missing accepted tail input: {path}")

    prepared = read(prepared_path)
    build = read(build_path)
    require(prepared.get("manifest") == meta(tail_manifest),
            "tail prepared receipt does not bind its workspace manifest")
    require(prepared.get("lock") == meta(tail_lock),
            "tail prepared receipt does not bind its workspace lock")
    require(prepared["manifest"] == meta(retained_manifest),
            "retained Cargo.toml differs from the accepted workspace manifest")
    require(prepared["lock"] == meta(retained_lock),
            "retained Cargo.lock differs from the accepted workspace lock")
    require(build.get("source_snapshot") == source,
            "accepted tail build is from a different source snapshot")
    binary = build.get("binary")
    require(isinstance(binary, dict), "accepted tail build has no binary record")
    binary_path = Path(str(binary.get("path", "")))
    require(binary_path.is_file(), "accepted tail binary is missing")
    require(meta(binary_path) == {key: binary[key] for key in ("bytes", "sha256")},
            "accepted tail binary changed")

    manifest_text = tail_manifest.read_text(encoding="utf-8")
    require('name = "parse_docx"' in manifest_text,
            "accepted build manifest has no parse_docx binary")
    require(f'path = {json.dumps(str(PARSE_TARGET))}' in manifest_text,
            "accepted build manifest does not bind parse_docx.rs")

    return {
        "prepared_record": file_ref(prepared_path),
        "build_record": file_ref(build_path),
        "retained_manifest": file_ref(retained_manifest),
        "retained_lock": file_ref(retained_lock),
        "workspace_manifest": file_ref(tail_manifest),
        "workspace_lock": file_ref(tail_lock),
        "parse_target": file_ref(PARSE_TARGET),
        "tail_target": file_ref(TAIL_TARGET),
        "tail_build_source": source,
        "tail_binary": {"path": str(binary_path), **meta(binary_path)},
    }


def _prepare_inputs(attempt: str, *, source: dict[str, object] | None = None) -> tuple[
    Path, Path, Path, dict[str, object], dict[str, dict[str, int | str]], dict[str, object]
]:
    data, tail_work, parse_work = attempt_paths(attempt)
    prepared_path = data / "parse-prepared.json"
    require(not prepared_path.exists(), f"parse preparation already exists: {prepared_path}")
    require(not parse_work.exists(), f"parse work directory already exists: {parse_work}")
    data.mkdir(parents=True, exist_ok=True)
    if source is None:
        source = snapshot()
    tail = _tail_inputs(data, tail_work, source)
    seeds, seed_bundle = _seed_bundle()
    return data, tail_work, parse_work, tail, seeds, seed_bundle


def prepare(attempt: str) -> None:
    source = snapshot()
    data, tail_work, parse_work, tail, seeds, seed_bundle = _prepare_inputs(
        attempt, source=source
    )
    parse_work.mkdir(parents=True, exist_ok=False)
    corpus = parse_work / "corpus"
    corpus.mkdir()
    copied_names: dict[str, str] = {}
    for name in seeds:
        destination_name = name.replace("/", "-")
        require(destination_name not in copied_names.values(),
                f"seed name collision after corpus flattening: {name}")
        shutil.copyfile(SEED_ROOT / name, corpus / destination_name)
        copied_names[name] = destination_name
    corpus_before = inventory(corpus)
    write(data / "parse-prepared.json", {
        "schema": "litchi-docx-parse-fuzz-prepared-v1",
        "version": 1,
        "kind": "parse_docx",
        "attempt": attempt,
        "prepared_utc": now(),
        "source_snapshot": source,
        "parse_target": file_ref(PARSE_TARGET),
        "tail_inputs": tail,
        "seed_bundle": seed_bundle,
        "seeds": seeds,
        "copied_seed_names": copied_names,
        "corpus_before": corpus_before,
        "work": {
            "path": str(parse_work),
            "corpus": str(corpus),
            "artifacts": str(parse_work / "artifacts"),
            "retained": str(data / "parse-post-run"),
        },
        "execution": {
            "target": TARGET,
            "flags": FLAGS,
            "target_dir": str(REPO / "target/fuzz-asan"),
            "binary": "parse_docx",
            "smoke_runs": 10000,
            "smoke_seed": 483,
            "smoke_max_len": 65536,
            "smoke_timeout_seconds": 10,
            "stdout_stderr_owner": "enclosing gate.py receipt",
            "cpu_lock_owner": "enclosing gate.py receipt",
        },
    })
    print(json.dumps({"prepared": str(data / "parse-prepared.json"),
                      "corpus_files": len(corpus_before)}), flush=True)


def verify_inputs(attempt: str, *, current_source: dict[str, object] | None = None) -> tuple[
    Path, Path, Path, dict[str, object]
]:
    data, tail_work, parse_work = attempt_paths(attempt)
    prepared_path = data / "parse-prepared.json"
    require(prepared_path.is_file(), f"missing parse preparation: {prepared_path}")
    prepared = read(prepared_path)
    require(prepared.get("schema") == "litchi-docx-parse-fuzz-prepared-v1",
            "unexpected parse preparation schema")
    require(prepared.get("kind") == "parse_docx" and prepared.get("attempt") == attempt,
            "parse preparation identity mismatch")
    if current_source is None:
        current_source = snapshot()
    require(prepared.get("source_snapshot") == current_source,
            "source changed after parse preparation")
    require(file_ref(PARSE_TARGET) == prepared["parse_target"],
            "parse_docx target source changed")
    tail = _tail_inputs(data, tail_work, current_source)
    require(tail == prepared["tail_inputs"],
            "accepted tail build-input custody changed")
    seeds, seed_bundle = _seed_bundle()
    require(seeds == prepared["seeds"], "v2 seed inventory changed")
    require(seed_bundle == prepared["seed_bundle"], "v2 seed receipt changed")
    require(parse_work.is_dir(), f"missing parse work directory: {parse_work}")
    corpus = parse_work / "corpus"
    require(inventory(corpus) == prepared["corpus_before"],
            "parse corpus changed before smoke")
    return data, tail_work, parse_work, prepared


def build(attempt: str) -> None:
    data, tail_work, parse_work, prepared = verify_inputs(attempt)
    build_path = data / "parse-build.json"
    destination = parse_work / "parse_docx"
    require(not build_path.exists(), f"parse build receipt already exists: {build_path}")
    require(not destination.exists(), f"parse binary destination already exists: {destination}")
    source_before = snapshot()
    require(source_before == prepared["source_snapshot"], "source changed before parse build")
    manifest = tail_work / "Cargo.toml"
    lock = tail_work / "Cargo.lock"
    manifest_before = meta(manifest)
    lock_before = meta(lock)
    argv = [
        "env",
        f"CARGO_TARGET_DIR={REPO / 'target/fuzz-asan'}",
        "RUSTC_BOOTSTRAP=1",
        f"RUSTFLAGS={FLAGS}",
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        str(manifest),
        "--target",
        TARGET,
        "--bin",
        "parse_docx",
    ]
    print(json.dumps(argv), flush=True)
    started = now()
    subprocess.run(argv, cwd=REPO, env=ENV, check=True)
    source_after = snapshot()
    require(source_after == source_before, "source changed during parse build")
    require(meta(manifest) == manifest_before and meta(lock) == lock_before,
            "accepted manifest or lock changed during parse build")
    origin = REPO / "target/fuzz-asan" / TARGET / "release/parse_docx"
    require(origin.is_file(), f"parse binary was not produced: {origin}")
    with origin.open("rb") as source_stream, destination.open("xb") as destination_stream:
        shutil.copyfileobj(source_stream, destination_stream)
    destination.chmod(0o755)
    origin_meta = meta(origin)
    destination_meta = meta(destination)
    require(origin_meta == destination_meta, "copied parse binary hash differs")
    write(build_path, {
        "schema": "litchi-docx-parse-fuzz-build-v1",
        "version": 1,
        "kind": "build",
        "attempt": attempt,
        "argv": argv,
        "cwd": str(REPO),
        "started_utc": started,
        "finished_utc": now(),
        "source_before": source_before,
        "source_after": source_after,
        "source_unchanged": source_before == source_after,
        "tail_inputs": prepared["tail_inputs"],
        "seed_bundle": prepared["seed_bundle"],
        "manifest_before": manifest_before,
        "lock_before": lock_before,
        "binary_origin": {"path": str(origin), **origin_meta},
        "binary": {"path": str(destination), **destination_meta},
        "stdout_stderr_owner": "enclosing gate.py receipt",
    })
    print(json.dumps({"build": str(build_path), "binary": str(destination)}), flush=True)


def smoke(attempt: str) -> None:
    source_before = snapshot()
    data, _tail_work, parse_work, prepared = verify_inputs(
        attempt, current_source=source_before
    )
    build_path = data / "parse-build.json"
    require(build_path.is_file(), f"missing parse build receipt: {build_path}")
    build_record = read(build_path)
    require(build_record.get("schema") == "litchi-docx-parse-fuzz-build-v1",
            "unexpected parse build schema")
    require(build_record.get("source_before") == source_before
            and build_record.get("source_after") == source_before,
            "parse build is not bound to the current source")
    binary_record = build_record.get("binary")
    require(isinstance(binary_record, dict), "parse build has no copied binary")
    binary = Path(str(binary_record.get("path", "")))
    require(binary == parse_work / "parse_docx", "parse binary path escaped its work dir")
    require(meta(binary) == {key: binary_record[key] for key in ("bytes", "sha256")},
            "parse binary changed after build")
    corpus = parse_work / "corpus"
    artifacts = parse_work / "artifacts"
    retained = data / "parse-post-run"
    smoke_path = data / "parse-smoke.json"
    require(not smoke_path.exists(), f"parse smoke receipt already exists: {smoke_path}")
    require(not artifacts.exists(), f"parse artifact directory already exists: {artifacts}")
    require(not retained.exists(), f"parse retained directory already exists: {retained}")
    corpus_before = inventory(corpus)
    artifacts.mkdir()
    argv = [
        str(binary),
        str(corpus),
        "-runs=10000",
        "-seed=483",
        "-max_len=65536",
        "-timeout=10",
        f"-artifact_prefix={artifacts}/",
    ]
    print(json.dumps(argv), flush=True)
    started = now()
    result = subprocess.run(argv, cwd=REPO, env=ENV)
    corpus_after = inventory(corpus)
    retained.mkdir(exist_ok=False)
    shutil.copytree(corpus, retained / "corpus")
    shutil.copytree(artifacts, retained / "artifacts")
    source_after = snapshot()
    binary_after = meta(binary)
    source_unchanged = source_before == source_after
    binary_unchanged = binary_after == {key: binary_record[key] for key in ("bytes", "sha256")}
    write(smoke_path, {
        "schema": "litchi-docx-parse-fuzz-smoke-v1",
        "version": 1,
        "kind": "smoke",
        "attempt": attempt,
        "argv": argv,
        "cwd": str(REPO),
        "started_utc": started,
        "finished_utc": now(),
        "exit_code": result.returncode,
        "source_before": source_before,
        "source_after": source_after,
        "source_unchanged": source_unchanged,
        "binary_before": dict(binary_record),
        "binary_after": {"path": str(binary), **binary_after},
        "tail_inputs": prepared["tail_inputs"],
        "seed_bundle": prepared["seed_bundle"],
        "corpus_before": corpus_before,
        "corpus_after": corpus_after,
        "retained": {
            "path": str(retained),
            "files": inventory(retained),
        },
        "stdout_stderr_owner": "enclosing gate.py receipt",
        "cpu_lock_owner": "enclosing gate.py receipt",
    })
    require(source_unchanged, "source changed during parse smoke")
    require(binary_unchanged, "parse binary changed during parse smoke")
    if result.returncode:
        raise SystemExit(result.returncode)
    print(json.dumps({"smoke": str(smoke_path), "exit_code": result.returncode}), flush=True)


def parse_args() -> tuple[str, str]:
    values = list(sys.argv[1:])
    if len(values) != 3 or values[0] != "--attempt":
        raise SystemExit(
            "usage: fuzz-parse.py --attempt ATTEMPT {prepare|build|smoke}"
        )
    attempt, command = values[1:]
    require(command in {"prepare", "build", "smoke"}, f"unknown command: {command}")
    return attempt, command


def main() -> None:
    attempt, command = parse_args()
    {"prepare": prepare, "build": build, "smoke": smoke}[command](attempt)


if __name__ == "__main__":
    main()
