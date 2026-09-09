#!/usr/bin/env python3
"""Verify the 0482 XML reader evidence bundle without running the harness."""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import re
import runpy
import sys
from typing import Any, Mapping

import analyze
from common import ENV_KEYS, ROOT, sha


class VerificationError(ValueError):
    pass


BUILD_SCHEMA = "xml-stream-audit-build-v1"
BUILD_LABELS = ("build-normal", "build-allocator")
XML_ASSET_SCHEMA = "xml-stream-audit-xml-assets-v1"
XML_ASSET_INVENTORY_PATH = "xml-assets.json"
XML_ASSET_COUNT = 77
FUZZ_SCHEMA = analyze.FUZZ_SCHEMA
FUZZ_DATA_PREFIX = analyze.FUZZ_DATA_PREFIX
FUZZ_SEED_COUNT = analyze.FUZZ_SEED_COUNT
FUZZ_TARGET = analyze.FUZZ_TARGET
FUZZ_FLAGS = (
    "-C passes=sancov-module -C llvm-args=-sanitizer-coverage-level=4 "
    "-C llvm-args=-sanitizer-coverage-inline-8bit-counters "
    "-C llvm-args=-sanitizer-coverage-pc-table "
    "-C llvm-args=-sanitizer-coverage-trace-compares -Z sanitizer=address --cfg fuzzing"
)
REQUIRED_VALIDATION_LABELS = analyze.REQUIRED_VALIDATION_LABELS
INSTRUMENTATIONS = analyze.INSTRUMENTATIONS
MODES = analyze.MODES
SIZES = analyze.SIZES
REQUIRED_BUILD_FIELDS = {
    "schema",
    "instrumentation",
    "source_snapshot",
    "argv",
    "git_revision",
    "environment",
    "started_utc",
    "finished_utc",
    "binary",
}
SHA256 = re.compile(r"^[0-9a-f]{64}$")
GIT_REVISION = re.compile(r"^[0-9a-f]{40,64}$")


def fail(message: str) -> None:
    raise VerificationError(message)


def _timestamp(value: Any, label: str) -> _datetime.datetime:
    if not isinstance(value, str) or not value:
        fail(f"{label}: timestamp missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: invalid timestamp: {error}")
    if parsed.tzinfo is None:
        fail(f"{label}: timestamp lacks timezone")
    return parsed


def _interval(start: Any, finish: Any, label: str) -> tuple[_datetime.datetime, _datetime.datetime]:
    beginning = _timestamp(start, f"{label}.started_utc")
    ending = _timestamp(finish, f"{label}.finished_utc")
    if ending < beginning:
        fail(f"{label}: finished before started")
    return beginning, ending


def read_json(path: Path, label: str | None = None) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        fail(f"{label or path}: cannot read JSON: {error}")


def _meta(path: Path) -> dict[str, Any]:
    if not path.is_file():
        fail(f"missing retained artifact: {path}")
    return {"bytes": path.stat().st_size, "sha256": sha(path)}


def _check_meta(value: Any, path: Path, label: str) -> None:
    if not isinstance(value, dict) or value.get("bytes") != _meta(path)["bytes"] or value.get("sha256") != _meta(path)["sha256"]:
        fail(f"{label}: retained artifact metadata differs")


def _is_xml_asset_path(value: Any) -> bool:
    if not isinstance(value, str):
        return False
    relative = Path(value)
    return (
        not relative.is_absolute()
        and ".." not in relative.parts
        and relative.parts[:1] == ("crates",)
        and "src" in relative.parts
        and relative.suffix == ".xml"
    )


def _check_snapshot(root: Path, value: Any, label: str) -> dict[str, str]:
    if not isinstance(value, dict) or set(value) != {"path", "sha256", "files"}:
        fail(f"{label}: source snapshot fields differ")
    relative = Path(value["path"]) if isinstance(value.get("path"), str) else Path("")
    if relative.is_absolute() or ".." in relative.parts:
        fail(f"{label}: source snapshot path escapes the evidence root")
    path = root / relative
    if not path.is_file() or sha(path) != value["sha256"] or path.stem != value["sha256"]:
        fail(f"{label}: source snapshot digest differs")
    if not isinstance(value["files"], int) or value["files"] <= 0:
        fail(f"{label}: source snapshot file count differs")
    manifest = read_json(path, f"{label}.manifest")
    if (
        not isinstance(manifest, dict)
        or len(manifest) != value["files"]
        or not all(isinstance(item, str) and SHA256.fullmatch(item) for item in manifest.values())
    ):
        fail(f"{label}: source manifest contents differ")
    return {str(name): str(digest) for name, digest in manifest.items()}


def _live_xml_assets(root: Path) -> dict[str, str]:
    if len(root.parents) <= 3:
        fail("live XML asset check: cannot locate repository root")
    repository = root.parents[3]
    crates = repository / "crates"
    if not crates.is_dir():
        fail(f"live XML asset check: repository crates directory is missing: {crates}")
    assets: dict[str, str] = {}
    for path in crates.rglob("*.xml"):
        relative = path.relative_to(repository)
        if path.is_file() and "src" in relative.parts:
            assets[relative.as_posix()] = sha(path)
    return assets


def _check_xml_inventory(
    root: Path,
    protocol: Mapping[str, Any],
    source_manifest: Mapping[str, str],
    live: bool = False,
) -> None:
    reference = protocol.get("xml_asset_inventory")
    if (
        not isinstance(reference, dict)
        or set(reference) != {"path", "sha256", "files"}
        or reference.get("path") != XML_ASSET_INVENTORY_PATH
        or reference.get("files") != XML_ASSET_COUNT
        or not isinstance(reference.get("sha256"), str)
        or SHA256.fullmatch(reference["sha256"]) is None
    ):
        fail("protocol.xml_asset_inventory binding differs")
    path = root / XML_ASSET_INVENTORY_PATH
    if not path.is_file() or sha(path) != reference["sha256"]:
        fail("xml-assets.json: retained inventory digest differs")
    inventory = read_json(path, "xml-assets.json")
    if not isinstance(inventory, dict) or set(inventory) != {"schema", "files", "assets"}:
        fail("xml-assets.json: inventory fields differ")
    assets = inventory.get("assets")
    if (
        inventory.get("schema") != XML_ASSET_SCHEMA
        or inventory.get("files") != XML_ASSET_COUNT
        or not isinstance(assets, dict)
        or len(assets) != XML_ASSET_COUNT
        or not all(_is_xml_asset_path(name) for name in assets)
        or not all(isinstance(digest, str) and SHA256.fullmatch(digest) for digest in assets.values())
    ):
        fail("xml-assets.json: inventory contents differ")
    if any(source_manifest.get(name) != digest for name, digest in assets.items()):
        fail("xml-assets.json: XML asset hash differs from retained source snapshot")
    source_assets = {name: digest for name, digest in source_manifest.items() if _is_xml_asset_path(name)}
    if source_assets != assets:
        fail("xml-assets.json: retained source snapshot XML inventory differs")
    if live and _live_xml_assets(root) != assets:
        fail("xml-assets.json: current repository XML inventory differs")


def _inventory(root: Path, relative: str, label: str) -> dict[str, dict[str, Any]]:
    directory = root / relative
    if not directory.is_dir():
        fail(f"{label}: retained directory is missing")
    values: dict[str, dict[str, Any]] = {}
    for path in sorted(directory.rglob("*")):
        if path.is_symlink():
            fail(f"{label}: retained directory contains a symlink")
        if path.is_file():
            values[path.relative_to(directory).as_posix()] = _meta(path)
    return values


def _inventory_digest(values: Mapping[str, Mapping[str, Any]]) -> str:
    encoded = json.dumps(values, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(encoded).hexdigest()


def _check_fuzz_file(root: Path, value: Any, relative: str, label: str) -> Path:
    if (
        not isinstance(value, dict)
        or set(value) != {"path", "bytes", "sha256"}
        or value.get("path") != relative
        or isinstance(value.get("bytes"), bool)
        or not isinstance(value.get("bytes"), int)
        or value["bytes"] <= 0
        or not isinstance(value.get("sha256"), str)
        or SHA256.fullmatch(value["sha256"]) is None
    ):
        fail(f"{label}: retained file reference differs")
    path = root / relative
    if not path.is_file() or _meta(path) != {"bytes": value["bytes"], "sha256": value["sha256"]}:
        fail(f"{label}: retained file digest differs")
    return path


def _check_fuzz_inventory(
    root: Path,
    value: Any,
    relative: str,
    minimum_files: int,
    label: str,
) -> dict[str, dict[str, Any]]:
    if (
        not isinstance(value, dict)
        or set(value) != {"path", "files", "sha256"}
        or value.get("path") != relative
        or isinstance(value.get("files"), bool)
        or not isinstance(value.get("files"), int)
        or value["files"] < minimum_files
        or not isinstance(value.get("sha256"), str)
        or SHA256.fullmatch(value["sha256"]) is None
    ):
        fail(f"{label}: retained directory inventory reference differs")
    values = _inventory(root, relative, label)
    if len(values) != value["files"] or _inventory_digest(values) != value["sha256"]:
        fail(f"{label}: retained directory inventory digest differs")
    return values


def _path_equal(actual: Any, expected: Path) -> bool:
    return isinstance(actual, str) and Path(actual).is_absolute() and Path(actual).resolve() == expected.resolve()


def _check_fuzz(
    root: Path,
    protocol: Mapping[str, Any],
    source_manifest: Mapping[str, str],
    live: bool = False,
) -> dict[str, _datetime.datetime]:
    fuzz = protocol.get("fuzz")
    if not isinstance(fuzz, dict):
        fail("protocol.fuzz: evidence binding is missing")
    cwd = fuzz.get("cwd")
    if not isinstance(cwd, str) or not Path(cwd).is_absolute():
        fail("protocol.fuzz.cwd: expected an absolute recorded repository path")
    recorded_cwd = Path(cwd)
    helper = _check_fuzz_file(root, fuzz.get("helper"), "fuzz.py", "protocol.fuzz.helper")
    if sha(helper) != sha(root / "fuzz.py"):
        fail("protocol.fuzz.helper: script digest differs")
    paths = {
        name: _check_fuzz_file(root, fuzz.get(name), f"{FUZZ_DATA_PREFIX}/{name}.json", f"protocol.fuzz.{name}")
        for name in ("prepared", "build", "smoke")
    }
    manifest_path = _check_fuzz_file(root, fuzz.get("manifest"), f"{FUZZ_DATA_PREFIX}/build-inputs/Cargo.toml", "protocol.fuzz.manifest")
    lock_path = _check_fuzz_file(root, fuzz.get("lock"), f"{FUZZ_DATA_PREFIX}/build-inputs/Cargo.lock.txt", "protocol.fuzz.lock")
    seeds = _check_fuzz_inventory(root, fuzz.get("seed_inventory"), "fuzz/seeds", FUZZ_SEED_COUNT, "protocol.fuzz.seed_inventory")
    post_run = _check_fuzz_inventory(root, fuzz.get("post_run_inventory"), f"{FUZZ_DATA_PREFIX}/post-run", 1, "protocol.fuzz.post_run_inventory")

    prepared = read_json(paths["prepared"], "fuzz/final/prepared.json")
    if not isinstance(prepared, dict) or set(prepared) != {"prepared_utc", "target_source", "manifest", "lock", "seeds", "corpus_before"}:
        fail("fuzz/final/prepared.json: record fields differ")
    _timestamp(prepared.get("prepared_utc"), "fuzz/final/prepared.prepared_utc")
    target_source = prepared.get("target_source")
    if (
        not isinstance(target_source, dict)
        or set(target_source) != {"bytes", "sha256"}
        or isinstance(target_source.get("bytes"), bool)
        or not isinstance(target_source.get("bytes"), int)
        or target_source["bytes"] <= 0
        or not isinstance(target_source.get("sha256"), str)
        or SHA256.fullmatch(target_source["sha256"]) is None
        or source_manifest.get("crates/xml-minifier/fuzz/fuzz_targets/minify_xml.rs") != target_source["sha256"]
    ):
        fail("fuzz/final/prepared.json: target source binding differs")
    for field, path in (("manifest", manifest_path), ("lock", lock_path)):
        record = prepared.get(field)
        if not isinstance(record, dict) or set(record) != {"bytes", "sha256"} or record != _meta(path):
            fail(f"fuzz/final/prepared.json: {field} binding differs")
    if source_manifest.get(f"docs/performance/results/change-0482/{FUZZ_DATA_PREFIX}/build-inputs/Cargo.toml") != fuzz["manifest"]["sha256"]:
        fail("fuzz/final/build-inputs/Cargo.toml: source snapshot binding differs")
    if prepared.get("seeds") != seeds or len(seeds) != FUZZ_SEED_COUNT:
        fail("fuzz/final/prepared.json: seed inventory differs")
    expected_corpus = {name.replace("/", "-"): value for name, value in seeds.items()}
    if prepared.get("corpus_before") != expected_corpus:
        fail("fuzz/final/prepared.json: prepared corpus differs from seeds")

    build = read_json(paths["build"], "fuzz/final/build.json")
    if not isinstance(build, dict) or set(build) != {"argv", "cwd", "started_utc", "finished_utc", "source_snapshot", "inputs", "binary"}:
        fail("fuzz/final/build.json: record fields differ")
    if build.get("cwd") != cwd or build.get("source_snapshot") != protocol.get("source_snapshot") or build.get("inputs") != prepared:
        fail("fuzz/final/build.json: source or input binding differs")
    build_started, build_finished = _interval(build.get("started_utc"), build.get("finished_utc"), "fuzz/final/build")
    binary = build.get("binary")
    if not isinstance(binary, dict):
        fail("fuzz/final/build.json: binary binding is missing")
    _check_binary(binary, binary, "fuzz/final/build.binary", live=live)
    binary_path = Path(binary["path"])
    work = binary_path.parent
    if binary_path.name != "minify_xml" or not work.is_absolute():
        fail("fuzz/final/build.json: binary path differs")
    build_argv = build.get("argv")
    expected_build_argv = [
        "env",
        f"CARGO_TARGET_DIR={(recorded_cwd / 'target/fuzz-asan').resolve()}",
        "RUSTC_BOOTSTRAP=1",
        f"RUSTFLAGS={FUZZ_FLAGS}",
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        str((work / "Cargo.toml").resolve()),
        "--target",
        FUZZ_TARGET,
        "--bin",
        "minify_xml",
    ]
    if build_argv != expected_build_argv:
        fail("fuzz/final/build.json: build command binding differs")
    if not _path_equal(binary.get("path"), work / "minify_xml"):
        fail("fuzz/final/build.json: binary path is not rooted at the retained work directory")

    smoke = read_json(paths["smoke"], "fuzz/final/smoke.json")
    if not isinstance(smoke, dict) or set(smoke) != {"argv", "cwd", "started_utc", "finished_utc", "exit_code", "binary_before", "binary_after", "corpus_before", "retained"}:
        fail("fuzz/final/smoke.json: record fields differ")
    if smoke.get("cwd") != cwd or smoke.get("exit_code") != 0:
        fail("fuzz/final/smoke.json: smoke run did not succeed in the recorded cwd")
    smoke_started, smoke_finished = _interval(smoke.get("started_utc"), smoke.get("finished_utc"), "fuzz/final/smoke")
    if smoke_started < build_finished:
        fail("fuzz/final/smoke: smoke started before the fuzzer build finished")
    if smoke.get("binary_before") != binary or smoke.get("binary_after") != binary:
        fail("fuzz/final/smoke: binary custody differs from the fuzzer build")
    if smoke.get("corpus_before") != prepared.get("corpus_before"):
        fail("fuzz/final/smoke: corpus before run differs from preparation")
    expected_smoke_argv = [
        str((work / "minify_xml").resolve()),
        str((work / "corpus").resolve()),
        "-runs=10000",
        "-seed=482",
        "-max_len=65536",
        "-timeout=10",
        f"-artifact_prefix={(work / 'artifacts').resolve()}/",
    ]
    if smoke.get("argv") != expected_smoke_argv:
        fail("fuzz/final/smoke.json: smoke command binding differs")
    if smoke.get("retained") != post_run:
        fail("fuzz/final/smoke.json: retained post-run inventory differs")
    post_run_root = root / f"{FUZZ_DATA_PREFIX}/post-run"
    if not (post_run_root / "corpus").is_dir() or not (post_run_root / "artifacts").is_dir() or not any(name.startswith("corpus/") for name in post_run):
        fail("fuzz/final/smoke.json: corpus and artifact retention is incomplete")
    return {"build_finished": build_finished, "smoke_finished": smoke_finished}


def _check_binary(value: Any, protocol_binary: Mapping[str, Any], label: str, live: bool = False) -> None:
    if not isinstance(value, dict) or value != dict(protocol_binary):
        fail(f"{label}: binary identity differs from frozen protocol")
    if (
        set(value) != {"path", "bytes", "sha256"}
        or not isinstance(value.get("path"), str)
        or not Path(value["path"]).is_absolute()
        or not isinstance(value.get("bytes"), int)
        or isinstance(value.get("bytes"), bool)
        or value["bytes"] <= 0
        or not isinstance(value.get("sha256"), str)
        or SHA256.fullmatch(value["sha256"]) is None
    ):
        fail(f"{label}: binary metadata is malformed")
    if live:
        path = Path(value["path"])
        if not path.is_file() or not os.access(path, os.X_OK):
            fail(f"{label}: executable is not present for live verification")
        current = {"path": str(path), **_meta(path)}
        if current != dict(protocol_binary):
            fail(f"{label}: executable changed after freeze")


def _check_artifacts(
    root: Path,
    receipt: Mapping[str, Any],
    prefix: Path,
    label: str,
    suffixes: tuple[str, ...] = (".stdout", ".stderr", ".resource", ".report.json"),
) -> None:
    artifacts = receipt.get("artifacts")
    if not isinstance(artifacts, dict):
        fail(f"{label}: receipt artifacts missing")
    expected_names = {
        prefix.with_suffix(suffix).name for suffix in suffixes
    }
    actual_names = set(artifacts)
    if not actual_names <= expected_names:
        fail(f"{label}: receipt retains an unknown artifact")
    for name, record in artifacts.items():
        _check_meta(record, prefix.parent / name, f"{label}/{name}")
    if receipt.get("exit_code") == 0 and actual_names != expected_names:
        fail(f"{label}: successful process did not retain all output artifacts")


def _check_pilot_argv(capture: Mapping[str, Any], binary: Mapping[str, Any], label: str) -> None:
    argv = capture.get("argv")
    if not isinstance(argv, list) or len(argv) != 18 or not all(isinstance(value, str) for value in argv):
        fail(f"{label}: successful pilot command vector is malformed")
    expected = [
        "/usr/bin/time", "-v", "-o", None, "/usr/bin/taskset", "-c", str(analyze.CPU),
        binary["path"], "--mode", capture.get("mode"), "--sizes", str(capture.get("size_bytes")),
        "--samples", "1", "--warmup", "1", "--json", None,
    ]
    for position, expected_value in enumerate(expected):
        if expected_value is not None and argv[position] != expected_value:
            fail(f"{label}: pilot command argument {position} differs")
    if not Path(argv[3]).is_absolute() or Path(argv[3]).parent.name != "pilots" or Path(argv[3]).name != f"{label}.resource":
        fail(f"{label}: pilot GNU-time resource path differs")
    if not Path(argv[17]).is_absolute() or Path(argv[17]).parent.name != "pilots" or Path(argv[17]).name != f"{label}.report.json":
        fail(f"{label}: pilot report path differs")


def _check_gate_receipt(
    root: Path,
    label: str,
    protocol: Mapping[str, Any],
    required: bool,
    expected_argv: list[str] | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Check one gate.py started/terminal pair, allowing old dev failures."""

    prefix = root / "validation" / label
    started = read_json(prefix.with_suffix(".started.json"), f"validation/{label}.started")
    terminal = read_json(prefix.with_suffix(".json"), f"validation/{label}")
    if not isinstance(started, dict) or not isinstance(terminal, dict):
        fail(f"validation/{label}: receipt must be an object")
    for field in ("argv", "cwd", "environment", "driver_sha256", "common_sha256", "started_utc", "source_before"):
        if terminal.get(field) != started.get(field):
            fail(f"validation/{label}: terminal changed started {field}")
    _interval(started.get("started_utc"), terminal.get("finished_utc"), f"validation/{label}")
    if not isinstance(terminal.get("driver_sha256"), str) or not SHA256.fullmatch(terminal["driver_sha256"]):
        fail(f"validation/{label}: gate driver digest is malformed")
    if not isinstance(terminal.get("common_sha256"), str) or not SHA256.fullmatch(terminal["common_sha256"]):
        fail(f"validation/{label}: common helper digest is malformed")
    if required and terminal.get("driver_sha256") != sha(root / "gate.py"):
        fail(f"validation/{label}: gate driver digest differs")
    if required and terminal.get("common_sha256") != sha(root / "common.py"):
        fail(f"validation/{label}: common helper digest differs")
    if required and terminal.get("environment") != protocol.get("environment"):
        fail(f"validation/{label}: required gate environment differs from protocol")
    if required and expected_argv is not None and started.get("argv") != expected_argv:
        fail(f"validation/{label}: required gate command differs from run-gates policy")
    _check_snapshot(root, started.get("source_before"), f"validation/{label}.source_before")
    _check_snapshot(root, terminal.get("source_after"), f"validation/{label}.source_after")
    _check_artifacts(root, terminal, prefix, f"validation/{label}", suffixes=(".stdout", ".stderr"))
    if required:
        if terminal.get("exit_code") != 0 or terminal.get("source_unchanged") is not True:
            fail(f"validation/{label}: required gate failed")
        if terminal.get("source_before") != protocol.get("source_snapshot") or terminal.get("source_after") != protocol.get("source_snapshot"):
            fail(f"validation/{label}: required gate source differs from protocol")
    return started, terminal


def _run_gate_commands(root: Path) -> Mapping[str, Any]:
    previous_common = sys.modules.pop("common", None)
    try:
        sys.path.insert(0, str(root))
        try:
            namespace = runpy.run_path(str(root / "run-gates.py"))
        finally:
            sys.path.pop(0)
    except Exception as error:
        fail(f"run-gates.py: cannot load final gate policy: {error}")
    finally:
        if previous_common is not None:
            sys.modules["common"] = previous_common
        else:
            sys.modules.pop("common", None)
    commands = namespace.get("GATES")
    if not isinstance(commands, dict) or not all(
        isinstance(label, str) and isinstance(command, list) and all(isinstance(item, str) for item in command)
        for label, command in commands.items()
    ):
        fail("run-gates.py: GATES policy is malformed")
    return commands


def _check_build_receipt(
    root: Path,
    protocol: Mapping[str, Any],
    instrumentation: str,
    gate_started: Mapping[str, Any],
    gate_terminal: Mapping[str, Any],
    live: bool = False,
) -> tuple[dict[str, Any], _datetime.datetime]:
    path = root / "builds" / f"{instrumentation}.json"
    value = read_json(path, f"builds/{instrumentation}.json")
    if not isinstance(value, dict) or set(value) != REQUIRED_BUILD_FIELDS:
        fail(f"builds/{instrumentation}.json: receipt fields differ")
    label = f"builds/{instrumentation}"
    if value.get("schema") != BUILD_SCHEMA or value.get("instrumentation") != instrumentation:
        fail(f"{label}: schema or instrumentation differs")
    if value.get("source_snapshot") != protocol.get("source_snapshot"):
        fail(f"{label}: source snapshot differs from protocol")
    _check_snapshot(root, value.get("source_snapshot"), f"{label}.source_snapshot")
    argv = value.get("argv")
    if not isinstance(argv, list) or not argv or not all(isinstance(item, str) for item in argv):
        fail(f"{label}.argv: expected command vector")
    if argv[:2] != ["cargo", "build"] or "--release" not in argv or "--locked" not in argv or "--bin" not in argv or "xml_stream_audit" not in argv:
        fail(f"{label}.argv: cargo release build binding differs")
    has_allocator_feature = "allocator-metrics" in argv
    if has_allocator_feature != (instrumentation == "allocator"):
        fail(f"{label}.argv: allocator feature binding differs")
    revision = value.get("git_revision")
    if not isinstance(revision, str) or GIT_REVISION.fullmatch(revision) is None:
        fail(f"{label}.git_revision: expected hexadecimal revision")
    environment = value.get("environment")
    if not isinstance(environment, dict) or set(environment) != set(ENV_KEYS) or not all(isinstance(item, str) for item in environment.values()):
        fail(f"{label}.environment: fields differ")
    if environment != protocol.get("environment"):
        fail(f"{label}.environment: differs from protocol")
    started_at, finished_at = _interval(value.get("started_utc"), value.get("finished_utc"), label)
    gate_started_at, gate_finished_at = _interval(gate_started.get("started_utc"), gate_terminal.get("finished_utc"), f"validation/build-{instrumentation}")
    if not (gate_started_at <= started_at <= finished_at <= gate_finished_at):
        fail(f"{label}: build interval is outside its gate interval")
    binary = value.get("binary")
    expected_binary = protocol.get("binaries", {}).get(instrumentation)
    if not isinstance(binary, dict) or binary != expected_binary:
        fail(f"{label}.binary: executable binding differs from protocol")
    _check_binary(binary, expected_binary, f"{label}.binary", live=live)
    if not isinstance(gate_started.get("argv"), list) or not gate_started["argv"]:
        fail(f"validation/build-{instrumentation}.argv: missing build driver command")
    gate_argv = gate_started["argv"]
    recorded_cwd = gate_started.get("cwd")
    expected_relative_script = Path("docs/performance/results/change-0482/build.py")
    if not isinstance(recorded_cwd, str) or not Path(recorded_cwd).is_absolute():
        fail(f"validation/build-{instrumentation}: recorded cwd is not absolute")
    script_argument = Path(gate_argv[2]) if len(gate_argv) > 2 and isinstance(gate_argv[2], str) else Path("")
    script_matches_recorded_cwd = (
        script_argument == expected_relative_script
        or (
            script_argument.is_absolute()
            and script_argument.resolve() == (Path(recorded_cwd) / expected_relative_script).resolve()
        )
    )
    if (
        len(gate_argv) != 4
        or not all(isinstance(item, str) for item in gate_argv)
        or gate_argv[0] not in ("python3", sys.executable)
        or gate_argv[1] != "-B"
        or not script_matches_recorded_cwd
        or gate_argv[3] != instrumentation
    ):
        fail(f"validation/build-{instrumentation}.argv: build.py binding differs")
    if gate_started.get("source_before") != value.get("source_snapshot") or gate_terminal.get("source_after") != value.get("source_snapshot"):
        fail(f"validation/build-{instrumentation}: source/build binding differs")
    return value, finished_at


def _check_validation_receipts(root: Path, protocol: Mapping[str, Any]) -> dict[str, Any]:
    directory = root / "validation"
    if not directory.is_dir():
        fail("validation directory is missing")
    for started_path in sorted(directory.glob("*.started.json")):
        terminal_path = started_path.with_name(started_path.name.removesuffix(".started.json") + ".json")
        if not terminal_path.is_file():
            fail(f"{started_path.name}: started receipt has no terminal receipt")
    terminals: dict[str, tuple[dict[str, Any], dict[str, Any]]] = {}
    for terminal_path in sorted(directory.glob("*.json")):
        label = terminal_path.stem
        if label.endswith(".started"):
            continue
        terminals[label] = _check_gate_receipt(root, label, protocol, required=False)
    required_labels = protocol.get("required_validation_labels", list(BUILD_LABELS))
    if not isinstance(required_labels, list) or len(set(required_labels)) != len(required_labels) or not all(isinstance(item, str) and item for item in required_labels):
        fail("protocol.required_validation_labels is malformed")
    if set(required_labels) != set(REQUIRED_VALIDATION_LABELS):
        fail("protocol.required_validation_labels bypasses the final gate policy")
    gate_commands = _run_gate_commands(root)
    for label in required_labels:
        if label not in BUILD_LABELS and label not in gate_commands:
            fail(f"run-gates.py: required gate {label} has no command policy")
        expected_argv = None if label in BUILD_LABELS else gate_commands[label]
        if label not in terminals:
            # Re-read through the normal helper to produce the precise missing
            # receipt error and to keep required/optional handling identical.
            terminals[label] = _check_gate_receipt(
                root,
                label,
                protocol,
                required=True,
                expected_argv=expected_argv,
            )
        else:
            _check_gate_receipt(
                root,
                label,
                protocol,
                required=True,
                expected_argv=expected_argv,
            )
    return {"labels": sorted(terminals), "required": sorted(required_labels), "receipts": len(terminals)}


def _check_capture(root: Path, protocol: Mapping[str, Any], spec: Mapping[str, Any], live: bool = False) -> None:
    label = str(spec["label"])
    prefix = root / "captures" / label
    started = read_json(prefix.with_suffix(".started.json"), f"{label}.started")
    terminal = read_json(prefix.with_suffix(".json"), f"{label}.receipt")
    if not isinstance(started, dict) or not isinstance(terminal, dict):
        fail(f"{label}: process receipts must be objects")
    if started.get("schema") != protocol["schema"] or terminal.get("capture") != started.get("capture"):
        fail(f"{label}: started/terminal capture binding differs")
    if started.get("capture") != spec:
        fail(f"{label}: capture specification differs from protocol")
    for field in ("cwd", "environment", "protocol_sha256", "source_before", "binary_before"):
        if terminal.get(field) != started.get(field):
            fail(f"{label}: terminal receipt changed started {field}")
    _check_snapshot(root, started.get("source_before"), f"{label}.source_before")
    _check_snapshot(root, terminal.get("source_after"), f"{label}.source_after")
    if started.get("source_before") != protocol.get("source_snapshot"):
        fail(f"{label}: source differs from frozen protocol snapshot")
    if started.get("source_before") != terminal.get("source_after") or terminal.get("source_unchanged") is not True:
        fail(f"{label}: source changed during retained process")
    binary_spec = protocol["binaries"][spec["instrumentation"]]
    _check_binary(started.get("binary_before"), binary_spec, f"{label}.binary_before", live=live)
    _check_binary(terminal.get("binary_after"), binary_spec, f"{label}.binary_after", live=live)
    if terminal.get("binary_unchanged") is not True:
        fail(f"{label}: binary changed during retained process")
    _check_artifacts(root, terminal, prefix, label)
    if terminal.get("exit_code") != 0:
        fail(f"{label}: formal process exited {terminal.get('exit_code')}")
    if terminal.get("protocol_sha256") != hashlib.sha256((root / "protocol.json").read_bytes()).hexdigest():
        fail(f"{label}: protocol digest differs")


def _check_pilots(
    root: Path,
    protocol: Mapping[str, Any],
    first_formal: _datetime.datetime,
    live: bool = False,
) -> dict[str, Any]:
    directory = root / "pilots"
    attempted: list[str] = []
    failures: list[str] = []
    successful: dict[tuple[str, str, int], list[tuple[_datetime.datetime, dict[str, Any]]]] = {}
    if not directory.exists():
        fail("pilots directory is missing; one successful final pilot is required per lane")
    for started_path in sorted(directory.glob("*.started.json")):
        terminal_path = started_path.with_name(started_path.name.removesuffix(".started.json") + ".json")
        if not terminal_path.is_file():
            fail(f"{started_path.stem}: pilot started receipt has no terminal receipt")
    for terminal_path in sorted(directory.glob("*.json")):
        if terminal_path.name.endswith(".started.json") or terminal_path.name.endswith(".report.json"):
            continue
        terminal = read_json(terminal_path, terminal_path.name)
        if not isinstance(terminal, dict):
            fail(f"{terminal_path}: pilot receipt is not an object")
        label = terminal_path.stem
        prefix = directory / label
        started = read_json(prefix.with_suffix(".started.json"), f"{label}.started")
        if not isinstance(started, dict) or terminal.get("capture") != started.get("capture"):
            fail(f"{label}: pilot started/terminal binding differs")
        for field in ("binary_before", "environment", "cwd"):
            if terminal.get(field) != started.get(field):
                fail(f"{label}: pilot terminal changed started {field}")
        _interval(started.get("started_utc"), terminal.get("finished_utc"), label)
        final_source = started.get("source_before") == protocol.get("source_snapshot")
        _check_snapshot(root, started.get("source_before"), f"{label}.source_before")
        _check_snapshot(root, terminal.get("source_after"), f"{label}.source_after")
        if started.get("source_before") != terminal.get("source_after") or terminal.get("source_unchanged") is not True:
            fail(f"{label}: pilot source changed during process")
        _check_artifacts(root, terminal, prefix, label)
        binary_before = started.get("binary_before")
        if not isinstance(binary_before, dict):
            fail(f"{label}: pilot binary_before is missing")
        _check_binary(binary_before, binary_before, f"{label}.binary_before", live=live)
        binary_after = terminal.get("binary_after")
        if binary_after is not None:
            _check_binary(binary_after, binary_before, f"{label}.binary_after", live=live)
        if terminal.get("binary_unchanged") is not (binary_after == binary_before):
            fail(f"{label}: pilot binary custody flag differs")
        attempted.append(label)
        if terminal.get("exit_code") != 0:
            failures.append(label)
            continue
        capture = terminal.get("capture")
        if not isinstance(capture, dict) or capture.get("instrumentation") not in INSTRUMENTATIONS or capture.get("mode") not in MODES or capture.get("size_bytes") not in SIZES:
            fail(f"{label}: successful pilot lane binding differs")
        if capture.get("samples") != 1 or capture.get("warmups") != 1 or capture.get("repeat") != 0:
            fail(f"{label}: successful pilot protocol differs")
        report_path = prefix.with_suffix(".report.json")
        resource_path = prefix.with_suffix(".resource")
        report = read_json(report_path, f"{label}.report")
        try:
            analyze._validate_report(report, capture, resource_path)
        except analyze.AnalysisError as error:
            fail(f"{label}: pilot report validation failed: {error}")
        _check_pilot_argv(capture, started["binary_before"], label)
        _, finished_at = _interval(started.get("started_utc"), terminal.get("finished_utc"), label)
        if finished_at >= first_formal:
            fail(f"{label}: successful pilot finished after formal captures started")
        if not isinstance(started.get("binary_before"), dict) or terminal.get("binary_after") != started.get("binary_before"):
            fail(f"{label}: successful pilot binary changed during process")
        lane = (capture["instrumentation"], capture["mode"], capture["size_bytes"])
        successful.setdefault(lane, []).append((finished_at, started))
    expected_lanes = {(instrumentation, mode, size) for instrumentation in INSTRUMENTATIONS for mode in MODES for size in SIZES}
    if set(successful) != expected_lanes:
        fail(f"successful final pilots omit lanes: {sorted(expected_lanes - set(successful))}")
    final_pilots: dict[str, str] = {}
    protocol_source = protocol.get("source_snapshot")
    for lane, candidates in successful.items():
        finished_at, started = max(candidates, key=lambda item: item[0])
        if started.get("source_before") != protocol_source:
            fail(f"pilot {lane}: final successful pilot source differs from protocol")
        expected_binary = protocol["binaries"][lane[0]]
        if started.get("binary_before") != expected_binary:
            fail(f"pilot {lane}: final successful pilot binary differs from protocol")
        final_pilots["/".join((lane[0], lane[1], str(lane[2])))] = finished_at.isoformat()
    return {"attempted": len(attempted), "failed": failures, "final_lanes": final_pilots}


def verify_bundle(root: Path = ROOT, live: bool = False) -> dict[str, Any]:
    protocol_path = root / "protocol.json"
    protocol = read_json(protocol_path, "protocol.json")
    if not isinstance(protocol, dict):
        fail("protocol.json is not an object")
    # ``analyze`` uses ROOT to resolve retained captures.  This assignment is
    # also what makes the negative evidence tests operate on a copied bundle.
    old_root = analyze.ROOT
    analyze.ROOT = root
    try:
        try:
            specs = analyze.validate_protocol(protocol)
        except analyze.AnalysisError as error:
            fail(str(error))
        scripts = protocol.get("scripts")
        required_scripts = {"common.py", "gate.py", "build.py", "run-gates.py", "run-measurements.py", "analyze.py", "verify.py", "test_evidence.py", "fuzz.py"}
        if not isinstance(scripts, dict) or set(scripts) != required_scripts:
            fail("protocol script custody differs")
        for name, digest in scripts.items():
            path = root / name
            if not path.is_file() or sha(path) != digest:
                fail(f"protocol script digest differs: {name}")
        source_manifest = _check_snapshot(root, protocol.get("source_snapshot"), "protocol.source_snapshot")
        _check_xml_inventory(root, protocol, source_manifest, live=live)
        fuzz_summary = _check_fuzz(root, protocol, source_manifest, live=live)
        if live:
            from gate import snapshot

            if snapshot() != protocol.get("source_snapshot"):
                fail("live repository source snapshot differs from protocol")
        for instrumentation, binary in protocol["binaries"].items():
            _check_binary(binary, binary, f"protocol.binaries.{instrumentation}", live=live)
        validation_summary = _check_validation_receipts(root, protocol)
        build_receipts = {}
        build_finished = {}
        revisions = set()
        for instrumentation in ("normal", "allocator"):
            build_reference = protocol["builds"][instrumentation]
            build_path = root / build_reference["path"]
            if not build_path.is_file():
                fail(f"protocol.builds.{instrumentation}: retained build receipt is missing")
            if sha(build_path) != build_reference["sha256"]:
                fail(f"protocol.builds.{instrumentation}: receipt digest differs")
            gate_started, gate_terminal = _check_gate_receipt(root, f"build-{instrumentation}", protocol, required=True)
            build, finished = _check_build_receipt(root, protocol, instrumentation, gate_started, gate_terminal, live=live)
            build_receipts[instrumentation] = build
            build_finished[instrumentation] = finished
            revisions.add(build["git_revision"])
        if len(revisions) != 1:
            fail("normal and allocator builds have different git revisions")
        formal_starts = []
        for spec in specs:
            started_path = root / "captures" / f"{spec['label']}.started.json"
            started_record = read_json(started_path, f"{spec['label']}.started")
            formal_starts.append(_timestamp(started_record.get("started_utc"), f"{spec['label']}.started_utc"))
            _check_capture(root, protocol, spec, live=live)
        if not formal_starts:
            fail("formal captures are missing")
        first_formal = min(formal_starts)
        for instrumentation, finished in build_finished.items():
            if finished >= first_formal:
                fail(f"build-{instrumentation}: successful build did not finish before formal captures")
        if fuzz_summary["smoke_finished"] >= first_formal:
            fail("xml-fuzz-smoke: successful smoke did not finish before formal captures")
        pilot_summary = _check_pilots(root, protocol, first_formal, live=live)
        try:
            rows = analyze.load_rows(protocol)
            recomputed = analyze.summarize(protocol, rows)
        except analyze.AnalysisError as error:
            fail(str(error))
        summary_path = root / "summary.json"
        try:
            stored = analyze.read_json(summary_path, "summary.json")
        except analyze.AnalysisError as error:
            fail(str(error))
        if stored != recomputed:
            fail("summary.json differs from data-only recomputation")
        return {
            "schema": "xml-stream-audit-comparison-verification-v1",
            "protocol_sha256": hashlib.sha256(protocol_path.read_bytes()).hexdigest(),
            "summary_sha256": sha(summary_path),
            "formal_captures": len(specs),
            "formal_samples": len(rows) * analyze.SAMPLES,
            "validation": validation_summary,
            "build_revisions": sorted(revisions),
            "fuzz_smoke_finished": fuzz_summary["smoke_finished"].isoformat(),
            "pilot_attempts": pilot_summary["attempted"],
            "pilot_failures": pilot_summary["failed"],
            "source_scope": "same executable/source revision; materialized versus streaming primitive",
            "data_only_recomputed": True,
        }
    finally:
        analyze.ROOT = old_root


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--live", action="store_true", help="also compare the current repository source snapshot")
    args = parser.parse_args()
    try:
        result = verify_bundle(args.root.resolve(), live=args.live)
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (VerificationError, analyze.AnalysisError) as error:
        print(f"verify.py: FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
