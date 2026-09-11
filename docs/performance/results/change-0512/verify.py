#!/usr/bin/env python3
"""Fail-closed replay checks for the bounded 0512 XLSX evidence bundle.

The 0512 capture is a single frozen build.  It is deliberately a diagnostic
and repeatability bundle: the normal lanes cover the four XLSX commit
selectors on three deterministic shapes, while the profile and hardware
lanes describe the same build.  This verifier therefore computes only
same-build repeat drift.  It never compares the 0512 binary with an older
binary or turns an absent metric into zero.

The script consumes the build/capture receipts, report/catalog pairs, raw
Callgrind and perf artifacts, and the source custody manifests.  It never
builds, captures, or writes a summary file; the summary returned on stdout is
intended for the coordinating agent to retain after the remaining gates have
completed.
"""

from __future__ import annotations

import csv
import hashlib
import json
import math
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from tools.summarize_crud_baseline import _validate_elapsed  # noqa: E402
from tools.validate_perf_corpus_binding import validate_binding  # noqa: E402


BASE_REVISION = "18d4d7fe1452bf43614a6efbd5feab4adb5a174f"
PREVIOUS_MANIFEST = HERE.parent / "change-0511" / "source-manifest.json"

CASES = (
    "xlsx_one_cell_commit",
    "xlsx_one_percent_commit",
    "xlsx_one_cell_commit_save",
    "xlsx_one_percent_commit_save",
)
SHAPES = ("tiny", "medium", "dense-wide")
NORMAL_LANES = ("preflight", "normal-r1", "normal-r2")
REPEAT_LANES = ("normal-r1", "normal-r2")
PROFILE_LANES = ("profile-r1", "profile-r2")
SAVED_CASES = {"xlsx_one_cell_commit_save", "xlsx_one_percent_commit_save"}
DRIFT_LIMITS = {"p50": 5.0, "mean": 5.0, "p95": 10.0, "p99": 15.0}
STATISTICS = ("p50", "mean", "p95", "p99")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
CALLGRIND_SUMMARY_RE = re.compile(r"^summary:\s*(\d+)\s*$", re.MULTILINE)
PERF_EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "page-faults",
    "context-switches",
    "cpu-migrations",
)
PERF_GROUP_EVENTS = ("cycles", "instructions", "branches", "branch-misses")
PERF_EVENT_ARGUMENT = "{cycles,instructions,branches,branch-misses},page-faults,context-switches,cpu-migrations"
EXPECTED_COMMIT_PATTERN = "*litchi_xlsx::workbook::edit::semantic::transaction::Edit*::commit"
EXPECTED_RUNNER_PATTERN = "*litchi_perf_baseline::run_xlsx_update_commit"
EXPECTED_REPORT_PREFIX = ["/usr/bin/time", "-v", "taskset", "-c", "2"]


class VerificationError(ValueError):
    """A missing, malformed, or inconsistent evidence item."""


def fail(message: str) -> None:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in result, f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def _reject_constant(value: str) -> None:
    fail(f"non-finite JSON number {value!r}")


def load(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_no_duplicate_pairs,
            parse_constant=_reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"cannot load {path}: {error}")


def sha(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        fail(f"cannot hash {path}: {error}")


def check_hash(value: Any, context: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{context} is not a lowercase SHA-256")
    return value


def canonical(value: Any, context: str = "value") -> str:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError) as error:
        fail(f"{context} is not canonical JSON: {error}")


def finite(value: Any, context: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool),
            f"{context} must be numeric")
    number = float(value)
    require(math.isfinite(number), f"{context} must be finite")
    return number


def regular(path: Path, context: str) -> Path:
    require(path.is_file() and not path.is_symlink(),
            f"{context} is missing, symlinked, or not a regular file")
    return path


def bundle_file(relative: str, context: str) -> Path:
    require(isinstance(relative, str) and relative, f"{context} must be a path")
    path = Path(relative)
    require(not path.is_absolute() and ".." not in path.parts and path.as_posix() == relative,
            f"{context} must be a relative traversal-free path")
    resolved = (HERE / path).resolve()
    require(resolved.is_relative_to(HERE.resolve()), f"{context} escapes the evidence bundle")
    return regular(HERE / path, context)


def repository_file(relative: str, context: str) -> Path:
    require(isinstance(relative, str) and relative, f"{context} must be a path")
    path = Path(relative)
    require(not path.is_absolute() and ".." not in path.parts and path.as_posix() == relative,
            f"{context} must be a relative repository path")
    resolved = (REPO / path).resolve()
    require(resolved.is_relative_to(REPO.resolve()), f"{context} escapes the repository")
    return regular(resolved, context)


def git_blob_sha(revision: str, relative: str) -> str:
    try:
        data = subprocess.check_output(["git", "show", f"{revision}:{relative}"], cwd=REPO)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read {revision}:{relative} from git: {error}")
    return hashlib.sha256(data).hexdigest()


def current_sources() -> dict[str, str]:
    """Reproduce run.py's source manifest without executing any benchmark."""

    try:
        names = subprocess.check_output(
            [
                "git",
                "ls-files",
                "--cached",
                "--others",
                "--exclude-standard",
                "-z",
                "crates",
                "tools/perf-baseline",
                "Cargo.toml",
                "Cargo.lock",
                "rust-toolchain.toml",
                ".cargo",
            ],
            cwd=REPO,
        ).decode("utf-8").split("\0")
    except (OSError, UnicodeError, subprocess.CalledProcessError) as error:
        fail(f"cannot enumerate source manifest: {error}")
    result: dict[str, str] = {}
    for name in sorted(item for item in names if item):
        if Path(name).suffix not in {".rs", ".toml", ".lock"}:
            continue
        result[name] = sha(repository_file(name, f"source {name}"))
    return result


def verify_custody() -> dict[str, Any]:
    """Authenticate the frozen source epoch and non-Rust compiled inputs."""

    plan = load(HERE / "plan.json")
    require(isinstance(plan, dict), "plan.json must be an object")
    require(plan.get("base_revision") == BASE_REVISION,
            "plan base revision does not match the frozen 0511 final")

    manifest = load(HERE / "source-manifest.json")
    require(isinstance(manifest, dict) and manifest, "source-manifest.json must be a non-empty object")
    for relative, value in manifest.items():
        require(isinstance(relative, str) and relative and not Path(relative).is_absolute(),
                f"source manifest path {relative!r} is invalid")
        check_hash(value, f"source manifest {relative}")
    require(manifest == load(PREVIOUS_MANIFEST),
            "0512 source manifest does not preserve the 0511 final source epoch")
    require(manifest == current_sources(),
            "current source tree differs from the source manifest captured before build")
    for relative, value in manifest.items():
        require(sha(repository_file(relative, f"source {relative}")) == value,
                f"source hash changed after build: {relative}")

    require(git_blob_sha(BASE_REVISION, "tools/perf-baseline/src/lib.rs") == manifest["tools/perf-baseline/src/lib.rs"],
            "source manifest does not match the declared base revision")

    adr = load(HERE / "adr-manifest.json")
    require(isinstance(adr, dict), "ADR manifest must be an object")
    require(adr.get("revision") == BASE_REVISION, "ADR manifest revision differs from base")
    files = adr.get("files")
    require(isinstance(files, dict) and len(files) == 30,
            "ADR manifest must contain the 30 reviewed files")
    for relative, value in files.items():
        check_hash(value, f"ADR {relative}")
        require(sha(repository_file(relative, f"ADR {relative}")) == value,
                f"ADR hash changed: {relative}")

    verifier_sources = load(HERE / "verifier-sources.json")
    require(isinstance(verifier_sources, dict) and verifier_sources,
            "verifier-sources.json must be a non-empty object")
    for relative, value in verifier_sources.items():
        check_hash(value, f"verifier source {relative}")
        require(sha(repository_file(relative, f"verifier source {relative}")) == value,
                f"verifier source hash changed: {relative}")

    fixtures = load(HERE / "compile-fixtures.json")
    expected_fixture_paths = {
        "test-data/rtf/watermark.rtf",
        "test-data/poi/test-data/spreadsheet/54016.xls",
    }
    require(isinstance(fixtures, dict) and set(fixtures) == expected_fixture_paths,
            "compile-fixtures.json must bind the two compiled harness fixtures")
    for relative, value in fixtures.items():
        check_hash(value, f"compile fixture {relative}")
        path = repository_file(relative, f"compile fixture {relative}")
        require(sha(path) == value, f"compile fixture hash changed: {relative}")
        require(git_blob_sha(BASE_REVISION, relative) == value,
                f"compile fixture is not from the declared base revision: {relative}")

    return {
        "base_revision": BASE_REVISION,
        "source_files": len(manifest),
        "source_manifest_sha256": sha(HERE / "source-manifest.json"),
        "adr_files": len(files),
        "verifier_sources": {name: value for name, value in sorted(verifier_sources.items())},
        "compile_fixtures": {name: value for name, value in sorted(fixtures.items())},
    }


def verify_build() -> dict[str, Any]:
    receipt = load(HERE / "build-receipt.json")
    require(isinstance(receipt, dict), "build receipt must be an object")
    require(receipt.get("lane") == "build", "build receipt lane is not build")
    require(receipt.get("exit_code") == 0, "build did not exit successfully")
    require(receipt.get("source_unchanged") is True, "build source was not unchanged")
    binary = check_hash(receipt.get("binary_sha256"), "build.binary_sha256")
    manifest_sha = check_hash(receipt.get("source_manifest_sha256"),
                               "build.source_manifest_sha256")
    require(manifest_sha == sha(HERE / "source-manifest.json"),
            "build source-manifest hash is not bound")
    require(receipt.get("log_sha256") == sha(HERE / "build.log"),
            "build log hash is not bound")
    command = receipt.get("command")
    expected = [
        "cargo", "build", "--release", "--bin", "litchi-perf-baseline",
        "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
    ]
    require(command == expected, "build command is not the frozen release harness build")
    finite(receipt.get("elapsed_seconds"), "build.elapsed_seconds")
    require(receipt.get("elapsed_seconds") > 0, "build elapsed time must be positive")
    return {
        "binary_sha256": binary,
        "source_manifest_sha256": manifest_sha,
        "build_log_sha256": sha(HERE / "build.log"),
    }


def command_value(command: list[str], flag: str, context: str) -> str:
    positions = [index for index, item in enumerate(command) if item == flag]
    require(len(positions) == 1 and positions[0] + 1 < len(command),
            f"{context} command must contain one {flag} value")
    value = command[positions[0] + 1]
    require(isinstance(value, str) and value, f"{context} {flag} value is empty")
    return value


def command_option_value(command: list[str], option: str, context: str) -> str:
    """Read either ``--option value`` or the argv-style ``--option=value``."""

    direct = [index for index, item in enumerate(command) if item == option]
    inline = [item[len(option) + 1:] for item in command if item.startswith(option + "=")]
    require(len(direct) + len(inline) == 1, f"{context} command must contain one {option} value")
    if inline:
        value = inline[0]
    else:
        require(direct[0] + 1 < len(command), f"{context} {option} value is missing")
        value = command[direct[0] + 1]
    require(isinstance(value, str) and value, f"{context} {option} value is empty")
    return value


def verify_command_common(command: Any, name: str, samples: int, warmup: int) -> list[str]:
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{name}.command is not an argv list")
    require(command[:5] == EXPECTED_REPORT_PREFIX,
            f"{name} is not pinned to CPU 2 under /usr/bin/time -v")
    binary_positions = [
        index for index, item in enumerate(command)
        if Path(item).name == "litchi-perf-baseline"
    ]
    require(len(binary_positions) == 1,
            f"{name} command does not identify exactly one normal harness binary")
    require(command_value(command, "--samples", name) == str(samples),
            f"{name} command sample count differs")
    require(command_value(command, "--warmup", name) == str(warmup),
            f"{name} command warmup count differs")
    report = command_value(command, "--json", name)
    catalog = command_value(command, "--corpus-manifest", name)
    require(Path(report).name == f"{name}-report.json", f"{name} report output is not bound")
    require(Path(catalog).name == f"{name}-catalog.json", f"{name} catalog output is not bound")
    return command


def verify_artifacts(
    name: str,
    receipt: dict[str, Any],
    build: dict[str, Any],
    required: set[str],
) -> None:
    context = name
    require(receipt.get("exit_code") == 0, f"{context} did not exit successfully")
    require(receipt.get("source_unchanged") is True,
            f"{context} source was not unchanged during capture")
    require(receipt.get("binary_sha256") == build["binary_sha256"],
            f"{context} binary hash does not match build")
    check_hash(receipt.get("binary_sha256"), f"{context}.binary_sha256")
    require(receipt.get("source_manifest_sha256") == build["source_manifest_sha256"],
            f"{context} source manifest hash does not match build")
    check_hash(receipt.get("source_manifest_sha256"), f"{context}.source_manifest_sha256")
    finite(receipt.get("elapsed_seconds"), f"{context}.elapsed_seconds")
    require(receipt.get("elapsed_seconds") > 0, f"{context} elapsed time must be positive")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{context} artifact hash map is missing")
    require(required <= set(artifacts),
            f"{context} receipt does not bind all required artifacts")
    for relative, value in artifacts.items():
        artifact = bundle_file(relative, f"{context} artifact")
        check_hash(value, f"{context} artifact {relative}")
        require(sha(artifact) == value, f"{context} artifact hash changed: {relative}")


def verify_report_identity(
    report: dict[str, Any],
    build: dict[str, Any],
    samples: int,
    warmup: int,
    cases: list[str],
    shapes: list[str],
    context: str,
) -> None:
    require(report.get("schema_version") == 1, f"{context} schema version is not 1")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{context}.tool is missing")
    for key, expected in {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "binary": "litchi-perf-baseline",
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
        "instrumentation": "none",
    }.items():
        require(tool.get(key) == expected, f"{context}.tool.{key} identity differs")
    require(tool.get("allocator_counter_revision") is None,
            f"{context} unexpectedly carries allocator instrumentation identity")

    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{context}.binary_identity is missing")
    require(identity.get("binary_sha256") == build["binary_sha256"],
            f"{context} binary identity does not match build")
    check_hash(identity.get("binary_sha256"), f"{context}.binary_identity.binary_sha256")
    require(isinstance(identity.get("path"), str) and identity["path"],
            f"{context}.binary_identity.path is missing")
    require(Path(identity["path"]).name == "litchi-perf-baseline",
            f"{context}.binary_identity.path does not name the normal harness")
    require(isinstance(identity.get("binary_bytes"), int) and identity["binary_bytes"] > 0,
            f"{context}.binary_identity.binary_bytes is invalid")
    require(identity.get("executable") is True and identity.get("profile") == "release",
            f"{context}.binary_identity executable/profile identity is invalid")

    environment = report.get("environment")
    require(isinstance(environment, dict), f"{context}.environment is missing")
    require(environment.get("git_revision") == BASE_REVISION,
            f"{context}.environment.git_revision is not the frozen base")
    require(isinstance(environment.get("git_worktree_dirty"), bool),
            f"{context}.environment.git_worktree_dirty is not boolean")
    require(environment.get("cpu_affinity") == "2",
            f"{context}.environment.cpu_affinity is not CPU 2")
    require(environment.get("allocator") == "Rust system allocator",
            f"{context} allocator environment identity is not normal")

    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{context}.configuration is missing")
    require(configuration.get("samples_per_case") == samples,
            f"{context} sample count is not {samples}")
    require(configuration.get("warmup_iterations_per_case") == warmup,
            f"{context} warmup count is not {warmup}")
    require(configuration.get("cases") == cases,
            f"{context} case selection differs from the frozen lane")
    require(configuration.get("xlsx_shapes") == shapes,
            f"{context} XLSX shape selection differs from the frozen lane")
    require(configuration.get("filesystem_process_isolated") is True,
            f"{context} is not process isolated")
    require(configuration.get("filesystem_fresh_child_per_sample") is True,
            f"{context} does not use a fresh child per sample")


def verify_sink(value: Any, context: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{context} sink is missing")
    required = {"accepted_bytes", "write_calls", "largest_write", "write_size_buckets"}
    require(set(value) == required, f"{context} sink schema differs")
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        require(isinstance(value[key], int) and not isinstance(value[key], bool) and value[key] >= 0,
                f"{context}.sink.{key} is not a non-negative integer")
    buckets = value["write_size_buckets"]
    require(isinstance(buckets, dict), f"{context}.sink.write_size_buckets is not an object")
    expected_buckets = {
        "bytes_0",
        "bytes_1_to_512",
        "bytes_513_to_4096",
        "bytes_4097_to_16384",
        "bytes_16385_to_65536",
        "bytes_over_65536",
    }
    require(set(buckets) == expected_buckets,
            f"{context}.sink.write_size_buckets schema differs")
    for key, item in buckets.items():
        require(isinstance(item, int) and not isinstance(item, bool) and item >= 0,
                f"{context}.sink.write_size_buckets.{key} is invalid")
    require(sum(buckets.values()) == value["write_calls"],
            f"{context}.sink write buckets do not sum to write_calls")
    return value


def verify_row(row: Any, case: str, shape: str, samples: int, context: str) -> dict[str, Any]:
    require(isinstance(row, dict), f"{context} is not an object")
    require(row.get("case") == case, f"{context}.case differs")
    corpus = row.get("corpus")
    require(isinstance(corpus, dict), f"{context}.corpus is missing")
    require(corpus.get("shape") == shape, f"{context}.corpus.shape differs")
    require(corpus.get("name") == f"xlsx-{shape}", f"{context}.corpus.name differs")
    require(corpus.get("generator") == "litchi-xlsx-synthetic-v1",
            f"{context}.corpus.generator differs")
    require(corpus.get("package_format") == "XLSX/OPC/ZIP",
            f"{context}.corpus.package_format differs")
    check_hash(corpus.get("archive_sha256"), f"{context}.corpus.archive_sha256")
    check_hash(corpus.get("target_payload_sha256"), f"{context}.corpus.target_payload_sha256")
    try:
        elapsed = _validate_elapsed(row, samples, context)
    except (ValueError, TypeError, KeyError) as error:
        fail(f"{context} elapsed vector is invalid: {error}")
    if case in SAVED_CASES:
        verify_sink(row.get("sink"), context)
        operation_metrics = row.get("operation_metrics")
        if operation_metrics is not None:
            require(isinstance(operation_metrics, dict),
                    f"{context}.operation_metrics is not an object")
    else:
        require(row.get("sink") is None, f"{context} unexpectedly publishes sink evidence")
        require(row.get("operation_metrics") is None,
                f"{context} unexpectedly publishes operation metrics")
    require(row.get("source") is None, f"{context} unexpectedly publishes source evidence")
    return {"row": row, "elapsed": elapsed}


def verify_capture(
    name: str,
    build: dict[str, Any],
    samples: int,
    warmup: int,
    cases: list[str],
    shapes: list[str],
    kind: str = "normal",
) -> tuple[dict[str, Any], dict[str, Any]]:
    receipt = load(HERE / f"{name}-receipt.json")
    require(isinstance(receipt, dict), f"{name} receipt must be an object")
    required = {f"{name}-report.json", f"{name}-catalog.json", f"{name}.log"}
    if kind == "profile":
        required.add(f"{name}.out")
        required.add(f"{name}-exclusive.txt")
        required.add(f"{name}-inclusive.txt")
    if kind == "hardware":
        required.add("hardware.csv")
    verify_artifacts(name, receipt, build, required)
    command = verify_command_common(receipt.get("command"), name, samples, warmup)
    if kind == "normal":
        require(receipt.get("scope") == "Native existing per-case timer; setup, generator, expected-output generation, verification and drops excluded according to runner",
                f"{name} timing scope is not the native operation scope")
        require("--case" in command and "--xlsx-shape" in command,
                f"{name} command omits explicit case/shape selections")
    elif kind == "profile":
        require(command[5:7] == ["valgrind", "--tool=callgrind"],
                f"{name} is not a Callgrind invocation")
        require("--collect-atstart=no" in command, f"{name} does not disable collection at start")
        symbols = load(HERE / "profile-symbols.json")
        require(isinstance(symbols, dict), "profile-symbols.json must be an object")
        require(symbols.get("commit") == EXPECTED_COMMIT_PATTERN,
                "profile commit selector is not the exact Edit::commit pattern")
        require(symbols.get("runner") == EXPECTED_RUNNER_PATTERN,
                "profile reset selector is not the exact update runner pattern")
        require(f"--toggle-collect={symbols['commit']}" in command,
                f"{name} does not bind exact Edit::commit collection")
        require(f"--zero-before={symbols['runner']}" in command,
                f"{name} does not reset at exact update runner entry")
        out = command_option_value(command, "--callgrind-out-file", name)
        require(Path(out).name == f"{name}.out", f"{name} Callgrind output is not bound")
        require(receipt.get("scope") == "Commit function only, reset at update-commit runner entry to exclude generator commit; 3 measured commits, no warmups; setup, final readback, save and drop excluded",
                f"{name} profile scope is not explicit")
    elif kind == "hardware":
        require(command[5:7] == ["perf", "stat"], f"{name} is not a perf-stat invocation")
        require(command_value(command, "-o", name).endswith("/hardware.csv"),
                f"{name} perf output is not hardware.csv")
        require(command_value(command, "-e", name) == PERF_EVENT_ARGUMENT,
                f"{name} perf event selection differs")
        require(receipt.get("scope") == "Whole child including fixture generation, opening, staging, 3 warmups, timed commits, final readback and teardown; not operation-local counters",
                f"{name} hardware scope is not explicit whole-child scope")
    else:
        fail(f"unknown capture kind {kind!r}")

    report = load(HERE / f"{name}-report.json")
    catalog = load(HERE / f"{name}-catalog.json")
    require(isinstance(report, dict) and isinstance(catalog, dict),
            f"{name} report/catalog must be objects")
    try:
        validate_binding(report, catalog)
    except Exception as error:
        fail(f"{name} report/catalog binding failed: {error}")
    verify_report_identity(report, build, samples, warmup, cases, shapes, name)
    rows = report.get("results")
    expected_keys = [(shape, case) for shape in shapes for case in cases]
    require(isinstance(rows, list) and len(rows) == len(expected_keys),
            f"{name} has the wrong number of report rows")
    parsed: dict[tuple[str, str], dict[str, Any]] = {}
    for index, ((shape, case), row) in enumerate(zip(expected_keys, rows)):
        parsed[case, shape] = verify_row(row, case, shape, samples, f"{name}.results[{index}]")
    if kind == "profile" or kind == "hardware":
        require(len(rows) == 1, f"{name} diagnostic lane must contain one row")
    return receipt, {"report": report, "catalog": catalog, "rows": parsed}


def verify_diagnostics_identity(
    name: str,
    capture: dict[str, Any],
    expected: dict[str, Any],
    context: str,
) -> None:
    key = ("xlsx_one_percent_commit", "dense-wide")
    row = capture["rows"].get(key)
    require(row is not None, f"{name} is missing the dense one-percent commit row")
    expected_row = expected["rows"][key]["row"]
    current = row["row"]
    require(canonical(current.get("corpus"), context + ".corpus") ==
            canonical(expected_row.get("corpus"), context + ".corpus"),
            f"{context} corpus identity differs from normal baseline")
    require(current.get("sink") == expected_row.get("sink"),
            f"{context} sink identity differs from normal baseline")
    require(current.get("source") is None and current.get("operation_metrics") is None,
            f"{context} unexpectedly publishes non-applicable metrics")


def verify_same_build_identity(
    captures: dict[str, dict[str, Any]],
) -> dict[tuple[str, str], dict[str, Any]]:
    """Require deterministic corpus/sink identities across every normal lane."""

    reference = captures["normal-r1"]["rows"]
    for lane, capture in captures.items():
        for key, item in capture["rows"].items():
            expected = reference[key]["row"]
            current = item["row"]
            context = f"{lane}.{key[0]}.{key[1]}"
            require(canonical(current.get("corpus"), context + ".corpus") ==
                    canonical(expected.get("corpus"), context + ".corpus"),
                    f"{context} corpus identity differs across same-build lanes")
            require(current.get("sink") == expected.get("sink"),
                    f"{context} sink identity differs across same-build lanes")
    return reference


def drift_rows(
    captures: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    left = captures["normal-r1"]["rows"]
    right = captures["normal-r2"]["rows"]
    for shape in SHAPES:
        for case in CASES:
            key = (case, shape)
            a = left[key]["elapsed"]["statistics"]
            b = right[key]["elapsed"]["statistics"]
            changes = {
                metric: (float(b[metric]) / float(a[metric]) - 1.0) * 100.0
                for metric in STATISTICS
            }
            exceeds = {
                metric: abs(changes[metric]) > DRIFT_LIMITS[metric]
                for metric in STATISTICS
            }
            result.append({
                "case": case,
                "shape": shape,
                "repeat_change_percent": changes,
                "drift_limits_percent": dict(DRIFT_LIMITS),
                "exceeds_drift_ceiling": exceeds,
                "any_exceeds_drift_ceiling": any(exceeds.values()),
            })
    return result


def verify_negative_short_vector(captures: dict[str, dict[str, Any]]) -> bool:
    source = captures["normal-r1"]["rows"][("xlsx_one_percent_commit", "dense-wide")]["row"]
    malformed = json.loads(json.dumps(source))
    malformed["elapsed_ns"]["samples"].pop()
    try:
        _validate_elapsed(malformed, 30, "negative short vector")
    except (ValueError, TypeError, KeyError):
        return True
    fail("negative short elapsed vector was accepted")


def verify_profile_annotations(name: str, receipt: dict[str, Any]) -> dict[str, Any]:
    raw = bundle_file(f"{name}.out", f"{name} raw Callgrind output")
    raw_text = raw.read_text(encoding="utf-8", errors="strict")
    summaries = CALLGRIND_SUMMARY_RE.findall(raw_text)
    require(len(summaries) == 1 and int(summaries[0]) > 0,
            f"{name} raw Callgrind output has no nonzero summary")

    inclusive = bundle_file(f"{name}-inclusive.txt", f"{name} inclusive annotation")
    inclusive_text = inclusive.read_text(encoding="utf-8", errors="strict")
    # callgrind_annotate prints the demangled target on a line ending in
    # ``(3x)``.  The end-anchored profile pattern excludes commit closures;
    # the same distinction is retained here when recognizing the annotation.
    target_counts: list[int] = []
    for line in inclusive_text.splitlines():
        # A ``>`` row is a direct callee edge.  ``<`` rows are callers and
        # may repeat the target's aggregate count for every descendant; they
        # do not establish how many times the runner entered commit.
        if ">" not in line:
            continue
        if "Edit::commit" not in line and "Edit>::commit" not in line:
            continue
        if "commit::" in line:
            continue
        target_counts.extend(int(item) for item in re.findall(r"\((\d+)x\)", line))
    require(target_counts == [3],
            f"{name} inclusive annotation does not confirm exactly three Edit::commit calls: {target_counts!r}")
    # The exclusive annotation is retained and receipt-bound for auditability;
    # it is intentionally not interpreted as a source/profile claim here.
    exclusive = bundle_file(f"{name}-exclusive.txt", f"{name} exclusive annotation")
    require(exclusive.stat().st_size > 0, f"{name} exclusive annotation is empty")
    return {
        "summary_ir": int(summaries[0]),
        "commit_calls": 3,
        "raw_sha256": sha(raw),
        "inclusive_sha256": sha(inclusive),
        "exclusive_sha256": sha(exclusive),
        "scope": "Edit::commit collection reset at run_xlsx_update_commit; annotations are diagnostic only",
    }


def verify_attribution_summary() -> dict[str, Any]:
    """Recompute the retained disjoint/overlapping profile attribution."""

    script = bundle_file("analyze.py", "attribution analyzer")
    summary_path = bundle_file("attribution-summary.json", "attribution summary")
    try:
        import importlib.util

        spec = importlib.util.spec_from_file_location("litchi_0512_analyze", script)
        require(spec is not None and spec.loader is not None,
                "cannot load the retained attribution analyzer")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        computed = module.analyze()
    except VerificationError:
        raise
    except Exception as error:
        fail(f"attribution analyzer failed: {error}")
    retained = load(summary_path)
    require(retained == computed,
            "attribution-summary.json differs from recomputation by analyze.py")
    return {
        "sha256": sha(summary_path),
        "analyzer_sha256": sha(script),
        "summary": retained,
    }


def parse_perf_csv(path: Path) -> dict[str, dict[str, Any]]:
    events: dict[str, dict[str, Any]] = {}
    try:
        with path.open(newline="", encoding="utf-8") as stream:
            rows = csv.reader(stream)
            for fields in rows:
                if not fields or not any(item.strip() for item in fields):
                    continue
                if fields[0].lstrip().startswith("#"):
                    continue
                require(len(fields) >= 5, f"hardware.csv event row has fewer than five fields")
                event = fields[2].strip()
                require(event and event not in events,
                        f"hardware.csv has a missing or duplicate event {event!r}")
                raw_count = fields[0].strip()
                require(raw_count not in {"<not counted>", "<not supported>", "<not available>"},
                        f"hardware event {event} was not counted")
                try:
                    count = int(raw_count.replace(",", ""))
                    runtime_ns = int(fields[3].strip().replace(",", ""))
                    running_percent = float(fields[4].strip().rstrip("%"))
                except (TypeError, ValueError) as error:
                    fail(f"hardware event {event} has malformed numeric fields: {error}")
                require(count >= 0 and runtime_ns > 0 and math.isfinite(running_percent),
                        f"hardware event {event} has invalid counter values")
                events[event] = {
                    "count": count,
                    "runtime_ns": runtime_ns,
                    "running_percent": running_percent,
                }
    except OSError as error:
        fail(f"cannot read hardware.csv: {error}")
    require(set(events) == set(PERF_EVENTS),
            f"hardware event set differs: {sorted(events)!r}")
    require(all(events[event]["running_percent"] >= 99.9 for event in PERF_EVENTS),
            "hardware event coverage is below 100% equivalent (99.9% floor)")
    runtimes = {events[event]["runtime_ns"] for event in PERF_GROUP_EVENTS}
    require(len(runtimes) == 1, "grouped hardware events do not share one runtime")
    require(events["cycles"]["count"] > 0 and events["instructions"]["count"] > 0,
            "grouped cycles/instructions counters are zero")
    require(events["branches"]["count"] > 0,
            "grouped branches counter is zero")
    require(events["branch-misses"]["count"] <= events["branches"]["count"],
            "branch misses exceed branches")
    return events


def verify_hardware(
    capture: dict[str, Any],
    expected: dict[str, Any],
) -> dict[str, Any]:
    row = capture["rows"][("xlsx_one_percent_commit", "dense-wide")]
    expected_row = expected["rows"][("xlsx_one_percent_commit", "dense-wide")]
    require(row["row"].get("sink") is None, "hardware diagnostic unexpectedly has a sink")
    require(canonical(row["row"].get("corpus"), "hardware corpus") ==
            canonical(expected_row["row"].get("corpus"), "hardware corpus"),
            "hardware corpus does not match normal dense one-percent commit")
    events = parse_perf_csv(bundle_file("hardware.csv", "hardware perf output"))
    return {
        "events": events,
        "event_count": len(events),
        "group_events": list(PERF_GROUP_EVENTS),
        "group_runtime_ns": events["cycles"]["runtime_ns"],
        "coverage_floor_percent": 99.9,
        "scope": "whole child including fixture generation, opening, staging, warmups, timed commits, final readback and teardown; operation-local counters are not claimed",
        "instrumented_timing_excluded": True,
    }


def verify() -> dict[str, Any]:
    custody = verify_custody()
    build = verify_build()

    normal: dict[str, dict[str, Any]] = {}
    for name, samples, warmup in (
        ("preflight", 1, 0),
        ("normal-r1", 30, 3),
        ("normal-r2", 30, 3),
    ):
        _, capture = verify_capture(
            name, build, samples, warmup, list(CASES), list(SHAPES), "normal"
        )
        normal[name] = capture

    reference = verify_same_build_identity(normal)
    profiles: dict[str, dict[str, Any]] = {}
    profile_annotations: dict[str, dict[str, Any]] = {}
    for name in PROFILE_LANES:
        _, capture = verify_capture(
            name, build, 3, 0, ["xlsx_one_percent_commit"], ["dense-wide"], "profile"
        )
        verify_diagnostics_identity(name, capture, normal["normal-r1"], name)
        profiles[name] = capture
        profile_annotations[name] = verify_profile_annotations(name, load(HERE / f"{name}-receipt.json"))

    _, hardware = verify_capture(
        "hardware", build, 30, 3, ["xlsx_one_percent_commit"], ["dense-wide"], "hardware"
    )
    verify_diagnostics_identity("hardware", hardware, normal["normal-r1"], "hardware")
    hardware_summary = verify_hardware(hardware, normal["normal-r1"])

    drift = drift_rows(normal)
    negative = verify_negative_short_vector(normal)
    attribution = verify_attribution_summary()
    return {
        "schema": "litchi-0512-verification-v1",
        "comparison_scope": "same-build repeat drift only; no cross-build speedup claim",
        "build": build,
        "custody": custody,
        "normal": {
            "lanes": list(NORMAL_LANES),
            "measured_lanes": list(REPEAT_LANES),
            "samples_per_lane": {"preflight": 1, "normal-r1": 30, "normal-r2": 30},
            "warmups_per_lane": {"preflight": 0, "normal-r1": 3, "normal-r2": 3},
            "rows_per_lane": 12,
            "total_rows": 36,
            "corpus_shapes": list(SHAPES),
            "cases": list(CASES),
            "same_build_corpus_sink_identity_verified": True,
        },
        "profile": {
            "lanes": list(PROFILE_LANES),
            "samples_per_lane": 3,
            "warmups_per_lane": 0,
            "rows_per_lane": 1,
            "annotations": profile_annotations,
        },
        "hardware": hardware_summary,
        "attribution": attribution,
        "same_build_drift": {
            "limits_percent": dict(DRIFT_LIMITS),
            "rows": drift,
            "flagged_rows": [item for item in drift if item["any_exceeds_drift_ceiling"]],
        },
        "negative_short_vector_rejected": negative,
        "allocation_metrics": "not collected by the normal XLSX commit selectors",
    }


def main() -> int:
    try:
        result = verify()
    except VerificationError as error:
        print(f"verification failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
