#!/usr/bin/env python3
"""Fail-closed replay verifier for the bounded 0515 XLSX evidence bundle.

0515 keeps one source/binary epoch and measures only same-build repeat drift.
The four Callgrind lanes are diagnostic attribution for changed-output parsing
and worksheet compaction; they do not authorize a performance claim.  This
program reads retained evidence, validates report/catalog bindings and hashes,
then replays the profile parser.  It never builds, captures, or writes a
summary file.
"""

from __future__ import annotations

import copy
import datetime as _datetime
import hashlib
import importlib.util
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


BASE_REVISION = "b698732363250547f05d235b8c8c090f2dd17c59"
PREVIOUS_MANIFEST = HERE.parent / "change-0514" / "before" / "source-manifest.json"
SCRATCH = Path("/tmp/litchi-goal-0515")
CASES = (
    "xlsx_one_cell_commit",
    "xlsx_one_percent_commit",
    "xlsx_one_cell_commit_save",
    "xlsx_one_percent_commit_save",
)
SHAPES = ("tiny", "medium", "dense-wide")
SAVED_CASES = {"xlsx_one_cell_commit_save", "xlsx_one_percent_commit_save"}
NORMAL_LANES = ("preflight", "normal-r1", "normal-r2")
PROFILE_LANES = ("commit-r1", "commit-r2", "compact-r1", "compact-r2")
CAPTURE_ORDER = ("preflight", "normal-r1", "commit-r1", "compact-r1",
                 "commit-r2", "compact-r2", "normal-r2")
STATISTICS = ("p50", "mean", "p95", "p99")
DRIFT_LIMITS = {"p50": 5.0, "mean": 5.0, "p95": 10.0, "p99": 15.0}
ALLOC_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
RSS_RE = re.compile(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", re.MULTILINE)
REPORT_PREFIX = ("/usr/bin/time", "-v", "taskset", "-c", "2")
EXPECTED_COMMIT_PATTERN = "*litchi_xlsx::workbook::edit::semantic::transaction::Edit*::commit"
EXPECTED_COMPACTION_PATTERN = "*litchi_xlsx::raw::compact::changed_worksheet"
EXPECTED_RUNNER_PATTERN = "*litchi_perf_baseline::run_xlsx_update_commit"
NATIVE_SCOPE = (
    "Native existing per-case timer; setup, generator, expected-output generation, "
    "verification and drops excluded according to runner"
)
COMMIT_PROFILE_SCOPE = (
    "Three commit bodies only, reset at runner entry; three caller contexts distinguish "
    "changed-output parser from source Store parser. Setup, readback, save and drops excluded."
)
COMPACTION_PROFILE_SCOPE = (
    "Changed worksheet compaction bodies only; 6 measured calls for 2 changed sheets x 3 "
    "commits, reset at runner entry. Grid validation, setup and readback excluded."
)


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
    except VerificationError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"cannot load {path}: {error}")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


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
    result = float(value)
    require(math.isfinite(result), f"{context} must be finite")
    return result


def regular(path: Path, context: str) -> Path:
    require(path.is_file() and not path.is_symlink(),
            f"{context} is missing, symlinked, or not a regular file")
    return path


def safe_relative(value: Any, context: str) -> Path:
    require(isinstance(value, str) and value, f"{context} is not a path")
    path = Path(value)
    require(not path.is_absolute() and ".." not in path.parts and path.as_posix() == value,
            f"{context} is unsafe")
    return path


def bundle_file(relative: str, context: str) -> Path:
    path = safe_relative(relative, context)
    resolved = (HERE / path).resolve()
    require(resolved.is_relative_to(HERE.resolve()), f"{context} escapes the evidence bundle")
    return regular(HERE / path, context)


def repository_file(relative: str, context: str) -> Path:
    path = safe_relative(relative, context)
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
    """Reproduce run.py's source manifest without building or measuring."""

    try:
        names = subprocess.check_output(
            [
                "git", "ls-files", "--cached", "--others", "--exclude-standard", "-z",
                "crates", "tools/perf-baseline", "Cargo.toml", "Cargo.lock",
                "rust-toolchain.toml", ".cargo",
            ],
            cwd=REPO,
        ).decode("utf-8").split("\0")
    except (OSError, UnicodeError, subprocess.CalledProcessError) as error:
        fail(f"cannot enumerate source manifest: {error}")
    result: dict[str, str] = {}
    for name in sorted(item for item in names if item):
        if Path(name).suffix in {".rs", ".toml", ".lock"}:
            result[name] = sha(repository_file(name, f"source {name}"))
    return result


def validate_manifest(value: Any, context: str) -> dict[str, str]:
    require(isinstance(value, dict) and value, f"{context} must be a non-empty object")
    result: dict[str, str] = {}
    for relative, digest in value.items():
        safe_relative(relative, f"{context} path")
        result[relative] = check_hash(digest, f"{context}.{relative}")
    return result


def verify_custody() -> dict[str, Any]:
    plan = load(HERE / "plan.json")
    require(isinstance(plan, dict) and plan.get("base_revision") == BASE_REVISION,
            "plan base revision does not match the frozen 0514 base")
    require(plan.get("performance_claim") == "none",
            "0515 plan authorizes an unexpected performance claim")
    protocol = plan.get("protocol")
    require(isinstance(protocol, dict), "0515 plan protocol is missing")
    require(protocol.get("profile_commit") and protocol.get("profile_compaction"),
            "0515 plan omits the two profile protocols")

    manifest = validate_manifest(load(HERE / "source-manifest.json"), "source manifest")
    previous = validate_manifest(load(PREVIOUS_MANIFEST), "0514 before source manifest")
    require(manifest == previous, "0515 source manifest differs from the 0514 before epoch")
    current = current_sources()
    require(manifest == current, "current source tree differs from the source manifest")
    for relative, digest in manifest.items():
        require(sha(repository_file(relative, f"source {relative}")) == digest,
                f"source hash changed after capture: {relative}")
        require(git_blob_sha(BASE_REVISION, relative) == digest,
                f"source is not from the declared base revision: {relative}")

    adr = load(HERE / "adr-manifest.json")
    require(isinstance(adr, dict) and adr.get("revision") == BASE_REVISION,
            "ADR manifest revision differs from the frozen base")
    adr_files = adr.get("files")
    require(isinstance(adr_files, dict) and len(adr_files) == 30,
            "ADR manifest must contain the 30 reviewed files")
    for relative, digest in adr_files.items():
        safe_relative(relative, f"ADR {relative}")
        check_hash(digest, f"ADR {relative}")
        require(sha(repository_file(relative, f"ADR {relative}")) == digest,
                f"ADR hash changed: {relative}")

    fixtures = load(HERE / "compile-fixtures.json")
    expected_fixtures = {
        "test-data/rtf/watermark.rtf",
        "test-data/poi/test-data/spreadsheet/54016.xls",
    }
    require(isinstance(fixtures, dict) and set(fixtures) == expected_fixtures,
            "compile-fixtures.json must bind the two compiled fixtures")
    for relative, digest in fixtures.items():
        check_hash(digest, f"fixture {relative}")
        require(sha(repository_file(relative, f"fixture {relative}")) == digest,
                f"fixture hash changed: {relative}")
        require(git_blob_sha(BASE_REVISION, relative) == digest,
                f"fixture is not from the declared base revision: {relative}")

    verifier_sources = load(HERE / "verifier-sources.json")
    require(isinstance(verifier_sources, dict) and verifier_sources,
            "verifier-sources.json must be a non-empty object")
    required_tools = {
        "tools/summarize_crud_baseline.py",
        "tools/validate_perf_corpus_binding.py",
        "tools/generate_corpus_manifest_v2.py",
        "docs/performance/results/change-0515/verify.py",
        "docs/performance/results/change-0515/analyze.py",
    }
    require(required_tools <= set(verifier_sources),
            "verifier-sources.json omits a canonical tool or 0515 verifier")
    for relative, digest in verifier_sources.items():
        safe_relative(relative, f"verifier source {relative}")
        check_hash(digest, f"verifier source {relative}")
        require(sha(repository_file(relative, f"verifier source {relative}")) == digest,
                f"verifier source hash changed: {relative}")

    return {
        "base_revision": BASE_REVISION,
        "source_files": len(manifest),
        "source_manifest_sha256": sha(HERE / "source-manifest.json"),
        "source_epoch": "base",
        "adr_files": len(adr_files),
        "compile_fixtures": {name: digest for name, digest in sorted(fixtures.items())},
        "verifier_sources": {name: digest for name, digest in sorted(verifier_sources.items())},
    }


def verify_build() -> dict[str, Any]:
    receipt = load(HERE / "build-receipt.json")
    require(isinstance(receipt, dict) and receipt.get("lane") == "build",
            "build receipt is not the 0515 build lane")
    require(receipt.get("exit_code") == 0, "build did not exit successfully")
    require(receipt.get("source_unchanged") is True, "build changed the source tree")
    binary = check_hash(receipt.get("binary_sha256"), "build.binary_sha256")
    source_digest = check_hash(receipt.get("source_manifest_sha256"),
                               "build.source_manifest_sha256")
    require(source_digest == sha(HERE / "source-manifest.json"),
            "build source manifest hash is not bound")
    require(receipt.get("log_sha256") == sha(HERE / "build.log"),
            "build log hash is not bound")
    elapsed = finite(receipt.get("elapsed_seconds"), "build.elapsed_seconds")
    require(elapsed > 0, "build elapsed time is not positive")
    command = receipt.get("command")
    expected = [
        "cargo", "build", "--release", "--bin", "litchi-perf-baseline",
        "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
    ]
    require(command == expected, "build command differs from run.py's release build")
    return {
        "binary_sha256": binary,
        "source_manifest_sha256": source_digest,
        "build_log_sha256": sha(HERE / "build.log"),
    }


def option(command: list[str], flag: str, context: str) -> str:
    positions = [index for index, value in enumerate(command) if value == flag]
    inline_prefix = flag + "="
    inline = [value[len(inline_prefix):] for value in command
              if value.startswith(inline_prefix)]
    require(len(positions) + len(inline) == 1,
            f"{context} command must contain one {flag} value")
    if inline:
        value = inline[0]
    else:
        require(positions[0] + 1 < len(command), f"{context} {flag} value is missing")
        value = command[positions[0] + 1]
    require(isinstance(value, str) and value, f"{context} {flag} value is empty")
    return value


def expected_command_values(name: str) -> tuple[str, str, int, int]:
    if name == "preflight":
        return ",".join(CASES), ",".join(SHAPES), 1, 0
    if name == "normal-r1" or name == "normal-r2":
        return ",".join(CASES), ",".join(SHAPES), 30, 3
    if name in PROFILE_LANES:
        return "xlsx_one_percent_commit", "dense-wide", 3, 0
    fail(f"unknown capture lane {name!r}")


def verify_command(command: Any, name: str) -> list[str]:
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{name}.command is not an argv list")
    require(tuple(command[:5]) == REPORT_PREFIX,
            f"{name} command is not pinned to CPU 2 under /usr/bin/time -v")
    case, shape, samples, warmup = expected_command_values(name)
    binaries = [item for item in command
                if Path(item).name == "litchi-perf-baseline"]
    require(len(binaries) == 1, f"{name} command does not identify one harness binary")
    binary_path = Path(binaries[0])
    require(binary_path.is_absolute() and binary_path.as_posix().startswith(SCRATCH.as_posix() + "/"),
            f"{name} command binary is outside the owned scratch role")
    require(option(command, "--case", name) == case, f"{name} case selection differs")
    require(option(command, "--xlsx-shape", name) == shape,
            f"{name} XLSX shape selection differs")
    require(option(command, "--samples", name) == str(samples),
            f"{name} sample count differs")
    require(option(command, "--warmup", name) == str(warmup),
            f"{name} warmup count differs")
    report = option(command, "--json", name)
    catalog = option(command, "--corpus-manifest", name)
    require(Path(report).name == f"{name}-report.json",
            f"{name} report path is not bound")
    require(Path(catalog).name == f"{name}-catalog.json",
            f"{name} catalog path is not bound")
    if name not in PROFILE_LANES:
        require(command[5] == binaries[0], f"{name} normal binary is in the wrong argv position")
        require("valgrind" not in command, f"{name} normal lane unexpectedly uses Valgrind")
    else:
        require(command[5:7] == ["valgrind", "--tool=callgrind"],
                f"{name} is not a Callgrind invocation")
        require("--collect-atstart=no" in command,
                f"{name} does not disable collection at start")
        symbols = load(HERE / "profile-symbols.json")
        require(isinstance(symbols, dict), "profile-symbols.json must be an object")
        commit_symbol = check_symbol(symbols.get("commit"), "profile-symbols.commit")
        runner_symbol = check_symbol(symbols.get("runner"), "profile-symbols.runner")
        compaction_symbol = check_symbol(symbols.get("compaction"), "profile-symbols.compaction")
        expected_symbol = compaction_symbol if name.startswith("compact-") else commit_symbol
        require(f"--toggle-collect={expected_symbol}" in command,
                f"{name} does not bind its exact collection selector")
        require(f"--zero-before={runner_symbol}" in command,
                f"{name} does not reset at the exact runner entry")
        if name.startswith("commit-"):
            require("--separate-callers=3" in command,
                    f"{name} does not bind global three-caller separation")
        else:
            require("--separate-callers=3" not in command,
                    f"{name} unexpectedly changes compaction caller contexts")
        output = [item[len("--callgrind-out-file="):]
                  for item in command if item.startswith("--callgrind-out-file=")]
        require(len(output) == 1 and Path(output[0]).name == f"{name}.out",
                f"{name} Callgrind output path is not bound")
    return command


def check_symbol(value: Any, context: str) -> str:
    require(isinstance(value, str) and value, f"{context} is missing")
    return value


def verify_artifacts(name: str, receipt: dict[str, Any], build: dict[str, Any],
                     required: set[str]) -> None:
    context = name
    require(receipt.get("exit_code") == 0, f"{context} did not exit successfully")
    require(receipt.get("source_unchanged") is True,
            f"{context} changed the source tree")
    require(receipt.get("binary_sha256") == build["binary_sha256"],
            f"{context} binary hash differs from build")
    check_hash(receipt.get("binary_sha256"), f"{context}.binary_sha256")
    require(receipt.get("source_manifest_sha256") == build["source_manifest_sha256"],
            f"{context} source manifest hash differs from build")
    check_hash(receipt.get("source_manifest_sha256"), f"{context}.source_manifest_sha256")
    elapsed = finite(receipt.get("elapsed_seconds"), f"{context}.elapsed_seconds")
    require(elapsed > 0, f"{context} receipt elapsed is not positive")
    require(isinstance(receipt.get("started_utc"), str),
            f"{context} receipt has no start timestamp")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == required,
            f"{context} receipt artifact inventory differs")
    for relative, digest in artifacts.items():
        artifact = bundle_file(relative, f"{context} artifact {relative}")
        check_hash(digest, f"{context} artifact {relative}")
        require(sha(artifact) == digest, f"{context} artifact hash changed: {relative}")


def verify_report_identity(report: dict[str, Any], build: dict[str, Any],
                           name: str, samples: int, warmup: int,
                           cases: list[str], shapes: list[str]) -> None:
    context = name
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
            f"{context} unexpectedly carries allocator instrumentation")

    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{context}.binary_identity is missing")
    require(identity.get("binary_sha256") == build["binary_sha256"],
            f"{context} report binary differs from build")
    check_hash(identity.get("binary_sha256"), f"{context}.binary_identity.binary_sha256")
    path = identity.get("path")
    require(isinstance(path, str) and Path(path).name == "litchi-perf-baseline",
            f"{context}.binary_identity.path is not the normal harness")
    require(isinstance(identity.get("binary_bytes"), int) and identity["binary_bytes"] > 0,
            f"{context}.binary_identity.binary_bytes is invalid")
    require(identity.get("executable") is True and identity.get("profile") == "release",
            f"{context}.binary_identity executable/profile identity is invalid")

    environment = report.get("environment")
    require(isinstance(environment, dict), f"{context}.environment is missing")
    require(environment.get("git_revision") == BASE_REVISION,
            f"{context}.environment.git_revision differs")
    require(isinstance(environment.get("git_worktree_dirty"), bool),
            f"{context}.environment.git_worktree_dirty is not boolean")
    require(environment.get("cpu_affinity") == "2", f"{context} is not pinned to CPU 2")
    require(environment.get("allocator") == "Rust system allocator",
            f"{context}.environment allocator differs")

    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{context}.configuration is missing")
    require(configuration.get("samples_per_case") == samples,
            f"{context} report sample count differs")
    require(configuration.get("warmup_iterations_per_case") == warmup,
            f"{context} report warmup count differs")
    require(configuration.get("cases") == cases, f"{context} case selection differs")
    require(configuration.get("xlsx_shapes") == shapes, f"{context} shape selection differs")
    # These selected XLSX cases use in-process loops, one process per capture.
    # Global filesystem-isolation configuration is not a per-row execution proof.


def verify_sink(value: Any, context: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{context}.sink is missing")
    expected = {"accepted_bytes", "write_calls", "largest_write", "write_size_buckets"}
    require(set(value) == expected, f"{context}.sink schema differs")
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        require(isinstance(value[key], int) and not isinstance(value[key], bool)
                and value[key] >= 0, f"{context}.sink.{key} is invalid")
    buckets = value["write_size_buckets"]
    bucket_names = {
        "bytes_0", "bytes_1_to_512", "bytes_513_to_4096", "bytes_4097_to_16384",
        "bytes_16385_to_65536", "bytes_over_65536",
    }
    require(isinstance(buckets, dict) and set(buckets) == bucket_names,
            f"{context}.sink.write_size_buckets schema differs")
    for key, count in buckets.items():
        require(isinstance(count, int) and not isinstance(count, bool) and count >= 0,
                f"{context}.sink.write_size_buckets.{key} is invalid")
    require(sum(buckets.values()) == value["write_calls"],
            f"{context}.sink buckets do not sum to write_calls")
    return value


def verify_metric_vector(value: Any, samples: int, context: str,
                         allow_pattern: bool = False) -> None:
    require(isinstance(value, dict), f"{context} is not a metric vector")
    status = value.get("status")
    require(status in {"measured", "not_applicable", "unavailable", "overflow"},
            f"{context}.status is invalid")
    require(isinstance(value.get("scope"), str) and value["scope"],
            f"{context}.scope is missing")
    values = value.get("values")
    if status != "measured":
        require(values is None, f"{context} publishes values with status {status}")
        return
    require(isinstance(values, list) and len(values) == samples,
            f"{context}.values must contain {samples} values")
    for index, item in enumerate(values):
        if allow_pattern:
            require(item in {"sequential", "random", "unknown"},
                    f"{context}.values[{index}] is not a read pattern")
        else:
            require(isinstance(item, int) and not isinstance(item, bool) and item >= 0,
                    f"{context}.values[{index}] is not a non-negative integer")


def verify_metric_tree(value: Any, samples: int, context: str) -> None:
    if isinstance(value, dict):
        # MetricVector objects have a values member and a scope.  Other
        # envelopes (source, process, sink) carry status without those fields.
        if "values" in value:
            verify_metric_vector(value, samples, context,
                                 allow_pattern=("pattern" in context))
            return
        if "status" in value:
            require(value["status"] in {"measured", "not_applicable", "unavailable", "overflow"},
                    f"{context}.status is invalid")
        for key, child in value.items():
            if key not in {"status", "scope"}:
                verify_metric_tree(child, samples, f"{context}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            verify_metric_tree(child, samples, f"{context}[{index}]")


def measured_values(value: Any, samples: int, context: str) -> list[int]:
    verify_metric_vector(value, samples, context)
    require(value.get("status") == "measured", f"{context} is not measured")
    return value["values"]


def verify_operation_metrics(row: dict[str, Any], samples: int, context: str,
                             elapsed_order: list[int]) -> dict[str, Any]:
    operation = row.get("operation_metrics")
    require(isinstance(operation, dict), f"{context}.operation_metrics is missing")
    require(operation.get("sample_count") == samples,
            f"{context} operation sample count differs")
    indices = operation.get("sample_indices")
    require(isinstance(indices, list) and sorted(indices) == list(range(samples))
            and len(set(indices)) == samples,
            f"{context} operation sample_indices is not a permutation")
    require(indices == elapsed_order,
            f"{context} operation vectors are not aligned to elapsed samples")
    require(operation.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{context} operation alignment is not explicit")
    require(operation.get("latency_claim") == "comparable_timed_operation",
            f"{context} operation latency claim differs")
    verify_metric_tree(operation, samples, f"{context}.operation_metrics")

    allocation = operation.get("allocation")
    require(isinstance(allocation, dict) and allocation.get("status") == "unavailable",
            f"{context} normal allocation must be unavailable")
    require(allocation.get("scope") == "operation_global_system_allocator",
            f"{context} normal allocation scope differs")
    for field in ALLOC_FIELDS:
        vector = allocation.get(field)
        require(isinstance(vector, dict) and vector.get("status") == "unavailable"
                and vector.get("scope") == "operation_global_system_allocator"
                and vector.get("values") is None,
                f"{context}.allocation.{field} fabricates unavailable values")

    if row["case"] in SAVED_CASES:
        sink = verify_sink(row.get("sink"), context)
        op_sink = operation.get("sink")
        require(isinstance(op_sink, dict) and op_sink.get("write_status") == "measured",
                f"{context} save sink vectors are not measured")
        for name in ("accepted_bytes", "write_calls", "largest_write"):
            values = measured_values(op_sink.get(name), samples,
                                     f"{context}.operation.sink.{name}")
            require(set(values) == {sink[name]},
                    f"{context} operation sink {name} differs from summary")
        buckets = op_sink.get("write_size_buckets")
        require(isinstance(buckets, dict) and buckets.get("status") == "measured",
                f"{context} operation sink buckets are not measured")
        for name, expected in sink["write_size_buckets"].items():
            values = measured_values(buckets.get(name), samples,
                                     f"{context}.operation.sink.write_size_buckets.{name}")
            require(set(values) == {expected},
                    f"{context} operation sink bucket {name} differs from summary")
    else:
        require(row.get("sink") is None, f"{context} non-save row has a sink")
    return operation


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
    except Exception as error:
        fail(f"{context} elapsed vector is invalid: {error}")
    require(row["elapsed_ns"].get("unit") == "ns", f"{context} elapsed unit differs")
    samples_vector = elapsed.get("samples")
    require(isinstance(samples_vector, list)
            and all(isinstance(item, int) and not isinstance(item, bool) and item > 0
                    for item in samples_vector),
            f"{context} elapsed vector contains an invalid value")
    require(row.get("source") is None, f"{context} unexpectedly publishes source evidence")
    if case in SAVED_CASES:
        verify_sink(row.get("sink"), context)
    else:
        require(row.get("sink") is None, f"{context} non-save row has a sink")
    operation = verify_operation_metrics(row, samples, context, elapsed["sample_order"])
    output = row.get("output_sha256")
    if output is not None:
        check_hash(output, f"{context}.output_sha256")
    return {"row": row, "elapsed": elapsed, "operation": operation}


def verify_capture(name: str, build: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    cases, shapes, samples, warmup = expected_command_values(name)
    is_profile = name in PROFILE_LANES
    receipt = load(HERE / f"{name}-receipt.json")
    require(isinstance(receipt, dict), f"{name} receipt must be an object")
    required = {f"{name}-report.json", f"{name}-catalog.json", f"{name}.log"}
    if is_profile:
        required.add(f"{name}.out")
    verify_artifacts(name, receipt, build, required)
    command = verify_command(receipt.get("command"), name)
    expected_scope = COMPACTION_PROFILE_SCOPE if name.startswith("compact-") else (
        COMMIT_PROFILE_SCOPE if name.startswith("commit-") else NATIVE_SCOPE
    )
    require(receipt.get("scope") == expected_scope, f"{name} scope is not explicit")

    report = load(HERE / f"{name}-report.json")
    catalog = load(HERE / f"{name}-catalog.json")
    require(isinstance(report, dict) and isinstance(catalog, dict),
            f"{name} report/catalog must be objects")
    try:
        validate_binding(report, catalog)
    except Exception as error:
        fail(f"{name} report/catalog binding failed: {error}")
    verify_report_identity(report, build, name, samples, warmup,
                           cases.split(","), shapes.split(","))
    rows = report.get("results")
    expected_keys = [(shape, case) for shape in shapes.split(",") for case in cases.split(",")]
    require(isinstance(rows, list) and len(rows) == len(expected_keys),
            f"{name} has the wrong report row count")
    parsed: dict[tuple[str, str], dict[str, Any]] = {}
    for index, ((shape, case), row) in enumerate(zip(expected_keys, rows)):
        parsed[case, shape] = verify_row(row, case, shape, samples,
                                        f"{name}.results[{index}]")
    if is_profile:
        require(len(rows) == 1, f"{name} profile report must contain one selected row")
    return receipt, {"report": report, "catalog": catalog, "rows": parsed, "command": command}


def verify_same_identity(captures: dict[str, dict[str, Any]]) -> None:
    reference = captures["normal-r1"]["rows"]
    for name, capture in captures.items():
        for key, item in capture["rows"].items():
            expected = reference[("xlsx_one_percent_commit", "dense-wide")]
            # Profile captures have only one selected row; normal captures
            # compare their corresponding key.  This keeps corpus/sink
            # identity independent of elapsed vectors and sample counts.
            if key in reference:
                expected = reference[key]
            current = item["row"]
            context = f"{name}/{key[0]}/{key[1]}"
            require(canonical(current.get("corpus"), context + ".corpus") ==
                    canonical(expected["row"].get("corpus"), context + ".corpus"),
                    f"{context} corpus identity differs")
            require(current.get("sink") == expected["row"].get("sink"),
                    f"{context} sink identity differs")


def percent(after: float, before: float) -> float:
    require(before > 0, "cannot compute a percent change from a non-positive value")
    return (after / before - 1.0) * 100.0


def native_drift(captures: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    first = captures["normal-r1"]["rows"]
    second = captures["normal-r2"]["rows"]
    result: list[dict[str, Any]] = []
    for shape in SHAPES:
        for case in CASES:
            key = case, shape
            left = first[key]["elapsed"]["statistics"]
            right = second[key]["elapsed"]["statistics"]
            changes = {metric: percent(float(right[metric]), float(left[metric]))
                       for metric in STATISTICS}
            exceeds = {metric: abs(changes[metric]) > DRIFT_LIMITS[metric]
                       for metric in STATISTICS}
            result.append({
                "case": case,
                "shape": shape,
                "first_statistics_ns": left,
                "second_statistics_ns": right,
                "repeat_change_percent": changes,
                "drift_limits_percent": dict(DRIFT_LIMITS),
                "exceeds_drift_ceiling": exceeds,
                "any_exceeds_drift_ceiling": any(exceeds.values()),
                "scope": "same-build normal repeat drift; no cross-build comparison",
            })
    return result


def parse_rss(path: Path) -> int:
    content = path.read_text(encoding="utf-8", errors="strict")
    values = RSS_RE.findall(content)
    require(len(values) == 1, f"{path.name} must contain exactly one GNU time RSS line")
    value = int(values[0])
    require(value > 0, f"{path.name} reports non-positive RSS")
    return value


def native_rss(captures: dict[str, dict[str, Any]]) -> dict[str, Any]:
    first = parse_rss(HERE / "normal-r1.log")
    second = parse_rss(HERE / "normal-r2.log")
    change = percent(float(second), float(first))
    del captures
    return {
        "first_lane": "normal-r1",
        "second_lane": "normal-r2",
        "first_peak_rss_kib": first,
        "second_peak_rss_kib": second,
        "change_percent": change,
        "drift_limit_percent": 5.0,
        "adverse_over_5_percent": abs(change) > 5.0,
        "scope": "same-build whole-child GNU time RSS for the native normal lanes; no candidate or operation-local RSS claim",
    }


def parse_time(value: Any, context: str) -> _datetime.datetime:
    require(isinstance(value, str) and value, f"{context} start timestamp is missing")
    try:
        return _datetime.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        fail(f"{context} start timestamp is invalid: {error}")


def verify_capture_order(receipts: dict[str, dict[str, Any]]) -> None:
    starts = [parse_time(receipts[name].get("started_utc"), name) for name in CAPTURE_ORDER]
    require(len(set(starts)) == len(starts) and starts == sorted(starts), "capture start order differs")
    for index, name in enumerate(CAPTURE_ORDER[:-1]):
        end = starts[index] + _datetime.timedelta(seconds=finite(receipts[name]["elapsed_seconds"], name))
        require(end <= starts[index + 1], f"{name} overlaps the next capture")
    plan_time = parse_time(load(HERE / "plan.json")["frozen_utc"], "plan")
    build_time = parse_time(load(HERE / "build-receipt.json")["started_utc"], "build")
    require(plan_time <= build_time <= starts[0], "plan/build/capture ordering differs")


def negative_vectors(capture: dict[str, Any]) -> dict[str, bool]:
    row = capture["rows"][("xlsx_one_percent_commit", "dense-wide")]["row"]
    short = copy.deepcopy(row)
    short["elapsed_ns"]["samples"].pop()
    short_rejected = False
    try:
        _validate_elapsed(short, 30, "negative short elapsed vector")
    except Exception:
        short_rejected = True
    require(short_rejected, "negative short elapsed vector was accepted")

    corrupt = copy.deepcopy(row)
    corrupt["elapsed_ns"]["samples"][0] = -1
    corrupt_rejected = False
    try:
        _validate_elapsed(corrupt, 30, "negative corrupt elapsed vector")
    except Exception:
        corrupt_rejected = True
    require(corrupt_rejected, "negative corrupt elapsed vector was accepted")

    operation = copy.deepcopy(row)
    operation["operation_metrics"]["sample_indices"].pop()
    operation_rejected = False
    try:
        elapsed = operation["elapsed_ns"]["sample_order"]
        verify_operation_metrics(operation, 30, "negative operation vector", elapsed)
    except Exception:
        operation_rejected = True
    require(operation_rejected, "negative operation vector was accepted")
    return {
        "short_elapsed_vector_rejected": short_rejected,
        "corrupt_elapsed_vector_rejected": corrupt_rejected,
        "short_operation_vector_rejected": operation_rejected,
    }


def verify_profile_annotations(name: str, receipt: dict[str, Any]) -> dict[str, Any]:
    raw = bundle_file(f"{name}.out", f"{name} raw profile")
    require(receipt["artifacts"].get(raw.name) == sha(raw), f"{name} raw hash differs")
    annotations = load(HERE / f"{name}-annotations.json")
    require(set(annotations) == {"inclusive", "exclusive"}, f"{name} annotation inventory differs")
    results = {}
    for kind, enabled in (("inclusive", "yes"), ("exclusive", "no")):
        item = annotations[kind]
        output = bundle_file(f"{name}-{kind}.txt", f"{name} {kind} annotation")
        command = item.get("command")
        require(isinstance(command, list) and command[:-1] == ["callgrind_annotate", "--inclusive=" + enabled, "--tree=both", "--threshold=100", "--auto=no"], f"{name}/{kind} annotation command differs")
        require(Path(command[-1]).name == raw.name, f"{name}/{kind} raw argument differs")
        require(item.get("exit_code") == 0 and item.get("stderr") == "", f"{name}/{kind} annotation failed")
        require(item.get("output") == output.name and item.get("output_sha256") == sha(output), f"{name}/{kind} output hash differs")
        require(item.get("profile_sha256") == sha(raw), f"{name}/{kind} raw hash differs")
        results[kind] = {"sha256": sha(output), "profile_sha256": sha(raw)}
    replay = load(HERE / "annotation-replay.json")
    rows = [item for item in replay["annotations"] if item["raw"] == raw.name]
    require(len(rows) == 2, f"{name} lacks two independent annotation replays")
    require({item["annotation"] for item in rows} == {f"{name}-inclusive.txt", f"{name}-exclusive.txt"}, f"{name} replay inventory differs")
    for item in rows:
        require(item.get("exit_code") == 0 and item.get("stderr_empty") is True and item.get("function_blocks_and_edges_identical") is True, f"{name} independent annotation replay failed")
        require(item["raw_sha256"] == sha(raw) and item["annotation_sha256"] == sha(HERE / item["annotation"]), f"{name} annotation replay hash differs")
    return {"raw_sha256": sha(raw), "annotations": results, "independent_block_replay_verified": True, "instrumented_rss_excluded": True}


def verify_attribution() -> dict[str, Any]:
    script = bundle_file("analyze.py", "0515 attribution analyzer")
    summary_path = bundle_file("attribution-summary.json", "0515 attribution summary")
    spec = importlib.util.spec_from_file_location("litchi_0515_analyze", script)
    require(spec is not None and spec.loader is not None,
            "cannot load the 0515 attribution analyzer")
    module = importlib.util.module_from_spec(spec)
    try:
        sys.modules[spec.name] = module
        spec.loader.exec_module(module)
        computed = module.analyze()
    except VerificationError:
        raise
    except Exception as error:
        fail(f"0515 attribution analyzer failed: {error}")
    retained = load(summary_path)
    require(retained == computed,
            "attribution-summary.json differs from recomputation by analyze.py")
    return {
        "summary_sha256": sha(summary_path),
        "analyzer_sha256": sha(script),
        "summary": retained,
    }


def verify_scratch_replay(receipts: dict[str, dict[str, Any]]) -> dict[str, Any]:
    expected = str(SCRATCH / "target/release/litchi-perf-baseline")
    for name, receipt in receipts.items():
        command = receipt.get("command")
        require(isinstance(command, list) and command.count(expected) == 1,
                f"{name} command is not bound to the owned scratch binary")
    if Path(expected).exists():
        require(sha(Path(expected)) == load(HERE / "build-receipt.json")["binary_sha256"],
                "live scratch binary differs from captured build")
    return {"owned_scratch": str(SCRATCH), "executable_not_required_for_replay": True,
            "retained_binary_hashes_used": True}


def verify() -> dict[str, Any]:
    custody = verify_custody()
    build = verify_build()
    captures: dict[str, dict[str, Any]] = {}
    receipts: dict[str, dict[str, Any]] = {}
    for name in CAPTURE_ORDER:
        receipt, capture = verify_capture(name, build)
        receipts[name] = receipt
        captures[name] = capture
    verify_capture_order(receipts)

    normal = {name: captures[name] for name in NORMAL_LANES}
    verify_same_identity({name: captures[name] for name in CAPTURE_ORDER})
    profile_annotations = {
        name: verify_profile_annotations(name, receipts[name]) for name in PROFILE_LANES
    }
    attribution = verify_attribution()
    negative = negative_vectors(normal["normal-r1"])
    drift = native_drift(normal)
    rss = native_rss(normal)
    replay = verify_scratch_replay(receipts)

    return {
        "schema": "litchi-0515-verification-v1",
        "performance_claim": "none",
        "claim_authorized": False,
        "comparison_scope": "same-build normal repeat drift and bounded diagnostic Callgrind attribution; no cross-build speedup claim",
        "custody": custody,
        "build": build,
        "native": {
            "lanes": list(NORMAL_LANES),
            "measured_lanes": ["normal-r1", "normal-r2"],
            "preflight_samples": 1,
            "preflight_warmups": 0,
            "samples_per_lane": 30,
            "warmups_per_lane": 3,
            "rows_per_lane": 12,
            "total_rows": 36,
            "measured_elapsed_samples": 720,
            "preflight_elapsed_samples": 12,
            "execution_scope": "One process per capture; iterations execute in process, with a fresh workbook staged outside each timer",
            "cases": list(CASES),
            "shapes": list(SHAPES),
            "drift_limits_percent": dict(DRIFT_LIMITS),
            "same_build_drift": drift,
            "whole_child_rss": rss,
            "rss_is_descriptive_only": True,
        },
        "profiles": {
            "lanes": list(PROFILE_LANES),
            "samples_per_lane": 3,
            "total_diagnostic_iterations": 12,
            "warmups_per_lane": 0,
            "rows_per_lane": 1,
            "annotations": profile_annotations,
            "attribution": attribution,
            "instrumented_elapsed_and_rss_excluded": True,
        },
        "hardware": {
            "status": "not_collected",
            "scope": "current 0515 plan has no hardware-counter lane",
        },
        "negative_vectors": negative,
        "report_catalog_bindings_verified": True,
        "artifact_and_identity_bindings_verified": True,
        "replay": replay,
    }


def main() -> int:
    try:
        print(json.dumps(verify(), indent=2, sort_keys=True))
    except VerificationError as error:
        print(f"verification failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
