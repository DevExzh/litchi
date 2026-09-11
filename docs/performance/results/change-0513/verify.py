#!/usr/bin/env python3
"""Fail-closed replay verifier for the 0513 XLSX metric-enabler bundle.

The bundle has two normal builds, four native ABBA captures (plus one
preflight per build), two allocator-only captures for the candidate, and one
candidate Callgrind capture.  This file only authenticates and replays those
artifacts.  It never builds the harness, runs a benchmark, or treats an
unavailable counter as zero.
"""

from __future__ import annotations

import copy
import datetime as _datetime
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


BASE_REVISION = "53a7c4a523e8518a7150d259ccf3435cb134e4a4"
PREVIOUS_MANIFEST = HERE.parent / "change-0512" / "source-manifest.json"
CASES = (
    "xlsx_one_cell_commit",
    "xlsx_one_percent_commit",
    "xlsx_one_cell_commit_save",
    "xlsx_one_percent_commit_save",
)
SHAPES = ("tiny", "medium", "dense-wide")
SAVED_CASES = {"xlsx_one_cell_commit_save", "xlsx_one_percent_commit_save"}
LANES = ("preflight", "r1", "r2")
ABBA = (("before", "r1"), ("after", "r1"), ("after", "r2"), ("before", "r2"))
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
RSS_RE = re.compile(r"Maximum resident set size \(kbytes\):\s*(\d+)")
CALLGRIND_SUMMARY_RE = re.compile(r"^summary:\s*(\d+)\s*$", re.MULTILINE)
CALLGRIND_FUNCTION_RE = re.compile(r"^\s*[\d,]+\s+\([^)]*\)\s+\*\s+(?P<function>.+?)\s*$")
CALLGRIND_EDGE_RE = re.compile(
    r"^\s*[\d,]+\s+\([^)]*\)\s+>\s+(?P<function>.+?)\s+\((?P<count>[\d,]+)x\)(?:\s+\[[^]]*\])?\s*$"
)
REPORT_PREFIX = ("/usr/bin/time", "-v", "taskset", "-c", "2")
PROFILE_TOGGLE = "--toggle-collect=*litchi_perf_baseline::xlsx_commit_save_operation"
PROFILE_SCOPE = (
    "Three xlsx_commit_save_operation helper calls only; fixture and expected-output "
    "commits/writes excluded; helper returns Commit before caller drop"
)
NATIVE_SCOPE = "Existing native commit or commit+save clock; setup, expected output, sink reservation, oracles and drop excluded"
ALLOC_SCOPE = "Allocation region begins before Instant and ends after elapsed; operation only, no setup/oracles/drop; instrumented elapsed and RSS excluded"


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
        require(key not in result, f"duplicate JSON key {key!r}")
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
    result = float(value)
    require(math.isfinite(result), f"{context} must be finite")
    return result


def regular(path: Path, context: str) -> Path:
    require(path.is_file() and not path.is_symlink(),
            f"{context} is missing, symlinked, or not a regular file")
    return path


def stage_dir(stage: str) -> Path:
    require(stage in {"before", "after"}, f"unknown stage {stage!r}")
    return HERE / stage


def manifest_path(stage: str) -> Path:
    # The control build's custody files are moved below before/.  The
    # candidate build remains at the capture root because capture.py reads it
    # there when it prepares after/ artifacts.
    return HERE / "before" / "source-manifest.json" if stage == "before" else HERE / "source-manifest.json"


def build_receipt_path(stage: str, allocator: bool = False) -> Path:
    if allocator:
        return stage_dir(stage) / "allocator-build-receipt.json"
    return HERE / "before" / "build-receipt.json" if stage == "before" else HERE / "build-receipt.json"


def bundle_file(path: Path, context: str) -> Path:
    root = HERE.resolve()
    resolved = path.resolve()
    require(resolved.is_relative_to(root), f"{context} escapes the evidence bundle")
    return regular(path, context)


def repository_file(relative: str, context: str) -> Path:
    path = Path(relative)
    require(relative and not path.is_absolute() and ".." not in path.parts and path.as_posix() == relative,
            f"{context} is not a safe repository path")
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
        fail(f"cannot enumerate source files: {error}")
    result: dict[str, str] = {}
    for name in sorted(item for item in names if item):
        if Path(name).suffix in {".rs", ".toml", ".lock"}:
            result[name] = sha(repository_file(name, f"source {name}"))
    return result


def verify_sources() -> dict[str, Any]:
    plan = load(HERE / "plan.json")
    require(isinstance(plan, dict) and plan.get("base_revision") == BASE_REVISION,
            "plan base revision does not match the frozen 0512 candidate")
    before = load(manifest_path("before"))
    after = load(manifest_path("after"))
    previous = load(PREVIOUS_MANIFEST)
    for label, manifest in (("before", before), ("after", after), ("previous", previous)):
        require(isinstance(manifest, dict) and manifest, f"{label} source manifest is empty")
        for name, value in manifest.items():
            require(isinstance(name, str) and name and not Path(name).is_absolute(),
                    f"{label} source path is invalid")
            check_hash(value, f"{label} source {name}")
    require(before == previous, "control manifest does not preserve the 0512 source epoch")
    require(after == current_sources(), "candidate source tree differs from its manifest")
    for name, digest in before.items():
        require(git_blob_sha(BASE_REVISION, name) == digest,
                f"control source is not from the declared base revision: {name}")
    for name, digest in after.items():
        require(sha(repository_file(name, f"candidate source {name}")) == digest,
                f"candidate source changed after build: {name}")
    changed = sorted(name for name in set(before) | set(after) if before.get(name) != after.get(name))
    # The focused metric enabler is intentionally isolated to the baseline
    # library and its source test.  Keep the exact paths here because an
    # unrelated harness or library edit would invalidate the comparison.
    expected_names = {
        "tools/perf-baseline/src/lib.rs",
        "tools/perf-baseline/src/xlsx_commit_metrics_tests.rs",
    }
    require(set(changed) == expected_names,
            f"source change set is outside the focused two-file batch: {changed!r}")
    require(all(name.startswith("tools/perf-baseline/") for name in changed),
            f"focused source change is outside perf-baseline: {changed!r}")
    helper_sources = []
    for name in changed:
        text = repository_file(name, f"candidate source {name}").read_text(encoding="utf-8")
        if re.search(r"\bfn\s+xlsx_commit_save_operation\s*\(", text):
            helper_sources.append((name, text))
    require(len(helper_sources) == 1,
            "the exact save helper must be defined in one focused source file")
    helper_name, helper_text = helper_sources[0]
    for fragment in (
        "#[inline(never)]",
        "fn xlsx_commit_save_operation",
        "Result<litchi_xlsx::edit::Commit",
        "let commit = edit.commit()?;",
        "commit.workbook().write_to(sink)?;",
    ):
        require(fragment in helper_text,
                f"save helper source is missing required fragment {fragment!r}")

    adr = load(HERE / "adr-manifest.json")
    require(isinstance(adr, dict) and adr.get("revision") == BASE_REVISION,
            "ADR manifest revision differs from the frozen base")
    files = adr.get("files")
    require(isinstance(files, dict) and len(files) == 30, "ADR manifest must contain 30 files")
    for name, digest in files.items():
        check_hash(digest, f"ADR {name}")
        require(sha(repository_file(name, f"ADR {name}")) == digest, f"ADR changed: {name}")

    verifier_sources = load(HERE / "verifier-sources.json")
    require(isinstance(verifier_sources, dict) and verifier_sources,
            "verifier-sources.json must be a non-empty object")
    for name, digest in verifier_sources.items():
        check_hash(digest, f"verifier source {name}")
        require(sha(repository_file(name, f"verifier source {name}")) == digest,
                f"verifier source changed: {name}")

    fixtures = load(HERE / "compile-fixtures.json")
    expected_fixtures = {"test-data/rtf/watermark.rtf", "test-data/poi/test-data/spreadsheet/54016.xls"}
    require(isinstance(fixtures, dict) and set(fixtures) == expected_fixtures,
            "compile-fixtures.json does not bind the two compiled fixtures")
    for name, digest in fixtures.items():
        check_hash(digest, f"fixture {name}")
        require(sha(repository_file(name, f"fixture {name}")) == digest,
                f"fixture changed: {name}")
        require(git_blob_sha(BASE_REVISION, name) == digest,
                f"fixture is not from the base revision: {name}")
    return {
        "base_revision": BASE_REVISION,
        "before_manifest_sha256": sha(manifest_path("before")),
        "after_manifest_sha256": sha(manifest_path("after")),
        "source_files_before": len(before),
        "source_files_after": len(after),
        "changed": changed,
        "save_helper_source": helper_name,
        "adr_files": len(files),
    }


def verify_build(stage: str, allocator: bool = False) -> dict[str, Any]:
    receipt_path = build_receipt_path(stage, allocator)
    receipt = load(receipt_path)
    label = f"{stage} {'allocator ' if allocator else ''}build"
    require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
            f"{label} did not exit successfully")
    require(receipt.get("source_unchanged") is True, f"{label} source changed during build")
    binary = check_hash(receipt.get("binary_sha256"), f"{label}.binary_sha256")
    source_hash = check_hash(receipt.get("source_manifest_sha256"), f"{label}.source_manifest_sha256")
    require(source_hash == sha(manifest_path(stage)), f"{label} source manifest is not bound")
    log_path = receipt_path.parent / ("allocator-build.log" if allocator else "build.log")
    require(receipt.get("log_sha256") == sha(log_path), f"{label} log hash is not bound")
    finite(receipt.get("elapsed_seconds"), f"{label}.elapsed_seconds")
    require(receipt["elapsed_seconds"] > 0, f"{label} elapsed time is not positive")
    command = receipt.get("command")
    require(isinstance(command, list) and command and all(isinstance(item, str) for item in command),
            f"{label}.command is not an argv list")
    expected_normal = [
        "cargo", "build", "--release", "--bin", "litchi-perf-baseline",
        "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
    ]
    expected_allocator = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--features", "allocator-metrics",
        "--bin", "litchi-perf-baseline-alloc",
    ]
    require(command == (expected_allocator if allocator else expected_normal),
            f"{label} command differs from the frozen build")
    return {"binary_sha256": binary, "source_manifest_sha256": source_hash,
            "log_sha256": sha(log_path), "allocator": allocator}


def option(command: list[str], flag: str, context: str) -> str:
    positions = [index for index, value in enumerate(command) if value == flag]
    require(len(positions) == 1 and positions[0] + 1 < len(command),
            f"{context} command must contain one {flag} value")
    value = command[positions[0] + 1]
    require(isinstance(value, str) and value, f"{context} {flag} value is empty")
    return value


def equals_option(command: list[str], flag: str, context: str) -> str:
    """Read one ``--flag=value`` argv item without accepting duplicates."""
    prefix = flag + "="
    values = [item[len(prefix):] for item in command if item.startswith(prefix)]
    require(flag not in command and len(values) == 1 and values[0],
            f"{context} command must contain one {prefix} value")
    return values[0]


def verify_artifacts(stage: str, name: str, receipt: dict[str, Any], build: dict[str, Any],
                     required: set[str]) -> None:
    context = f"{stage}/{name}"
    require(receipt.get("exit_code") == 0, f"{context} did not exit successfully")
    require(receipt.get("source_unchanged") is True, f"{context} source changed during capture")
    require(receipt.get("binary_sha256") == build["binary_sha256"],
            f"{context} binary hash differs from build")
    check_hash(receipt.get("binary_sha256"), f"{context}.binary_sha256")
    require(receipt.get("source_manifest_sha256") == build["source_manifest_sha256"],
            f"{context} source manifest differs from build")
    check_hash(receipt.get("source_manifest_sha256"), f"{context}.source_manifest_sha256")
    finite(receipt.get("elapsed_seconds"), f"{context}.elapsed_seconds")
    require(receipt["elapsed_seconds"] > 0, f"{context} receipt elapsed is not positive")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and required <= set(artifacts),
            f"{context} receipt does not bind all required artifacts")
    for relative, digest in artifacts.items():
        path = Path(relative)
        require(relative and not path.is_absolute() and ".." not in path.parts and path.as_posix() == relative,
                f"{context} artifact path is unsafe: {relative!r}")
        artifact = bundle_file(stage_dir(stage) / path, f"{context} artifact {relative}")
        check_hash(digest, f"{context} artifact {relative}")
        require(sha(artifact) == digest, f"{context} artifact hash changed: {relative}")


def verify_command(command: Any, name: str, samples: int, warmup: int,
                   allocator: bool = False, profile: bool = False) -> list[str]:
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{name}.command is not an argv list")
    require(tuple(command[:5]) == REPORT_PREFIX, f"{name} is not pinned to CPU 2")
    basename = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    binaries = [item for item in command if Path(item).name == basename]
    require(len(binaries) == 1, f"{name} command does not identify one {basename}")
    require(option(command, "--samples", name) == str(samples), f"{name} sample count differs")
    require(option(command, "--warmup", name) == str(warmup), f"{name} warmup count differs")
    report = option(command, "--json", name)
    catalog = option(command, "--corpus-manifest", name)
    require(Path(report).name == f"{name}-report.json", f"{name} report path is not bound")
    require(Path(catalog).name == f"{name}-catalog.json", f"{name} catalog path is not bound")
    if profile:
        require(command[5:7] == ["valgrind", "--tool=callgrind"],
                f"{name} is not a Callgrind command")
        require("--collect-atstart=no" in command and PROFILE_TOGGLE in command,
                f"{name} does not bind the exact save-helper profile boundary")
        output = equals_option(command, "--callgrind-out-file", name)
        require(Path(output).name == f"{name}.out", f"{name} Callgrind output path is not bound")
    else:
        require("--case" in command and "--xlsx-shape" in command,
                f"{name} omits explicit XLSX case/shape selectors")
    return command


def verify_report_identity(report: dict[str, Any], build: dict[str, Any], samples: int,
                           warmup: int, cases: list[str], shapes: list[str],
                           context: str, allocator: bool = False) -> None:
    require(report.get("schema_version") == 1, f"{context} schema version is not 1")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{context}.tool is missing")
    expected = {
        "name": "litchi-perf-baseline", "version": "0.1.0",
        "binary": "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline",
        "profile": "release",
        "target_os": "linux", "target_arch": "x86_64",
        "instrumentation": "system_allocator_operation_scoped" if allocator else "none",
        "allocator_counter_revision": "serialized_region_peak_v3" if allocator else None,
    }
    for key, value in expected.items():
        require(tool.get(key) == value, f"{context}.tool.{key} identity differs")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{context}.binary_identity is missing")
    require(identity.get("binary_sha256") == build["binary_sha256"],
            f"{context} report binary does not match build")
    check_hash(identity.get("binary_sha256"), f"{context}.binary_identity.binary_sha256")
    require(Path(str(identity.get("path"))).name == expected["binary"],
            f"{context}.binary_identity path does not name the selected binary")
    require(identity.get("executable") is True and identity.get("profile") == "release",
            f"{context}.binary_identity executable/profile is invalid")
    require(isinstance(identity.get("binary_bytes"), int) and identity["binary_bytes"] > 0,
            f"{context}.binary_identity.binary_bytes is invalid")
    environment = report.get("environment")
    require(isinstance(environment, dict), f"{context}.environment is missing")
    require(environment.get("git_revision") == BASE_REVISION,
            f"{context}.environment.git_revision differs")
    require(isinstance(environment.get("git_worktree_dirty"), bool),
            f"{context}.environment.git_worktree_dirty is not boolean")
    require(environment.get("cpu_affinity") == "2", f"{context} is not pinned to CPU 2")
    expected_allocator = "CountingSystemAllocator(std::alloc::System)" if allocator else "Rust system allocator"
    require(environment.get("allocator") == expected_allocator,
            f"{context}.environment allocator identity differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{context}.configuration is missing")
    require(configuration.get("samples_per_case") == samples,
            f"{context} report sample count differs")
    require(configuration.get("warmup_iterations_per_case") == warmup,
            f"{context} report warmup count differs")
    require(configuration.get("cases") == cases, f"{context} case selection differs")
    require(configuration.get("xlsx_shapes") == shapes, f"{context} shape selection differs")
    require(configuration.get("filesystem_process_isolated") is True,
            f"{context} is not process isolated")
    require(configuration.get("filesystem_fresh_child_per_sample") is True,
            f"{context} does not use a fresh child per sample")


def verify_sink(value: Any, context: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{context}.sink is missing")
    required = {"accepted_bytes", "write_calls", "largest_write", "write_size_buckets"}
    require(set(value) == required, f"{context}.sink schema differs")
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        require(isinstance(value[key], int) and not isinstance(value[key], bool) and value[key] >= 0,
                f"{context}.sink.{key} is invalid")
    buckets = value["write_size_buckets"]
    expected = {"bytes_0", "bytes_1_to_512", "bytes_513_to_4096", "bytes_4097_to_16384",
                "bytes_16385_to_65536", "bytes_over_65536"}
    require(isinstance(buckets, dict) and set(buckets) == expected,
            f"{context}.sink.write_size_buckets schema differs")
    for key, count in buckets.items():
        require(isinstance(count, int) and not isinstance(count, bool) and count >= 0,
                f"{context}.sink.write_size_buckets.{key} is invalid")
    require(sum(buckets.values()) == value["write_calls"],
            f"{context}.sink write buckets do not sum to write_calls")
    return value


def verify_metric_vector(value: Any, samples: int, context: str, allow_pattern: bool = False) -> None:
    require(isinstance(value, dict), f"{context} must be a metric vector")
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
        # AllocationMetrics and the other metric envelopes also have
        # ``status``/``scope`` fields, but only a three-key object is a
        # MetricVector.  Classifying an envelope as a vector would incorrectly
        # require a top-level ``values`` member.
        if "status" in value and "scope" in value and set(value) <= {"status", "scope", "values"}:
            verify_metric_vector(value, samples, context,
                                 allow_pattern=("pattern" in context))
            return
        for key, child in value.items():
            verify_metric_tree(child, samples, f"{context}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            verify_metric_tree(child, samples, f"{context}[{index}]")


def verify_operation_metrics(row: dict[str, Any], samples: int, context: str,
                             allocator: bool, elapsed_order: list[int] | None = None,
                             require_normal_allocation: bool = False) -> dict[str, Any]:
    operation = row.get("operation_metrics")
    require(isinstance(operation, dict), f"{context}.operation_metrics is missing")
    require(operation.get("sample_count") == samples, f"{context} operation sample count differs")
    indices = operation.get("sample_indices")
    require(isinstance(indices, list) and sorted(indices) == list(range(samples)) and len(set(indices)) == samples,
            f"{context} operation sample_indices is not a permutation")
    if elapsed_order is not None:
        require(indices == elapsed_order, f"{context} operation vectors are not aligned to elapsed samples")
    require(operation.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{context} operation alignment is not explicit")
    verify_metric_tree(operation, samples, f"{context}.operation_metrics")
    allocation = operation.get("allocation")
    if allocator:
        require(isinstance(allocation, dict) and allocation.get("status") == "measured",
                f"{context} allocator metrics are not measured")
        require(allocation.get("scope") == "operation_global_system_allocator",
                f"{context} allocator scope differs")
        values: dict[str, list[int]] = {}
        for field in ALLOC_FIELDS:
            vector = allocation.get(field)
            require(isinstance(vector, dict) and vector.get("status") == "measured",
                    f"{context} allocator field {field} is not measured")
            raw = vector.get("values")
            require(isinstance(raw, list) and len(raw) == samples,
                    f"{context} allocator field {field} has wrong cardinality")
            require(all(isinstance(item, int) and not isinstance(item, bool) and item >= 0 for item in raw),
                    f"{context} allocator field {field} has invalid values")
            values[field] = raw
        increments = [peak - base for peak, base in zip(
            values["region_peak_live_bytes"], values["live_bytes_before"])]
        require(min(increments) >= 0, f"{context} region peak precedes live-byte baseline")
        require(all(
            region >= after >= 0 and region <= peak_after
            for region, after, peak_after in zip(
                values["region_peak_live_bytes"],
                values["live_bytes_after"],
                values["peak_live_bytes_after"],
            )
        ), f"{context} region peak is outside live/high-water bounds")
        return {
            "status": allocation["status"],
            "scope": allocation["scope"],
            "constant_fields": {
                field: sorted(set(values[field])) for field in ALLOC_FIELDS
            },
            "incremental_region_peak_range_bytes": [min(increments), max(increments)],
        }
    if require_normal_allocation:
        require(isinstance(allocation, dict),
                f"{context} normal allocation status is missing")
    if allocation is not None:
        require(isinstance(allocation, dict) and allocation.get("status") == "unavailable",
                f"{context} normal allocation must be unavailable")
        require(allocation.get("scope") == "operation_global_system_allocator",
                f"{context} normal allocation scope differs")
        for field in ALLOC_FIELDS:
            vector = allocation.get(field)
            require(isinstance(vector, dict) and vector.get("status") == "unavailable",
                    f"{context} normal allocation field {field} is not unavailable")
            require(vector.get("scope") == "operation_global_system_allocator",
                    f"{context} normal allocation field {field} scope differs")
            require(vector.get("values") is None,
                    f"{context} normal allocation field {field} fabricated values")
    return {"status": "unavailable" if allocation is not None else "absent"}


def verify_row(row: Any, case: str, shape: str, samples: int, context: str,
               allocator: bool, require_operation: bool = False,
               align_operation: bool = False,
               require_normal_allocation: bool = False) -> dict[str, Any]:
    require(isinstance(row, dict), f"{context} is not an object")
    require(row.get("case") == case, f"{context}.case differs")
    corpus = row.get("corpus")
    require(isinstance(corpus, dict) and corpus.get("shape") == shape,
            f"{context}.corpus shape differs")
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
    if case in SAVED_CASES:
        verify_sink(row.get("sink"), context)
    else:
        require(row.get("sink") is None, f"{context} non-save row has a sink")
    require(row.get("source") is None, f"{context} unexpectedly publishes source evidence")
    operation = row.get("operation_metrics")
    if require_operation:
        require(isinstance(operation, dict), f"{context} operation metrics are missing")
    if operation is not None:
        allocation = verify_operation_metrics(row, samples, context, allocator,
                                              elapsed["sample_order"] if (allocator or align_operation) else None,
                                              require_normal_allocation=require_normal_allocation)
    else:
        require(not allocator, f"{context} allocator row has no operation metrics")
        allocation = {"status": "absent"}
    output = row.get("output_sha256")
    if output is not None:
        check_hash(output, f"{context}.output_sha256")
    return {"row": row, "elapsed": elapsed, "allocation": allocation}


def verify_capture(stage: str, lane: str, build: dict[str, Any], samples: int,
                   warmup: int, allocator: bool = False, profile: bool = False) -> tuple[dict[str, Any], dict[str, Any]]:
    name = "allocator-" + lane if allocator else lane
    directory = stage_dir(stage)
    receipt = load(directory / f"{name}-receipt.json")
    required = {f"{name}-report.json", f"{name}-catalog.json", f"{name}.log"}
    if profile:
        required.add(f"{name}.out")
        # Annotation files are retained and receipt-bound by the sealing step.
        required.update({f"{name}-inclusive.txt", f"{name}-exclusive.txt"})
    verify_artifacts(stage, name, receipt, build, required)
    command = verify_command(receipt.get("command"), name, samples, warmup, allocator, profile)
    expected_scope = ALLOC_SCOPE if allocator else (PROFILE_SCOPE if profile else NATIVE_SCOPE)
    require(receipt.get("scope") == expected_scope,
            f"{stage}/{name} scope is not explicit")
    report = load(directory / f"{name}-report.json")
    catalog = load(directory / f"{name}-catalog.json")
    require(isinstance(report, dict) and isinstance(catalog, dict),
            f"{stage}/{name} report/catalog must be objects")
    try:
        validate_binding(report, catalog)
    except Exception as error:
        fail(f"{stage}/{name} report/catalog binding failed: {error}")
    report_cases = ["xlsx_one_percent_commit_save"] if profile else list(CASES)
    report_shapes = ["dense-wide"] if profile else list(SHAPES)
    verify_report_identity(report, build, samples, warmup, report_cases, report_shapes,
                           f"{stage}/{name}", allocator)
    rows = report.get("results")
    expected_keys = [(shape, case) for shape in report_shapes for case in report_cases]
    require(isinstance(rows, list) and len(rows) == len(expected_keys),
            f"{stage}/{name} has the wrong report row count")
    parsed: dict[tuple[str, str], dict[str, Any]] = {}
    for index, ((shape, case), row) in enumerate(zip(expected_keys, rows)):
        parsed[case, shape] = verify_row(
            row, case, shape, samples, f"{stage}/{name}.results[{index}]", allocator,
            require_operation=(allocator or stage == "after"),
            align_operation=(allocator or stage == "after"),
            require_normal_allocation=(stage == "after" and not allocator),
        )
    if profile:
        require(len(rows) == 1, f"{stage}/{name} profile report must contain one selected row")
        selected = parsed.get(("xlsx_one_percent_commit_save", "dense-wide"))
        require(selected is not None, f"{stage}/{name} lacks the selected dense save row")
        annotation = verify_profile_annotations(stage, name, receipt)
    else:
        annotation = None
    return receipt, {"report": report, "catalog": catalog, "rows": parsed, "annotations": annotation, "command": command}


def verify_profile_annotations(stage: str, name: str, receipt: dict[str, Any]) -> dict[str, Any]:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{stage}/{name} artifact map is missing")
    raw = bundle_file(stage_dir(stage) / f"{name}.out", f"{stage}/{name} raw Callgrind")
    text = raw.read_text(encoding="utf-8", errors="strict")
    summaries = CALLGRIND_SUMMARY_RE.findall(text)
    require(len(summaries) == 1 and int(summaries[0]) > 0,
            f"{stage}/{name} raw Callgrind summary is missing")
    inclusive = bundle_file(stage_dir(stage) / f"{name}-inclusive.txt", f"{stage}/{name} inclusive annotation")
    exclusive = bundle_file(stage_dir(stage) / f"{name}-exclusive.txt", f"{stage}/{name} exclusive annotation")
    inclusive_text = inclusive.read_text(encoding="utf-8", errors="strict")
    require(inclusive.stat().st_size > 0 and exclusive.stat().st_size > 0,
            f"{stage}/{name} profile annotations are empty")
    # callgrind_annotate renders a function record as a ``*`` line followed
    # by its direct callee ``>`` lines.  The callee edge is on its own line;
    # looking for helper and commit names on one line would therefore accept
    # an unrelated caller aggregate.  ``<`` rows are deliberately ignored.
    edges: list[tuple[str, str, int]] = []
    current_function: str | None = None
    for line in inclusive_text.splitlines():
        function_match = CALLGRIND_FUNCTION_RE.match(line)
        if function_match:
            current_function = function_match.group("function").strip()
            continue
        edge_match = CALLGRIND_EDGE_RE.match(line)
        if edge_match and current_function is not None:
            edges.append((
                current_function,
                edge_match.group("function").strip(),
                int(edge_match.group("count").replace(",", "")),
            ))
    helper_to_commit = [count for parent, child, count in edges
                        if "xlsx_commit_save_operation" in parent and "Edit::commit" in child]
    runner_to_helper = [count for parent, child, count in edges
                        if "run_xlsx_update_commit_save" in parent
                        and "xlsx_commit_save_operation" in child]
    writer_tokens = ("PackageWriter", "CountingSink", "Workbook::write_to", "Writer<")
    helper_to_writer = [count for parent, child, count in edges
                        if "xlsx_commit_save_operation" in parent
                        and any(token in child for token in writer_tokens)]
    require(helper_to_commit == [3],
            f"{stage}/{name} annotation does not prove three direct helper-to-commit calls: {helper_to_commit!r}")
    require(runner_to_helper == [3],
            f"{stage}/{name} annotation does not prove three direct runner-to-helper calls: {runner_to_helper!r}")
    # The exact writer symbol is intentionally not frozen: generic write_to
    # can inline.  When callgrind retains a writer/package/sink edge, it must
    # retain the three helper invocations.  The source-level helper check
    # above proves that the write is still part of the helper when the generic
    # method is inlined away.
    require(any(count == 3 for count in helper_to_writer),
            f"{stage}/{name} annotation has no three-call retained writer subtree")
    return {
        "summary_ir": int(summaries[0]),
        "commit_direct_calls": 3,
        "runner_to_helper_calls": 3,
        "helper_to_writer_calls": 3,
        "raw_sha256": sha(raw),
        "inclusive_sha256": sha(inclusive),
        "exclusive_sha256": sha(exclusive),
        "scope": PROFILE_SCOPE,
    }


def verify_row_identity(current: dict[str, Any], expected: dict[str, Any], context: str) -> None:
    require(canonical(current.get("corpus"), context + ".corpus") ==
            canonical(expected.get("corpus"), context + ".corpus"),
            f"{context} corpus identity differs")
    require(current.get("sink") == expected.get("sink"),
            f"{context} sink identity differs")
    require(current.get("output_sha256") == expected.get("output_sha256"),
            f"{context} output identity differs")


def capture_identity(captures: dict[tuple[str, str], dict[str, Any]], context: str) -> None:
    reference = captures["before", "r1"]["rows"]
    for (stage, lane), capture in captures.items():
        for key, current in capture["rows"].items():
            expected = reference[key]["row"]
            verify_row_identity(
                current["row"], expected,
                f"{context}/{stage}/{lane}/{key[0]}/{key[1]}",
            )


def percent(after: float, before: float) -> float:
    require(before > 0, "cannot compute a percent change from a non-positive value")
    return (after / before - 1.0) * 100.0


def native_comparisons(captures: dict[tuple[str, str], dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    pairs: list[dict[str, Any]] = []
    drift: list[dict[str, Any]] = []
    for repeat in ("r1", "r2"):
        before = captures["before", repeat]["rows"]
        after = captures["after", repeat]["rows"]
        for shape in SHAPES:
            for case in CASES:
                key = (case, shape)
                b = before[key]["elapsed"]["statistics"]
                a = after[key]["elapsed"]["statistics"]
                changes = {metric: percent(float(a[metric]), float(b[metric])) for metric in STATISTICS}
                throughput = percent(float(b["mean"]), float(a["mean"]))
                adverse = any(changes[metric] > 5.0 for metric in STATISTICS) or throughput < -5.0
                pairs.append({
                    "repeat": repeat, "case": case, "shape": shape,
                    "before_statistics_ns": b, "after_statistics_ns": a,
                    "latency_change_percent": changes,
                    "throughput_change_percent": throughput,
                    "adverse_latency_or_throughput_over_5_percent": adverse,
                })
    for stage in ("before", "after"):
        first = captures[stage, "r1"]["rows"]
        second = captures[stage, "r2"]["rows"]
        for shape in SHAPES:
            for case in CASES:
                key = (case, shape)
                left = first[key]["elapsed"]["statistics"]
                right = second[key]["elapsed"]["statistics"]
                changes = {metric: percent(float(right[metric]), float(left[metric])) for metric in STATISTICS}
                exceeds = {metric: abs(changes[metric]) > DRIFT_LIMITS[metric] for metric in STATISTICS}
                drift.append({
                    "stage": stage, "case": case, "shape": shape,
                    "repeat_change_percent": changes,
                    "drift_limits_percent": dict(DRIFT_LIMITS),
                    "exceeds_drift_ceiling": exceeds,
                    "any_exceeds_drift_ceiling": any(exceeds.values()),
                })
    return pairs, drift, []


def rss(path: Path) -> int:
    text = path.read_text(encoding="utf-8", errors="strict")
    values = RSS_RE.findall(text)
    require(len(values) == 1, f"{path} must contain one whole-child RSS line")
    result = int(values[0])
    require(result > 0, f"{path} reports non-positive RSS")
    return result


def native_rss() -> list[dict[str, Any]]:
    result = []
    for repeat in ("r1", "r2"):
        before = rss(stage_dir("before") / f"{repeat}.log")
        after = rss(stage_dir("after") / f"{repeat}.log")
        change = percent(float(after), float(before))
        result.append({
            "repeat": repeat,
            "scope": "whole child: all 12 rows, fixture generation, setup, oracles and drop",
            "before_peak_rss_kib": before, "after_peak_rss_kib": after,
            "change_percent": change, "adverse_over_5_percent": change > 5.0,
        })
    return result


def verify_capture_order(receipts: dict[tuple[str, str], dict[str, Any]]) -> None:
    starts = []
    for stage, repeat in ABBA:
        value = receipts[stage, repeat].get("started_utc")
        require(isinstance(value, str) and value, f"{stage}/{repeat} has no start time")
        try:
            starts.append(_datetime.datetime.fromisoformat(value.replace("Z", "+00:00")))
        except ValueError as error:
            fail(f"{stage}/{repeat} has an invalid start time: {error}")
    require(starts == sorted(starts) and len(set(starts)) == len(starts),
            "native captures are not serial ABBA")


def verify_negative_vector(capture: dict[str, Any]) -> dict[str, bool]:
    row = capture["rows"][("xlsx_one_percent_commit_save", "dense-wide")]["row"]
    short = copy.deepcopy(row)
    short["elapsed_ns"]["samples"].pop()
    short_rejected = False
    try:
        _validate_elapsed(short, 100, "negative short vector")
    except Exception:
        short_rejected = True
    require(short_rejected, "negative short elapsed vector was accepted")
    corrupt = copy.deepcopy(row)
    corrupt["elapsed_ns"]["samples"][0] += 1
    corrupt_rejected = False
    try:
        _validate_elapsed(corrupt, 100, "negative corrupt vector")
    except Exception:
        corrupt_rejected = True
    require(corrupt_rejected, "negative corrupt elapsed vector was accepted")
    return {"short_vector_rejected": short_rejected, "corrupt_vector_rejected": corrupt_rejected}


def verify() -> dict[str, Any]:
    custody = verify_sources()
    builds = {stage: verify_build(stage) for stage in ("before", "after")}
    allocator_build = verify_build("after", allocator=True)
    normal: dict[tuple[str, str], dict[str, Any]] = {}
    receipts: dict[tuple[str, str], dict[str, Any]] = {}
    for stage in ("before", "after"):
        preflight_receipt, preflight = verify_capture(stage, "preflight", builds[stage], 1, 0)
        normal[stage, "preflight"] = preflight
        receipts[stage, "preflight"] = preflight_receipt
    for stage, repeat in ABBA:
        receipt, capture = verify_capture(stage, repeat, builds[stage], 100, 3)
        normal[stage, repeat] = capture
        receipts[stage, repeat] = receipt
    capture_identity(normal, "native")
    verify_capture_order(receipts)

    allocator: dict[str, dict[str, Any]] = {}
    allocator_receipts: dict[str, dict[str, Any]] = {}
    for repeat in ("r1", "r2"):
        receipt, capture = verify_capture("after", repeat, allocator_build, 10, 1, allocator=True)
        allocator[repeat] = capture
        allocator_receipts[repeat] = receipt
        for key, row in capture["rows"].items():
            native = normal["after", "r1"]["rows"][key]["row"]
            require(canonical(row["row"].get("corpus"), "allocator corpus") ==
                    canonical(native.get("corpus"), "allocator corpus"),
                    f"allocator/{repeat}/{key} corpus differs from candidate native")
            require(row["row"].get("sink") == native.get("sink"),
                    f"allocator/{repeat}/{key} sink differs from candidate native")
    profile_receipt, profile = verify_capture("after", "profile", builds["after"], 3, 0, profile=True)
    del profile_receipt
    profile_key = ("xlsx_one_percent_commit_save", "dense-wide")
    verify_row_identity(
        profile["rows"][profile_key]["row"],
        normal["after", "r1"]["rows"][profile_key]["row"],
        "profile/after/dense-wide/xlsx_one_percent_commit_save",
    )
    pairs, drift, _ = native_comparisons({key: normal[key] for key in normal if key[1] in {"r1", "r2"}})
    allocation_rows = {
        repeat: {
            f"{case}/{shape}": item["allocation"]
            for (case, shape), item in capture["rows"].items()
        }
        for repeat, capture in allocator.items()
    }
    negative = verify_negative_vector(normal["after", "r1"])
    return {
        "schema": "litchi-0513-verification-v1",
        "comparison_scope": "before/after native commit and commit+save evidence; allocator timing/RSS and profile IR are excluded from native deltas",
        "custody": custody,
        "builds": {"before": builds["before"], "after": builds["after"], "after_allocator": allocator_build},
        "native": {
            "rows_per_capture": 12, "cases": list(CASES), "shapes": list(SHAPES),
            "preflight_samples": 1, "preflight_warmups": 0,
            "samples_per_capture": 100, "warmups_per_capture": 3,
            "total_elapsed_samples": 4 * 12 * 100,
            "abba_order": [f"{stage}/{repeat}" for stage, repeat in ABBA],
            "pairs": pairs, "same_role_drift": drift,
            "whole_child_rss": native_rss(),
            "adverse_flags_are_observations": True,
            "corpus_sink_equivalence_verified": True,
        },
        "allocator": {
            "candidate_only": True, "rows_per_capture": 12, "samples_per_row": 10,
            "warmups_per_row": 1, "total_samples": 2 * 12 * 10,
            "instrumented_elapsed_and_rss_excluded": True,
            "normal_allocation_status": "unavailable_or_absent; no zero inferred",
            "rows": allocation_rows,
            "corpus_sink_equivalence_verified": True,
            "incremental_peak_scope": "region_peak_live_bytes minus live_bytes_before; incremental operation demand, not process/document peak",
        },
        "profile": profile["annotations"],
        "negative_vector": negative,
        "report_catalog_bindings_verified": True,
        "artifact_and_identity_bindings_verified": True,
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
