#!/usr/bin/env python3
"""Validate and compare the conditional 0542 eager XLSX read controls.

The ordinary eager controls are ``xlsx_open_owned`` and ``xlsx_first_cell``.
Each case is run separately for the medium and dense-wide XLSX corpora, with
two repeats per build in baseline/candidate/candidate/baseline order.  This
analyzer verifies the frozen plan, command receipts, source/binary custody,
report schema, corpus identity, elapsed vectors, and whole-child RSS before
making any comparison.

The lane is diagnostic.  It has no speedup threshold.  A positive change over
five percent is retained as an individual adverse review item; an absolute
same-build repeat change over five percent is retained as a drift review item.
The analyzer never starts a process or mutates a capture artifact.  ``--write``
only writes the deterministic JSON returned by ``analyze()``.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import re
import sys
from typing import Any

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))

import eager_run as EAGER  # noqa: E402  (sibling runner owns frozen job schema)


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
EAGER_PLAN_PATH = HERE / "eager-plan.json"
RUN_PATH = HERE / "run.py"
EAGER_RUN_PATH = HERE / "eager_run.py"
ANALYZER_PATH = Path(__file__).resolve()

STAGES = ("baseline", "candidate")
REPEATS = (1, 2)
TIMING_STATS = ("p50", "p95", "p99", "mean")
ADVERSE_THRESHOLD_PERCENT = 5.0
EXPECTED_TOP_LEVEL = {
    "schema_version", "tool", "binary_identity", "environment",
    "configuration", "parallel_metrics", "results",
}
EXPECTED_TOOL = {
    "name": "litchi-perf-baseline",
    "version": "0.1.0",
    "binary": "litchi-perf-baseline",
    "profile": "release",
    "target_os": "linux",
    "target_arch": "x86_64",
    "instrumentation": "none",
}
EXPECTED_ENVIRONMENT_FIELDS = {
    "rustc_version", "git_revision", "git_worktree_dirty",
    "logical_cpus_available", "allocator", "rustflags", "cargo_build_target",
    "perf_event_paranoid", "os", "kernel", "cpu_model", "total_memory_bytes",
    "page_size_bytes", "filesystem_type", "source_destination_same_device",
    "cpu_affinity", "storage_identifier",
}
EXPECTED_CONFIGURATION_FIELDS = {
    "samples_per_case", "warmup_iterations_per_case", "filesystem_cache_states",
    "filesystem_fresh_child_per_sample", "filesystem_process_isolated",
    "filesystem_root_selected", "cases", "corpus_shapes", "payload_kinds",
    "writer_shapes", "xlsx_shapes", "xlsb_shapes", "xlsx_cell_crud_shapes",
    "xlsx_row_visibility_shapes", "semantic_shapes", "rtf_variants",
    "range_simulation", "execution_workers", "opc_cache_lock_diagnostics",
}


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory eager evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def write_json(path: Path, value: Any) -> None:
    try:
        path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n",
                        encoding="utf-8")
    except OSError as error:
        raise EvidenceError(f"cannot write JSON {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside eager evidence: {path}") from error


def digest(value: Any, label: str, length: int = 64) -> str:
    require(isinstance(value, str)
            and re.fullmatch(r"[0-9a-f]{%d}" % length, value) is not None,
            f"{label} is not a lowercase SHA-{length * 4} digest")
    return value


def nonnegative_integer(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")
    return value


def positive_integer(value: Any, label: str) -> int:
    result = nonnegative_integer(value, label)
    require(result > 0, f"{label} is not positive")
    return result


def finite_number(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")
    return float(value)


def nonempty_string(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label} is not a nonempty string")
    return value


def same_number(actual: Any, expected: Any, label: str) -> None:
    finite_number(actual, label)
    finite_number(expected, f"expected {label}")
    require(math.isclose(float(actual), float(expected), rel_tol=1e-12, abs_tol=1e-9),
            f"{label}: {actual!r} != {expected!r}")


def student_t_critical_95(degrees: int) -> float:
    values = (
        12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262,
        2.228, 2.201, 2.179, 2.160, 2.145, 2.131, 2.120, 2.110, 2.101,
        2.093, 2.086, 2.080, 2.074, 2.069, 2.064, 2.060, 2.056, 2.052,
        2.048, 2.045, 2.042,
    )
    if degrees == 0:
        return 0.0
    if degrees <= len(values):
        return values[degrees - 1]
    z = 1.959_963_984_540_054
    z2 = z * z
    z3 = z2 * z
    z5 = z3 * z2
    z7 = z5 * z2
    d = float(degrees)
    return (z + (z3 + z) / (4.0 * d)
            + (5.0 * z5 + 16.0 * z3 + 3.0 * z) / (96.0 * d * d)
            + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * z)
            / (384.0 * d * d * d))


def expected_statistics(values: list[int]) -> dict[str, Any]:
    require(values, "empty elapsed vector")
    for index, value in enumerate(values):
        nonnegative_integer(value, f"elapsed sample {index}")
    ordered = sorted(values)
    mean = 0.0
    squared_deviation_sum = 0.0
    for index, value in enumerate(ordered):
        numeric = float(value)
        count = float(index + 1)
        delta = numeric - mean
        next_mean = mean + delta / count
        squared_deviation_sum += delta * (numeric - next_mean)
        mean = next_mean
    standard_deviation = (
        math.sqrt(squared_deviation_sum / (len(values) - 1))
        if len(values) > 1 else 0.0
    )
    margin = (
        student_t_critical_95(len(values) - 1) * standard_deviation
        / math.sqrt(len(values))
        if len(values) > 1 else 0.0
    )
    return {
        "unit": "ns",
        "samples": ordered,
        "p50": (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) // 2,
        "p95": ordered[min(math.ceil(len(ordered) * 0.95) - 1, len(ordered) - 1)],
        "p99": ordered[min(math.ceil(len(ordered) * 0.99) - 1, len(ordered) - 1)],
        "min": ordered[0],
        "max": ordered[-1],
        "mean": mean,
        "standard_deviation": standard_deviation,
        "confidence_interval_95": {
            "method": "two-sided Student's t interval for the mean",
            "lower": max(0.0, mean - margin),
            "upper": mean + margin,
        },
    }


def validate_statistics(value: Any, count: int, label: str) -> tuple[dict[str, Any], list[int]]:
    require(isinstance(value, dict), f"{label} is not an aggregate object")
    samples = value.get("samples")
    require(isinstance(samples, list) and len(samples) == count,
            f"{label}.samples has wrong cardinality")
    for index, item in enumerate(samples):
        nonnegative_integer(item, f"{label}.samples[{index}]")
    order = value.get("sample_order")
    require(isinstance(order, list) and sorted(order) == list(range(count)),
            f"{label}.sample_order is not a permutation")
    expected = expected_statistics(samples)
    # The Rust harness sorts the retained vector before writing its aggregate.
    require(value.get("unit") == expected["unit"], f"{label}.unit differs")
    require(value.get("samples") == expected["samples"],
            f"{label}.samples are not sorted")
    for key in ("min", "p50", "p95", "p99", "max", "mean", "standard_deviation"):
        same_number(value.get(key), expected[key], f"{label}.{key}")
    interval = value.get("confidence_interval_95")
    require(isinstance(interval, dict), f"{label}.confidence_interval_95 is not an object")
    require(interval.get("method") == expected["confidence_interval_95"]["method"],
            f"{label} confidence interval method differs")
    for key in ("lower", "upper"):
        same_number(interval.get(key), expected["confidence_interval_95"][key],
                    f"{label}.confidence_interval_95.{key}")
    return expected, list(samples)


def validate_rss(path: Path) -> dict[str, Any]:
    value = read_json(path)
    require(isinstance(value, dict), f"{relative(path)} RSS is not an object")
    require(set(value) == {"max_rss_kib", "elapsed_seconds", "user_seconds",
                           "system_seconds"},
            f"{relative(path)} RSS fields differ")
    positive_integer(value["max_rss_kib"], f"{relative(path)}.max_rss_kib")
    for key in ("elapsed_seconds", "user_seconds", "system_seconds"):
        require(finite_number(value[key], f"{relative(path)}.{key}") >= 0.0,
                f"{relative(path)}.{key} is negative")
    return value


def parse_time(value: Any, label: str) -> str:
    # ISO-8601 lexical ordering is valid for the UTC timestamps emitted by
    # run.py.  Keep the original strings in the report for exact replay.
    nonempty_string(value, label)
    require(value.endswith("+00:00") or value.endswith("Z"),
            f"{label} is not an explicit UTC timestamp")
    return value


def validate_plan() -> tuple[dict[str, Any], dict[str, Any]]:
    primary = EAGER.primary_plan()
    eager = read_json(EAGER_PLAN_PATH)
    EAGER.validate_plan(eager, primary)
    require(eager.get("primary_plan_sha256") == sha256(PLAN_PATH),
            "eager plan primary hash differs")
    require(eager.get("run_script_sha256") == sha256(RUN_PATH),
            "eager plan run hash differs")
    require(eager.get("eager_run_script_sha256") == sha256(EAGER_RUN_PATH),
            "eager plan eager runner hash differs")
    require(eager.get("analyzer_script_sha256") == sha256(ANALYZER_PATH),
            "eager plan analyzer hash differs")
    return primary, eager


def binary_identity(stage: str, eager: dict[str, Any]) -> dict[str, Any]:
    identity = EAGER.binary_metadata(stage)
    require(identity["sha256"] == eager["binary_sha256"][stage],
            f"{stage} binary hash differs from eager plan")
    require(identity["source_manifest_sha256"] == eager["source_manifest_sha256"][stage],
            f"{stage} binary source hash differs from eager plan")
    build_receipt_path = HERE / stage / "build-normal.receipt.json"
    require(build_receipt_path.is_file() and not build_receipt_path.is_symlink(),
            f"{stage} normal build receipt is missing")
    build_receipt = read_json(build_receipt_path)
    require(isinstance(build_receipt, dict), f"{stage} build receipt is not an object")
    require(build_receipt.get("exit_code") == 0
            and build_receipt.get("binary_sha256") is None,
            f"{stage} normal build receipt is not successful")
    require(build_receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{stage} build receipt plan hash differs")
    require(build_receipt.get("script_sha256") == sha256(RUN_PATH),
            f"{stage} build receipt run hash differs")
    require(build_receipt.get("source_manifest_sha256") ==
            eager["source_manifest_sha256"][stage],
            f"{stage} build receipt source hash differs")
    command = build_receipt.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{stage} build command is not a string list")
    require("cargo" in command and "build" in command and "--release" in command
            and "--locked" in command and "--manifest-path" in command
            and "--bin" in command,
            f"{stage} build command is incomplete")
    require(command[command.index("--bin") + 1] == "litchi-perf-baseline",
            f"{stage} build binary differs")
    require("--features" not in command,
            f"{stage} normal build unexpectedly enables a feature")
    artifacts = build_receipt.get("artifacts")
    require(isinstance(artifacts, dict)
            and set(artifacts) == {"build-normal.stdout", "build-normal.stderr"},
            f"{stage} build artifact inventory differs")
    for name, artifact_digest in artifacts.items():
        artifact = HERE / stage / name
        require(artifact.is_file() and not artifact.is_symlink()
                and sha256(artifact) == artifact_digest,
                f"{stage} build artifact digest differs: {name}")
    return identity


def validate_configuration(value: Any, job: dict[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}.configuration is not an object")
    require(set(value) == EXPECTED_CONFIGURATION_FIELDS,
            f"{label}.configuration fields differ")
    require(value["samples_per_case"] == job["samples"]
            and value["warmup_iterations_per_case"] == job["warmup"],
            f"{label}.configuration iteration counts differ")
    require(value["cases"] == [job["case"]],
            f"{label}.configuration cases differ")
    require(value["xlsx_shapes"] == [job["shape"]],
            f"{label}.configuration XLSX shapes differ")
    require(value["filesystem_cache_states"] == ["warm", "cold-requested"],
            f"{label}.configuration cache states differ")
    require(value["filesystem_fresh_child_per_sample"] is True
            and value["filesystem_process_isolated"] is True
            and value["filesystem_root_selected"] is False,
            f"{label}.configuration process isolation differs")
    require(value["corpus_shapes"] == ["tiny", "many-small", "few-large", "wide-root"],
            f"{label}.configuration corpus shapes differ")
    require(value["payload_kinds"] == ["compressible", "incompressible"],
            f"{label}.configuration payload kinds differ")
    require(value["writer_shapes"] == ["tiny", "large", "payload-heavy"],
            f"{label}.configuration writer shapes differ")
    require(value["xlsb_shapes"] == ["tiny", "medium", "large", "sparse"],
            f"{label}.configuration XLSB shapes differ")
    require(value["xlsx_cell_crud_shapes"] == ["medium", "dense-sparse"],
            f"{label}.configuration cell CRUD defaults differ")
    require(value["xlsx_row_visibility_shapes"] == ["medium", "large"],
            f"{label}.configuration row visibility defaults differ")
    require(value["semantic_shapes"] == ["tiny", "medium", "large"]
            and value["rtf_variants"] == ["plain"],
            f"{label}.configuration semantic defaults differ")
    require(value["range_simulation"] == {
        "fixed_latency_us": 100,
        "request_overhead_us": 25,
        "bandwidth_bytes_per_second": 52_428_800,
        "max_physical_range_bytes": 4_096,
    }, f"{label}.configuration range defaults differ")
    require(value["execution_workers"] == [1]
            and value["opc_cache_lock_diagnostics"] is False,
            f"{label}.configuration worker defaults differ")
    return value


def validate_parallel_metrics(value: Any, job: dict[str, Any], corpus_sha: str,
                              label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}.parallel_metrics is not an object")
    expected = {"schema_version", "scope", "claim", "configured_worker_budget",
                "observed_process_thread_count", "cases"}
    require(set(value) == expected, f"{label}.parallel_metrics fields differ")
    require(value.get("schema_version") == 1
            and value.get("scope") == "explicit_local_execution_only"
            and value.get("claim") == "descriptive",
            f"{label}.parallel_metrics envelope differs")
    worker = value.get("configured_worker_budget")
    require(worker == {"status": "measured", "value": [1],
                       "scope": "configuration.execution_workers"},
            f"{label}.parallel_metrics worker budget differs")
    process = value.get("observed_process_thread_count")
    require(isinstance(process, dict)
            and process.get("status") == "unavailable",
            f"{label}.parallel_metrics process count differs")
    nonempty_string(process.get("scope"), f"{label}.parallel_metrics process scope")
    nonempty_string(process.get("reason"), f"{label}.parallel_metrics process reason")
    cases = value.get("cases")
    require(isinstance(cases, list) and len(cases) == 1,
            f"{label}.parallel_metrics cases differ")
    case = cases[0]
    require(isinstance(case, dict)
            and case.get("case") == job["case"]
            and case.get("corpus_sha256") == corpus_sha,
            f"{label}.parallel_metrics case identity differs")
    digest(case.get("corpus_sha256"), f"{label}.parallel_metrics corpus SHA")
    for key in ("configured_worker_count", "observed_local_worker_count",
                "deterministic_task_count", "deterministic_chunk_count"):
        metric = case.get(key)
        require(isinstance(metric, dict) and metric.get("status") == "not_applicable",
                f"{label}.parallel_metrics.{key} differs")
        nonempty_string(metric.get("scope"), f"{label}.parallel_metrics.{key}.scope")
        nonempty_string(metric.get("reason"), f"{label}.parallel_metrics.{key}.reason")
    lock_wait = case.get("lock_wait_ns")
    require(isinstance(lock_wait, dict) and lock_wait.get("status") == "unavailable",
            f"{label}.parallel_metrics lock wait differs")
    nonempty_string(lock_wait.get("scope"), f"{label}.parallel_metrics lock scope")
    nonempty_string(lock_wait.get("reason"), f"{label}.parallel_metrics lock reason")
    return value


def corpus_dimensions(shape: str) -> tuple[int, int, int, int]:
    if shape == "medium":
        return 4, 32, 32, 41
    if shape == "dense-wide":
        return 2, 256, 256, 1_311
    raise EvidenceError(f"unsupported eager corpus shape {shape}")


def validate_corpus(value: Any, shape: str, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}.corpus is not an object")
    required = {
        "name", "generator", "package_format", "shape", "payload_kind",
        "compression", "entry_count", "archive_member_count", "entry_bytes",
        "uncompressed_payload_bytes", "archive_bytes", "archive_sha256",
        "target_entry", "target_payload_bytes", "target_payload_sha256", "xlsx",
    }
    require(set(value) == required, f"{label}.corpus fields differ")
    require(value["name"] == f"xlsx-{shape}", f"{label}.corpus.name differs")
    require(value["generator"] == "litchi-xlsx-synthetic-v1",
            f"{label}.corpus.generator differs")
    require(value["package_format"] == "XLSX/OPC/ZIP"
            and value["shape"] == shape
            and value["payload_kind"] == "deterministic-integer-grid"
            and value["compression"] == "deflate",
            f"{label}.corpus format identity differs")
    for key in ("entry_count", "archive_member_count", "entry_bytes",
                "uncompressed_payload_bytes", "archive_bytes", "target_payload_bytes"):
        positive_integer(value[key], f"{label}.corpus.{key}")
    digest(value["archive_sha256"], f"{label}.corpus.archive_sha256")
    digest(value["target_payload_sha256"], f"{label}.corpus.target_payload_sha256")
    require(value["target_entry"] == "Sheet1!A1"
            and value["target_payload_bytes"] == 1
            and value["target_payload_sha256"] ==
            "5feceb66ffc86f38d952786c6d696c79c2dbc239dd4e91b46729d73a27fb57e9",
            f"{label}.corpus target identity differs")
    sheet_count, rows, columns, updates = corpus_dimensions(shape)
    xlsx = value["xlsx"]
    require(isinstance(xlsx, dict)
            and set(xlsx) == {"sheet_count", "rows_per_sheet", "columns_per_sheet",
                               "one_percent_update_count", "source_members"},
            f"{label}.corpus.xlsx fields differ")
    require(xlsx["sheet_count"] == sheet_count
            and xlsx["rows_per_sheet"] == rows
            and xlsx["columns_per_sheet"] == columns
            and xlsx["one_percent_update_count"] == updates,
            f"{label}.corpus.xlsx dimensions differ")
    source_members = xlsx["source_members"]
    require(isinstance(source_members, dict)
            and set(source_members) == {"workbook", "worksheets", "shared_strings", "styles"},
            f"{label}.corpus source members differ")
    require(source_members["workbook"] == "xl/workbook.xml"
            and source_members["shared_strings"] is None
            and source_members["styles"] == "xl/styles.xml",
            f"{label}.corpus ancillary members differ")
    worksheets = source_members["worksheets"]
    expected_worksheets = [f"xl/worksheets/sheet{i}.xml"
                           for i in range(1, sheet_count + 1)]
    require(worksheets == expected_worksheets,
            f"{label}.corpus worksheet members differ")
    return value


def validate_report(raw: Any, primary: dict[str, Any], eager: dict[str, Any],
                    job: dict[str, Any], binary: dict[str, Any]) -> dict[str, Any]:
    label = job["name"]
    require(isinstance(raw, dict), f"{label} report is not an object")
    require(set(raw) == EXPECTED_TOP_LEVEL, f"{label} top-level report fields differ")
    require(raw.get("schema_version") == 1, f"{label} schema version differs")
    require(raw.get("tool") == EXPECTED_TOOL, f"{label}.tool identity differs")
    binary_raw = raw.get("binary_identity")
    require(isinstance(binary_raw, dict)
            and set(binary_raw) == {"path", "binary_sha256", "binary_bytes",
                                    "mode_bits", "executable", "profile"},
            f"{label}.binary_identity fields differ")
    require(binary_raw["path"] == binary["path"]
            and binary_raw["binary_sha256"] == binary["sha256"]
            and binary_raw["binary_bytes"] == binary["bytes"]
            and binary_raw["executable"] is True
            and binary_raw["profile"] == "release",
            f"{label}.binary_identity differs from retained binary")
    nonnegative_integer(binary_raw["mode_bits"], f"{label}.binary_identity.mode_bits")
    environment = raw.get("environment")
    require(isinstance(environment, dict)
            and set(environment) == EXPECTED_ENVIRONMENT_FIELDS,
            f"{label}.environment fields differ")
    require(environment["git_revision"] == primary["revision"]
            and environment["cpu_affinity"] == str(eager["cpu"])
            and environment["os"] == "linux",
            f"{label}.environment identity differs")
    configuration = validate_configuration(raw.get("configuration"), job, label)
    results = raw.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label} result count differs")
    result = results[0]
    require(isinstance(result, dict)
            and set(result) == {"case", "corpus", "elapsed_ns", "sink"},
            f"{label} result fields differ")
    require(result["case"] == job["case"], f"{label} result case differs")
    corpus = validate_corpus(result["corpus"], job["shape"], label)
    # Recheck the corpus binding now that the exact manifest is available.
    parallel = validate_parallel_metrics(raw["parallel_metrics"], job,
                                         corpus["archive_sha256"], label)
    elapsed_stats, elapsed_values = validate_statistics(
        result["elapsed_ns"], job["samples"], f"{label}.elapsed_ns"
    )
    require(result["sink"] is None, f"{label}.sink unexpectedly contains output evidence")
    # These ordinary read cases intentionally have no source-backed or
    # publication identity.  The exact result-key check above rejects a
    # hidden source/output/phase field; retain an explicit contract in output.
    logical_identity = {
        "case": result["case"],
        "corpus": corpus,
        "configuration": configuration,
        "parallel_metrics": parallel,
        "source": "absent",
        "output_sha256": "absent",
        "sink": None,
        "operation_metrics": "absent",
    }
    return {
        "name": label,
        "stage": job["stage"],
        "repeat": job["repeat"],
        "group": job["group"],
        "case": job["case"],
        "shape": job["shape"],
        "samples": job["samples"],
        "timing": {"elapsed_ns": elapsed_stats},
        "elapsed_values": elapsed_values,
        "rss": None,
        "report_identity": {
            "schema_version": raw["schema_version"],
            "tool": raw["tool"],
            "binary_identity": binary_raw,
            "environment": environment,
            "configuration": configuration,
            "parallel_metrics": parallel,
        },
        "logical_identity": logical_identity,
        "identity_sha256": hashlib.sha256(
            json.dumps(logical_identity, sort_keys=True, separators=(",", ":")).encode()
        ).hexdigest(),
        "raw_identity": {
            "source_field": "absent",
            "output_field": "absent",
            "sink_field": "explicit_null",
            "operation_metrics_field": "absent",
        },
    }


def expected_artifacts(name: str) -> set[str]:
    return {f"{name}.json", f"{name}.stdout", f"{name}.stderr",
            f"{name}.rss.json"}


def validate_receipt(stage: str, job: dict[str, Any], eager: dict[str, Any],
                     binary: dict[str, Any], source_manifest_sha: str,
                     working_manifest_sha: str) -> tuple[dict[str, Any], dict[str, Any]]:
    stage_dir = HERE / stage
    name = job["name"]
    receipt_path = stage_dir / f"{name}.receipt.json"
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict), f"{stage}/{name} receipt is not an object")
    expected_fields = {
        "command", "start_utc", "end_utc", "seconds", "exit_code",
        "binary_sha256", "source_manifest_sha256",
        "working_source_manifest_sha256", "script_sha256", "plan_sha256",
        "environment", "artifacts",
    }
    require(set(receipt) == expected_fields, f"{stage}/{name} receipt fields differ")
    parse_time(receipt["start_utc"], f"{stage}/{name}.start_utc")
    parse_time(receipt["end_utc"], f"{stage}/{name}.end_utc")
    require(receipt["start_utc"] <= receipt["end_utc"],
            f"{stage}/{name} receipt interval is reversed")
    require(finite_number(receipt["seconds"], f"{stage}/{name}.seconds") >= 0,
            f"{stage}/{name}.seconds is negative")
    require(receipt["exit_code"] == 0, f"{stage}/{name} child failed")
    require(receipt["binary_sha256"] == binary["sha256"]
            and receipt["source_manifest_sha256"] == source_manifest_sha
            and receipt["working_source_manifest_sha256"] == working_manifest_sha
            and receipt["script_sha256"] == sha256(RUN_PATH)
            and receipt["plan_sha256"] == sha256(PLAN_PATH),
            f"{stage}/{name} receipt identity differs")
    environment = receipt["environment"]
    require(isinstance(environment, dict)
            and set(environment) == {"RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS",
                                     "LD_PRELOAD", "MALLOC_CONF", "GLIBC_TUNABLES",
                                     "TMPDIR"},
            f"{stage}/{name} receipt environment differs")
    expected_command = EAGER.expected_capture_command(Path(binary["path"]), job, eager)
    require(receipt["command"] == expected_command,
            f"{stage}/{name} receipt command differs")
    artifacts = receipt["artifacts"]
    require(isinstance(artifacts, dict)
            and set(artifacts) == expected_artifacts(name),
            f"{stage}/{name} receipt artifact inventory differs")
    for artifact_name, artifact_digest in artifacts.items():
        artifact = stage_dir / artifact_name
        require(artifact.is_file() and not artifact.is_symlink()
                and sha256(artifact) == artifact_digest,
                f"{stage}/{name} artifact digest differs: {artifact_name}")
    binding_path = stage_dir / f"{name}.binding.json"
    binding = read_json(binding_path)
    require(isinstance(binding, dict), f"{stage}/{name} binding is not an object")
    expected_binding_fields = {
        "schema", "stage", "repeat", "group", "name", "case", "shape",
        "source_checkout", "retained_baseline", "primary_plan_sha256",
        "eager_plan_sha256", "run_script_sha256", "eager_run_script_sha256",
        "analyzer_script_sha256", "source_manifest_sha256",
        "working_source_manifest_sha256", "binary_sha256", "receipt_sha256",
        "command_sha256", "output", "rss",
    }
    require(set(binding) == expected_binding_fields,
            f"{stage}/{name} binding fields differ")
    require(binding["schema"] == EAGER.BINDING_SCHEMA
            and binding["stage"] == stage
            and binding["repeat"] == job["repeat"]
            and binding["group"] == job["group"]
            and binding["name"] == name
            and binding["case"] == job["case"]
            and binding["shape"] == job["shape"]
            and binding["source_checkout"] == job["source_checkout"]
            and binding["retained_baseline"] == job["retained_baseline"],
            f"{stage}/{name} binding job identity differs")
    require(binding["primary_plan_sha256"] == sha256(PLAN_PATH)
            and binding["eager_plan_sha256"] == sha256(EAGER_PLAN_PATH)
            and binding["run_script_sha256"] == sha256(RUN_PATH)
            and binding["eager_run_script_sha256"] == sha256(EAGER_RUN_PATH)
            and binding["analyzer_script_sha256"] == sha256(ANALYZER_PATH)
            and binding["source_manifest_sha256"] == source_manifest_sha
            and binding["working_source_manifest_sha256"] == working_manifest_sha
            and binding["binary_sha256"] == binary["sha256"]
            and binding["receipt_sha256"] == sha256(receipt_path),
            f"{stage}/{name} binding digest differs")
    require(binding["command_sha256"] == hashlib.sha256(
        json.dumps(expected_command, separators=(",", ":")).encode()
    ).hexdigest(), f"{stage}/{name} binding command hash differs")
    require(binding["output"] == relative(stage_dir / f"{name}.json")
            and binding["rss"] == relative(stage_dir / f"{name}.rss.json"),
            f"{stage}/{name} binding artifact paths differ")
    raw = read_json(stage_dir / f"{name}.json")
    row = validate_report(raw, *validate_plan(), job, binary)
    row["rss"] = {"scope": "whole_child_process",
                   **validate_rss(stage_dir / f"{name}.rss.json")}
    row["report_sha256"] = sha256(stage_dir / f"{name}.json")
    row["receipt_sha256"] = sha256(receipt_path)
    row["receipt_start_utc"] = receipt["start_utc"]
    row["receipt_end_utc"] = receipt["end_utc"]
    return receipt, row


def validate_stage(stage: str, eager: dict[str, Any], primary: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / stage
    require(stage_dir.is_dir(), f"{stage} stage directory is missing")
    manifest_path = stage_dir / "source-manifest.json"
    require(manifest_path.is_file() and not manifest_path.is_symlink(),
            f"{stage} source manifest is missing")
    _, manifest_sha = EAGER.stage_manifest(stage)
    require(manifest_sha == eager["source_manifest_sha256"][stage],
            f"{stage} source manifest differs from eager plan")
    binary = binary_identity(stage, eager)
    expected_jobs = [job for job in eager["jobs"] if job["stage"] == stage]
    expected_names = {job["name"] for job in expected_jobs}
    actual_names = {path.name[:-len(".receipt.json")]
                    for path in stage_dir.glob("eager-*.receipt.json")}
    require(actual_names == expected_names,
            f"{stage} eager receipt set differs: {sorted(actual_names ^ expected_names)}")
    rows: list[dict[str, Any]] = []
    receipts: list[dict[str, Any]] = []
    for job in expected_jobs:
        _, working_manifest_sha = EAGER.stage_manifest("candidate")
        receipt, row = validate_receipt(stage, job, eager, binary, manifest_sha,
                                        working_manifest_sha)
        rows.append(row)
        receipts.append(receipt)
    ordered = sorted(receipts, key=lambda item: item["start_utc"])
    require(all(left["end_utc"] <= right["start_utc"]
                for left, right in zip(ordered, ordered[1:])),
            f"{stage} eager child receipts overlap")
    # The same build must reproduce each ordinary read's logical identity
    # across repeats.  Timing and process RSS remain intentionally free to vary.
    for case in EAGER.CASES:
        for shape in EAGER.SHAPES:
            matching = [row for row in rows if row["case"] == case and row["shape"] == shape]
            require(len(matching) == 2, f"{stage} {case}/{shape} repeat count differs")
            require(matching[0]["logical_identity"] == matching[1]["logical_identity"],
                    f"{stage} {case}/{shape} logical identity drifts")
    rows.sort(key=lambda row: (row["repeat"], EAGER.CASES.index(row["case"]),
                               EAGER.SHAPES.index(row["shape"])))
    return {
        "stage": stage,
        "manifest_sha256": manifest_sha,
        "binary_identity": binary,
        "rows": rows,
        "custody": {
            "receipt_count": len(receipts),
            "receipts_non_overlapping": True,
            "source_manifest_entries": len(read_json(manifest_path)),
        },
    }


def capture_order(baseline: dict[str, Any], candidate: dict[str, Any],
                  eager: dict[str, Any]) -> list[dict[str, Any]]:
    groups: dict[str, list[dict[str, Any]]] = {}
    for evidence in (baseline, candidate):
        for row in evidence["rows"]:
            groups.setdefault(row["group"], []).append(row)
    result: list[dict[str, Any]] = []
    previous_end: str | None = None
    for group in EAGER.NATIVE_ORDER:
        rows = groups.get(group)
        require(isinstance(rows, list) and len(rows) == len(EAGER.CASES) * len(EAGER.SHAPES),
                f"eager capture group {group} is incomplete")
        ordered = sorted(rows, key=lambda row: row["receipt_start_utc"])
        expected = [job["name"] for job in eager["jobs"] if job["group"] == group]
        require([row["name"] for row in ordered] == expected,
                f"eager capture group {group} job order differs")
        start, end = ordered[0]["receipt_start_utc"], ordered[-1]["receipt_end_utc"]
        if previous_end is not None:
            require(previous_end <= start, f"eager ABBA groups overlap at {group}")
        previous_end = end
        result.append({
            "group": group,
            "stage": ordered[0]["stage"],
            "repeat": ordered[0]["repeat"],
            "start_utc": start,
            "end_utc": end,
            "source_checkout": eager["stage_source_protocol"][group],
        })
    return result


def metric_record(baseline: Any, candidate: Any, label: str) -> dict[str, Any]:
    base = finite_number(baseline, f"{label}.baseline")
    after = finite_number(candidate, f"{label}.candidate")
    require(base > 0.0, f"{label}.baseline is not positive")
    change = (after / base - 1.0) * 100.0
    return {
        "baseline": baseline,
        "candidate": candidate,
        "change_percent": change,
        "adverse_over_five_percent": change > ADVERSE_THRESHOLD_PERCENT,
    }


def review_item(kind: str, lane: str, key: dict[str, Any], metric: str, stat: str,
                record: dict[str, Any]) -> dict[str, Any]:
    return {
        "review_id": f"{kind}:{lane}:{key['repeat']}:{key['case']}:{key['shape']}:{metric}:{stat}",
        "kind": kind,
        "lane": lane,
        "repeat": key["repeat"],
        "case": key["case"],
        "shape": key["shape"],
        "metric": metric,
        "stat": stat,
        "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
        "change_percent": record["change_percent"],
        "baseline": record["baseline"],
        "candidate": record["candidate"],
        "status": "individual-review-required",
        "reason": (
            "Metric increased by more than the frozen 5% review threshold; "
            "the eager lane has no automatic speedup admission gate."
        ),
    }


def compare_pair(before: dict[str, Any], after: dict[str, Any], lane: str,
                key: dict[str, Any], reviews: list[dict[str, Any]],
                adverse: list[dict[str, Any]]) -> dict[str, Any]:
    elapsed: dict[str, Any] = {}
    for stat in TIMING_STATS:
        record = metric_record(before["timing"]["elapsed_ns"][stat],
                               after["timing"]["elapsed_ns"][stat],
                               f"{lane} {key} elapsed {stat}")
        elapsed[stat] = record
        if record["adverse_over_five_percent"]:
            item = review_item("adverse_metric", lane, key, "elapsed_ns", stat, record)
            adverse.append(item)
            reviews.append(item)
    rss = metric_record(before["rss"]["max_rss_kib"], after["rss"]["max_rss_kib"],
                        f"{lane} {key} max RSS")
    if rss["adverse_over_five_percent"]:
        item = review_item("adverse_metric", lane, key, "max_rss_kib", "peak", rss)
        adverse.append(item)
        reviews.append(item)
    return {"elapsed_ns": elapsed, "max_rss_kib": {"peak": rss}}


def compare(baseline: dict[str, Any], candidate: dict[str, Any],
            eager: dict[str, Any]) -> dict[str, Any]:
    left = {(row["repeat"], row["case"], row["shape"]): row
            for row in baseline["rows"]}
    right = {(row["repeat"], row["case"], row["shape"]): row
             for row in candidate["rows"]}
    require(set(left) == set(right), "eager baseline/candidate matrices differ")
    comparisons: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    reviews: list[dict[str, Any]] = []
    for repeat, case, shape in sorted(left):
        before, after = left[(repeat, case, shape)], right[(repeat, case, shape)]
        require(before["logical_identity"] == after["logical_identity"],
                f"eager logical identity differs for {repeat}/{case}/{shape}")
        key = {"repeat": repeat, "case": case, "shape": shape}
        comparisons.append({
            **key,
            "identity_equal": True,
            "metrics": compare_pair(before, after, "candidate-vs-baseline", key,
                                     reviews, adverse),
        })

    drift: list[dict[str, Any]] = []
    drift_flags: list[dict[str, Any]] = []
    for stage, evidence in (("baseline", baseline), ("candidate", candidate)):
        rows = {(row["repeat"], row["case"], row["shape"]): row
                for row in evidence["rows"]}
        for case in EAGER.CASES:
            for shape in EAGER.SHAPES:
                before, after = rows[(1, case, shape)], rows[(2, case, shape)]
                key = {"repeat": "1-to-2", "case": case, "shape": shape}
                metrics: dict[str, Any] = {"elapsed_ns": {}}
                for stat in TIMING_STATS:
                    record = metric_record(before["timing"]["elapsed_ns"][stat],
                                           after["timing"]["elapsed_ns"][stat],
                                           f"{stage} repeat drift {case}/{shape} {stat}")
                    metrics["elapsed_ns"][stat] = record
                    if abs(record["change_percent"]) > ADVERSE_THRESHOLD_PERCENT:
                        item = review_item("same_build_drift", stage, key,
                                           "elapsed_ns", stat, record)
                        item["repeat_first"] = 1
                        item["repeat_second"] = 2
                        drift_flags.append(item)
                        reviews.append(item)
                rss = metric_record(before["rss"]["max_rss_kib"],
                                    after["rss"]["max_rss_kib"],
                                    f"{stage} repeat drift {case}/{shape} max RSS")
                metrics["max_rss_kib"] = {"peak": rss}
                if abs(rss["change_percent"]) > ADVERSE_THRESHOLD_PERCENT:
                    item = review_item("same_build_drift", stage, key,
                                       "max_rss_kib", "peak", rss)
                    item["repeat_first"] = 1
                    item["repeat_second"] = 2
                    drift_flags.append(item)
                    reviews.append(item)
                drift.append({
                    "stage": stage,
                    "case": case,
                    "shape": shape,
                    "repeat_first": 1,
                    "repeat_second": 2,
                    "metrics": metrics,
                })
    return {
        "timing_comparisons": comparisons,
        "same_build_drift": drift,
        "adverse_flags_over_five_percent": adverse,
        "same_build_drift_over_five_percent": drift_flags,
        "individual_reviews": reviews,
        "individual_review_count": len(reviews),
        "all_metrics_compared": ["elapsed_ns.p50", "elapsed_ns.p95",
                                 "elapsed_ns.p99", "elapsed_ns.mean",
                                 "max_rss_kib.peak"],
        "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
        "no_speedup_requirement": True,
        "runtime_comparison_performed": True,
    }


def analyze(output: Path | None = None) -> dict[str, Any]:
    primary, eager = validate_plan()
    baseline = validate_stage("baseline", eager, primary)
    candidate = validate_stage("candidate", eager, primary)
    result = {
        "status": "pass",
        "schema": "litchi-0542-eager-analysis-v1",
        "primary_plan_sha256": sha256(PLAN_PATH),
        "eager_plan_sha256": sha256(EAGER_PLAN_PATH),
        "capture_authority": {"path": relative(RUN_PATH), "sha256": sha256(RUN_PATH)},
        "eager_runner_authority": {"path": relative(EAGER_RUN_PATH),
                                    "sha256": sha256(EAGER_RUN_PATH)},
        "analysis_authority": {"path": relative(ANALYZER_PATH),
                                "sha256": sha256(ANALYZER_PATH)},
        "baseline": baseline,
        "candidate": candidate,
        "capture_order": capture_order(baseline, candidate, eager),
        "comparison": compare(baseline, candidate, eager),
        "scope": eager["purpose"],
        "no_speedup_requirement": True,
        "output": relative(output or (HERE / "eager-analysis.json")),
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write", action="store_true",
                        help="write eager-analysis.json after read-only analysis")
    parser.add_argument("--output", type=Path,
                        help="optional output path used with --write")
    args = parser.parse_args()
    try:
        require(args.output is None or args.write,
                "--output requires --write")
        output = args.output or (HERE / "eager-analysis.json")
        value = analyze(output)
        if args.write:
            output.parent.mkdir(parents=True, exist_ok=True)
            write_json(output, value)
            print(f"0542 eager analysis verified: {output}")
        else:
            print(json.dumps(value, indent=2, sort_keys=True))
    except (EvidenceError, AssertionError, OSError, json.JSONDecodeError) as error:
        print(f"analyze_eager.py: error: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
