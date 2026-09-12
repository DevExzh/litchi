#!/usr/bin/env python3
"""Validate and compare the supplemental eager XLSX guard reports.

``eager_guard.py`` is the frozen capture authority.  Its normal report
validator was written for the source-backed primary lane, whose reports carry
phase vectors under ``result.source``.  The eager consumer deliberately emits
only the total operation and sink evidence, so this companion validator keeps
the frozen receipt, binding, and artifact checks while validating the eager
report envelope on its own.  In particular, it requires the source field to
be absent and makes no source or phase observation.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import re
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
GUARD_PATH = HERE / "eager_guard.py"
PLAN_PATH = HERE / "plan.json"
EAGER_PLAN_PATH = HERE / "eager-guard-plan.json"
ANALYZER_PATH = Path(__file__).resolve()


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory eager guard artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def load_module(path: Path, name: str) -> Any:
    require(path.is_file(), f"missing helper {path}")
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


# Importing the frozen wrapper gives this analyzer the same plan, job, receipt,
# binding, and artifact custody rules used by capture.  It does not invoke the
# wrapper's capture path.
GUARD = load_module(GUARD_PATH, "xlsx_0525_frozen_eager_guard")
BASE = GUARD.BASE


def sha256(path: Path) -> str:
    return GUARD.sha256(path)


def read_json(path: Path) -> Any:
    return GUARD.read_json(path)


def relative(path: Path) -> str:
    return GUARD.relative(path)


def digest(value: Any, label: str) -> None:
    require(isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value) is not None,
            f"{label} is not a lowercase SHA-256 digest")


def nonnegative_integer(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")


def positive_integer(value: Any, label: str) -> None:
    nonnegative_integer(value, label)
    require(value > 0, f"{label} is not positive")


def string(value: Any, label: str) -> None:
    require(isinstance(value, str) and value, f"{label} is not a nonempty string")


def plans() -> tuple[dict[str, Any], dict[str, Any]]:
    # Use the frozen wrapper's plan checks, including the supplemental plan's
    # binding to the primary plan.  Keeping this call in one place prevents a
    # companion analyzer from silently drifting from capture configuration.
    return GUARD.plans()


def validate_not_applicable(value: Any, label: str) -> None:
    """Validate a descriptive unavailable operation-metric section.

    The eager report has no source, process, publication, or phase counters.
    Their explicit ``not_applicable`` records are retained as schema evidence;
    this routine rejects a hidden numeric observation in one of those records.
    """

    require(isinstance(value, dict), f"{label} is not an object")
    # A grouping object such as cfb_phases.open has no status of its own;
    # each leaf metric beneath it carries the explicit unavailable status.
    if "status" in value:
        require(value.get("status") == "not_applicable",
                f"{label}.status is not not_applicable")
    for key, child in value.items():
        if key == "status":
            continue
        if key in ("scope", "counter_scope"):
            string(child, f"{label}.{key}")
            continue
        require(isinstance(child, dict),
                f"{label}.{key} contains an unexpected scalar observation")
        validate_not_applicable(child, f"{label}.{key}")


def vector(value: Any, count: int, label: str) -> list[Any]:
    require(isinstance(value, list), f"{label} is not a vector")
    require(len(value) == count,
            f"{label} has {len(value)} values, expected {count}")
    return value


def constant_vector(value: Any, count: int, label: str) -> int:
    values = vector(value, count, label)
    for index, item in enumerate(values):
        nonnegative_integer(item, f"{label}[{index}]")
    require(all(item == values[0] for item in values),
            f"{label} is not constant across samples")
    return values[0]


def validate_sink(result_sink: Any, metrics_sink: Any, count: int, label: str) -> dict[str, Any]:
    require(isinstance(result_sink, dict), f"{label}.sink is not an object")
    require(set(result_sink) == {
        "accepted_bytes", "largest_write", "write_calls", "write_size_buckets",
    }, f"{label}.sink fields are unexpected")
    for key in ("accepted_bytes", "largest_write", "write_calls"):
        nonnegative_integer(result_sink[key], f"{label}.sink.{key}")
    require(result_sink["accepted_bytes"] > 0, f"{label}.sink.accepted_bytes is zero")
    require(result_sink["write_calls"] > 0, f"{label}.sink.write_calls is zero")
    require(0 < result_sink["largest_write"] <= 65536,
            f"{label}.sink.largest_write exceeds the bounded writer")
    expected_buckets = {
        "bytes_0", "bytes_1_to_512", "bytes_513_to_4096",
        "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536",
    }
    buckets = result_sink["write_size_buckets"]
    require(isinstance(buckets, dict) and set(buckets) == expected_buckets,
            f"{label}.sink.write_size_buckets fields are unexpected")
    for key, value in buckets.items():
        nonnegative_integer(value, f"{label}.sink.write_size_buckets.{key}")
    require(sum(buckets.values()) == result_sink["write_calls"],
            f"{label}.sink buckets do not reconcile with write calls")

    require(isinstance(metrics_sink, dict),
            f"{label}.operation_metrics.sink is not an object")
    require(set(metrics_sink) == {
        "status", "output_bytes", "write_status", "accepted_bytes", "write_calls",
        "largest_write", "write_size_buckets",
    }, f"{label}.operation_metrics.sink fields are unexpected")
    require(metrics_sink.get("status") == "not_applicable",
            f"{label}.operation_metrics.sink.status is not not_applicable")
    require(metrics_sink.get("write_status") == "measured",
            f"{label}.operation_metrics.sink.write_status is not measured")
    expected_scopes = {
        "accepted_bytes": "logical_sink_accepted_write_bytes",
        "largest_write": "logical_sink_largest_accepted_write",
        "write_calls": "logical_sink_accepted_write_calls",
    }
    for key, scope in expected_scopes.items():
        metric = metrics_sink.get(key)
        require(isinstance(metric, dict), f"{label}.operation_metrics.sink.{key} is not an object")
        require(metric.get("status") == "measured",
                f"{label}.operation_metrics.sink.{key}.status is not measured")
        require(metric.get("scope") == scope,
                f"{label}.operation_metrics.sink.{key}.scope differs")
        observed = constant_vector(metric.get("values"), count,
                                   f"{label}.operation_metrics.sink.{key}.values")
        require(observed == result_sink[key],
                f"{label}.operation_metrics.sink.{key} disagrees with sink")
    metric_buckets = metrics_sink.get("write_size_buckets")
    require(isinstance(metric_buckets, dict)
            and set(metric_buckets) == expected_buckets | {"status"},
            f"{label}.operation_metrics.sink.write_size_buckets fields are unexpected")
    require(metric_buckets.get("status") == "measured",
            f"{label}.operation_metrics.sink.write_size_buckets.status is not measured")
    for key in expected_buckets:
        metric = metric_buckets[key]
        require(isinstance(metric, dict),
                f"{label}.operation_metrics.sink.write_size_buckets.{key} is not an object")
        require(metric.get("status") == "measured"
                and metric.get("scope") == "logical_sink_accepted_write_size_bucket_counts",
                f"{label}.operation_metrics.sink.write_size_buckets.{key} metadata differs")
        observed = constant_vector(metric.get("values"), count,
                                   f"{label}.operation_metrics.sink.write_size_buckets.{key}.values")
        require(observed == buckets[key],
                f"{label}.operation_metrics.sink.write_size_buckets.{key} disagrees with sink")
    output_bytes = metrics_sink.get("output_bytes")
    validate_not_applicable(output_bytes, f"{label}.operation_metrics.sink.output_bytes")
    return {
        "accepted_bytes": result_sink["accepted_bytes"],
        "largest_write": result_sink["largest_write"],
        "write_calls": result_sink["write_calls"],
        "write_size_buckets": dict(sorted(buckets.items())),
    }


def validate_corpus(corpus: Any, shape: str, label: str) -> dict[str, Any]:
    require(isinstance(corpus, dict), f"{label}.corpus is not an object")
    required_strings = {
        "name", "generator", "package_format", "shape", "payload_kind",
        "compression", "target_entry", "archive_sha256", "target_payload_sha256",
    }
    require(required_strings <= set(corpus), f"{label}.corpus is incomplete")
    for key in ("name", "generator", "payload_kind", "compression", "target_entry"):
        string(corpus[key], f"{label}.corpus.{key}")
    require(corpus["shape"] == shape, f"{label}.corpus.shape differs")
    require(corpus["package_format"] == "XLSX/OPC/ZIP",
            f"{label}.corpus.package_format differs")
    digest(corpus["archive_sha256"], f"{label}.corpus.archive_sha256")
    digest(corpus["target_payload_sha256"],
           f"{label}.corpus.target_payload_sha256")
    for key in ("entry_count", "archive_member_count", "entry_bytes",
                "uncompressed_payload_bytes", "archive_bytes", "target_payload_bytes"):
        positive_integer(corpus.get(key), f"{label}.corpus.{key}")
    xlsx = corpus.get("xlsx")
    require(isinstance(xlsx, dict), f"{label}.corpus.xlsx is not an object")
    for key in ("sheet_count", "rows_per_sheet", "columns_per_sheet",
                "one_percent_update_count"):
        positive_integer(xlsx.get(key), f"{label}.corpus.xlsx.{key}")
    members = xlsx.get("source_members")
    require(isinstance(members, dict), f"{label}.corpus.xlsx.source_members is not an object")
    require(members.get("workbook") == "xl/workbook.xml",
            f"{label}.corpus.xlsx workbook member differs")
    worksheets = members.get("worksheets")
    require(isinstance(worksheets, list) and len(worksheets) == xlsx["sheet_count"],
            f"{label}.corpus.xlsx worksheet members differ")
    require(all(isinstance(item, str) and item.startswith("xl/worksheets/sheet")
                and item.endswith(".xml") for item in worksheets),
            f"{label}.corpus.xlsx worksheet member names differ")
    require(members.get("shared_strings") is None
            and members.get("styles") == "xl/styles.xml",
            f"{label}.corpus.xlsx ancillary members differ")
    return corpus


def validate_operation_metrics(operation: Any, count: int, label: str) -> dict[str, Any]:
    require(isinstance(operation, dict), f"{label}.operation_metrics is not an object")
    expected_keys = {
        "sample_count", "sample_indices", "alignment", "latency_claim",
        "source", "process", "sink", "publication", "materialization", "cfb_phases",
    }
    require(set(operation) == expected_keys,
            f"{label}.operation_metrics fields are unexpected")
    require(operation["sample_count"] == count,
            f"{label}.operation_metrics.sample_count differs")
    require(operation["sample_indices"] == list(range(count)),
            f"{label}.operation_metrics.sample_indices differs")
    require(operation["alignment"] == "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{label}.operation_metrics.alignment differs")
    require(operation["latency_claim"] == "comparable_timed_operation",
            f"{label}.operation_metrics.latency_claim differs")
    for key in ("source", "process", "publication", "materialization", "cfb_phases"):
        validate_not_applicable(operation[key], f"{label}.operation_metrics.{key}")
    return {
        "sample_count": count,
        "alignment": operation["alignment"],
        "latency_claim": operation["latency_claim"],
        "source_field": "absent",
        "phase_fields": "not_applicable",
    }


def validate_parallel_metrics(parallel: Any, case: str, corpus_sha: str,
                              label: str) -> None:
    require(isinstance(parallel, dict), f"{label}.parallel_metrics is not an object")
    require(set(parallel) == {
        "schema_version", "scope", "claim", "configured_worker_budget",
        "observed_process_thread_count", "cases",
    }, f"{label}.parallel_metrics fields are unexpected")
    require(parallel.get("schema_version") == 1
            and parallel.get("scope") == "explicit_local_execution_only"
            and parallel.get("claim") == "descriptive",
            f"{label}.parallel_metrics envelope differs")
    worker_budget = parallel.get("configured_worker_budget")
    require(isinstance(worker_budget, dict)
            and worker_budget.get("status") == "measured"
            and worker_budget.get("value") == [1]
            and worker_budget.get("scope") == "configuration.execution_workers",
            f"{label}.parallel_metrics worker budget differs")
    process_count = parallel.get("observed_process_thread_count")
    require(isinstance(process_count, dict)
            and process_count.get("status") == "unavailable",
            f"{label}.parallel_metrics process thread count differs")
    string(process_count.get("scope"),
           f"{label}.parallel_metrics process thread count scope")
    string(process_count.get("reason"),
           f"{label}.parallel_metrics process thread count reason")
    cases = parallel.get("cases")
    require(isinstance(cases, list) and len(cases) == 1,
            f"{label}.parallel_metrics case count differs")
    item = cases[0]
    require(isinstance(item, dict), f"{label}.parallel_metrics case is not an object")
    require(item.get("case") == case and item.get("corpus_sha256") == corpus_sha,
            f"{label}.parallel_metrics case identity differs")
    digest(item.get("corpus_sha256"),
           f"{label}.parallel_metrics.corpus_sha256")
    for key in ("configured_worker_count", "observed_local_worker_count",
                "deterministic_task_count", "deterministic_chunk_count"):
        metric = item.get(key)
        require(isinstance(metric, dict)
                and metric.get("status") == "not_applicable",
                f"{label}.parallel_metrics.{key} differs")
        string(metric.get("scope"), f"{label}.parallel_metrics.{key}.scope")
        string(metric.get("reason"), f"{label}.parallel_metrics.{key}.reason")
    lock_wait = item.get("lock_wait_ns")
    require(isinstance(lock_wait, dict)
            and lock_wait.get("status") == "unavailable",
            f"{label}.parallel_metrics.lock_wait_ns differs")
    string(lock_wait.get("scope"), f"{label}.parallel_metrics.lock_wait_ns.scope")
    string(lock_wait.get("reason"), f"{label}.parallel_metrics.lock_wait_ns.reason")


def identity_digest(identity: dict[str, Any]) -> str:
    encoded = json.dumps(identity, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(encoded).hexdigest()


def validate_eager_report(raw: Any, primary: dict[str, Any], eager: dict[str, Any],
                          job: dict[str, Any], binary: dict[str, Any]) -> dict[str, Any]:
    label = job["name"]
    require(isinstance(raw, dict), f"{label} report is not an object")
    require(set(raw) == {
        "schema_version", "tool", "binary_identity", "environment", "configuration",
        "parallel_metrics", "results",
    }, f"{label} report fields are unexpected")
    require(raw.get("schema_version") == 1, f"{label} schema version is not 1")
    tool = raw.get("tool")
    require(isinstance(tool, dict), f"{label}.tool is not an object")
    require(set(tool) == {
        "name", "version", "binary", "profile", "target_os", "target_arch",
        "instrumentation",
    }, f"{label}.tool fields are unexpected")
    require(tool.get("name") == "litchi-perf-baseline"
            and tool.get("version") == "0.1.0"
            and tool.get("binary") == "litchi-perf-baseline"
            and tool.get("profile") == "release"
            and tool.get("target_os") == "linux"
            and tool.get("target_arch") == "x86_64"
            and tool.get("instrumentation") == "none",
            f"{label}.tool identity differs")
    binary_identity = raw.get("binary_identity")
    require(isinstance(binary_identity, dict), f"{label}.binary_identity is not an object")
    require(set(binary_identity) == {
        "path", "binary_sha256", "binary_bytes", "mode_bits", "executable", "profile",
    }, f"{label}.binary_identity fields are unexpected")
    require(binary_identity.get("binary_sha256") == binary["sha256"],
            f"{label} report is bound to the wrong binary")
    digest(binary_identity.get("binary_sha256"), f"{label}.binary_identity.binary_sha256")
    require(binary_identity.get("profile") == "release",
            f"{label}.binary_identity.profile differs")
    require(binary_identity.get("path") == binary.get("path"),
            f"{label}.binary_identity.path differs")
    positive_integer(binary_identity.get("binary_bytes"),
                     f"{label}.binary_identity.binary_bytes")
    nonnegative_integer(binary_identity.get("mode_bits"),
                        f"{label}.binary_identity.mode_bits")
    require(binary_identity.get("executable") is True,
            f"{label}.binary_identity.executable is not true")
    environment = raw.get("environment")
    require(isinstance(environment, dict), f"{label}.environment is not an object")
    require(environment.get("git_revision") == primary["revision"],
            f"{label}.environment.git_revision differs")
    require(environment.get("cpu_affinity") == str(primary["cpu"]),
            f"{label}.environment.cpu_affinity differs")
    require(environment.get("os") == "linux", f"{label}.environment.os differs")
    configuration = raw.get("configuration")
    require(isinstance(configuration, dict), f"{label}.configuration is not an object")
    require(configuration.get("cases") == [job["case"]],
            f"{label}.configuration.cases differs")
    require(configuration.get("xlsx_cell_crud_shapes") == [job["shape"]],
            f"{label}.configuration.xlsx_cell_crud_shapes differs")
    require(configuration.get("samples_per_case") == job["samples"],
            f"{label}.configuration.samples_per_case differs")
    require(configuration.get("warmup_iterations_per_case") == job["warmup"],
            f"{label}.configuration.warmup_iterations_per_case differs")
    require(configuration.get("filesystem_fresh_child_per_sample") is True
            and configuration.get("filesystem_process_isolated") is True,
            f"{label}.configuration child isolation differs")
    require(configuration.get("execution_workers") == [1],
            f"{label}.configuration.execution_workers differs")
    results = raw.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label} must contain one result")
    result = results[0]
    require(isinstance(result, dict), f"{label} result is not an object")
    require(set(result) == {
        "case", "corpus", "elapsed_ns", "sink", "output_sha256", "operation_metrics",
    }, f"{label} result fields are unexpected")
    require(result.get("case") == job["case"], f"{label} result case differs")
    corpus = validate_corpus(result.get("corpus"), job["shape"], label)
    validate_parallel_metrics(raw["parallel_metrics"], job["case"],
                              corpus["archive_sha256"], label)
    # ``elapsed_ns.samples`` is the sorted statistics vector.  The separate
    # sample_order permutation maps those sorted positions back to acquisition
    # indices; it is already checked as a complete permutation by
    # BASE.verify_elapsed.  No phase vector exists here with which to perform
    # an additional alignment check.
    elapsed_stats, _elapsed_values = BASE.verify_elapsed(result, job["samples"], label)
    operation_summary = validate_operation_metrics(result.get("operation_metrics"),
                                                   job["samples"], label)
    sink = validate_sink(result.get("sink"), result["operation_metrics"]["sink"],
                         job["samples"], label)
    output_sha = result.get("output_sha256")
    digest(output_sha, f"{label}.output_sha256")
    identity = {
        "configuration": configuration,
        "corpus": corpus,
        "sink": sink,
        "output_sha256": output_sha,
    }
    return {
        "name": label,
        "kind": "primary",
        "guard": None,
        "repeat": job["repeat"],
        "case": job["case"],
        "shape": job["shape"],
        "samples": job["samples"],
        "timing": {"elapsed_ns": elapsed_stats},
        "rss": None,
        "identity": identity,
        "identity_sha256": identity_digest(identity),
        "schema": {
            "source_field": "absent",
            "operation_metrics": operation_summary,
            "phase_metrics": "not_applicable",
        },
    }


def validate_stage(stage: str, primary: dict[str, Any], eager: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / stage
    require(stage_dir.is_dir(), f"{stage} source stage is missing")
    manifest = stage_dir / "source-manifest.json"
    require(manifest.is_file() and not manifest.is_symlink(),
            f"{stage} source manifest is missing")
    manifest_sha = sha256(manifest)
    binary_identity = read_json(stage_dir / "binary-normal.json")
    require(isinstance(binary_identity, dict), f"{stage} normal binary identity is not an object")
    binary_sha = binary_identity.get("sha256")
    digest(binary_sha, f"{stage} binary-normal.sha256")
    binary = {"sha256": binary_sha, "path": binary_identity.get("path")}
    rows: list[dict[str, Any]] = []
    jobs = GUARD.eager_jobs(eager)
    expected_names = {job["name"] for job in jobs}
    actual_names = {
        path.name[:-len(".receipt.json")]
        for path in stage_dir.glob("eager-*.receipt.json")
    }
    require(actual_names == expected_names,
            f"{stage} eager receipt set differs from plan: "
            f"{sorted(actual_names ^ expected_names)}")
    for job in jobs:
        name = job["name"]
        # These are the frozen capture wrapper's receipt, artifact, binding,
        # binary, source-manifest, command, and RSS custody checks.
        GUARD.validate_guard_receipt(stage_dir, name, job, primary, eager,
                                     binary, manifest_sha)
        report_path = stage_dir / f"{name}.json"
        raw = read_json(report_path)
        row = validate_eager_report(raw, primary, eager, job, binary)
        row["rss"] = {"scope": "whole_child_process",
                       **BASE.validate_rss(stage_dir / f"{name}.rss.json")}
        row["report_sha256"] = sha256(report_path)
        rows.append(row)
    rows.sort(key=lambda row: (row["repeat"], row["shape"]))
    require(len(rows) == len(expected_names), f"{stage} eager row count differs")
    return {
        "stage": stage,
        "rows": rows,
        "manifest_sha256": manifest_sha,
        "binary_sha256": binary_sha,
        "validation": {
            "expected_matrix": True,
            "receipt_bindings": True,
            "report_and_rss_valid": True,
            "source_evidence": "absent_by_eager_schema",
            "phase_evidence": "not_applicable",
        },
    }


def change_percent(baseline: Any, candidate: Any, label: str) -> float:
    require(isinstance(baseline, (int, float)) and not isinstance(baseline, bool),
            f"{label}.baseline is not numeric")
    require(isinstance(candidate, (int, float)) and not isinstance(candidate, bool),
            f"{label}.candidate is not numeric")
    require(math.isfinite(float(baseline)) and math.isfinite(float(candidate)),
            f"{label} contains a non-finite value")
    require(float(baseline) > 0, f"{label}.baseline is not positive")
    return (float(candidate) / float(baseline) - 1.0) * 100.0


def metric_record(baseline: Any, candidate: Any, label: str) -> dict[str, Any]:
    return {
        "baseline": baseline,
        "candidate": candidate,
        "change_percent": change_percent(baseline, candidate, label),
    }


def compare(baseline: dict[str, Any], candidate: dict[str, Any], eager: dict[str, Any]) -> dict[str, Any]:
    left_rows = baseline["rows"]
    right_rows = candidate["rows"]
    left = {(row["repeat"], row["shape"]): row for row in left_rows}
    right = {(row["repeat"], row["shape"]): row for row in right_rows}
    require(len(left_rows) == len(left) and len(right_rows) == len(right),
            "eager baseline/candidate rows contain duplicate logical identities")
    require(set(left) == set(right), "eager baseline/candidate matrices differ")
    comparisons: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    for key in sorted(left):
        before, after = left[key], right[key]
        require(before["identity"] == after["identity"],
                f"eager logical identity differs for {key}")
        metrics: dict[str, Any] = {"elapsed_ns": {}}
        for stat in ("p50", "p95", "p99", "mean"):
            record = metric_record(before["timing"]["elapsed_ns"][stat],
                                   after["timing"]["elapsed_ns"][stat],
                                   f"eager {key} elapsed {stat}")
            record["adverse_over_five_percent"] = record["change_percent"] > 5.0
            metrics["elapsed_ns"][stat] = record
            if record["adverse_over_five_percent"]:
                adverse.append({"repeat": key[0], "shape": key[1],
                                "metric": "elapsed_ns", "stat": stat, **record})
        rss_record = metric_record(before["rss"]["max_rss_kib"],
                                   after["rss"]["max_rss_kib"],
                                   f"eager {key} peak RSS")
        rss_record["adverse_over_five_percent"] = rss_record["change_percent"] > 5.0
        metrics["max_rss_kib"] = rss_record
        if rss_record["adverse_over_five_percent"]:
            adverse.append({"repeat": key[0], "shape": key[1],
                            "metric": "max_rss_kib", "stat": "peak", **rss_record})
        comparisons.append({
            "repeat": key[0], "shape": key[1], "identity_equal": True,
            "metrics": metrics,
        })
    return {
        "rows": comparisons,
        "adverse_flags_over_five_percent": adverse,
        "gate": eager["gate"],
        "all_adverse_metrics_retained": True,
        "no_primary_gain_claim": True,
        "phase_metrics": {
            "status": "not_applicable",
            "reason": "eager reports contain no source or phase vectors",
        },
        "runtime_comparison_performed": True,
    }


def analyzer_identity() -> dict[str, str]:
    return {"path": relative(ANALYZER_PATH), "sha256": sha256(ANALYZER_PATH)}


def authority_identity() -> dict[str, str]:
    return {"path": relative(GUARD_PATH), "sha256": sha256(GUARD_PATH)}


def analyze(stage: str | None) -> dict[str, Any]:
    primary, eager = plans()
    common = {
        "plan_sha256": sha256(EAGER_PLAN_PATH),
        "primary_plan_sha256": sha256(PLAN_PATH),
        "capture_authority": authority_identity(),
        "analyzer": analyzer_identity(),
        "scope": eager["purpose"],
    }
    if stage in ("baseline", "candidate"):
        return {
            "status": "pass",
            "stage": stage,
            **common,
            "evidence": validate_stage(stage, primary, eager),
        }
    baseline = validate_stage("baseline", primary, eager)
    candidate = validate_stage("candidate", primary, eager)
    return {
        "status": "pass",
        "stage": "compare",
        **common,
        "baseline": baseline,
        "candidate": candidate,
        "comparison": compare(baseline, candidate, eager),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", choices=("baseline", "candidate", "compare"))
    parser.add_argument("action", choices=("analyze",), nargs="?", default="analyze")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    stage = None if args.stage == "compare" else args.stage
    output = args.output or HERE / ("eager-guard-comparison.json" if stage is None
                                    else f"{stage}/eager-guard-analysis.json")
    try:
        result = analyze(stage)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    # BASE.verify_elapsed contributes its legacy EvidenceError subclass for
    # the shared statistics checks; it is also a ValueError.  Normalize that
    # validation failure at this CLI boundary instead of leaking a traceback.
    except (EvidenceError, ValueError, OSError, json.JSONDecodeError, AssertionError) as error:
        print(f"analyze_eager_guard.py: error: {error}", file=sys.stderr)
        return 2
    print(f"0525 eager guard verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
