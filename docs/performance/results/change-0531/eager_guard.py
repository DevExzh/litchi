#!/usr/bin/env python3
"""Capture and verify the conditional 0531 eager XLSX guard.

The guard is deliberately separate from the native 0531 driver.  It reuses
``run.py`` for source, binary, child, temporary-directory, RSS, and receipt
custody, while keeping its own supplemental plan hash in every binding.  The
four captures are ordered baseline-r1, candidate-r1, candidate-r2,
baseline-r2.  Both baseline captures run the retained baseline binary while
the candidate source checkout is installed, which makes the binary/source
relationship explicit in each receipt.

The eager report has no source-backed phase vector.  A report is accepted only
when it omits ``result.source`` (the serde representation of the eager path)
and recursively marks the
source/process/publication/materialization/CFB operation-metric sections as
``not_applicable``.  Sink vectors, phase absence, corpus identity, and output
digest are checked directly; the source-backed validator is never coerced to
accept an eager report.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import re
import sys
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
EAGER_PLAN_PATH = HERE / "eager-plan.json"
RUN_PATH = HERE / "run.py"
ANALYZER_PATH = HERE / "analyze.py"


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory eager-guard artifact."""


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


# Importing the frozen modules is read-only.  RUN owns capture receipts and
# NUMERIC.BASE supplies the already-bound elapsed/RSS validators.
RUN = load_module(RUN_PATH, "xlsx_0531_frozen_run_for_eager")
NUMERIC = load_module(ANALYZER_PATH, "xlsx_0531_frozen_numeric_for_eager")
BASE = NUMERIC.BASE


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
        raise EvidenceError(f"cannot write {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence directory: {path}") from error


def digest(value: Any, label: str) -> None:
    require(isinstance(value, str)
            and re.fullmatch(r"[0-9a-f]{64}", value) is not None,
            f"{label} is not a lowercase SHA-256 digest")


def nonnegative_integer(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")


def positive_integer(value: Any, label: str) -> None:
    nonnegative_integer(value, label)
    require(value > 0, f"{label} is not positive")


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def nonempty_string(value: Any, label: str) -> None:
    require(isinstance(value, str) and value, f"{label} is not a nonempty string")


def plans() -> tuple[dict[str, Any], dict[str, Any]]:
    primary = read_json(PLAN_PATH)
    eager = read_json(EAGER_PLAN_PATH)
    require(isinstance(primary, dict) and isinstance(eager, dict),
            "plan documents are not objects")
    require(eager.get("schema") == "litchi-0531-eager-guard-plan-v1",
            "eager plan schema differs")
    require(eager.get("primary_plan_sha256") == sha256(PLAN_PATH),
            "eager plan is bound to a different primary plan")
    require(eager.get("run_script_sha256") == sha256(RUN_PATH),
            "eager plan is bound to a different capture driver")
    require(eager.get("guard_script_sha256") == sha256(Path(__file__)),
            "eager plan is bound to a different eager guard")
    require(eager.get("case") == "xlsx_eager_cell_values_one_percent_edit_save",
            "eager case differs")
    require(eager.get("shapes") == ["medium", "dense-sparse"],
            "eager shape matrix differs")
    require(eager.get("repeats") == 2 and eager.get("warmup") == 10
            and eager.get("samples") == 30,
            "eager repeat/warmup/sample counts differ")
    require(eager.get("cpu") == primary.get("cpu") == 2,
            "eager CPU differs from the frozen primary plan")
    require(eager.get("native_order") == [
        "baseline-r1", "candidate-r1", "candidate-r2", "baseline-r2",
    ], "eager capture order differs")
    protocol = eager.get("stage_source_protocol")
    require(isinstance(protocol, dict)
            and set(protocol) == set(eager["native_order"]),
            "eager source checkout protocol is incomplete")
    require(protocol["baseline-r1"] ==
            "retained baseline binary under candidate source checkout",
            "baseline-r1 source protocol differs")
    require(protocol["candidate-r1"] ==
            "candidate binary under candidate source checkout",
            "candidate-r1 source protocol differs")
    require(protocol["candidate-r2"] ==
            "candidate binary under candidate source checkout",
            "candidate-r2 source protocol differs")
    require(protocol["baseline-r2"] ==
            "retained baseline binary under candidate source checkout",
            "baseline-r2 source protocol differs")
    require(eager.get("no_speedup_requirement") is True,
            "eager guard unexpectedly declares a speedup gate")
    threshold = eager.get("adverse_threshold_percent")
    finite_number(threshold, "eager adverse threshold")
    require(float(threshold) == 5.0, "eager adverse threshold differs")
    primary_case = primary.get("primary")
    require(isinstance(primary_case, dict)
            and primary_case.get("case") ==
            "xlsx_source_backed_cell_values_one_percent_edit_save",
            "primary plan case differs")
    conditional = primary.get("conditional_lanes")
    require(isinstance(conditional, dict)
            and isinstance(conditional.get("eager"), str),
            "primary eager conditional lane is missing")
    for stage, field in (("baseline", "baseline"), ("candidate", "candidate")):
        manifest = HERE / stage / "source-manifest.json"
        identity = read_json(HERE / stage / "binary-normal.json")
        require(isinstance(identity, dict), f"{stage} binary identity is not an object")
        require(identity.get("source_manifest_sha256") ==
                sha256(manifest), f"{stage} binary source manifest differs")
        require(identity.get("sha256") == eager["binary_sha256"][field],
                f"{stage} binary differs from frozen eager plan")
        require(sha256(manifest) == eager["source_manifest_sha256"][field],
                f"{stage} source manifest differs from frozen eager plan")
        digest(identity.get("sha256"), f"{stage} binary SHA")
        digest(sha256(manifest), f"{stage} source manifest SHA")
    return primary, eager


def group_key(stage: str, repeat: int) -> str:
    return f"{stage}-r{repeat}"


def group_spec(stage: str, repeat: int, eager: dict[str, Any]) -> dict[str, Any]:
    require(stage in ("baseline", "candidate"), f"unsupported eager stage {stage}")
    require(repeat in (1, 2), f"unsupported eager repeat {repeat}")
    group = group_key(stage, repeat)
    require(group in eager["native_order"], f"{group} is not in frozen ABBA order")
    if stage == "baseline":
        binary_stage = "baseline"
        retained_baseline = True
        working_manifest_stage = "candidate"
    else:
        binary_stage = "candidate"
        retained_baseline = False
        working_manifest_stage = "candidate"
    return {
        "stage": stage,
        "repeat": repeat,
        "group": group,
        "binary_stage": binary_stage,
        "retained_baseline": retained_baseline,
        "source_manifest_stage": stage,
        "working_manifest_stage": working_manifest_stage,
    }


def jobs(stage: str, repeat: int, eager: dict[str, Any]) -> list[dict[str, Any]]:
    spec = group_spec(stage, repeat, eager)
    return [
        {
            "name": f"eager-r{repeat}-{shape}",
            "stage": stage,
            "repeat": repeat,
            "shape": shape,
            "case": eager["case"],
            "warmup": eager["warmup"],
            "samples": eager["samples"],
            "group": spec["group"],
        }
        for shape in eager["shapes"]
    ]


def binary_identity(stage: str, eager: dict[str, Any]) -> tuple[dict[str, Any], Path]:
    path = HERE / stage / "binary-normal.json"
    identity = read_json(path)
    require(isinstance(identity, dict), f"{stage} binary identity is not an object")
    binary_path_value = identity.get("path")
    require(isinstance(binary_path_value, str) and binary_path_value,
            f"{stage} binary path is missing")
    binary_path = Path(binary_path_value)
    require(binary_path.is_file() and not binary_path.is_symlink(),
            f"{stage} retained normal binary is not a regular file")
    binary_sha = identity.get("sha256")
    digest(binary_sha, f"{stage} binary identity SHA")
    require(binary_sha == eager["binary_sha256"][stage],
            f"{stage} binary identity differs from eager plan")
    require(sha256(binary_path) == binary_sha,
            f"{stage} retained normal binary digest differs")
    positive_integer(identity.get("bytes"), f"{stage} binary bytes")
    require(identity["bytes"] == binary_path.stat().st_size,
            f"{stage} binary byte count differs")
    return {"sha256": binary_sha, "path": binary_path_value,
            "bytes": identity["bytes"]}, binary_path


def expected_capture_command(binary: Path, job: dict[str, Any],
                             eager: dict[str, Any]) -> list[str]:
    stage_dir = HERE / job["stage"]
    name = job["name"]
    rss = stage_dir / f"{name}.rss.json"
    output = stage_dir / f"{name}.json"
    return [
        "taskset", "-c", str(eager["cpu"]), "/usr/bin/time", "-f",
        '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
        "-o", str(rss), str(binary), "--warmup", str(job["warmup"]),
        "--samples", str(job["samples"]), "--case", job["case"],
        "--xlsx-cell-crud-shape", job["shape"], "--json", str(output),
    ]


def assert_capture_slot_is_next(stage: str, repeat: int, eager: dict[str, Any]) -> None:
    target = group_key(stage, repeat)
    order = eager["native_order"]
    index = order.index(target)
    for previous in order[:index]:
        previous_stage, previous_repeat = previous.rsplit("-r", 1)
        for job in jobs(previous_stage, int(previous_repeat), eager):
            require((HERE / previous_stage / f"{job['name']}.receipt.json").is_file(),
                    f"capture order prerequisite is incomplete: {previous}/{job['name']}")
    for later in order[index + 1:]:
        later_stage, later_repeat = later.rsplit("-r", 1)
        for job in jobs(later_stage, int(later_repeat), eager):
            require(not (HERE / later_stage / f"{job['name']}.receipt.json").exists(),
                    f"capture order already advanced past {target}")


def capture(stage: str, repeat: int) -> None:
    primary, eager = plans()
    assert_capture_slot_is_next(stage, repeat, eager)
    group = group_spec(stage, repeat, eager)
    identity, binary = binary_identity(group["binary_stage"], eager)
    stage_dir = HERE / stage
    source_manifest = stage_dir / "source-manifest.json"
    working_manifest = HERE / group["working_manifest_stage"] / "source-manifest.json"
    require(source_manifest.is_file(), f"{stage} source manifest is missing")
    require(working_manifest.is_file(), f"{group['working_manifest_stage']} source manifest is missing")
    source_manifest_sha = sha256(source_manifest)
    working_manifest_sha = sha256(working_manifest)
    for job in jobs(stage, repeat, eager):
        name = job["name"]
        for suffix in (".json", ".stdout", ".stderr", ".rss.json",
                       ".receipt.json", ".binding.json"):
            require(not (stage_dir / f"{name}{suffix}").exists(),
                    f"capture artifact already exists: {stage}/{name}{suffix}")
        command = expected_capture_command(binary, job, eager)
        RUN.run(stage, name, command, binary,
                retained_baseline=group["retained_baseline"])
        receipt_path = stage_dir / f"{name}.receipt.json"
        receipt = read_json(receipt_path)
        require(receipt.get("working_source_manifest_sha256") == working_manifest_sha,
                f"{name} working source manifest does not follow ABBA protocol")
        binding = {
            "schema": "litchi-0531-eager-guard-binding-v1",
            "stage": stage,
            "repeat": repeat,
            "group": group["group"],
            "name": name,
            "source_checkout": eager["stage_source_protocol"][group["group"]],
            "retained_baseline": group["retained_baseline"],
            "primary_plan_sha256": sha256(PLAN_PATH),
            "eager_plan_sha256": sha256(EAGER_PLAN_PATH),
            "run_script_sha256": sha256(RUN_PATH),
            "guard_script_sha256": sha256(Path(__file__)),
            "source_manifest_sha256": source_manifest_sha,
            "working_source_manifest_sha256": working_manifest_sha,
            "binary_sha256": identity["sha256"],
            "receipt_sha256": sha256(receipt_path),
            "command_sha256": hashlib.sha256(
                json.dumps(command, separators=(",", ":")).encode()
            ).hexdigest(),
            "output": relative(stage_dir / f"{name}.json"),
            "rss": relative(stage_dir / f"{name}.rss.json"),
        }
        write_json(stage_dir / f"{name}.binding.json", binding)
    print(f"0531 eager capture passed: {stage} repeat {repeat}")


def validate_not_applicable(value: Any, label: str) -> None:
    """Require an explicitly unavailable diagnostic subtree.

    Grouping objects such as ``cfb_phases.open`` may omit their own status;
    every leaf still has to carry ``status: not_applicable``.  Descriptive
    scope/counter-scope/reason strings are the only permitted scalar leaves.
    """

    require(isinstance(value, dict), f"{label} is not an object")
    if "status" in value:
        require(value.get("status") == "not_applicable",
                f"{label}.status is not not_applicable")
    for key, child in value.items():
        if key == "status":
            continue
        if key in ("scope", "counter_scope", "reason"):
            nonempty_string(child, f"{label}.{key}")
            continue
        require(isinstance(child, dict),
                f"{label}.{key} contains an unexpected scalar observation")
        validate_not_applicable(child, f"{label}.{key}")


def constant_vector(value: Any, count: int, label: str) -> int:
    require(isinstance(value, list) and len(value) == count,
            f"{label} does not have {count} values")
    for index, item in enumerate(value):
        nonnegative_integer(item, f"{label}[{index}]")
    require(all(item == value[0] for item in value),
            f"{label} is not constant across samples")
    return value[0]


def validate_sink(result_sink: Any, metrics_sink: Any, count: int,
                  label: str) -> dict[str, Any]:
    require(isinstance(result_sink, dict), f"{label}.sink is not an object")
    sink_fields = {"accepted_bytes", "largest_write", "write_calls",
                   "write_size_buckets"}
    require(set(result_sink) == sink_fields,
            f"{label}.sink fields are unexpected")
    for key in ("accepted_bytes", "largest_write", "write_calls"):
        nonnegative_integer(result_sink[key], f"{label}.sink.{key}")
    require(result_sink["accepted_bytes"] > 0,
            f"{label}.sink.accepted_bytes is zero")
    require(result_sink["write_calls"] > 0,
            f"{label}.sink.write_calls is zero")
    require(0 < result_sink["largest_write"] <= 65_536,
            f"{label}.sink.largest_write exceeds bounded writer")
    bucket_fields = {
        "bytes_0", "bytes_1_to_512", "bytes_513_to_4096",
        "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536",
    }
    buckets = result_sink["write_size_buckets"]
    require(isinstance(buckets, dict) and set(buckets) == bucket_fields,
            f"{label}.sink.write_size_buckets fields are unexpected")
    for key, value in buckets.items():
        nonnegative_integer(value, f"{label}.sink.write_size_buckets.{key}")
    require(sum(buckets.values()) == result_sink["write_calls"],
            f"{label}.sink buckets do not reconcile with write calls")

    require(isinstance(metrics_sink, dict),
            f"{label}.operation_metrics.sink is not an object")
    metrics_fields = {
        "status", "output_bytes", "write_status", "accepted_bytes",
        "write_calls", "largest_write", "write_size_buckets",
    }
    require(set(metrics_sink) == metrics_fields,
            f"{label}.operation_metrics.sink fields are unexpected")
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
        require(isinstance(metric, dict),
                f"{label}.operation_metrics.sink.{key} is not an object")
        require(metric.get("status") == "measured"
                and metric.get("scope") == scope,
                f"{label}.operation_metrics.sink.{key} metadata differs")
        observed = constant_vector(metric.get("values"), count,
                                   f"{label}.operation_metrics.sink.{key}.values")
        require(observed == result_sink[key],
                f"{label}.operation_metrics.sink.{key} disagrees with sink")
    metric_buckets = metrics_sink.get("write_size_buckets")
    require(isinstance(metric_buckets, dict)
            and set(metric_buckets) == bucket_fields | {"status"},
            f"{label}.operation_metrics.sink.write_size_buckets fields are unexpected")
    require(metric_buckets.get("status") == "measured",
            f"{label}.operation_metrics.sink.write_size_buckets.status is not measured")
    for key in bucket_fields:
        metric = metric_buckets[key]
        require(isinstance(metric, dict),
                f"{label}.operation_metrics.sink.write_size_buckets.{key} is not an object")
        require(metric.get("status") == "measured"
                and metric.get("scope") ==
                "logical_sink_accepted_write_size_bucket_counts",
                f"{label}.operation_metrics.sink.write_size_buckets.{key} metadata differs")
        observed = constant_vector(metric.get("values"), count,
                                   f"{label}.operation_metrics.sink.write_size_buckets.{key}.values")
        require(observed == buckets[key],
                f"{label}.operation_metrics.sink.write_size_buckets.{key} disagrees with sink")
    validate_not_applicable(metrics_sink["output_bytes"],
                            f"{label}.operation_metrics.sink.output_bytes")
    return {
        "accepted_bytes": result_sink["accepted_bytes"],
        "largest_write": result_sink["largest_write"],
        "write_calls": result_sink["write_calls"],
        "write_size_buckets": dict(sorted(buckets.items())),
    }


def validate_corpus(corpus: Any, shape: str, label: str) -> dict[str, Any]:
    require(isinstance(corpus, dict), f"{label}.corpus is not an object")
    required = {
        "name", "generator", "package_format", "shape", "payload_kind",
        "compression", "target_entry", "archive_sha256", "target_payload_sha256",
    }
    require(required <= set(corpus), f"{label}.corpus is incomplete")
    for key in ("name", "generator", "payload_kind", "compression", "target_entry"):
        nonempty_string(corpus[key], f"{label}.corpus.{key}")
    require(corpus["shape"] == shape, f"{label}.corpus.shape differs")
    require(corpus["package_format"] == "XLSX/OPC/ZIP",
            f"{label}.corpus.package_format differs")
    digest(corpus["archive_sha256"], f"{label}.corpus.archive_sha256")
    digest(corpus["target_payload_sha256"],
           f"{label}.corpus.target_payload_sha256")
    for key in ("entry_count", "archive_member_count", "entry_bytes",
                "uncompressed_payload_bytes", "archive_bytes",
                "target_payload_bytes"):
        positive_integer(corpus.get(key), f"{label}.corpus.{key}")
    xlsx = corpus.get("xlsx")
    require(isinstance(xlsx, dict), f"{label}.corpus.xlsx is not an object")
    for key in ("sheet_count", "rows_per_sheet", "columns_per_sheet",
                "one_percent_update_count"):
        positive_integer(xlsx.get(key), f"{label}.corpus.xlsx.{key}")
    members = xlsx.get("source_members")
    require(isinstance(members, dict),
            f"{label}.corpus.xlsx.source_members is not an object")
    require(members.get("workbook") == "xl/workbook.xml",
            f"{label}.corpus.xlsx workbook member differs")
    worksheets = members.get("worksheets")
    require(isinstance(worksheets, list)
            and len(worksheets) == xlsx["sheet_count"],
            f"{label}.corpus.xlsx worksheet members differ")
    require(all(isinstance(item, str)
                and item.startswith("xl/worksheets/sheet")
                and item.endswith(".xml") for item in worksheets),
            f"{label}.corpus.xlsx worksheet member names differ")
    require(members.get("shared_strings") is None
            and members.get("styles") == "xl/styles.xml",
            f"{label}.corpus.xlsx ancillary members differ")
    return corpus


def validate_parallel_metrics(parallel: Any, case: str, corpus_sha: str,
                              label: str) -> None:
    require(isinstance(parallel, dict),
            f"{label}.parallel_metrics is not an object")
    fields = {
        "schema_version", "scope", "claim", "configured_worker_budget",
        "observed_process_thread_count", "cases",
    }
    require(set(parallel) == fields,
            f"{label}.parallel_metrics fields are unexpected")
    require(parallel.get("schema_version") == 1
            and parallel.get("scope") == "explicit_local_execution_only"
            and parallel.get("claim") == "descriptive",
            f"{label}.parallel_metrics envelope differs")
    worker = parallel.get("configured_worker_budget")
    require(isinstance(worker, dict)
            and worker.get("status") == "measured"
            and worker.get("value") == [1]
            and worker.get("scope") == "configuration.execution_workers",
            f"{label}.parallel_metrics worker budget differs")
    process_count = parallel.get("observed_process_thread_count")
    require(isinstance(process_count, dict)
            and process_count.get("status") == "unavailable",
            f"{label}.parallel_metrics process thread count differs")
    nonempty_string(process_count.get("scope"),
                    f"{label}.parallel_metrics process thread count scope")
    nonempty_string(process_count.get("reason"),
                    f"{label}.parallel_metrics process thread count reason")
    cases = parallel.get("cases")
    require(isinstance(cases, list) and len(cases) == 1,
            f"{label}.parallel_metrics case count differs")
    item = cases[0]
    require(isinstance(item, dict), f"{label}.parallel_metrics case is not an object")
    require(item.get("case") == case and item.get("corpus_sha256") == corpus_sha,
            f"{label}.parallel_metrics case identity differs")
    digest(item.get("corpus_sha256"), f"{label}.parallel_metrics corpus SHA")
    for key in ("configured_worker_count", "observed_local_worker_count",
                "deterministic_task_count", "deterministic_chunk_count"):
        metric = item.get(key)
        require(isinstance(metric, dict)
                and metric.get("status") == "not_applicable",
                f"{label}.parallel_metrics.{key} differs")
        nonempty_string(metric.get("scope"),
                        f"{label}.parallel_metrics.{key}.scope")
        nonempty_string(metric.get("reason"),
                        f"{label}.parallel_metrics.{key}.reason")
    lock_wait = item.get("lock_wait_ns")
    require(isinstance(lock_wait, dict)
            and lock_wait.get("status") == "unavailable",
            f"{label}.parallel_metrics.lock_wait_ns differs")
    nonempty_string(lock_wait.get("scope"),
                    f"{label}.parallel_metrics.lock_wait_ns.scope")
    nonempty_string(lock_wait.get("reason"),
                    f"{label}.parallel_metrics.lock_wait_ns.reason")


def identity_digest(identity: dict[str, Any]) -> str:
    return hashlib.sha256(
        json.dumps(identity, sort_keys=True, separators=(",", ":")).encode()
    ).hexdigest()


def validate_operation_metrics(operation: Any, count: int, label: str) -> dict[str, Any]:
    require(isinstance(operation, dict),
            f"{label}.operation_metrics is not an object")
    expected = {
        "sample_count", "sample_indices", "alignment", "latency_claim",
        "source", "process", "sink", "publication", "materialization",
        "cfb_phases",
    }
    require(set(operation) == expected,
            f"{label}.operation_metrics fields are unexpected")
    require(operation.get("sample_count") == count,
            f"{label}.operation_metrics.sample_count differs")
    require(operation.get("sample_indices") == list(range(count)),
            f"{label}.operation_metrics.sample_indices differs")
    require(operation.get("alignment") ==
            "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{label}.operation_metrics.alignment differs")
    require(operation.get("latency_claim") == "comparable_timed_operation",
            f"{label}.operation_metrics.latency_claim differs")
    for key in ("source", "process", "publication", "materialization",
                "cfb_phases"):
        validate_not_applicable(operation[key],
                                f"{label}.operation_metrics.{key}")
    return {
        "sample_count": count,
        "alignment": operation["alignment"],
        "latency_claim": operation["latency_claim"],
        "source": "explicit_not_applicable",
        "phase_metrics": "not_applicable",
    }


def validate_eager_report(raw: Any, primary: dict[str, Any],
                          eager: dict[str, Any], job: dict[str, Any],
                          binary: dict[str, Any]) -> dict[str, Any]:
    label = job["name"]
    require(isinstance(raw, dict), f"{label} report is not an object")
    require(set(raw) == {
        "schema_version", "tool", "binary_identity", "environment",
        "configuration", "parallel_metrics", "results",
    }, f"{label} report fields are unexpected")
    require(raw.get("schema_version") == 1,
            f"{label}.schema_version differs")
    tool = raw.get("tool")
    require(isinstance(tool, dict), f"{label}.tool is not an object")
    require(set(tool) == {
        "name", "version", "binary", "profile", "target_os", "target_arch",
        "instrumentation",
    }, f"{label}.tool fields are unexpected")
    require(tool == {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "binary": "litchi-perf-baseline",
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
        "instrumentation": "none",
    }, f"{label}.tool identity differs")
    binary_raw = raw.get("binary_identity")
    require(isinstance(binary_raw, dict),
            f"{label}.binary_identity is not an object")
    require(set(binary_raw) == {
        "path", "binary_sha256", "binary_bytes", "mode_bits", "executable",
        "profile",
    }, f"{label}.binary_identity fields are unexpected")
    require(binary_raw.get("path") == binary["path"],
            f"{label}.binary_identity.path differs")
    require(binary_raw.get("binary_sha256") == binary["sha256"],
            f"{label}.binary_identity SHA differs")
    digest(binary_raw.get("binary_sha256"),
           f"{label}.binary_identity.binary_sha256")
    positive_integer(binary_raw.get("binary_bytes"),
                     f"{label}.binary_identity.binary_bytes")
    require(binary_raw["binary_bytes"] == binary["bytes"],
            f"{label}.binary_identity.binary_bytes differs")
    nonnegative_integer(binary_raw.get("mode_bits"),
                        f"{label}.binary_identity.mode_bits")
    require(binary_raw.get("executable") is True,
            f"{label}.binary_identity.executable is not true")
    require(binary_raw.get("profile") == "release",
            f"{label}.binary_identity.profile differs")
    environment = raw.get("environment")
    require(isinstance(environment, dict),
            f"{label}.environment is not an object")
    require(environment.get("git_revision") == primary["revision"],
            f"{label}.environment.git_revision differs")
    require(environment.get("cpu_affinity") == str(eager["cpu"]),
            f"{label}.environment.cpu_affinity differs")
    require(environment.get("os") == "linux", f"{label}.environment.os differs")
    configuration = raw.get("configuration")
    require(isinstance(configuration, dict),
            f"{label}.configuration is not an object")
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
        "case", "corpus", "elapsed_ns", "sink", "output_sha256",
        "operation_metrics",
    }, f"{label} result fields are unexpected")
    require(result.get("case") == job["case"],
            f"{label} result case differs")
    # This is the eager/source distinction.  Keep the field absent in the raw
    # identity rather than silently passing it to the source-backed analyzer,
    # which would manufacture phase evidence.
    require("source" not in result,
            f"{label}.result.source is present for eager case")
    corpus = validate_corpus(result.get("corpus"), job["shape"], label)
    validate_parallel_metrics(raw.get("parallel_metrics"), job["case"],
                              corpus["archive_sha256"], label)
    elapsed_stats, elapsed_values = BASE.verify_elapsed(
        result, job["samples"], label)
    operation = validate_operation_metrics(result.get("operation_metrics"),
                                           job["samples"], label)
    sink = validate_sink(result.get("sink"),
                         result["operation_metrics"]["sink"],
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
        "stage": job["stage"],
        "repeat": job["repeat"],
        "shape": job["shape"],
        "case": job["case"],
        "samples": job["samples"],
        "timing": {"elapsed_ns": elapsed_stats},
        "rss": None,
        "identity": identity,
        "identity_sha256": identity_digest(identity),
        "raw_identity": {
            "source_field": "absent_by_eager_schema",
            "sink": "summary_and_vectors_reconciled",
            "phase_fields": "explicit_not_applicable",
            "output_sha256": output_sha,
        },
        "operation_metrics": operation,
        "elapsed_values": elapsed_values,
    }


def validate_binding(stage_dir: Path, job: dict[str, Any],
                     eager: dict[str, Any], binary: dict[str, Any],
                     source_manifest_sha: str,
                     working_manifest_sha: str,
                     receipt_sha: str) -> dict[str, Any]:
    path = stage_dir / f"{job['name']}.binding.json"
    binding = read_json(path)
    require(isinstance(binding, dict), f"{relative(path)} is not an object")
    require(set(binding) == {
        "schema", "stage", "repeat", "group", "name", "source_checkout",
        "retained_baseline", "primary_plan_sha256", "eager_plan_sha256",
        "run_script_sha256", "guard_script_sha256", "source_manifest_sha256",
        "working_source_manifest_sha256", "binary_sha256", "receipt_sha256",
        "command_sha256", "output", "rss",
    }, f"{relative(path)} fields are unexpected")
    require(binding.get("schema") == "litchi-0531-eager-guard-binding-v1",
            f"{relative(path)} schema differs")
    require(binding.get("stage") == job["stage"]
            and binding.get("repeat") == job["repeat"]
            and binding.get("group") == job["group"]
            and binding.get("name") == job["name"],
            f"{relative(path)} job identity differs")
    group = group_spec(job["stage"], job["repeat"], eager)
    require(binding.get("source_checkout") == eager["stage_source_protocol"][job["group"]],
            f"{relative(path)} source checkout description differs")
    require(binding.get("retained_baseline") == group["retained_baseline"],
            f"{relative(path)} retained-baseline flag differs")
    require(binding.get("primary_plan_sha256") == sha256(PLAN_PATH),
            f"{relative(path)} primary plan binding differs")
    require(binding.get("eager_plan_sha256") == sha256(EAGER_PLAN_PATH),
            f"{relative(path)} eager plan binding differs")
    require(binding.get("run_script_sha256") == sha256(RUN_PATH),
            f"{relative(path)} run script binding differs")
    require(binding.get("guard_script_sha256") == sha256(Path(__file__)),
            f"{relative(path)} guard script binding differs")
    require(binding.get("source_manifest_sha256") == source_manifest_sha,
            f"{relative(path)} source manifest binding differs")
    require(binding.get("working_source_manifest_sha256") == working_manifest_sha,
            f"{relative(path)} working source binding differs")
    require(binding.get("binary_sha256") == binary["sha256"],
            f"{relative(path)} binary binding differs")
    require(binding.get("receipt_sha256") == receipt_sha,
            f"{relative(path)} receipt binding differs")
    require(binding.get("output") == relative(stage_dir / f"{job['name']}.json")
            and binding.get("rss") == relative(stage_dir / f"{job['name']}.rss.json"),
            f"{relative(path)} output paths differ")
    digest(binding.get("command_sha256"), f"{relative(path)} command SHA")
    return binding


def validate_receipt(stage_dir: Path, job: dict[str, Any], eager: dict[str, Any],
                     binary: dict[str, Any], source_manifest_sha: str,
                     working_manifest_sha: str) -> dict[str, Any]:
    name = job["name"]
    receipt_path = stage_dir / f"{name}.receipt.json"
    receipt = read_json(receipt_path)
    require(receipt.get("exit_code") == 0,
            f"{name} child did not exit successfully")
    require(receipt.get("binary_sha256") == binary["sha256"],
            f"{name} receipt binary differs")
    require(receipt.get("source_manifest_sha256") == source_manifest_sha,
            f"{name} receipt source manifest differs")
    require(receipt.get("working_source_manifest_sha256") == working_manifest_sha,
            f"{name} receipt working source differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{name} receipt primary plan differs")
    require(receipt.get("script_sha256") == sha256(RUN_PATH),
            f"{name} receipt run script differs")
    command = receipt.get("command")
    expected = expected_capture_command(Path(binary["path"]), job, eager)
    require(command == expected, f"{name} receipt command differs")
    artifacts = receipt.get("artifacts")
    expected_artifacts = {
        f"{name}.json", f"{name}.stdout", f"{name}.stderr",
        f"{name}.rss.json",
    }
    require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts,
            f"{name} receipt artifact inventory differs")
    for filename, artifact_sha in artifacts.items():
        artifact = stage_dir / filename
        require(artifact.is_file() and not artifact.is_symlink(),
                f"{name} artifact is not regular: {filename}")
        require(sha256(artifact) == artifact_sha,
                f"{name} artifact digest differs: {filename}")
    for filename in (f"{name}.json", f"{name}.rss.json"):
        require((stage_dir / filename).is_file(),
                f"{name} required output is missing: {filename}")
    validate_binding(stage_dir, job, eager, binary, source_manifest_sha,
                     working_manifest_sha, sha256(receipt_path))
    return receipt


def validate_stage(stage: str, eager: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / stage
    require(stage_dir.is_dir(), f"{stage} stage directory is missing")
    manifest = stage_dir / "source-manifest.json"
    require(manifest.is_file() and not manifest.is_symlink(),
            f"{stage} source manifest is missing")
    source_manifest_sha = sha256(manifest)
    group_manifests = {
        repeat: sha256(HERE / group_spec(stage, repeat, eager)["working_manifest_stage"]
                       / "source-manifest.json")
        for repeat in (1, 2)
    }
    binary, _binary_path = binary_identity(stage, eager)
    expected_jobs = [job for repeat in (1, 2) for job in jobs(stage, repeat, eager)]
    expected_names = {job["name"] for job in expected_jobs}
    actual_names = {
        path.name[:-len(".receipt.json")]
        for path in stage_dir.glob("eager-*.receipt.json")
    }
    require(actual_names == expected_names,
            f"{stage} eager receipt set differs from plan: "
            f"{sorted(actual_names ^ expected_names)}")
    rows: list[dict[str, Any]] = []
    receipts: list[dict[str, Any]] = []
    for job in expected_jobs:
        receipt = validate_receipt(stage_dir, job, eager, binary,
                                   source_manifest_sha,
                                   group_manifests[job["repeat"]])
        raw = read_json(stage_dir / f"{job['name']}.json")
        row = validate_eager_report(raw, read_json(PLAN_PATH), eager, job, binary)
        row["rss"] = {
            "scope": "whole_child_process",
            **BASE.validate_rss(stage_dir / f"{job['name']}.rss.json"),
        }
        row["report_sha256"] = sha256(stage_dir / f"{job['name']}.json")
        row["receipt_start_utc"] = receipt.get("start_utc")
        row["receipt_end_utc"] = receipt.get("end_utc")
        nonempty_string(row["receipt_start_utc"], f"{job['name']} receipt start")
        nonempty_string(row["receipt_end_utc"], f"{job['name']} receipt end")
        rows.append(row)
        receipts.append(receipt)
    rows.sort(key=lambda row: (row["repeat"], eager["shapes"].index(row["shape"])))
    for shape in eager["shapes"]:
        shape_rows = [row for row in rows if row["shape"] == shape]
        require(len(shape_rows) == 2, f"{stage} {shape} eager repeat count differs")
        require(shape_rows[0]["identity"] == shape_rows[1]["identity"],
                f"{stage} {shape} eager output identity drifts between repeats")
    ordered = sorted(receipts, key=lambda receipt: receipt["start_utc"])
    require(all(left["end_utc"] <= right["start_utc"]
                for left, right in zip(ordered, ordered[1:])),
            f"{stage} eager child receipts overlap")
    return {
        "stage": stage,
        "rows": rows,
        "manifest_sha256": source_manifest_sha,
        "binary_identity": binary,
        "custody": {
            "receipt_count": len(receipts),
            "receipts_non_overlapping": True,
            "source_manifest_entries": len(read_json(manifest)),
        },
        "validation": {
            "expected_matrix": True,
            "receipt_bindings": True,
            "raw_source_identity": "explicit_null",
            "raw_sink_identity": "reconciled",
            "raw_phase_identity": "explicit_not_applicable",
            "raw_output_identity": "digest_and_matched_identity",
        },
    }


def capture_order(baseline: dict[str, Any], candidate: dict[str, Any],
                  eager: dict[str, Any]) -> list[dict[str, Any]]:
    by_group: dict[str, list[dict[str, Any]]] = {}
    for evidence in (baseline, candidate):
        for row in evidence["rows"]:
            by_group.setdefault(group_key(row["stage"], row["repeat"]), []).append(row)
    result: list[dict[str, Any]] = []
    previous_end: str | None = None
    for group in eager["native_order"]:
        rows = by_group.get(group)
        require(isinstance(rows, list) and len(rows) == len(eager["shapes"]),
                f"capture group {group} is incomplete")
        rows.sort(key=lambda row: row["receipt_start_utc"])
        require([row["shape"] for row in rows] == eager["shapes"],
                f"capture group {group} shape order differs")
        start = rows[0]["receipt_start_utc"]
        end = rows[-1]["receipt_end_utc"]
        if previous_end is not None:
            require(previous_end <= start, f"capture ABBA groups overlap at {group}")
        previous_end = end
        result.append({
            "group": group,
            "start_utc": start,
            "end_utc": end,
            "stage": rows[0]["stage"],
            "repeat": rows[0]["repeat"],
            "source_checkout": eager["stage_source_protocol"][group],
        })
    return result


def percent_change(baseline: Any, candidate: Any, label: str) -> float:
    finite_number(baseline, f"{label}.baseline")
    finite_number(candidate, f"{label}.candidate")
    require(float(baseline) > 0.0, f"{label}.baseline is not positive")
    return (float(candidate) / float(baseline) - 1.0) * 100.0


def metric_record(baseline: Any, candidate: Any, label: str) -> dict[str, Any]:
    change = percent_change(baseline, candidate, label)
    return {
        "baseline": baseline,
        "candidate": candidate,
        "change_percent": change,
        "adverse_over_five_percent": change > 5.0,
    }


def review_item(kind: str, lane: str, key: dict[str, Any], metric: str,
                stat: str, record: dict[str, Any], threshold: float) -> dict[str, Any]:
    return {
        "review_id": f"{kind}:{lane}:{key['repeat']}:{key['shape']}:{metric}:{stat}",
        "kind": kind,
        "lane": lane,
        "repeat": key["repeat"],
        "shape": key["shape"],
        "metric": metric,
        "stat": stat,
        "threshold_percent": threshold,
        "change_percent": record["change_percent"],
        "baseline": record["baseline"],
        "candidate": record["candidate"],
        "status": "individual-review-required",
        "reason": (
            "Metric increased by more than the frozen 5% review threshold; "
            "this guard has no automatic speedup admission gate."
        ),
    }


def compare_metric_set(before: dict[str, Any], after: dict[str, Any],
                      lane: str, key: dict[str, Any], threshold: float,
                      reviews: list[dict[str, Any]],
                      adverse: list[dict[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for stat in ("p50", "p95", "p99", "mean"):
        record = metric_record(before["timing"]["elapsed_ns"][stat],
                               after["timing"]["elapsed_ns"][stat],
                               f"{lane} {key['repeat']}/{key['shape']} elapsed {stat}")
        result[stat] = record
        if record["adverse_over_five_percent"]:
            item = review_item("adverse_metric", lane, key, "elapsed_ns", stat,
                               record, threshold)
            adverse.append(item)
            reviews.append(item)
    rss = metric_record(before["rss"]["max_rss_kib"],
                        after["rss"]["max_rss_kib"],
                        f"{lane} {key['repeat']}/{key['shape']} peak RSS")
    if rss["adverse_over_five_percent"]:
        item = review_item("adverse_metric", lane, key, "max_rss_kib", "peak",
                           rss, threshold)
        adverse.append(item)
        reviews.append(item)
    return {"elapsed_ns": result, "max_rss_kib": rss}


def compare(baseline: dict[str, Any], candidate: dict[str, Any],
            eager: dict[str, Any]) -> dict[str, Any]:
    threshold = float(eager["adverse_threshold_percent"])
    left = {(row["repeat"], row["shape"]): row for row in baseline["rows"]}
    right = {(row["repeat"], row["shape"]): row for row in candidate["rows"]}
    require(set(left) == set(right), "eager baseline/candidate matrices differ")
    require(len(left) == len(baseline["rows"]) == len(candidate["rows"]),
            "eager rows contain duplicate logical identities")
    comparisons: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    reviews: list[dict[str, Any]] = []
    for repeat, shape in sorted(left):
        before, after = left[(repeat, shape)], right[(repeat, shape)]
        require(before["identity"] == after["identity"],
                f"eager logical identity differs for {(repeat, shape)}")
        key = {"repeat": repeat, "shape": shape}
        comparisons.append({
            "repeat": repeat,
            "shape": shape,
            "identity_equal": True,
            "metrics": compare_metric_set(before, after, "candidate-vs-baseline",
                                            key, threshold, reviews, adverse),
        })

    drift_records: list[dict[str, Any]] = []
    drift_flags: list[dict[str, Any]] = []
    for stage_name, evidence in (("baseline", baseline), ("candidate", candidate)):
        rows = {(row["repeat"], row["shape"]): row for row in evidence["rows"]}
        for shape in eager["shapes"]:
            before = rows[(1, shape)]
            after = rows[(2, shape)]
            metrics: dict[str, Any] = {}
            for stat in ("p50", "p95", "p99", "mean"):
                record = metric_record(
                    before["timing"]["elapsed_ns"][stat],
                    after["timing"]["elapsed_ns"][stat],
                    f"{stage_name} same-build {shape} elapsed {stat}")
                metrics.setdefault("elapsed_ns", {})[stat] = record
                if abs(record["change_percent"]) > threshold:
                    key = {"repeat": "1-to-2", "shape": shape}
                    item = review_item("same_build_drift", stage_name, key,
                                       "elapsed_ns", stat, record, threshold)
                    item["repeat_first"] = 1
                    item["repeat_second"] = 2
                    drift_flags.append(item)
                    reviews.append(item)
            rss = metric_record(before["rss"]["max_rss_kib"],
                                after["rss"]["max_rss_kib"],
                                f"{stage_name} same-build {shape} peak RSS")
            metrics["max_rss_kib"] = {"peak": rss}
            if abs(rss["change_percent"]) > threshold:
                key = {"repeat": "1-to-2", "shape": shape}
                item = review_item("same_build_drift", stage_name, key,
                                   "max_rss_kib", "peak", rss, threshold)
                item["repeat_first"] = 1
                item["repeat_second"] = 2
                drift_flags.append(item)
                reviews.append(item)
            drift_records.append({
                "stage": stage_name,
                "shape": shape,
                "repeat_first": 1,
                "repeat_second": 2,
                "metrics": metrics,
            })
    return {
        "timing_comparisons": comparisons,
        "same_build_drift": drift_records,
        "adverse_flags_over_five_percent": adverse,
        "same_build_drift_over_five_percent": drift_flags,
        "individual_reviews": reviews,
        "individual_review_count": len(reviews),
        "all_metrics_compared": ["elapsed_ns.p50", "elapsed_ns.p95",
                                 "elapsed_ns.p99", "elapsed_ns.mean",
                                 "max_rss_kib.peak"],
        "threshold_percent": threshold,
        "no_speedup_requirement": True,
        "phase_metrics": {
            "status": "not_applicable",
            "reason": "eager raw reports have explicit null source and no source-backed phase vectors",
        },
        "runtime_comparison_performed": True,
    }


def analyze(output: Path) -> dict[str, Any]:
    primary, eager = plans()
    baseline = validate_stage("baseline", eager)
    candidate = validate_stage("candidate", eager)
    ordering = capture_order(baseline, candidate, eager)
    result = {
        "status": "pass",
        "stage": "compare",
        "plan_sha256": sha256(EAGER_PLAN_PATH),
        "primary_plan_sha256": sha256(PLAN_PATH),
        "capture_authority": {
            "path": relative(RUN_PATH),
            "sha256": sha256(RUN_PATH),
        },
        "guard_authority": {
            "path": relative(Path(__file__)),
            "sha256": sha256(Path(__file__)),
        },
        "baseline": baseline,
        "candidate": candidate,
        "capture_order": ordering,
        "comparison": compare(baseline, candidate, eager),
        "scope": eager["purpose"],
        "no_speedup_requirement": True,
        "output": relative(output),
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--capture", action="store_true",
                       help="capture one frozen ABBA stage/repeat")
    modes.add_argument("--analyze", action="store_true",
                       help="validate both stages and write eager-comparison.json")
    parser.add_argument("--stage", choices=("baseline", "candidate"))
    parser.add_argument("--repeat", type=int, choices=(1, 2))
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        if args.capture:
            require(args.stage is not None and args.repeat is not None,
                    "--capture requires --stage and --repeat")
            require(args.output is None, "--capture does not take --output")
            capture(args.stage, args.repeat)
            return 0
        require(args.stage is None and args.repeat is None,
                "--analyze does not take --stage or --repeat")
        output = args.output or HERE / "eager-comparison.json"
        output.parent.mkdir(parents=True, exist_ok=True)
        value = analyze(output)
        write_json(output, value)
    except (EvidenceError, OSError, json.JSONDecodeError, AssertionError) as error:
        print(f"eager_guard.py: error: {error}", file=sys.stderr)
        return 2
    print(f"0531 eager guard verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
