#!/usr/bin/env python3
"""Validate and compare the frozen 0544 native XLSX completion pilot.

The 0544 candidate forwards the result of worksheet materialization after the
shared traversal has accepted EOF, while retaining the source-backed semantic
and error-order contracts.  This analyzer keeps the native admission contract
narrow: every XLSX
primary child must improve the measured workflow elapsed p50 and mean by the
thresholds in ``plan.json``, and its planning p50 must improve by the planning
threshold.  The elapsed timer covers the harness's open, planning, commit, and
publication workflow; the separate ``/usr/bin/time`` observation covers the
whole fresh child for RSS diagnostics.
The XLSX guard rows exercise unmanaged, managed, vendor-extension, and
noncompact paths.  A standalone planning refusal guard is captured by the
separate bound harness lane and its source is included identically in both builds.  Eager read guards are validated by their separate lane and
are not folded into this source-backed row matrix.

The script is safe to invoke before capture.  Missing or incomplete evidence
is reported as ``pending`` and no synthetic row or performance result is
created.  Once both native stages are complete, comparison retains every
timing metric, every greater-than-five-percent adverse change, same-build
repeat drift, and matched-child bootstrap interval.  Publication p50 is a
diagnostic phase only; there is no inherited publication gate.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import random
import sys
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
HELPER = HERE.parent / "change-0521" / "analyze.py"
_helper_spec = importlib.util.spec_from_file_location("xlsx_0521_numerical_0544", HELPER)
if _helper_spec is None or _helper_spec.loader is None:
    raise ImportError(f"cannot load numerical helper: {HELPER}")
BASE = importlib.util.module_from_spec(_helper_spec)
_helper_spec.loader.exec_module(BASE)

# The retained helper's artifact paths are module globals.  Rebinding them
# keeps this analyzer independent of every earlier campaign directory.
BASE.HERE = HERE
BASE.PLAN = HERE / "plan.json"
BASE.BOOTSTRAP_SEED = 5_440_544

_BASE_SOURCE_IDENTITY = BASE.source_identity
_BASE_XLSX_VALIDATE_RESULT = BASE.validate_result

PRIMARY_CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
TIMING_STATS = ("p50", "p95", "p99", "mean")
ADVERSE_THRESHOLD_PERCENT = 5.0
ALLOCATION_VECTOR_NAMES = (
    "plan_allocation_metrics",
    "commit_allocation_metrics",
    "publication_allocation_metrics",
)
GATE_FIELDS = (
    "total_p50_reduction_percent",
    "total_mean_reduction_percent",
    "planning_p50_reduction_percent",
    "planning_ir_reduction_percent",
    "require_every_shape_repeat",
)
OPTIONAL_GATE_FIELDS: tuple[str, ...] = ()
OPTIONAL_ALLOCATION_SCOPE = (
    "commit-only legacy diagnostics; planning/publication separately validated by "
    "allocation-analysis.json"
)


def __getattr__(name: str) -> Any:
    return getattr(BASE, name)


def _copy_without_allocation_vectors(source: dict[str, Any]) -> dict[str, Any]:
    """Remove per-sample allocator diagnostics from logical identity.

    The vectors are performance observations rather than semantic oracles.
    Both normal and allocator reports therefore share one identity contract,
    while the optional allocator lane below still validates the vectors.
    """

    xlsx = source.get("xlsx_cell_values")
    if not isinstance(xlsx, dict):
        return source
    result = dict(source)
    result_xlsx = dict(xlsx)
    for name in ALLOCATION_VECTOR_NAMES:
        result_xlsx.pop(name, None)
    result["xlsx_cell_values"] = result_xlsx
    return result


def _source_identity(source: dict[str, Any], count: int) -> dict[str, Any]:
    return _BASE_SOURCE_IDENTITY(_copy_without_allocation_vectors(source), count)


BASE.source_identity = _source_identity


def _require(condition: bool, message: str) -> None:
    BASE.require(condition, message)


def _sha(path: Path) -> str:
    return BASE.sha(path)


def _read_json(path: Path) -> Any:
    return BASE.read_json(path)


def _check_binary(folder: Path, allocator: bool) -> dict[str, Any]:
    """Validate a retained binary using the 0544 runner's target layout.

    The frozen 0544 plan owns the target directory, while ``run.py`` keeps
    retained binaries in its ``retained-binaries`` child.  The 0521 helper
    predates that layout and assumes its first owned path is the binary
    directory, so retain all of its custody checks here with the runner's
    explicit path binding.
    """

    label = "alloc" if allocator else "normal"
    identity = _read_json(folder / f"binary-{label}.json")
    _require(isinstance(identity, dict), f"{folder.name} binary identity is not an object")
    path = Path(identity.get("path", ""))
    plan = _read_json(BASE.PLAN)
    owned_paths = plan.get("owned_paths")
    _require(isinstance(owned_paths, list) and owned_paths
             and all(isinstance(item, str) and item for item in owned_paths),
             "plan.owned_paths is not a nonempty string list")
    expected_path = Path(owned_paths[0]) / "retained-binaries" / f"{folder.name}-{label}"
    _require(path == expected_path,
             f"{folder.name} binary path is unexpected: {path} != {expected_path}")
    digest = identity.get("sha256")
    _require(isinstance(digest, str) and len(digest) == 64
             and all(char in "0123456789abcdef" for char in digest),
             f"{folder.name} binary digest is malformed")
    BASE.nonnegative_integer(identity.get("bytes"), f"{folder.name} binary size")
    _require(identity["bytes"] > 0, f"{folder.name} binary size is zero")
    _require(not path.is_symlink(), f"{folder.name} binary is a symlink")
    if path.exists():
        _require(path.is_file(), f"{folder.name} binary is not a regular file")
        _require(identity["sha256"] == _sha(path),
                 f"{folder.name} binary digest mismatch")
        _require(identity["bytes"] == path.stat().st_size,
                 f"{folder.name} binary size mismatch")
    else:
        # Captures hash the live executable before and after each child. Once
        # owned builds are removed, replay checks their retained custody chain.
        cleanup = _read_json(HERE / "cleanup.json")
        _require(cleanup.get("owned_paths_absent") is True
                 and cleanup.get("accessible_process_references") == []
                 and cleanup.get("removed") == owned_paths
                 and all(not Path(name).exists() for name in owned_paths),
                 f"{folder.name} missing binary has no completed cleanup record")
    manifest = folder / "source-manifest.json"
    _require(identity.get("source_manifest_sha256") == _sha(manifest),
             f"{folder.name} binary manifest mismatch")
    build_name = "build-alloc" if allocator else "build-normal"
    build_receipt_path = folder / f"{build_name}.receipt.json"
    build_receipt = _read_json(build_receipt_path)
    _require(identity.get("build_receipt_sha256") == _sha(build_receipt_path),
             f"{folder.name} binary build receipt mismatch")
    _require(build_receipt.get("exit_code") == 0
             and build_receipt.get("binary_sha256") is None,
             f"{folder.name} build receipt is invalid")
    _require(build_receipt.get("plan_sha256") == _sha(BASE.PLAN),
             f"{folder.name} build plan mismatch")
    _require(build_receipt.get("script_sha256") == _sha(HERE / "run.py"),
             f"{folder.name} build script mismatch")
    _require(build_receipt.get("source_manifest_sha256") == _sha(manifest),
             f"{folder.name} build manifest mismatch")
    executable = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    command = build_receipt.get("command")
    _require(isinstance(command, list), f"{folder.name} build receipt command is not a list")
    for option in ("cargo", "build", "--release", "--locked", "--manifest-path", "--bin"):
        _require(option in command, f"{folder.name} build receipt is missing {option}")
    _require(command[command.index("--bin") + 1] == executable,
             f"{folder.name} build receipt binary is unexpected")
    if allocator:
        _require("--features" in command
                 and command[command.index("--features") + 1] == "allocator-metrics",
                 f"{folder.name} allocator build is missing allocator-metrics")
    else:
        _require("--features" not in command,
                 f"{folder.name} normal build unexpectedly enables allocator metrics")
    expected_artifacts = {f"{build_name}.stdout", f"{build_name}.stderr"}
    _require(set(build_receipt.get("artifacts", {})) == expected_artifacts,
             f"{folder.name} build artifact inventory mismatch")
    for filename, artifact_digest in build_receipt["artifacts"].items():
        artifact = folder / filename
        _require(artifact.is_file() and _sha(artifact) == artifact_digest,
                 f"{folder.name} build artifact mismatch: {filename}")
    return {"sha256": identity["sha256"], "bytes": identity["bytes"],
            "path": str(path), "binary": executable,
            "build_receipt_sha256": identity["build_receipt_sha256"]}


def _nonnegative_number(value: Any, label: str) -> None:
    BASE.finite_number(value, label)
    _require(float(value) >= 0.0, f"{label} is negative")


def _reduction_percent(baseline: Any, candidate: Any, label: str) -> float:
    _nonnegative_number(baseline, f"{label}.baseline")
    _nonnegative_number(candidate, f"{label}.candidate")
    _require(float(baseline) > 0.0, f"{label}.baseline must be positive")
    return (float(baseline) - float(candidate)) / float(baseline) * 100.0


def _percent_change(baseline: Any, candidate: Any, label: str) -> float | None:
    _nonnegative_number(baseline, f"{label}.baseline")
    _nonnegative_number(candidate, f"{label}.candidate")
    if float(baseline) == 0.0:
        return 0.0 if float(candidate) == 0.0 else None
    return (float(candidate) / float(baseline) - 1.0) * 100.0


def _comparison_record(baseline: Any, candidate: Any, label: str) -> dict[str, Any]:
    return {
        "baseline": baseline,
        "candidate": candidate,
        "change_percent": _percent_change(baseline, candidate, label),
    }


def _validated_gates(plan: dict[str, Any]) -> dict[str, Any]:
    gates = plan.get("gates")
    _require(isinstance(gates, dict), "plan.gates is not an object")
    expected_fields = set(GATE_FIELDS) | set(OPTIONAL_GATE_FIELDS)
    _require(set(gates) == expected_fields,
             "plan gate inventory differs from frozen 0544 plan")
    for field in GATE_FIELDS[:-1]:
        value = gates.get(field)
        BASE.finite_number(value, f"plan.gates.{field}")
        _require(float(value) >= 0.0, f"plan.gates.{field} is negative")
    _require(gates.get("require_every_shape_repeat") is True,
             "plan.gates.require_every_shape_repeat must be true")
    for field in OPTIONAL_GATE_FIELDS:
        value = gates.get(field)
        BASE.finite_number(value, f"plan.gates.{field}")
        _require(float(value) >= 0.0, f"plan.gates.{field} is negative")
    return gates


def _read_plan() -> dict[str, Any]:
    plan = _read_json(BASE.PLAN)
    _require(isinstance(plan, dict), "plan is not an object")
    _validated_gates(plan)
    _require(plan.get("primary", {}).get("case") == PRIMARY_CASE,
             "frozen primary case is not the expected XLSX source-backed case")
    return plan


def _expected_native_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    """Use the prior helper's exact naming convention for native children."""

    primary = plan.get("primary")
    _require(isinstance(primary, dict), "plan.primary is not an object")
    guards = plan.get("guards")
    _require(isinstance(guards, list), "plan.guards is not a list")
    jobs: list[dict[str, Any]] = []
    for repeat in range(1, int(primary["repeats"]) + 1):
        for shape in primary["shapes"]:
            jobs.append({
                "name": f"native-r{repeat}-primary-{shape}",
                "kind": "primary",
                "guard": None,
                "repeat": repeat,
                "case": primary["case"],
                "shape": shape,
                "warmup": int(primary["warmup"]),
                "samples": int(primary["samples"]),
            })
    guard_repeats = int(plan.get("guard_repeats", primary["repeats"]))
    guard_warmup = int(plan.get("guard_warmup", primary["warmup"]))
    guard_samples = int(plan.get("guard_samples", primary["samples"]))
    for repeat in range(1, guard_repeats + 1):
        for guard, item in enumerate(guards):
            _require(isinstance(item, dict), f"plan.guards[{guard}] is not an object")
            for shape in item["shapes"]:
                jobs.append({
                    "name": f"native-r{repeat}-guard{guard}-{shape}",
                    "kind": "guard",
                    "guard": guard,
                    "repeat": repeat,
                    "case": item["case"],
                    "shape": shape,
                    "warmup": guard_warmup,
                    "samples": guard_samples,
                })
    return jobs


def _expected_allocation_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    config = plan.get("allocation")
    primary = plan.get("primary")
    _require(isinstance(config, dict), "plan.allocation is not an object")
    _require(isinstance(primary, dict), "plan.primary is not an object")
    return [
        {
            "name": f"alloc-r{repeat}-{shape}",
            "kind": "allocation",
            "guard": None,
            "repeat": repeat,
            "case": primary["case"],
            "shape": shape,
            "warmup": int(config["warmup"]),
            "samples": int(config["samples"]),
        }
        for repeat in range(1, int(config["repeats"]) + 1)
        for shape in config["shapes"]
    ]


def _job_key(row: dict[str, Any]) -> tuple[Any, ...]:
    return (row["kind"], row["guard"], row["repeat"], row["case"], row["shape"])


def _is_xlsx_job(job: dict[str, Any]) -> bool:
    return job["case"].startswith("xlsx_")


def _format_for_job(job: dict[str, Any]) -> str:
    case = job["case"]
    if case.startswith("xlsx_"):
        return "XLSX/OPC/ZIP"
    if case.startswith("docx_"):
        return "DOCX/OPC/ZIP"
    if case.startswith("pptx_"):
        return "PPTX/OPC/ZIP"
    raise BASE.EvidenceError(f"unsupported 0544 guard case: {case}")


def _validate_configuration(configuration: Any, plan: dict[str, Any],
                            job: dict[str, Any]) -> dict[str, Any]:
    label = job["name"]
    _require(isinstance(configuration, dict), f"{label}.configuration is not an object")
    _require(configuration.get("cases") == [job["case"]],
             f"{label} case configuration is unexpected")
    _require(configuration.get("samples_per_case") == job["samples"],
             f"{label} sample count is unexpected")
    _require(configuration.get("warmup_iterations_per_case") == job["warmup"],
             f"{label} warmup is unexpected")
    # run.py deliberately binds every native child through this option, even
    # for every XLSX guard.  Requiring the recorded value catches a child
    # launched with a wrong shape or an unplanned default.
    _require(configuration.get("xlsx_cell_crud_shapes") == [job["shape"]],
             f"{label} xlsx-cell-crud shape configuration is unexpected")
    return configuration


def _validate_tool_and_environment(raw: dict[str, Any], plan: dict[str, Any],
                                   job: dict[str, Any], binary: dict[str, Any],
                                   allocator: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    label = job["name"]
    _require(raw.get("schema_version") == 1, f"{label} schema version is not 1")
    tool = raw.get("tool")
    _require(isinstance(tool, dict), f"{label}.tool is not an object")
    expected_binary = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    _require(tool.get("binary") == expected_binary, f"{label} tool binary is unexpected")
    _require(tool.get("profile") == "release", f"{label} tool profile is not release")
    _require(tool.get("instrumentation") == (
        "system_allocator_operation_scoped" if allocator else "none"),
             f"{label} tool instrumentation is unexpected")
    binary_identity = raw.get("binary_identity")
    _require(isinstance(binary_identity, dict), f"{label}.binary_identity is not an object")
    _require(binary_identity.get("binary_sha256") == binary["sha256"],
             f"{label} report is bound to the wrong binary")
    _require(binary_identity.get("profile") == "release",
             f"{label} binary profile is unexpected")
    environment = raw.get("environment")
    _require(isinstance(environment, dict), f"{label}.environment is not an object")
    _require(environment.get("git_revision") == plan["revision"],
             f"{label} revision is unexpected")
    _require(environment.get("cpu_affinity") == str(plan["cpu"]),
             f"{label} CPU affinity is unexpected")
    return tool, environment


def _validate_operation_metrics(raw: dict[str, Any], count: int,
                                sink: dict[str, Any], label: str) -> dict[str, Any]:
    """Check the raw aligned metrics envelope when the harness emits it.

    The top-level ``sink`` remains the source of the native identity.  This
    check verifies that its promoted per-sample view has the same cardinality
    and values, so a malformed metrics envelope cannot silently pass as an
    oracle.  Other diagnostic subtrees are retained by the raw report and are
    intentionally not converted into latency claims here.
    """

    metrics = raw.get("results", [{}])[0].get("operation_metrics")
    if metrics is None:
        return {"status": "absent"}
    _require(isinstance(metrics, dict), f"{label}.operation_metrics is not an object")
    _require(metrics.get("sample_count") == count,
             f"{label}.operation_metrics sample count differs")
    _require(metrics.get("sample_indices") == list(range(count)),
             f"{label}.operation_metrics sample indices differ")
    _require(metrics.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
             f"{label}.operation_metrics alignment is unexpected")
    _require(metrics.get("latency_claim") == "comparable_timed_operation",
             f"{label}.operation_metrics latency claim is unexpected")
    sink_metrics = metrics.get("sink")
    if isinstance(sink_metrics, dict):
        for field in ("accepted_bytes", "write_calls", "largest_write"):
            item = sink_metrics.get(field)
            _require(isinstance(item, dict), f"{label}.operation_metrics.sink.{field} is missing")
            _require(item.get("status") == "measured",
                     f"{label}.operation_metrics.sink.{field} is not measured")
            values = item.get("values")
            _require(isinstance(values, list) and len(values) == count,
                     f"{label}.operation_metrics.sink.{field}.values has wrong length")
            _require(all(value == sink[field] for value in values),
                     f"{label}.operation_metrics.sink.{field} disagrees with sink summary")
        buckets = sink_metrics.get("write_size_buckets")
        if isinstance(buckets, dict):
            _require(buckets.get("status") == "measured",
                     f"{label}.operation_metrics.sink.write_size_buckets is not measured")
            for field, value in sink["write_size_buckets"].items():
                item = buckets.get(field)
                _require(isinstance(item, dict),
                         f"{label}.operation_metrics.sink.write_size_buckets.{field} is missing")
                _require(item.get("status") == "measured",
                         f"{label}.operation_metrics.sink.write_size_buckets.{field} is not measured")
                values = item.get("values")
                _require(isinstance(values, list) and len(values) == count,
                         f"{label}.operation_metrics.sink.write_size_buckets.{field}.values has wrong length")
                _require(all(value == sink["write_size_buckets"][field] for value in values),
                         f"{label}.operation_metrics.sink.write_size_buckets.{field} disagrees with sink summary")
    return {"status": "pass", "sample_count": count}


def _validate_sink(sink: Any, label: str) -> dict[str, Any]:
    _require(isinstance(sink, dict), f"{label}.sink is not an object")
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        BASE.nonnegative_integer(sink.get(key), f"{label}.sink.{key}")
    buckets = sink.get("write_size_buckets")
    _require(isinstance(buckets, dict), f"{label} sink buckets are missing")
    for key, value in buckets.items():
        BASE.nonnegative_integer(value, f"{label}.sink.write_size_buckets.{key}")
    _require(sink["largest_write"] <= 65_536,
             f"{label} exceeds the bounded sink write size")
    _require(sum(buckets.values()) == sink["write_calls"],
             f"{label} sink buckets do not reconcile with write calls")
    return sink


def _validate_generic_result(raw: dict[str, Any], plan: dict[str, Any],
                             job: dict[str, Any], binary: dict[str, Any],
                             allocator: bool) -> dict[str, Any]:
    """Validate any non-XLSX guard without assuming XLSX phase evidence."""

    label = job["name"]
    _validate_tool_and_environment(raw, plan, job, binary, allocator)
    configuration = _validate_configuration(raw.get("configuration"), plan, job)
    results = raw.get("results")
    _require(isinstance(results, list) and len(results) == 1,
             f"{label} must contain one result")
    result = results[0]
    _require(isinstance(result, dict), f"{label} result is not an object")
    _require(result.get("case") == job["case"], f"{label} result case is unexpected")
    corpus = result.get("corpus")
    # The fixed source-edit corpus is selected through the plan's semantic
    # guard shape ``medium`` but its manifest names the actual media-rich
    # archive.  Keep both identities explicit: the command/configuration must
    # carry the planned shape, while the result must carry the fixed corpus
    # manifest identity.
    _require(isinstance(corpus, dict) and corpus.get("shape") == "media-rich",
             f"{label} corpus identity is unexpected")
    _require(corpus.get("package_format") == _format_for_job(job),
             f"{label} corpus package format is unexpected")
    elapsed_stats, elapsed_values = BASE.verify_elapsed(result, job["samples"], label)
    source = result.get("source")
    _require(isinstance(source, dict), f"{label} has no source evidence")
    sink = _validate_sink(result.get("sink"), label)
    output_sha = result.get("output_sha256")
    _require(isinstance(output_sha, str) and len(output_sha) == 64
             and all(char in "0123456789abcdef" for char in output_sha),
             f"{label} output digest is missing")
    source_identity = _source_identity(source, job["samples"])
    identity = {
        "configuration": configuration,
        "corpus": corpus,
        "sink": sink,
        "source": source_identity,
        "output_sha256": output_sha,
    }
    metrics_status = _validate_operation_metrics(raw, job["samples"], sink, label)
    return {
        "name": label,
        "kind": job["kind"],
        "guard": job["guard"],
        "repeat": job["repeat"],
        "case": job["case"],
        "shape": job["shape"],
        "samples": job["samples"],
        "timing": {"elapsed_ns": elapsed_stats},
        "phase_time_share": {},
        "rss": None,
        "allocation": [],
        "identity": identity,
        "identity_sha256": BASE.identity_digest(identity),
        "raw_metrics": metrics_status,
        "elapsed_values": elapsed_values,
    }


def _validate_result(raw: dict[str, Any], plan: dict[str, Any],
                     job: dict[str, Any], binary: dict[str, Any],
                     allocator: bool) -> dict[str, Any]:
    case = job["case"]
    if _is_xlsx_job(job):
        row = _BASE_XLSX_VALIDATE_RESULT(raw, plan, job, binary, allocator)
        # The retained validator already checked phase sums, source/output
        # hashes, sink bucket arithmetic, allocation balance and report stats.
        row["raw_metrics"] = _validate_operation_metrics(
            raw, job["samples"], raw["results"][0]["sink"], job["name"])
        return row
    if case.startswith("docx_") or case.startswith("pptx_"):
        return _validate_generic_result(raw, plan, job, binary, allocator)
    raise BASE.EvidenceError(f"unsupported 0544 case: {case}")


# BASE.check_receipt resolves ``validate_result`` in its module namespace.
# Install the dispatch above so its custody/artifact checks remain reused for
# both XLSX and the format-aware guards.
BASE.validate_result = _validate_result
# Bind the retained helper's binary validator to the 0544 runner layout.
BASE.check_binary = _check_binary


def _check_stage(stage: str, plan: dict[str, Any], jobs: list[dict[str, Any]],
                 allocator: bool = False) -> dict[str, Any]:
    folder = HERE / stage
    _require(folder.is_dir(), f"{stage} stage directory is missing")
    manifest = folder / "source-manifest.json"
    _require(manifest.is_file() and not manifest.is_symlink(),
             f"{stage} source manifest is missing")
    manifest_value = _read_json(manifest)
    _require(isinstance(manifest_value, dict) and manifest_value,
             f"{stage} source manifest is empty")
    binary = BASE.check_binary(folder, allocator)
    prefix = "alloc-" if allocator else "native-"
    expected_names = {job["name"] for job in jobs}
    actual_names = {
        path.name[:-len(".receipt.json")]
        for path in folder.glob("*.receipt.json")
        if path.name.startswith(prefix)
    }
    _require(actual_names == expected_names,
             f"{stage} {'allocation' if allocator else 'native'} receipt set differs from plan: "
             f"{sorted(actual_names ^ expected_names)}")
    rows: list[dict[str, Any]] = []
    receipts: list[dict[str, Any]] = []
    for job in jobs:
        receipt, row = BASE.check_receipt(folder, plan, job, binary, allocator)
        row["stage"] = stage
        if not allocator:
            row["rss"] = {"scope": "whole_child_process",
                           **BASE.validate_rss(folder / f"{job['name']}.rss.json")}
        rows.append(row)
        receipts.append(receipt)
    ordered_receipts = sorted(receipts, key=lambda value: value["start_utc"])
    _require(all(left["end_utc"] <= right["start_utc"]
                 for left, right in zip(ordered_receipts, ordered_receipts[1:])),
             f"{stage} child receipts overlap")
    rows.sort(key=lambda row: (_job_key(row), row["name"]))
    lane_name = "allocation" if allocator else "native"
    return {
        "stage": stage,
        "manifest_sha256": _sha(manifest),
        "binary_identity": binary,
        lane_name: {"rows": rows, "row_count": len(rows),
                    "total_samples": sum(row["samples"] for row in rows)},
        "custody": {"receipt_count": len(receipts),
                     "receipts_non_overlapping": True,
                     "source_manifest_entries": len(manifest_value)},
    }


def _missing_stage_artifacts(stage: str, plan: dict[str, Any],
                             allocator: bool = False) -> list[str]:
    folder = HERE / stage
    if not folder.is_dir():
        return [str(folder.relative_to(HERE))]
    jobs = _expected_allocation_jobs(plan) if allocator else _expected_native_jobs(plan)
    missing: list[str] = []
    for name in ("source-manifest.json", "binary-" + ("alloc" if allocator else "normal") + ".json",
                 "build-" + ("alloc" if allocator else "normal") + ".receipt.json"):
        if not (folder / name).is_file():
            missing.append(f"{stage}/{name}")
    for job in jobs:
        for suffix in (".receipt.json", ".json", ".stdout", ".stderr"):
            if not (folder / (job["name"] + suffix)).is_file():
                missing.append(f"{stage}/{job['name']}{suffix}")
        if not allocator and not (folder / (job["name"] + ".rss.json")).is_file():
            missing.append(f"{stage}/{job['name']}.rss.json")
    return missing


def _stage_or_pending(stage: str, plan: dict[str, Any], allocator: bool = False) -> dict[str, Any]:
    jobs = _expected_allocation_jobs(plan) if allocator else _expected_native_jobs(plan)
    missing = _missing_stage_artifacts(stage, plan, allocator)
    if missing:
        return {
            "status": "pending",
            "stage": stage,
            "missing_artifacts": missing,
            "expected_row_count": len(jobs),
            "lane": "allocation" if allocator else "native",
        }
    return {"status": "pass", "evidence": _check_stage(stage, plan, jobs, allocator)}


def _timing(row: dict[str, Any], phase: str, label: str) -> dict[str, Any]:
    timing = row.get("timing")
    _require(isinstance(timing, dict), f"{label}.timing is not an object")
    value = timing.get(phase)
    _require(isinstance(value, dict), f"{label}.timing.{phase} is not an object")
    for stat in TIMING_STATS:
        _nonnegative_number(value.get(stat), f"{label}.timing.{phase}.{stat}")
    return value


def _primary_rows(evidence: dict[str, Any], plan: dict[str, Any], label: str) -> dict[tuple[int, str], dict[str, Any]]:
    native = evidence.get("native")
    _require(isinstance(native, dict), f"{label}.native is not an object")
    rows = [row for row in native.get("rows", [])
            if isinstance(row, dict) and row.get("kind") == "primary"
            and row.get("guard") is None]
    primary = plan["primary"]
    expected = {(repeat, shape)
                for repeat in range(1, int(primary["repeats"]) + 1)
                for shape in primary["shapes"]}
    actual = [(row.get("repeat"), row.get("shape")) for row in rows]
    _require(len(rows) == len(expected),
             f"{label} primary native row count differs from plan: {len(rows)} != {len(expected)}")
    _require(len(set(actual)) == len(actual), f"{label} primary native rows contain duplicates")
    _require(set(actual) == expected,
             f"{label} primary native matrix differs from plan: {sorted(set(actual) ^ expected)}")
    return {(row["repeat"], row["shape"]): row for row in rows}


def _native_admission(baseline: dict[str, Any], candidate: dict[str, Any],
                      plan: dict[str, Any]) -> dict[str, Any]:
    gates = _validated_gates(plan)
    left_rows = _primary_rows(baseline, plan, "baseline")
    right_rows = _primary_rows(candidate, plan, "candidate")
    rows: list[dict[str, Any]] = []
    for repeat, shape in sorted(left_rows):
        left = left_rows[(repeat, shape)]
        right = right_rows[(repeat, shape)]
        left_elapsed = _timing(left, "elapsed_ns", f"primary {repeat}/{shape} baseline")
        right_elapsed = _timing(right, "elapsed_ns", f"primary {repeat}/{shape} candidate")
        left_plan = _timing(left, "plan_ns", f"primary {repeat}/{shape} baseline")
        right_plan = _timing(right, "plan_ns", f"primary {repeat}/{shape} candidate")
        total_p50 = _reduction_percent(left_elapsed["p50"], right_elapsed["p50"],
                                        f"primary {repeat}/{shape} total p50")
        total_mean = _reduction_percent(left_elapsed["mean"], right_elapsed["mean"],
                                         f"primary {repeat}/{shape} total mean")
        planning_p50 = _reduction_percent(left_plan["p50"], right_plan["p50"],
                                           f"primary {repeat}/{shape} planning p50")
        checks = {
            "native_primary_total_p50": {
                "baseline": left_elapsed["p50"], "candidate": right_elapsed["p50"],
                "reduction_percent": total_p50,
                "required_reduction_percent": float(gates["total_p50_reduction_percent"]),
                "passed": total_p50 >= float(gates["total_p50_reduction_percent"]),
            },
            "native_primary_total_mean": {
                "baseline": left_elapsed["mean"], "candidate": right_elapsed["mean"],
                "reduction_percent": total_mean,
                "required_reduction_percent": float(gates["total_mean_reduction_percent"]),
                "passed": total_mean >= float(gates["total_mean_reduction_percent"]),
            },
            "native_primary_planning_p50": {
                "metric": "plan_ns",
                "baseline": left_plan["p50"], "candidate": right_plan["p50"],
                "reduction_percent": planning_p50,
                "required_reduction_percent": float(gates["planning_p50_reduction_percent"]),
                "passed": planning_p50 >= float(gates["planning_p50_reduction_percent"]),
            },
            "planning_ir_conditional": {
                "metric": "planning_ir",
                "status": "unmeasured",
                "baseline": None,
                "candidate": None,
                "required_reduction_percent": float(gates["planning_ir_reduction_percent"]),
                "used_for_gate": False,
                "note": "Callgrind planning Ir is conditional and is not manufactured by native analysis.",
            },
        }
        rows.append({
            "repeat": repeat,
            "shape": shape,
            **checks,
            "passed": all(value["passed"] for value in checks.values()
                           if value.get("used_for_gate", True)),
        })
    return {
        "rows": rows,
        "passed": all(row["passed"] for row in rows),
        "decision": "eligible-for-conditional-lanes" if all(row["passed"] for row in rows) else "reject",
        "scope": (
            "Every planned XLSX primary shape/repeat must meet measured workflow elapsed p50, "
            "total mean, and plan_ns p50 gates. Publication p50 and planning Ir "
            "are diagnostics/conditional evidence and do not enter native admission."
        ),
    }


def _flag_adverse(adverse: list[dict[str, Any]], *, lane: str, key: tuple[Any, ...],
                  phase: str, stat: str, value: dict[str, Any]) -> None:
    change = value.get("change_percent")
    adverse_change = (change is not None and change > ADVERSE_THRESHOLD_PERCENT)
    zero_baseline = value.get("baseline") == 0 and value.get("candidate", 0) > 0
    if adverse_change or zero_baseline:
        adverse.append({
            "lane": lane,
            "case": key[3],
            "shape": key[4],
            "repeat": key[2],
            "guard": key[1],
            "phase": phase,
            "stat": stat,
            "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
            "baseline_zero_adverse": zero_baseline,
            **value,
        })


def _compare_timing_lane(baseline: dict[str, Any], candidate: dict[str, Any],
                         lane: str, rng: random.Random,
                         comparisons: list[dict[str, Any]],
                         adverse: list[dict[str, Any]],
                         drift: list[dict[str, Any]]) -> None:
    left_rows = {BASE.job_key(row): row for row in baseline[lane]["rows"]}
    right_rows = {BASE.job_key(row): row for row in candidate[lane]["rows"]}
    _require(set(left_rows) == set(right_rows), f"{lane} baseline/candidate row keys differ")
    for key in sorted(left_rows):
        left, right = left_rows[key], right_rows[key]
        _require(left["identity"] == right["identity"],
                 f"baseline/candidate logical identity differs for {lane} {key}")
        phases = tuple(sorted(set(left["timing"]) | set(right["timing"])))
        metrics: dict[str, Any] = {}
        for phase in phases:
            left_phase = _timing(left, phase, f"{lane} {key} baseline")
            right_phase = _timing(right, phase, f"{lane} {key} candidate")
            metrics[phase] = {}
            for stat in TIMING_STATS:
                value = _comparison_record(left_phase[stat], right_phase[stat],
                                           f"{lane} {key} {phase} {stat}")
                metrics[phase][stat] = value
                _flag_adverse(adverse, lane=lane, key=key, phase=phase,
                              stat=stat, value=value)
        record: dict[str, Any] = {
            "lane": lane,
            "case": key[3],
            "shape": key[4],
            "repeat": key[2],
            "guard": key[1],
            "metrics": metrics,
            "identity_equal": True,
            "bootstrap": {
                "elapsed_ns_p50": BASE.bootstrap_median_ratio(
                    left["timing"]["elapsed_ns"]["samples"],
                    right["timing"]["elapsed_ns"]["samples"], rng),
            },
        }
        if "plan_ns" in left["timing"] and "plan_ns" in right["timing"]:
            record["bootstrap"]["planning_ns_p50"] = BASE.bootstrap_median_ratio(
                left["timing"]["plan_ns"]["samples"],
                right["timing"]["plan_ns"]["samples"], rng)
        if left.get("rss") is not None or right.get("rss") is not None:
            _require(isinstance(left.get("rss"), dict) and isinstance(right.get("rss"), dict),
                     f"matched {lane} RSS is incomplete for {key}")
            rss = _comparison_record(left["rss"]["max_rss_kib"],
                                     right["rss"]["max_rss_kib"],
                                     f"{lane} {key} rss")
            record["rss"] = rss
            _flag_adverse(adverse, lane=lane, key=key, phase="rss",
                          stat="max_rss_kib", value=rss)
        comparisons.append(record)

    # Every same-build repeat pair is retained with the complete value record.
    groups: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
    for row in baseline[lane]["rows"] + candidate[lane]["rows"]:
        groups.setdefault((row["stage"], row["kind"], row["guard"],
                           row["case"], row["shape"]), []).append(row)
    for group, rows in sorted(groups.items(), key=lambda item: str(item[0])):
        by_repeat = {row["repeat"]: row for row in rows}
        for first_repeat in sorted(by_repeat):
            for second_repeat in sorted(by_repeat):
                if second_repeat <= first_repeat:
                    continue
                first, second = by_repeat[first_repeat], by_repeat[second_repeat]
                for phase in sorted(set(first["timing"]) | set(second["timing"])):
                    first_phase = _timing(first, phase, f"drift {group} first")
                    second_phase = _timing(second, phase, f"drift {group} second")
                    for stat in TIMING_STATS:
                        value = _comparison_record(
                            first_phase[stat], second_phase[stat],
                            f"drift {group} {phase} {stat}")
                        change = value.get("change_percent")
                        if (change is not None and abs(change) > ADVERSE_THRESHOLD_PERCENT) \
                                or (value.get("baseline") == 0 and value.get("candidate", 0) > 0):
                            drift.append({
                                "lane": lane,
                                "stage": group[0],
                                "case": group[3],
                                "shape": group[4],
                                "kind": group[1],
                                "guard": group[2],
                                "repeat_first": first_repeat,
                                "repeat_second": second_repeat,
                                "phase": phase,
                                "stat": stat,
                                "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
                                "baseline_zero_drift": value.get("baseline") == 0
                                and value.get("candidate", 0) > 0,
                                **value,
                            })
                if first.get("rss") is not None or second.get("rss") is not None:
                    _require(isinstance(first.get("rss"), dict)
                             and isinstance(second.get("rss"), dict),
                             f"same-build RSS is incomplete for {group}")
                    value = _comparison_record(
                        first["rss"]["max_rss_kib"], second["rss"]["max_rss_kib"],
                        f"drift {group} rss")
                    change = value.get("change_percent")
                    if (change is not None and abs(change) > ADVERSE_THRESHOLD_PERCENT) \
                            or (value.get("baseline") == 0 and value.get("candidate", 0) > 0):
                        drift.append({
                            "lane": lane,
                            "stage": group[0],
                            "case": group[3],
                            "shape": group[4],
                            "kind": group[1],
                            "guard": group[2],
                            "repeat_first": first_repeat,
                            "repeat_second": second_repeat,
                            "phase": "rss",
                            "stat": "max_rss_kib",
                            "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
                            **value,
                        })


def _optional_allocation_stage(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    """Validate a complete optional allocator stage without requiring it."""

    folder = HERE / stage
    if not folder.is_dir() or not list(folder.glob("alloc-*.receipt.json")):
        return {
            "status": "not-captured",
            "stage": stage,
            "used_for_gate": False,
            "scope": OPTIONAL_ALLOCATION_SCOPE,
        }
    missing = _missing_stage_artifacts(stage, plan, True)
    if missing:
        return {
            "status": "pending",
            "stage": stage,
            "used_for_gate": False,
            "missing_artifacts": missing,
            "scope": OPTIONAL_ALLOCATION_SCOPE,
        }
    return {
        "status": "pass",
        "stage": stage,
        "used_for_gate": False,
        "evidence": _check_stage(stage, plan, _expected_allocation_jobs(plan), True),
        "scope": OPTIONAL_ALLOCATION_SCOPE,
    }


def _numeric_summary(values: list[Any], label: str) -> dict[str, Any]:
    _require(isinstance(values, list) and values, f"{label} is empty")
    for index, value in enumerate(values):
        BASE.nonnegative_integer(value, f"{label}[{index}]")
    ordered = sorted(values)
    return {
        "samples": ordered,
        "p50": (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) / 2.0,
        "mean": sum(ordered) / len(ordered),
    }


def _compare_optional_allocations(baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    """Compare the retained commit-only allocator diagnostics.

    Planning and publication vectors are checked by the separate allocation
    analyzer; this legacy lane must not imply that it validates them.
    """

    left_state = baseline.get("allocation", {})
    right_state = candidate.get("allocation", {})
    if left_state.get("status") == right_state.get("status") == "not-captured":
        return {
            "status": "not-captured",
            "used_for_gate": False,
            "scope": OPTIONAL_ALLOCATION_SCOPE,
        }
    if left_state.get("status") != "pass" or right_state.get("status") != "pass":
        return {
            "status": "incomplete-pair",
            "used_for_gate": False,
            "baseline_status": left_state.get("status"),
            "candidate_status": right_state.get("status"),
            "scope": OPTIONAL_ALLOCATION_SCOPE,
        }
    left_rows = {BASE.job_key(row): row for row in left_state["evidence"]["allocation"]["rows"]}
    right_rows = {BASE.job_key(row): row for row in right_state["evidence"]["allocation"]["rows"]}
    _require(set(left_rows) == set(right_rows), "optional allocator baseline/candidate row keys differ")
    rows: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    for key in sorted(left_rows):
        left, right = left_rows[key], right_rows[key]
        _require(left["identity"] == right["identity"],
                 f"optional allocator logical identity differs for {key}")
        _require(len(left["allocation"]) == len(right["allocation"]),
                 f"optional allocator sample counts differ for {key}")
        metrics: dict[str, Any] = {}
        for field in BASE.ALLOCATION_REPORT_FIELDS:
            left_values = [sample[field] for sample in left["allocation"]]
            right_values = [sample[field] for sample in right["allocation"]]
            left_summary = _numeric_summary(left_values, f"optional allocator baseline {key} {field}")
            right_summary = _numeric_summary(right_values, f"optional allocator candidate {key} {field}")
            value = _comparison_record(left_summary["p50"], right_summary["p50"],
                                       f"optional allocator {key} {field} p50")
            metrics[field] = {
                **value,
                "baseline_summary": left_summary,
                "candidate_summary": right_summary,
                "delta_p50": right_summary["p50"] - left_summary["p50"],
            }
            if value["change_percent"] is not None and value["change_percent"] > ADVERSE_THRESHOLD_PERCENT:
                adverse.append({
                    "lane": "allocation",
                    "case": key[3], "shape": key[4], "repeat": key[2], "guard": key[1],
                    "metric": field, "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
                    **value,
                })
        rows.append({
            "case": key[3], "shape": key[4], "repeat": key[2],
            "identity_equal": True, "metrics": metrics,
        })
    return {
        "status": "pass",
        "used_for_gate": False,
        "rows": rows,
        "adverse_flags_over_five_percent": adverse,
        "scope": OPTIONAL_ALLOCATION_SCOPE,
    }


def _compare_stages(baseline: dict[str, Any], candidate: dict[str, Any],
                    plan: dict[str, Any]) -> dict[str, Any]:
    comparisons: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    drift: list[dict[str, Any]] = []
    rng = random.Random(BASE.BOOTSTRAP_SEED)
    _compare_timing_lane(baseline["evidence"], candidate["evidence"],
                         "native", rng, comparisons, adverse, drift)
    return {
        "timing_comparisons": comparisons,
        "adverse_flags_over_five_percent": adverse,
        "same_build_drift_over_five_percent": drift,
        "bootstrap": {
            "iterations": BASE.BOOTSTRAP_ITERATIONS,
            "seed": BASE.BOOTSTRAP_SEED,
            "scope": "matched-child within-stage resampling; candidate median / baseline median",
        },
    }


def _analysis_for_stage(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    native = _stage_or_pending(stage, plan, False)
    allocation = _optional_allocation_stage(stage, plan)
    if native["status"] == "pending":
        return {
            "status": "pending",
            "stage": stage,
            "plan_sha256": _sha(BASE.PLAN),
            "structured_gates": dict(plan["gates"]),
            "native": native,
            "allocation": allocation,
            "conditional_lanes": {
                "profile": {"status": "unmeasured",
                             "reason": "Native capture is incomplete; no planning Ir result is manufactured.",
                             "required_reduction_percent": float(plan["gates"]["planning_ir_reduction_percent"]),
                             "used_for_gate": False},
            },
        }
    evidence = native["evidence"]
    return {
        "status": "pass",
        "stage": stage,
        "plan_sha256": _sha(BASE.PLAN),
        "structured_gates": dict(plan["gates"]),
        "evidence": evidence,
        "allocation": allocation,
        "conditional_lanes": {
            "profile": {"status": "unmeasured",
                         "reason": "Native stage validated; planning Ir requires a separate bound profile.",
                         "required_reduction_percent": float(plan["gates"]["planning_ir_reduction_percent"]),
                         "used_for_gate": False},
        },
    }


def analyze(stage: str | None = None) -> dict[str, Any]:
    plan = _read_plan()
    if stage in ("baseline", "candidate"):
        result = _analysis_for_stage(stage, plan)
    else:
        baseline = _analysis_for_stage("baseline", plan)
        candidate = _analysis_for_stage("candidate", plan)
        if baseline["status"] == "pending" or candidate["status"] == "pending":
            result = {
                "status": "pending",
                "stage": "compare",
                "plan_sha256": _sha(BASE.PLAN),
                "structured_gates": dict(plan["gates"]),
                "baseline": baseline,
                "candidate": candidate,
                "admission_status": "pending",
                "allocation_diagnostics": _compare_optional_allocations(baseline, candidate),
                "conditional_lanes": {
                    "profile": {"status": "unmeasured",
                                 "reason": "Both native stages must validate before planning Ir is considered.",
                                 "required_reduction_percent": float(plan["gates"]["planning_ir_reduction_percent"]),
                                 "used_for_gate": False},
                },
            }
        else:
            native_admission = _native_admission(
                baseline["evidence"], candidate["evidence"], plan)
            allocation_diagnostics = _compare_optional_allocations(baseline, candidate)
            result = {
                "status": "pass",
                "stage": "compare",
                "plan_sha256": _sha(BASE.PLAN),
                "structured_gates": dict(plan["gates"]),
                "baseline": baseline["evidence"],
                "candidate": candidate["evidence"],
                "comparison": _compare_stages(baseline, candidate, plan),
                "native_admission": native_admission,
                "admission_status": native_admission["decision"],
                "allocation_diagnostics": allocation_diagnostics,
                "conditional_lanes": {
                    "profile": {"status": "unmeasured",
                                 "reason": (
                                     "Native gates passed; planning Ir requires a separate "
                                     "bound profile and is not inferred from elapsed time."
                                     if native_admission["passed"] else
                                     "Native gate failed; planning Ir profile is not admitted."
                                 ),
                                 "required_reduction_percent": float(plan["gates"]["planning_ir_reduction_percent"]),
                                 "used_for_gate": False},
                },
            }
    result["scope"] = "0544 native XLSX shared worksheet traversal pilot with XLSX guard rows"
    result["numerical_verifier"] = {
        "path": str(HELPER.relative_to(HERE.parent)),
        "sha256": _sha(HELPER),
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("baseline", "candidate", "compare"))
    parser.add_argument("--output", type=Path)
    parser.add_argument("output_positional", nargs="?", type=Path)
    args = parser.parse_args()
    stage = None if args.stage in (None, "compare") else args.stage
    output = args.output or args.output_positional
    if output is None:
        output = HERE / (f"analysis-{stage}.json" if stage else "comparison.json")
    try:
        result = analyze(stage)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    except BASE.EvidenceError as error:
        print(f"evidence check failed: {error}", file=sys.stderr)
        return 1
    print(f"0544 {result['stage']} evidence {result['status']}: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
