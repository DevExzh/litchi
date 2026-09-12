#!/usr/bin/env python3
"""Validate and compare the matched 0524 CFB/OLE2 measurements.

The campaign has two source-bound stages.  This module deliberately validates
the retained stage artifacts and their bindings; it does not inspect the
current checkout to decide whether a historical stage was captured correctly.
That replay responsibility belongs to the campaign verifier.  The one
exception in the receipt matrix is intentional: the baseline native repeat-2
control was run while the candidate source manifest was active, so that the
control and candidate measurements occupy the planned execution order while
the baseline binary remains the measured executable.

Native elapsed statistics and whole-child RSS are the latency evidence.
Allocator vectors are checked and compared separately; allocator-instrumented
elapsed values never enter the native timing comparison.  Callgrind Ir is
validated by ``analyze_profiles.py`` and is an independent admission gate.
"""

from __future__ import annotations

import argparse
import datetime
import importlib.util
import json
import math
from pathlib import Path
from typing import Any

import run as RUN


HERE = RUN.HERE
REPO = RUN.REPO
TARGET = RUN.TARGET
STAGE = RUN.STAGE
EXECUTION_STAGE = RUN.EXECUTION_STAGE
FOLDER = RUN.FOLDER
SCRATCH = RUN.SCRATCH
jobs = RUN.jobs
sha = RUN.sha

HELPER = HERE.parent / "change-0511" / "verify.py"
_helper_spec = importlib.util.spec_from_file_location("cfb_0511_checks_0524", HELPER)
if _helper_spec is None or _helper_spec.loader is None:
    raise ImportError(f"cannot load numerical helper: {HELPER}")
OLD = importlib.util.module_from_spec(_helper_spec)
_helper_spec.loader.exec_module(OLD)

METRICS = ("p50", "p95", "p99", "mean")
ELAPSED_SUMMARY_KEYS = (
    "min", "p50", "p95", "p99", "max", "mean", "standard_deviation",
    "confidence_interval_95",
)
ALLOCATION_FIELDS = tuple(OLD.ALLOCATOR_VECTOR_FIELDS) + (
    "incremental_region_peak_live_bytes",
)
PRIMARY_ALLOCATION_FIELDS = (
    "allocation_calls", "allocated_bytes", "incremental_region_peak_live_bytes",
)
STAGES = ("baseline", "candidate", "final")
ANALYSIS_STAGES = ("baseline", "candidate")
SHA256_LENGTH = 64


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence item."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def configure(stage: str, execution_stage: str | None = None) -> None:
    """Select the evidence stage and keep the imported driver in sync.

    ``analyze_profiles.py`` imports this module and calls the numerical
    helpers directly.  The assignments below are therefore required in
    addition to ``run.configure``: names imported from a module are otherwise
    stale after the driver's stage switch.
    """

    require(stage in STAGES, f"unknown evidence stage: {stage!r}")
    RUN.configure(stage, execution_stage)
    globals().update(
        STAGE=RUN.STAGE,
        EXECUTION_STAGE=RUN.EXECUTION_STAGE,
        FOLDER=RUN.FOLDER,
        SCRATCH=RUN.SCRATCH,
    )


def read(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def sha256(path: Path) -> str:
    return sha(path)


def is_hash(value: Any) -> bool:
    return isinstance(value, str) and len(value) == SHA256_LENGTH and all(
        character in "0123456789abcdef" for character in value
    )


def check_hash(value: Any, label: str) -> str:
    require(is_hash(value), f"{label} is not a lowercase SHA-256")
    return value


def plan_data() -> dict[str, Any]:
    plan = read(HERE / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("cpu") == 2, "plan CPU differs from the pinned CPU")
    groups = plan.get("groups")
    require(isinstance(groups, dict), "plan groups are missing")
    require(groups.get("xls", {}).get("cases") == [
        "xls_semantic_open",
        "xls_eager_open_list_worksheets",
        "xls_eager_open_one_cell",
        "xls_source_backed_open",
        "xls_source_backed_open_list_worksheets",
        "xls_source_backed_open_one_cell",
        "xls_owned_source_open",
        "xls_owned_source_open_list_worksheets",
        "xls_owned_source_open_one_cell",
    ], "plan XLS case matrix differs")
    require(groups.get("cfb", {}).get("cases") == ["cfb_open"],
            "plan CFB case matrix differs")
    require(groups.get("cfb", {}).get("shapes") == [
        "tiny", "many-small", "few-large"
    ], "plan CFB shape matrix differs")
    require(groups.get("cfb", {}).get("payload") == "incompressible",
            "plan CFB payload differs")

    native = plan.get("native")
    require(isinstance(native, dict), "plan native lane is missing")
    require(native.get("repeats") == 2 and native.get("warmup") == 20
            and native.get("samples") == 1000,
            "plan native counts differ")
    require(native.get("order") == [
        "baseline r1 xls/cfb",
        "candidate r1 xls/cfb",
        "candidate r2 cfb/xls",
        "baseline r2 cfb/xls",
    ], "plan native order differs")

    allocation = plan.get("allocation")
    require(isinstance(allocation, dict), "plan allocation lane is missing")
    require(allocation.get("repeats") == 2 and allocation.get("warmup") == 3
            and allocation.get("samples") == 30,
            "plan allocation counts differ")

    profile = plan.get("profile")
    require(isinstance(profile, dict), "plan profile lane is missing")
    require(profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 5,
            "plan profile counts differ")
    require(profile.get("jobs") == [
        "xls-owned", "cfb-tiny", "cfb-many-small", "cfb-few-large"
    ], "plan profile job matrix differs")
    require(profile.get("xls_owner") == (
        "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
    ), "plan XLS profile owner differs")
    require(profile.get("cfb_owner") == "litchi_cfb::file::OleFile<R>::open",
            "plan CFB profile owner differs")

    hardware = plan.get("hardware")
    require(isinstance(hardware, dict), "plan hardware lane is missing")
    require(hardware.get("repeats") == 2 and hardware.get("samples") == 1000
            and hardware.get("warmup") == 0,
            "plan hardware counts differ")

    review = plan.get("review")
    require(isinstance(review, dict), "plan review policy is missing")
    require(review.get("same_build_adverse_percent") == 5
            and review.get("matched_adverse_percent") == 5,
            "plan review thresholds differ")
    require(review.get("primary_cases") == [
        "xls_source_backed_open",
        "xls_source_backed_open_one_cell",
        "xls_owned_source_open",
        "xls_owned_source_open_one_cell",
    ], "plan primary case list differs")

    owned = plan.get("owned_paths")
    require(owned == [
        "/tmp/litchi-goal-0524",
        "/home/zhuhe/litchi-goal-0524-target",
    ], "plan owned paths differ")
    return plan


def stage_manifest(stage: str | None = None) -> dict[str, str]:
    selected = stage or STAGE
    require(selected in STAGES, f"unknown manifest stage: {selected!r}")
    path = HERE / selected / "source-manifest.json"
    value = read(path)
    require(isinstance(value, dict) and value, f"{selected} source manifest is empty")
    for relative, digest in value.items():
        require(isinstance(relative, str) and relative
                and not Path(relative).is_absolute(),
                f"{selected} source manifest has an invalid path")
        check_hash(digest, f"{selected} source manifest {relative}")
    return value


def _expected_execution_stage(name: str, stage: str | None = None) -> str:
    selected = stage or STAGE
    require(selected in STAGES, f"unknown receipt stage: {selected!r}")
    # This is the planned order's final control.  The baseline executable was
    # restored while the candidate source manifest remained active.
    if selected == "baseline" and name in {"native-r2-xls", "native-r2-cfb"}:
        return "candidate"
    return selected


def receipt(name: str, kind: str | None = None, allow_failure: bool = False,
            expected_execution_stage: str | None = None) -> dict[str, Any]:
    """Validate a stage-local receipt and its stage-local artifact hashes."""

    path = FOLDER / f"{name}.receipt.json"
    value = read(path)
    require(isinstance(value, dict), f"{path.name} is not an object")
    if not allow_failure:
        require(value.get("exit_code") == 0, f"{path.name} did not exit successfully")
    else:
        require(isinstance(value.get("exit_code"), int), f"{path.name} exit code is invalid")
    require(value.get("script_sha256") == sha(HERE / "run.py"),
            f"{path.name} script hash differs")
    require(value.get("plan_sha256") == sha(HERE / "plan.json"),
            f"{path.name} plan hash differs")
    require(value.get("source_manifest_sha256") == sha(FOLDER / "source-manifest.json"),
            f"{path.name} stage manifest hash differs")

    execution = value.get("execution_stage")
    require(execution in STAGES, f"{path.name} execution stage is invalid")
    expected = _expected_execution_stage(name)
    if expected_execution_stage is not None:
        expected = expected_execution_stage
    require(execution == expected,
            f"{path.name} execution stage {execution!r} differs from {expected!r}")
    execution_manifest = HERE / execution / "source-manifest.json"
    require(value.get("execution_manifest_sha256") == sha(execution_manifest),
            f"{path.name} execution manifest hash differs")

    require(isinstance(value.get("seconds"), (int, float))
            and not isinstance(value["seconds"], bool)
            and value["seconds"] > 0,
            f"{path.name} duration is invalid")
    try:
        start = datetime.datetime.fromisoformat(value["start_utc"])
        end = datetime.datetime.fromisoformat(value["end_utc"])
    except (KeyError, TypeError, ValueError) as error:
        raise EvidenceError(f"{path.name} timestamps are invalid") from error
    require(start < end, f"{path.name} timestamps are not increasing")
    environment = value.get("environment")
    require(isinstance(environment, dict), f"{path.name} environment is missing")
    require(all(item is None for item in environment.values()),
            f"{path.name} has an uncontrolled environment")

    if kind is None:
        require(value.get("binary_sha256") is None,
                f"{path.name} unexpectedly carries a binary hash")
    else:
        require(kind in {"normal", "alloc"}, f"unknown binary kind {kind!r}")
        meta = read(FOLDER / f"binary-{kind}.json")
        require(value.get("binary_sha256") == meta.get("sha256"),
                f"{path.name} binary hash differs")

    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict), f"{path.name} artifact inventory is missing")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename,
                f"{path.name} has a non-local artifact name")
        check_hash(digest, f"{path.name}/{filename}")
        artifact = FOLDER / filename
        require(artifact.is_file() and not artifact.is_symlink(),
                f"{path.name} artifact is missing: {filename}")
        require(sha(artifact) == digest, f"{path.name} artifact hash differs: {filename}")
    return value


def _binary_path_is_absent_with_custody(path: Path, plan: dict[str, Any]) -> None:
    cleanup_path = HERE / "cleanup.json"
    require(cleanup_path.is_file(), f"cleaned binary has no custody record: {path}")
    cleanup = read(cleanup_path)
    require(isinstance(cleanup, dict) and cleanup.get("owned_paths_absent") is True,
            "cleanup record does not prove owned paths are absent")
    removed = cleanup.get("removed", cleanup.get("removed_paths"))
    require(isinstance(removed, list) and str(SCRATCH.parent) in removed,
            "cleanup record does not name the 0524 scratch root")
    require(all(item in removed for item in plan["owned_paths"]),
            "cleanup record does not name every planned owned path")


def binary(kind: str) -> dict[str, Any]:
    """Validate one stage binary descriptor, build receipt, and custody."""

    require(kind in {"normal", "alloc"}, f"unknown binary kind {kind!r}")
    plan = plan_data()
    meta = read(FOLDER / f"binary-{kind}.json")
    require(isinstance(meta, dict), f"binary-{kind}.json is not an object")
    check_hash(meta.get("sha256"), f"binary-{kind}.json.sha256")
    require(isinstance(meta.get("bytes"), int) and meta["bytes"] > 0,
            f"binary-{kind}.json bytes are invalid")
    build_path = FOLDER / f"build-{kind}.receipt.json"
    require(meta.get("build_receipt_sha256") == sha(build_path),
            f"binary-{kind} build receipt hash differs")
    require(meta.get("source_manifest_sha256") == sha(FOLDER / "source-manifest.json"),
            f"binary-{kind} source manifest hash differs")
    build = receipt(f"build-{kind}")
    require(build.get("binary_sha256") is None,
            f"build-{kind} receipt unexpectedly has a binary hash")
    expected = [
        "env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
        "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
        "--bin", "litchi-perf-baseline-alloc" if kind == "alloc" else "litchi-perf-baseline",
        "--target-dir", str(TARGET),
    ]
    if kind == "alloc":
        expected += ["--features", "allocator-metrics"]
    require(build.get("command") == expected, f"build-{kind} command differs")

    path = Path(meta.get("path", ""))
    require(path == SCRATCH / kind, f"binary-{kind} path differs from the stage")
    if path.exists():
        require(path.is_file() and not path.is_symlink(),
                f"binary-{kind} is not a regular file")
        require(sha(path) == meta["sha256"] and path.stat().st_size == meta["bytes"],
                f"binary-{kind} descriptor does not match the executable")
    else:
        _binary_path_is_absent_with_custody(path, plan)
    return {
        "path": str(path),
        "sha256": meta["sha256"],
        "bytes": meta["bytes"],
        "build_receipt_sha256": meta["build_receipt_sha256"],
        "source_manifest_sha256": meta["source_manifest_sha256"],
    }


def normalized(value: Any) -> Any:
    if isinstance(value, list):
        require(value, "identity vector is empty")
        values = [normalized(item) for item in value]
        require(all(item == values[0] for item in values),
                "identity vector varies across samples")
        return values[0]
    if isinstance(value, dict):
        return {key: normalized(item) for key, item in sorted(value.items())}
    return value


def identity(row: dict[str, Any]) -> dict[str, Any]:
    return {
        "case": row["case"],
        "corpus": row["corpus"],
        "sink": row.get("sink"),
        "output_sha256": row.get("output_sha256"),
        "source": normalized(row.get("source"))
        if row.get("source") is not None else None,
    }


def validate_row(row: dict[str, Any], samples: int, allocator: bool) -> dict[str, Any]:
    context = row["case"] + "/" + row["corpus"]["shape"]
    elapsed = OLD._validate_elapsed(row, samples, context)
    if row["case"] != "cfb_open":
        OLD.verify_xls_row(row, samples, context, allocator)
    else:
        require(row.get("source") is None and row.get("sink") is None,
                f"{context} CFB row unexpectedly publishes source/sink evidence")
        require(row.get("output_sha256") is None,
                f"{context} CFB row unexpectedly publishes an output hash")
        OLD.verify_operation_metrics(
            row, samples, context, allocator, elapsed["sample_order"]
        )
    operation = row.get("operation_metrics")
    require(isinstance(operation, dict), f"{context} operation metrics are missing")
    allocation = operation.get("allocation")
    require(isinstance(allocation, dict), f"{context} allocation envelope is missing")
    if allocator:
        require(allocation.get("status") == "measured",
                f"{context} allocator metrics are not measured")
        require(allocation.get("scope") == "operation_global_system_allocator",
                f"{context} allocator scope differs")
        for index in range(samples):
            vectors = {
                field: allocation[field]["values"][index]
                for field in OLD.ALLOCATOR_VECTOR_FIELDS
            }
            require(vectors["failed_allocation_calls"] == 0,
                    f"{context} recorded a failed allocation")
            require(
                vectors["live_bytes_after"]
                == vectors["live_bytes_before"] + vectors["allocated_bytes"]
                - vectors["deallocated_bytes"],
                f"{context} live-byte balance does not reconcile",
            )
            require(
                vectors["region_peak_live_bytes"]
                >= max(vectors["live_bytes_before"], vectors["live_bytes_after"]),
                f"{context} region peak is below live bytes",
            )
    else:
        require(allocation.get("status") == "unavailable",
                f"{context} normal allocation metrics are not explicitly unavailable")
        require(allocation.get("scope") == "operation_global_system_allocator",
                f"{context} normal allocation scope differs")
        require("values" not in allocation,
                f"{context} normal allocation envelope contains values")
    return row


def validate_report(path: Path, job: dict[str, Any], kind: str = "normal") -> list[dict[str, Any]]:
    """Validate a report while the caller-selected stage is active."""

    report = read(path)
    meta = read(FOLDER / f"binary-{kind}.json")
    OLD.verify_report_identity(
        report, {"binary_sha256": meta["sha256"]}, job["samples"],
        job["warmup"], kind == "alloc", path.name,
    )
    environment = report.get("environment")
    require(environment.get("git_revision") == read(HERE / "plan.json")["revision"],
            f"{path.name} revision differs from the plan")
    require(environment.get("cpu_affinity") == str(read(HERE / "plan.json")["cpu"]),
            f"{path.name} CPU affinity differs from the plan")
    binary_identity = report.get("binary_identity")
    require(binary_identity.get("path") == meta["path"],
            f"{path.name} binary path differs")
    require(binary_identity.get("binary_bytes") == meta["bytes"],
            f"{path.name} binary size differs")

    selection = job["selection"]
    configuration = report["configuration"]
    require(configuration.get("cases") == selection["cases"],
            f"{path.name} case selection differs")
    if "shapes" in selection:
        require(configuration.get("corpus_shapes") == selection["shapes"],
                f"{path.name} CFB shape selection differs")
        require(configuration.get("payload_kinds") == [selection["payload"]],
                f"{path.name} payload selection differs")

    rows = report.get("results")
    require(isinstance(rows, list), f"{path.name} results are not a list")
    if "shapes" in selection:
        expected = {
            (case, shape)
            for case in selection["cases"]
            for shape in selection["shapes"]
        }
        actual = {(row.get("case"), row.get("corpus", {}).get("shape"))
                  for row in rows}
    else:
        expected = {(case, "xls-source-backed") for case in selection["cases"]}
        actual = {(row.get("case"), "xls-source-backed") for row in rows}
    require(len(rows) == len(actual) and actual == expected,
            f"{path.name} result matrix differs")

    catalog_path = path.with_name(path.name.replace(".json", ".catalog.json"))
    catalog = read(catalog_path)
    OLD.validate_binding(report, catalog)
    return [validate_row(row, job["samples"], kind == "alloc") for row in rows]


def _elapsed_summary(elapsed: dict[str, Any]) -> dict[str, Any]:
    return {key: elapsed[key] for key in ELAPSED_SUMMARY_KEYS if key in elapsed}


def _validate_rss(path: Path) -> dict[str, Any]:
    value = read(path)
    require(isinstance(value, dict), f"{path.name} RSS is not an object")
    require(set(value) == {
        "max_rss_kib", "elapsed_seconds", "user_seconds", "system_seconds"
    }, f"{path.name} RSS fields differ")
    require(isinstance(value["max_rss_kib"], int)
            and not isinstance(value["max_rss_kib"], bool)
            and value["max_rss_kib"] >= 0,
            f"{path.name} max RSS is invalid")
    for key in ("elapsed_seconds", "user_seconds", "system_seconds"):
        require(isinstance(value[key], (int, float))
                and not isinstance(value[key], bool)
                and math.isfinite(float(value[key]))
                and value[key] >= 0,
                f"{path.name} {key} is invalid")
    return value


def _command_option(command: list[str], option: str) -> str:
    positions = [index for index, item in enumerate(command) if item == option]
    require(len(positions) == 1 and positions[0] + 1 < len(command),
            f"command option {option} is missing or repeated")
    return command[positions[0] + 1]


def _validate_capture_command(command: list[str], job: dict[str, Any],
                              kind: str, lane: str) -> None:
    plan = plan_data()
    require(command[:3] == ["taskset", "-c", str(plan["cpu"])],
            f"{job['name']} CPU command differs")
    require(str(SCRATCH / kind) in command,
            f"{job['name']} binary path is missing")
    require(_command_option(command, "--case") == ",".join(job["selection"]["cases"]),
            f"{job['name']} case option differs")
    require(_command_option(command, "--samples") == str(job["samples"]),
            f"{job['name']} sample option differs")
    require(_command_option(command, "--warmup") == str(job["warmup"]),
            f"{job['name']} warmup option differs")
    require(_command_option(command, "--json") == str(FOLDER / f"{job['name']}.json"),
            f"{job['name']} report path differs")
    require(_command_option(command, "--corpus-manifest")
            == str(FOLDER / f"{job['name']}.catalog.json"),
            f"{job['name']} catalog path differs")
    if "shapes" in job["selection"]:
        require(_command_option(command, "--shape")
                == ",".join(job["selection"]["shapes"]),
                f"{job['name']} shape option differs")
        require(_command_option(command, "--payload") == job["selection"]["payload"],
                f"{job['name']} payload option differs")
    else:
        require("--shape" not in command and "--payload" not in command,
                f"{job['name']} XLS command unexpectedly has CFB selectors")
    if lane == "native":
        require("/usr/bin/time" in command,
                f"{job['name']} native command lacks RSS observer")
        require(_command_option(command, "-o") == str(FOLDER / f"{job['name']}.rss.json"),
                f"{job['name']} RSS path differs")
    else:
        require("/usr/bin/time" not in command and "valgrind" not in command,
                f"{job['name']} allocator command has another observer")


def _row_item(row: dict[str, Any], job: dict[str, Any], lane: str,
              report_name: str) -> dict[str, Any]:
    ident = identity(row)
    item = {
        "lane": lane,
        "repeat": job["repeat"],
        "case": row["case"],
        "shape": row["corpus"]["shape"],
        "report": report_name,
        "samples": job["samples"],
        "identity": ident,
        "identity_equal": True,
    }
    if lane == "native":
        elapsed = row["elapsed_ns"]
        item["elapsed_ns"] = {metric: elapsed[metric] for metric in METRICS}
        item["timing_stats"] = _elapsed_summary(elapsed)
        item["operations_per_second"] = 1e9 / elapsed["mean"]
        item["standard_deviation_ns"] = elapsed["standard_deviation"]
        item["confidence_interval_95"] = elapsed["confidence_interval_95"]
    else:
        allocation = row["operation_metrics"]["allocation"]
        require(allocation.get("status") == "measured",
                f"{report_name} allocation report is not measured")
        vectors = {
            field: list(allocation[field]["values"])
            for field in OLD.ALLOCATOR_VECTOR_FIELDS
        }
        vectors["incremental_region_peak_live_bytes"] = [
            peak - before
            for peak, before in zip(
                vectors["region_peak_live_bytes"], vectors["live_bytes_before"]
            )
        ]
        item["allocation"] = vectors
        item["allocation_status"] = allocation["status"]
        item["allocation_scope"] = allocation["scope"]
    return item


def _variation(base: Any, current: Any) -> dict[str, Any]:
    require(isinstance(base, (int, float)) and not isinstance(base, bool),
            "variation baseline is not numeric")
    require(isinstance(current, (int, float)) and not isinstance(current, bool),
            "variation current value is not numeric")
    base_float = float(base)
    current_float = float(current)
    require(math.isfinite(base_float) and math.isfinite(current_float),
            "variation value is not finite")
    delta = current_float - base_float
    if base_float == 0:
        change = 0.0 if current_float == 0 else None
    else:
        change = delta / base_float * 100.0
    return {
        "first": base,
        "second": current,
        "delta": delta,
        "change_percent": change,
        "over_five_percent": change is None or abs(change) > 5.0,
    }


def _numeric_leaves(value: Any, prefix: str = "") -> list[tuple[str, float | int]]:
    if isinstance(value, dict):
        leaves: list[tuple[str, float | int]] = []
        for key in sorted(value):
            if key == "method":
                continue
            child = f"{prefix}.{key}" if prefix else key
            leaves.extend(_numeric_leaves(value[key], child))
        return leaves
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return [(prefix, value)]
    return []


def _stage_variations(rows: list[dict[str, Any]], rss: list[dict[str, Any]],
                      threshold: float) -> list[dict[str, Any]]:
    variations: list[dict[str, Any]] = []
    native = {
        (item["case"], item["shape"], item["repeat"]): item
        for item in rows if item["lane"] == "native"
    }
    cases = sorted({(item["case"], item["shape"]) for item in rows
                    if item["lane"] == "native"})
    for case, shape in cases:
        first = native.get((case, shape, 1))
        second = native.get((case, shape, 2))
        if first is None or second is None:
            continue
        for stat, left in _numeric_leaves(first["timing_stats"]):
            right = dict(_numeric_leaves(second["timing_stats"])).get(stat)
            if right is None:
                continue
            record = _variation(left, right)
            record["over_threshold"] = record["change_percent"] is None or abs(record["change_percent"]) > threshold
            if record["over_threshold"]:
                variations.append({
                    "lane": "native", "case": case, "shape": shape,
                    "repeat_first": 1, "repeat_second": 2,
                    "metric": "elapsed_ns." + stat, **record,
                })
        for rss_key in ("max_rss_kib", "elapsed_seconds", "user_seconds", "system_seconds"):
            record = _variation(first["rss"][rss_key], second["rss"][rss_key])
            record["over_threshold"] = record["change_percent"] is None or abs(record["change_percent"]) > threshold
            if record["over_threshold"]:
                variations.append({
                    "lane": "native", "case": case, "shape": shape,
                    "repeat_first": 1, "repeat_second": 2,
                    "metric": "rss." + rss_key, **record,
                })

    allocator = {
        (item["case"], item["shape"], item["repeat"]): item
        for item in rows if item["lane"] == "alloc"
    }
    for case, shape in sorted({(key[0], key[1]) for key in allocator}):
        first = allocator.get((case, shape, 1))
        second = allocator.get((case, shape, 2))
        if first is None or second is None:
            continue
        for field in ALLOCATION_FIELDS:
            left_values = first["allocation"][field]
            right_values = second["allocation"][field]
            require(len(left_values) == len(right_values),
                    f"allocation repeat vectors differ for {case}/{shape}/{field}")
            # Allocation captures are vectors rather than a timing statistic;
            # report the aggregate and retain its direction for review.
            left = max(left_values) if field == "incremental_region_peak_live_bytes" else sum(left_values)
            right = max(right_values) if field == "incremental_region_peak_live_bytes" else sum(right_values)
            record = _variation(left, right)
            record["over_threshold"] = record["change_percent"] is None or abs(record["change_percent"]) > threshold
            if record["over_threshold"]:
                variations.append({
                    "lane": "allocation", "case": case, "shape": shape,
                    "repeat_first": 1, "repeat_second": 2,
                    "metric": "allocation." + field, **record,
                })
    return variations


def analyze(stage: str | None = None) -> dict[str, Any]:
    """Validate one complete stage and return deterministic numeric evidence."""

    selected = stage or STAGE
    require(selected in ANALYSIS_STAGES,
            f"numeric analysis is only defined for baseline/candidate stages: {selected!r}")
    configure(selected)
    plan = plan_data()
    manifest = stage_manifest(selected)
    metadata = {kind: binary(kind) for kind in ("normal", "alloc")}
    rows: list[dict[str, Any]] = []
    identities: dict[tuple[str, str], dict[str, Any]] = {}
    rss: list[dict[str, Any]] = []
    all_receipts: list[dict[str, Any]] = []

    for lane in ("native", "alloc"):
        kind = "alloc" if lane == "alloc" else "normal"
        for job in jobs(lane):
            current_receipt = receipt(job["name"], kind)
            all_receipts.append(current_receipt)
            _validate_capture_command(current_receipt["command"], job, kind, lane)
            report_path = FOLDER / f"{job['name']}.json"
            captured = validate_report(report_path, job, kind)
            if lane == "native":
                rss_value = _validate_rss(FOLDER / f"{job['name']}.rss.json")
            for row in captured:
                item = _row_item(row, job, lane, report_path.name)
                if lane == "native":
                    item["rss"] = dict(rss_value)
                    rss.append({"name": job["name"], **rss_value})
                key = (item["case"], item["shape"])
                current_identity = item["identity"]
                if key in identities:
                    require(identities[key] == current_identity,
                            f"{selected} logical identity changed for {key}")
                else:
                    identities[key] = current_identity
                rows.append(item)

    require(len([item for item in rows if item["lane"] == "native"]) == 24,
            f"{selected} native row matrix is incomplete")
    require(len([item for item in rows if item["lane"] == "alloc"]) == 24,
            f"{selected} allocation row matrix is incomplete")
    require(sum(item["samples"] for item in rows if item["lane"] == "native") == 24000,
            f"{selected} native sample count is not 24000")
    require(sum(item["samples"] for item in rows if item["lane"] == "alloc") == 720,
            f"{selected} allocation sample count is not 720")

    native_by_key = {
        (item["case"], item["shape"], item["repeat"]): item
        for item in rows if item["lane"] == "native"
    }
    for item in (item for item in rows if item["lane"] == "alloc"):
        counterpart = native_by_key.get((item["case"], item["shape"], item["repeat"]))
        require(counterpart is not None,
                f"{selected} allocation row has no native counterpart")
        require(counterpart["identity"] == item["identity"],
                f"{selected} allocator/native identity differs for {item['case']}/{item['shape']}")

    ordered = sorted(all_receipts, key=lambda item: item["start_utc"])
    require(all(left["end_utc"] <= right["start_utc"]
                for left, right in zip(ordered, ordered[1:])),
            f"{selected} stage receipts overlap")
    same_build = _stage_variations(rows, rss, plan["review"]["same_build_adverse_percent"])
    same_build = [dict(stage=selected, **item) for item in same_build]
    return {
        "schema": "cfb_ole2_matched_numeric_analysis_v1",
        "status": "pass",
        "stage": selected,
        "scope": plan["scope"],
        "priority": plan["priority"],
        "plan_sha256": sha(HERE / "plan.json"),
        "source_manifest_sha256": sha(HERE / selected / "source-manifest.json"),
        "execution_manifest_bindings": sorted({
            item["execution_stage"] for item in all_receipts
        }),
        "metadata": metadata,
        "native_samples": 24000,
        "allocation_samples": 720,
        "rows": rows,
        "identities": list(identities.values()),
        "whole_child_rss": rss,
        "same_build_variations_over_five_percent": same_build,
        "helpers": {
            "change-0511/verify.py": sha(HELPER),
            "tools/summarize_crud_baseline.py": sha(REPO / "tools/summarize_crud_baseline.py"),
            "tools/validate_perf_corpus_binding.py": sha(
                REPO / "tools/validate_perf_corpus_binding.py"
            ),
            "change-0524/run.py": sha(HERE / "run.py"),
        },
        "limits": (
            "Native timing and /usr/bin/time whole-child RSS are matched evidence. "
            "Allocator vectors are separate instrumentation and their elapsed values "
            "are excluded from timing claims. No physical-I/O, cold-cache, stable-tail, "
            "provider, native-producer, or scaling claim; constructor Ir is gated by "
            "the independent profile analysis. ODF remains deferred."
        ),
    }


def _percent_change(baseline: float, candidate: float) -> float | None:
    if baseline == 0:
        return 0.0 if candidate == 0 else None
    return (candidate / baseline - 1.0) * 100.0


def _delta(baseline: Any, candidate: Any) -> Any:
    if isinstance(baseline, dict) and isinstance(candidate, dict):
        require(set(baseline) == set(candidate), "matched nested statistic keys differ")
        return {
            key: (baseline[key] if key == "method"
                  else _delta(baseline[key], candidate[key]))
            for key in sorted(baseline)
        }
    if isinstance(baseline, str) and isinstance(candidate, str):
        require(baseline == candidate, "matched statistic labels differ")
        return baseline
    require(isinstance(baseline, (int, float)) and not isinstance(baseline, bool),
            "matched statistic baseline is not numeric")
    require(isinstance(candidate, (int, float)) and not isinstance(candidate, bool),
            "matched statistic candidate is not numeric")
    baseline_float = float(baseline)
    candidate_float = float(candidate)
    require(math.isfinite(baseline_float) and math.isfinite(candidate_float),
            "matched statistic is not finite")
    change = _percent_change(baseline_float, candidate_float)
    return {
        "baseline": baseline,
        "candidate": candidate,
        "delta": candidate_float - baseline_float,
        "change_percent": change,
        "improvement_percent": -change if change is not None else None,
    }


def _allocation_delta(left: list[int], right: list[int]) -> dict[str, Any]:
    require(len(left) == len(right) and left, "matched allocation vectors differ")
    deltas = [candidate - baseline for baseline, candidate in zip(left, right)]
    if len(left) == 1:
        baseline_mean, candidate_mean = float(left[0]), float(right[0])
    else:
        baseline_mean = sum(left) / len(left)
        candidate_mean = sum(right) / len(right)
    baseline_total, candidate_total = sum(left), sum(right)
    total_change = _percent_change(float(baseline_total), float(candidate_total))
    return {
        "baseline": left,
        "candidate": right,
        "delta": deltas,
        "baseline_total": baseline_total,
        "candidate_total": candidate_total,
        "delta_total": candidate_total - baseline_total,
        "total_change_percent": total_change,
        "baseline_mean": baseline_mean,
        "candidate_mean": candidate_mean,
        "delta_mean": candidate_mean - baseline_mean,
        "baseline_max": max(left),
        "candidate_max": max(right),
        "delta_max": max(right) - max(left),
        "max_change_percent": _percent_change(float(max(left)), float(max(right))),
    }


def _matched_flag(lane: str, case: str, shape: str, repeat: int, metric: str,
                  value: dict[str, Any], threshold: float) -> dict[str, Any]:
    change = value.get("change_percent")
    adverse = change is None or change > threshold
    return {
        "lane": lane,
        "case": case,
        "shape": shape,
        "repeat": repeat,
        "metric": metric,
        "threshold_percent": threshold,
        "adverse": adverse,
        **value,
    }


def _delta_records(value: Any, prefix: str = "") -> list[tuple[str, dict[str, Any]]]:
    """Return scalar comparison records, including nested CI leaves."""

    if isinstance(value, dict):
        if {"baseline", "candidate", "change_percent"}.issubset(value):
            return [(prefix, value)]
        records: list[tuple[str, dict[str, Any]]] = []
        for key in sorted(value):
            if key == "method":
                continue
            child = f"{prefix}.{key}" if prefix else key
            records.extend(_delta_records(value[key], child))
        return records
    return []


def compare_stages(baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    """Compare two already-validated stage documents."""

    require(baseline.get("stage") == "baseline", "baseline document has wrong stage")
    require(candidate.get("stage") == "candidate", "candidate document has wrong stage")
    baseline_rows = {
        (item["lane"], item["repeat"], item["case"], item["shape"]): item
        for item in baseline["rows"]
    }
    candidate_rows = {
        (item["lane"], item["repeat"], item["case"], item["shape"]): item
        for item in candidate["rows"]
    }
    require(set(baseline_rows) == set(candidate_rows),
            "baseline/candidate row keys differ")
    review = plan_data()["review"]
    threshold = review["matched_adverse_percent"]
    timing_comparisons: list[dict[str, Any]] = []
    adverse_flags: list[dict[str, Any]] = []
    allocation_comparisons: list[dict[str, Any]] = []

    for key in sorted(baseline_rows):
        left, right = baseline_rows[key], candidate_rows[key]
        require(left["identity"] == right["identity"],
                f"baseline/candidate logical identity differs for {key}")
        lane, repeat, case, shape = key
        if lane == "native":
            stats: dict[str, Any] = {}
            for metric in sorted(left["timing_stats"]):
                value = _delta(left["timing_stats"][metric], right["timing_stats"][metric])
                stats[metric] = value
                for leaf, leaf_value in _delta_records(value):
                    flag = _matched_flag(
                        "native", case, shape, repeat,
                        "elapsed_ns." + metric + ("." + leaf if leaf else ""),
                        leaf_value, threshold,
                    )
                    if flag["adverse"]:
                        adverse_flags.append(flag)
            rss_stats: dict[str, Any] = {}
            for metric in sorted(left["rss"]):
                value = _delta(left["rss"][metric], right["rss"][metric])
                rss_stats[metric] = value
                flag = _matched_flag(
                    "native", case, shape, repeat, "rss." + metric, value, threshold
                )
                if flag["adverse"]:
                    adverse_flags.append(flag)
            timing_comparisons.append({
                "lane": lane,
                "repeat": repeat,
                "case": case,
                "shape": shape,
                "identity_equal": True,
                "timing_stats": stats,
                # Keep the phase-style name used by earlier campaign
                # reports as a compatibility view; there is only one native
                # phase in this CFB experiment.
                "metrics": {"elapsed_ns": stats},
                "rss": rss_stats,
            })
        else:
            metrics = {
                field: _allocation_delta(left["allocation"][field], right["allocation"][field])
                for field in ALLOCATION_FIELDS
            }
            require(left["allocation_status"] == right["allocation_status"]
                    and left["allocation_scope"] == right["allocation_scope"],
                    f"matched allocation envelope differs for {key}")
            allocation_comparisons.append({
                "lane": lane,
                "repeat": repeat,
                "case": case,
                "shape": shape,
                "identity_equal": True,
                "status": left["allocation_status"],
                "scope": left["allocation_scope"],
                "metrics": metrics,
            })

    same_build = list(baseline.get("same_build_variations_over_five_percent", []))
    same_build.extend(candidate.get("same_build_variations_over_five_percent", []))

    primary: list[dict[str, Any]] = []
    native_primary = {
        (item["case"], item["repeat"]): item
        for item in baseline_rows.values()
        if item["lane"] == "native"
    }
    candidate_primary = {
        (item["case"], item["repeat"]): item
        for item in candidate_rows.values()
        if item["lane"] == "native"
    }
    for case in review["primary_cases"]:
        for repeat in (1, 2):
            left = native_primary.get((case, repeat))
            right = candidate_primary.get((case, repeat))
            require(left is not None and right is not None,
                    f"primary case is missing from matched native rows: {case}/{repeat}")
            value = _delta(left["timing_stats"]["p50"], right["timing_stats"]["p50"])
            improvement = value["improvement_percent"]
            primary.append({
                "case": case,
                "repeat": repeat,
                "shape": left["shape"],
                "p50": value,
                "required_improvement_percent": 3.0,
                "passes": improvement is not None and improvement >= 3.0,
            })

    allocation_guard: list[dict[str, Any]] = []
    for record in allocation_comparisons:
        for field in PRIMARY_ALLOCATION_FIELDS:
            metric = record["metrics"][field]
            change = (metric["max_change_percent"]
                      if field == "incremental_region_peak_live_bytes"
                      else metric["total_change_percent"])
            adverse = change is None or change > threshold
            allocation_guard.append({
                "case": record["case"], "shape": record["shape"],
                "repeat": record["repeat"], "metric": field,
                "threshold_percent": threshold,
                "change_basis": "max" if field == "incremental_region_peak_live_bytes" else "total",
                "change_percent": change,
                "adverse": adverse,
                **metric,
            })

    primary_pass = bool(primary) and all(item["passes"] for item in primary)
    allocation_pass = not any(item["adverse"] for item in allocation_guard)
    return {
        "status": "pass",
        "identity_equal": True,
        "timing_comparisons": timing_comparisons,
        "allocation_comparisons": allocation_comparisons,
        "matched_adverse_flags_over_five_percent": adverse_flags,
        "adverse_flags_over_five_percent": adverse_flags,
        "same_build_variations_over_five_percent": same_build,
        "same_build_drift_over_five_percent": same_build,
        "admission": {
            "primary_workflow_p50": {
                "required_improvement_percent": 3.0,
                "paired_repeats": primary,
                "all_four_cases_both_repeats_pass": primary_pass,
            },
            "allocation_guard": {
                "required": "calls, allocated bytes, and incremental region peak do not grow materially",
                "material_growth_threshold_percent": threshold,
                "metrics": allocation_guard,
                "passes": allocation_pass,
            },
            "constructor_ir": {
                "required": "selected constructor Ir decreases in both repeats",
                "status": "separate_root_profile_gate",
                "passed": None,
            },
            "quality": {
                "required": "quality gates and correctness checks pass",
                "status": "separate_root_gate",
                "passed": None,
            },
            "eligible_before_external_gates": primary_pass and allocation_pass,
        },
        "policy": review["policy"],
        "limits": (
            "Matched deltas are descriptive evidence for this fixed workload. "
            "A production decision also requires independent constructor-Ir and "
            "quality gates; no broad provider, producer, scaling, or ODF claim."
        ),
    }


def compare() -> dict[str, Any]:
    baseline = analyze("baseline")
    candidate = analyze("candidate")
    return {
        "schema": "cfb_ole2_matched_comparison_v1",
        "status": "pass",
        "stage": "compare",
        "scope": plan_data()["scope"],
        "priority": plan_data()["priority"],
        "plan_sha256": sha(HERE / "plan.json"),
        "baseline": baseline,
        "candidate": candidate,
        "comparison": compare_stages(baseline, candidate),
        "helpers": baseline["helpers"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", nargs="?", type=Path,
                        help="JSON destination (single stage or comparison)")
    parser.add_argument("--stage", choices=STAGES, default="baseline",
                        help="validate one stage (default: baseline)")
    parser.add_argument("--compare", action="store_true",
                        help="validate and compare both stages")
    parser.add_argument("--output", dest="output_option", type=Path,
                        help="JSON destination (alternative to positional path)")
    args = parser.parse_args()
    if args.output is not None and args.output_option is not None:
        parser.error("provide output either positionally or with --output")
    if args.compare:
        output = args.output_option or args.output or HERE / "comparison.json"
    else:
        output = args.output_option or args.output or HERE / args.stage / "analysis.json"
    try:
        document = compare() if args.compare else analyze(args.stage)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")
    except (EvidenceError, AssertionError, KeyError, OSError, TypeError, ValueError) as error:
        print(f"analyze.py: evidence check failed: {error}")
        return 2
    print(f"0524 {document['stage']} evidence verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
