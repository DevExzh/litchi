#!/usr/bin/env python3
"""Validate and compare the 0521 XLSX cell-values evidence.

This file intentionally knows about the report envelope and the 0521 plan,
but not about implementation internals.  It is safe to run after either stage
has been captured::

    python3 analyze.py --stage baseline
    python3 analyze.py                 # compare once both stages exist

The normal native lane owns latency and whole-child RSS.  The allocator lane
is reported separately; its phase timings never enter the native comparison.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import random
import sys
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
PLAN = HERE / "plan.json"
PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")
TIMING_PHASES = PHASES + ("reopen_ns",)
TIMING_STATS = ("p50", "p95", "p99", "mean")
ALLOCATION_FIELDS = (
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
ALLOCATION_REPORT_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "incremental_region_peak_live_bytes",
)
BOOTSTRAP_ITERATIONS = 2000
BOOTSTRAP_SEED = 5_210_521
THRESHOLD_PERCENT = 5.0


class EvidenceError(ValueError):
    """A missing, malformed, or internally inconsistent evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha(path: Path) -> str:
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


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool),
            f"{label} is not numeric")
    require(math.isfinite(float(value)), f"{label} is not finite")


def nonnegative_integer(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")


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
    z2, z3, z5, z7 = z * z, z * z * z, z * z * z * z * z, z * z * z * z * z * z * z
    d = float(degrees)
    return (z + (z3 + z) / (4.0 * d)
            + (5.0 * z5 + 16.0 * z3 + 3.0 * z) / (96.0 * d * d)
            + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * z)
            / (384.0 * d * d * d))


def report_stats(values: list[int]) -> dict[str, Any]:
    require(values, "empty statistics vector")
    for index, value in enumerate(values):
        nonnegative_integer(value, f"sample[{index}]")
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
    standard_deviation = math.sqrt(squared_deviation_sum / (len(values) - 1)) if len(values) > 1 else 0.0
    margin = (student_t_critical_95(len(values) - 1) * standard_deviation
              / math.sqrt(len(values))) if len(values) > 1 else 0.0
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


def verify_report_stats(actual: dict[str, Any], values: list[int], label: str) -> dict[str, Any]:
    expected = report_stats(values)
    require(actual.get("unit") == "ns", f"{label}.unit is not ns")
    require(actual.get("samples") == expected["samples"], f"{label}.samples are not sorted")
    for key in ("min", "p50", "p95", "p99", "max", "mean", "standard_deviation"):
        require(key in actual, f"{label} is missing {key}")
        same_number(actual[key], expected[key], f"{label}.{key}")
    interval = actual.get("confidence_interval_95")
    require(isinstance(interval, dict), f"{label}.confidence_interval_95 is not an object")
    require(interval.get("method") == expected["confidence_interval_95"]["method"],
            f"{label} has a different confidence-interval method")
    for key in ("lower", "upper"):
        same_number(interval.get(key), expected["confidence_interval_95"][key],
                    f"{label}.confidence_interval_95.{key}")
    return expected


def expected_jobs(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    primary = plan["primary"]
    jobs: list[dict[str, Any]] = []
    if lane == "native":
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
            for guard, item in enumerate(plan["guards"]):
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
    config = plan["allocation"]
    for repeat in range(1, int(config["repeats"]) + 1):
        for shape in config["shapes"]:
            jobs.append({
                "name": f"alloc-r{repeat}-{shape}",
                "kind": "allocation",
                "guard": None,
                "repeat": repeat,
                "case": primary["case"],
                "shape": shape,
                "warmup": int(config["warmup"]),
                "samples": int(config["samples"]),
            })
    return jobs


def job_key(job: dict[str, Any]) -> tuple[Any, ...]:
    return (job["kind"], job["guard"], job["repeat"], job["case"], job["shape"])


def check_vector(value: Any, count: int, label: str) -> list[Any]:
    require(isinstance(value, list), f"{label} is not a vector")
    require(len(value) == count, f"{label} has {len(value)} values, expected {count}")
    return value


def canonical_value(value: Any, count: int, label: str) -> Any:
    if isinstance(value, list):
        values = check_vector(value, count, label)
        for index, item in enumerate(values):
            canonical_value(item, count, f"{label}[{index}]")
        require(all(item == values[0] for item in values),
                f"{label} is not constant across retained samples")
        return canonical_value(values[0], count, f"{label}[0]")
    if isinstance(value, dict):
        return {key: canonical_value(item, count, f"{label}.{key}")
                for key, item in sorted(value.items())}
    return value


def source_identity(source: dict[str, Any], count: int) -> dict[str, Any]:
    require(isinstance(source, dict), "result.source is not an object")
    result: dict[str, Any] = {}
    for key, value in sorted(source.items()):
        if key != "xlsx_cell_values":
            result[key] = canonical_value(value, count, f"source.{key}")
            continue
        require(isinstance(value, dict), "source.xlsx_cell_values is not an object")
        inner: dict[str, Any] = {}
        for inner_key, inner_value in sorted(value.items()):
            if inner_key in TIMING_PHASES:
                check_vector(inner_value, count, f"source.xlsx_cell_values.{inner_key}")
            elif inner_key == "commit_allocation_metrics":
                check_vector(inner_value, count, "source.xlsx_cell_values.commit_allocation_metrics")
            else:
                inner[inner_key] = canonical_value(
                    inner_value, count, f"source.xlsx_cell_values.{inner_key}")
        result[key] = inner
    return result


def identity_digest(identity: dict[str, Any]) -> str:
    encoded = json.dumps(identity, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(encoded).hexdigest()


def normalize_iteration_counts(identity: dict[str, Any]) -> dict[str, Any]:
    """Normalize only the independently planned native/allocator counts."""

    normalized = dict(identity)
    configuration = dict(identity["configuration"])
    for field in ("samples_per_case", "warmup_iterations_per_case"):
        configuration[field] = "<validated-planned-count>"
    normalized["configuration"] = configuration
    return normalized


def allocation_sample(sample: Any, label: str, measured: bool) -> dict[str, Any]:
    require(isinstance(sample, dict), f"{label} is not an object")
    require(sample.get("scope") == "operation_global_system_allocator",
            f"{label}.scope is unexpected")
    status = sample.get("status")
    if not measured:
        require(status == "unavailable", f"{label} must be unavailable in the normal binary")
        require(all(field not in sample for field in ALLOCATION_FIELDS),
                f"{label} contains numeric fields despite unavailable status")
        return {"status": status, "scope": sample["scope"]}
    require(status == "measured", f"{label}.status is not measured")
    values: dict[str, int] = {}
    for field in ALLOCATION_FIELDS:
        value = sample.get(field)
        nonnegative_integer(value, f"{label}.{field}")
        values[field] = value
    require(values["failed_allocation_calls"] == 0,
            f"{label} recorded a failed allocation")
    require(values["live_bytes_before"] + values["allocated_bytes"]
            == values["live_bytes_after"] + values["deallocated_bytes"],
            f"{label} live-byte balance does not reconcile")
    require(values["peak_live_bytes_before"] >= values["live_bytes_before"],
            f"{label} pre-operation peak is below live bytes")
    require(values["peak_live_bytes_after"] >= values["peak_live_bytes_before"],
            f"{label} peak moved backwards")
    require(values["peak_live_bytes_after"] >= values["live_bytes_after"],
            f"{label} post-operation peak is below live bytes")
    require(values["region_peak_live_bytes"] >= max(values["live_bytes_before"], values["live_bytes_after"]),
            f"{label} region peak is below live bytes")
    require(values["region_peak_live_bytes"] <= values["peak_live_bytes_after"],
            f"{label} region peak exceeds process peak")
    return {"status": status, "scope": sample["scope"], **values,
            "incremental_region_peak_live_bytes":
                values["region_peak_live_bytes"] - values["live_bytes_before"]}


def validate_rss(path: Path) -> dict[str, Any]:
    value = read_json(path)
    require(isinstance(value, dict), f"{path.name} RSS is not an object")
    require(set(value) == {"max_rss_kib", "elapsed_seconds", "user_seconds", "system_seconds"},
            f"{path.name} RSS fields are unexpected")
    nonnegative_integer(value["max_rss_kib"], f"{path.name}.max_rss_kib")
    for key in ("elapsed_seconds", "user_seconds", "system_seconds"):
        finite_number(value[key], f"{path.name}.{key}")
        require(value[key] >= 0, f"{path.name}.{key} is negative")
    return value


def verify_elapsed(raw: dict[str, Any], count: int, label: str) -> tuple[dict[str, Any], list[int]]:
    elapsed = raw.get("elapsed_ns")
    require(isinstance(elapsed, dict), f"{label}.elapsed_ns is not an object")
    values = check_vector(elapsed.get("samples"), count, f"{label}.elapsed_ns.samples")
    for index, value in enumerate(values):
        nonnegative_integer(value, f"{label}.elapsed_ns.samples[{index}]")
    order = elapsed.get("sample_order")
    require(isinstance(order, list) and sorted(order) == list(range(count)),
            f"{label}.elapsed_ns.sample_order is not a permutation")
    return verify_report_stats(elapsed, values, f"{label}.elapsed_ns"), values


def validate_result(raw: dict[str, Any], plan: dict[str, Any], job: dict[str, Any],
                    binary: dict[str, Any], allocator: bool) -> dict[str, Any]:
    label = job["name"]
    require(raw.get("schema_version") == 1, f"{label} schema version is not 1")
    tool = raw.get("tool")
    require(isinstance(tool, dict), f"{label}.tool is not an object")
    expected_binary = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    require(tool.get("binary") == expected_binary, f"{label} tool binary is unexpected")
    require(tool.get("profile") == "release", f"{label} tool profile is not release")
    require(tool.get("instrumentation") == (
        "system_allocator_operation_scoped" if allocator else "none"),
            f"{label} instrumentation is unexpected")
    binary_identity = raw.get("binary_identity")
    require(isinstance(binary_identity, dict), f"{label}.binary_identity is not an object")
    require(binary_identity.get("binary_sha256") == binary["sha256"],
            f"{label} report is bound to the wrong binary")
    require(binary_identity.get("profile") == "release", f"{label} binary profile is unexpected")
    environment = raw.get("environment")
    require(isinstance(environment, dict), f"{label}.environment is not an object")
    require(environment.get("git_revision") == plan["revision"], f"{label} revision is unexpected")
    require(environment.get("cpu_affinity") == str(plan["cpu"]), f"{label} CPU affinity is unexpected")
    configuration = raw.get("configuration")
    require(isinstance(configuration, dict), f"{label}.configuration is not an object")
    require(configuration.get("cases") == [job["case"]], f"{label} case configuration is unexpected")
    require(configuration.get("xlsx_cell_crud_shapes") == [job["shape"]],
            f"{label} shape configuration is unexpected")
    require(configuration.get("samples_per_case") == job["samples"], f"{label} sample count is unexpected")
    require(configuration.get("warmup_iterations_per_case") == job["warmup"], f"{label} warmup is unexpected")
    results = raw.get("results")
    require(isinstance(results, list) and len(results) == 1, f"{label} must contain one result")
    result = results[0]
    require(isinstance(result, dict), f"{label} result is not an object")
    require(result.get("case") == job["case"], f"{label} result case is unexpected")
    corpus = result.get("corpus")
    require(isinstance(corpus, dict) and corpus.get("shape") == job["shape"],
            f"{label} corpus identity is unexpected")
    elapsed_stats, elapsed_values = verify_elapsed(result, job["samples"], label)
    source = result.get("source")
    require(isinstance(source, dict), f"{label} has no source evidence")
    xlsx = source.get("xlsx_cell_values")
    require(isinstance(xlsx, dict), f"{label} has no XLSX source evidence")
    require(xlsx.get("implementation") in ("source-backed", "managed-source-backed"),
            f"{label} implementation is not source-backed")
    managed = xlsx.get("cache_mode") == "managed-budget"
    require(xlsx.get("cache_mode") in ("unmanaged-control", "managed-budget"),
            f"{label} cache mode is unexpected")
    require(xlsx["implementation"] == ("managed-source-backed" if managed else "source-backed"),
            f"{label} implementation and cache mode disagree")
    require(xlsx.get("cache_budget_managed") is managed,
            f"{label} cache mode and budget flag disagree")
    phase_values: dict[str, list[int]] = {}
    for phase in TIMING_PHASES:
        values = check_vector(xlsx.get(phase), job["samples"], f"{label}.{phase}")
        for index, value in enumerate(values):
            nonnegative_integer(value, f"{label}.{phase}[{index}]")
        phase_values[phase] = values
    for sorted_index, acquisition_index in enumerate(result["elapsed_ns"]["sample_order"]):
        phase_sum = sum(phase_values[phase][acquisition_index] for phase in PHASES)
        require(phase_sum == elapsed_values[sorted_index],
                f"{label} phase sum does not match elapsed sample {sorted_index}")
    phase_sums = [sum(phase_values[phase][index] for phase in PHASES)
                  for index in range(job["samples"])]
    require(sorted(range(job["samples"]), key=lambda index: (phase_sums[index], index))
            == result["elapsed_ns"]["sample_order"],
            f"{label} phase sums do not reproduce elapsed sample order")
    allocation_values = check_vector(xlsx.get("commit_allocation_metrics"), job["samples"],
                                     f"{label}.commit_allocation_metrics")
    allocation_rows = []
    for index, value in enumerate(allocation_values):
        allocation_rows.append({
            "index": index,
            **allocation_sample(value, f"{label}.commit_allocation_metrics[{index}]", allocator),
        })
    identity = {
        "configuration": configuration,
        "corpus": corpus,
        "sink": result.get("sink"),
        "source": source_identity(source, job["samples"]),
        "output_sha256": result.get("output_sha256"),
    }
    require(isinstance(identity["sink"], dict), f"{label}.sink is not an object")
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        nonnegative_integer(identity["sink"].get(key), f"{label}.sink.{key}")
    buckets = identity["sink"].get("write_size_buckets")
    require(isinstance(buckets, dict), f"{label} sink buckets are missing")
    for key, value in buckets.items():
        nonnegative_integer(value, f"{label}.sink.write_size_buckets.{key}")
    require(identity["sink"]["largest_write"] <= 65536,
            f"{label} exceeds the bounded sink write size")
    require(sum(buckets.values()) == identity["sink"]["write_calls"],
            f"{label} sink buckets do not reconcile with write calls")
    require(isinstance(identity["output_sha256"], str) and len(identity["output_sha256"]) == 64,
            f"{label} output digest is missing")
    source_output = xlsx.get("output_sha256")
    require(isinstance(source_output, list) and all(value == identity["output_sha256"] for value in source_output),
            f"{label} source output digest does not match result output")
    timing = {"elapsed_ns": elapsed_stats}
    for phase in TIMING_PHASES:
        timing[phase] = report_stats(phase_values[phase])
    row = {
        "name": label,
        "kind": job["kind"],
        "guard": job["guard"],
        "repeat": job["repeat"],
        "case": job["case"],
        "shape": job["shape"],
        "samples": job["samples"],
        "timing": timing,
        "phase_time_share": {
            phase: sum(phase_values[phase]) / sum(elapsed_values) for phase in PHASES
        },
        "rss": None,
        "allocation": allocation_rows,
        "identity": identity,
        "identity_sha256": identity_digest(identity),
    }
    return row


def expected_command(receipt: dict[str, Any], plan: dict[str, Any], job: dict[str, Any],
                     binary: dict[str, Any], allocator: bool, folder: Path) -> None:
    command = receipt.get("command")
    require(isinstance(command, list), f"{job['name']} receipt command is not a list")
    require(command[:3] == ["taskset", "-c", str(plan["cpu"])],
            f"{job['name']} does not use the planned CPU")
    require(binary["path"] in command, f"{job['name']} does not invoke the bound binary")
    for option, value in (("--warmup", job["warmup"]), ("--samples", job["samples"]),
                          ("--case", job["case"]), ("--xlsx-cell-crud-shape", job["shape"]),
                          ("--json", str(folder / (job["name"] + ".json")))):
        require(command.count(option) == 1 and command[command.index(option) + 1] == str(value),
                f"{job['name']} command has an unexpected {option}")
    if allocator:
        require("/usr/bin/time" not in command and "valgrind" not in command,
                f"{job['name']} allocator command has a native profiler")
    else:
        require("/usr/bin/time" in command, f"{job['name']} has no whole-child RSS observer")
        require("-o" in command and command[command.index("-o") + 1] == str(folder / (job["name"] + ".rss.json")),
                f"{job['name']} RSS path is not bound to its artifact")


def build_command(receipt: dict[str, Any], executable: str, allocator: bool) -> None:
    command = receipt.get("command")
    require(isinstance(command, list), "build receipt command is not a list")
    for option in ("cargo", "build", "--release", "--locked", "--manifest-path", "--bin"):
        require(option in command, f"build receipt is missing {option}")
    require(command[command.index("--bin") + 1] == executable, "build receipt binary is unexpected")
    if allocator:
        require("--features" in command and command[command.index("--features") + 1] == "allocator-metrics",
                "allocator build is missing allocator-metrics")
    else:
        require("--features" not in command, "normal build unexpectedly enables allocator metrics")


def check_receipt(folder: Path, plan: dict[str, Any], job: dict[str, Any], binary: dict[str, Any],
                  allocator: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    name = job["name"]
    receipt_path = folder / f"{name}.receipt.json"
    receipt = read_json(receipt_path)
    require(receipt.get("exit_code") == 0, f"{name} did not exit successfully")
    require(receipt.get("binary_sha256") == binary["sha256"], f"{name} receipt binary mismatch")
    require(receipt.get("plan_sha256") == sha(PLAN), f"{name} receipt plan mismatch")
    require(receipt.get("script_sha256") == sha(HERE / "run.py"), f"{name} receipt script mismatch")
    manifest_path = folder / "source-manifest.json"
    require(receipt.get("source_manifest_sha256") == sha(manifest_path), f"{name} manifest mismatch")
    expected_command(receipt, plan, job, binary, allocator, folder)
    suffixes = [".json", ".stdout", ".stderr"]
    if not allocator:
        suffixes.append(".rss.json")
    expected_artifacts = {name + suffix for suffix in suffixes}
    require(set(receipt.get("artifacts", {})) == expected_artifacts,
            f"{name} receipt artifact inventory mismatch")
    for filename, digest in receipt["artifacts"].items():
        artifact = folder / filename
        require(artifact.is_file() and not artifact.is_symlink(), f"{name} artifact is not regular: {filename}")
        require(sha(artifact) == digest, f"{name} artifact digest mismatch: {filename}")
    raw = read_json(folder / f"{name}.json")
    row = validate_result(raw, plan, job, binary, allocator)
    if not allocator:
        row["rss"] = {"scope": "whole_child_process",
                       **validate_rss(folder / f"{name}.rss.json")}
    return receipt, row


def check_binary(folder: Path, allocator: bool) -> dict[str, Any]:
    label = "alloc" if allocator else "normal"
    identity = read_json(folder / f"binary-{label}.json")
    require(isinstance(identity, dict), f"{folder.name} binary identity is not an object")
    path = Path(identity.get("path", ""))
    plan = read_json(PLAN)
    require(path == Path(plan['owned_paths'][0]) / f"{folder.name}-{label}",
            f"{folder.name} binary path is unexpected")
    digest = identity.get("sha256")
    require(isinstance(digest, str) and len(digest) == 64
            and all(char in '0123456789abcdef' for char in digest),
            f"{folder.name} binary digest is malformed")
    nonnegative_integer(identity.get("bytes"), f"{folder.name} binary size")
    require(identity['bytes'] > 0, f"{folder.name} binary size is zero")
    require(not path.is_symlink(), f"{folder.name} binary is a symlink")
    if path.exists():
        require(path.is_file(), f"{folder.name} binary is not a regular file")
        require(identity['sha256'] == sha(path), f"{folder.name} binary digest mismatch")
        require(identity['bytes'] == path.stat().st_size, f"{folder.name} binary size mismatch")
    else:
        # Captures hash the live executable before and after each child. Once
        # owned builds are removed, replay checks their retained custody chain.
        cleanup = read_json(HERE / 'cleanup.json')
        require(cleanup.get('owned_paths_absent') is True
                and cleanup.get('accessible_process_references') == []
                and cleanup.get('removed') == plan['owned_paths']
                and all(not Path(name).exists() for name in plan['owned_paths']),
                f"{folder.name} missing binary has no completed cleanup record")
    require(identity.get("source_manifest_sha256") == sha(folder / "source-manifest.json"),
            f"{folder.name} binary manifest mismatch")
    build_name = "build-alloc" if allocator else "build-normal"
    build_receipt_path = folder / f"{build_name}.receipt.json"
    build_receipt = read_json(build_receipt_path)
    require(identity.get("build_receipt_sha256") == sha(build_receipt_path),
            f"{folder.name} binary build receipt mismatch")
    require(build_receipt.get("exit_code") == 0 and build_receipt.get("binary_sha256") is None,
            f"{folder.name} build receipt is invalid")
    require(build_receipt.get("plan_sha256") == sha(PLAN), f"{folder.name} build plan mismatch")
    require(build_receipt.get("script_sha256") == sha(HERE / "run.py"), f"{folder.name} build script mismatch")
    require(build_receipt.get("source_manifest_sha256") == sha(folder / "source-manifest.json"),
            f"{folder.name} build manifest mismatch")
    executable = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    build_command(build_receipt, executable, allocator)
    expected = {f"{build_name}.stdout", f"{build_name}.stderr"}
    require(set(build_receipt.get("artifacts", {})) == expected,
            f"{folder.name} build artifact inventory mismatch")
    for filename, digest in build_receipt["artifacts"].items():
        artifact = folder / filename
        require(artifact.is_file() and sha(artifact) == digest, f"{folder.name} build artifact mismatch: {filename}")
    return {"sha256": identity["sha256"], "bytes": identity["bytes"], "path": str(path),
            "binary": executable, "build_receipt_sha256": identity["build_receipt_sha256"]}


def check_stage(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    require(folder.is_dir(), f"{stage} stage directory is missing")
    manifest = folder / "source-manifest.json"
    require(manifest.is_file() and not manifest.is_symlink(), f"{stage} source manifest is missing")
    manifest_value = read_json(manifest)
    require(isinstance(manifest_value, dict) and manifest_value, f"{stage} source manifest is empty")
    identities = {
        "normal": check_binary(folder, False),
        "allocator": check_binary(folder, True),
    }
    rows: dict[str, list[dict[str, Any]]] = {"native": [], "allocation": []}
    receipts: list[dict[str, Any]] = []
    for lane, allocator in (("native", False), ("allocation", True)):
        jobs = expected_jobs(plan, "native" if lane == "native" else "alloc")
        expected_names = {job["name"] for job in jobs}
        actual_names = {
            path.name[:-len(".receipt.json")]
            for path in folder.glob("*.receipt.json")
            if path.name.startswith("native-") or path.name.startswith("alloc-")
        }
        prefix = "native-" if lane == "native" else "alloc-"
        actual_names = {name for name in actual_names if name.startswith(prefix)}
        require(actual_names == expected_names,
                f"{stage} {lane} receipt set differs from plan: {sorted(actual_names ^ expected_names)}")
        for job in jobs:
            receipt, row = check_receipt(folder, plan, job, identities["allocator" if allocator else "normal"], allocator)
            row["stage"] = stage
            rows[lane].append(row)
            receipts.append(receipt)
    ordered_receipts = sorted(receipts, key=lambda item: item["start_utc"])
    require(all(left["end_utc"] <= right["start_utc"]
                for left, right in zip(ordered_receipts, ordered_receipts[1:])),
            f"{stage} child receipts overlap")
    for lane in rows:
        rows[lane].sort(key=lambda row: (job_key(row), row["name"]))
    # Allocator and normal reports use different binary identities, but their
    # logical corpus/sink/source/output identity must agree for the primary.
    native_by_shape = {(row["case"], row["shape"]): row for row in rows["native"]
                       if row["kind"] == "primary"}
    for row in rows["allocation"]:
        counterpart = native_by_shape.get((row["case"], row["shape"]))
        require(counterpart is not None, f"{stage} allocator row has no native counterpart")
        require(normalize_iteration_counts(row["identity"])
                == normalize_iteration_counts(counterpart["identity"]),
                f"{stage} allocator and native logical identities differ for {row['shape']}")
    return {
        "stage": stage,
        "manifest_sha256": sha(manifest),
        "binary_identities": identities,
        "native": {"rows": rows["native"], "row_count": len(rows["native"]),
                    "total_samples": sum(row["samples"] for row in rows["native"])},
        "allocation": {"rows": rows["allocation"], "row_count": len(rows["allocation"]),
                        "total_samples": sum(row["samples"] for row in rows["allocation"])},
        "custody": {"receipt_count": len(receipts), "receipts_non_overlapping": True,
                     "source_manifest_entries": len(manifest_value)},
    }


def percent_change(baseline: float, candidate: float) -> float | None:
    if baseline == 0:
        return 0.0 if candidate == 0 else None
    return (candidate / baseline - 1.0) * 100.0


def comparison_record(baseline: Any, candidate: Any) -> dict[str, Any]:
    change = percent_change(float(baseline), float(candidate))
    return {"baseline": baseline, "candidate": candidate, "change_percent": change}


def bootstrap_median_ratio(baseline: list[int], candidate: list[int], rng: random.Random) -> dict[str, Any]:
    require(len(baseline) == len(candidate) and baseline, "bootstrap vectors do not match")
    ratios: list[float] = []
    count = len(baseline)
    for _ in range(BOOTSTRAP_ITERATIONS):
        left = sorted(rng.choice(baseline) for _ in range(count))
        right = sorted(rng.choice(candidate) for _ in range(count))
        left_median = (left[(count - 1) // 2] + left[count // 2]) / 2.0
        right_median = (right[(count - 1) // 2] + right[count // 2]) / 2.0
        if left_median == 0:
            ratios.append(1.0 if right_median == 0 else math.inf)
        else:
            ratios.append(right_median / left_median)
    ratios.sort()
    require(all(math.isfinite(value) for value in ratios), "bootstrap ratio is not finite")
    return {"iterations": BOOTSTRAP_ITERATIONS, "seed": BOOTSTRAP_SEED,
            "low": ratios[math.ceil(len(ratios) * 0.025) - 1],
            "high": ratios[math.ceil(len(ratios) * 0.975) - 1],
            "unit": "within-child resampling; candidate median / baseline median"}


def compare_stages(baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    comparisons: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    drift: list[dict[str, Any]] = []
    rng = random.Random(BOOTSTRAP_SEED)

    def compare_lane(lane: str, include_bootstrap: bool) -> None:
        left_rows = {job_key(row): row for row in baseline[lane]["rows"]}
        right_rows = {job_key(row): row for row in candidate[lane]["rows"]}
        require(set(left_rows) == set(right_rows), f"{lane} baseline/candidate row keys differ")
        for key in sorted(left_rows):
            left, right = left_rows[key], right_rows[key]
            metrics: dict[str, Any] = {}
            for phase in ("elapsed_ns",) + TIMING_PHASES:
                metrics[phase] = {}
                for stat in TIMING_STATS:
                    value = comparison_record(left["timing"][phase][stat], right["timing"][phase][stat])
                    metrics[phase][stat] = value
                    change = value["change_percent"]
                    if change is not None and change > THRESHOLD_PERCENT:
                        adverse.append({"lane": lane, "case": key[3], "shape": key[4],
                                        "repeat": key[2], "guard": key[1], "phase": phase,
                                        "stat": stat, **value})
            record = {"lane": lane, "case": key[3], "shape": key[4], "repeat": key[2],
                      "guard": key[1], "metrics": metrics,
                      "identity_equal": left["identity"] == right["identity"]}
            require(record["identity_equal"],
                    f"baseline/candidate logical identity differs for {lane} {key}")
            if include_bootstrap:
                # The sorted samples are retained in every timing statistic.
                # Bootstrap from elapsed acquisition values is descriptive and
                # explicitly scoped to this matched pair of children.
                record["elapsed_ns_p50_bootstrap"] = bootstrap_median_ratio(
                    left["timing"]["elapsed_ns"]["samples"],
                    right["timing"]["elapsed_ns"]["samples"], rng)
            if lane == "native":
                left_rss, right_rss = left["rss"], right["rss"]
                rss_change = percent_change(left_rss["max_rss_kib"], right_rss["max_rss_kib"])
                record["rss"] = comparison_record(left_rss["max_rss_kib"], right_rss["max_rss_kib"])
                if rss_change is not None and rss_change > THRESHOLD_PERCENT:
                    adverse.append({"lane": lane, "case": key[3], "shape": key[4],
                                    "repeat": key[2], "guard": key[1], "phase": "rss",
                                    "stat": "max_rss_kib", **record["rss"]})
            comparisons.append(record)
        # Same-build drift compares matched repeat children within each stage.
        groups: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
        for row in baseline[lane]["rows"] + candidate[lane]["rows"]:
            groups.setdefault((row["stage"], row["kind"], row["guard"], row["case"], row["shape"]), []).append(row)
        for group, rows in sorted(groups.items(), key=lambda item: str(item[0])):
            by_repeat = {row["repeat"]: row for row in rows}
            for first_repeat in sorted(by_repeat):
                for second_repeat in sorted(by_repeat):
                    if second_repeat <= first_repeat:
                        continue
                    first, second = by_repeat[first_repeat], by_repeat[second_repeat]
                    for phase in ("elapsed_ns",) + TIMING_PHASES:
                        for stat in TIMING_STATS:
                            value = comparison_record(first["timing"][phase][stat], second["timing"][phase][stat])
                            change = value["change_percent"]
                            if change is not None and abs(change) > THRESHOLD_PERCENT:
                                drift.append({"lane": lane, "stage": group[0], "case": group[3],
                                              "shape": group[4], "kind": group[1], "guard": group[2],
                                              "repeat_first": first_repeat, "repeat_second": second_repeat,
                                              "phase": phase, "stat": stat, **value})

    compare_lane("native", True)
    allocation_comparisons: list[dict[str, Any]] = []
    left_alloc = {job_key(row): row for row in baseline["allocation"]["rows"]}
    right_alloc = {job_key(row): row for row in candidate["allocation"]["rows"]}
    require(set(left_alloc) == set(right_alloc), "allocator baseline/candidate row keys differ")
    for key in sorted(left_alloc):
        left, right = left_alloc[key], right_alloc[key]
        require(left["identity"] == right["identity"],
                f"baseline/candidate allocator logical identity differs for {key}")
        require(len(left["allocation"]) == len(right["allocation"]),
                f"allocator sample counts differ for {key}")
        metrics = {field: {"baseline": [], "candidate": [], "delta": []}
                   for field in ALLOCATION_REPORT_FIELDS}
        for left_sample, right_sample in zip(left["allocation"], right["allocation"]):
            require(left_sample["status"] == right_sample["status"] == "measured",
                    f"allocator sample status is not measured for {key}")
            for field in ALLOCATION_REPORT_FIELDS:
                left_value, right_value = left_sample[field], right_sample[field]
                metrics[field]["baseline"].append(left_value)
                metrics[field]["candidate"].append(right_value)
                metrics[field]["delta"].append(right_value - left_value)
        allocation_comparisons.append({"case": key[3], "shape": key[4], "repeat": key[2],
                                       "identity_equal": True, "metrics": metrics})
    return {"timing_comparisons": comparisons, "allocation_comparisons": allocation_comparisons,
            "adverse_flags_over_five_percent": adverse,
            "same_build_drift_over_five_percent": drift,
            "bootstrap": {"iterations": BOOTSTRAP_ITERATIONS, "seed": BOOTSTRAP_SEED,
                          "scope": "primary elapsed p50 within matched child; descriptive only"}}


def analyze(stage: str | None) -> dict[str, Any]:
    plan = read_json(PLAN)
    require(isinstance(plan, dict), "plan is not an object")
    if stage in ("baseline", "candidate"):
        return {"status": "pass", "stage": stage, "plan_sha256": sha(PLAN),
                "evidence": check_stage(stage, plan),
                "scope": "0521 XLSX source-backed cell-values native and allocator evidence"}
    baseline = check_stage("baseline", plan)
    candidate = check_stage("candidate", plan)
    return {"status": "pass", "stage": "compare", "plan_sha256": sha(PLAN),
            "baseline": baseline, "candidate": candidate,
            "comparison": compare_stages(baseline, candidate),
            "scope": "0521 XLSX source-backed cell-values native and allocator evidence"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("baseline", "candidate", "compare"),
                        help="analyze one retained stage, or compare both stages")
    parser.add_argument("--output", type=Path, help="JSON output path")
    parser.add_argument("output_positional", nargs="?", type=Path,
                        help="compatibility alias for --output")
    args = parser.parse_args()
    stage = args.stage
    if stage == "compare":
        stage = None
    output = args.output or args.output_positional
    if output is None:
        output = HERE / (f"analysis-{stage}.json" if stage else "comparison.json")
    try:
        result = analyze(stage)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    except EvidenceError as error:
        print(f"evidence check failed: {error}", file=sys.stderr)
        return 1
    print(f"0521 {result['stage']} evidence verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
