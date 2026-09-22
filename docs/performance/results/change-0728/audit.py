#!/usr/bin/env python3
"""Independently replay the frozen 0728 OLE2 baseline evidence.

This audit reads only the retained plan, custody receipts, manifest and raw
probe JSON.  It never imports or executes ``analyze.py``, Cargo, a Rust
binary, a native profiler, or a capture command.  Timing statistics and
within-process phase fractions are recomputed here, and the resulting process
rows are compared with the primary ``analysis.json`` using a documented tight
floating-point tolerance.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import sys
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
CAPTURES = PACKET / "captures"
HEX = set("0123456789abcdef")
LANES = ("native", "allocation")
ROUTES = (("format", "reuse"), ("container", "reuse"), ("container", "rewrite"))
STAT_KEYS = ("p50", "mean", "p95", "p99", "maximum")
REL_TOL = 1.0e-12
ABS_TOL = 1.0e-9


class AuditError(Exception):
    """A custody, matrix, oracle, or replay failure."""


class Pending(Exception):
    """The terminal capture artifacts do not exist yet."""

    def __init__(self, items: Iterable[str]):
        self.items = tuple(dict.fromkeys(items))
        super().__init__("terminal evidence is pending")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def read_json(path: Path) -> Any:
    if not path.is_file() or path.is_symlink():
        raise Pending((path.relative_to(PACKET).as_posix(),))
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AuditError(f"invalid JSON {path}: {error}") from error


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing or symlinked file: {path}")
    return sha_bytes(path.read_bytes())


def require_digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")
    return value


def require_int(value: Any, label: str, minimum: int = 0) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= minimum,
            f"{label} is not an integer >= {minimum}")
    return value


def safe_capture_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str), f"{label} path is not a string")
    relative = Path(raw)
    require(not relative.is_absolute() and ".." not in relative.parts,
            f"{label} path escapes the capture directory")
    resolved = (CAPTURES / relative).resolve()
    root = CAPTURES.resolve()
    require(resolved == root or root in resolved.parents,
            f"{label} path is outside the capture directory")
    require(resolved.is_file() and not resolved.is_symlink(),
            f"{label} is missing or symlinked: {raw}")
    return resolved


def load_plan() -> dict[str, Any]:
    plan = read_json(PACKET / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("cpu") == 12, "CPU binding changed")
    require(plan.get("cycles") == 2, "native cycle count changed")
    require(plan.get("processes_per_route_cycle") == 3,
            "processes per route changed")
    require(plan.get("samples") == 30 and plan.get("warmups") == 2,
            "native sample plan changed")
    require(plan.get("cases") == ["docfloat", "docnohf", "ppt45543"],
            "case order changed")
    require(plan.get("routes") == ["format-default", "container-reuse", "container-rewrite"],
            "route matrix changed")
    require(plan.get("native_processes") == 54 and plan.get("allocation_processes") == 27,
            "process total changed")
    require(plan.get("allocation_processes_per_route") == 3,
            "allocation process count changed")
    require(plan.get("allocation_samples") == 1 and plan.get("allocation_warmups") == 0,
            "allocation sample plan changed")
    return plan


def load_cases(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    rows = read_json(PACKET / "cases.json")
    require(isinstance(rows, list), "cases receipt is not a list")
    require([row.get("case") for row in rows] == plan["cases"],
            "case receipt order changed")
    result: dict[str, dict[str, Any]] = {}
    for row in rows:
        require(isinstance(row, dict), "case receipt row is not an object")
        case = row.get("case")
        require(isinstance(case, str) and case not in result, "case identity is invalid")
        path = row.get("path")
        require(isinstance(path, str) and not Path(path).is_absolute()
                and ".." not in Path(path).parts, f"{case}: fixture path changed")
        fixture = ROOT / path
        require(fixture.is_file() and not fixture.is_symlink(), f"missing fixture: {fixture}")
        require(row.get("bytes") == fixture.stat().st_size, f"{case}: fixture size changed")
        require(row.get("sha256") == sha(fixture), f"{case}: fixture digest changed")
        require(row.get("format") in ("doc", "ppt"), f"{case}: fixture format missing")
        result[case] = row
    return result


def load_contract(cases: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    contract = read_json(PACKET / "oracle-contract.json")
    require(isinstance(contract, dict) and set(contract) == set(cases),
            "oracle contract case set changed")
    required = {
        "directory_metadata_fields", "allocation_ownership_contract",
        "semantic_witness", "control_names", "identity",
    }
    for case, value in contract.items():
        require(isinstance(value, dict) and set(value) == required,
                f"{case}: oracle contract schema changed")
        require(isinstance(value["directory_metadata_fields"], list)
                and value["directory_metadata_fields"],
                f"{case}: directory metadata contract is empty")
        require(isinstance(value["allocation_ownership_contract"], str)
                and value["allocation_ownership_contract"],
                f"{case}: allocation ownership contract is empty")
        require(isinstance(value["semantic_witness"], dict),
                f"{case}: semantic witness is not an object")
        require(isinstance(value["control_names"], list)
                and value["control_names"], f"{case}: control names are empty")
        require(isinstance(value["identity"], dict),
                f"{case}: identity is not an object")
    return contract


def load_builds() -> dict[str, Any]:
    builds = read_json(PACKET / "builds.json")
    require(isinstance(builds, dict), "builds receipt is not an object")
    binaries = builds.get("binaries")
    require(isinstance(binaries, list) and len(binaries) == 2,
            "binary receipt count changed")
    names = {Path(row.get("path", "")).name for row in binaries}
    require(names == {"ole_format_save_probe", "ole_format_save_probe_alloc"},
            "binary names changed")

    source = builds.get("source_sha256")
    probe = builds.get("probe_sha256")
    require(isinstance(source, dict) and source, "source custody map is empty")
    require(isinstance(probe, dict) and probe, "probe custody map is empty")
    for relative, expected in source.items():
        require(isinstance(relative, str) and not Path(relative).is_absolute()
                and ".." not in Path(relative).parts, f"source path is unsafe: {relative}")
        require_digest(expected, f"source custody {relative}")
        require(sha(ROOT / relative) == expected, f"source custody changed: {relative}")
    for relative, expected in probe.items():
        require(isinstance(relative, str) and not Path(relative).is_absolute()
                and ".." not in Path(relative).parts, f"probe path is unsafe: {relative}")
        require_digest(expected, f"probe custody {relative}")
        require(sha(PACKET / relative) == expected, f"probe custody changed: {relative}")
    return builds


def verify_binary_custody(builds: dict[str, Any], frozen: dict[str, Any]) -> None:
    expected = builds["binaries"]
    actual = frozen.get("binaries")
    require(actual == expected, "freeze binary identities differ from builds receipt")
    cleanup_path = PACKET / "cleanup.json"
    cleanup = read_json(cleanup_path) if cleanup_path.is_file() else None
    witnesses = cleanup.get("identities", []) if isinstance(cleanup, dict) else []
    require(isinstance(witnesses, list), "cleanup identities are not a list")
    for row in expected:
        path = Path(row["path"])
        if path.is_file():
            require(not path.is_symlink() and sha(path) == row["sha256"]
                    and path.stat().st_size == row["bytes"],
                    f"live binary identity changed: {path}")
        else:
            require(isinstance(cleanup, dict) and cleanup.get("removed") is True,
                    f"missing binary has no cleanup witness: {path}")
            require(row in witnesses, f"cleanup witness changed: {path}")


def verify_freeze(plan: dict[str, Any], cases: dict[str, dict[str, Any]],
                  builds: dict[str, Any]) -> dict[str, Any]:
    frozen = read_json(PACKET / "freeze.json")
    require(isinstance(frozen, dict), "freeze is not an object")
    bindings = frozen.get("bindings")
    require(isinstance(bindings, dict) and bindings, "freeze bindings are empty")
    required = {
        "docs/performance/results/change-0728/plan.json",
        "docs/performance/results/change-0728/cases.json",
        "docs/performance/results/change-0728/builds.json",
        "docs/performance/results/change-0728/capture.py",
        "docs/performance/results/change-0728/analyze.py",
        "docs/performance/results/change-0728/hypothesis.md",
        "docs/performance/results/change-0728/constraints.json",
        "docs/performance/results/change-0728/environment.json",
        "docs/performance/results/change-0728/qualification.json",
        "docs/performance/results/change-0728/oracle-contract.json",
    }
    require(required <= set(bindings), "freeze omits a required packet binding")
    for relative, expected in bindings.items():
        require(isinstance(relative, str) and not Path(relative).is_absolute()
                and ".." not in Path(relative).parts, f"freeze path is unsafe: {relative}")
        require_digest(expected, f"freeze binding {relative}")
        require(sha(ROOT / relative) == expected, f"frozen binding changed: {relative}")
    verify_binary_custody(builds, frozen)
    return frozen


def expected_matrix(plan: dict[str, Any], cases: dict[str, dict[str, Any]],
                    frozen: dict[str, Any]) -> list[dict[str, Any]]:
    binaries = {Path(row["path"]).name: row["path"] for row in frozen["binaries"]}
    expected: list[dict[str, Any]] = []
    for lane in LANES:
        cycles = range(plan["cycles"]) if lane == "native" else range(1)
        binary_name = "ole_format_save_probe" if lane == "native" else "ole_format_save_probe_alloc"
        binary = binaries[binary_name]
        samples = plan["samples"] if lane == "native" else plan["allocation_samples"]
        warmups = plan["warmups"] if lane == "native" else plan["allocation_warmups"]
        for cycle in cycles:
            cells = [(case, operation, policy)
                     for case in cases.values() for operation, policy in ROUTES]
            if cycle % 2:
                cells.reverse()
            for case, operation, policy in cells:
                for repeat in range(3):
                    name = f"{lane}-c{cycle}-{case['case']}-{operation}-{policy}-r{repeat}.json"
                    command = [
                        "taskset", "-c", str(plan["cpu"]), binary,
                        "--case", case["case"], "--input", case["path"],
                        "--operation", operation, "--policy", policy,
                        "--samples", str(samples), "--warmups", str(warmups),
                    ]
                    expected.append({
                        "lane": lane, "cycle": cycle, "case": case["case"],
                        "operation": operation, "policy": policy, "repeat": repeat,
                        "output": name, "stderr": name + ".stderr", "command": command,
                    })
    return expected


def oracle_guard(oracle: Any, label: str) -> None:
    require(isinstance(oracle, dict), f"{label}: oracle is not an object")
    require(oracle.get("oracle_ok") is True, f"{label}: oracle_ok is false")
    require(oracle.get("failure_reasons") == [], f"{label}: oracle has failures")
    for key, value in oracle.items():
        if isinstance(value, bool):
            require(value, f"{label}: oracle field {key} is false")


def inventory_streams(report: dict[str, Any], key: str, label: str) -> list[Any]:
    inventory = report.get(key)
    require(isinstance(inventory, dict), f"{label}: {key} is not an object")
    streams = inventory.get("streams")
    require(isinstance(streams, list) and streams, f"{label}: {key}.streams is empty")
    return streams


def check_report_header(report: dict[str, Any], row: dict[str, Any], case: dict[str, Any],
                        samples: int, warmups: int, native: bool) -> None:
    label = f"{row['lane']}/{row['cycle']}/{row['case']}/{row['operation']}/{row['policy']}/{row['repeat']}"
    require(report.get("schema_version") == 1, f"{label}: schema changed")
    for key, expected in {
        "case": row["case"], "format": case["format"], "operation": row["operation"],
        "policy": row["policy"], "input": case["path"], "samples_requested": samples,
        "warmups": warmups, "timing_claim": native,
        "allocator_instrumented": not native,
        "policy_applied": row["operation"] == "container",
        "policy_application_scope": (
            "common_container_editor" if row["operation"] == "container"
            else "not_applied_public_format_route"
        ),
    }.items():
        require(report.get(key) == expected, f"{label}: header {key} changed")
    expected_scope = (
        "public_format_open_edit_commit" if row["operation"] == "format"
        else "common_container_open_replace_and_validate_finish_control"
    )
    require(report.get("scope") == expected_scope, f"{label}: scope changed")
    require(report.get("source_sha256") == case["sha256"], f"{label}: source digest changed")
    require_digest(report.get("expected_output_sha256"), f"{label}: expected output hash")
    require_digest(report.get("replacements_sha256"), f"{label}: replacement hash")
    source_inventory = report.get("source_inventory")
    require(isinstance(source_inventory, dict), f"{label}: source inventory missing")
    require(source_inventory.get("file_bytes") == case["bytes"],
            f"{label}: source size changed")
    require(isinstance(report.get("policy_contract"), str) and report["policy_contract"],
            f"{label}: policy contract missing")


def validate_report(report: Any, row: dict[str, Any], case: dict[str, Any],
                    plan: dict[str, Any], route_outputs: dict[tuple[str, str, str], str],
                    identities: dict[str, dict[str, Any]],
                    allocations: dict[tuple[str, str, str], Any],
                    contract: dict[str, dict[str, Any]]) -> tuple[dict[str, list[float]], dict[str, list[float]], Any, str]:
    require(isinstance(report, dict), f"{row}: probe report is not an object")
    native = row["lane"] == "native"
    samples_requested = plan["samples"] if native else plan["allocation_samples"]
    warmups = plan["warmups"] if native else plan["allocation_warmups"]
    check_report_header(report, row, case, samples_requested, warmups, native)
    label = f"{row['lane']}/{row['cycle']}/{row['case']}/{row['operation']}/{row['policy']}/{row['repeat']}"
    expected_contract = contract[row["case"]]

    require(report.get("directory_metadata_fields") == expected_contract["directory_metadata_fields"],
            f"{label}: directory metadata contract changed")
    require(report.get("allocation_ownership_contract") == expected_contract["allocation_ownership_contract"],
            f"{label}: allocation ownership contract changed")
    expected_oracle = report.get("expected_oracle")
    require(isinstance(expected_oracle, dict), f"{label}: expected oracle is not an object")
    require(expected_oracle.get("semantic_witness") == expected_contract["semantic_witness"],
            f"{label}: semantic witness changed")
    proof = report.get("changed_length_proof")
    require(isinstance(proof, dict)
            and proof.get("format_specific_semantic_length_proven") is True,
            f"{label}: format-specific length proof failed")
    controls = report.get("oracle_controls")
    require(isinstance(controls, list)
            and [item.get("name") for item in controls] == expected_contract["control_names"],
            f"{label}: oracle control names changed")
    for control in controls:
        require(isinstance(control, dict) and control.get("status") == "rejected"
                and control.get("rejected") is True
                and isinstance(control.get("failure_reasons"), list)
                and bool(control["failure_reasons"])
                and all(isinstance(reason, str) and reason
                        for reason in control["failure_reasons"]),
                f"{label}: oracle control did not reject with a reason")

    oracle_guard(expected_oracle, f"{label}: expected")
    require(expected_oracle.get("directory_metadata_differences") == [],
            f"{label}: expected self-oracle reports metadata differences")
    expected_streams = inventory_streams(report, "expected_output_inventory", label)
    length_proof = proof
    require(isinstance(length_proof, dict), f"{label}: length proof missing")
    require(length_proof.get("logical_stream_length_change_proven") is True
            and length_proof.get("any_stream_length_changed") is True,
            f"{label}: logical changed-length proof failed")
    require(require_int(length_proof.get("changed_stream_count"), f"{label}: changed stream count", 1) >= 1,
            f"{label}: no changed stream")
    identity_keys = (
        "source_sha256", "expected_output_sha256", "replacements_sha256",
        "source_inventory", "expected_output_inventory", "replacements",
        "changed_length_proof",
    )
    identity = {key: report.get(key) for key in identity_keys}
    require(identity == expected_contract["identity"],
            f"{label}: frozen oracle identity changed")
    if row["case"] in identities:
        require(identity == identities[row["case"]], f"{label}: case identity changed")
    else:
        identities[row["case"]] = identity

    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == samples_requested,
            f"{label}: sample count changed")
    arrays: dict[str, list[float]] = {}
    phase_fractions: dict[str, list[float]] = {}
    output_hash: str | None = None
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("index") == index,
                f"{label}: sample ordinal changed")
        oracle_guard(sample.get("oracle"), f"{label}: sample {index}")
        sample_inventory = sample.get("output_inventory")
        require(isinstance(sample_inventory, dict), f"{label}: output inventory missing")
        require(sample_inventory.get("streams") == expected_streams,
                f"{label}: output stream inventory differs from expected")
        require(sample.get("oracle", {}).get("semantic_witness")
                == expected_contract["semantic_witness"],
                f"{label}: sample semantic witness changed")
        raw_directory = sample.get("oracle", {}).get("raw_directory")
        require(isinstance(raw_directory, dict),
                f"{label}: sample raw directory oracle is missing")
        require(raw_directory.get("source_expected_difference_bytes") == 0,
                f"{label}: source/expected raw directory difference is nonzero")
        if row["operation"] == "format" or row["policy"] == "reuse":
            require(raw_directory.get("source_output_difference_bytes")
                    == raw_directory.get("expected_output_difference_bytes") == 0,
                    f"{label}: reuse/format raw directory difference is nonzero")
        sample_hash = require_digest(sample.get("output_sha256"), f"{label}: output hash")
        if output_hash is None:
            output_hash = sample_hash
        require(sample_hash == output_hash, f"{label}: sample output hash drifted")
        route = (row["case"], row["operation"], row["policy"])
        if route in route_outputs:
            require(sample_hash == route_outputs[route], f"{label}: route output hash drifted")
        else:
            route_outputs[route] = sample_hash

        if native:
            require("allocations" not in sample, f"{label}: native sample has allocations")
            phase = sample.get("phase_ns")
            require(isinstance(phase, dict), f"{label}: timing phases missing")
            expected_keys = {"whole_ns"} if row["operation"] == "format" else {
                "open_ns", "stage_ns", "finish_ns", "whole_ns"
            }
            require(set(phase) == expected_keys, f"{label}: phase set changed")
            times: dict[str, int] = {}
            for key, value in phase.items():
                times[key] = require_int(value, f"{label}: {key}")
                arrays.setdefault(key, []).append(times[key])
            if row["operation"] == "container":
                require(sum(times[key] for key in ("open_ns", "stage_ns", "finish_ns"))
                        <= times["whole_ns"], f"{label}: phases exceed whole time")
                require(times["whole_ns"] > 0, f"{label}: zero whole time")
                for key in ("open_ns", "stage_ns", "finish_ns"):
                    phase_fractions.setdefault(key, []).append(
                        times[key] / times["whole_ns"] * 100.0
                    )
        else:
            require("phase_ns" not in sample, f"{label}: allocation sample has timing")
            allocation = sample.get("allocations")
            require(isinstance(allocation, dict), f"{label}: allocation sample missing")
            route = (row["case"], row["operation"], row["policy"])
            if route in allocations:
                require(allocation == allocations[route], f"{label}: allocation result drifted")
            else:
                allocations[route] = allocation
            expected_allocation_keys = {"whole"} if row["operation"] == "format" else {
                "open", "stage", "finish"
            }
            require(set(allocation) == expected_allocation_keys,
                    f"{label}: allocation phase set changed")

        if row["operation"] == "format":
            require(sample_hash == report["expected_output_sha256"],
                    f"{label}: public output differs from expected output")
    require(output_hash is not None, f"{label}: no output hash")
    if row["operation"] == "format":
        require(report["expected_output_sha256"] == output_hash,
                f"{label}: public output identity changed")
    return arrays, phase_fractions, samples[0].get("allocations"), output_hash


def median(values: list[float]) -> float:
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / 2


def stats(values: list[float]) -> dict[str, float | int]:
    require(values, "empty independent timing vector")
    ordered = sorted(values)
    return {
        "n": len(values),
        "p50": median(ordered),
        "mean": math.fsum(values) / len(values),
        "p95": ordered[math.ceil(0.95 * len(ordered)) - 1],
        "p99": ordered[math.ceil(0.99 * len(ordered)) - 1],
        "maximum": ordered[-1],
    }


def close_json(actual: Any, expected: Any, path: str = "analysis") -> None:
    if isinstance(actual, bool) or isinstance(expected, bool):
        require(actual is expected, f"{path}: boolean differs")
        return
    if isinstance(actual, dict) or isinstance(expected, dict):
        require(isinstance(actual, dict) and isinstance(expected, dict),
                f"{path}: object shape differs")
        require(set(actual) == set(expected), f"{path}: object keys differ")
        for key in expected:
            close_json(actual[key], expected[key], f"{path}.{key}")
        return
    if isinstance(actual, list) or isinstance(expected, list):
        require(isinstance(actual, list) and isinstance(expected, list)
                and len(actual) == len(expected), f"{path}: list shape differs")
        for index, (left, right) in enumerate(zip(actual, expected)):
            close_json(left, right, f"{path}[{index}]")
        return
    if isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
        require(math.isclose(float(actual), float(expected), rel_tol=REL_TOL, abs_tol=ABS_TOL),
                f"{path}: {actual!r} differs from {expected!r}")
        return
    require(actual == expected, f"{path}: value differs")


def audit_capture(plan: dict[str, Any], cases: dict[str, dict[str, Any]],
                  contract: dict[str, dict[str, Any]],
                  frozen: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    manifest_path = CAPTURES / "manifest.json"
    if not manifest_path.is_file():
        raise Pending(("captures/manifest.json",))
    manifest = read_json(manifest_path)
    require(manifest.get("status") == "complete", "capture manifest is not complete")
    require(manifest.get("freeze_sha256") == sha(PACKET / "freeze.json"),
            "capture freeze hash changed")
    require(manifest.get("bindings_start") == frozen["bindings"]
            and manifest.get("bindings_end") == frozen["bindings"],
            "capture bindings changed during run")

    expected = expected_matrix(plan, cases, frozen)
    runs = manifest.get("runs")
    require(isinstance(runs, list) and len(runs) == len(expected),
            f"capture process count changed: expected {len(expected)}")
    actual_keys = []
    raw_paths: set[str] = set()
    identities: dict[str, dict[str, Any]] = {}
    allocations: dict[tuple[str, str, str], Any] = {}
    route_outputs: dict[tuple[str, str, str], str] = {}
    process_rows: list[dict[str, Any]] = []
    for run, planned in zip(runs, expected):
        require(isinstance(run, dict), "capture run is not an object")
        for key in ("lane", "cycle", "case", "operation", "policy", "repeat"):
            require(run.get(key) == planned[key], f"matrix row {planned}: {key} changed")
        require(run.get("exit_code") == 0, f"{planned}: child process failed")
        require(run.get("command") == planned["command"], f"{planned}: command changed")
        output = safe_capture_path(run.get("output"), f"{planned}: output")
        stderr = safe_capture_path(run.get("stderr"), f"{planned}: stderr")
        output_rel = output.relative_to(CAPTURES).as_posix()
        stderr_rel = stderr.relative_to(CAPTURES).as_posix()
        require(output_rel not in raw_paths and stderr_rel not in raw_paths,
                f"{planned}: raw path reused")
        raw_paths.update((output_rel, stderr_rel))
        require(run.get("output") == planned["output"] and run.get("stderr") == planned["stderr"],
                f"{planned}: raw names changed")
        require(run.get("sha256") == sha(output), f"{planned}: output hash changed")
        require(run.get("stderr_sha256") == sha(stderr), f"{planned}: stderr hash changed")
        case = cases[planned["case"]]
        report = read_json(output)
        arrays, phase_fractions, allocation, output_hash = validate_report(
            report, planned, case, plan, route_outputs, identities, allocations, contract
        )
        timing = {key: stats(values) for key, values in arrays.items()}
        fractions = {key: stats(values) for key, values in phase_fractions.items()}
        process_rows.append({
            "lane": planned["lane"], "cycle": planned["cycle"], "case": planned["case"],
            "operation": planned["operation"], "policy": planned["policy"],
            "repeat": planned["repeat"], "output_sha256": output_hash,
            "timing": timing, "phase_percent": fractions, "allocations": allocation,
        })
        actual_keys.append(tuple(planned[key] for key in
                                 ("lane", "cycle", "case", "operation", "policy", "repeat")))
    require(len(actual_keys) == len(set(actual_keys)), "duplicate matrix identity")

    expected_raw = {row["output"] for row in expected} | {row["stderr"] for row in expected}
    require(raw_paths == expected_raw, "capture raw path set differs")
    on_disk_raw = {path.relative_to(CAPTURES).as_posix() for path in CAPTURES.rglob("*")
                   if path.is_file() and path.name != "manifest.json"}
    require(on_disk_raw == expected_raw, "capture directory contains unmanifested raw file")

    analysis = read_json(PACKET / "analysis.json")
    independent = {
        "disposition": "unchanged production baseline; no speedup claim",
        "processes": process_rows,
        "case_identities": identities,
    }
    close_json(analysis, independent)
    return manifest, independent


def custody_record(plan: dict[str, Any], builds: dict[str, Any], frozen: dict[str, Any],
                   manifest: dict[str, Any], independent: dict[str, Any]) -> dict[str, Any]:
    """Return custody data independent of whether binaries are live or cleaned."""
    return {
        "plan_sha256": sha(PACKET / "plan.json"),
        "freeze_sha256": sha(PACKET / "freeze.json"),
        "manifest_sha256": sha(CAPTURES / "manifest.json"),
        "analysis_sha256": sha(PACKET / "analysis.json"),
        "qualification_sha256": sha(PACKET / "qualification.json"),
        "oracle_contract_sha256": sha(PACKET / "oracle-contract.json"),
        "frozen_binding_count": len(frozen["bindings"]),
        "source_sha256": builds["source_sha256"],
        "probe_sha256": builds["probe_sha256"],
        "fixtures": {
            case: {"path": row["path"], "bytes": row["bytes"], "sha256": row["sha256"]}
            for case, row in sorted(load_cases(plan).items())
        },
        "binaries": [
            {"path": row["path"], "bytes": row["bytes"], "sha256": row["sha256"]}
            for row in frozen["binaries"]
        ],
        "processes": len(independent["processes"]),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--draft", action="store_true",
                        help="validate available custody and report pending terminal files")
    args = parser.parse_args()
    try:
        plan = load_plan()
        cases = load_cases(plan)
        contract = load_contract(cases)
        builds = load_builds()
        frozen = verify_freeze(plan, cases, builds)
        if args.draft and not (CAPTURES / "manifest.json").is_file():
            print(json.dumps({"status": "DRAFT PASS", "pending": ["captures/manifest.json"]},
                             sort_keys=True))
            return 0
        manifest, independent = audit_capture(plan, cases, contract, frozen)
        record = {
            "packet": "change-0728",
            "disposition": independent["disposition"],
            "matrix": {"processes": len(independent["processes"]), "lanes": list(LANES),
                        "routes": [list(route) for route in ROUTES]},
            "statistics": {"quantiles": "nearest-rank p95/p99; midpoint p50",
                            "fields": list(STAT_KEYS), "mean": "math.fsum / n",
                            "relative_tolerance": REL_TOL,
                            "absolute_tolerance": ABS_TOL},
            "custody": custody_record(plan, builds, frozen, manifest, independent),
            "processes": independent["processes"],
            "case_identities": independent["case_identities"],
            "analysis_match": True,
        }
        (PACKET / "audit.json").write_text(json.dumps(record, indent=2, sort_keys=True) + "\n",
                                             encoding="utf-8")
        print(json.dumps({"status": "PASS", "processes": len(independent["processes"]),
                          "analysis_match": True}, sort_keys=True))
        return 0
    except Pending as pending:
        print("PENDING: " + "; ".join(pending.items), file=sys.stderr)
        return 2
    except (AuditError, OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
