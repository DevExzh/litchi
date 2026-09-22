#!/usr/bin/env python3
"""Independently audit the frozen 0729 DOC phase-attribution capture.

The primary analyzer is deliberately not imported here.  This replay checks
the packet and binary custody, reconstructs the complete 72-process matrix,
validates the public DOC oracle and the profiled event traces, recomputes the
per-process timing vectors and statistics, and compares its result with
``analysis.json``.  It does not invoke Cargo, the probe, or a native command.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
CAPTURES = PACKET / "captures"
HEX = set("0123456789abcdef")
CASES = ("docfloat", "docnohf")
ROUTES = ("ordinary-opaque", "ordinary-split", "profiled-empty", "profiled-clock")
COMPARISONS = (
    ("ordinary-opaque", "ordinary-split"),
    ("ordinary-split", "profiled-empty"),
    ("profiled-empty", "profiled-clock"),
)
OUTER = ("open_ns", "edit_ns", "replace_ns", "commit_ns", "output_copy_ns")
STAT_KEYS = ("n", "p50", "mean", "p95", "p99", "maximum")
REL_TOL = 1.0e-12
ABS_TOL = 1.0e-9
OPEN_PHASES = ("StrictOwnerValidation", "PublicReaderValidation", "SourceRetention")
COMMIT_PHASES = (
    "Finish",
    "StrictOwnerValidation",
    "PublicReaderValidation",
    "SourceRetention",
    "Patch",
)
REQUIRED_FREEZE_FILES = {
    "docs/performance/results/change-0729/plan.json",
    "docs/performance/results/change-0729/cases.json",
    "docs/performance/results/change-0729/builds.json",
    "docs/performance/results/change-0729/capture.py",
    "docs/performance/results/change-0729/analyze.py",
    "docs/performance/results/change-0729/hypothesis.md",
    "docs/performance/results/change-0729/constraints.json",
    "docs/performance/results/change-0729/environment.json",
    "docs/performance/results/change-0729/oracle-contract.json",
    "docs/performance/results/change-0729/prior-oracle-contract.json",
}


class AuditError(Exception):
    """A custody, matrix, oracle, timing, or statistics failure."""


class Pending(Exception):
    """Terminal capture artifacts have not been produced yet."""

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


def digest(value: Any, label: str) -> str:
    require(
        isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
        f"{label} is not a lowercase SHA-256 digest",
    )
    return value


def integer(value: Any, label: str, minimum: int | None = 0) -> int:
    require(isinstance(value, int) and not isinstance(value, bool), f"{label} is not an integer")
    if minimum is not None:
        require(value >= minimum, f"{label} is below {minimum}")
    return value


def relative_root(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def safe_capture_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str), f"{label} path is not a string")
    relative = Path(raw)
    require(
        not relative.is_absolute() and ".." not in relative.parts,
        f"{label} path escapes captures: {raw}",
    )
    resolved = (CAPTURES / relative).resolve()
    base = CAPTURES.resolve()
    require(base == resolved or base in resolved.parents, f"{label} path escapes captures: {raw}")
    require(resolved.is_file() and not resolved.is_symlink(), f"{label} is missing or symlinked: {raw}")
    return resolved


def load_plan() -> dict[str, Any]:
    plan = read_json(PACKET / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("cpu") == 12, "CPU binding changed")
    require(plan.get("cycles") == 3 and plan.get("rounds") == 3, "cycle or round count changed")
    require(plan.get("cases") == list(CASES), "case order changed")
    require(plan.get("routes") == list(ROUTES), "route order changed")
    require(plan.get("samples") == 50 and plan.get("warmups") == 3, "sample plan changed")
    require(plan.get("processes") == 72, "process total changed")
    require(plan.get("measured_lifecycles") == 3600, "measured lifecycle count changed")
    require(plan.get("comparisons") == [list(pair) for pair in COMPARISONS], "comparison plan changed")
    require(plan.get("observer_flag_percent") == 5, "observer flag threshold changed")
    schedule = plan.get("schedule")
    require(isinstance(schedule, list) and len(schedule) == 72, "schedule is not the fixed 72-process matrix")
    seen: list[tuple[Any, ...]] = []
    for index, row in enumerate(schedule):
        require(isinstance(row, dict), f"schedule row {index} is not an object")
        require(set(row) == {"cycle", "repeat", "case", "route"}, f"schedule row {index} shape changed")
        require(row["cycle"] in range(3), f"schedule row {index} cycle changed")
        require(row["repeat"] in range(3), f"schedule row {index} repeat changed")
        require(row["case"] in CASES and row["route"] in ROUTES, f"schedule row {index} identity changed")
        seen.append((row["cycle"], row["repeat"], row["case"], row["route"]))
    expected = {
        (cycle, repeat, case, route)
        for cycle in range(3)
        for repeat in range(3)
        for case in CASES
        for route in ROUTES
    }
    require(len(seen) == len(set(seen)) == len(expected), "schedule contains duplicate or missing process cells")
    require(set(seen) == expected, "schedule does not cover all cycle/round/case/route cells")
    return plan


def load_cases(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    value = read_json(PACKET / "cases.json")
    require(isinstance(value, list) and [row.get("case") for row in value] == list(CASES), "case receipt changed")
    result: dict[str, dict[str, Any]] = {}
    for row in value:
        require(isinstance(row, dict), "case receipt row is not an object")
        case = row.get("case")
        require(case in CASES and case not in result, f"invalid case identity: {case!r}")
        path = row.get("path")
        require(isinstance(path, str), f"{case}: fixture path is missing")
        relative = Path(path)
        require(not relative.is_absolute() and ".." not in relative.parts, f"{case}: fixture path escapes root")
        fixture = ROOT / relative
        require(fixture.is_file() and not fixture.is_symlink(), f"{case}: fixture is missing")
        require(row.get("format") == "doc", f"{case}: format changed")
        require(row.get("bytes") == fixture.stat().st_size, f"{case}: fixture size changed")
        require(row.get("sha256") == sha(fixture), f"{case}: fixture digest changed")
        require("different-length" in str(row.get("edit", "")), f"{case}: edit contract changed")
        result[case] = row
    return result


def load_contract(cases: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    contract = read_json(PACKET / "oracle-contract.json")
    prior = read_json(PACKET / "prior-oracle-contract.json")
    require(isinstance(contract, dict) and set(contract) == set(cases), "oracle contract case set changed")
    require(isinstance(prior, dict) and set(prior) == set(cases), "prior oracle contract case set changed")
    required = {
        "directory_metadata_fields",
        "allocation_ownership_contract",
        "semantic_witness",
        "control_names",
        "identity",
        "headers",
    }
    for case in CASES:
        value = contract[case]
        old = prior[case]
        require(isinstance(value, dict) and required <= set(value), f"{case}: oracle contract schema changed")
        for key in ("directory_metadata_fields", "semantic_witness", "identity"):
            require(value.get(key) == old.get(key), f"{case}: inherited oracle identity changed at {key}")
        require(isinstance(value["directory_metadata_fields"], list) and value["directory_metadata_fields"],
                f"{case}: directory metadata contract is empty")
        require(isinstance(value["allocation_ownership_contract"], str) and value["allocation_ownership_contract"],
                f"{case}: ownership contract is empty")
        require(isinstance(value["semantic_witness"], dict), f"{case}: semantic witness is not an object")
        require(isinstance(value["control_names"], list) and value["control_names"], f"{case}: controls are empty")
        require(isinstance(value["headers"], dict) and value["headers"], f"{case}: header contract is empty")
        for key, header in value["headers"].items():
            require(isinstance(key, str) and isinstance(header, str) and header, f"{case}: invalid header contract")
    return contract


def verify_map(value: Any, base: Path, label: str) -> dict[str, str]:
    require(isinstance(value, dict) and value, f"{label} custody map is empty")
    result: dict[str, str] = {}
    for raw, expected in value.items():
        require(isinstance(raw, str), f"{label} path is not a string")
        relative = Path(raw)
        require(not relative.is_absolute() and ".." not in relative.parts, f"{label} path escapes root: {raw}")
        path = base / relative
        require(path.is_file() and not path.is_symlink(), f"{label} file is missing: {path}")
        require(sha(path) == digest(expected, f"{label} {raw}"), f"{label} changed: {raw}")
        result[raw] = expected
    return result


def load_builds() -> dict[str, Any]:
    builds = read_json(PACKET / "builds.json")
    require(isinstance(builds, dict), "builds receipt is not an object")
    binaries = builds.get("binaries")
    require(isinstance(binaries, list) and len(binaries) == 1, "binary receipt count changed")
    row = binaries[0]
    require(Path(str(row.get("path", ""))).name == "doc_phase_probe", "probe binary name changed")
    digest(row.get("sha256"), "probe binary")
    integer(row.get("bytes"), "probe binary bytes", 1)
    verify_map(builds.get("source_sha256"), ROOT, "source")
    verify_map(builds.get("probe_sha256"), PACKET, "probe")
    return builds


def cleanup_witnesses() -> dict[str, dict[str, Any]]:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return {}
    value = read_json(path)
    require(isinstance(value, dict), "cleanup receipt is not an object")
    result: dict[str, dict[str, Any]] = {}
    for row in value.get("identities", []):
        require(isinstance(row, dict), "cleanup identity is not an object")
        raw = row.get("path")
        require(isinstance(raw, str), "cleanup identity path is missing")
        result[str(Path(raw).resolve())] = row
    return result


def verify_binary(row: dict[str, Any], witnesses: dict[str, dict[str, Any]]) -> None:
    path = Path(row["path"])
    if path.is_file():
        require(not path.is_symlink() and sha(path) == row["sha256"] and path.stat().st_size == row["bytes"],
                f"live binary identity changed: {path}")
        return
    witness = witnesses.get(str(path.resolve()))
    require(witness is not None, f"missing binary has no cleanup witness: {path}")
    require(witness == row, f"cleanup witness differs for {path}")


def verify_freeze(builds: dict[str, Any]) -> dict[str, Any]:
    frozen = read_json(PACKET / "freeze.json")
    require(isinstance(frozen, dict), "freeze is not an object")
    bindings = frozen.get("bindings")
    require(isinstance(bindings, dict) and bindings, "freeze bindings are empty")
    require(REQUIRED_FREEZE_FILES <= set(bindings), "freeze omits required custody files")
    for raw, expected in bindings.items():
        require(isinstance(raw, str) and not Path(raw).is_absolute() and ".." not in Path(raw).parts,
                f"freeze binding path is unsafe: {raw}")
        path = ROOT / raw
        require(path.is_file() and not path.is_symlink(), f"freeze binding is missing: {raw}")
        require(sha(path) == digest(expected, f"freeze binding {raw}"), f"freeze binding changed: {raw}")
    binaries = frozen.get("binaries")
    require(binaries == builds.get("binaries"), "freeze binary identity differs from builds receipt")
    witnesses = cleanup_witnesses()
    for row in binaries:
        verify_binary(row, witnesses)
    return frozen


def oracle_guard(oracle: Any, witness: dict[str, Any], label: str) -> None:
    require(isinstance(oracle, dict), f"{label}: oracle is not an object")
    require(oracle.get("failure_reasons") == [], f"{label}: oracle has failures")
    require(oracle.get("semantic_witness") == witness, f"{label}: semantic witness changed")

    def booleans(item: Any, path: str) -> None:
        if isinstance(item, dict):
            for key, value in item.items():
                if isinstance(value, bool):
                    require(value, f"{path}.{key} is false")
                else:
                    booleans(value, f"{path}.{key}")
        elif isinstance(item, list):
            for index, value in enumerate(item):
                booleans(value, f"{path}[{index}]")

    booleans(oracle, label)
    raw = oracle.get("raw_directory")
    require(isinstance(raw, dict), f"{label}: raw directory oracle missing")
    for key in ("source_expected_difference_bytes", "expected_output_difference_bytes", "source_output_difference_bytes"):
        require(raw.get(key) == 0, f"{label}: raw directory {key} is nonzero")
    require(oracle.get("directory_metadata_differences") == [], f"{label}: directory metadata differs")


def stats(values: list[int | float]) -> dict[str, int | float]:
    require(values, "empty independent timing vector")
    ordered = sorted(values)
    return {
        "n": len(values),
        "p50": statistics.median(ordered),
        "mean": statistics.mean(ordered),
        "p95": ordered[math.ceil(0.95 * len(ordered)) - 1],
        "p99": ordered[math.ceil(0.99 * len(ordered)) - 1],
        "maximum": ordered[-1],
    }


def close_json(actual: Any, expected: Any, path: str = "analysis") -> None:
    if isinstance(actual, bool) or isinstance(expected, bool):
        require(actual is expected, f"{path}: boolean differs")
    elif isinstance(actual, dict) or isinstance(expected, dict):
        require(isinstance(actual, dict) and isinstance(expected, dict), f"{path}: object shape differs")
        require(set(actual) == set(expected), f"{path}: object keys differ")
        for key in expected:
            close_json(actual[key], expected[key], f"{path}.{key}")
    elif isinstance(actual, list) or isinstance(expected, list):
        require(isinstance(actual, list) and isinstance(expected, list) and len(actual) == len(expected),
                f"{path}: list shape differs")
        for index, (left, right) in enumerate(zip(actual, expected)):
            close_json(left, right, f"{path}[{index}]")
    elif isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
        require(math.isclose(float(actual), float(expected), rel_tol=REL_TOL, abs_tol=ABS_TOL),
                f"{path}: {actual!r} differs from {expected!r}")
    else:
        require(actual == expected, f"{path}: value differs")


def check_trace(trace: Any, phases: tuple[str, ...], whole: int, label: str, owner: str,
                split_owner: int) -> tuple[dict[str, int], dict[str, float], dict[str, float]]:
    require(isinstance(trace, dict), f"{label}: diagnostic trace is not an object")
    require(trace.get("event_count") == len(phases) * 2, f"{label}: event count changed")
    require(trace.get("overflow") is False and trace.get("balanced") is True
            and trace.get("sequence_ok") is True, f"{label}: event flags failed")
    require(trace.get("expected_phases") == list(phases), f"{label}: expected event phases changed")
    events = trace.get("events")
    spans = trace.get("spans")
    require(isinstance(events, list) and len(events) == len(phases) * 2, f"{label}: event list changed")
    require(isinstance(spans, list) and len(spans) == len(phases), f"{label}: span list changed")
    require(split_owner > 0, f"{label}: owner interval is zero")
    times: dict[str, int] = {}
    whole_percent: dict[str, float] = {}
    parent_percent: dict[str, float] = {}
    previous = 0
    total = 0
    for index, phase in enumerate(phases):
        started, finished = events[index * 2:index * 2 + 2]
        require(isinstance(started, dict) and isinstance(finished, dict), f"{label}: event is not an object")
        require((started.get("kind"), started.get("phase"), started.get("outcome"))
                == ("started", phase, "started"), f"{label}: start event changed")
        require((finished.get("kind"), finished.get("phase"), finished.get("outcome"))
                == ("finished", phase, "success"), f"{label}: finish event changed")
        start_ns = integer(started.get("t_ns"), f"{label}: start timestamp")
        finish_ns = integer(finished.get("t_ns"), f"{label}: finish timestamp")
        require(previous <= start_ns <= finish_ns <= whole, f"{label}: event timestamp is outside owner whole")
        previous = finish_ns
        duration = finish_ns - start_ns
        span = spans[index]
        require(span == {
            "phase": phase,
            "outcome": "success",
            "start_ns": start_ns,
            "finish_ns": finish_ns,
            "duration_ns": duration,
        }, f"{label}: span duration or identity changed")
        # Keep the same owner-qualified keys as analyze.py while using the
        # process label only for diagnostics in failure messages.
        key = f"{owner}.{phase}"
        times[key] = duration
        whole_percent[key] = duration / whole * 100.0
        parent_percent[key] = duration / split_owner * 100.0
        total += duration
    require(total <= split_owner, f"{label}: nested event spans exceed owner interval")
    return times, whole_percent, parent_percent


def sample_times(sample: dict[str, Any], route: str, label: str) -> tuple[dict[str, int], dict[str, float], dict[str, float]]:
    whole = integer(sample.get("whole_ns"), f"{label}: whole_ns", 1)
    times: dict[str, int] = {"whole_ns": whole}
    whole_percent: dict[str, float] = {}
    parent_percent: dict[str, float] = {}
    split = sample.get("split")
    if route == "ordinary-opaque":
        require(split is None, f"{label}: opaque route has split timing")
        if "phase_ns" in sample:
            require(sample.get("phase_ns") == {"whole_ns": whole}, f"{label}: opaque compatibility timing changed")
        if "public_phase_ns" in sample:
            require(sample.get("public_phase_ns") is None, f"{label}: opaque public split alias changed")
    else:
        require(isinstance(split, dict), f"{label}: split timing is missing")
        require(set(split) == set(OUTER) | {"split_sum_ns", "whole_residual_ns"}, f"{label}: split fields changed")
        for key in OUTER + ("split_sum_ns",):
            integer(split.get(key), f"{label}: {key}")
        # The five windows are sequentially sampled inside the monotonic outer
        # interval.  A negative residual would therefore be a timing-boundary
        # corruption, rather than an attributable phase; keep it unqualified.
        residual = integer(split.get("whole_residual_ns"), f"{label}: whole_residual_ns")
        require(split["split_sum_ns"] == sum(split[key] for key in OUTER), f"{label}: split sum changed")
        require(residual == whole - split["split_sum_ns"], f"{label}: outer residual changed")
        if "public_phase_ns" in sample:
            require(sample.get("public_phase_ns") == split, f"{label}: public split projection changed")
        if "phase_ns" in sample:
            expected_phase = {
                "open_ns": split["open_ns"],
                "stage_ns": split["edit_ns"] + split["replace_ns"],
                "finish_ns": split["commit_ns"],
                "whole_ns": whole,
            }
            require(sample.get("phase_ns") == expected_phase, f"{label}: compatibility phase projection changed")
        times.update(split)
        for key in OUTER + ("whole_residual_ns",):
            whole_percent[key] = split[key] / whole * 100.0

    diagnostics = sample.get("diagnostics")
    observer = sample.get("observer_clock_control_ns")
    if route != "profiled-clock":
        require(diagnostics is None and observer is None, f"{label}: unprofiled route carries diagnostics")
    else:
        require(isinstance(diagnostics, dict) and set(diagnostics) == {"open", "commit"},
                f"{label}: profiled diagnostic owners changed")
        observer_ns = integer(observer, f"{label}: observer clock control")
        times["observer_clock_control_ns"] = observer_ns
        open_times, open_whole, open_parent = check_trace(
            diagnostics["open"], OPEN_PHASES, whole, f"{label}.open", "open", split["open_ns"]
        )
        commit_times, commit_whole, commit_parent = check_trace(
            diagnostics["commit"], COMMIT_PHASES, whole, f"{label}.commit", "commit", split["commit_ns"]
        )
        times.update(open_times)
        times.update(commit_times)
        whole_percent.update(open_whole)
        whole_percent.update(commit_whole)
        parent_percent.update(open_parent)
        parent_percent.update(commit_parent)
    return times, whole_percent, parent_percent


def check_report_header(report: dict[str, Any], row: dict[str, Any], case: dict[str, Any],
                        plan: dict[str, Any], contract: dict[str, dict[str, Any]]) -> None:
    label = f"{row['cycle']}/{row['repeat']}/{row['case']}/{row['route']}"
    require(report.get("schema_version") == 1, f"{label}: schema changed")
    expected = {
        "case": row["case"],
        "format": "doc",
        "operation": "format",
        "route": row["route"],
        "policy": "reuse",
        "policy_applied": False,
        "policy_application_scope": "not_applied_public_format_route",
        "input": case["path"],
        "samples_requested": plan["samples"],
        "warmups": plan["warmups"],
        "timing_claim": True,
        "allocator_instrumented": False,
        "text_utf16_units": 45,
        "scope": "public_doc_open_edit_replace_commit_output_copy",
        "mode": "doc_public_phase_attribution",
    }
    for key, value in expected.items():
        require(report.get(key) == value, f"{label}: header {key} changed")
    for key, value in contract[row["case"]]["headers"].items():
        require(report.get(key) == value, f"{label}: contract header {key} changed")
    require(report.get("source_sha256") == case["sha256"], f"{label}: source digest changed")
    source_inventory = report.get("source_inventory")
    require(isinstance(source_inventory, dict), f"{label}: source inventory is missing")
    require(source_inventory.get("file_bytes") == case["bytes"], f"{label}: source size changed")
    digest(report.get("expected_output_sha256"), f"{label}: expected output digest")
    digest(report.get("replacements_sha256"), f"{label}: replacements digest")
    require(report.get("directory_metadata_fields") == contract[row["case"]]["directory_metadata_fields"],
            f"{label}: directory metadata fields changed")
    require(report.get("allocation_ownership_contract") == contract[row["case"]]["allocation_ownership_contract"],
            f"{label}: ownership contract changed")


def validate_report(report: Any, row: dict[str, Any], case: dict[str, Any], plan: dict[str, Any],
                    contracts: dict[str, dict[str, Any]], identities: dict[str, dict[str, Any]],
                    outputs: dict[str, str]) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], str]:
    require(isinstance(report, dict), f"{row}: report is not an object")
    check_report_header(report, row, case, plan, contracts)
    label = f"{row['cycle']}/{row['repeat']}/{row['case']}/{row['route']}"
    contract = contracts[row["case"]]
    expected_oracle = report.get("expected_oracle")
    oracle_guard(expected_oracle, contract["semantic_witness"], f"{label}: expected")
    require([item.get("name") for item in report.get("oracle_controls", [])] == contract["control_names"],
            f"{label}: oracle control names changed")
    controls = report.get("oracle_controls")
    require(isinstance(controls, list) and controls, f"{label}: oracle controls are missing")
    for control in controls:
        require(isinstance(control, dict) and control.get("status") == "rejected"
                and control.get("rejected") is True and isinstance(control.get("failure_reasons"), list)
                and control["failure_reasons"], f"{label}: oracle control did not reject")

    proof = report.get("changed_length_proof")
    require(isinstance(proof, dict), f"{label}: length proof is missing")
    require(proof.get("logical_stream_length_change_proven") is True
            and proof.get("format_specific_semantic_length_proven") is True
            and proof.get("any_stream_length_changed") is True, f"{label}: length proof failed")
    identity_keys = ("source_sha256", "expected_output_sha256", "replacements_sha256",
                     "source_inventory", "expected_output_inventory", "replacements", "changed_length_proof")
    identity = {key: report.get(key) for key in identity_keys}
    require(identity == contract["identity"], f"{label}: frozen oracle identity changed")
    if row["case"] in identities:
        require(identity == identities[row["case"]], f"{label}: case identity drifted")
    else:
        identities[row["case"]] = identity

    expected_inventory = report.get("expected_output_inventory")
    require(isinstance(expected_inventory, dict), f"{label}: expected output inventory is missing")
    expected_streams = expected_inventory.get("streams")
    require(isinstance(expected_streams, list) and expected_streams, f"{label}: expected stream inventory missing")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == plan["samples"], f"{label}: sample count changed")
    arrays: dict[str, list[int]] = {}
    whole_percent: dict[str, list[float]] = {}
    parent_percent: dict[str, list[float]] = {}
    output_hash: str | None = None
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("index") == index and sample.get("route") == row["route"],
                f"{label}: sample identity changed")
        oracle_guard(sample.get("oracle"), contract["semantic_witness"], f"{label}: sample {index}")
        sample_hash = digest(sample.get("output_sha256"), f"{label}: sample output hash")
        if output_hash is None:
            output_hash = sample_hash
        require(sample_hash == output_hash, f"{label}: sample output hash drifted")
        if row["case"] in outputs:
            require(sample_hash == outputs[row["case"]], f"{label}: route output hash drifted")
        else:
            outputs[row["case"]] = sample_hash
        require(sample.get("output_inventory") == report["expected_output_inventory"],
                f"{label}: output inventory differs from expected")
        times, fractions, parents = sample_times(sample, row["route"], label)
        for key, value in times.items():
            arrays.setdefault(key, []).append(value)
        for key, value in fractions.items():
            whole_percent.setdefault(key, []).append(value)
        for key, value in parents.items():
            parent_percent.setdefault(key, []).append(value)
        require(sample["oracle"].get("output_stream_bytes_match_expected") is True,
                f"{label}: sample output byte oracle failed")
    require(output_hash is not None, f"{label}: no sample output hash")
    return (
        {key: stats(value) for key, value in arrays.items()},
        {key: stats(value) for key, value in whole_percent.items()},
        {key: stats(value) for key, value in parent_percent.items()},
        output_hash,
    )


def expected_comparisons(plan: dict[str, Any], rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    lookup = {(row["cycle"], row["repeat"], row["case"], row["route"]): row for row in rows}
    require(len(lookup) == len(rows) == 72, "independent process identity is not unique")
    result: list[dict[str, Any]] = []
    for cycle in range(plan["cycles"]):
        for repeat in range(plan["rounds"]):
            for case in plan["cases"]:
                for before, after in COMPARISONS:
                    left = lookup[cycle, repeat, case, before]["timing"]["whole_ns"]
                    right = lookup[cycle, repeat, case, after]["timing"]["whole_ns"]
                    metrics: dict[str, Any] = {}
                    for metric in ("p50", "mean"):
                        before_ns = left[metric]
                        after_ns = right[metric]
                        delta = (after_ns / before_ns - 1.0) * 100.0
                        metrics[metric] = {
                            "before_ns": before_ns,
                            "after_ns": after_ns,
                            "delta_percent": delta,
                            "flag": abs(delta) > plan["observer_flag_percent"],
                        }
                    result.append({
                        "cycle": cycle,
                        "repeat": repeat,
                        "case": case,
                        "before": before,
                        "after": after,
                        "metrics": metrics,
                    })
    return result


def audit_capture(plan: dict[str, Any], cases: dict[str, dict[str, Any]],
                  contracts: dict[str, dict[str, Any]], frozen: dict[str, Any]) -> dict[str, Any]:
    manifest_path = CAPTURES / "manifest.json"
    if not manifest_path.is_file():
        raise Pending(("captures/manifest.json",))
    manifest = read_json(manifest_path)
    require(isinstance(manifest, dict) and manifest.get("status") == "complete", "capture manifest is not complete")
    require(manifest.get("freeze_sha256") == sha(PACKET / "freeze.json"), "capture freeze binding changed")
    require(manifest.get("bindings_start") == frozen["bindings"] == manifest.get("bindings_end"),
            "capture bindings changed during run")
    runs = manifest.get("runs")
    schedule = plan["schedule"]
    require(isinstance(runs, list) and len(runs) == len(schedule) == 72, "capture process count changed")
    identities: dict[str, dict[str, Any]] = {}
    outputs: dict[str, str] = {}
    rows: list[dict[str, Any]] = []
    raw_paths: set[str] = set()
    for run, planned in zip(runs, schedule):
        require(isinstance(run, dict), "capture run is not an object")
        require(all(run.get(key) == value for key, value in planned.items()), f"capture matrix row changed: {planned}")
        require(run.get("exit_code") == 0, f"capture child failed: {planned}")
        case = cases[planned["case"]]
        binary_path = frozen["binaries"][0]["path"]
        expected_name = f"c{planned['cycle']}-{planned['case']}-{planned['route']}-r{planned['repeat']}.json"
        expected_command = [
            "taskset", "-c", str(plan["cpu"]), binary_path,
            "--case", planned["case"], "--input", case["path"], "--route", planned["route"],
            "--samples", str(plan["samples"]), "--warmups", str(plan["warmups"]),
        ]
        require(run.get("command") == expected_command, f"capture command changed: {planned}")
        output = safe_capture_path(run.get("output"), f"{planned}: output")
        stderr = safe_capture_path(run.get("stderr"), f"{planned}: stderr")
        output_rel = output.relative_to(CAPTURES).as_posix()
        stderr_rel = stderr.relative_to(CAPTURES).as_posix()
        require(output_rel == expected_name and stderr_rel == expected_name + ".stderr", f"capture names changed: {planned}")
        require(not ({output_rel, stderr_rel} & raw_paths), f"capture raw path reused: {planned}")
        raw_paths.update((output_rel, stderr_rel))
        require(run.get("sha256") == sha(output), f"capture output hash changed: {planned}")
        require(run.get("stderr_sha256") == sha(stderr), f"capture stderr hash changed: {planned}")
        report = read_json(output)
        timing, whole, parents, output_hash = validate_report(
            report, planned, case, plan, contracts, identities, outputs
        )
        rows.append({
            **planned,
            "output_sha256": output_hash,
            "timing": timing,
            "whole_percent": whole,
            "parent_percent": parents,
        })
    expected_raw = {
        f"c{row['cycle']}-{row['case']}-{row['route']}-r{row['repeat']}.json" for row in schedule
    }
    expected_raw |= {name + ".stderr" for name in expected_raw}
    require(raw_paths == expected_raw, "capture raw path set differs from schedule")
    on_disk = {
        path.relative_to(CAPTURES).as_posix()
        for path in CAPTURES.rglob("*")
        if path.is_file() and path.name != "manifest.json"
    }
    require(on_disk == expected_raw, "capture directory contains unmanifested raw file")

    comparisons = expected_comparisons(plan, rows)
    independent = {
        "disposition": "unchanged source public DOC attribution; no production speedup claim",
        "processes": rows,
        "comparisons": comparisons,
        "case_identities": identities,
    }
    analysis = read_json(PACKET / "analysis.json")
    close_json(analysis, independent)
    return independent


def custody_record(builds: dict[str, Any], frozen: dict[str, Any], independent: dict[str, Any]) -> dict[str, Any]:
    return {
        "plan_sha256": sha(PACKET / "plan.json"),
        "freeze_sha256": sha(PACKET / "freeze.json"),
        "manifest_sha256": sha(CAPTURES / "manifest.json"),
        "analysis_sha256": sha(PACKET / "analysis.json"),
        "oracle_contract_sha256": sha(PACKET / "oracle-contract.json"),
        "prior_oracle_contract_sha256": sha(PACKET / "prior-oracle-contract.json"),
        "frozen_binding_count": len(frozen["bindings"]),
        "source_sha256": builds["source_sha256"],
        "probe_sha256": builds["probe_sha256"],
        "binaries": builds["binaries"],
        "processes": len(independent["processes"]),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--draft", action="store_true", help="return a pending draft result before terminal capture")
    args = parser.parse_args()
    try:
        plan = load_plan()
        cases = load_cases(plan)
        contracts = load_contract(cases)
        builds = load_builds()
        frozen = verify_freeze(builds)
        if args.draft and not (CAPTURES / "manifest.json").is_file():
            print(json.dumps({"status": "DRAFT PASS", "pending": ["captures/manifest.json"]}, sort_keys=True))
            return 0
        independent = audit_capture(plan, cases, contracts, frozen)
        record = {
            "packet": "change-0729",
            "disposition": independent["disposition"],
            "matrix": {"processes": len(independent["processes"]), "cycles": plan["cycles"],
                        "rounds": plan["rounds"], "cases": list(CASES), "routes": list(ROUTES)},
            "statistics": {"quantiles": "nearest-rank p95/p99; midpoint p50",
                            "fields": list(STAT_KEYS), "mean": "statistics.mean",
                            "relative_tolerance": REL_TOL, "absolute_tolerance": ABS_TOL,
                            "outer_timing_fields": list(OUTER),
                            "event_owners": {"open": list(OPEN_PHASES), "commit": list(COMMIT_PHASES)}},
            "custody": custody_record(builds, frozen, independent),
            "processes": independent["processes"],
            "comparisons": independent["comparisons"],
            "case_identities": independent["case_identities"],
            "analysis_match": True,
        }
        (PACKET / "audit.json").write_text(json.dumps(record, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({"status": "PASS", "processes": len(independent["processes"]), "analysis_match": True}, sort_keys=True))
        return 0
    except Pending as pending:
        if args.draft:
            print(json.dumps({"status": "DRAFT PENDING", "pending": list(pending.items)}, sort_keys=True))
            return 0
        print("PENDING: " + "; ".join(pending.items), file=sys.stderr)
        return 2
    except (AuditError, OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
