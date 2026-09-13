#!/usr/bin/env python3
"""Validate 0546 XLSX allocation vectors and planning memory gates.

The allocator lane is an evidence lane for the native post-EOF completion pilot.
It retains every planning, commit, and publication vector and compares the
matched shape/repeat children without turning allocator calls into a latency
claim.  Only ordinary planning allocated bytes and the incremental planning
peak are admission gates: for each shape and repeat, the largest candidate
value must be no greater than 1.01 times the smallest baseline value.  Calls
and reallocations remain diagnostics, and no call-count reduction is required.

``analyze`` is read-only and deterministic.  The command-line entry point
may write the requested JSON result after validation; it never launches a
build or benchmark child.
"""

from __future__ import annotations

import argparse
import json
import statistics
from pathlib import Path
from typing import Any

import analyze as native


HERE = Path(__file__).resolve().parent
PHASES = ("plan", "commit", "publication")
FIELDS = (
    "allocation_calls",
    "reallocation_calls",
    "allocated_bytes",
    "incremental_region_peak_live_bytes",
)
MAX_CANDIDATE_RATIO = 1.01
ALLOCATION_GATES = HERE / "allocation-gates.json"


def _require(condition: bool, message: str) -> None:
    native.BASE.require(condition, message)


def _read_json(path: Path) -> Any:
    return native._read_json(path)


def _finite_number(value: Any, label: str) -> float:
    native.BASE.finite_number(value, label)
    return float(value)


def _allocation_gates(plan: dict[str, Any]) -> dict[str, Any]:
    """Read and validate the two ordinary planning gates.

    The call-count gate used by the previous pilot is deliberately rejected if
    it is reintroduced.  The 0546 gate file is frozen separately from the
    native latency gate object so this analyzer cannot silently change the
    native admission schema while adding an allocation bound.
    """

    gates = _read_json(ALLOCATION_GATES)
    _require(isinstance(gates, dict), "0546 allocation gates are not an object")
    expected = {
        "status",
        "planning_allocated_bytes_max_increase_percent",
        "planning_incremental_peak_live_bytes_max_increase_percent",
        "comparison",
        "scope",
        "created_utc",
    }
    _require(set(gates) == expected,
             "0546 allocation gate inventory contains an unexpected field")
    _require(gates.get("status") in {"frozen-before-first-allocation-capture",
                                      "frozen-before-build-and-capture"},
             "0546 allocation gates are not frozen before capture")
    _require("planning_allocation_calls_reduction_percent" not in gates,
             "0546 must not contain a hard planning allocation call gate")
    _require(gates.get("comparison") ==
             "maximum candidate <= 1.01 * minimum baseline in every shape/repeat",
             "0546 allocation comparison rule differs")
    result: dict[str, Any] = {
        "planning_allocated_bytes_max_increase_percent": _finite_number(
            gates["planning_allocated_bytes_max_increase_percent"],
            "allocation-gates.planning_allocated_bytes_max_increase_percent"),
        "planning_incremental_peak_live_bytes_max_increase_percent": _finite_number(
            gates["planning_incremental_peak_live_bytes_max_increase_percent"],
            "allocation-gates.planning_incremental_peak_live_bytes_max_increase_percent"),
    }
    for key, value in result.items():
        _require(value >= 0.0, f"allocation-gates.{key} is negative")
        _require(value == 1.0, f"allocation-gates.{key} must be 1.0")
    result["planning_allocated_bytes_max_candidate_ratio"] = (
        1.0 + result["planning_allocated_bytes_max_increase_percent"] / 100.0)
    result["planning_incremental_region_peak_max_candidate_ratio"] = (
        1.0 + result["planning_incremental_peak_live_bytes_max_increase_percent"] / 100.0)
    _require(result["planning_allocated_bytes_max_candidate_ratio"] == MAX_CANDIDATE_RATIO,
             "0546 allocated-byte ratio is not 1.01")
    _require(result["planning_incremental_region_peak_max_candidate_ratio"] == MAX_CANDIDATE_RATIO,
             "0546 incremental-peak ratio is not 1.01")
    return result


def _planning_ratio_gate(baseline: list[int], candidate: list[int],
                         ratio: float, label: str) -> dict[str, Any]:
    _require(baseline and candidate, f"{label} has no samples")
    before = min(baseline)
    after = max(candidate)
    _require(before >= 0 and after >= 0, f"{label} has a negative value")
    # A zero baseline is safe only when the candidate remains zero.  Avoid a
    # division by zero while still recording the exact worst pair.
    if before == 0:
        candidate_ratio: float | None = 1.0 if after == 0 else None
        passed = after == 0
    else:
        candidate_ratio = after / before
        passed = candidate_ratio <= ratio
    return {
        "baseline_min": before,
        "candidate_max": after,
        "candidate_to_baseline_ratio": candidate_ratio,
        "max_allowed_ratio": ratio,
        "passed": passed,
    }


def _change_percent(baseline: float, candidate: float) -> float | None:
    if baseline == 0.0:
        return 0.0 if candidate == 0.0 else None
    return (candidate / baseline - 1.0) * 100.0


def _phase_vectors(stage: str, job: dict[str, Any], report: dict[str, Any]) -> dict[str, dict[str, list[int]]]:
    label = f"{stage}/{job['name']}"
    results = report.get("results")
    _require(isinstance(results, list) and len(results) == 1,
             f"{label} must contain one result")
    result = results[0]
    _require(isinstance(result, dict), f"{label} result is not an object")
    source = result.get("source")
    _require(isinstance(source, dict), f"{label} has no source evidence")
    xlsx = source.get("xlsx_cell_values")
    _require(isinstance(xlsx, dict), f"{label} has no XLSX source evidence")
    phases: dict[str, dict[str, list[int]]] = {}
    for phase in PHASES:
        allocation_name = phase + "_allocation_metrics"
        time_name = phase + "_ns"
        samples = xlsx.get(allocation_name)
        elapsed = xlsx.get(time_name)
        _require(isinstance(samples, list), f"{label}.{allocation_name} is not a vector")
        _require(isinstance(elapsed, list) and len(samples) == len(elapsed) == job["samples"],
                 f"{label}.{phase} vector cardinality differs from plan")
        parsed = [native.BASE.allocation_sample(
            sample, f"{label}/{phase}[{index}]", True)
                  for index, sample in enumerate(samples)]
        phases[phase] = {
            field: [sample[field] for sample in parsed]
            for field in FIELDS
        }
    return phases


def _stage_evidence(stage: str, plan: dict[str, Any], jobs: list[dict[str, Any]]) -> dict[str, Any]:
    checked = native._check_stage(stage, plan, jobs, True)
    rows = {row["name"]: row for row in checked["allocation"]["rows"]}
    _require(set(rows) == {job["name"] for job in jobs},
             f"{stage} allocation row set differs from plan")
    evidence: dict[str, Any] = {}
    for job in jobs:
        name = job["name"]
        report = _read_json(HERE / stage / f"{name}.json")
        phases = _phase_vectors(stage, job, report)
        evidence[name] = {
            "case": job["case"],
            "shape": job["shape"],
            "repeat": job["repeat"],
            "samples": job["samples"],
            "phases": phases,
            "identity": rows[name]["identity"],
        }
    return {"stage": stage, "manifest_sha256": checked["manifest_sha256"],
            "rows": evidence, "custody": checked["custody"]}


def _pending_stage(stage: str, plan: dict[str, Any],
                   jobs: list[dict[str, Any]]) -> dict[str, Any] | None:
    """Return a pending record when capture artifacts have not arrived yet.

    The existence check is intentionally shallow.  Once every expected path
    exists, ``_stage_evidence`` performs the strict schema, hash, and custody
    validation and reports malformed evidence as an error rather than hiding
    it as pending.
    """

    missing = native._missing_stage_artifacts(stage, plan, True)
    if not missing:
        return None
    return {
        "status": "pending",
        "stage": stage,
        "lane": "allocation",
        "expected_row_count": len(jobs),
        "missing_artifacts": missing,
    }


def _metric_record(baseline: list[int], candidate: list[int]) -> dict[str, Any]:
    baseline_p50 = statistics.median(baseline)
    candidate_p50 = statistics.median(candidate)
    return {
        "baseline": list(baseline),
        "candidate": list(candidate),
        "baseline_p50": baseline_p50,
        "candidate_p50": candidate_p50,
        "change_percent": _change_percent(float(baseline_p50), float(candidate_p50)),
    }


def _compare(plan: dict[str, Any], baseline: dict[str, Any],
             candidate: dict[str, Any], jobs: list[dict[str, Any]]) -> dict[str, Any]:
    gates = _allocation_gates(plan)
    rows: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    decisions: list[dict[str, Any]] = []
    for job in jobs:
        name = job["name"]
        left = baseline["rows"][name]
        right = candidate["rows"][name]
        _require(left["identity"] == right["identity"],
                 f"baseline/candidate allocation identity differs for {name}")
        for phase in PHASES:
            metrics: dict[str, Any] = {}
            for field in FIELDS:
                value = _metric_record(left["phases"][phase][field],
                                       right["phases"][phase][field])
                metrics[field] = value
                change = value["change_percent"]
                if change is None or change > native.ADVERSE_THRESHOLD_PERCENT:
                    adverse.append({
                        "name": name,
                        "shape": job["shape"],
                        "repeat": job["repeat"],
                        "phase": phase,
                        "field": field,
                        "threshold_percent": native.ADVERSE_THRESHOLD_PERCENT,
                        **value,
                    })
            rows.append({
                "name": name,
                "shape": job["shape"],
                "repeat": job["repeat"],
                "phase": phase,
                "metrics": metrics,
            })

        planning = {
            "allocated_bytes": _planning_ratio_gate(
                left["phases"]["plan"]["allocated_bytes"],
                right["phases"]["plan"]["allocated_bytes"],
                gates["planning_allocated_bytes_max_candidate_ratio"],
                f"{name}/plan/allocated_bytes"),
            "incremental_region_peak_live_bytes": _planning_ratio_gate(
                left["phases"]["plan"]["incremental_region_peak_live_bytes"],
                right["phases"]["plan"]["incremental_region_peak_live_bytes"],
                gates["planning_incremental_region_peak_max_candidate_ratio"],
                f"{name}/plan/incremental_region_peak_live_bytes"),
        }
        decisions.append({
            "name": name,
            "shape": job["shape"],
            "repeat": job["repeat"],
            "planning": planning,
            "calls_diagnostic": _metric_record(
                left["phases"]["plan"]["allocation_calls"],
                right["phases"]["plan"]["allocation_calls"]),
            "reallocations_diagnostic": _metric_record(
                left["phases"]["plan"]["reallocation_calls"],
                right["phases"]["plan"]["reallocation_calls"]),
            "passed": all(item["passed"] for item in planning.values()),
        })
    return {
        "rows": rows,
        "decisions": decisions,
        "planning_gate_passed": all(item["passed"] for item in decisions),
        "adverse_flags": adverse,
        "gate_values": gates,
    }


def analyze() -> dict[str, Any]:
    """Validate complete baseline/candidate allocation evidence read-only."""

    plan = _read_json(HERE / "plan.json")
    _require(isinstance(plan, dict), "plan is not an object")
    primary = plan.get("primary")
    _require(isinstance(primary, dict)
             and primary.get("case") == native.PRIMARY_CASE,
             "0546 primary case is not the expected XLSX source-backed case")
    plan_gates = plan.get("gates")
    _require(isinstance(plan_gates, dict), "plan.gates is not an object")
    _require("planning_allocation_calls_reduction_percent" not in plan_gates,
             "0546 plan must not contain a hard planning allocation call gate")
    # Validate the separate allocation gate inventory before loading any
    # allocator report.  This also rejects the legacy call-count gate.
    if not ALLOCATION_GATES.is_file():
        return {
            "status": "pending",
            "stage": "compare",
            "plan_sha256": native._sha(HERE / "plan.json"),
            "missing_artifacts": ["allocation-gates.json"],
            "scope": "0546 XLSX allocation evidence is incomplete",
        }
    _allocation_gates(plan)
    jobs = native._expected_allocation_jobs(plan)
    pending = {
        stage: _pending_stage(stage, plan, jobs)
        for stage in ("baseline", "candidate")
    }
    if any(value is not None for value in pending.values()):
        return {
            "status": "pending",
            "stage": "compare",
            "plan_sha256": native._sha(HERE / "plan.json"),
            "allocation_gates_sha256": native._sha(ALLOCATION_GATES),
            "baseline": pending["baseline"],
            "candidate": pending["candidate"],
            "scope": "0546 XLSX allocation evidence is incomplete",
        }
    baseline = _stage_evidence("baseline", plan, jobs)
    candidate = _stage_evidence("candidate", plan, jobs)
    comparison = _compare(plan, baseline, candidate, jobs)
    return {
        "status": "pass",
        "stage": "compare",
        "plan_sha256": native._sha(HERE / "plan.json"),
        "allocation_gates_sha256": native._sha(ALLOCATION_GATES),
        "baseline": baseline,
        "candidate": candidate,
        **comparison,
        "scope": (
            "0546 XLSX planning/commit/publication allocation vectors; ordinary "
            "planning allocated bytes and incremental peak gates only; calls and "
            "reallocations are diagnostics; no latency interpretation"
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path,
                        default=HERE / "allocation-analysis.json")
    args = parser.parse_args()
    try:
        result = analyze()
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n",
                               encoding="utf-8")
    except native.BASE.EvidenceError as error:
        print(f"evidence check failed: {error}")
        return 1
    print(f"0546 allocation evidence {result['status']}: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
