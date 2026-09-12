"""Validate and compare the 0527 XLSX row-reconstruction evidence.

The report and custody checks are inherited from the retained 0521 verifier.
This adapter adds the structured 0527 pilot admission rules: every primary
shape/repeat must meet the total p50, total mean, and commit p50 reductions,
and every allocator shape/repeat must meet the allocation-call reduction.
Allocator reallocation calls remain an independently reported metric and are
never substituted for allocation calls in the gate.

The conditional profile, hardware, and eager lanes are intentionally not
treated as passing here.  They are eligible only after the pilot gates pass
and remain ``unmeasured`` until their own bound evidence is analyzed.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
HELPER = HERE.parent / "change-0521" / "analyze.py"
_helper_spec = importlib.util.spec_from_file_location("xlsx_0521_numerical_0527", HELPER)
if _helper_spec is None or _helper_spec.loader is None:
    raise ImportError(f"cannot load numerical helper: {HELPER}")
BASE = importlib.util.module_from_spec(_helper_spec)
_helper_spec.loader.exec_module(BASE)

# The retained verifier resolves all artifacts relative to these module
# globals.  Keep those bindings local to the 0527 result bundle.
BASE.HERE = HERE
BASE.PLAN = HERE / "plan.json"
BASE.BOOTSTRAP_SEED = 5270527


def __getattr__(name: str) -> Any:
    return getattr(BASE, name)


GATE_FIELDS = (
    "total_p50_reduction_percent",
    "commit_p50_reduction_percent",
    "total_mean_reduction_percent",
    "allocation_calls_reduction_percent",
    "commit_ir_reduction_percent",
    "require_every_shape_repeat",
)
CONDITIONAL_LANES = ("profile", "hardware", "eager")


def _validated_gates(plan: dict[str, Any]) -> dict[str, Any]:
    gates = plan.get("gates")
    BASE.require(isinstance(gates, dict), "plan structured gates are missing")
    BASE.require(set(gates) == set(GATE_FIELDS),
                 "plan structured gate inventory differs")
    for field in GATE_FIELDS[:-1]:
        value = gates.get(field)
        BASE.finite_number(value, f"plan.gates.{field}")
        BASE.require(float(value) >= 0.0,
                     f"plan.gates.{field} is negative")
    BASE.require(gates.get("require_every_shape_repeat") is True,
                 "plan.gates.require_every_shape_repeat must be true")
    return gates


def _admission_thresholds(plan: dict[str, Any]) -> dict[str, Any]:
    """Expose the frozen structured gates using the retained report names."""

    gates = _validated_gates(plan)
    return {
        "native_primary_total_p50_improvement_percent": float(
            gates["total_p50_reduction_percent"]),
        "native_primary_total_mean_improvement_percent": float(
            gates["total_mean_reduction_percent"]),
        "native_primary_commit_p50_improvement_percent": float(
            gates["commit_p50_reduction_percent"]),
        "allocation_calls_reduction_percent": float(
            gates["allocation_calls_reduction_percent"]),
        "profile_commit_ir_reduction_percent": float(
            gates["commit_ir_reduction_percent"]),
        "source": "plan.gates",
    }


def _read_plan() -> dict[str, Any]:
    plan = BASE.read_json(BASE.PLAN)
    BASE.require(isinstance(plan, dict), "plan is not an object")
    _validated_gates(plan)
    return plan


def _nonnegative_number(value: Any, label: str) -> None:
    BASE.finite_number(value, label)
    BASE.require(float(value) >= 0.0, f"{label} is negative")


def _reduction_percent(baseline: Any, candidate: Any, label: str) -> float:
    """Return the candidate reduction relative to a positive baseline."""

    _nonnegative_number(baseline, f"{label}.baseline")
    _nonnegative_number(candidate, f"{label}.candidate")
    BASE.require(float(baseline) > 0.0, f"{label}.baseline must be positive")
    return (float(baseline) - float(candidate)) / float(baseline) * 100.0


def _primary_rows(evidence: dict[str, Any], plan: dict[str, Any], label: str) -> dict[tuple[int, str], dict[str, Any]]:
    """Return the exact primary matrix, rejecting omissions and duplicates."""

    native = evidence.get("native")
    BASE.require(isinstance(native, dict), f"{label}.native is not an object")
    raw_rows = native.get("rows")
    BASE.require(isinstance(raw_rows, list), f"{label}.native.rows is not a list")
    rows = [row for row in raw_rows
            if isinstance(row, dict)
            and row.get("kind") == "primary"
            and row.get("guard") is None]
    primary = plan.get("primary")
    BASE.require(isinstance(primary, dict), "plan primary lane is missing")
    expected = {(repeat, shape)
                for repeat in range(1, int(primary["repeats"]) + 1)
                for shape in primary["shapes"]}
    actual = [(row.get("repeat"), row.get("shape")) for row in rows]
    for repeat, shape in actual:
        BASE.require(isinstance(repeat, int) and not isinstance(repeat, bool),
                     f"{label} primary row repeat is not an integer")
        BASE.require(isinstance(shape, str) and shape,
                     f"{label} primary row shape is not a non-empty string")
    BASE.require(len(rows) == len(expected),
                 f"{label} primary native row count differs from plan: "
                 f"{len(rows)} != {len(expected)}")
    BASE.require(len(set(actual)) == len(actual),
                 f"{label} primary native rows contain duplicate keys")
    actual_set = set(actual)
    BASE.require(actual_set == expected,
                 f"{label} primary native matrix differs from plan: "
                 f"{sorted(actual_set ^ expected)}")
    return {(row["repeat"], row["shape"]): row for row in rows}


def _allocation_rows(evidence: dict[str, Any], plan: dict[str, Any], label: str) -> dict[tuple[int, str], dict[str, Any]]:
    """Return the exact allocator matrix and reject duplicate rows."""

    allocation = evidence.get("allocation")
    BASE.require(isinstance(allocation, dict),
                 f"{label}.allocation is not an object")
    raw_rows = allocation.get("rows")
    BASE.require(isinstance(raw_rows, list),
                 f"{label}.allocation.rows is not a list")
    rows = [row for row in raw_rows
            if isinstance(row, dict) and row.get("kind") == "allocation"]
    config = plan.get("allocation")
    primary = plan.get("primary")
    BASE.require(isinstance(config, dict), "plan allocation lane is missing")
    BASE.require(isinstance(primary, dict), "plan primary lane is missing")
    expected = {(repeat, shape)
                for repeat in range(1, int(config["repeats"]) + 1)
                for shape in config["shapes"]}
    actual = [(row.get("repeat"), row.get("shape")) for row in rows]
    for repeat, shape in actual:
        BASE.require(isinstance(repeat, int) and not isinstance(repeat, bool),
                     f"{label} allocation row repeat is not an integer")
        BASE.require(isinstance(shape, str) and shape,
                     f"{label} allocation row shape is not a non-empty string")
    BASE.require(len(rows) == len(expected),
                 f"{label} allocation row count differs from plan: "
                 f"{len(rows)} != {len(expected)}")
    BASE.require(len(set(actual)) == len(actual),
                 f"{label} allocation rows contain duplicate keys")
    actual_set = set(actual)
    BASE.require(actual_set == expected,
                 f"{label} allocation matrix differs from plan: "
                 f"{sorted(actual_set ^ expected)}")
    expected_case = primary.get("case")
    for row in rows:
        BASE.require(row.get("case") == expected_case,
                     f"{label} allocation row case differs from primary")
    return {(row["repeat"], row["shape"]): row for row in rows}


def _timing(row: dict[str, Any], key: str, label: str) -> dict[str, Any]:
    timing = row.get("timing")
    BASE.require(isinstance(timing, dict), f"{label}.timing is not an object")
    value = timing.get(key)
    BASE.require(isinstance(value, dict), f"{label}.timing.{key} is not an object")
    for stat in ("p50", "mean"):
        BASE.require(stat in value, f"{label}.timing.{key}.{stat} is missing")
        _nonnegative_number(value[stat], f"{label}.timing.{key}.{stat}")
    return value


def _native_admission(result: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    """Evaluate the three native pilot gates for every planned pair."""

    thresholds = _admission_thresholds(plan)
    baseline = _primary_rows(result["baseline"], plan, "baseline")
    candidate = _primary_rows(result["candidate"], plan, "candidate")
    rows: list[dict[str, Any]] = []
    for repeat, shape in sorted(baseline):
        left = baseline[(repeat, shape)]
        right = candidate[(repeat, shape)]
        left_elapsed = _timing(left, "elapsed_ns", f"primary {repeat}/{shape} baseline")
        right_elapsed = _timing(right, "elapsed_ns", f"primary {repeat}/{shape} candidate")
        left_commit = _timing(left, "commit_ns", f"primary {repeat}/{shape} baseline")
        right_commit = _timing(right, "commit_ns", f"primary {repeat}/{shape} candidate")
        total_p50 = _reduction_percent(
            left_elapsed["p50"], right_elapsed["p50"],
            f"primary {repeat}/{shape} total p50")
        total_mean = _reduction_percent(
            left_elapsed["mean"], right_elapsed["mean"],
            f"primary {repeat}/{shape} total mean")
        commit_p50 = _reduction_percent(
            left_commit["p50"], right_commit["p50"],
            f"primary {repeat}/{shape} commit p50")
        checks = {
            "native_primary_total_p50": {
                "baseline": left_elapsed["p50"],
                "candidate": right_elapsed["p50"],
                "reduction_percent": total_p50,
                "required_reduction_percent": thresholds[
                    "native_primary_total_p50_improvement_percent"],
                "passed": total_p50 >= thresholds[
                    "native_primary_total_p50_improvement_percent"],
            },
            "native_primary_total_mean": {
                "baseline": left_elapsed["mean"],
                "candidate": right_elapsed["mean"],
                "reduction_percent": total_mean,
                "required_reduction_percent": thresholds[
                    "native_primary_total_mean_improvement_percent"],
                "passed": total_mean >= thresholds[
                    "native_primary_total_mean_improvement_percent"],
            },
            "native_primary_commit_p50": {
                "baseline": left_commit["p50"],
                "candidate": right_commit["p50"],
                "reduction_percent": commit_p50,
                "required_reduction_percent": thresholds[
                    "native_primary_commit_p50_improvement_percent"],
                "passed": commit_p50 >= thresholds[
                    "native_primary_commit_p50_improvement_percent"],
            },
        }
        rows.append({
            "repeat": repeat,
            "shape": shape,
            **checks,
            "passed": all(check["passed"] for check in checks.values()),
        })
    return {
        "rows": rows,
        "passed": all(row["passed"] for row in rows),
        "scope": (
            "Every planned primary shape/repeat must meet total elapsed p50 "
            "and mean reductions plus commit p50 reduction."
        ),
    }


def _numeric_summary(values: list[Any], label: str) -> dict[str, Any]:
    BASE.require(isinstance(values, list) and values,
                 f"{label} is empty")
    for index, value in enumerate(values):
        BASE.nonnegative_integer(value, f"{label}[{index}]")
    ordered = sorted(values)
    middle = (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) / 2.0
    return {"samples": ordered, "p50": middle,
            "mean": sum(ordered) / len(ordered)}


def _allocation_admission(result: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    """Evaluate allocation-call reduction while retaining realloc separately."""

    thresholds = _admission_thresholds(plan)
    baseline = _allocation_rows(result["baseline"], plan, "baseline")
    candidate = _allocation_rows(result["candidate"], plan, "candidate")
    rows: list[dict[str, Any]] = []
    for repeat, shape in sorted(baseline):
        left_samples = baseline[(repeat, shape)].get("allocation")
        right_samples = candidate[(repeat, shape)].get("allocation")
        BASE.require(isinstance(left_samples, list)
                     and isinstance(right_samples, list),
                     f"allocator {repeat}/{shape} samples are not lists")
        BASE.require(len(left_samples) == len(right_samples) and left_samples,
                     f"allocator {repeat}/{shape} sample counts differ or are empty")
        left_calls: list[Any] = []
        right_calls: list[Any] = []
        left_realloc: list[Any] = []
        right_realloc: list[Any] = []
        for index, (left, right) in enumerate(zip(left_samples, right_samples)):
            BASE.require(isinstance(left, dict) and isinstance(right, dict),
                         f"allocator {repeat}/{shape} sample {index} is not an object")
            BASE.require(left.get("status") == right.get("status") == "measured",
                         f"allocator {repeat}/{shape} sample {index} is not measured")
            for sample, calls, realloc, side in (
                (left, left_calls, left_realloc, "baseline"),
                (right, right_calls, right_realloc, "candidate"),
            ):
                BASE.nonnegative_integer(
                    sample.get("allocation_calls"),
                    f"allocator {repeat}/{shape} {side} sample {index}.allocation_calls")
                BASE.nonnegative_integer(
                    sample.get("reallocation_calls"),
                    f"allocator {repeat}/{shape} {side} sample {index}.reallocation_calls")
                calls.append(sample["allocation_calls"])
                realloc.append(sample["reallocation_calls"])
        left_call_summary = _numeric_summary(left_calls,
                                              f"allocator {repeat}/{shape} baseline allocation_calls")
        right_call_summary = _numeric_summary(right_calls,
                                               f"allocator {repeat}/{shape} candidate allocation_calls")
        call_reduction = _reduction_percent(
            left_call_summary["p50"], right_call_summary["p50"],
            f"allocator {repeat}/{shape} allocation_calls p50")
        left_realloc_summary = _numeric_summary(
            left_realloc, f"allocator {repeat}/{shape} baseline reallocation_calls")
        right_realloc_summary = _numeric_summary(
            right_realloc, f"allocator {repeat}/{shape} candidate reallocation_calls")
        checks = {
            "allocation_calls": {
                "baseline": left_call_summary,
                "candidate": right_call_summary,
                "reduction_percent": call_reduction,
                "required_reduction_percent": thresholds[
                    "allocation_calls_reduction_percent"],
                "passed": call_reduction >= thresholds[
                    "allocation_calls_reduction_percent"],
            },
            "reallocation_calls": {
                "baseline": left_realloc_summary,
                "candidate": right_realloc_summary,
                "delta_p50": right_realloc_summary["p50"] - left_realloc_summary["p50"],
                "used_for_gate": False,
                "note": "Reported separately; no percent reinterpretation as allocation_calls.",
            },
        }
        rows.append({
            "repeat": repeat,
            "shape": shape,
            **checks,
            "passed": checks["allocation_calls"]["passed"],
        })
    return {
        "rows": rows,
        "passed": all(row["passed"] for row in rows),
        "scope": (
            "Every planned allocator shape/repeat must reduce the measured "
            "allocation_calls p50; reallocation_calls are diagnostic only."
        ),
    }


def _conditional_lanes(pilot_passed: bool) -> dict[str, dict[str, str]]:
    if pilot_passed:
        reason = "Pilot gates passed; this lane remains unmeasured pending its bound evidence."
    else:
        reason = "Pilot gate failed; this conditional lane was not admitted and is unmeasured."
    return {
        lane: {"status": "unmeasured", "reason": reason}
        for lane in CONDITIONAL_LANES
    }


def _pilot_admission(result: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    native = _native_admission(result, plan)
    allocation = _allocation_admission(result, plan)
    passed = native["passed"] and allocation["passed"]
    return {
        "native": native,
        "allocation": allocation,
        "passed": passed,
        "decision": "eligible-for-conditional-lanes" if passed else "reject",
        "conditional_lanes": _conditional_lanes(passed),
        "scope": (
            "0527 pilot admission requires every planned native and allocator "
            "shape/repeat gate to pass."
        ),
    }


def _stage_lane_status() -> dict[str, dict[str, str]]:
    return {
        lane: {
            "status": "unmeasured",
            "reason": "Requires the completed baseline/candidate comparison and pilot gates.",
        }
        for lane in CONDITIONAL_LANES
    }


def analyze(stage: str | None = None) -> dict[str, Any]:
    """Run retained evidence validation and add the 0527 structured gates."""

    plan = _read_plan()
    result = BASE.analyze(stage)
    result["scope"] = "0527 XLSX source-backed changed-row reconstruction pilot comparison"
    result["structured_gates"] = dict(plan["gates"])
    if result.get("stage") == "compare":
        pilot = _pilot_admission(result, plan)
        result["native_admission"] = pilot["native"]
        result["allocation_admission"] = pilot["allocation"]
        result["pilot_admission"] = pilot
        result["admission_status"] = pilot["decision"]
        result["conditional_lanes"] = pilot["conditional_lanes"]
    else:
        result["conditional_lanes"] = _stage_lane_status()
    result["numerical_verifier"] = {
        "path": str(HELPER.relative_to(HERE.parent)),
        "sha256": hashlib.sha256(HELPER.read_bytes()).hexdigest(),
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("baseline", "candidate", "compare"))
    parser.add_argument("--output", type=Path)
    parser.add_argument("output_positional", nargs="?", type=Path)
    args = parser.parse_args()
    stage = None if args.stage == "compare" else args.stage
    output = args.output or args.output_positional
    if output is None:
        output = HERE / (f"analysis-{stage}.json" if stage else "comparison.json")
    try:
        result = analyze(stage)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    except BASE.EvidenceError as error:
        print(f"evidence check failed: {error}", file=sys.stderr)
        return 1
    print(f"0527 {result['stage']} evidence verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
