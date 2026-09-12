"""Validate and compare the 0529 XML-minifier pilot evidence.

The report and custody checks are inherited from the retained 0521 verifier.
This adapter adds the 0529 pilot admission rules: every primary shape/repeat
must meet total p50, total mean, and publication p50 reductions, and every
allocator shape/repeat must meet the publication allocation-call reduction.
The harness records the allocation region around
``publish_multi_commit_to_stream`` (including the returned ``MultiSnapshot``
drop) separately from the existing commit allocation region.  Commit
allocation counters remain validated diagnostics and never substitute for the
publication metric in a gate.

Profile, hardware, and eager lanes remain unmeasured until the pilot passes
and their own bound evidence is analyzed.  In particular, publication Ir is a
conditional gate; a failed pilot rejects it without treating an absent
profile as a pass.
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
_helper_spec = importlib.util.spec_from_file_location("xlsx_0521_numerical_0529", HELPER)
if _helper_spec is None or _helper_spec.loader is None:
    raise ImportError(f"cannot load numerical helper: {HELPER}")
BASE = importlib.util.module_from_spec(_helper_spec)
_helper_spec.loader.exec_module(BASE)

# The retained verifier resolves all artifacts relative to these module
# globals.  Keep those bindings local to the 0529 result bundle.
BASE.HERE = HERE
BASE.PLAN = HERE / "plan.json"
BASE.BOOTSTRAP_SEED = 5290529


def __getattr__(name: str) -> Any:
    return getattr(BASE, name)


GATE_FIELDS = (
    "total_p50_reduction_percent",
    "total_mean_reduction_percent",
    "allocation_calls_reduction_percent",
    "publication_p50_reduction_percent",
    "publication_ir_reduction_percent",
    "require_every_shape_repeat",
)
CONDITIONAL_LANES = ("profile", "hardware", "eager")

ALLOCATION_VECTOR_NAMES = (
    "commit_allocation_metrics",
    "publication_allocation_metrics",
)

# The retained helper calls this global while it validates a row.  Keep the
# two allocation vectors out of logical identity while still checking their
# lengths and sample contents separately below.  Shallow copying is enough:
# only the XLSX source map is changed and the helper never mutates it.
_BASE_SOURCE_IDENTITY = BASE.source_identity
_BASE_VALIDATE_RESULT = BASE.validate_result


def _source_identity(source: dict[str, Any], count: int) -> dict[str, Any]:
    require = BASE.require
    require(isinstance(source, dict), "result.source is not an object")
    xlsx = source.get("xlsx_cell_values")
    require(isinstance(xlsx, dict), "source.xlsx_cell_values is not an object")
    identity_source = dict(source)
    identity_xlsx = dict(xlsx)
    for name in ALLOCATION_VECTOR_NAMES:
        identity_xlsx.pop(name, None)
    identity_source["xlsx_cell_values"] = identity_xlsx
    return _BASE_SOURCE_IDENTITY(identity_source, count)


BASE.source_identity = _source_identity


def _validate_allocation_vector(
    xlsx: dict[str, Any], job: dict[str, Any], allocator: bool, name: str,
) -> list[dict[str, Any]]:
    """Validate one allocation vector and return normalized samples.

    ``commit_allocation_metrics`` and ``publication_allocation_metrics`` have
    the same Sample schema, but they measure different scopes.  Keeping this
    function parameterized prevents the old commit vector from being silently
    reused for the publication gate.
    """

    label = f"{job['name']}.{name}"
    values = BASE.check_vector(xlsx.get(name), job["samples"], label)
    rows: list[dict[str, Any]] = []
    for index, value in enumerate(values):
        rows.append({
            "index": index,
            **BASE.allocation_sample(value, f"{label}[{index}]", allocator),
        })
    return rows


def _validate_result(
    raw: dict[str, Any], plan: dict[str, Any], job: dict[str, Any],
    binary: dict[str, Any], allocator: bool,
) -> dict[str, Any]:
    """Run retained validation, then validate and retain both vector scopes."""

    row = _BASE_VALIDATE_RESULT(raw, plan, job, binary, allocator)
    result = raw["results"][0]
    source = result["source"]
    xlsx = source["xlsx_cell_values"]
    # The retained helper has already validated and normalized the commit
    # vector as row["allocation"]. Re-run it here to make the scope explicit,
    # then attach the new publication vector under a distinct row key.
    commit_rows = _validate_allocation_vector(
        xlsx, job, allocator, "commit_allocation_metrics")
    BASE.require(commit_rows == row["allocation"],
                 f"{job['name']} commit allocation normalization changed")
    row["publication_allocation"] = _validate_allocation_vector(
        xlsx, job, allocator, "publication_allocation_metrics")
    row["allocation_metric_scopes"] = {
        "commit": {
            "field": "commit_allocation_metrics",
            "used_for_gate": False,
            "scope": "edit.commit region; diagnostic only",
        },
        "publication": {
            "field": "publication_allocation_metrics",
            "used_for_gate": allocator,
            "scope": (
                "publish_multi_commit_to_stream region, including returned "
                "MultiSnapshot drop"
            ),
        },
    }
    return row


BASE.validate_result = _validate_result


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
    """Expose the frozen 0529 gates with explicit metric scopes."""

    gates = _validated_gates(plan)
    return {
        "native_primary_total_p50_improvement_percent": float(
            gates["total_p50_reduction_percent"]),
        "native_primary_total_mean_improvement_percent": float(
            gates["total_mean_reduction_percent"]),
        "native_primary_publication_p50_improvement_percent": float(
            gates["publication_p50_reduction_percent"]),
        "publication_allocation_calls_reduction_percent": float(
            gates["allocation_calls_reduction_percent"]),
        "publication_ir_reduction_percent": float(
            gates["publication_ir_reduction_percent"]),
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
    """Evaluate total and publication timing gates for every primary pair.

    Commit p50 is retained as a diagnostic because the harness still records
    it, but it has no 0529 threshold.  The candidate's intended work is in
    the XML publication path, so silently carrying forward the 0527 commit
    gate would test the wrong operation.
    """

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
        publication_p50 = _reduction_percent(
            _timing(left, "publication_ns", f"primary {repeat}/{shape} baseline")["p50"],
            _timing(right, "publication_ns", f"primary {repeat}/{shape} candidate")["p50"],
            f"primary {repeat}/{shape} publication p50")
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
            "native_primary_publication_p50": {
                "baseline": _timing(left, "publication_ns",
                                     f"primary {repeat}/{shape} baseline")["p50"],
                "candidate": _timing(right, "publication_ns",
                                      f"primary {repeat}/{shape} candidate")["p50"],
                "reduction_percent": publication_p50,
                "required_reduction_percent": thresholds[
                    "native_primary_publication_p50_improvement_percent"],
                "passed": publication_p50 >= thresholds[
                    "native_primary_publication_p50_improvement_percent"],
            },
            "commit_p50_diagnostic": {
                "baseline": left_commit["p50"],
                "candidate": right_commit["p50"],
                "reduction_percent": commit_p50,
                "used_for_gate": False,
                "note": "Retained commit timing diagnostic; no 0529 commit threshold.",
            },
        }
        rows.append({
            "repeat": repeat,
            "shape": shape,
            **checks,
            "passed": all(check["passed"] for check in checks.values()
                           if check.get("used_for_gate", True)),
        })
    return {
        "rows": rows,
        "passed": all(row["passed"] for row in rows),
        "scope": (
            "Every planned primary shape/repeat must meet total elapsed p50, "
            "total mean, and publication p50 reductions; commit p50 is diagnostic."
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
    """Evaluate publication allocation-call reduction per allocator pair.

    The retained 0521 ``allocation`` row member is the commit region.  The
    0529 row member ``publication_allocation`` is the region around
    ``publish_multi_commit_to_stream`` and is the only vector used here.
    """

    thresholds = _admission_thresholds(plan)
    baseline = _allocation_rows(result["baseline"], plan, "baseline")
    candidate = _allocation_rows(result["candidate"], plan, "candidate")
    rows: list[dict[str, Any]] = []
    for repeat, shape in sorted(baseline):
        left_samples = baseline[(repeat, shape)].get("publication_allocation")
        right_samples = candidate[(repeat, shape)].get("publication_allocation")
        BASE.require(isinstance(left_samples, list)
                     and isinstance(right_samples, list),
                     f"publication allocator {repeat}/{shape} samples are not lists")
        BASE.require(len(left_samples) == len(right_samples) and left_samples,
                     f"publication allocator {repeat}/{shape} sample counts differ or are empty")
        left_calls: list[Any] = []
        right_calls: list[Any] = []
        left_realloc: list[Any] = []
        right_realloc: list[Any] = []
        for index, (left, right) in enumerate(zip(left_samples, right_samples)):
            BASE.require(isinstance(left, dict) and isinstance(right, dict),
                         f"publication allocator {repeat}/{shape} sample {index} is not an object")
            BASE.require(left.get("status") == right.get("status") == "measured",
                         f"publication allocator {repeat}/{shape} sample {index} is not measured")
            for sample, calls, realloc, side in (
                (left, left_calls, left_realloc, "baseline"),
                (right, right_calls, right_realloc, "candidate"),
            ):
                BASE.nonnegative_integer(
                    sample.get("allocation_calls"),
                    f"publication allocator {repeat}/{shape} {side} sample {index}.allocation_calls")
                BASE.nonnegative_integer(
                    sample.get("reallocation_calls"),
                    f"publication allocator {repeat}/{shape} {side} sample {index}.reallocation_calls")
                calls.append(sample["allocation_calls"])
                realloc.append(sample["reallocation_calls"])
        left_call_summary = _numeric_summary(left_calls,
                                              f"publication allocator {repeat}/{shape} baseline allocation_calls")
        right_call_summary = _numeric_summary(right_calls,
                                               f"publication allocator {repeat}/{shape} candidate allocation_calls")
        call_reduction = _reduction_percent(
            left_call_summary["p50"], right_call_summary["p50"],
            f"publication allocator {repeat}/{shape} allocation_calls p50")
        left_realloc_summary = _numeric_summary(
            left_realloc, f"publication allocator {repeat}/{shape} baseline reallocation_calls")
        right_realloc_summary = _numeric_summary(
            right_realloc, f"publication allocator {repeat}/{shape} candidate reallocation_calls")
        checks = {
            "allocation_calls": {
                "metric": "publication_allocation_metrics",
                "baseline": left_call_summary,
                "candidate": right_call_summary,
                "reduction_percent": call_reduction,
                "required_reduction_percent": thresholds[
                    "publication_allocation_calls_reduction_percent"],
                "passed": call_reduction >= thresholds[
                    "publication_allocation_calls_reduction_percent"],
            },
            "reallocation_calls": {
                "metric": "publication_allocation_metrics",
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
            "publication_allocation_metrics allocation_calls p50; commit "
            "allocation counters and publication reallocation_calls are diagnostic only."
        ),
    }


def _publication_allocation_comparisons(
    baseline: dict[str, Any], candidate: dict[str, Any],
) -> list[dict[str, Any]]:
    """Compare the new publication vector while retaining all report fields.

    ``BASE.compare_stages`` already compares the old commit vector.  This
    companion comparison deliberately uses a different row key, so a future
    change cannot accidentally make commit allocations look like publication
    allocations merely because both use the same Sample schema.
    """

    left_rows = {BASE.job_key(row): row for row in baseline["allocation"]["rows"]}
    right_rows = {BASE.job_key(row): row for row in candidate["allocation"]["rows"]}
    BASE.require(set(left_rows) == set(right_rows),
                 "publication allocator baseline/candidate row keys differ")
    comparisons: list[dict[str, Any]] = []
    for key in sorted(left_rows):
        left, right = left_rows[key], right_rows[key]
        BASE.require(left["identity"] == right["identity"],
                     f"publication allocator logical identity differs for {key}")
        left_samples = left.get("publication_allocation")
        right_samples = right.get("publication_allocation")
        BASE.require(isinstance(left_samples, list) and isinstance(right_samples, list),
                     f"publication allocator samples are not lists for {key}")
        BASE.require(len(left_samples) == len(right_samples) and left_samples,
                     f"publication allocator sample counts differ or are empty for {key}")
        metrics: dict[str, Any] = {}
        for field in BASE.ALLOCATION_REPORT_FIELDS:
            left_values: list[int] = []
            right_values: list[int] = []
            for index, (left_sample, right_sample) in enumerate(
                zip(left_samples, right_samples)
            ):
                BASE.require(
                    left_sample.get("status") == right_sample.get("status") == "measured",
                    f"publication allocator sample status is not measured for {key} at {index}",
                )
                BASE.nonnegative_integer(
                    left_sample.get(field),
                    f"publication allocator baseline {key} sample {index}.{field}",
                )
                BASE.nonnegative_integer(
                    right_sample.get(field),
                    f"publication allocator candidate {key} sample {index}.{field}",
                )
                left_values.append(left_sample[field])
                right_values.append(right_sample[field])
            metrics[field] = {
                "baseline": left_values,
                "candidate": right_values,
                "delta": [right_value - left_value
                          for left_value, right_value in zip(left_values, right_values)],
                "baseline_summary": _numeric_summary(
                    left_values, f"publication allocator baseline {key} {field}"),
                "candidate_summary": _numeric_summary(
                    right_values, f"publication allocator candidate {key} {field}"),
            }
        comparisons.append({
            "case": key[3],
            "shape": key[4],
            "repeat": key[2],
            "identity_equal": True,
            "metric": "publication_allocation_metrics",
            "scope": (
                "publish_multi_commit_to_stream region including returned "
                "MultiSnapshot drop"
            ),
            "metrics": metrics,
        })
    return comparisons


def _adapt_comparison(result: dict[str, Any]) -> None:
    """Make the retained comparison's allocation lane publication-scoped."""

    comparison = result.get("comparison")
    BASE.require(isinstance(comparison, dict), "comparison object is missing")
    old_commit = comparison.pop("allocation_comparisons", None)
    BASE.require(isinstance(old_commit, list),
                 "retained commit allocation comparison is missing")
    publication = _publication_allocation_comparisons(
        result["baseline"], result["candidate"])
    # Keep the established key for downstream readers, but bind it explicitly
    # to publication metrics.  The previous key contents remain available as
    # separately named commit diagnostics.
    comparison["allocation_comparisons"] = publication
    comparison["publication_allocation_comparisons"] = publication
    comparison["commit_allocation_diagnostics"] = old_commit
    comparison["allocation_metric"] = "publication_allocation_metrics"
    comparison["commit_allocation_metric"] = "commit_allocation_metrics"
    comparison["allocation_scope"] = (
        "Publication region includes publish_multi_commit_to_stream and the "
        "returned MultiSnapshot drop."
    )


def _conditional_lanes(pilot_passed: bool, plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    if pilot_passed:
        reason = "Pilot gates passed; this lane remains unmeasured pending its bound evidence."
    else:
        reason = (
            "Pilot gate failed; conditional capture was not admitted and this "
            "lane is unmeasured."
        )
    configured = plan.get("conditional_lanes")
    lanes = tuple(configured) if isinstance(configured, dict) else CONDITIONAL_LANES
    result: dict[str, dict[str, Any]] = {}
    for lane in lanes:
        record: dict[str, Any] = {"status": "unmeasured", "reason": reason}
        if lane == "profile":
            record.update({
                "metric": "publication_ir",
                "required_reduction_percent": float(
                    _admission_thresholds(plan)["publication_ir_reduction_percent"]),
                "used_for_gate": False,
            })
            if pilot_passed:
                record["reason"] = (
                    "Pilot gates passed; publication Ir remains unmeasured until "
                    "bound profile evidence is analyzed."
                )
        result[lane] = record
    return result


def _pilot_admission(result: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    native = _native_admission(result, plan)
    allocation = _allocation_admission(result, plan)
    passed = native["passed"] and allocation["passed"]
    return {
        "native": native,
        "allocation": allocation,
        "passed": passed,
        "decision": "eligible-for-conditional-lanes" if passed else "reject",
        "conditional_lanes": _conditional_lanes(passed, plan),
        "scope": (
            "0529 pilot admission requires every planned primary and publication "
            "allocator shape/repeat gate to pass. Publication Ir is conditional."
        ),
    }


def _stage_lane_status(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    configured = plan.get("conditional_lanes")
    lanes = tuple(configured) if isinstance(configured, dict) else CONDITIONAL_LANES
    return {
        lane: {
            "status": "unmeasured",
            "reason": "Requires the completed baseline/candidate comparison and pilot gates.",
            **({
                "metric": "publication_ir",
                "required_reduction_percent": float(
                    _admission_thresholds(plan)["publication_ir_reduction_percent"]),
                "used_for_gate": False,
            } if lane == "profile" else {}),
        }
        for lane in lanes
    }


def analyze(stage: str | None = None) -> dict[str, Any]:
    """Run retained evidence validation and add the 0529 structured gates."""

    plan = _read_plan()
    result = BASE.analyze(stage)
    result["scope"] = "0529 XML-minifier XLSX source-backed publication pilot comparison"
    result["structured_gates"] = dict(plan["gates"])
    result["allocation_metric_scopes"] = {
        "publication": {
            "field": "publication_allocation_metrics",
            "comparison_key": "allocation_comparisons",
            "used_for_gate": True,
            "scope": (
                "publish_multi_commit_to_stream, including returned "
                "MultiSnapshot drop"
            ),
        },
        "commit": {
            "field": "commit_allocation_metrics",
            "comparison_key": "commit_allocation_diagnostics",
            "used_for_gate": False,
            "scope": "edit.commit region; diagnostic only",
        },
    }
    if result.get("stage") == "compare":
        _adapt_comparison(result)
        pilot = _pilot_admission(result, plan)
        result["native_admission"] = pilot["native"]
        result["allocation_admission"] = pilot["allocation"]
        result["pilot_admission"] = pilot
        result["admission_status"] = pilot["decision"]
        result["conditional_lanes"] = pilot["conditional_lanes"]
    else:
        result["conditional_lanes"] = _stage_lane_status(plan)
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
    print(f"0529 {result['stage']} evidence verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
