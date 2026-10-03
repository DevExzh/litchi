#!/usr/bin/env python3
"""Compare two 0894 observable-probe JSONL receipts.

The comparison is deliberately stricter than a digest-only check.  It binds
the complete input census, worksheet census, four leg identities, operation
errors and source chains, partial callback state, transcript framing/digests,
per-worksheet tolerance reports, aggregate tolerance reports, and source
revision stability.  It ignores only the process-local ``SourceVersion.id``
fields; the public id is an opaque identity allocated independently in each
probe process and cannot establish cross-process semantic equivalence.  The
revision and stable fields remain compared.

The report also lists disagreements among the four legs within each binary.
Those are diagnostic pre-existing behavior and do not become a failure by
themselves.  A baseline/candidate disagreement is a failure, including when
the baseline already had a disagreement between its legs.
"""

from __future__ import annotations

import argparse
import copy
import json
from pathlib import Path
from typing import Any


SCHEMA = "litchi.xlsx-observable-differential-comparison-0894.v1"
PROBE_SCHEMA = "litchi.xlsx-observable-differential-0894.v1"
IGNORED_FIELDS = (
    {
        "field": "legs[*].source_version_before.id",
        "reason": "SourceVersion.id is a process-local opaque identity allocated independently by each probe process; revision remains compared.",
    },
    {
        "field": "legs[*].source_version_after.id",
        "reason": "SourceVersion.id is a process-local opaque identity allocated independently by each probe process; revision remains compared.",
    },
)


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def read_jsonl(path: Path) -> tuple[dict[str, dict[str, Any]], dict[str, Any], list[dict[str, Any]]]:
    inputs: dict[str, dict[str, Any]] = {}
    summary: dict[str, Any] | None = None
    other: list[dict[str, Any]] = []
    with path.open(encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, 1):
            if not line.strip():
                continue
            value = json.loads(line)
            require(isinstance(value, dict), f"{path}:{line_number} is not an object")
            require(value.get("schema") == PROBE_SCHEMA, f"{path}:{line_number} schema mismatch")
            kind = value.get("kind")
            if kind == "input":
                input_value = value.get("input")
                require(isinstance(input_value, dict), f"{path}:{line_number} input metadata missing")
                name = input_value.get("path")
                require(isinstance(name, str) and name, f"{path}:{line_number} input path missing")
                require(name not in inputs, f"{path}:{line_number} duplicate input {name}")
                inputs[name] = value
            elif kind == "summary":
                require(summary is None, f"{path}:{line_number} duplicate summary")
                summary = value
            else:
                other.append(value)
    require(summary is not None, f"{path} has no summary record")
    return inputs, summary, other


def normalize_source_version(value: Any) -> Any:
    if not isinstance(value, dict):
        return value
    return {key: value[key] for key in value if key != "id"}


def normalize_leg(leg: dict[str, Any]) -> dict[str, Any]:
    value = copy.deepcopy(leg)
    for key in ("source_version_before", "source_version_after"):
        value[key] = normalize_source_version(value.get(key))
    return value


def normalize_input(record: dict[str, Any]) -> dict[str, Any]:
    value = copy.deepcopy(record)
    value["legs"] = [normalize_leg(leg) for leg in value.get("legs", [])]
    return value


def _diff(left: Any, right: Any, prefix: str = "$") -> list[dict[str, Any]]:
    if type(left) is not type(right):
        return [{"path": prefix, "baseline": left, "candidate": right}]
    if isinstance(left, dict):
        changes: list[dict[str, Any]] = []
        for key in sorted(set(left) | set(right)):
            child = f"{prefix}.{key}"
            if key not in left:
                changes.append({"path": child, "baseline": None, "candidate": right[key]})
            elif key not in right:
                changes.append({"path": child, "baseline": left[key], "candidate": None})
            else:
                changes.extend(_diff(left[key], right[key], child))
        return changes
    if isinstance(left, list):
        changes = []
        for index in range(max(len(left), len(right))):
            child = f"{prefix}[{index}]"
            if index >= len(left):
                changes.append({"path": child, "baseline": None, "candidate": right[index]})
            elif index >= len(right):
                changes.append({"path": child, "baseline": left[index], "candidate": None})
            else:
                changes.extend(_diff(left[index], right[index], child))
        return changes
    if left != right:
        return [{"path": prefix, "baseline": left, "candidate": right}]
    return []


def core_leg(leg: dict[str, Any]) -> dict[str, Any]:
    """Projection used only to describe pre-existing four-leg differences."""

    value = normalize_leg(leg)
    value.pop("leg", None)
    value.pop("source_version_before", None)
    value.pop("source_version_after", None)
    guard = value.get("guard")
    if isinstance(guard, dict):
        guard.pop("scope", None)
    return value


def leg_key(leg: dict[str, Any]) -> tuple[Any, Any, Any]:
    return (leg.get("sheet_position"), leg.get("sheet_name"), leg.get("leg"))


def grouped_legs(record: dict[str, Any]) -> dict[tuple[Any, Any], dict[str, dict[str, Any]]]:
    grouped: dict[tuple[Any, Any], dict[str, dict[str, Any]]] = {}
    for leg in record.get("legs", []):
        require(isinstance(leg, dict), "leg is not an object")
        key = (leg.get("sheet_position"), leg.get("sheet_name"))
        name = leg.get("leg")
        require(isinstance(name, str) and name, "leg name is missing")
        require(name not in grouped.setdefault(key, {}), f"duplicate leg {key}/{name}")
        grouped[key][name] = leg
    return grouped


def shape_mismatches(inputs: dict[str, dict[str, Any]], role: str) -> list[dict[str, Any]]:
    mismatches: list[dict[str, Any]] = []
    for path in sorted(inputs):
        record = inputs[path]
        status = record.get("status")
        sheets = record.get("sheets")
        legs = record.get("legs")
        if not isinstance(sheets, list) or not isinstance(legs, list):
            mismatches.append({"role": role, "path": path, "reason": "sheets/legs are not arrays"})
            continue
        if status == "ok":
            expected = len(sheets) * 4
            if len(legs) != expected:
                mismatches.append({"role": role, "path": path, "reason": "leg_count", "expected": expected, "actual": len(legs)})
            expected_keys = {
                (sheet.get("position"), sheet.get("name"), leg)
                for sheet in sheets
                if isinstance(sheet, dict)
                for leg in ("cold_source_visit", "cold_source_cells", "warm_source_visit", "eager_model_cells")
            }
            actual_keys = {leg_key(leg) for leg in legs if isinstance(leg, dict)}
            if actual_keys != expected_keys:
                mismatches.append(
                    {
                        "role": role,
                        "path": path,
                        "reason": "leg_keys",
                        "missing": [list(key) for key in sorted(expected_keys - actual_keys, key=str)],
                        "unexpected": [list(key) for key in sorted(actual_keys - expected_keys, key=str)],
                    }
                )
        elif legs:
            mismatches.append(
                {
                    "role": role,
                    "path": path,
                    "reason": "nonempty_legs_for_source_open_failure",
                    "actual": len(legs),
                }
            )
    return mismatches


def summary_shape(summary: dict[str, Any], inputs: dict[str, dict[str, Any]], role: str) -> list[dict[str, Any]]:
    expected_inputs = len(inputs)
    expected_worksheets = sum(len(record.get("sheets", [])) for record in inputs.values())
    expected_legs = sum(len(record.get("legs", [])) for record in inputs.values())
    mismatches = []
    for key, expected in (("inputs", expected_inputs), ("worksheets", expected_worksheets), ("legs", expected_legs)):
        if summary.get(key) != expected:
            mismatches.append({"role": role, "field": key, "expected": expected, "actual": summary.get(key)})
    return mismatches


def preexisting_mismatches(inputs: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    mismatches: list[dict[str, Any]] = []
    for path in sorted(inputs):
        record = inputs[path]
        for sheet, legs in sorted(grouped_legs(record).items(), key=lambda item: str(item[0])):
            if len(legs) < 2:
                continue
            names = sorted(legs)
            reference = core_leg(legs[names[0]])
            differences = []
            for name in names[1:]:
                changes = _diff(reference, core_leg(legs[name]))
                if changes:
                    differences.append({"against": names[0], "leg": name, "changes": changes[:32], "change_count": len(changes)})
            if differences:
                mismatches.append({"path": path, "sheet": list(sheet), "legs": names, "differences": differences})
    return mismatches


def compare(baseline_path: Path, candidate_path: Path) -> dict[str, Any]:
    baseline, baseline_summary, baseline_other = read_jsonl(baseline_path)
    candidate, candidate_summary, candidate_other = read_jsonl(candidate_path)
    baseline_shape = shape_mismatches(baseline, "baseline")
    candidate_shape = shape_mismatches(candidate, "candidate")
    baseline_summary_shape = summary_shape(baseline_summary, baseline, "baseline")
    candidate_summary_shape = summary_shape(candidate_summary, candidate, "candidate")
    input_paths = sorted(set(baseline) | set(candidate))
    input_mismatches: list[dict[str, Any]] = []
    leg_mismatches: list[dict[str, Any]] = []
    for path in input_paths:
        if path not in baseline or path not in candidate:
            input_mismatches.append({"path": path, "baseline_present": path in baseline, "candidate_present": path in candidate})
            continue
        left = normalize_input(baseline[path])
        right = normalize_input(candidate[path])
        changes = _diff(left, right)
        if changes:
            input_mismatches.append({"path": path, "changes": changes[:64], "change_count": len(changes)})
        left_legs = {leg_key(leg): normalize_leg(leg) for leg in baseline[path].get("legs", [])}
        right_legs = {leg_key(leg): normalize_leg(leg) for leg in candidate[path].get("legs", [])}
        for key in sorted(set(left_legs) | set(right_legs), key=str):
            if key not in left_legs or key not in right_legs:
                leg_mismatches.append({"path": path, "leg_key": list(key), "baseline_present": key in left_legs, "candidate_present": key in right_legs})
                continue
            changes = _diff(left_legs[key], right_legs[key])
            if changes:
                leg_mismatches.append({"path": path, "leg_key": list(key), "changes": changes[:64], "change_count": len(changes)})
    # Summary records use scalar ``legs``/``inputs`` counters, so applying
    # input normalization here treats the integer leg count as an iterable.
    # Their schemas are already checked by read_jsonl; compare them directly.
    summary_changes = _diff(baseline_summary, candidate_summary)
    other_changes = _diff(baseline_other, candidate_other)
    result = {
        "schema": SCHEMA,
        "status": "pass"
        if not input_mismatches
        and not leg_mismatches
        and not baseline_shape
        and not candidate_shape
        and not summary_changes
        and not baseline_summary_shape
        and not candidate_summary_shape
        and not other_changes
        else "mismatch",
        "baseline": str(baseline_path),
        "candidate": str(candidate_path),
        "inputs": len(input_paths),
        "baseline_input_count": len(baseline),
        "candidate_input_count": len(candidate),
        "input_mismatches": input_mismatches,
        "leg_mismatches": leg_mismatches,
        "structural_mismatches": baseline_shape + candidate_shape,
        "summary_mismatches": summary_changes,
        "summary_shape_mismatches": baseline_summary_shape + candidate_summary_shape,
        "discovery_record_mismatches": other_changes,
        "preexisting_baseline_leg_mismatches": preexisting_mismatches(baseline),
        "preexisting_candidate_leg_mismatches": preexisting_mismatches(candidate),
        "ignored_fields": list(IGNORED_FIELDS),
        "performance_claim": None,
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", type=Path, required=True)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = compare(args.baseline, args.candidate)
    require(not args.output.exists(), f"refusing to overwrite comparison receipt: {args.output}")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("x", encoding="utf-8") as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write("\n")
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0 if result["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
