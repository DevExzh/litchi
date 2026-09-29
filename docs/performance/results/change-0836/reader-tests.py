#!/usr/bin/env python3
"""Offline mutation tests for the 0836 evidence reader.

The tests exercise the strict statistics mapping on synthetic data and, once
the root agent has retained native reports, mutate real multi-sample reports.
Every mutation must be rejected by the same reader used for admission.  This
file never starts a native process, builds a binary, or edits a retained
report; it only writes the small ``reader-tests.json`` result when requested.
"""

from __future__ import annotations

from copy import deepcopy
import hashlib
import json
from pathlib import Path
import sys
from typing import Any

import reader


HERE = Path(__file__).resolve().parent


def _statistics(acquisition: list[int]) -> dict[str, Any]:
    ordered = sorted(enumerate(acquisition), key=lambda item: (item[1], item[0]))
    values = [value for _index, value in ordered]
    order = [index for index, _value in ordered]
    count = len(values)
    mean = sum(values) / count
    running = 0.0
    squared = 0.0
    for index, value in enumerate(values):
        n = index + 1
        delta = value - running
        next_mean = running + delta / n
        squared += delta * (value - next_mean)
        running = next_mean
    deviation = (squared / (count - 1)) ** 0.5 if count > 1 else 0.0
    critical = reader._LEGACY._student_t_critical_95(count - 1)
    margin = critical * deviation / count**0.5 if count > 1 else 0.0
    nearest = lambda rank: values[min(count - 1, (rank * count + 99) // 100 - 1)]
    midpoint = lambda left, right: left // 2 + right // 2 + (left % 2 + right % 2) // 2
    return {
        "unit": "ns",
        "samples": values,
        "sample_order": order,
        "min": values[0],
        "p50": midpoint(values[(count - 1) // 2], values[count // 2]),
        "p95": nearest(95),
        "p99": nearest(99),
        "max": values[-1],
        "mean": mean,
        "standard_deviation": deviation,
        "confidence_interval_95": {
            "method": "two-sided Student's t interval for the mean",
            "lower": max(0.0, mean - margin),
            "upper": mean + margin,
        },
    }


def _rejects(name: str, value: dict[str, Any], checks: list[dict[str, Any]]) -> None:
    try:
        reader._strict_stats(value, f"synthetic.{name}", len(value["samples"]))
    except (reader.QualificationError, AssertionError, KeyError, TypeError, ValueError):
        checks.append({"name": name, "rejected": True})
    else:
        raise AssertionError(f"accepted statistics mutation: {name}")


def _manifest_candidates() -> list[tuple[Path, dict[str, Any]]]:
    candidates: list[tuple[Path, dict[str, Any]]] = []
    for path in sorted(HERE.glob("*.json")):
        if path.name in {"reader-tests.json", "validation.json"}:
            continue
        try:
            value = reader.read(path)
        except reader.QualificationError:
            continue
        if not isinstance(value, dict) or value.get("status") != "commands_pass":
            continue
        rows = value.get("rows")
        if not isinstance(rows, list):
            continue
        if any(isinstance(row, dict) and
               reader._field(row, "stage") in reader.STAGES and
               isinstance(reader._field(row, "report", "report_descriptor", "reportdesc", "report_desc"), dict)
               for row in rows):
            candidates.append((path, value))
    return candidates


def _row_data(manifest_path: Path, row: dict[str, Any]) -> tuple[Path, str, Any, int, int]:
    stage = reader._field(row, "stage")
    state = reader._field(row, "state", "cache_state", "cache_states")
    samples = reader._field(row, "samples")
    warmup = reader._field(row, "warmup")
    require = reader.require
    require(stage in reader.STAGES, "mutation row has no valid stage")
    require(type(samples) is int and samples > 0, "mutation row has invalid samples")
    require(type(warmup) is int and warmup >= 0, "mutation row has invalid warmup")
    descriptor = reader._row_descriptor(
        row, ("report", "report_descriptor", "reportdesc", "report_desc"),
        "mutation.report",
    )
    report_path = reader._descriptor(descriptor, manifest_path.parent, "mutation.report")
    return report_path, stage, state, samples, warmup


def _rejects_report(name: str, base: dict[str, Any], stage: str, state: Any,
                    samples: int, warmup: int, checks: list[dict[str, Any]],
                    mutate: Any) -> None:
    candidate = deepcopy(base)
    mutate(candidate)
    try:
        reader.validate_report(candidate, stage, state, samples, warmup)
    except (reader.QualificationError, AssertionError, KeyError, TypeError, ValueError):
        checks.append({"name": name, "rejected": True})
    else:
        raise AssertionError(f"accepted report mutation: {name}")


def _report_mutations(checks: list[dict[str, Any]]) -> str:
    candidates = _manifest_candidates()
    if not candidates:
        raise AssertionError("no retained stage-aware manifest is available")

    selected: tuple[Path, dict[str, Any], dict[str, Any], str, Any, int, int] | None = None
    cold_selected: tuple[Path, dict[str, Any], dict[str, Any], str, Any, int, int] | None = None
    for manifest_path, manifest in candidates:
        for row in manifest["rows"]:
            if not isinstance(row, dict):
                continue
            try:
                report_path, stage, state, samples, warmup = _row_data(manifest_path, row)
                report = reader.read(report_path)
            except (reader.QualificationError, KeyError, TypeError, ValueError):
                continue
            if not isinstance(report, dict):
                continue
            item = (manifest_path, report_path, report, stage, state, samples, warmup)
            if samples > 1 and selected is None:
                selected = item
            if "cold-verified" in reader._parse_states(state, "mutation state") and cold_selected is None:
                cold_selected = item
            if selected is not None and cold_selected is not None:
                break
        if selected is not None and cold_selected is not None:
            break
    if selected is None:
        raise AssertionError("no retained multi-sample report is available")

    _manifest_path, report_path, base, stage, state, samples, warmup = selected
    reader.validate_report(base, stage, state, samples, warmup)
    result = next(item for item in base["results"]
                   if item.get("cache_state") == reader._parse_states(state, "mutation state")[0])
    stats_mutation = lambda value: value["results"][0]["elapsed_ns"]["sample_order"].__setitem__(
        1, value["results"][0]["elapsed_ns"]["sample_order"][0]
    )
    if samples <= 1:
        raise AssertionError("selected report unexpectedly has one sample")
    _rejects_report("multisample-stat-order-permutation", base, stage, state,
                    samples, warmup, checks, stats_mutation)

    def mutate_elapsed(value: dict[str, Any]) -> None:
        state_name = reader._parse_states(state, "mutation state")[0]
        target = next(item for item in value["results"]
                      if item.get("cache_state") == state_name)
        evidence = next(item for record in value["filesystem_evidence"]
                        for item in record["samples"]
                        if item.get("cache_state") == state_name and item.get("sample_index") == 0)
        target["elapsed_ns"]["samples"][0] += 1
        # Keep the vector internally plausible; the evidence binding must still
        # reject the one-nanosecond disagreement.
        evidence["elapsed_ns"] += 0

    _rejects_report("multisample-evidence-stat-mapping", base, stage, state,
                    samples, warmup, checks, mutate_elapsed)

    if cold_selected is not None:
        _cm, _cr, cold_base, cold_stage, cold_state, cold_samples, cold_warmup = cold_selected
        reader.validate_report(cold_base, cold_stage, cold_state, cold_samples, cold_warmup)

        def mutate_cold_proof(value: dict[str, Any]) -> None:
            cold = next(sample for record in value["filesystem_evidence"]
                        for sample in record["samples"]
                        if sample.get("cache_state") == "cold-verified")
            cold.pop("cold_verified", None)

        _rejects_report("cold-proof-removed", cold_base, cold_stage, cold_state,
                        cold_samples, cold_warmup, checks, mutate_cold_proof)

        def mutate_cold_alignment(value: dict[str, Any]) -> None:
            evidence = value["filesystem_evidence"][0]
            proof = evidence["cold_verified_samples"][0]
            proof["read_bytes_delta"] += 1

        _rejects_report("cold-proof-counter-mismatch", cold_base, cold_stage, cold_state,
                        cold_samples, cold_warmup, checks, mutate_cold_alignment)
    else:
        raise AssertionError("no retained cold report is available")

    return hashlib.sha256(report_path.read_bytes()).hexdigest()


def main() -> None:
    acquisition_input = [
        300, 100, 200, 100, 500, 400, 200, 600,
        700, 900, 800, 1000, 1100, 1300, 1200, 1400,
    ]
    valid = _statistics(acquisition_input)
    reader._strict_stats(valid, "synthetic.valid", len(acquisition_input))
    reconstructed = reader.reconstruct_acquisition_samples(valid)
    if reconstructed != acquisition_input:
        raise AssertionError(f"sample-order reconstruction differs: {reconstructed!r}")
    checks: list[dict[str, Any]] = []

    wrong_order = deepcopy(valid)
    wrong_order["sample_order"][0], wrong_order["sample_order"][1] = (
        wrong_order["sample_order"][1], wrong_order["sample_order"][0]
    )
    _rejects("tie-break-order", wrong_order, checks)
    wrong_permutation = deepcopy(valid)
    wrong_permutation["sample_order"][1] = wrong_permutation["sample_order"][0]
    _rejects("duplicate-sample-order", wrong_permutation, checks)
    wrong_value = deepcopy(valid)
    wrong_value["samples"][1] = 99
    _rejects("unsorted-statistics", wrong_value, checks)
    wrong_mean = deepcopy(valid)
    wrong_mean["mean"] = 0.0
    _rejects("mean-mismatch", wrong_mean, checks)
    wrong_interval = deepcopy(valid)
    wrong_interval["confidence_interval_95"]["upper"] += 1.0
    _rejects("confidence-interval-mismatch", wrong_interval, checks)

    if "--preflight" in sys.argv:
        report = reader.read(HERE / "qualification-00.json")
        reader.validate_report(report, "baseline", "warm,cold-verified", 1, 0)
        def remove_proof(value):
            cold = next(s for s in value["filesystem_evidence"][0]["samples"] if s["cache_state"] == "cold-verified")
            del cold["cold_verified"]
        _rejects_report("qualification-cold-proof-removed", report, "baseline", "warm,cold-verified", 1, 0, checks, remove_proof)
        report_sha256 = hashlib.sha256((HERE / "qualification-00.json").read_bytes()).hexdigest()
    else:
        report_sha256 = _report_mutations(checks)
    result = {
        "status": "pass",
        "statistics": {
            "valid_unsorted_acquisition": acquisition_input,
            "sorted_samples": valid["samples"],
            "sample_order": valid["sample_order"],
            "reconstructed_acquisition": reconstructed,
        },
        "mutations": checks,
        "report_sha256": report_sha256,
    }
    destination = HERE / ("reader-preflight-tests.json" if "--preflight" in sys.argv else "reader-tests.json")
    if "--check" in sys.argv:
        if reader.read(destination) != result:
            raise AssertionError("reader-tests.json differs from fresh mutation result")
    else:
        with destination.open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2, sort_keys=True)
            stream.write("\n")
    print(f"reader mutation tests PASS: {len(checks)} mutations rejected")


if __name__ == "__main__":
    main()
