#!/usr/bin/env python3
"""Offline mutation and statistics preflight checks for the 0835 reader.

The synthetic statistics cases run before any native qualification exists.  If
the retained qualification reports are available, the second group exercises
the report gate against representative forged evidence as well.
"""

from __future__ import annotations

from copy import deepcopy
import hashlib
import json
from pathlib import Path
import sys

import reader


HERE = Path(__file__).resolve().parent


def _statistics(acquisition: list[int]) -> dict[str, object]:
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


def _rejects(name: str, value: dict[str, object], checks: list[dict[str, object]]) -> None:
    try:
        reader._strict_stats(value, f"synthetic.{name}", len(value["samples"]))
    except (reader.QualificationError, AssertionError, KeyError, TypeError, ValueError):
        checks.append({"name": name, "rejected": True})
    else:
        raise AssertionError(f"accepted statistics mutation: {name}")


def _report_mutations(checks: list[dict[str, object]]) -> str | None:
    qualification = HERE / "qualification.json"
    build = HERE / "build-baseline.json"
    if not qualification.exists() or not build.exists():
        return None
    manifest = reader.read(qualification)
    rows = manifest.get("rows")
    if not isinstance(rows, list) or not rows or not isinstance(rows[0], dict):
        raise AssertionError("qualification manifest is not ready for mutation checks")
    descriptor = rows[0].get("report")
    report_path = reader._descriptor(descriptor, HERE, "mutation.report")
    base = reader.read(report_path)
    binary = reader.read(build).get("binary")

    def validate(value: dict[str, object]) -> None:
        reader.validate_report(value, reader.CASES[0], reader.STATES, 1, 0, binary)

    validate(base)
    mutations = {
        "wrong-git-revision": lambda value: value["environment"].update(git_revision="0" * 40),
        "wrong-corpus-hash": lambda value: value["results"][0]["corpus"].update(archive_sha256="0" * 64),
        "missing-cold-proof": lambda value: _remove_existing_cold_proof(value),
        "duplicate-child-pid": lambda value: value["filesystem_evidence"][0]["samples"][1].update(
            child_process_id=value["filesystem_evidence"][0]["samples"][0]["child_process_id"]
        ),
    }
    for name, mutate in mutations.items():
        candidate = deepcopy(base)
        mutate(candidate)
        try:
            validate(candidate)
        except (reader.QualificationError, AssertionError, KeyError, TypeError, ValueError):
            checks.append({"name": name, "rejected": True})
        else:
            raise AssertionError(f"accepted report mutation: {name}")
    return hashlib.sha256(report_path.read_bytes()).hexdigest()


def _remove_existing_cold_proof(value: dict[str, object]) -> None:
    samples = value["filesystem_evidence"][0]["samples"]
    cold = next(sample for sample in samples if sample.get("cache_state") == "cold-verified")
    if "cold_verified" not in cold:
        raise AssertionError("cold qualification sample has no cold_verified key to mutate")
    cold.pop("cold_verified")


def main() -> None:
    acquisition_input = [
        300, 100, 200, 100, 500, 400, 200, 600, 700, 800,
        900, 1000, 1100, 1200, 1300, 1400, 1500, 1600, 1700, 1800,
        1900, 2000, 2100, 2200, 2300, 2400, 2500, 2600, 2700, 2800,
    ]
    valid = _statistics(acquisition_input)
    reader._strict_stats(valid, "synthetic.valid", 30)
    acquisition = reader.reconstruct_acquisition_samples(valid)
    if acquisition != acquisition_input:
        raise AssertionError(f"sample-order reconstruction differs: {acquisition!r}")
    checks: list[dict[str, object]] = []
    wrong_order = deepcopy(valid)
    wrong_order["sample_order"][0], wrong_order["sample_order"][1] = (
        wrong_order["sample_order"][1], wrong_order["sample_order"][0]
    )
    _rejects("tie-break-order", wrong_order, checks)
    wrong_permutation = deepcopy(valid)
    wrong_permutation["sample_order"][1] = wrong_permutation["sample_order"][0]
    _rejects("duplicate-sample-order", wrong_permutation, checks)
    wrong_values = deepcopy(valid)
    wrong_values["samples"][1] = 99
    _rejects("unsorted-values", wrong_values, checks)
    wrong_mean = deepcopy(valid)
    wrong_mean["mean"] = 0.0
    _rejects("mean-mismatch", wrong_mean, checks)
    wrong_interval = deepcopy(valid)
    wrong_interval["confidence_interval_95"]["upper"] += 1.0
    _rejects("confidence-interval-mismatch", wrong_interval, checks)
    report_sha256 = _report_mutations(checks)
    result = {
        "status": "pass",
        "statistics": {
            "valid_unsorted_acquisition": acquisition_input,
            "sorted_samples": valid["samples"],
            "sample_order": valid["sample_order"],
            "reconstructed_acquisition": acquisition,
        },
        "mutations": checks,
        "report_sha256": report_sha256,
    }
    destination = HERE / "reader-tests.json"
    if "--check" in sys.argv:
        if reader.read(destination) != result:
            raise AssertionError("reader-tests.json differs from fresh mutation result")
    else:
        with destination.open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2, sort_keys=True)
            stream.write("\n")
    print(f"reader preflight PASS: {len(checks)} mutations rejected")


if __name__ == "__main__":
    main()
