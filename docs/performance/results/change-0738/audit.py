#!/usr/bin/env python3
"""Independently replay the 0738 unchanged-owner capture matrix.

This validator duplicates the statistical calculations in ``analyze.py``
without importing its implementation.  It verifies the frozen source and
binary custody, the complete schedule and capture manifest, every oracle
field, and the derived timing, drift-window, bootstrap, and allocation
reports before writing the audit receipt.
"""

from __future__ import annotations

import json
import math
import random
import statistics as st
from pathlib import Path
from typing import Any

from contract import P, ROOT, command, read, sha, validate_report
from run import guard


CASES = ("primary", "secondary")
NATIVE_ARMS = ("archive", "prior", "restored-a", "restored-b")
TIMING_FIELDS = ("p50", "mean", "p95", "p99", "maximum")
WINDOWS = (("first10", 0, 10), ("middle30", 10, 40), ("last10", 40, 50))
ALLOC_FIELDS = (
    "allocated_bytes",
    "deallocated_bytes",
    "allocation_calls",
    "peak_live_bytes",
    "retained_bytes",
)


class AuditError(Exception):
    """An independent audit assertion failed."""


def need(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def integer(value: Any, label: str, minimum: int | None = None) -> int:
    need(isinstance(value, int) and not isinstance(value, bool), f"{label}: not an integer")
    if minimum is not None:
        need(value >= minimum, f"{label}: below {minimum}")
    return value


def basename(value: Any, label: str) -> str:
    need(isinstance(value, str) and value, f"{label}: empty path")
    path = Path(value)
    need(
        not path.is_absolute() and path.name == value and value not in (".", ".."),
        f"{label}: unsafe capture path",
    )
    return value


def statistics(values: list[int | float]) -> dict[str, int | float]:
    """Compute the report's public statistics independently."""
    need(values, "empty sample sequence")
    ordered = sorted(values)
    result: dict[str, int | float] = {
        "p50": st.median(values),
        "mean": st.mean(values),
    }
    if len(ordered) > 1:
        result.update(
            p95=ordered[math.ceil(0.95 * len(ordered)) - 1],
            p99=ordered[math.ceil(0.99 * len(ordered)) - 1],
            maximum=ordered[-1],
        )
    return result


def percent(before: int | float, after: int | float) -> float:
    need(before != 0, "percentage denominator is zero")
    return 100 * (after / before - 1)


def summarize(values: list[float], plan: dict[str, Any]) -> dict[str, Any]:
    """Recompute paired summaries and nearest-rank bootstrap bounds."""
    need(values, "empty comparison")
    resamples = integer(plan["bootstrap_resamples"], "bootstrap resamples", 1)
    # Seed each summary independently, matching the public report contract.
    rng = random.Random(plan["bootstrap_seed"])
    bootstrap = sorted(
        st.median(rng.choices(values, k=len(values))) for _ in range(resamples)
    )
    lower = math.ceil(0.025 * len(bootstrap)) - 1
    upper = math.ceil(0.975 * len(bootstrap)) - 1
    threshold = plan["review_threshold_percent"]
    return {
        "values": values,
        "median": st.median(values),
        "minimum": min(values),
        "maximum": max(values),
        "bootstrap_median_95": [bootstrap[lower], bootstrap[upper]],
        "flagged_pairs": [i for i, value in enumerate(values) if abs(value) > threshold],
    }


def close(actual: Any, expected: Any, path: str = "analysis") -> None:
    """Compare JSON values while retaining exact shape and numeric values."""
    if isinstance(actual, bool) or isinstance(expected, bool):
        need(actual is expected, f"{path}: boolean differs")
        return
    if isinstance(actual, dict) or isinstance(expected, dict):
        need(isinstance(actual, dict) and isinstance(expected, dict),
             f"{path}: object type differs")
        need(set(actual) == set(expected), f"{path}: object shape differs")
        for key in expected:
            close(actual[key], expected[key], f"{path}.{key}")
        return
    if isinstance(actual, list) or isinstance(expected, list):
        need(isinstance(actual, list) and isinstance(expected, list),
             f"{path}: list type differs")
        need(len(actual) == len(expected), f"{path}: list length differs")
        for index, (left, right) in enumerate(zip(actual, expected)):
            close(left, right, f"{path}[{index}]")
        return
    if isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
        need(math.isclose(float(actual), float(expected), rel_tol=0.0, abs_tol=1e-9),
             f"{path}: number differs: {actual!r} != {expected!r}")
        return
    need(actual == expected, f"{path}: value differs: {actual!r} != {expected!r}")


def packet_plan() -> dict[str, Any]:
    plan = read(P / "plan.json")
    need(isinstance(plan, dict), "plan is not an object")
    for key, expected in {
        "kind": "unchanged-owner-sample-layout-control",
        "cpu": 12,
        "native_processes": 72,
        "allocation_processes": 24,
        "bootstrap_seed": 7338,
        "bootstrap_resamples": 10_000,
        "review_threshold_percent": 5,
        "comparisons": [
            ["archive", "prior"],
            ["prior", "restored-a"],
            ["restored-a", "restored-b"],
        ],
        "allocation_comparisons": [
            ["archive", "prior"],
            ["prior", "restored-a"],
            ["restored-a", "restored-b"],
        ],
    }.items():
        need(plan.get(key) == expected, f"plan changed: {key}")

    schedule = plan.get("schedule")
    need(isinstance(schedule, list) and len(schedule) == 96, "schedule count changed")
    expected_native = {
        (case, arm, repeat)
        for case in CASES
        for arm in NATIVE_ARMS
        for repeat in range(9)
    }
    expected_allocation = {
        (case, arm, repeat)
        for case in CASES
        for arm in NATIVE_ARMS
        for repeat in range(3)
    }
    native_seen: set[tuple[str, str, int]] = set()
    allocation_seen: set[tuple[str, str, int]] = set()
    arm_contract = {
        "archive": ("archive", "legacy"),
        "prior": ("prior", "legacy"),
        "restored-a": ("restored", "legacy"),
        "restored-b": ("restored", "legacy"),
    }
    for index, row in enumerate(schedule):
        need(isinstance(row, dict), f"schedule row {index}: not an object")
        for key in ("lane", "repeat", "case", "arm", "build", "lifecycle", "samples", "warmups"):
            need(key in row, f"schedule row {index}: missing {key}")
        lane = row["lane"]
        key = (row["case"], row["arm"], row["repeat"])
        need(
            tuple(row[key_name] for key_name in ("build", "lifecycle"))
            == arm_contract.get(row["arm"]),
            f"schedule row {index}: arm contract changed",
        )
        if lane == "native":
            need(key in expected_native and key not in native_seen,
                 f"schedule row {index}: invalid or duplicate native row")
            native_seen.add(key)
            need(row["samples"] == 50 and row["warmups"] == 3,
                 f"schedule row {index}: native sampling changed")
        elif lane == "allocation":
            need(key in expected_allocation and key not in allocation_seen,
                 f"schedule row {index}: invalid or duplicate allocation row")
            allocation_seen.add(key)
            need(row["samples"] == 1 and row["warmups"] == 0,
                 f"schedule row {index}: allocation sampling changed")
        else:
            raise AuditError(f"schedule row {index}: unknown lane {lane!r}")
    need(native_seen == expected_native, "native schedule matrix incomplete")
    need(allocation_seen == expected_allocation, "allocation schedule matrix incomplete")
    return plan


def packet_cases() -> dict[str, dict[str, Any]]:
    rows = read(P / "cases.json")
    need(isinstance(rows, list) and [row.get("id") for row in rows] == list(CASES),
         "case order changed")
    result: dict[str, dict[str, Any]] = {}
    for row in rows:
        need(isinstance(row, dict), "case row is not an object")
        case = row.get("id")
        need(case in CASES and case not in result, f"invalid case: {case!r}")
        path = row.get("path")
        need(
            isinstance(path, str)
            and not Path(path).is_absolute()
            and ".." not in Path(path).parts,
            f"{case}: unsafe fixture path",
        )
        fixture = ROOT / path
        need(fixture.is_file() and not fixture.is_symlink(), f"{case}: fixture missing")
        need(
            fixture.stat().st_size == row.get("bytes") and sha(fixture) == row.get("sha256"),
            f"{case}: fixture custody changed",
        )
        result[case] = row
    need(set(result) == set(CASES), "case set changed")
    return result


def preflight_bindings() -> None:
    preflight = read(P / "preflight.json")
    need(preflight.get("status") == "passed", "preflight did not pass")
    need(preflight.get("freeze_sha256") == sha(P / "freeze.json"),
         "preflight freeze binding changed")


def capture_rows(
    plan: dict[str, Any], cases: dict[str, dict[str, Any]]
) -> tuple[list[dict[str, Any]], dict[tuple[str, str, str, int], dict[str, Any]]]:
    captures = P / "captures"
    manifest = read(captures / "manifest.json")
    need(manifest.get("status") == "complete", "capture manifest is not complete")
    need(manifest.get("freeze_sha256") == sha(P / "freeze.json"),
         "capture freeze binding changed")
    need(manifest.get("preflight_sha256") == sha(P / "preflight.json"),
         "capture preflight binding changed")
    runs = manifest.get("runs")
    schedule = plan["schedule"]
    need(isinstance(runs, list) and len(runs) == len(schedule) == 96,
         "capture count changed")
    processes: list[dict[str, Any]] = []
    lookup: dict[tuple[str, str, str, int], dict[str, Any]] = {}
    expected_files = {"manifest.json"}
    previous_end = 0
    output_names: set[str] = set()
    stderr_names: set[str] = set()
    native_samples = 0
    native_processes = 0
    allocation_processes = 0

    for index, (row, expected) in enumerate(zip(runs, schedule)):
        need(isinstance(row, dict), f"capture row {index}: not an object")
        for key, value in expected.items():
            need(row.get(key) == value, f"capture row {index}: schedule field {key} changed")
        need(row.get("command") == command(expected), f"capture row {index}: command changed")
        need(row.get("exit_code") == 0, f"capture row {index}: child failed")
        start = integer(row.get("start_monotonic_ns"), f"capture row {index}: start", 0)
        end = integer(row.get("end_monotonic_ns"), f"capture row {index}: end", 0)
        need(start >= previous_end and end > start, f"capture row {index}: time order changed")
        previous_end = end

        output_name = basename(row.get("output"), f"capture row {index}: output")
        need(output_name.endswith(".json"), f"capture row {index}: output is not JSON")
        stderr_name = output_name.removesuffix(".json") + ".stderr"
        if "stderr" in row:
            stderr_name = basename(row["stderr"], f"capture row {index}: stderr")
            need(stderr_name == output_name.removesuffix(".json") + ".stderr",
                 f"capture row {index}: stderr name changed")
        need(output_name not in output_names and stderr_name not in stderr_names,
             f"capture row {index}: duplicate output or stderr")
        output_names.add(output_name)
        stderr_names.add(stderr_name)
        expected_files.update((output_name, stderr_name))

        output = captures / output_name
        stderr = captures / stderr_name
        need(output.is_file() and not output.is_symlink(),
             f"capture row {index}: output missing")
        need(stderr.is_file() and not stderr.is_symlink(),
             f"capture row {index}: stderr missing")
        need(row.get("sha256") == sha(output), f"capture row {index}: output hash changed")
        need(row.get("stderr_sha256") == sha(stderr),
             f"capture row {index}: stderr hash changed")
        need(stderr.read_bytes() == b"", f"capture row {index}: stderr is not empty")

        report = validate_report(read(output), expected)
        key = (expected["lane"], expected["case"], expected["arm"], expected["repeat"])
        need(key not in lookup, f"capture row {index}: duplicate process key")
        item = {field: expected[field] for field in expected}
        if expected["lane"] == "native":
            native_processes += 1
            times = []
            for sample_index, sample in enumerate(report["samples"]):
                need(sample.get("index") == sample_index,
                     f"capture row {index}: sample index changed")
                phase = sample.get("phase_ns")
                need(isinstance(phase, dict) and set(phase) == {"whole_ns"},
                     f"capture row {index}: timing schema changed")
                times.append(integer(phase["whole_ns"],
                                     f"capture row {index}: whole_ns", 1))
            need(len(times) == 50, f"capture row {index}: native sample count changed")
            native_samples += len(times)
            item["times_ns"] = times
            item["stats"] = statistics(times)
            item["windows"] = {
                name: statistics(times[start:stop]) for name, start, stop in WINDOWS
            }
            item["last10_vs_first10_percent"] = percent(
                st.median(times[:10]), st.median(times[-10:])
            )
        else:
            allocation_processes += 1
            samples = report.get("samples")
            need(isinstance(samples, list) and len(samples) == 1,
                 f"capture row {index}: allocation sample count changed")
            sample = samples[0]
            allocations = sample.get("allocations")
            need(isinstance(allocations, dict) and "whole" in allocations
                 and isinstance(allocations["whole"], dict),
                 f"capture row {index}: allocation schema changed")
            whole = allocations["whole"]
            need(set(whole) == set(ALLOC_FIELDS),
                 f"capture row {index}: allocation fields changed")
            item["allocation"] = {
                field: integer(whole[field], f"capture row {index}: {field}", 0)
                for field in ALLOC_FIELDS
            }
        processes.append(item)
        lookup[key] = item

    need(native_processes == 72 and allocation_processes == 24,
         "capture lane counts changed")
    need(native_samples == 3600, f"native sample count changed: {native_samples}")
    entries = list(captures.iterdir())
    need(
        {entry.name for entry in entries} == expected_files
        and all(entry.is_file() and not entry.is_symlink() for entry in entries),
        "capture directory has unmanifested or invalid files",
    )
    return processes, lookup


def native_comparisons(
    plan: dict[str, Any], lookup: dict[tuple[str, str, str, int], dict[str, Any]]
) -> list[dict[str, Any]]:
    comparisons: list[dict[str, Any]] = []
    for case in CASES:
        for left, right in plan["comparisons"]:
            pairs = [
                (lookup[("native", case, left, repeat)],
                 lookup[("native", case, right, repeat)])
                for repeat in range(9)
            ]
            metrics = {
                metric: summarize(
                    [percent(a["stats"][metric], b["stats"][metric]) for a, b in pairs],
                    plan,
                )
                for metric in TIMING_FIELDS
            }
            windows = {
                name: summarize(
                    [percent(a["windows"][name]["p50"], b["windows"][name]["p50"])
                     for a, b in pairs],
                    plan,
                )
                for name, _start, _stop in WINDOWS
            }
            comparisons.append({"case": case, "left": left, "right": right,
                                "metrics": metrics, "windows": windows})
    return comparisons


def allocation_comparisons(
    plan: dict[str, Any], lookup: dict[tuple[str, str, str, int], dict[str, Any]]
) -> list[dict[str, Any]]:
    comparisons: list[dict[str, Any]] = []
    for case in CASES:
        for left, right in plan["allocation_comparisons"]:
            pairs = [
                (lookup[("allocation", case, left, repeat)],
                 lookup[("allocation", case, right, repeat)])
                for repeat in range(3)
            ]
            fields = {
                field: {
                    "before": [a["allocation"][field] for a, _b in pairs],
                    "after": [b["allocation"][field] for _a, b in pairs],
                    "differences": [b["allocation"][field] - a["allocation"][field]
                                    for a, b in pairs],
                }
                for field in ALLOC_FIELDS
            }
            comparisons.append({"case": case, "left": left, "right": right,
                                "fields": fields})
    return comparisons


def main() -> int:
    # Keep the production-source guard rooted at the real workspace.  The
    # replay harness can replace this module's P in an isolated packet root.
    guard()
    plan = packet_plan()
    cases = packet_cases()
    preflight_bindings()

    processes, lookup = capture_rows(plan, cases)
    comparisons = native_comparisons(plan, lookup)
    allocation = allocation_comparisons(plan, lookup)
    result = {
        "status": "passed",
        "native_processes": 72,
        "allocation_processes": 24,
        "native_samples": 3600,
        "processes": processes,
        "comparisons": comparisons,
        "allocation_comparisons": allocation,
        "scope": "Unchanged production owner; harness layout sensitivity only. Full-window gates retained; no candidate reinstatement.",
    }
    analysis = read(P / "analysis.json")
    close(analysis, result)

    receipt = {
        "schema_version": 1,
        "status": "passed",
        "analysis_match": True,
        "analysis_sha256": sha(P / "analysis.json"),
        "freeze_sha256": sha(P / "freeze.json"),
        "preflight_sha256": sha(P / "preflight.json"),
        "capture_sha256": sha(P / "captures" / "manifest.json"),
        "native_processes": 72,
        "allocation_processes": 24,
        "native_samples": 3600,
        "native_comparisons": 6,
        "allocation_comparisons": 6,
        "bootstrap_seed": plan["bootstrap_seed"],
        "bootstrap_resamples": plan["bootstrap_resamples"],
        "fresh_scope": "No fresh arm; all four arms use the legacy lifecycle and full 50-sample windows.",
    }
    (P / "audit.json").write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n",
                                   encoding="utf-8")
    print("PASS independent 0738 custody, schedule, 3600-sample replay, native bootstrap, and allocation audit")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (AuditError, OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"FAIL: {error}")
        raise SystemExit(1)
