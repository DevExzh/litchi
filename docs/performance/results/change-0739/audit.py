#!/usr/bin/env python3
"""Independently replay the 0739 cross-slide-copy capture matrix.

The capture runner owns process invocation and the report contract.  This
module may use those custody checks, but it reconstructs every process
statistic and group summary from the raw JSON report without importing
``analyze.py``.  It is intentionally usable after the capture directory has
been populated; before then it only needs to remain syntax-checkable.
"""

from __future__ import annotations

import copy
import json
import math
import random
import statistics as st
from pathlib import Path
from typing import Any

from run import P, PHASES, command, guard, projection, read, schedule, sha, verify_freeze


TIMING_FIELDS = ("p50", "mean", "p95", "p99", "maximum")
ALLOC_FIELDS = (
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


class AuditError(Exception):
    """Raised when custody, raw data, or independently derived values differ."""


def need(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def integer(value: Any, label: str, minimum: int | None = None) -> int:
    need(isinstance(value, int) and not isinstance(value, bool), f"{label}: not an integer")
    if minimum is not None:
        need(value >= minimum, f"{label}: below {minimum}")
    return value


def number(value: Any, label: str) -> int | float:
    """Validate a finite JSON number without narrowing wall-clock timestamps."""
    need(
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and math.isfinite(float(value)),
        f"{label}: not a finite number",
    )
    return value


def basename(value: Any, label: str) -> str:
    """Require a manifest path to be a relative single filename."""
    need(isinstance(value, str) and value, f"{label}: empty path")
    path = Path(value)
    need(
        not path.is_absolute()
        and len(path.parts) == 1
        and path.name == value
        and value not in (".", ".."),
        f"{label}: unsafe path",
    )
    return value


def stats(values: list[int | float]) -> dict[str, int | float]:
    """Recompute midpoint, arithmetic mean, and nearest-rank tails."""
    need(values, "empty timing sequence")
    ordered = sorted(values)
    n = len(ordered)
    return {
        "p50": st.median(values),
        "mean": st.mean(values),
        "p95": ordered[math.ceil(0.95 * n) - 1],
        "p99": ordered[math.ceil(0.99 * n) - 1],
        "maximum": ordered[-1],
    }


def percent(before: int | float, after: int | float) -> float:
    need(before != 0, "percentage denominator is zero")
    return 100 * (after / before - 1)


def summary(values: list[int | float]) -> dict[str, Any]:
    """Recompute a process-group summary and its deterministic median bootstrap."""
    need(values, "empty group sequence")
    # Each public summary starts a fresh generator with the declared seed.
    # Keep the order and nearest-rank indices explicit so this calculation is
    # independent of the implementation in analyze.py.
    rng = random.Random(7339)
    bootstrap = sorted(
        st.median(rng.choices(values, k=len(values))) for _ in range(10_000)
    )
    return {
        "median": st.median(values),
        "minimum": min(values),
        "maximum": max(values),
        "bootstrap_median_95": [bootstrap[249], bootstrap[9749]],
        "values": values,
    }


def close(actual: Any, expected: Any, path: str = "analysis") -> None:
    """Compare JSON trees, allowing only a documented floating-point tolerance."""
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
        need(
            math.isclose(float(actual), float(expected), rel_tol=1e-12, abs_tol=1e-9),
            f"{path}: number differs: {actual!r} != {expected!r}",
        )
        return
    need(actual == expected, f"{path}: value differs: {actual!r} != {expected!r}")


def validate_plan() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    plan = read("plan.json")
    need(isinstance(plan, dict), "plan is not an object")
    need(plan.get("cpu") == 12, "plan CPU changed")
    need(plan.get("cases") == [
        "pptx_cross_copy_plain_lifecycle",
        "pptx_cross_copy_media_rich_lifecycle",
    ], "plan case order changed")
    need(plan.get("native") == {"repeats": 9, "samples": 30, "warmups": 3},
         "native plan changed")
    need(plan.get("allocation") == {"repeats": 3, "samples": 1, "warmups": 0,
                                     "interpretation": plan["allocation"]["interpretation"]},
         "allocation plan shape changed")
    rows = schedule()
    need(len(rows) == 24, f"schedule length changed: {len(rows)}")
    need(sum(row["lane"] == "native" for row in rows) == 18,
         "native process count changed")
    need(sum(row["lane"] == "allocation" for row in rows) == 6,
         "allocation process count changed")
    for row in rows:
        if row["lane"] == "native":
            need(row["samples"] == 30 and row["warmups"] == 3,
                 "native schedule sampling changed")
        else:
            need(row["samples"] == 1 and row["warmups"] == 0,
                 "allocation schedule sampling changed")
    return plan, rows


def allocation_copy(report: dict[str, Any], index: int, lane: str) -> dict[str, Any]:
    """Validate and retain the complete operation-scoped allocator envelope."""
    result = report["results"][0]
    allocation = result["operation_metrics"]["allocation"]
    need(isinstance(allocation, dict), f"capture {index}: allocation is not an object")
    expected_status = "measured" if lane == "allocation" else "unavailable"
    need(allocation.get("status") == expected_status,
         f"capture {index}: allocation status changed")
    need(set(allocation) == {"status", "scope", *ALLOC_FIELDS},
         f"capture {index}: allocator field set changed")
    for field in ALLOC_FIELDS:
        value = allocation.get(field)
        need(isinstance(value, dict), f"capture {index}: allocation field {field} missing")
        need(value.get("status") == expected_status,
             f"capture {index}: allocation field {field} status changed")
        if lane == "allocation":
            values = value.get("values")
            need(isinstance(values, list) and len(values) == 1,
                 f"capture {index}: allocation field {field} length changed")
            integer(values[0], f"capture {index}: allocation {field}", 0)
        else:
            need("values" not in value,
                 f"capture {index}: unavailable allocation field has values")
    # The analyzer retains all allocation metadata.  Copy the raw tree so the
    # independent result has the same complete contract, including scopes.
    return copy.deepcopy(allocation)


def capture_rows(
    rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[tuple[str, int, str], dict[str, Any]]]:
    captures = P / "captures"
    need(captures.is_dir() and not captures.is_symlink(), "capture directory missing")
    manifest = read(captures / "manifest.json")
    need(isinstance(manifest, list) and len(manifest) == len(rows),
         "capture manifest length changed")
    processes: list[dict[str, Any]] = []
    lookup: dict[tuple[str, int, str], dict[str, Any]] = {}
    expected_files = {"manifest.json"}
    previous_end = 0

    for index, (row, record) in enumerate(zip(rows, manifest)):
        need(isinstance(record, dict), f"capture {index}: manifest row is not an object")
        for key, value in row.items():
            need(record.get(key) == value, f"capture {index}: schedule field {key} changed")
        expected_output = f"captures/{index:03d}.json"
        need(record.get("command") == command(row, P / "captures" / f"{index:03d}.json"),
             f"capture {index}: command changed")
        need(record.get("exit") == 0, f"capture {index}: child failed")
        start = integer(record.get("monotonic_start"), f"capture {index}: start", 0)
        end = integer(record.get("monotonic_end"), f"capture {index}: end", 0)
        need(previous_end <= start < end, f"capture {index}: monotonic order changed")
        previous_end = end
        number(record.get("started"), f"capture {index}: wall start")
        number(record.get("ended"), f"capture {index}: wall end")
        need(record["started"] <= record["ended"], f"capture {index}: wall order changed")

        files = record.get("files")
        expected_names = {
            expected_output,
            f"captures/{index:03d}.stdout",
            f"captures/{index:03d}.stderr",
        }
        need(isinstance(files, dict) and set(files) == expected_names,
             f"capture {index}: file manifest changed")
        for relative, digest in files.items():
            basename(relative.rsplit("/", 1)[-1], f"capture {index}: file")
            path = P / relative
            need(path.is_file() and not path.is_symlink(),
                 f"capture {index}: file missing: {relative}")
            need(sha(path) == digest, f"capture {index}: file hash changed: {relative}")
        need((P / f"captures/{index:03d}.stderr").read_bytes() == b"",
             f"capture {index}: stderr is not empty")
        expected_files.update(expected_names)

        report = read(f"captures/{index:03d}.json")
        # run.projection is the shared report/oracle contract validator.  No
        # timing statistic below relies on its derived values.
        projected = projection(report, row)
        oracle = read("oracle.json")
        need(projected == oracle[row["case"]],
             f"capture {index}: qualification oracle differs")

        result = report["results"][0]
        source = result["source"]["pptx_cross_copy"]
        item = row.copy()
        key = (row["lane"], row["repeat"], row["case"])
        need(key not in lookup, f"capture {index}: duplicate process key")
        if row["lane"] == "native":
            lifecycle = source["lifecycle_ns"]
            order = result["elapsed_ns"]["sample_order"]
            need(len(lifecycle) == row["samples"] == 30,
                 f"capture {index}: lifecycle sample count changed")
            need(len(order) == len(lifecycle) and sorted(order) == list(range(len(order))),
                 f"capture {index}: sample_order is not a permutation")
            ordered = [0] * len(lifecycle)
            for sample_order, value in zip(order, lifecycle):
                ordered[sample_order] = value
            values = {phase: source[phase] for phase in PHASES}
            for phase, vector in values.items():
                need(isinstance(vector, list) and len(vector) == len(lifecycle),
                     f"capture {index}: {phase} sample count changed")
                for sample, value in enumerate(vector):
                    integer(value, f"capture {index}: {phase}[{sample}]", 1)
            residual = []
            for sample, total in enumerate(lifecycle):
                integer(total, f"capture {index}: lifecycle[{sample}]", 1)
                remainder = total - sum(
                    values[phase][sample] for phase in ("plan_ns", "commit_ns", "publication_ns")
                )
                need(remainder >= 0, f"capture {index}: negative unassigned residual")
                residual.append(remainder)
            values["unassigned_ns"] = residual
            item["stats"] = {phase: stats(vector) for phase, vector in values.items()}
            item["phase_share_percent"] = {
                phase: st.median(
                    100 * value / total
                    for value, total in zip(values[phase], lifecycle)
                )
                for phase in ("plan_ns", "commit_ns", "publication_ns", "unassigned_ns")
            }
            item["ordered_lifecycle_ns"] = ordered
            item["last10_vs_first10_percent"] = percent(
                st.median(ordered[:10]), st.median(ordered[-10:])
            )
        else:
            item["allocations"] = allocation_copy(report, index, row["lane"])
        processes.append(item)
        lookup[key] = item

    need({entry.name for entry in captures.iterdir()} ==
         {name.rsplit("/", 1)[-1] for name in expected_files},
         "capture directory has unmanifested files")
    return processes, lookup


def build_result(
    plan: dict[str, Any],
    rows: list[dict[str, Any]],
    processes: list[dict[str, Any]],
    lookup: dict[tuple[str, int, str], dict[str, Any]],
) -> dict[str, Any]:
    groups = []
    for case in plan["cases"]:
        selected = [
            process for process in processes
            if process["lane"] == "native" and process["case"] == case
        ]
        need(len(selected) == 9, f"{case}: native process count changed")
        phase_names = tuple(selected[0]["stats"])
        metrics = {
            phase: {
                metric: summary([process["stats"][phase][metric] for process in selected])
                for metric in TIMING_FIELDS
            }
            for phase in phase_names
        }
        flags = []
        for phase, metric_values in metrics.items():
            for metric, value in metric_values.items():
                spread = (value["maximum"] / value["minimum"] - 1) * 100
                if spread > 5:
                    flags.append({
                        "phase": phase,
                        "metric": metric,
                        "across_process_spread_percent": spread,
                    })
        groups.append({
            "case": case,
            "metrics_ns": metrics,
            "phase_share_percent": {
                phase: summary([process["phase_share_percent"][phase] for process in selected])
                for phase in selected[0]["phase_share_percent"]
            },
            "spread_flags": flags,
            "order_flags": [
                {
                    "repeat": process["repeat"],
                    "change_percent": process["last10_vs_first10_percent"],
                }
                for process in selected
                if abs(process["last10_vs_first10_percent"]) > 5
            ],
        })

    allocation_comparisons = []
    # There is no timing comparison in this descriptive batch.  Retain the
    # per-process allocator envelopes in ``processes``; this explicit check
    # keeps the six allocation rows bound to their schedule and raw reports.
    for row in rows:
        if row["lane"] == "allocation":
            need(("allocation", row["repeat"], row["case"]) in lookup,
                 "allocation schedule lookup incomplete")

    return {
        "status": "passed",
        "native_processes": 18,
        "native_samples": 540,
        "allocation_processes": 6,
        "processes": processes,
        "groups": groups,
        "scope": "Current source descriptive process baseline. Phase wall-clock fractions are not CPU fractions or causal removable costs. Reopen is outside lifecycle; unassigned includes ingress/snapshots and clock/loop overhead. No speedup or historical comparison.",
    }


def main() -> int:
    # Source custody, freeze binding, command construction, report contract,
    # and qualification oracle validation remain shared with run.py.  All
    # scalar and grouping work above is independent of analyze.py.
    guard()
    verify_freeze()
    plan, rows = validate_plan()
    processes, lookup = capture_rows(rows)
    result = build_result(plan, rows, processes, lookup)
    analysis = read("analysis.json")
    close(analysis, result)

    receipt = {
        "schema_version": 1,
        "status": "passed",
        "analysis_match": True,
        "analysis_sha256": sha(P / "analysis.json"),
        "freeze_sha256": sha(P / "freeze.json"),
        "capture_sha256": sha(P / "captures" / "manifest.json"),
        "native_processes": 18,
        "native_samples": 540,
        "allocation_processes": 6,
        "bootstrap_seed": 7339,
        "bootstrap_resamples": 10_000,
        "bootstrap_ci_indices": [249, 9749],
        "phase_share_definition": "median of paired within-sample phase/lifecycle ratios; unassigned is lifecycle minus plan, commit, and publication",
    }
    (P / "audit.json").write_text(
        json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    print("PASS independent 0739 custody, 24-process replay, 540 native samples, and allocation audit")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (AuditError, OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"FAIL: {error}")
        raise SystemExit(1)
