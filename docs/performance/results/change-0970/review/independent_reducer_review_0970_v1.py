#!/usr/bin/env python3
"""Independent read-only replay/check of the 0970 reducer v6 output.

This file intentionally uses only the frozen packet JSON/sidecars and standard
library arithmetic; it does not import the sealed reducer or protocol modules.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import random
import statistics
from pathlib import Path
from typing import Any

SEED = 970001
RESAMPLES = 10_000
EVENTS = ["task-clock", "cycles", "instructions", "context-switches", "cpu-migrations", "page-faults"]
TIME_FIELDS = ["user_s", "system_s", "voluntary_context_switches", "involuntary_context_switches", "minor_faults", "major_faults", "max_rss_kb"]
CELLS = ["docnohf/e1/default", "docnohf/e32/default", "docfloat/e1/default", "docfloat/e32/default"]


class ReviewError(Exception):
    pass


def load(path: Path) -> Any:
    def hook(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ReviewError(f"duplicate JSON key {key}: {path}")
            result[key] = value
        return result
    return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=hook)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def meta(path: Path) -> dict[str, Any]:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def qlinear(values: list[float], q: float) -> float:
    ordered = sorted(float(x) for x in values)
    position = (len(ordered) - 1) * q
    low = math.floor(position)
    high = math.ceil(position)
    if low == high:
        return ordered[low]
    return ordered[low] + (ordered[high] - ordered[low]) * (position - low)


def median(values: list[float]) -> float:
    return qlinear(values, 0.5)


def nearest(values: list[float], q: float) -> float:
    ordered = sorted(float(x) for x in values)
    return ordered[max(1, math.ceil(q * len(ordered))) - 1]


def ratio(numerator: float, denominator: float, reason: str) -> dict[str, Any]:
    if denominator == 0:
        return {"value": None, "reason": reason}
    value = numerator / denominator
    return {"value": value, "reason": None}


def robust(values: list[float]) -> dict[str, Any]:
    med = median(values)
    p25 = nearest(values, 0.25)
    p75 = nearest(values, 0.75)
    p95 = nearest(values, 0.95)
    mad = median([abs(x - med) for x in values])
    return {
        "count": len(values), "median": med, "p25": p25, "p75": p75, "p95": p95,
        "iqr": p75 - p25, "mad": mad,
        "iqr_over_median": ratio(p75 - p25, med, "zero_median"),
        "mad_over_median": ratio(mad, med, "zero_median"),
        "p95_over_median": ratio(p95, med, "zero_median"),
        "percentile_definition": "nearest_rank for p25/p75/p95; linear median",
    }


def population_cv(values: list[float]) -> float:
    mean = sum(values) / len(values)
    sd = math.sqrt(sum((x - mean) ** 2 for x in values) / len(values))
    return sd / mean


def child(packet: dict[str, Any]) -> dict[str, Any]:
    values = [float(x) for x in packet["report"]["elapsed_ns"]]
    p50 = qlinear(values, 0.50)
    p95 = qlinear(values, 0.95)
    p99 = qlinear(values, 0.99)
    return {
        "packet_id": packet["packet_id"], "pair": packet["context"]["pair"],
        "condition": packet["context"]["condition"], "order_position": packet["context"]["order_position"],
        "p50_ns": p50, "p95_ns": p95, "p99_ns": p99, "mean_ns": sum(values) / len(values),
        "min_ns": min(values), "max_ns": max(values), "spread_ns": max(values) - min(values),
        "within_tail": ratio(p95, p50, "zero_p50"),
        "values": values,
    }


def bootstrap_median(values: list[float]) -> dict[str, Any]:
    rng = random.Random(SEED)
    medians = [median([values[rng.randrange(len(values))] for _ in values]) for _ in range(RESAMPLES)]
    return {"median": median(values), "lower": qlinear(medians, 0.025), "upper": qlinear(medians, 0.975), "seed": SEED, "resamples": RESAMPLES}


def paired_ratio(pairs: list[tuple[float, float]]) -> dict[str, Any]:
    ratios = [pin / unpin for pin, unpin in pairs]
    result = bootstrap_median(ratios)
    result["ratios"] = ratios
    return result


def pair_diff(pairs: list[tuple[float, float]], metric: str) -> dict[str, Any]:
    diffs = [pin - unpin for pin, unpin in pairs]
    ordered = sorted(diffs)
    ratios = [pin / unpin for pin, unpin in pairs if unpin > 0]
    result: dict[str, Any] = {
        "metric": metric, "pinned_minus_unpinned": diffs, "difference_median": median(diffs),
        "difference_p25": ordered[max(0, int(len(diffs) * 0.25) - 1)],
        "difference_p75": ordered[max(0, int(len(diffs) * 0.75) - 1)],
    }
    if ratios:
        result["pinned_over_unpinned"] = {"median": median(ratios), "ratios": ratios, "finite": all(math.isfinite(x) for x in ratios)}
    else:
        result["pinned_over_unpinned"] = {"median": None, "ratios": [], "finite": False, "reason": "zero_unpinned"}
    return result


def summary_measure(summary: dict[str, Any], measure: str) -> float | None:
    value = summary[measure]
    if isinstance(value, dict):
        value = value.get("value")
    return None if value is None else float(value)


def dispersion(pairs: list[tuple[float, float]], measure: str) -> dict[str, Any]:
    pin_summary = robust([x[0] for x in pairs])
    unpin_summary = robust([x[1] for x in pairs])
    pin_value = summary_measure(pin_summary, measure)
    unpin_value = summary_measure(unpin_summary, measure)
    point = {"value": None, "reason": "undefined_dispersion"} if pin_value is None or unpin_value is None else ratio(pin_value, unpin_value, f"zero_unpinned_{measure}")
    rng = random.Random(SEED)
    boots: list[float] = []
    for _ in range(RESAMPLES):
        sample = [pairs[rng.randrange(len(pairs))] for _ in pairs]
        pin = summary_measure(robust([x[0] for x in sample]), measure)
        unpin = summary_measure(robust([x[1] for x in sample]), measure)
        if pin is None or unpin is None or unpin == 0:
            return {"point": point, "bootstrap": {"median": None, "lower": None, "upper": None, "seed": SEED, "resamples": RESAMPLES, "reason": f"zero_unpinned_{measure}"}, "measure": measure}
        boots.append(pin / unpin)
    return {"point": point, "bootstrap": {"median": median(boots), "lower": qlinear(boots, 0.025), "upper": qlinear(boots, 0.975), "seed": SEED, "resamples": RESAMPLES}, "measure": measure}


def close(a: Any, b: Any, path: str, errors: list[str], tol: float = 1e-9) -> None:
    if isinstance(a, bool) or isinstance(b, bool):
        if a != b:
            errors.append(f"{path}: {a!r} != {b!r}")
    elif isinstance(a, (int, float)) and isinstance(b, (int, float)):
        if not math.isclose(float(a), float(b), rel_tol=1e-10, abs_tol=tol):
            errors.append(f"{path}: {a!r} != {b!r}")
    elif a != b:
        errors.append(f"{path}: {a!r} != {b!r}")


def compare(actual: Any, expected: Any, path: str, errors: list[str]) -> None:
    if isinstance(expected, dict):
        if not isinstance(actual, dict):
            errors.append(f"{path}: expected object")
            return
        if set(actual) != set(expected):
            errors.append(f"{path}: key set mismatch")
            return
        for key in expected:
            compare(actual[key], expected[key], f"{path}.{key}", errors)
    elif isinstance(expected, list):
        if not isinstance(actual, list) or len(actual) != len(expected):
            errors.append(f"{path}: list shape mismatch")
            return
        for index, item in enumerate(expected):
            compare(actual[index], item, f"{path}[{index}]", errors)
    else:
        close(actual, expected, path, errors)


def parse_perf(path: Path) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    for raw in path.read_text().splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        fields = line.split(",")
        event = fields[2]
        runtime = int(fields[3]) if len(fields) > 3 and fields[3] else None
        percent = float(fields[4]) if len(fields) > 4 and fields[4] else None
        rows[event] = {
            "value": float(fields[0]), "unit": fields[1], "runtime_ns": runtime,
            "percent": percent, "multiplexed": percent is not None and percent < 99.999,
        }
    return {"events": rows, "all_events_present": set(rows) == set(EVENTS), "multiplexed_events": [event for event, row in rows.items() if row["multiplexed"]]}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--base", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    base = args.base.resolve()
    capture = base / "capture/diagnostic-v3"
    output_path = base / "analysis/affinity-diagnostic-v6-final.json"
    reducer_path = base / "tools/reduce_0970_v6.py"
    errors: list[str] = []
    source = reducer_path.read_text(encoding="utf-8")
    static_requirements = {
        "context_adapter": "for name in (\"pair\", \"condition\", \"order_position\")" in source and "adapted[name] = context[name]" in source,
        "population_sample_cv": "sum((value - mean_value) ** 2 for value in values) / len(values)" in source and "within_cv_definition" in source,
        "nearest_rank": "rank = max(1, math.ceil(q * len(ordered)))" in source,
        "iqr_mad_ratios": 'for measure in ("iqr", "mad", "iqr_over_median", "mad_over_median")' in source,
        "whole_process_scope": '"frequency_claim": False' in source and '"derived_whole_process"' in source,
        "diagnostic_only": '"qualification": None' in source and '"decision": None' in source,
        "exclusive_output": "write_json_exclusive(output_path.resolve(), result)" in source,
    }
    for key, value in static_requirements.items():
        if not value:
            errors.append(f"static reducer requirement failed: {key}")

    result = load(output_path)
    if result.get("schema") != "litchi.performance.0970.diagnostic_reduction.v2" or result.get("status") != "complete" or result.get("packet_count") != 96 or result.get("expected_packet_count") != 96 or result.get("errors") != [] or result.get("qualification") is not None or result.get("decision") is not None or result.get("raw_packets_retained") is not True:
        errors.append("v6 top-level status/custody/diagnostic fields mismatch")
    packets = {}
    manifest = load(capture / "capture-manifest-v2.json")
    for packet_path_text in manifest["packets"]:
        packet_path = Path(packet_path_text)
        packet = load(packet_path)
        packets[packet["packet_id"]] = packet
    if len(packets) != 96:
        errors.append(f"raw packet count {len(packets)} != 96")

    children: dict[str, dict[int, dict[str, dict[str, Any]]]] = {}
    sidecars: dict[str, dict[str, Any]] = {}
    for packet in packets.values():
        row = child(packet)
        cell_id = packet["context"]["cell_id"]
        children.setdefault(cell_id, {}).setdefault(row["pair"], {})[row["condition"]] = row
        sidecars[packet["packet_id"]] = {
            "launch": load(Path(packet["sidecars"]["launch"]["path"])),
            "time": load(Path(packet["sidecars"]["time"]["path"])),
            "perf": parse_perf(Path(packet["sidecars"]["perf"]["path"])),
        }
    cells_result = result.get("cells", {})
    if set(cells_result) != set(CELLS):
        errors.append("v6 cell set mismatch")
    child_checks = 0
    between_checks = 0
    bootstrap_checks = 0
    for cell_id in CELLS:
        expected_cell = cells_result.get(cell_id, {})
        expected_children = {row["packet_id"]: row for row in expected_cell.get("children", [])}
        expected_pairs = children.get(cell_id, {})
        if len(expected_children) != 24 or len(expected_pairs) != 12:
            errors.append(f"{cell_id}: child/pair count mismatch")
        for pair in range(1, 13):
            for condition in ("pinned", "unpinned"):
                source_row = expected_pairs.get(pair, {}).get(condition)
                if source_row is None:
                    errors.append(f"{cell_id}: missing {condition} pair {pair}")
                    continue
                actual_row = expected_children.get(source_row["packet_id"])
                if actual_row is None:
                    errors.append(f"{cell_id}: output missing child {source_row['packet_id']}")
                    continue
                expected_child = {key: value for key, value in source_row.items() if key != "values"}
                compare(actual_row, expected_child, f"{cell_id}.children[{source_row['packet_id']}]", errors)
                child_checks += 1
        for condition in ("pinned", "unpinned"):
            rows = [expected_pairs[pair][condition] for pair in range(1, 13)]
            actual_between = expected_cell.get("between_child", {}).get(condition, {})
            for metric, output_key in (("p50_ns", "p50"), ("p95_ns", "p95")):
                expected_summary = robust([row[metric] for row in rows])
                compare(actual_between.get(output_key), expected_summary, f"{cell_id}.between_child.{condition}.{output_key}", errors)
                between_checks += 1
            values = [row["p50_ns"] for row in rows]
            within_iqrs = []
            within_cvs = []
            for row in rows:
                ordered = sorted(row["values"])
                within_iqrs.append(ordered[int(0.75 * (len(ordered) - 1))] - ordered[int(0.25 * (len(ordered) - 1))])
                within_cvs.append(population_cv(row["values"]))
            within_iqr_median = median(within_iqrs)
            within_cv_median = median(within_cvs)
            between_mean = sum(values) / len(values)
            between_sd = math.sqrt(sum((value - between_mean) ** 2 for value in values) / len(values))
            expected_extra = {
                "within_iqr_median": within_iqr_median,
                "between_iqr_over_within_iqr": ratio(robust(values)["iqr"], within_iqr_median, "zero_within_iqr"),
                "within_cv_definition": "median of 12 child population SD(elapsed_ns)/mean(elapsed_ns) values over each child's 80 samples",
                "within_cv_median": within_cv_median,
                "between_cv_population": ratio(between_sd, between_mean, "zero_between_mean"),
                "between_cv_over_within_cv": ratio(between_sd / between_mean, within_cv_median, "zero_within_cv"),
            }
            compare(actual_between, {**{key: actual_between.get(key) for key in ("p50", "p95", "order_stratified", "first_child", "last_child")}, **expected_extra}, f"{cell_id}.between_child.{condition}.extras", errors)
            # Explicitly verify order stratification independently.
            order_actual = actual_between.get("order_stratified", {})
            for position in ("first", "second"):
                expected_order = robust([row["p50_ns"] for row in rows if row["order_position"] == position])
                compare(order_actual.get(position), expected_order, f"{cell_id}.order.{condition}.{position}", errors)

        paired_rows = [(expected_pairs[pair]["pinned"], expected_pairs[pair]["unpinned"]) for pair in range(1, 13)]
        actual_paired = expected_cell.get("paired", {})
        metric_pairs = {
            "p50_ns": [(row[0]["p50_ns"], row[1]["p50_ns"]) for row in paired_rows],
            "p95_ns": [(row[0]["p95_ns"], row[1]["p95_ns"]) for row in paired_rows],
            "p99_ns": [(row[0]["p99_ns"], row[1]["p99_ns"]) for row in paired_rows],
            "within_tail": [(row[0]["within_tail"]["value"], row[1]["within_tail"]["value"]) for row in paired_rows],
        }
        for metric, pairs in metric_pairs.items():
            compare(actual_paired.get(metric), paired_ratio(pairs), f"{cell_id}.paired.{metric}", errors)
            bootstrap_checks += 1
        for metric in TIME_FIELDS:
            pairs = [(sidecars[row[0]["packet_id"]]["time"][metric], sidecars[row[1]["packet_id"]]["time"][metric]) for row in paired_rows]
            compare(actual_paired.get(metric), pair_diff(pairs, metric), f"{cell_id}.paired.{metric}", errors)
        for event in EVENTS:
            pairs = [(sidecars[row[0]["packet_id"]]["perf"]["events"][event]["value"], sidecars[row[1]["packet_id"]]["perf"]["events"][event]["value"]) for row in paired_rows]
            compare(actual_paired.get(f"perf.{event}"), pair_diff(pairs, f"perf.{event}"), f"{cell_id}.paired.perf.{event}", errors)
        dispersion_actual = actual_paired.get("between_child_dispersion", {})
        for metric_name, metric_pairs_name in (("p50", "p50_ns"), ("p95", "p95_ns")):
            pairs = [(row[0][metric_pairs_name], row[1][metric_pairs_name]) for row in paired_rows]
            for measure in ("iqr", "mad", "iqr_over_median", "mad_over_median"):
                compare(dispersion_actual.get(metric_name, {}).get(measure), dispersion(pairs, measure), f"{cell_id}.paired.dispersion.{metric_name}.{measure}", errors)
                bootstrap_checks += 1
        expected_affinity = {"pinned": [[31]], "unpinned": [list(range(32))]}
        compare(actual_paired.get("observed_affinity"), {"definition": "Categorical actual_affinity from each validated launch sidecar; no numeric affinity difference is inferred.", **expected_affinity}, f"{cell_id}.paired.observed_affinity", errors)
        compare(actual_paired.get("multiplexed_events"), [], f"{cell_id}.paired.multiplexed_events", errors)
        # Derived whole-process ratios and telemetry are checked from raw sidecars.
        for numerator, denominator, key in (("cycles", "instructions", "cycles_over_instructions"), ("cycles", "task-clock", "cycles_over_task_clock")):
            pairs = []
            for pin, unpin in paired_rows:
                pin_rows = sidecars[pin["packet_id"]]["perf"]["events"]
                unpin_rows = sidecars[unpin["packet_id"]]["perf"]["events"]
                pairs.append((pin_rows[numerator]["value"] / pin_rows[denominator]["value"], unpin_rows[numerator]["value"] / unpin_rows[denominator]["value"]))
            compare(actual_paired.get("derived_whole_process", {}).get(key), {**bootstrap_median([x / y for x, y in pairs]), "ratios": [x / y for x, y in pairs], "scope": "whole-process perf telemetry", "frequency_claim": False}, f"{cell_id}.derived.{key}", errors)
            bootstrap_checks += 1
        if actual_paired.get("derived_whole_process", {}).get("available") is not True or actual_paired.get("derived_whole_process", {}).get("note") != "descriptive whole-process telemetry; not CPU-frequency evidence and not harness-region attribution":
            errors.append(f"{cell_id}: derived whole-process scope note mismatch")
        for packet_id, side in sidecars.items():
            if packet_id not in expected_cell.get("telemetry", {}):
                continue
            expected_telemetry = {"launch": side["launch"], "time": side["time"], "perf": side["perf"], "scope": "whole-process wrappers; harness elapsed_ns is separate"}
            compare(expected_cell["telemetry"][packet_id], expected_telemetry, f"{cell_id}.telemetry.{packet_id}", errors)
    review = {
        "schema": "litchi.performance.0970.independent_reducer_review.v1",
        "generated_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "status": "pass" if not errors else "blocked",
        "scope": "Independent read-only arithmetic replay against frozen raw packets and reducer v6 output; no workload/build and no shared reducer import.",
        "bindings": {
            "reducer": meta(reducer_path), "output": meta(output_path),
            "manifest": meta(capture / "capture-manifest-v2.json"),
            "raw_audit_receipt": meta(base / "review/raw-independent-audit-receipt-0970-v1.json"),
            "correction_receipt": meta(base / "receipts/reducer-correction-0970-v3.json"),
        },
        "static_fix_checks": static_requirements,
        "replay_checks": {
            "raw_packets": len(packets), "cells": len(CELLS), "children_recomputed": child_checks,
            "between_summaries_recomputed": between_checks, "bootstrap_groups_recomputed": bootstrap_checks,
            "output_status_complete": result.get("status") == "complete", "output_errors_empty": result.get("errors") == [],
            "qualification_null": result.get("qualification") is None, "decision_null": result.get("decision") is None,
            "raw_packets_retained": result.get("raw_packets_retained") is True,
        },
        "definition_verification": {
            "context_adapter": "context.pair/condition/order_position copied before child summary",
            "child_percentiles": "linear interpolation at (n-1)*q for p50/p95/p99; within_tail p95/p50",
            "between_percentiles": "nearest-rank ceil(q*n), one-indexed, for p25/p75/p95; linear median",
            "within_cv": "per-child population SD/mean over 80 elapsed samples, then linear median across 12 children",
            "between_cv": "population SD/mean across 12 child p50 centers",
            "within_iqr": "v6 retained order-statistic indices floor(.75*(n-1))-floor(.25*(n-1)) for each 80-sample child",
            "dispersion": "pinned/unpinned aggregate IQR, MAD, IQR/median, MAD/median point ratios with paired bootstrap",
            "telemetry": "raw launch/time/perf retained; whole-process scope explicit; derived ratios frequency_claim false",
            "zero_handling": "null with explicit reason; no zero substitution",
        },
        "interpretation": "Pass resolves the prior reducer adapter/tooling failure for this frozen capture. The result remains descriptive historical-baseline affinity telemetry and makes no host-cause, candidate qualification, optimization, or adoption claim.",
        "errors": errors,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(review, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": review["status"], "errors": len(errors), "output": str(args.output)}, sort_keys=True))
    return 0 if not errors else 2


if __name__ == "__main__":
    raise SystemExit(main())
