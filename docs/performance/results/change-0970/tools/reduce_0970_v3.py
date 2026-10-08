#!/usr/bin/env python3
"""Offline reducer for the 0970 baseline affinity diagnostic."""
from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any

from protocol_0970_v3 import (
    BOOTSTRAP_RESAMPLES,
    BOOTSTRAP_SEED,
    EVENTS,
    EXPECTED_ALLOWED_AFFINITY,
    EXPECTED_PIN_CPU,
    MANIFEST_SCHEMA,
    REDUCTION_SCHEMA,
    STAGE_SCHEMA,
    ProtocolError,
    child_summary,
    check_binding,
    check_packet,
    cells,
    finite_number,
    median,
    paired_bootstrap_ratio,
    robust_summary,
    safe_ratio,
    bootstrap_median,
    sha256,
    strict_load_json,
    validate_affinity_controls,
    validate_manifest,
    write_json_exclusive,
)
from run_0970_v3 import check_plan, validate_go


def metric_value(packet: dict[str, Any], sidecars: dict[str, Any], name: str) -> float:
    if name in {"p50_ns", "p95_ns", "p99_ns", "mean_ns", "within_tail"}:
        summary = child_summary(packet)
        value = summary[name]
        if isinstance(value, dict):
            value = value.get("value")
        return finite_number(value, f"child.{name}", positive=True)
    time_row = sidecars["time"]
    if name in time_row:
        return finite_number(time_row[name], f"time.{name}", nonnegative=True)
    if name.startswith("perf."):
        event = name[5:]
        return finite_number(sidecars["perf"]["events"][event]["value"], f"perf.{event}", nonnegative=True)
    raise ProtocolError(f"unknown metric {name}")


def pair_differences(rows: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any]]], metric: str) -> dict[str, Any]:
    differences = []
    ratios = []
    for pinned_packet, unpinned_packet, pinned_sidecars, unpinned_sidecars in rows:
        pinned = metric_value(pinned_packet, pinned_sidecars, metric)
        unpinned = metric_value(unpinned_packet, unpinned_sidecars, metric)
        difference = pinned - unpinned
        if not math.isfinite(difference):
            raise ProtocolError(f"{metric}: non-finite paired difference")
        differences.append(difference)
        if unpinned > 0:
            ratios.append(pinned / unpinned)
    result: dict[str, Any] = {
        "metric": metric,
        "pinned_minus_unpinned": differences,
        "difference_median": median(differences),
        "difference_p25": sorted(differences)[max(0, int(len(differences) * 0.25) - 1)],
        "difference_p75": sorted(differences)[max(0, int(len(differences) * 0.75) - 1)],
    }
    if ratios:
        result["pinned_over_unpinned"] = {
            "median": median(ratios),
            "ratios": ratios,
            "finite": all(math.isfinite(value) for value in ratios),
        }
    else:
        result["pinned_over_unpinned"] = {"median": None, "ratios": [], "finite": False, "reason": "zero_unpinned"}
    return result


def reduce_cell(identity: str, packets: list[dict[str, Any]], sidecars: dict[str, dict[str, Any]]) -> dict[str, Any]:
    by_pair: dict[int, dict[str, tuple[dict[str, Any], dict[str, Any]]]] = {}
    for packet in packets:
        pair = packet["context"]["pair"]
        condition = packet["context"]["condition"]
        by_pair.setdefault(pair, {})
        if condition in by_pair[pair]:
            raise ProtocolError(f"{identity}: duplicate {condition} pair {pair}")
        by_pair[pair][condition] = (packet, sidecars[packet["packet_id"]])
    if set(by_pair) != set(range(1, 13)) or any(set(row) != {"pinned", "unpinned"} for row in by_pair.values()):
        raise ProtocolError(f"{identity}: pair/condition matrix incomplete")
    children: list[dict[str, Any]] = []
    for pair in range(1, 13):
        for condition in ("pinned", "unpinned"):
            packet, _ = by_pair[pair][condition]
            children.append(child_summary(packet))
    between: dict[str, Any] = {}
    for condition in ("pinned", "unpinned"):
        rows = [row for row in children if row["condition"] == condition]
        p50_values = [row["p50_ns"] for row in rows]
        p95_values = [row["p95_ns"] for row in rows]
        p50_summary = robust_summary(p50_values)
        p95_summary = robust_summary(p95_values)
        within_iqrs = []
        for pair in range(1, 13):
            packet, _ = by_pair[pair][condition]
            values = [float(value) for value in packet["report"]["elapsed_ns"]]
            ordered = sorted(values)
            within_iqrs.append(ordered[int(0.75 * (len(ordered) - 1))] - ordered[int(0.25 * (len(ordered) - 1))])
        order_rows = {
            position: robust_summary([row["p50_ns"] for row in rows if row["order_position"] == position])
            for position in ("first", "second")
        }
        between_mean = sum(p50_values) / len(p50_values)
        between_sd = (sum((value - between_mean) ** 2 for value in p50_values) / len(p50_values)) ** 0.5
        within_mean = sum(within_iqrs) / len(within_iqrs)
        within_sd = (sum((value - within_mean) ** 2 for value in within_iqrs) / len(within_iqrs)) ** 0.5
        between[condition] = {
            "p50": p50_summary,
            "p95": p95_summary,
            "within_iqr_median": median(within_iqrs),
            "between_iqr_over_within_iqr": safe_ratio(p50_summary["iqr"], median(within_iqrs), "zero_within_iqr"),
            "between_cv_over_within_cv": safe_ratio(
                safe_ratio(between_sd, between_mean, "zero_between_mean")["value"] or 0.0,
                safe_ratio(within_sd, within_mean, "zero_within_mean")["value"] or 0.0,
                "zero_within_cv",
            ),
            "order_stratified": order_rows,
            "first_child": rows[0],
            "last_child": rows[-1],
        }
    paired_rows = []
    for pair in range(1, 13):
        pinned, pinned_sidecars = by_pair[pair]["pinned"]
        unpinned, unpinned_sidecars = by_pair[pair]["unpinned"]
        paired_rows.append((pinned, unpinned, pinned_sidecars, unpinned_sidecars))
    paired: dict[str, Any] = {}
    for metric in ("p50_ns", "p95_ns", "p99_ns", "within_tail"):
        paired[metric] = paired_bootstrap_ratio(
            [(metric_value(row[0], row[2], metric), metric_value(row[1], row[3], metric)) for row in paired_rows],
            numerator="pinned_over_unpinned",
        )
    for metric in ("user_s", "system_s", "voluntary_context_switches", "involuntary_context_switches", "minor_faults", "major_faults", "max_rss_kb"):
        paired[metric] = pair_differences(paired_rows, metric)
    for event in EVENTS:
        paired[f"perf.{event}"] = pair_differences(paired_rows, f"perf.{event}")
    multiplexed = sorted({
        event
        for _, _, pinned_sidecars, unpinned_sidecars in paired_rows
        for sidecars in (pinned_sidecars, unpinned_sidecars)
        for event in sidecars["perf"]["multiplexed_events"]
    })
    paired["multiplexed_events"] = multiplexed
    paired["whole_process_scope"] = {
        "perf": "perf measures the launched GNU time command and inherited launcher/baseline process tree; perf bookkeeping is excluded",
        "gnu_time": "launcher startup plus same-PID execed historical baseline",
        "harness": "historical baseline Instant samples only",
        "derived_ratios": ["cycles/instructions", "cycles/task-clock"],
        "frequency_claim": False,
    }
    def derived_ratio(numerator: str, denominator: str) -> dict[str, Any]:
        paired_values = []
        for pinned_packet, unpinned_packet, pinned_sidecars, unpinned_sidecars in paired_rows:
            pin_den = metric_value(pinned_packet, pinned_sidecars, denominator)
            unpin_den = metric_value(unpinned_packet, unpinned_sidecars, denominator)
            if pin_den <= 0 or unpin_den <= 0:
                raise ProtocolError(f"{numerator}/{denominator}: non-positive denominator")
            pin_value = metric_value(pinned_packet, pinned_sidecars, numerator) / pin_den
            unpin_value = metric_value(unpinned_packet, unpinned_sidecars, numerator) / unpin_den
            if unpin_value <= 0 or not math.isfinite(pin_value) or not math.isfinite(unpin_value):
                raise ProtocolError(f"{numerator}/{denominator}: invalid derived operand")
            paired_values.append(pin_value / unpin_value)
        result = bootstrap_median(paired_values)
        result["ratios"] = paired_values
        result["scope"] = "whole-process perf telemetry"
        result["frequency_claim"] = False
        return result

    paired["derived_whole_process"] = {
        "cycles_over_instructions": derived_ratio("perf.cycles", "perf.instructions"),
        "cycles_over_task_clock": derived_ratio("perf.cycles", "perf.task-clock"),
        "available": True,
        "note": "descriptive whole-process telemetry; not CPU-frequency evidence and not harness-region attribution",
    }
    # Retain raw per-child telemetry with no inferred zero values.
    telemetry = {}
    for pair in range(1, 13):
        for condition in ("pinned", "unpinned"):
            packet, side = by_pair[pair][condition]
            telemetry[packet["packet_id"]] = {
                "time": side["time"],
                "perf": side["perf"],
                "scope": "whole-process wrappers; harness elapsed_ns is separate",
            }
    return {
        "cell_id": identity,
        "children": children,
        "between_child": between,
        "paired": paired,
        "telemetry": telemetry,
    }


def reduce_capture(plan_path: Path, manifest_path: Path, output_path: Path) -> dict[str, Any]:
    plan = check_plan(plan_path, require_pre_capture=False)
    manifest = strict_load_json(manifest_path.resolve())
    go_row = manifest.get("root_go") or {}
    if not isinstance(go_row, dict) or not isinstance(go_row.get("path"), str) or not isinstance(go_row.get("sha256"), str):
        raise ProtocolError("manifest root GO binding is missing")
    go_binding = validate_go(Path(go_row["path"]), plan_path.resolve(), plan)
    paths = validate_manifest(manifest, plan_path.resolve(), plan, Path(go_binding["path"]), go_binding["sha256"])
    errors: list[str] = []
    if manifest.get("complete") is not True:
        errors.append("capture manifest is not marked complete")
    stage_value = manifest.get("stage")
    if not isinstance(stage_value, str):
        errors.append("capture stage binding is missing")
    else:
        stage_path = Path(stage_value).resolve()
        try:
            stage = strict_load_json(stage_path)
            if stage.get("schema") != STAGE_SCHEMA or stage.get("no_retry") is not True or stage.get("status") != "complete":
                errors.append("capture stage is not complete/no-retry")
            if stage.get("packet_count") != len(paths) or stage.get("expected_packet_count") != len(cells()) * 12 * 2:
                errors.append("capture stage packet count binding drifted")
            if stage.get("allowed_affinity") != EXPECTED_ALLOWED_AFFINITY or stage.get("pin_cpu") != EXPECTED_PIN_CPU:
                errors.append("capture stage audited affinity binding drifted")
            if stage.get("lc_all") != "C" or stage.get("lang") != "C" or stage.get("python_dontwritebytecode") != "1":
                errors.append("capture stage environment binding drifted")
        except (OSError, ProtocolError, ValueError, TypeError, UnicodeError, json.JSONDecodeError) as exc:
            errors.append(f"capture stage invalid: {exc}")
    try:
        controls_path = check_binding(manifest.get("affinity_controls"), "capture affinity controls")
        controls = strict_load_json(controls_path)
        validate_affinity_controls(controls)
    except (OSError, ProtocolError, ValueError, TypeError, UnicodeError, json.JSONDecodeError) as exc:
        errors.append(f"capture affinity controls invalid: {exc}")
    packets: list[dict[str, Any]] = []
    sidecars: dict[str, dict[str, Any]] = {}
    for packet_path in paths:
        try:
            packet = strict_load_json(packet_path)
            parsed = check_packet(packet, plan)
            packets.append(packet)
            sidecars[packet["packet_id"]] = parsed
        except (OSError, ProtocolError, ValueError, KeyError, TypeError, UnicodeError, json.JSONDecodeError) as exc:
            errors.append(f"{packet_path}: {exc}")
    expected_count = len(cells()) * 12 * 2
    if len(paths) != expected_count:
        errors.append(f"packet count {len(paths)} != {expected_count}")
    if len(packets) != len(paths):
        errors.append("one or more packets failed custody/report validation")
    result: dict[str, Any] = {
        "schema": REDUCTION_SCHEMA,
        "status": "incomplete",
        "diagnostic_only": True,
        "qualification": None,
        "decision": None,
        "plan": str(plan_path.resolve()),
        "manifest": str(manifest_path.resolve()),
        "packet_count": len(paths),
        "expected_packet_count": expected_count,
        "raw_packets_retained": True,
        "errors": errors,
        "seed": BOOTSTRAP_SEED,
        "bootstrap_resamples": BOOTSTRAP_RESAMPLES,
        "cells": {},
        "interpretation": None,
    }
    if errors:
        result["interpretation"] = "incomplete custody/oracle/telemetry packet set; no diagnostic conclusion"
        write_json_exclusive(output_path.resolve(), result)
        return result
    grouped: dict[str, list[dict[str, Any]]] = {row["cell_id"]: [] for row in cells()}
    seen: set[str] = set()
    for packet in packets:
        packet_id = packet["packet_id"]
        if packet_id in seen:
            raise ProtocolError("duplicate packet identity")
        seen.add(packet_id)
        grouped[packet["context"]["cell_id"]].append(packet)
    for identity, rows in grouped.items():
        if len(rows) != 24:
            raise ProtocolError(f"{identity}: expected 24 packets")
        result["cells"][identity] = reduce_cell(identity, rows, sidecars)
    result["status"] = "complete"
    result["interpretation"] = "descriptive pinned-versus-unpinned historical-baseline telemetry; no host-cause or qualification claim"
    write_json_exclusive(output_path.resolve(), result)
    return result


def self_test() -> int:
    values = [100.0] * 12
    summary = robust_summary(values)
    if summary["mad"] != 0.0 or summary["iqr"] != 0.0:
        raise SystemExit("constant dispersion summary failed")
    if safe_ratio(summary["iqr"], 0.0, "zero_within_iqr") != {"value": None, "reason": "zero_within_iqr"}:
        raise SystemExit("zero within-IQR ratio failed")
    interval = paired_bootstrap_ratio([(100.0, 100.0)] * 12, numerator="pinned_over_unpinned")
    if not all(math.isfinite(interval[key]) for key in ("median", "lower", "upper")):
        raise SystemExit("constant paired bootstrap failed")
    valid_controls = {
        "schema": "litchi.performance.0970.affinity_controls.v2",
        "allowed": list(EXPECTED_ALLOWED_AFFINITY),
        "pin_cpu": EXPECTED_PIN_CPU,
        "controls": {
            "unpinned": {
                "condition": "unpinned", "before": list(EXPECTED_ALLOWED_AFFINITY),
                "target": list(EXPECTED_ALLOWED_AFFINITY), "after": list(EXPECTED_ALLOWED_AFFINITY),
                "restored": list(EXPECTED_ALLOWED_AFFINITY),
            },
            "pinned": {
                "condition": "pinned", "before": [EXPECTED_PIN_CPU], "target": [EXPECTED_PIN_CPU],
                "after": [EXPECTED_PIN_CPU], "restored": [EXPECTED_PIN_CPU],
            },
        },
    }
    validate_affinity_controls(valid_controls)
    swapped = json.loads(json.dumps(valid_controls))
    swapped["controls"]["pinned"]["target"] = list(EXPECTED_ALLOWED_AFFINITY)
    try:
        validate_affinity_controls(swapped)
    except ProtocolError:
        pass
    else:
        raise SystemExit("swapped affinity-control arms were accepted")
    missing = dict(valid_controls)
    missing.pop("controls")
    try:
        validate_affinity_controls(missing)
    except ProtocolError:
        pass
    else:
        raise SystemExit("missing affinity controls were accepted")
    print(json.dumps({"schema": REDUCTION_SCHEMA, "status": "self_test_pass", "zero_dispersion_null": True, "constant_bootstrap_finite": True, "swapped_affinity_rejected": True, "missing_affinity_rejected": True, "workload_started": False}))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        return self_test()
    if args.plan is None or args.manifest is None or args.output is None:
        print("reducer requires --plan, --manifest, and --output", file=sys.stderr)
        return 2
    try:
        result = reduce_capture(args.plan.resolve(), args.manifest.resolve(), args.output.resolve())
        print(json.dumps({"schema": REDUCTION_SCHEMA, "status": result["status"], "output": str(args.output.resolve()), "decision": None}))
        return 0 if result["status"] in {"complete", "incomplete"} else 2
    except (OSError, ProtocolError, ValueError, KeyError, TypeError, UnicodeError, json.JSONDecodeError) as exc:
        print(f"reducer rejected input: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
