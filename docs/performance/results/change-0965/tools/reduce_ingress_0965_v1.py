#!/usr/bin/env python3
"""Reduce one completed 0965 packet into descriptive ingress diagnostics.

The reducer is deliberately separate from the runner's arithmetic.  It
rechecks plan, custody, launch, terminal, report-oracle, affinity, cleanup,
sample-count, and packet-path bindings before calculating nearest-rank
descriptors and pair ratios.  It never reads historical sample vectors and
never emits an adoption or screen decision.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import os
from pathlib import Path
import statistics
from collections import defaultdict
from typing import Any, Iterable


ROOT = Path("/home/zhuhe/litchi-performance-0965").resolve()
PLAN_SCHEMA = "litchi.performance.0965.ingress-diagnostic-plan.v1"
RUNNER_SCHEMA = "litchi.performance.0965.diagnostic-capture.v1"
REDUCER_SCHEMA = "litchi.performance.0965.ingress-diagnostic-reduction.v1"
DESIGN_SCHEMA = "prospective-design-v3"
EXPECTED_DESIGN_SHA = "0dbff34e930aec3d63350ebf6cce9d0c8a86afdcd8f176f50e52a7e6fe2c06bf"
CPU_MASK = [12, 13, 14, 15]
WARMUP = 20
SAMPLES = 256
CASES = ("medium-one-percent", "dense-sparse-one-percent", "medium-one-edit")
PAIR_SCHEDULE = {
    "medium-one-percent": ("RR", "RC", "CC", "CR"),
    "dense-sparse-one-percent": ("RC", "CC", "CR", "RR"),
    "medium-one-edit": ("CC", "CR", "RR", "RC"),
}
PHASES = ("ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns")
QUANTILES = (0.50, 0.95, 0.99)
Q_NAMES = ("p50", "p95", "p99")
TIME_KEYS = {
    "Maximum resident set size (kbytes)": "max_rss_kib",
    "User time (seconds)": "user_seconds",
    "System time (seconds)": "system_seconds",
    "Major (requiring I/O) page faults": "major_page_faults",
    "Minor (reclaiming a frame) page faults": "minor_page_faults",
    "Voluntary context switches": "voluntary_context_switches",
    "Involuntary context switches": "involuntary_context_switches",
}


class ReductionError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReductionError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as exc:
        raise ReductionError(f"cannot hash {path}: {exc}") from exc
    return digest.hexdigest()


def descriptor(path: Path, executable: bool | None = None) -> dict[str, Any]:
    path = path.expanduser().resolve()
    require(path.is_file() and not path.is_symlink(), f"regular file required: {path}")
    mode = path.stat().st_mode & 0o777
    is_executable = bool(mode & 0o111)
    if executable is not None:
        require(is_executable is executable, f"executable mismatch: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "mode": mode, "executable": is_executable, "sha256": sha256(path)}


def bind(value: Any, label: str, executable: bool | None = None) -> dict[str, Any]:
    require(isinstance(value, dict) and isinstance(value.get("path"), str), f"{label} binding missing")
    actual = descriptor(Path(value["path"]), executable)
    for key in ("path", "bytes", "mode", "executable", "sha256"):
        if key in value:
            require(actual[key] == value[key], f"{label}.{key} changed")
    return actual


def read_json(path: Path) -> Any:
    path = path.expanduser().resolve()
    require(path.is_file() and not path.is_symlink(), f"JSON missing: {path}")
    require(path.stat().st_size <= 128 * 1024 * 1024, f"JSON too large: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ReductionError(f"invalid JSON {path}: {exc}") from exc


def nearest_rank(values: Iterable[int | float], q: float) -> int | float:
    ordered = sorted(values)
    require(bool(ordered), "empty quantile vector")
    rank = max(1, math.ceil(q * len(ordered)))
    return ordered[rank - 1]


def compact(value: int | float) -> int | float:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def quantiles(values: Iterable[int | float]) -> dict[str, int | float]:
    vector = list(values)
    return {name: compact(nearest_rank(vector, q)) for name, q in zip(Q_NAMES, QUANTILES)}


def iso_seconds(value: str) -> float:
    try:
        return dt.datetime.fromisoformat(value).timestamp()
    except (TypeError, ValueError) as exc:
        raise ReductionError(f"invalid timestamp {value!r}") from exc


def parse_time(path: Path) -> dict[str, float | int]:
    values: dict[str, float | int] = {}
    for line in path.read_text(errors="replace").splitlines():
        if ":" not in line:
            continue
        key, raw = line.split(":", 1)
        name = TIME_KEYS.get(key.strip())
        if name is None:
            continue
        token = raw.strip().split()[0] if raw.strip() else ""
        try:
            value: float | int = float(token) if name in {"user_seconds", "system_seconds"} else int(token)
        except ValueError as exc:
            raise ReductionError(f"invalid time metric {name}") from exc
        require(math.isfinite(float(value)) and value >= 0, f"invalid time metric {name}")
        values[name] = value
    require(set(TIME_KEYS.values()) <= values.keys(), "complete time metric set missing")
    require(type(values.get("max_rss_kib")) is int and values["max_rss_kib"] > 0, "positive max RSS missing")
    return values


def schedule_rows(plan: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    case_by_id = {case["case_id"]: case for case in plan["cases"]}
    for case_id in CASES:
        for pair_index, pair_type in enumerate(PAIR_SCHEDULE[case_id]):
            pair_id = f"{case_id}__pair{pair_index:02d}__{pair_type}"
            for position, role in enumerate(pair_type):
                case = case_by_id[case_id]
                rows.append({
                    "child_id": f"native__{pair_id}__p{position}", "lane": "native", "case_id": case_id,
                    "pair_id": pair_id, "pair_type": pair_type, "pair_index": pair_index, "position": position,
                    "role": role, "binary_role": "baseline" if role == "R" else "candidate",
                    "comparison_id": "retained-baseline-vs-retained-candidate", "shape": case["shape"],
                    "workload": case["workload"], "workers": 1, "warmup_iterations": WARMUP,
                    "measured_samples": SAMPLES, "cpu_mask": list(CPU_MASK),
                })
    return rows


def validate_plan(plan_path: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    plan_path = plan_path.expanduser().resolve()
    plan = read_json(plan_path)
    require(plan.get("schema") == PLAN_SCHEMA and plan.get("status") == "authorized", "reduction requires an authorized 0965 plan")
    require(plan.get("scope", {}).get("children") == 24 and plan.get("scope", {}).get("measured_samples") == SAMPLES, "plan schedule changed")
    plan_binding = descriptor(plan_path)
    runner = bind(plan.get("runner", {}).get("descriptor"), "runner", True)
    reducer = bind(plan.get("reducer", {}).get("descriptor"), "reducer", True)
    require(plan.get("runner", {}).get("schema") == RUNNER_SCHEMA, "runner schema changed")
    require(plan.get("reducer", {}).get("schema") == REDUCER_SCHEMA, "reducer schema changed")
    require(plan.get("analysis", {}).get("adoption_allowed") is False and plan.get("analysis", {}).get("screen_reopened") is False, "analysis policy changed")
    require(plan.get("diagnostic_rule", {}).get("frozen_0964_screen_rule_relaxed") is False, "frozen screen rule changed")
    design = plan.get("design", {})
    require(design.get("schema") == DESIGN_SCHEMA and design.get("status") == "sealed", "sealed design v3 binding missing")
    design_binding = bind(design.get("descriptor"), "design v3")
    require(design_binding["sha256"] == EXPECTED_DESIGN_SHA, "sealed design v3 hash changed")
    bind(plan.get("capture", {}).get("host_snapshot_before"), "host-before")
    inherited = plan.get("inherited_0964", {})
    old_plan = bind(inherited.get("plan"), "inherited plan")
    bind(inherited.get("raw_audit"), "inherited audit")
    require(inherited.get("rejection_remains_binding") is True, "inherited rejection policy changed")
    source = plan.get("source", {})
    for key in ("baseline_manifest", "candidate_manifest", "candidate_allowlist", "baseline_binding", "candidate_freeze", "candidate_quality", "candidate_build_result", "required_perf_baseline_lock", "candidate_perf_baseline_lock", "candidate_root_lock_overlay"):
        bind(source.get(key), f"source.{key}")
    for key in ("R", "C"):
        bind(plan.get("binaries", {}).get(key), f"binary {key}", True)
    require(plan.get("binaries", {}).get("R", {}).get("sha256") == "5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3", "baseline hash changed")
    require(plan.get("binaries", {}).get("C", {}).get("sha256") == "d21863a6613800391a01914865fc4eaecdf54ad3061447be8b6be7d47d0dc910", "candidate hash changed")
    require(plan.get("cases") and [case.get("case_id") for case in plan["cases"]] == list(CASES), "case order changed")
    contract = plan.get("report_contract", {}).get("native", {})
    require(contract.get("schema") == "litchi.xlsx-managed-compression-benchmark.v1" and contract.get("warmup_iterations") == WARMUP and contract.get("measured_samples") == SAMPLES, "report contract changed")
    expected = schedule_rows(plan)
    require(plan.get("schedule", {}).get("children") == expected, "child schedule changed")
    require(plan.get("capture", {}).get("host_snapshot_after", {}).get("status") == "required_after_terminal", "host-after policy changed")
    return plan, {"plan": plan_binding, "runner": runner, "reducer": reducer, "old_plan": old_plan}


def validate_report(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any]) -> list[dict[str, Any]]:
    contract = plan["report_contract"]["native"]
    require(report.get("schema") == contract["schema"] and report.get("status") == "completed", f"{row['child_id']} report schema/status mismatch")
    for key, expected in (("shape", row["shape"]), ("workload", row["workload"]), ("workers", 1), ("warmup_iterations", WARMUP), ("measured_samples", SAMPLES)):
        require(report.get(key) == expected, f"{row['child_id']} report root {key} changed")
    require(report.get("performance_claim") == contract["performance_claim"], f"{row['child_id']} performance claim changed")
    case_cfg = next(case for case in plan["cases"] if case["case_id"] == row["case_id"])
    case = report.get("case")
    require(isinstance(case, dict), f"{row['child_id']} case missing")
    expected_case = {
        "id": f"{row['shape']}-{row['workload']}-workers1", "shape": row["shape"], "workload": row["workload"], "workers": 1,
        "update_count": case_cfg["update_count"], "source_bytes": case_cfg["source_bytes"], "source_sha256": case_cfg["source_sha256"],
        "output_bytes": case_cfg["output_bytes"], "output_sha256": case_cfg["output_sha256"], "semantic_sha256": case_cfg["semantic_sha256"],
        "expected_cpu_tasks": case_cfg["expected_cpu_tasks"], "admission_expected_worker_permits": 0, "admission_observed_worker_permits": 0,
        "admission_cpu_tasks_used": case_cfg["expected_cpu_tasks"], "preflight_actual_commit_inverse_verified": True,
        "context_limits": case_cfg["context_limits"], "sink_capacity_bytes": case_cfg["context_limits"]["output_bytes"],
        "sink_max_write_bytes": 65536, "warmup_iterations": WARMUP, "measured_samples": SAMPLES,
    }
    for key, value in expected_case.items():
        require(case.get(key) == value, f"{row['child_id']} case oracle {key} changed")
    samples = case.get("samples")
    require(isinstance(samples, list) and len(samples) == SAMPLES, f"{row['child_id']} sample count changed")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("sample") == index, f"{row['child_id']} sample index changed")
        values = []
        for phase in PHASES:
            value = sample.get(phase)
            require(type(value) is int and value > 0, f"{row['child_id']} {phase} invalid")
            values.append(value)
        require(sum(values[:4]) <= values[4], f"{row['child_id']} phase sum exceeds total")
        for key, value in (("output_bytes", case_cfg["output_bytes"]), ("output_bytes_used", case_cfg["output_bytes"]), ("cpu_tasks_used", case_cfg["expected_cpu_tasks"]), ("output_sha256", case_cfg["output_sha256"]), ("semantic_sha256", case_cfg["semantic_sha256"])):
            require(sample.get(key) == value, f"{row['child_id']} sample oracle {key} changed")
        require(sample.get("fixture_inverse_restored_exact_source") is True, f"{row['child_id']} inverse oracle failed")
        require(type(sample.get("sink_write_calls")) is int and sample["sink_write_calls"] > 0, f"{row['child_id']} sink write calls invalid")
        require(type(sample.get("sink_largest_write")) is int and 0 < sample["sink_largest_write"] <= 65536, f"{row['child_id']} sink bound changed")
        resources = sample.get("resources_after_drop")
        require(isinstance(resources, dict) and all(resources.get(name) == 0 for name in ("memory", "workers", "io_concurrency", "objects")), f"{row['child_id']} resources not clean")
    return samples


def validate_child(packet: Path, row: dict[str, Any], plan: dict[str, Any], plan_binding: dict[str, Any], index: int) -> dict[str, Any]:
    child = packet / "children" / row["child_id"]
    require(child.is_dir() and not child.is_symlink(), f"missing child {row['child_id']}")
    start_path = child / "start.json"
    launch_path = child / "launch.json"
    terminal_path = child / "terminal.json"
    validation_path = child / "validation.json"
    report_path = child / "output" / "report.json"
    for path in (start_path, launch_path, terminal_path, validation_path, report_path, child / "time-v.txt"):
        require(path.is_file() and not path.is_symlink(), f"missing child artifact {path}")
    start = read_json(start_path)
    launch = read_json(launch_path)
    terminal = read_json(terminal_path)
    validation = read_json(validation_path)
    report = read_json(report_path)
    require(start.get("schema") == RUNNER_SCHEMA + ".child-start" and start.get("index") == index, f"{row['child_id']} start schema/index changed")
    require(start.get("child") == row, f"{row['child_id']} planned child mismatch")
    require(start.get("plan") == plan_binding, f"{row['child_id']} plan binding changed")
    binary_key = "R" if row["role"] == "R" else "C"
    require(start.get("binary") == bind(plan["binaries"][binary_key], f"{row['child_id']} binary", True), f"{row['child_id']} binary receipt changed")
    require(launch.get("observed_affinity") == CPU_MASK and launch.get("expected_affinity") == CPU_MASK and launch.get("affinity_verified_before_wait") is True, f"{row['child_id']} affinity contract failed")
    require(terminal.get("schema") == RUNNER_SCHEMA + ".child-terminal" and terminal.get("status") == "completed" and terminal.get("return_code") == 0, f"{row['child_id']} terminal failed")
    require(terminal.get("terminal_before_report_parse") is True and terminal.get("run", {}).get("reaped") is True and terminal.get("run", {}).get("process_group_clear") is True, f"{row['child_id']} cleanup/terminal contract failed")
    require(terminal.get("run", {}).get("pid") == launch.get("pid") and terminal.get("run", {}).get("process_group_id") == launch.get("pid"), f"{row['child_id']} process identity changed")
    require(validation.get("schema") == RUNNER_SCHEMA + ".child-validation" and validation.get("status") == "passed", f"{row['child_id']} validation failed")
    require(validation.get("plan") == plan_binding, f"{row['child_id']} validation plan binding changed")
    actual_report = descriptor(report_path)
    require(validation.get("report") == actual_report, f"{row['child_id']} validation/report hash changed")
    samples = validate_report(report, row, plan)
    time_metrics = parse_time(child / "time-v.txt")
    require(validation.get("time_metrics") == time_metrics, f"{row['child_id']} time metrics changed")
    start_time = iso_seconds(start.get("started_utc"))
    terminal_time = iso_seconds(terminal.get("ended_utc"))
    require(terminal_time >= start_time, f"{row['child_id']} terminal precedes start")
    return {
        "row": row, "child": child, "start": start, "launch": launch, "terminal": terminal,
        "validation": validation, "report": report, "report_path": actual_report, "samples": samples,
        "time_metrics": time_metrics, "start_seconds": start_time, "terminal_seconds": terminal_time,
    }


def child_stats(samples: list[dict[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for phase in PHASES:
        values = [sample[phase] for sample in samples]
        q = quantiles(values)
        ordered = sorted(((value, index) for index, value in enumerate(values)), key=lambda item: (-item[0], item[1]))
        top = ordered[:5]
        result[phase] = {"quantiles_ns": q, "max_ns": max(values), "mean_ns": compact(statistics.fmean(values)), "p99_sample_indices": [index for value, index in ordered if value == q["p99"]], "top5_mean_ns": compact(statistics.fmean(value for value, _ in top)), "top5_sample_indices": [index for _, index in top]}
    return result


def ratio(numerator: int | float, denominator: int | float) -> float:
    require(denominator > 0, "ratio denominator is not positive")
    value = float(numerator) / float(denominator)
    require(math.isfinite(value) and value >= 0, "ratio is not finite")
    return value


def geometric_mean(values: list[float]) -> float:
    require(values and all(value > 0 and math.isfinite(value) for value in values), "geometric mean requires positive finite values")
    return math.exp(statistics.fmean(math.log(value) for value in values))


def pair_summary(first: dict[str, Any], second: dict[str, Any], pair: dict[str, Any]) -> dict[str, Any]:
    first_stats = first["stats"]
    second_stats = second["stats"]
    cross = pair["pair_type"] in {"RC", "CR"}
    phase_result: dict[str, Any] = {}
    for phase in PHASES:
        first_q = first_stats[phase]["quantiles_ns"]
        second_q = second_stats[phase]["quantiles_ns"]
        if cross:
            baseline = first if first["row"]["role"] == "R" else second
            candidate = second if first["row"]["role"] == "R" else first
            baseline_q = baseline["stats"][phase]["quantiles_ns"]
            candidate_q = candidate["stats"][phase]["quantiles_ns"]
            q_ratio = {name: ratio(candidate_q[name], baseline_q[name]) for name in Q_NAMES}
            q_delta = {name: candidate_q[name] - baseline_q[name] for name in Q_NAMES}
            baseline_top5 = baseline["stats"][phase]["top5_mean_ns"]
            candidate_top5 = candidate["stats"][phase]["top5_mean_ns"]
            top5_ratio = ratio(candidate_top5, baseline_top5)
            phase_result[phase] = {
                "baseline_child": baseline["row"]["child_id"], "candidate_child": candidate["row"]["child_id"],
                "baseline_quantiles_ns": baseline_q, "candidate_quantiles_ns": candidate_q,
                "candidate_over_baseline_ratio": q_ratio, "candidate_minus_baseline_ns": q_delta,
                "baseline_top5_mean_ns": baseline_top5, "candidate_top5_mean_ns": candidate_top5,
                "candidate_over_baseline_top5_ratio": top5_ratio,
                "candidate_p99_above_1_05": q_ratio["p99"] > 1.05,
                "candidate_p95_above_1_05": q_ratio["p95"] > 1.05,
                "candidate_top5_above_1_05": top5_ratio > 1.05,
                "baseline_p99_sample_indices": baseline["stats"][phase]["p99_sample_indices"],
                "candidate_p99_sample_indices": candidate["stats"][phase]["p99_sample_indices"],
                "baseline_top5_sample_indices": baseline["stats"][phase]["top5_sample_indices"],
                "candidate_top5_sample_indices": candidate["stats"][phase]["top5_sample_indices"],
            }
        else:
            directional = {name: ratio(second_q[name], first_q[name]) for name in Q_NAMES}
            absolute = {name: max(directional[name], ratio(first_q[name], second_q[name])) for name in Q_NAMES}
            first_top5 = first["stats"][phase]["top5_mean_ns"]
            second_top5 = second["stats"][phase]["top5_mean_ns"]
            top5_directional = ratio(second_top5, first_top5)
            top5_absolute = max(top5_directional, ratio(first_top5, second_top5))
            phase_result[phase] = {
                "first_child": first["row"]["child_id"], "second_child": second["row"]["child_id"],
                "first_quantiles_ns": first_q, "second_quantiles_ns": second_q,
                "second_over_first_ratio": directional, "absolute_ratio": absolute,
                "first_top5_mean_ns": first_top5, "second_top5_mean_ns": second_top5,
                "second_over_first_top5_ratio": top5_directional, "absolute_top5_ratio": top5_absolute,
                "absolute_p99_above_1_05": absolute["p99"] > 1.05,
                "absolute_p95_above_1_05": absolute["p95"] > 1.05,
                "absolute_top5_above_1_05": top5_absolute > 1.05,
                "first_p99_sample_indices": first["stats"][phase]["p99_sample_indices"],
                "second_p99_sample_indices": second["stats"][phase]["p99_sample_indices"],
                "first_top5_sample_indices": first["stats"][phase]["top5_sample_indices"],
                "second_top5_sample_indices": second["stats"][phase]["top5_sample_indices"],
            }
    return {"pair_id": pair["pair_id"], "case_id": pair["case_id"], "pair_type": pair["pair_type"], "children": [first["row"]["child_id"], second["row"]["child_id"]], "phases": phase_result}


def classify(case_id: str, pairs: dict[str, dict[str, Any]]) -> dict[str, Any]:
    rc = pairs["RC"]["phases"]["ingress_ns"]["candidate_over_baseline_ratio"]["p99"]
    cr = pairs["CR"]["phases"]["ingress_ns"]["candidate_over_baseline_ratio"]["p99"]
    rr = pairs["RR"]["phases"]["ingress_ns"]["absolute_ratio"]["p99"]
    cc = pairs["CC"]["phases"]["ingress_ns"]["absolute_ratio"]["p99"]
    rc_p95 = pairs["RC"]["phases"]["ingress_ns"]["candidate_over_baseline_ratio"]["p95"]
    cr_p95 = pairs["CR"]["phases"]["ingress_ns"]["candidate_over_baseline_ratio"]["p95"]
    rc_top5 = pairs["RC"]["phases"]["ingress_ns"]["candidate_over_baseline_top5_ratio"]
    cr_top5 = pairs["CR"]["phases"]["ingress_ns"]["candidate_over_baseline_top5_ratio"]
    order_spread = max(rc, cr) / min(rc, cr)
    flags = {
        "RC_p99_above_1_05": rc > 1.05,
        "CR_p99_above_1_05": cr > 1.05,
        "RR_absolute_p99_above_1_05": rr > 1.05,
        "CC_absolute_p99_above_1_05": cc > 1.05,
    }
    controls_stable = rr <= 1.05 and cc <= 1.05
    if flags["RC_p99_above_1_05"] and flags["CR_p99_above_1_05"] and controls_stable:
        label = "concordant"
    elif not flags["RC_p99_above_1_05"] and not flags["CR_p99_above_1_05"] and controls_stable:
        label = "nonreproduced"
    else:
        label = "mixed"
    return {
        "case_id": case_id, "label": label, "threshold": 1.05,
        "cross_order_p99_ratios": {"RC": rc, "CR": cr},
        "cross_order_p95_ratios": {"RC": rc_p95, "CR": cr_p95},
        "cross_order_top5_ratios": {"RC": rc_top5, "CR": cr_top5},
        "cross_order_spread": order_spread,
        "same_binary_absolute_p99_ratios": {"RR": rr, "CC": cc},
        "flags": flags,
        "diagnostic_only": True,
        "source_build_host_causal_attribution": False,
    }


def reduce_packet(plan_path: Path, packet_path: Path) -> dict[str, Any]:
    plan, bindings = validate_plan(plan_path)
    packet = packet_path.expanduser().resolve()
    require(packet.is_dir() and not packet.is_symlink(), f"packet missing: {packet}")
    capture_start_path = packet / "capture-start.json"
    capture_terminal_path = packet / "capture-terminal.json"
    completed_path = packet / "completed-packet.json"
    capture_start = read_json(capture_start_path)
    capture_terminal = read_json(capture_terminal_path)
    completed = read_json(completed_path)
    require(capture_start.get("schema") == RUNNER_SCHEMA + ".capture-start" and capture_start.get("status") == "started", "capture start schema/status changed")
    require(capture_start.get("plan") == bindings["plan"], "capture start plan binding changed")
    require(capture_start.get("runner") == bindings["runner"], "capture start runner binding changed")
    require(capture_start.get("serial") is True and capture_start.get("children_expected") == 24 and capture_start.get("warmup_iterations") == WARMUP and capture_start.get("measured_samples") == SAMPLES, "capture start schedule changed")
    require(capture_terminal.get("schema") == RUNNER_SCHEMA + ".capture-terminal" and capture_terminal.get("status") == "completed", "capture terminal is not complete")
    require(capture_terminal.get("children_expected") == 24 and capture_terminal.get("children_completed") == 24 and capture_terminal.get("retry_or_replacement") is False and capture_terminal.get("terminal_before_parse") is True, "capture terminal child contract changed")
    require(capture_terminal.get("capture_start") == descriptor(capture_start_path), "capture terminal/start binding changed")
    require(completed.get("schema") == RUNNER_SCHEMA + ".completed-packet" and completed.get("status") == "completed", "completed packet schema/status changed")
    require(completed.get("plan") == bindings["plan"] and completed.get("runner") == bindings["runner"], "completed packet binding changed")
    rows = plan["schedule"]["children"]
    host_after_path = Path(plan["capture"]["host_snapshot_after"]["path"])
    require(host_after_path.is_file() and not host_after_path.is_symlink(), "host-after snapshot is missing")
    host_after = read_json(host_after_path)
    require(isinstance(host_after, dict) and isinstance(host_after.get("created_utc"), str) and isinstance(host_after.get("scope"), str), "host-after snapshot schema is incomplete")
    packet_children = completed.get("children")
    require(isinstance(packet_children, list) and len(packet_children) == 24 and len({item.get("child_id") for item in packet_children}) == 24, "completed packet child count changed")
    child_records: list[dict[str, Any]] = []
    for index, row in enumerate(rows):
        record = validate_child(packet, row, plan, bindings["plan"], index)
        child_records.append(record)
    for first, second in zip(child_records, child_records[1:]):
        require(second["start_seconds"] >= first["terminal_seconds"], "children overlapped or were not terminal-first")
    completed_by_id = {item.get("child_id"): item for item in packet_children}
    for record in child_records:
        item = completed_by_id.get(record["row"]["child_id"])
        require(item is not None and item.get("validation") == descriptor(record["child"] / "validation.json"), f"completed packet validation binding changed for {record['row']['child_id']}")

    for record in child_records:
        record["stats"] = child_stats(record["samples"])
    by_pair = {pair["pair_id"]: pair for pair in plan["schedule"]["pairs"]}
    records_by_id = {record["row"]["child_id"]: record for record in child_records}
    pair_results: dict[str, dict[str, Any]] = {}
    per_case: dict[str, dict[str, Any]] = {}
    for case_id in CASES:
        case_pairs: dict[str, dict[str, Any]] = {}
        for pair_type in PAIR_SCHEDULE[case_id]:
            pair = next(item for item in plan["schedule"]["pairs"] if item["case_id"] == case_id and item["pair_type"] == pair_type)
            first = records_by_id[pair["children"][0]]
            second = records_by_id[pair["children"][1]]
            result = pair_summary(first, second, pair)
            pair_results[pair["pair_id"]] = result
            case_pairs[pair_type] = result
        per_case[case_id] = {"pairs": case_pairs, "classification": classify(case_id, case_pairs)}

    order = sorted(child_records, key=lambda record: record["start_seconds"])
    return {
        "schema": REDUCER_SCHEMA,
        "status": "descriptive_read_only",
        "decision": {"adoption_allowed": False, "screen_reopened": False, "causal_claim": False, "0964_rejection_preserved": True},
        "scope": {"lane": "native", "children": 24, "pairs": 12, "cases": list(CASES), "warmup_iterations": WARMUP, "samples_per_child": SAMPLES, "phases": list(PHASES), "nearest_rank": {"p50_rank": 128, "p95_rank": 244, "p99_rank": 254}, "pooled_samples_used": False, "historical_samples_used": False, "allocator_reports_read": False},
        "bindings": {"plan": bindings["plan"], "runner": bindings["runner"], "reducer": bindings["reducer"], "capture_start": descriptor(capture_start_path), "capture_terminal": descriptor(capture_terminal_path), "completed_packet": descriptor(completed_path), "host_before": bind(plan["capture"]["host_snapshot_before"], "host-before"), "host_after": descriptor(host_after_path)},
        "child_summaries": [{"child_id": record["row"]["child_id"], "case_id": record["row"]["case_id"], "pair_id": record["row"]["pair_id"], "pair_type": record["row"]["pair_type"], "position": record["row"]["position"], "role": record["row"]["role"], "binary_role": record["row"]["binary_role"], "start_utc": record["start"].get("started_utc"), "stats": record["stats"], "report": record["report_path"], "time_metrics": record["time_metrics"]} for record in order],
        "pairs": pair_results,
        "case_results": per_case,
        "interpretation_limits": ["RR/RC/CR/CC pairs are descriptive controls within this fresh packet; no confidence interval or pooled-sample effect is computed.", "A concordant, nonreproduced, or mixed label does not attribute the result to the source, build, host, or allocator and does not qualify the candidate.", "The 0964 1.05 safety rule remains binding and is not reopened by this diagnostic.", "Any failed custody, terminal, affinity, oracle, or independent audit check rejects the packet; no adaptive samples or replacements are allowed."],
    }


def self_test() -> dict[str, Any]:
    values = list(range(1, SAMPLES + 1))
    require(nearest_rank(values, 0.50) == 128 and nearest_rank(values, 0.95) == 244 and nearest_rank(values, 0.99) == 254, "nearest-rank self-test failed")
    require(round(geometric_mean([1.0, 4.0]), 12) == 2.0, "geometric mean self-test failed")
    return {"schema": REDUCER_SCHEMA + ".self-test", "status": "passed", "samples": SAMPLES, "nearest_rank": {"p50": 128, "p95": 244, "p99": 254}, "geometric_mean_1_4": 2.0}


def write_once(path: Path, value: Any) -> None:
    path = path.expanduser().resolve()
    require(not path.exists() and not path.is_symlink(), f"output already exists: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write((json.dumps(value, indent=2, sort_keys=True) + "\n").encode())
        stream.flush()
        os.fsync(stream.fileno())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--packet", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    try:
        if args.self_test:
            print(json.dumps(self_test(), sort_keys=True))
            return 0
        require(args.plan is not None and args.packet is not None and args.output is not None, "--plan, --packet, and --output are required")
        write_once(args.output, reduce_packet(args.plan, args.packet))
    except ReductionError as exc:
        parser.error(str(exc))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
