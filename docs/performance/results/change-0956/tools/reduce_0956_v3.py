#!/usr/bin/env python3
"""Offline paired reduction for the 0956 native width experiment.

The reducer consumes a sealed packet only; it never starts a benchmark.  It
retains child identity and raw-vector provenance, computes child nearest-rank
p50/p95/p99/mean values, pairs adjacent A/B legs inside each frozen block,
and bootstraps the six block ratios.  A process/RSS metric is explicitly
whole-process (including harness/oracle work from ``/usr/bin/time``).
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import math
import random
import statistics
from pathlib import Path
from typing import Any, Iterable

HERE = Path(__file__).resolve().parent
_spec = importlib.util.spec_from_file_location("capture_0956_v3", HERE / "capture_0956_v3.py")
_capture = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(_capture)

SCHEMA = "litchi.performance.0956.native-width-reduction.v3"


def utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise ValueError(f"expected regular file: {path}")
    return {"path": str(path.resolve()), "bytes": path.stat().st_size, "sha256": sha256(path)}


def write_exclusive(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()


def read_json(path: Path) -> dict[str, Any]:
    with path.open("rb") as stream:
        value = json.loads(stream.read().decode("utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"JSON object required: {path}")
    return value


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def nearest_rank(values: list[float], quantile: float) -> float:
    require(values, "cannot quantile an empty vector")
    ordered = sorted(values)
    index = max(0, min(len(ordered) - 1, math.ceil(quantile * len(ordered)) - 1))
    return ordered[index]


def stats(values: list[float]) -> dict[str, float]:
    require(values, "empty timing vector")
    return {
        "count": float(len(values)),
        "mean": float(statistics.fmean(values)),
        "p50": nearest_rank(values, 0.50),
        "p95": nearest_rank(values, 0.95),
        "p99": nearest_rank(values, 0.99),
        "min": float(min(values)),
        "max": float(max(values)),
    }


def ratio(numerator: float, denominator: float, label: str) -> dict[str, Any]:
    if denominator == 0:
        return {"status": "undefined_zero_denominator", "label": label, "numerator": numerator, "denominator": denominator, "ratio": None}
    return {"status": "defined", "label": label, "numerator": numerator, "denominator": denominator, "ratio": numerator / denominator}


def geometric_mean(values: list[float]) -> float:
    require(values, "empty geometric-mean vector")
    require(all(value >= 0 and math.isfinite(value) for value in values), "invalid geometric-mean value")
    if any(value == 0 for value in values):
        return 0.0
    return math.exp(statistics.fmean(math.log(value) for value in values))


def bootstrap_geometric_mean(values: list[float], *, seed: int, resamples: int) -> dict[str, Any]:
    require(len(values) == 6, "the frozen bootstrap unit is six block ratios")
    require(all(value >= 0 and math.isfinite(value) for value in values), "bootstrap requires nonnegative finite ratios")
    rng = random.Random(seed)
    estimates: list[float] = []
    for _ in range(resamples):
        estimates.append(geometric_mean([rng.choice(values) for _ in values]))
    point = geometric_mean(values)
    return {
        "seed": seed,
        "resamples": resamples,
        "unit": "six block ratios sampled with replacement",
        "statistic": "geometric_mean",
        "point_geometric_mean": point,
        "ci95": [nearest_rank(estimates, 0.025), nearest_rank(estimates, 0.975)],
    }


def parse_child_time(path: Path) -> dict[str, float | int]:
    return _capture.parse_time(path)


def _same_descriptor(actual_path: Path, claimed: Any, label: str) -> dict[str, Any]:
    require(isinstance(claimed, dict), f"{label} descriptor missing")
    actual = descriptor(actual_path)
    require(claimed == actual, f"{label} descriptor mismatch")
    return actual


def child_report(packet: Path, row: dict[str, Any], plan: dict[str, Any], index: int) -> tuple[dict[str, Any], dict[str, Any]]:
    child = packet / "children" / row["id"]
    require(child.is_dir() and not child.is_symlink(), f"missing child directory {child}")
    start_path, terminal_path, validation_path = child / "start.json", child / "terminal.json", child / "validation.json"
    post_path = child / "post-terminal-artifacts.json"
    start, terminal, validation = read_json(start_path), read_json(terminal_path), read_json(validation_path)
    require(start.get("index") == index and terminal.get("index") == index, f"child index mismatch: {row['id']}")
    require(start.get("row") == row and terminal.get("row") == row, f"child schedule row mismatch: {row['id']}")
    expected = _capture._expected_child_contract(row, plan, child)
    for record_name, record in (("start", start), ("terminal", terminal)):
        require(record.get("argv") == expected["argv"], f"{record_name} argv mismatch: {row['id']}")
        require(record.get("cwd") == expected["cwd"], f"{record_name} cwd mismatch: {row['id']}")
        require(record.get("environment") == expected["environment"], f"{record_name} environment mismatch: {row['id']}")
        require(record.get("deadline_seconds") == expected["deadline_seconds"], f"{record_name} deadline mismatch: {row['id']}")
        require(record.get("cpu_mask_requested") == expected["cpu_mask"], f"{record_name} requested affinity mismatch: {row['id']}")
    require(start.get("terminal_path") == str(terminal_path), f"start terminal path mismatch: {row['id']}")
    require(terminal.get("status") == "completed" and terminal.get("returncode") == 0, f"child terminal did not complete: {row['id']}")
    require(terminal.get("deadline_exceeded") is False and terminal.get("forced_cleanup") is False and terminal.get("process_group_clear") is True, f"child process custody failed: {row['id']}")
    require(terminal.get("affinity_status") == "verified" and terminal.get("cpu_mask_observed") == expected["cpu_mask"], f"affinity was not observed on frozen mask: {row['id']}")
    report_path = child / "workload-output" / plan["cli"]["output_file"]
    require(str(report_path.resolve()) == expected["report"], f"report output path mismatch: {row['id']}")
    required_files = {"stdout": child / "stdout.log", "stderr": child / "stderr.log", "time": child / "time-v.txt", "report": report_path}
    require(post_path.is_file() and not post_path.is_symlink(), f"post-terminal artifact list missing: {row['id']}")
    post = read_json(post_path)
    expected_artifacts = {name: descriptor(path) for name, path in required_files.items()}
    require(post == expected_artifacts, f"post-terminal artifact descriptors mismatch: {row['id']}")
    for name, path in required_files.items():
        require(str(path.resolve()) == expected[name], f"{name} output path mismatch: {row['id']}")
    report_desc = _same_descriptor(report_path, post.get("report"), f"report/{row['id']}")
    _same_descriptor(required_files["time"], post.get("time"), f"time/{row['id']}")
    _same_descriptor(required_files["stdout"], post.get("stdout"), f"stdout/{row['id']}")
    _same_descriptor(required_files["stderr"], post.get("stderr"), f"stderr/{row['id']}")
    _same_descriptor(terminal_path, validation.get("terminal"), f"validation terminal/{row['id']}")
    _same_descriptor(report_path, validation.get("report"), f"validation report/{row['id']}")
    require(validation.get("status") == "passed" and validation.get("row") == row, f"child validation did not pass or row changed: {row['id']}")
    report = read_json(report_path)
    _capture.validate_report(report, row, plan)
    time_values = parse_child_time(required_files["time"])
    custody = {
        "row": row, "index": index, "child": str(child.resolve()),
        "start": descriptor(start_path), "terminal": descriptor(terminal_path), "validation": descriptor(validation_path),
        "post_terminal_artifacts": descriptor(post_path), "stdout": descriptor(required_files["stdout"]),
        "stderr": descriptor(required_files["stderr"]), "time": descriptor(required_files["time"]), "report": report_desc,
        "time_values": time_values,
    }
    return report, custody

def child_metric_values(report: dict[str, Any], time_values: dict[str, float | int], metric: str) -> list[float]:
    if metric in _capture.TIME_FIELDS.values():
        value = time_values[metric]
        return [float(value)]
    samples = report["case"]["samples"]
    return [float(sample[metric]) for sample in samples]


def summarize_child(report: dict[str, Any], time_values: dict[str, float | int]) -> dict[str, Any]:
    metric_names = ["ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns"]
    result: dict[str, Any] = {}
    for metric in metric_names:
        result[metric] = stats(child_metric_values(report, time_values, metric))
    for metric in ("max_rss_kib", "user_seconds", "system_seconds", "major_page_faults", "minor_page_faults", "voluntary_context_switches", "involuntary_context_switches"):
        result[metric] = stats(child_metric_values(report, time_values, metric))
    return result


def read_packet(plan_path: Path, packet: Path) -> tuple[dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    plan = read_json(plan_path)
    rows = _capture.validate_plan(plan, for_execution=False)
    require(packet.is_dir() and not packet.is_symlink(), f"packet is not a directory: {packet}")
    start_path = packet / "capture-start.json"
    outer_path = packet / "capture-terminal.json"
    start = read_json(start_path)
    outer = read_json(outer_path)
    require(start.get("schema") == _capture.RUNNER_SCHEMA + ".outer-start", "capture-start schema mismatch")
    require(start.get("row_count") == len(rows), "capture-start row count mismatch")
    require(start.get("plan") == descriptor(plan_path), "capture-start plan descriptor mismatch")
    require(start.get("runner") == plan["helpers"]["capture"], "capture-start runner binding mismatch")
    require(start.get("reducer") == plan["helpers"]["reducer"], "capture-start reducer binding mismatch")
    require(start.get("controls") == plan["helpers"]["controls"], "capture-start controls binding mismatch")
    require(start.get("source_manifest") == descriptor(Path(plan["source"]["source_manifest"])), "capture-start source manifest binding mismatch")
    require(outer.get("schema") == _capture.RUNNER_SCHEMA + ".outer-terminal", "capture-terminal schema mismatch")
    require(outer.get("status") == "completed", "capture outer terminal is not completed")
    require(outer.get("children_started") == len(rows) and outer.get("children_expected") == len(rows), "capture child count is not 144")
    require(outer.get("capture_start") == descriptor(start_path), "capture-terminal start binding mismatch")
    require(outer.get("plan") == descriptor(plan_path), "capture-terminal plan descriptor mismatch")
    require(outer.get("runner") == plan["helpers"]["capture"], "capture-terminal runner binding mismatch")
    children: list[dict[str, Any]] = []
    for index, row in enumerate(rows):
        report, custody = child_report(packet, row, plan, index)
        children.append({"row": row, "report": report, "custody": custody, "stats": summarize_child(report, custody["time_values"])})
    return plan, children, outer

def block_pairs(items: list[dict[str, Any]], metric: str, quantile: str = "p50") -> list[dict[str, Any]]:
    require(items, "block pairing received no children")
    comparisons = {item["row"]["comparison_id"] for item in items}
    cases = {item["row"]["case_id"] for item in items}
    require(len(comparisons) == 1 and len(cases) == 1, "block pairing mixed groups")
    comparison = next(iter(comparisons))
    case = next(iter(cases))
    by_key = {(item["row"]["comparison_id"], item["row"]["case_id"], item["row"]["block"], item["row"]["position"]): item for item in items}
    output: list[dict[str, Any]] = []
    for block in range(6):
                values: list[dict[str, Any]] = []
                for left, right in ((0, 1), (2, 3)):
                    first, second = by_key[(comparison, case, block, left)], by_key[(comparison, case, block, right)]
                    if first["row"]["role"] == "A" and second["row"]["role"] == "B":
                        serial, treatment = first, second
                    elif first["row"]["role"] == "B" and second["row"]["role"] == "A":
                        treatment, serial = first, second
                    else:
                        raise ValueError(f"invalid adjacent block roles: {comparison}/{case}/b{block}")
                    serial_value = serial["stats"][metric][quantile]
                    treatment_value = treatment["stats"][metric][quantile]
                    values.append({"treatment_id": treatment["row"]["id"], "serial_id": serial["row"]["id"], **ratio(treatment_value, serial_value, metric)})
                defined = [item["ratio"] for item in values if item["status"] == "defined"]
                output.append({"comparison_id": comparison, "case_id": case, "block": block, "metric": metric, "quantile": quantile, "adjacent_process_ratios": values, "block_geometric_mean": geometric_mean(defined) if len(defined) == 2 else None, "status": "defined" if len(defined) == 2 else "undefined"})
    return output


def reduce(plan_path: Path, packet: Path, output: Path) -> None:
    plan, children, outer = read_packet(plan_path, packet)
    metrics = ["ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns", "max_rss_kib", "user_seconds", "system_seconds", "major_page_faults", "minor_page_faults", "voluntary_context_switches", "involuntary_context_switches"]
    groups: list[dict[str, Any]] = []
    for comparison in ("width1-vs-width2", "width1-vs-width4"):
        for case in ("medium-one-percent", "dense-sparse-one-percent", "medium-one-edit"):
            selected = [item for item in children if item["row"]["comparison_id"] == comparison and item["row"]["case_id"] == case]
            require(len(selected) == 24, f"unexpected child count for {comparison}/{case}")
            child_stats = [{"id": item["row"]["id"], "role": item["row"]["role"], "block": item["row"]["block"], "position": item["row"]["position"], "width": item["row"]["width"], "metrics": {metric: item["stats"][metric] for metric in metrics}} for item in selected]
            group: dict[str, Any] = {"comparison_id": comparison, "case_id": case, "children": child_stats, "metrics": {}}
            for metric in metrics:
                quantile_results: dict[str, Any] = {}
                for quantile in ("p50", "p95", "p99", "mean"):
                    ratios = block_pairs(selected, metric, quantile)
                    block_values = [item["block_geometric_mean"] for item in ratios if item["status"] == "defined"]
                    pair_values = [pair["ratio"] for item in ratios for pair in item["adjacent_process_ratios"] if pair["status"] == "defined"]
                    quantile_result: dict[str, Any] = {"block_ratios": ratios, "process_pair_ratio_count": len(pair_values), "process_pair_ratios": pair_values, "threshold": 1.05, "ratio_direction": "treatment_width_over_serial_width"}
                    if len(block_values) == 6:
                        quantile_result["bootstrap"] = bootstrap_geometric_mean(block_values, seed=plan["execution"]["bootstrap"]["seed"], resamples=plan["execution"]["bootstrap"]["resamples"])
                        quantile_result["geometric_mean_of_block_ratios"] = geometric_mean(block_values)
                        quantile_result["any_process_pair_over_1_05"] = any(value > 1.05 for value in pair_values)
                        quantile_result["any_block_over_1_05"] = any(value > 1.05 for value in block_values)
                    else:
                        quantile_result["status"] = "undefined_due_to_zero_denominator"
                    quantile_results[quantile] = quantile_result
                group["metrics"][metric] = {"child_stats": [{"id": item["row"]["id"], "role": item["row"]["role"], "block": item["row"]["block"], "position": item["row"]["position"], "width": item["row"]["width"], "stats": item["stats"][metric]} for item in selected], "quantiles": quantile_results, "threshold": 1.05, "ratio_direction": "treatment_width_over_serial_width"}
            phase_metrics = ["ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns"]
            primary_quantiles: dict[str, Any] = {}
            for quantile in ("p50", "p95", "p99"):
                publication_point = group["metrics"]["publication_ns"]["quantiles"][quantile].get("geometric_mean_of_block_ratios")
                total_point = group["metrics"]["total_ns"]["quantiles"][quantile].get("geometric_mean_of_block_ratios")
                rss_point = group["metrics"]["max_rss_kib"]["quantiles"][quantile].get("geometric_mean_of_block_ratios")
                primary_quantiles[quantile] = {
                    "publication_point_ratio": publication_point,
                    "total_point_ratio": total_point,
                    "rss_point_ratio": rss_point,
                    "aggregate_point_over_1_05": any(value is None or value > 1.05 for value in (publication_point, total_point, rss_point)),
                    "latency_primary_metrics": ["publication_ns", "total_ns"],
                    "rss_review_metric": "max_rss_kib",
                    "rss_is_review_only": True,
                }
            group["gate_summary"] = {
                "primary_aggregate": primary_quantiles,
                "p50_all_phase_process_pairs_le_1_05": all(not group["metrics"][metric]["quantiles"]["p50"].get("any_process_pair_over_1_05", True) for metric in phase_metrics),
                "p95_all_phase_process_pairs_le_1_05": all(not group["metrics"][metric]["quantiles"]["p95"].get("any_process_pair_over_1_05", True) for metric in phase_metrics),
                "p99_phase_review_trigger": any(group["metrics"][metric]["quantiles"]["p99"].get("any_process_pair_over_1_05", True) for metric in phase_metrics),
                "rss_review_trigger": group["metrics"]["max_rss_kib"]["quantiles"]["p50"].get("any_process_pair_over_1_05", True),
                "nonphase_process_metrics_descriptive_only": True,
                "status": "descriptive_review_only",
            }
            group["interpretation"] = "descriptive only; no OS-thread, allocation, or adoption claim"
            groups.append(group)
    result = {
        "schema": SCHEMA,
        "status": "completed",
        "created_utc": utc(),
        "plan": descriptor(plan_path),
        "packet": {"path": str(packet.resolve()), "outer_terminal": descriptor(packet / "capture-terminal.json"), "children": len(children)},
        "outer_terminal": outer,
        "metrics": {"phase_vectors": ["ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns"], "process_metrics": ["max_rss_kib", "user_seconds", "system_seconds", "major_page_faults", "minor_page_faults", "voluntary_context_switches", "involuntary_context_switches"], "quantiles": {"method": "nearest rank", "p50": 0.5, "p95": 0.95, "p99": 0.99}, "rss_scope": "whole process including setup/oracle work as measured by /usr/bin/time -v", "bootstrap": plan["execution"]["bootstrap"]},
        "child_count": len(children),
        "groups": groups,
        "adoption": {"claim": "none", "allocation_claim": "none", "native_width_claim": "none until separately reviewed", "decision": "not computed by this reducer"},
    }
    write_exclusive(output, result)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", type=Path, required=True)
    ap.add_argument("--packet", type=Path, required=True)
    ap.add_argument("--output", type=Path, required=True)
    args = ap.parse_args()
    try:
        reduce(args.plan, args.packet, args.output)
    except Exception as exc:
        print(f"reduce_0956_v3: {exc}", file=__import__("sys").stderr)
        return 1
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
