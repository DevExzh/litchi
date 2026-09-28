"""Independent arithmetic audit for the 0827 ordinary-save packet.

This module deliberately does not import :mod:`analysis`.  It rereads raw
reports and RSS receipts, recomputes the harness quantiles, Welford means,
paired native ratios, and observer allocation counters, and then compares its
small deterministic witness with ``analysis.json``.
"""

from __future__ import annotations

import hashlib
import json
import math
import random
import statistics
import sys
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN_SCHEMA = "litchi.performance.0827.plan.v1"
RAW_SCHEMA = "litchi.performance.0827.raw-audit.v1"
ALIGNMENT = "elapsed_ns.samples_by_elapsed_then_sample_index"
SEED = 827827
RESAMPLES = 10_000
LOW = 250
HIGH = 9749
FORMATS = ("docx", "xlsx", "pptx")
PHASES = ("lifecycle", "edit", "atomic_publish", "counting_publish")
LEGS = ("before", "after")
CASES = tuple(
    (fmt, phase, f"{fmt}_real_file_ordinary_save_{phase}")
    for fmt in FORMATS for phase in PHASES
)
ALLOC = ("allocation_calls", "deallocation_calls", "reallocation_calls",
         "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
         "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
         "peak_live_bytes_after", "region_peak_live_bytes")
DERIVED_ALLOC = ALLOC + ("net_live", "peak_above_entry")
OBSERVER_METRICS = DERIVED_ALLOC + ("rss_kib",)
PROCESS_VECTOR_NAMES = (
    "user_cpu_ticks", "system_cpu_ticks", "clock_ticks_per_second",
    "minor_faults", "major_faults", "voluntary_context_switches",
    "nonvoluntary_context_switches", "rss_delta_bytes", "peak_rss_bytes",
    "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
    "syscr", "syscw",
)


class AuditError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise AuditError(f"invalid JSON {path}: {error}") from error


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def artifact(value: Any, label: str) -> Path:
    require(isinstance(value, dict) and isinstance(value.get("path"), str),
            f"{label} descriptor malformed")
    raw = Path(value["path"])
    candidates = [raw] if raw.is_absolute() else [PACKET / raw, ROOT / raw]
    path = next((p.resolve() for p in candidates if p.is_file() and not p.is_symlink()), None)
    require(path is not None, f"{label} missing")
    require(path.stat().st_size == value.get("bytes")
            and sha(path) == value.get("sha256"), f"{label} identity changed")
    return path


def integer(value: Any, label: str, positive: bool = False) -> None:
    require(isinstance(value, int) and not isinstance(value, bool)
            and (value > 0 if positive else value >= 0), f"{label} invalid")


def finite(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} invalid")


def midpoint(values: list[int]) -> int:
    ordered = sorted(values)
    return (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) // 2


def nearest(values: list[int], quantile: float) -> int:
    ordered = sorted(values)
    return ordered[max(1, math.ceil(len(ordered) * quantile)) - 1]


def welford(values: list[int]) -> float:
    mean = 0.0
    count = 0
    for value in values:
        count += 1
        mean += (value - mean) / count
    return mean


def sample_stats(values: list[int]) -> dict[str, Any]:
    require(values and all(isinstance(x, int) and x > 0 for x in values),
            "elapsed vector invalid")
    ordered = sorted(values)
    return {"count": len(values), "min": ordered[0], "p50": midpoint(ordered),
            "p95": nearest(ordered, .95), "p99": nearest(ordered, .99),
            "max": ordered[-1], "mean": welford(ordered)}


def parse_alloc(report: dict[str, Any], samples: int, mode: str,
                label: str) -> dict[str, list[float]] | None:
    """Parse raw per-sample allocator vectors without comparing timings."""
    require(mode in {"native", "observer"}, f"{label} mode changed")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label} result cardinality changed")
    metrics = results[0].get("operation_metrics")
    require(isinstance(metrics, dict), f"{label} operation metrics missing")
    allocation_value = metrics.get("allocation")
    require(isinstance(allocation_value, dict), f"{label} allocation missing")
    if mode == "native":
        require(set(allocation_value) == {"status", "scope", *ALLOC},
                f"{label} native allocation envelope changed")
        require(allocation_value.get("status") == "unavailable"
                and allocation_value.get("scope") == "operation_global_system_allocator",
                f"{label} native allocation unexpectedly measured")
        for name in ALLOC:
            require(allocation_value[name] == {
                "status": "unavailable", "scope": "operation_global_system_allocator"
            }, f"{label} native allocation {name} changed")
        return None
    require(set(allocation_value) == {"status", "scope", *ALLOC},
            f"{label} allocation envelope changed")
    require(allocation_value.get("status") == "measured"
            and allocation_value.get("scope") == "operation_global_system_allocator",
            f"{label} allocation status changed")
    result: dict[str, list[float]] = {}
    for name in ALLOC:
        item = allocation_value.get(name)
        require(isinstance(item, dict) and set(item) == {"status", "scope", "values"}
                and item.get("status") == "measured"
                and item.get("scope") == "operation_global_system_allocator"
                and isinstance(item.get("values"), list)
                and len(item["values"]) == samples,
                f"{label} allocation {name} changed")
        values = item["values"]
        require(all(type(x) is int and 0 <= x < (1 << 64) for x in values),
                f"{label} allocation {name} is not an unsigned vector")
        result[name] = list(values)
    result["net_live"] = [a - b for a, b in zip(result["live_bytes_after"],
                                                  result["live_bytes_before"])]
    result["peak_above_entry"] = [a - b for a, b in zip(result["region_peak_live_bytes"],
                                                         result["live_bytes_before"])]
    for index in range(samples):
        require(result["live_bytes_after"][index] - result["live_bytes_before"][index]
                == result["allocated_bytes"][index] - result["deallocated_bytes"][index],
                f"{label} live-byte conservation changed")
        require(result["peak_live_bytes_before"][index] >= result["live_bytes_before"][index]
                and result["peak_live_bytes_after"][index]
                >= result["peak_live_bytes_before"][index]
                and result["region_peak_live_bytes"][index]
                >= max(result["live_bytes_before"][index], result["live_bytes_after"][index])
                and result["region_peak_live_bytes"][index]
                <= result["peak_live_bytes_after"][index],
                f"{label} allocation peak bounds changed")
    require(all(math.isfinite(x) for x in result["net_live"]),
            f"{label} signed net-live vector changed")
    require(all(x >= 0 for x in result["peak_above_entry"]),
            f"{label} peak-above-entry vector changed")
    return result


def parse_report(report: dict[str, Any], *, expected_case: str,
                 expected_samples: int, expected_warmup: int, mode: str,
                 label: str) -> dict[str, Any]:
    """Parse one historical/current report for schema and arithmetic only."""
    require(mode in {"native", "observer"}, f"{label} mode changed")
    require(report.get("schema_version") == 1, f"{label} report schema changed")
    tool = report.get("tool")
    expected_binary = ("litchi-perf-baseline" if mode == "native"
                       else "litchi-perf-baseline-alloc")
    require(isinstance(tool, dict) and tool.get("name") == "litchi-perf-baseline"
            and tool.get("binary") == expected_binary
            and tool.get("profile") == "release", f"{label} tool changed")
    if mode == "native":
        require(tool.get("instrumentation") == "none"
                and "allocator_counter_revision" not in tool,
                f"{label} native instrumentation changed")
    else:
        require(tool.get("instrumentation") in {
            "system_allocator_operation_scoped",
            "ordinary_save_procfs_and_system_allocator_operation_scoped",
        } and tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{label} observer instrumentation changed")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("samples_per_case") == expected_samples
            and configuration.get("warmup_iterations_per_case") == expected_warmup
            and configuration.get("cases") == [expected_case]
            and configuration.get("filesystem_fresh_child_per_sample") is True
            and configuration.get("filesystem_process_isolated") is True,
            f"{label} configuration changed")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1
            and isinstance(results[0], dict)
            and results[0].get("case") == expected_case,
            f"{label} result changed")
    elapsed = results[0].get("elapsed_ns")
    require(isinstance(elapsed, dict) and elapsed.get("unit") == "ns",
            f"{label} elapsed envelope changed")
    values = elapsed.get("samples")
    order = elapsed.get("sample_order")
    require(isinstance(values, list) and len(values) == expected_samples
            and all(type(x) is int and x >= 0 for x in values)
            and values == sorted(values)
            and isinstance(order, list) and len(order) == expected_samples
            and all(type(x) is int and 0 <= x < expected_samples for x in order)
            and sorted(order) == list(range(expected_samples)),
            f"{label} elapsed vectors changed")
    require(all(values[index] != values[index + 1] or order[index] < order[index + 1]
               for index in range(expected_samples - 1)),
            f"{label} elapsed tie order changed")
    arithmetic = sample_stats(values)
    require(elapsed.get("min") == arithmetic["min"]
            and elapsed.get("p50") == arithmetic["p50"]
            and elapsed.get("p95") == arithmetic["p95"]
            and elapsed.get("p99") == arithmetic["p99"]
            and elapsed.get("max") == arithmetic["max"],
            f"{label} elapsed quantiles changed")
    finite(elapsed.get("mean"), f"{label} elapsed mean")
    require(abs(float(elapsed["mean"]) - arithmetic["mean"]) < 1e-12,
            f"{label} elapsed mean changed")
    metrics = results[0].get("operation_metrics")
    require(isinstance(metrics, dict) and metrics.get("sample_count") == expected_samples
            and metrics.get("sample_indices") == order
            and metrics.get("alignment") == ALIGNMENT,
            f"{label} operation-metric alignment changed")
    if mode == "native":
        require(metrics.get("latency_claim") == "comparable_timed_operation"
                and isinstance(metrics.get("process"), dict)
                and metrics["process"].get("status") == "unavailable",
                f"{label} native metrics changed")
    else:
        require(metrics.get("latency_claim") ==
                "allocator_instrumented_elapsed_not_latency_claim",
                f"{label} observer latency claim changed")
        process = metrics.get("process")
        require(isinstance(process, dict)
                and process.get("status") in {"measured", "unavailable"},
                f"{label} process metrics changed")
        if process.get("status") == "measured":
            for name in PROCESS_VECTOR_NAMES:
                item = process.get(name)
                if item is None:
                    continue
                require(isinstance(item, dict) and isinstance(item.get("values"), list)
                        and len(item["values"]) == expected_samples,
                        f"{label} process {name} vector changed")
                for value in item["values"]:
                    finite(value, f"{label} process {name}")
    allocation = parse_alloc(report, expected_samples, mode, label)
    return {"case": expected_case, "samples": values, "sample_order": order,
            "stats": arithmetic, "allocation": allocation}


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(len(values) == 6, "native bootstrap requires six ratios")
    rng = random.Random(SEED)
    estimates = sorted(statistics.median(rng.choice(values) for _ in values)
                       for _ in range(RESAMPLES))
    return {"estimate": statistics.median(values), "ci_low": estimates[LOW],
            "ci_high": estimates[HIGH], "seed": SEED, "resamples": RESAMPLES,
            "low_rank": LOW, "high_rank": HIGH}


def orders(plan: dict[str, Any], lane: str) -> list[list[str]]:
    values = plan["lanes"][lane].get("orders", [])
    result = []
    for value in values:
        if isinstance(value, list):
            arm = [str(x).lower() for x in value]
        elif str(value).lower() in {"ba", "forward"}:
            arm = ["before", "after"]
        elif str(value).lower() in {"ab", "reverse"}:
            arm = ["after", "before"]
        else:
            raise AuditError(f"unknown order {value!r}")
        require(arm in (list(LEGS), list(reversed(LEGS))), "arm order changed")
        result.append(arm)
    return result


def plan() -> dict[str, Any]:
    value = read(PACKET / "plan.json")
    require(value.get("schema") == PLAN_SCHEMA, "plan schema changed")
    require(value.get("bootstrap", {}).get("seed") == SEED
            and value["bootstrap"].get("resamples") == RESAMPLES
            and value["bootstrap"].get("low_rank") == LOW
            and value["bootstrap"].get("high_rank") == HIGH,
            "bootstrap contract changed")
    require(len(value.get("cases", [])) == 12, "case count changed")
    require(value["lanes"]["native"]["blocks"] == 6
            and value["lanes"]["native"]["samples"] == 30
            and value["lanes"]["observer"]["blocks"] == 2
            and value["lanes"]["observer"]["samples"] == 3,
            "lane samples changed")
    return value


def receipt_rows(directory: Path, label: str) -> list[dict[str, Any]]:
    complete = read(directory / "complete.json")
    require(complete.get("status") == "pass", f"{label} completion failed")
    raw = complete.get("receipts")
    path = artifact(raw, f"{label} receipts") if isinstance(raw, dict) else directory / "receipts.json"
    rows = read(path)
    require(isinstance(rows, list), f"{label} receipts malformed")
    return rows


def allocation(report: dict[str, Any], samples: int, label: str) -> dict[str, list[float]]:
    result = parse_alloc(report, samples, "observer", label)
    require(result is not None, f"{label} allocation unavailable")
    return result


def read_lane(plan_value: dict[str, Any], lane: str, directory_name: str,
              qualification_leg: str | None = None) -> list[dict[str, Any]]:
    rows = receipt_rows(PACKET / directory_name, directory_name)
    section = plan_value["lanes"][lane]
    if lane == "qualification":
        wanted_orders = [[qualification_leg]]
    else:
        wanted_orders = orders(plan_value, lane)
    wanted = []
    for block, arm_order in enumerate(wanted_orders):
        for case in CASES:
            for leg in arm_order:
                wanted.append((block, case, leg))
    require(len(rows) == len(wanted), f"{directory_name} report count changed")
    result = []
    for row, (block, case, leg) in zip(rows, wanted):
        require(row.get("exit_code") == 0 and row.get("block") == block
                and row.get("leg") == leg and row.get("samples") == section["samples"]
                and row.get("warmup") == section["warmup"],
                f"{directory_name} receipt identity changed")
        report_path = artifact(row.get("report"), f"{directory_name}/{case[2]} report")
        report = read(report_path)
        result_value = report["results"][0]
        elapsed = result_value["elapsed_ns"]
        samples = elapsed.get("samples")
        require(report.get("schema_version") == 1 and result_value.get("case") == case[2]
                and isinstance(samples, list) and len(samples) == section["samples"]
                and samples == sorted(samples), f"{directory_name}/{case[2]} report changed")
        values = sample_stats(samples)
        require(elapsed.get("p50") == values["p50"]
                and elapsed.get("p95") == values["p95"]
                and elapsed.get("p99") == values["p99"]
                and elapsed.get("min") == values["min"]
                and elapsed.get("max") == values["max"]
                and abs(float(elapsed.get("mean")) - values["mean"]) < 1e-12,
                f"{directory_name}/{case[2]} raw arithmetic changed")
        rss_path = artifact(row.get("rss"), f"{directory_name}/{case[2]} RSS")
        rss = rss_path.read_text(encoding="utf-8").strip()
        require(rss.isdigit() and int(rss) > 0, f"{directory_name}/{case[2]} RSS changed")
        metrics = None
        if lane == "observer":
            metrics = allocation(report, section["samples"], f"{directory_name}/{case[2]}")
        result.append({"case": case, "leg": leg, "block": block, "stats": values,
                       "samples": samples, "rss_kib": int(rss), "allocation": metrics,
                       "report": str(report_path.relative_to(PACKET)),
                       "report_sha256": sha(report_path)})
    return result


def raw_rows(native: list[dict[str, Any]], observer: list[dict[str, Any]]) -> dict[str, Any]:
    native_result = []
    lookup = {(x["case"][2], x["block"], x["leg"]): x for x in native}
    for fmt, phase, name in CASES:
        rows = {leg: [lookup[(name, block, leg)] for block in range(6)] for leg in LEGS}
        summary = {leg: {metric: statistics.median(x["stats"][metric] for x in rows[leg])
                         for metric in ("p50", "p95", "p99", "mean")}
                   | {"rss_kib": statistics.median(x["rss_kib"] for x in rows[leg])}
                   for leg in LEGS}
        paired = {}
        for metric in ("p50", "p95", "p99", "mean", "rss_kib"):
            ratios = [(rows["after"][i]["rss_kib"] / rows["before"][i]["rss_kib"]
                       if metric == "rss_kib" else
                       rows["after"][i]["stats"][metric] / rows["before"][i]["stats"][metric])
                      for i in range(6)]
            paired[metric] = {"ratios": ratios, "median_ratio": statistics.median(ratios),
                              "bootstrap": bootstrap(ratios)}
        native_result.append({"case": [fmt, phase, name], "native": summary,
                              "paired": paired,
                              "reports": [{"leg": leg, "block": x["block"],
                                           "report": x["report"],
                                           "report_sha256": x["report_sha256"]}
                                          for leg in LEGS for x in rows[leg]]})
    observer_result = []
    lookup = {(x["case"][2], x["block"], x["leg"]): x for x in observer}
    for fmt, phase, name in CASES:
        by_leg = {leg: [lookup[(name, block, leg)] for block in range(2)] for leg in LEGS}
        metrics = {}
        for metric in OBSERVER_METRICS:
            if metric == "rss_kib":
                before = [x["rss_kib"] for x in by_leg["before"]]
                after = [x["rss_kib"] for x in by_leg["after"]]
                before_values = [[x["rss_kib"]] for x in by_leg["before"]]
                after_values = [[x["rss_kib"]] for x in by_leg["after"]]
            else:
                before = [statistics.median(x["allocation"][metric]) for x in by_leg["before"]]
                after = [statistics.median(x["allocation"][metric]) for x in by_leg["after"]]
                before_values = [x["allocation"][metric] for x in by_leg["before"]]
                after_values = [x["allocation"][metric] for x in by_leg["after"]]
            metrics[metric] = {"before_process_medians": before,
                               "after_process_medians": after,
                               "before_values": before_values,
                               "after_values": after_values,
                               "before_median": statistics.median(before),
                               "after_median": statistics.median(after),
                               "increase": statistics.median(after) > statistics.median(before)}
        observer_result.append({"case": [fmt, phase, name], "metrics": metrics})
    return {"native": native_result, "observer": observer_result}


def compare(value: dict[str, Any]) -> None:
    analysis_path = PACKET / "analysis.json"
    require(analysis_path.is_file(), "analysis.json missing")
    analyzed = read(analysis_path)
    require(analyzed.get("schema") == "litchi.performance.0827.ordinary-save-analysis.v1",
            "analysis schema changed")
    by_case = {row["case"]: row for row in analyzed["native_rows"]}
    for row in value["rows"]["native"]:
        name = row["case"][2]
        actual = by_case[name]
        for leg in LEGS:
            for metric in ("p50", "p95", "p99", "mean"):
                require(actual[leg]["process_median"][metric] == row["native"][leg][metric],
                        f"native {name}/{leg}/{metric} differs")
            require(actual[leg]["rss_process_median"] == row["native"][leg]["rss_kib"],
                    f"native {name}/{leg}/rss differs")
        for metric, paired in row["paired"].items():
            actual_pair = actual["paired"]["metrics"][metric]
            require(actual_pair["median_ratio"] == paired["median_ratio"],
                    f"native pair {name}/{metric} differs")
            if paired["bootstrap"] is not None:
                require(actual_pair["bootstrap"]["ci_low"] == paired["bootstrap"]["ci_low"]
                        and actual_pair["bootstrap"]["ci_high"] == paired["bootstrap"]["ci_high"],
                        f"native CI {name}/{metric} differs")
    analyzed_observer = {row["case"]: row for row in analyzed["observer_rows"]}
    for row in value["rows"]["observer"]:
        actual = analyzed_observer[row["case"][2]]
        for metric in OBSERVER_METRICS:
            expected = row["metrics"][metric]
            observed = (actual["before"]["rss_process_median"] if metric == "rss_kib"
                        else actual["before"]["allocation"]["median_over_blocks"][metric])
            require(observed == expected["before_median"],
                    f"observer before {row['case'][2]}/{metric} differs")
            observed = (actual["after"]["rss_process_median"] if metric == "rss_kib"
                        else actual["after"]["allocation"]["median_over_blocks"][metric])
            require(observed == expected["after_median"],
                    f"observer after {row['case'][2]}/{metric} differs")


def derive() -> dict[str, Any]:
    p = plan()
    qualification_before = read_lane(p, "qualification", "qualification-before", "before")
    qualification_after = read_lane(p, "qualification", "qualification-after", "after")
    native = read_lane(p, "native", "native")
    observer = read_lane(p, "observer", "observer")
    require(len(qualification_before) == len(qualification_after) == 12
            and len(native) == 144 and len(observer) == 48,
            "raw lane cardinality changed")
    value = {"schema": RAW_SCHEMA,
             "counts": {"qualification_reports": 24, "qualification_samples": 24,
                        "native_reports": 144, "native_samples": 4320,
                        "observer_reports": 48, "observer_samples": 144,
                        "reports": 216, "samples": 4488},
             "rows": raw_rows(native, observer),
             "qualification": [{"leg": x["leg"], "case": list(x["case"]),
                                "block": x["block"], "report": x["report"]}
                               for x in qualification_before + qualification_after]}
    compare(value)
    return value


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    require(args in (["--write"], ["--check"]), "use exactly --write or --check")
    try:
        value = derive()
        encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
        path = PACKET / "raw-audit.json"
        if args == ["--write"]:
            require(not path.exists(), "refusing to overwrite raw-audit.json")
            path.write_text(encoded, encoding="utf-8")
        else:
            require(path.is_file() and path.read_text(encoding="utf-8") == encoded,
                    "raw-audit.json does not replay byte-for-byte")
        print("0827 independent raw audit PASS: 216 reports, 4488 samples")
        return 0
    except (AuditError, OSError, ValueError, KeyError, TypeError, AssertionError) as error:
        print(f"0827 raw audit failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
