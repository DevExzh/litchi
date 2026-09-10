#!/usr/bin/env python3
"""Compare matched 0501 provider lifecycle reports and retain all flags."""

from __future__ import annotations

import argparse
import json
import re
import statistics
from pathlib import Path
from typing import Any

from custody import HERE, canonical_json, safe_relative, sha_file


REVIEW_PERCENT = 5.0
TIMINGS = ("api_sum_ns", "open_ns", "plan_ns", "publication_ns", "open_source_ns", "open_destination_ns")
PHASES = ("baseline", "opened", "planned", "published", "drop_result", "drop_plan", "drop_view", "drop_caller_sources", "drop_sink")
CACHE_COUNTERS = ("hits", "cold_loads", "waiter_joins", "successful_loads", "failed_loads", "evictions", "bypasses", "oversized_bypasses", "allocation_bypasses", "budget_reservation_failures")
READ_COUNTERS = ("logical_calls", "requested_bytes", "returned_bytes", "short_reads", "delayed_calls")
BUDGET_COUNTERS = ("memory_used", "input_bytes_used", "output_bytes_used", "work_used", "objects_used", "depth_used")


def percentile(values: list[int], fraction: float) -> float:
    ordered = sorted(values)
    if not ordered:
        raise ValueError("empty distribution")
    position = (len(ordered) - 1) * fraction
    left = int(position)
    right = min(left + 1, len(ordered) - 1)
    return ordered[left] + (ordered[right] - ordered[left]) * (position - left)


def distribution(values: list[int]) -> dict[str, Any]:
    return {
        "samples": len(values),
        "min": min(values),
        "p50": percentile(values, 0.50),
        "p95": percentile(values, 0.95),
        "p99": percentile(values, 0.99),
        "max": max(values),
        "mean": statistics.fmean(values),
    }


def relative(before: float, after: float) -> float | None:
    return None if before == 0 else (after - before) * 100.0 / before


def resource_rss(path: Path) -> int:
    matches = re.findall(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", path.read_text(errors="replace"), re.MULTILINE)
    if len(matches) != 1:
        raise ValueError(f"expected one GNU-time RSS line in {path}")
    return int(matches[0]) * 1024


def report_fingerprint(report: dict[str, Any]) -> dict[str, Any]:
    return {
        "source_archive_sha256": report["source_archive_sha256"],
        "destination_archive_sha256": report["destination_archive_sha256"],
        "expected_output_sha256": report["expected_output_sha256"],
        "expected_output_bytes": report["expected_output_bytes"],
        "source_revision": report["source_revision"],
    }


def counter_fingerprint(report: dict[str, Any]) -> list[Any]:
    result = []
    for row in report["samples_raw"]:
        sample = []
        for phase in row["phases"]:
            owners = []
            for owner in ("source", "destination"):
                read = phase[f"{owner}_reads"]
                if read["availability"] == "unavailable":
                    read_value = None
                else:
                    read_value = tuple(read[key] for key in READ_COUNTERS) + tuple(read["request_size_counts"])
                cache = phase[f"{owner}_cache"]
                if cache["availability"] == "unavailable":
                    cache_value = None
                else:
                    cache_value = tuple(cache[key] for key in CACHE_COUNTERS)
                budget = phase[f"{owner}_budget"]
                budget_value = tuple(budget[key] for key in BUDGET_COUNTERS)
                owners.append((read_value, cache_value, budget_value))
            sample.append((phase["label"], tuple(owners)))
        result.append(tuple(sample))
    return result


def load_phase(phase: str, protocol: dict[str, Any]) -> dict[str, dict[str, Any]]:
    folder = HERE / phase
    if not folder.is_dir():
        raise ValueError(f"missing phase directory: {phase}")
    verifier_spec = __import__("importlib.util").util.spec_from_file_location("verify_report_compare", HERE / "verify-report.py")
    if verifier_spec is None or verifier_spec.loader is None:
        raise ValueError("cannot import report verifier")
    verifier = __import__("importlib.util").util.module_from_spec(verifier_spec)
    verifier_spec.loader.exec_module(verifier)
    records: dict[str, dict[str, Any]] = {}
    for receipt_path in sorted(folder.glob("*.receipt.json")):
        receipt = json.loads(receipt_path.read_text())
        if receipt.get("status") != "pass" or receipt.get("cleanup_verified") is not True:
            raise ValueError(f"{receipt_path}: capture did not pass and clean up")
        name = receipt["name"]
        report_path = safe_relative(receipt["artifacts"]["report"]["path"])
        if sha_file(report_path) != receipt["artifacts"]["report"]["sha256"]:
            raise ValueError(f"{receipt_path}: report hash mismatch")
        report = verifier.load(report_path)
        checked = verifier.check_report(report)
        if checked["corpus"] != receipt["lane"]["corpus"] or checked["provider"] != receipt["lane"]["provider"]:
            raise ValueError(f"{receipt_path}: report lane mismatch")
        resource_path = safe_relative(receipt["artifacts"]["resource"]["path"])
        records[name] = {
            "name": name,
            "lane": receipt["lane"],
            "receipt": receipt,
            "report": report,
            "identity": report_fingerprint(report),
            "counters": counter_fingerprint(report),
            "rss_bytes": resource_rss(resource_path),
        }
    return records


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--include-range", action="store_true")
    args = parser.parse_args()
    protocol = json.loads((HERE / "protocol.json").read_text())
    before = load_phase("before", protocol)
    after = load_phase("after", protocol)
    core_names = {f"{lane['corpus']}-{lane['provider_label']}-{lane['repeat'].lower()}" for lane in protocol["core_order"]}
    optional_names = {f"{lane['corpus']}-{lane['provider_label']}-{lane['repeat'].lower()}" for lane in protocol["optional_range"]["order"]}
    expected = core_names | (optional_names if args.include_range else set())
    if set(before) != expected or set(after) != expected:
        raise ValueError(f"lane inventory mismatch: expected {sorted(expected)}, before {sorted(before)}, after {sorted(after)}")
    comparisons = []
    flags = []
    for name in sorted(expected):
        left, right = before[name], after[name]
        if left["identity"] != right["identity"]:
            flags.append({"lane": name, "metric": "identity", "kind": "oracle", "before": left["identity"], "after": right["identity"]})
        if left["counters"] != right["counters"]:
            flags.append({"lane": name, "metric": "source_cache_budget_counters", "kind": "oracle", "before_after_equal": False})
        for metric in TIMINGS:
            before_values = [row["timings"][metric] for row in left["report"]["samples_raw"]]
            after_values = [row["timings"][metric] for row in right["report"]["samples_raw"]]
            bdist, adist = distribution(before_values), distribution(after_values)
            for statistic in ("p50", "p95", "p99", "mean"):
                change = relative(bdist[statistic], adist[statistic])
                row = {"lane": name, "metric": metric, "statistic": statistic, "before": bdist[statistic], "after": adist[statistic], "relative_percent": change}
                comparisons.append(row)
                if change is not None and abs(change) > REVIEW_PERCENT:
                    flags.append({**row, "kind": "timing", "threshold_percent": REVIEW_PERCENT})
        bthroughput = [left["report"]["expected_output_bytes"] * 1_000_000_000 / row["timings"]["api_sum_ns"] for row in left["report"]["samples_raw"]]
        athroughput = [right["report"]["expected_output_bytes"] * 1_000_000_000 / row["timings"]["api_sum_ns"] for row in right["report"]["samples_raw"]]
        bd, ad = distribution([int(value) for value in bthroughput]), distribution([int(value) for value in athroughput])
        change = relative(bd["p50"], ad["p50"])
        row = {"lane": name, "metric": "throughput_bytes_per_second", "statistic": "p50", "before": bd["p50"], "after": ad["p50"], "relative_percent": change}
        comparisons.append(row)
        if change is not None and abs(change) > REVIEW_PERCENT:
            flags.append({**row, "kind": "throughput", "threshold_percent": REVIEW_PERCENT})
        rss_change = relative(left["rss_bytes"], right["rss_bytes"])
        rss_row = {"lane": name, "metric": "whole_child_rss_bytes", "statistic": "endpoint", "before": left["rss_bytes"], "after": right["rss_bytes"], "relative_percent": rss_change}
        comparisons.append(rss_row)
        if rss_change is not None and abs(rss_change) > REVIEW_PERCENT:
            flags.append({**rss_row, "kind": "rss", "threshold_percent": REVIEW_PERCENT})
    result = {
        "change": 501,
        "lanes": sorted(expected),
        "before_reports": len(before),
        "after_reports": len(after),
        "samples_per_report": protocol["samples"],
        "comparisons": comparisons,
        "flags": flags,
        "review_percent": REVIEW_PERCENT,
        "claims": protocol["claims"],
    }
    (HERE / "comparison.json").write_bytes(canonical_json(result))
    lines = [
        "# 0501 matched comparison",
        "",
        "Every report passed the independent lifecycle oracle. Flags are retained for review; they do not by themselves authorize the candidate.",
        "",
        "| Lane | Metric | Statistic | Before | After | Relative change |",
        "| --- | --- | --- | ---: | ---: | ---: |",
    ]
    for row in comparisons:
        change = "n/a" if row["relative_percent"] is None else f"{row['relative_percent']:+.3f}%"
        lines.append(f"| {row['lane']} | {row['metric']} | {row['statistic']} | {row['before']:.3f} | {row['after']:.3f} | {change} |")
    lines.extend(["", f"Retained review flags: {len(flags)}.", "", "The whole-child RSS endpoint includes setup, correctness checks, and report serialization. Source/cache counters are logical and managed-resource evidence, not physical I/O or allocation attribution."])
    (HERE / "comparison.md").write_text("\n".join(lines) + "\n")
    print(json.dumps({"status": "pass", "lanes": len(expected), "flags": len(flags)}, sort_keys=True))


if __name__ == "__main__":
    main()
