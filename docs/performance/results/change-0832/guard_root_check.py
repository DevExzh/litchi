"""Recompute guard statistics directly from raw reports without guard imports."""
import json
import math
from pathlib import Path
import random

import driver as custody

P = custody.P / "promotion-guard"


def read(path):
    return json.loads(path.read_text())


def median(values):
    x = sorted(values)
    return (x[(len(x) - 1) // 2] + x[len(x) // 2]) / 2


def main():
    analysis = read(P / "analysis.json")
    assert analysis["status"] == "pass"
    cases = {}
    for native, observer in zip(analysis["native"]["cases"], analysis["observer"]["cases"]):
        name = native["id"]
        assert name == observer["id"]
        metrics = {}
        for block in range(6):
            for leg in ("before", "after"):
                path = P / "native" / name / f"{block:02d}-{leg}.json"
                assert custody.sha(path) == analysis["capture"]["report_hashes"][str(path.relative_to(P))]
                samples = read(path)["results"][0]["elapsed_ns"]["samples"]
                assert len(samples) == 500 and samples == sorted(samples)
                metrics[block, leg] = {
                    **{f"p{q}_ns": samples[math.ceil(len(samples) * q / 100) - 1] for q in (50, 95, 99)},
                    "mean_ns": sum(samples) / len(samples),
                    "rss_kib": int(path.with_suffix(".rss").read_text())}
        for metric in metrics[0, "before"]:
            ratios = [metrics[b, "after"][metric] / metrics[b, "before"][metric] for b in range(6)]
            rng = random.Random(832128)
            bootstrap = sorted(median([ratios[rng.randrange(6)] for _ in range(6)]) for _ in range(10000))
            retained = native["paired_after_over_before"][metric]
            assert ratios == retained["block_values"]
            assert (median(ratios), bootstrap[250], bootstrap[9749]) == (
                retained["bootstrap"]["estimate"], retained["bootstrap"]["ci_low"], retained["bootstrap"]["ci_high"])
        counters = {}
        for block in range(2):
            for leg in ("before", "after"):
                path = P / "observer" / name / f"{block:02d}-{leg}.json"
                assert custody.sha(path) == analysis["capture"]["report_hashes"][str(path.relative_to(P))]
                allocation = read(path)["results"][0]["operation_metrics"]["allocation"]
                values = {key: value["values"] for key, value in allocation.items() if isinstance(value, dict)}
                assert len(values) == 11 and all(len(value) == 3 for value in values.values())
                assert values["failed_allocation_calls"] == [0, 0, 0]
                for index in range(3):
                    assert values["live_bytes_after"][index] - values["live_bytes_before"][index] == values["allocated_bytes"][index] - values["deallocated_bytes"][index]
                counter = {key: median(value) for key, value in values.items()}
                counter["peak_above_entry_bytes"] = median([p-e for p, e in zip(values["region_peak_live_bytes"], values["live_bytes_before"])])
                counter["net_live_bytes"] = median([a-b for a, b in zip(values["live_bytes_after"], values["live_bytes_before"])])
                assert counter == observer["legs"][leg]["block_values"][block]
                counters[block, leg] = counter
        for metric, retained in observer["after_minus_before"].items():
            delta = [counters[b, "after"][metric] - counters[b, "before"][metric] for b in range(2)]
            assert delta == retained["block_values"] and median(delta) == retained["median_delta"]
        cases[name] = {"native_reports": 12, "observer_reports": 4, "statistics_match": True}
    custody.write(custody.P / "guard-root-check.json", {
        "status": "pass", "cases": cases,
        "analysis_sha256": custody.sha(P / "analysis.json"),
        "script_sha256": custody.sha(Path(__file__)),
        "scope": "All 32 comparative reports: independent quantiles, paired bootstrap, RSS, allocator medians and signed peak deltas."})
    print("0832 root guard recomputation PASS: all 32 comparative reports")


if __name__ == "__main__":
    main()
