"""Replay the root adoption decision and baseline-adjusted peak diagnostics."""
import statistics
import driver as d


def derive():
    analysis = d.read(d.P / "analysis.json")
    audit = d.read(d.P / "audit.json")
    assert analysis["status"] == audit["status"] == "pass"
    assert audit["analysis"]["matched"] is True
    assert audit["analysis"]["sha256"] == d.sha(d.P / "analysis.json")
    native = {row["id"]: row for row in analysis["native"]["cases"]}
    observer = {row["id"]: row for row in analysis["observer"]["cases"]}
    flags = []
    peaks = {}
    for name, row in native.items():
        ratios = row["paired_after_over_before"]
        if ratios["p50_ns"]["bootstrap"]["ci_low"] > 1.05:
            flags.append({"case": name, "metric": "native_p50_lower_ci"})
        if ratios["rss_kib"]["bootstrap"]["estimate"] > 1.05:
            flags.append({"case": name, "metric": "native_paired_rss"})
        for field in ("allocation_calls", "allocated_bytes", "region_peak_live_bytes"):
            if any(value > 0 for value in observer[name]["after_minus_before"][field]["block_values"]):
                flags.append({"case": name, "metric": field})
        peaks[name] = {}
        for leg in ("before", "after"):
            blocks = []
            for block in range(2):
                path = d.P / "observer" / f"{block:02}-{leg}-{name}.json"
                counters = d.read(path)["results"][0]["operation_metrics"]["allocation"]
                assert counters["failed_allocation_calls"]["values"] == [0, 0, 0]
                entry = counters["live_bytes_before"]["values"]
                peak = counters["region_peak_live_bytes"]["values"]
                blocks.append({"entry_live_bytes": statistics.median(entry),
                    "peak_above_entry_bytes": statistics.median(p - e for p, e in zip(peak, entry)),
                    "report_sha256": d.sha(path)})
            peaks[name][leg] = {"blocks": blocks,
                "entry_live_bytes": statistics.median(b["entry_live_bytes"] for b in blocks),
                "peak_above_entry_bytes": statistics.median(b["peak_above_entry_bytes"] for b in blocks)}
    real = observer["real-edit"]
    before = real["legs"]["before"]["midpoint_medians"]["allocated_bytes"]
    after = real["legs"]["after"]["midpoint_medians"]["allocated_bytes"]
    reduction = (before - after) / before
    assert reduction >= .15 and not flags
    assert real["after_minus_before"]["allocated_bytes"]["block_values"] == [-524288, -524288]
    assert real["after_minus_before"]["allocation_calls"]["block_values"] == [-1, -1]
    return {"schema": "litchi.performance.0831.root-decision.v1", "adopt": True,
        "reason": "Remove the unused empty-column-action map: the frozen requested-byte benefit gate passes with preserved output and no frozen regression flag.",
        "analysis_sha256": d.sha(d.P / "analysis.json"), "audit_sha256": d.sha(d.P / "audit.json"),
        "real_edit_requested_bytes": {"before": before, "after": after,
            "reduction_bytes": before-after, "reduction_fraction": reduction},
        "regression_flags": flags,
        "descriptive_spread_flags": {name: row["spread_flags"] for name, row in native.items()},
        "descriptive_tail_flags": {name: row["tail_flags"] for name, row in native.items()},
        "peak_diagnostic": {"derivation": "Within each observer process, median of region_peak_live_bytes minus live_bytes_before; midpoint median across two processes.", "cases": peaks},
        "limits": ["One small real XLSX fixture and fixed synthetic cases; no cross-format speedup claim.",
            "Real lifecycle confidence interval includes 1.0; no full-save speedup claim.",
            "Real-edit peak above entry is unchanged; no real peak-memory or process-RSS reduction claim.",
            "Twenty native spread flags and ten p99/p50 diagnostics limit tail claims.",
            "Observer elapsed times are excluded from latency claims."]}


def main():
    value = derive()
    path = d.P / "decision.json"
    if path.exists():
        assert d.read(path) == value, "retained root decision did not replay"
    else:
        d.write(path, value)
    print("0831 root decision PASS: adopt; requested-byte reduction", value["real_edit_requested_bytes"]["reduction_fraction"])


if __name__ == "__main__":
    main()
