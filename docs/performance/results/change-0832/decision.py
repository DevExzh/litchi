"""Replay the root adoption decision and baseline-adjusted peak diagnostics."""
import statistics
import driver as d


def derive():
    analysis = d.read(d.P / "analysis.json")
    audit = d.read(d.P / "audit.json")
    assert analysis["status"] == audit["status"] == "pass"
    assert analysis["schema"] == "litchi.performance.0832.allocation-analysis.v1"
    assert audit["schema"] == "litchi.performance.0832.independent-audit.v1"
    assert analysis["base"] == audit["base"] == d.BASE
    assert analysis["counts"] == audit["counts"]
    assert analysis["counts"]["reports"] == 180 and analysis["counts"]["samples"] == 20304
    for evidence in (analysis, audit):
        assert evidence["qualification"]["oracle_checked"] is True
        assert evidence["qualification"]["reports"] == evidence["qualification"]["samples"] == 36
        assert evidence["claims"]["observer_latency"] is False
        assert evidence["claims"]["automatic_adoption_decision"] is False
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
    for name, case in peaks.items():
        for before_block, after_block in zip(case["before"]["blocks"], case["after"]["blocks"]):
            if after_block["peak_above_entry_bytes"] > before_block["peak_above_entry_bytes"]:
                flags.append({"case": name, "metric": "peak_above_entry_bytes"})
    promotion = d.read(d.P / "promotion-guard/analysis.json")
    assert promotion["status"] == "pass"
    assert promotion["schema"] == "litchi.performance.0832.column-promotion-guard.analysis.v1"
    assert promotion["base"] == d.BASE
    assert promotion["counts"]["reports"] == 48 and promotion["counts"]["samples"] == 12040
    assert promotion["qualification"]["reports"] == promotion["qualification"]["samples"] == 16
    assert promotion["qualification"]["output_preservation"]["status"] == "pass"
    oracles = promotion["qualification"]["lifecycle_oracle"]
    assert set(oracles) == {"disjoint-two-records", "overlap-128-wide-complete"}
    assert all(row["equal_across_before_after_and_native_observer"] is True for row in oracles.values())
    assert promotion["claims"]["observer_latency"] is False
    assert promotion["claims"]["automatic_adoption_decision"] is False
    frozen = d.read(d.P / "promotion-guard/freeze.json")
    assert promotion["reader"]["sha256"] == frozen["scripts"]["reader.py"]["sha256"]
    assert promotion["reader"]["sha256"] == d.sha(d.P / "promotion-guard/reader.py")
    assert promotion["freeze"]["sha256"] == d.sha(d.P / "promotion-guard/freeze.json")
    correction = d.read(d.P / "promotion-guard/reader-correction-v2.json")
    corrected = d.read(d.P / "promotion-guard/corrected-analysis.json")
    assert correction["status"] == corrected["status"] == "pass"
    assert correction["missing_globals"] == ["SOURCES", "write_once"]
    assert correction["loader_sha256"] == d.sha(d.P / "promotion-guard/replay_reader_v2.py")
    assert correction["previous_loader_sha256"] == d.sha(d.P / "promotion-guard/replay_reader.py")
    assert corrected["correction_sha256"] == d.sha(d.P / "promotion-guard/reader-correction-v2.json")
    assert corrected["analysis_sha256"] == d.sha(d.P / "promotion-guard/analysis.json")
    assert corrected["corrected_admission_sha256"] == d.sha(d.P / "promotion-guard/corrected-admission.json")
    assert isinstance(promotion["regression_flags"], list)
    for flag in promotion["regression_flags"]:
        assert isinstance(flag, dict)
        flags.append({"guard": "promotion", **flag})
    adopt = reduction >= .50 and not flags
    return {"schema": "litchi.performance.0832.root-decision.v1", "adopt": adopt,
        "reason": "Adopt inline first-record storage only if the frozen 50% requested-byte gate passes with preserved output and no unresolved regression flag.",
        "analysis_sha256": d.sha(d.P / "analysis.json"), "audit_sha256": d.sha(d.P / "audit.json"),
        "promotion_guard": {"analysis_sha256": d.sha(d.P / "promotion-guard/analysis.json"),
                            "reader_correction_sha256": d.sha(d.P / "promotion-guard/reader-correction-v2.json"),
                            "corrected_analysis_sha256": d.sha(d.P / "promotion-guard/corrected-analysis.json"),
                            "regression_flags": promotion["regression_flags"]},
        "real_edit_requested_bytes": {"before": before, "after": after,
            "reduction_bytes": before-after, "reduction_fraction": reduction},
        "regression_flags": flags,
        "descriptive_spread_flags": {name: row["spread_flags"] for name, row in native.items()},
        "descriptive_tail_flags": {name: row["tail_flags"] for name, row in native.items()},
        "peak_diagnostic": {"derivation": "Within each observer process, median of region_peak_live_bytes minus live_bytes_before; midpoint median across two processes.", "cases": peaks},
        "limits": ["One small real XLSX fixture and fixed synthetic cases; no cross-format speedup claim.",
            "Full-output byte preservation is checked by separate pinned oracle tests; timed real edits check admitted outcomes.",
            "Supplemental controlled promotion fixtures do not establish general producer or nonempty-column-action performance.",
            "Process RSS includes harness setup; observer elapsed times do not support latency claims."]}



def main():
    value = derive()
    path = d.P / "decision.json"
    if path.exists():
        assert d.read(path) == value, "retained root decision did not replay"
    else:
        d.write(path, value)
    print("0832 root decision replay PASS; adopt", value["adopt"], "requested-byte reduction", value["real_edit_requested_bytes"]["reduction_fraction"])


if __name__ == "__main__":
    main()
