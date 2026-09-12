#!/usr/bin/env python3
"""Compare source-bound baseline and candidate native DOCX captures.

This driver reuses ``analyze_native`` for receipt/artifact validation and the
``capture.validate`` row contract.  It compares each route independently:
baseline r1 -> candidate after-r1 and baseline r2 -> candidate after-r2.
"""

import argparse
import json
from collections import Counter
from pathlib import Path

import analyze_native as native


HERE = native.HERE
PHASES = native.PHASES
STATS = native.STATS


def require(condition, message):
    if not condition:
        raise RuntimeError(message)


def variant_for(lane):
    return "candidate" if lane.startswith("after-") else "baseline"


def verify_anchor(lane, campaign):
    """Bind a campaign to the corresponding frozen manifest/build receipt."""
    variant = variant_for(lane)
    manifest = HERE / variant / "source-manifest.json"
    build_receipt_path = HERE / variant / "build-receipt.json"
    if manifest.is_file() and build_receipt_path.is_file():
        build = native.read_json(build_receipt_path)
        require(campaign["binding"]["source_manifest_sha256"] == native.sha256(manifest),
                f"{lane}: receipts are not bound to {manifest}")
        require(campaign["binding"]["binary_sha256"] == build.get("binary_sha256"),
                f"{lane}: receipts are not bound to {variant} build")
    plan = HERE / "plan.json"
    if plan.is_file():
        require(campaign["binding"]["plan_sha256"] == native.sha256(plan),
                f"{lane}: receipt plan binding differs from plan.json")
    if variant == "candidate":
        candidate_plan = HERE / "candidate-plan.json"
        if candidate_plan.is_file():
            expected = native.sha256(candidate_plan)
            for case in campaign["cases"].values():
                require(case["receipt"].get("candidate_plan_sha256") == expected,
                        f"{lane}/{case['name']}: candidate plan binding")


def check_output_identity(campaigns):
    names = sorted(native.EXPECTED_NAMES)
    for name in names:
        identities = {tuple(sorted(campaign["cases"][name]["identity"].items()))
                      for campaign in campaigns}
        require(len(identities) == 1, f"{name}: baseline/candidate output identity differs")
    return {name: campaigns[0]["cases"][name]["identity"] for name in names}


def candidate_metric_record(baseline, candidate, field, stat, seed, iterations):
    baseline_value = native.metric_value(baseline, field, "p50" if stat == "rss" else stat)
    candidate_value = native.metric_value(candidate, field, "p50" if stat == "rss" else stat)
    change = native.ratio(candidate_value, baseline_value)
    record = {
        "metric": field,
        "stat": stat,
        "baseline": baseline_value,
        "candidate": candidate_value,
        "ratio_candidate_over_baseline": change,
        "delta_pct": native.pct(change),
        "threshold_pct": native.threshold(stat),
    }
    if field == "rss_kib":
        # GNU time emits one high-water RSS value per child, so no synthetic
        # phase-local bootstrap is reported for this metric.
        record["bootstrap95_ci_ratio_of_medians"] = None
    elif stat == "p50":
        left_groups = []
        right_groups = []
        for repeat in ("0", "1"):
            left = [int(row[field]) for row in baseline["rows"] if row["repeat"] == repeat]
            right = [int(row[field]) for row in candidate["rows"] if row["repeat"] == repeat]
            require(len(left) == len(right) == 30,
                    f"{baseline['name']} / {candidate['name']}: repeat group count")
            left_groups.append(left)
            right_groups.append(right)
        interval = native.bootstrap_ratio_of_medians(
            left_groups, right_groups,
            native.stable_seed(seed, baseline["name"], candidate["name"], field), iterations)
        interval["unit"] = "unpaired measured route rows, resampled within each internal repeat; candidate median / baseline median"
        record["bootstrap95_ci_ratio_of_medians"] = interval
    return record


def compare_pair(baseline, candidate, baseline_lane, candidate_lane, seed, iterations):
    records = []
    for name in sorted(native.EXPECTED_NAMES):
        base_case = baseline["cases"][name]
        candidate_case = candidate["cases"][name]
        metrics = {}
        for field in (*PHASES, "rss_kib"):
            stat_names = ("rss",) if field == "rss_kib" else STATS
            metrics[field] = {
                stat: candidate_metric_record(base_case, candidate_case, field, stat, seed, iterations)
                for stat in stat_names
            }
        flags = [
            {"metric": record["metric"], "stat": record["stat"],
             "delta_pct": record["delta_pct"], "threshold_pct": record["threshold_pct"]}
            for values in metrics.values() for record in values.values()
            if record["delta_pct"] is not None and record["delta_pct"] > record["threshold_pct"]
        ]
        counters = {}
        for field in native.COUNTERS:
            base_value = base_case["stats"]["counters"][field]
            candidate_value = candidate_case["stats"]["counters"][field]
            change = native.ratio(candidate_value, base_value)
            counters[field] = {
                "baseline": base_value,
                "candidate": candidate_value,
                "delta": candidate_value - base_value,
                "delta_pct": native.pct(change),
            }
        records.append({
            "case": name,
            "baseline_campaign": baseline_lane,
            "candidate_campaign": candidate_lane,
            "metrics": metrics,
            "output_identity": {"baseline": base_case["identity"],
                                "candidate": candidate_case["identity"],
                                "equal": base_case["identity"] == candidate_case["identity"]},
            "counters": counters,
            "adverse_flags": flags,
        })
    return records


def public_binding(campaign):
    binding = {key: value for key, value in campaign["binding"].items()}
    candidate_plans = {
        case["receipt"].get("candidate_plan_sha256")
        for case in campaign["cases"].values()
        if case["receipt"].get("candidate_plan_sha256") is not None
    }
    if candidate_plans:
        require(len(candidate_plans) == 1, f"{campaign['lane']}: candidate plan binding differs")
        binding["candidate_plan_sha256"] = next(iter(candidate_plans))
    return binding


def p50_cell(record, field):
    metric = record["metrics"][field]["p50" if field != "rss_kib" else "rss"]
    unit = "KiB" if field == "rss_kib" else "ns"
    return (f"{metric['ratio_candidate_over_baseline']:.4f} "
            f"({metric['baseline']}→{metric['candidate']} {unit})")


def flag_cell(record, flag):
    metric = record["metrics"][flag["metric"]][flag["stat"]]
    unit = "KiB" if flag["metric"] == "rss_kib" else "ns"
    return (f"{flag['metric']}.{flag['stat']} {flag['delta_pct']:+.2f}% "
            f"({metric['baseline']}→{metric['candidate']} {unit}; >{flag['threshold_pct']:.0f}%)")


def write_markdown(result, path):
    lines = [
        "# 0517 baseline/candidate native comparison", "",
        "Each record compares the same API route and workload in one baseline campaign and its matching candidate campaign. Native rows were validated through `capture.validate`; profile and hardware lanes are excluded.",
        "All phase metrics are in nanoseconds and use nearest-rank p50/p95/p99 plus arithmetic mean over 60 measured rows. RSS is the one whole-child GNU time maximum in KiB.",
        f"The p50 intervals use a fixed seed `{result['seed']}` and {result['bootstrap_iterations']} unpaired route bootstrap iterations. Each route is resampled within each internal repeat, then candidate median / baseline median is calculated. The two internal repeats are not independent processes, so intervals are descriptive only. RSS has no phase-local bootstrap because each child contributes one high-water observation.",
        "", "## Pair summary", "",
        "A positive delta is a candidate regression. Adverse flags are directional and use 5% for p50/mean/RSS, 10% for p95, and 15% for p99.", "",
        "| Baseline | Candidate | Records | Flagged records | Flags |",
        "| --- | --- | ---: | ---: | ---: |",
    ]
    for pair in result["pairs"]:
        records = result["comparisons"][pair["id"]]
        flags = [flag for record in records for flag in record["adverse_flags"]]
        lines.append(f"| {pair['baseline']} | {pair['candidate']} | {len(records)} | {sum(bool(record['adverse_flags']) for record in records)} | {len(flags)} |")
    lines += ["", "## p50 route values", "",
              "Each cell is candidate / baseline followed by the absolute p50 values. This table covers all 24 cases in both campaign pairs.", "",
              "| Pair | Workload | elapsed p50 ratio | publish p50 ratio | RSS ratio |",
              "| --- | --- | ---: | ---: | ---: |"]
    for pair in result["pairs"]:
        for record in result["comparisons"][pair["id"]]:
            lines.append(f"| {pair['id']} | {record['case']} | {p50_cell(record, 'elapsed_ns')} | {p50_cell(record, 'publish_ns')} | {p50_cell(record, 'rss_kib')} |")
    lines += ["", "## Identity and counters", "",
              "Output identities are checked before comparison. Work and source-read counters are reported per case in JSON; this table shows whether any changed.", "",
              "| Pair | Identity matches | Work changed | Source read calls changed | Source bytes changed |",
              "| --- | ---: | ---: | ---: | ---: |"]
    for pair in result["pairs"]:
        records = result["comparisons"][pair["id"]]
        work_changed = sum(record["counters"]["budget_after_work"]["delta"] != 0 for record in records)
        read_calls_changed = sum(record["counters"]["source_read_calls"]["delta"] != 0 for record in records)
        source_bytes_changed = sum(any(record["counters"][field]["delta"] != 0 for field in ("budget_after_input", "source_requested_bytes", "source_returned_bytes")) for record in records)
        lines.append(f"| {pair['id']} | {sum(record['output_identity']['equal'] for record in records)}/{len(records)} | {work_changed} | {read_calls_changed} | {source_bytes_changed} |")
    lines += ["", "## Adverse flags", ""]
    for pair in result["pairs"]:
        lines += [f"### {pair['baseline']} → {pair['candidate']}", ""]
        records = result["comparisons"][pair["id"]]
        flagged = [record for record in records if record["adverse_flags"]]
        if not flagged:
            lines.append("No adverse flags.\n")
            continue
        lines += ["| Workload | Flags |", "| --- | --- |"]
        for record in flagged:
            values = "; ".join(flag_cell(record, flag) for flag in record["adverse_flags"])
            lines.append(f"| {record['case']} | {values} |")
        lines.append("")
    lines += ["The JSON contains all 48 case comparisons, every phase/statistic, RSS, bindings, output identities, and bootstrap metadata.", ""]
    path.write_text("\n".join(lines))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", nargs=2, default=["r1", "r2"], metavar=("R1", "R2"))
    parser.add_argument("--candidate", nargs=2, default=["after-r1", "after-r2"], metavar=("AFTER_R1", "AFTER_R2"))
    parser.add_argument("--seed", type=int, default=native.DEFAULT_SEED)
    parser.add_argument("--bootstrap-iterations", type=int, default=native.DEFAULT_BOOTSTRAPS)
    parser.add_argument("--json", type=Path, default=HERE / "candidate-comparison.json")
    parser.add_argument("--markdown", type=Path, default=HERE / "candidate-comparison.md")
    args = parser.parse_args()
    require(args.baseline[0] != args.baseline[1] and args.candidate[0] != args.candidate[1],
            "baseline and candidate campaigns must each be distinct")
    require(args.bootstrap_iterations > 0, "bootstrap iterations must be positive")
    baseline = {lane: native.load_campaign(lane) for lane in args.baseline}
    candidate = {lane: native.load_campaign(lane) for lane in args.candidate}
    for lane, campaign in (*baseline.items(), *candidate.items()):
        verify_anchor(lane, campaign)
    require(baseline[args.baseline[0]]["binding"] == baseline[args.baseline[1]]["binding"],
            "baseline campaigns do not share one source/binary/plan binding")
    require(candidate[args.candidate[0]]["binding"] == candidate[args.candidate[1]]["binding"],
            "candidate campaigns do not share one source/binary/plan binding")
    all_campaigns = [baseline[args.baseline[0]], baseline[args.baseline[1]],
                     candidate[args.candidate[0]], candidate[args.candidate[1]]]
    identities = check_output_identity(all_campaigns)
    pairs = [
        {"id": "r1", "baseline": args.baseline[0], "candidate": args.candidate[0]},
        {"id": "r2", "baseline": args.baseline[1], "candidate": args.candidate[1]},
    ]
    comparisons = {
        pair["id"]: compare_pair(baseline[pair["baseline"]], candidate[pair["candidate"]],
                                 pair["baseline"], pair["candidate"], args.seed,
                                 args.bootstrap_iterations)
        for pair in pairs
    }
    flag_counts = Counter(
        f"{flag['metric']}.{flag['stat']}"
        for records in comparisons.values() for record in records
        for flag in record["adverse_flags"])
    result = {
        "schema": "managed_paragraph_native_candidate_comparison_v1",
        "pairs": pairs,
        "seed": args.seed,
        "bootstrap_iterations": args.bootstrap_iterations,
        "thresholds_percent": {"p50": 5, "mean": 5, "p95": 10, "p99": 15, "rss": 5},
        "scope": "same-API native baseline/candidate comparison; 24 cases per pair; warm=false rows only; publish_ns includes returned Snapshot drop; RSS is whole-child",
        "bindings": {"baseline": {lane: public_binding(baseline[lane]) for lane in args.baseline},
                     "candidate": {lane: public_binding(candidate[lane]) for lane in args.candidate}},
        "output_identities": identities,
        "comparisons": comparisons,
        "adverse_flag_counts": dict(sorted(flag_counts.items())),
    }
    args.json.parent.mkdir(parents=True, exist_ok=True)
    args.json.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    write_markdown(result, args.markdown)
    print(json.dumps({"pairs": len(pairs), "records": sum(map(len, comparisons.values())),
                      "adverse_flags": sum(flag_counts.values()),
                      "json": str(args.json), "markdown": str(args.markdown)}, sort_keys=True))


if __name__ == "__main__":
    main()
