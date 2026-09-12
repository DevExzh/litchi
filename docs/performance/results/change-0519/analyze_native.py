#!/usr/bin/env python3
"""Validate and summarize the source-bound native DOCX campaign.

The input children are deliberately kept separate from Callgrind and perf
captures.  This script imports ``capture.validate`` so the row-level contract
has one implementation, then adds receipt, artifact, and campaign checks.
"""

import argparse
import csv
import hashlib
import importlib.util
import json
import math
import random
import re
import sys
from pathlib import Path


HERE = Path(__file__).resolve().parent
EXPECTED_NAMES = {
    f"p{paragraphs}-k{count}-{source}-{mode}"
    for paragraphs in (128, 512)
    for count in (1, 8, 32)
    for source in ("owned", "file")
    for mode in ("repeated", "batch")
}
PHASES = ("elapsed_ns", "open_ns", "edit_ns", "commit_ns", "publish_ns", "drop_ns")
STATS = ("p50", "p95", "p99", "mean")
BUDGET_COUNTERS = (
    "budget_before_memory", "budget_live_memory", "budget_after_memory",
    "budget_before_input", "budget_live_input", "budget_after_input",
    "budget_before_output", "budget_live_output", "budget_after_output",
    "budget_before_objects", "budget_live_objects", "budget_after_objects",
    "budget_before_work", "budget_live_work", "budget_after_work",
)
CACHE_COUNTERS = (
    "cache_before_cold_loads", "cache_before_successful_loads", "cache_before_hits",
    "cache_live_cold_loads", "cache_live_successful_loads", "cache_live_hits",
    "cache_live_retained_bytes", "cache_live_retained_entries", "cache_live_in_flight_loads",
)
SOURCE_COUNTERS = (
    "source_read_calls", "source_requested_bytes", "source_returned_bytes",
    "source_zero_length_calls",
)
COUNTERS = BUDGET_COUNTERS + CACHE_COUNTERS + SOURCE_COUNTERS
# This batch expects the Work charge to remain unchanged.  Retained live
# gauges are still retained and compared as descriptive data.  Release,
# output, and source-read fields remain guard fields in compare_candidate.py.
ALLOWED_COUNTER_DIFFERENCES = frozenset(
    field for field in COUNTERS
    if field.startswith("budget_live_") and field != "budget_live_work"
    or field.startswith("cache_live_")
)
GUARD_COUNTERS = tuple(field for field in COUNTERS if field not in ALLOWED_COUNTER_DIFFERENCES)
RSS_RE = re.compile(r"Maximum resident set size \(kbytes\):\s*(\d+)")
DEFAULT_SEED = 0x0519A11
DEFAULT_BOOTSTRAPS = 4000


def load_capture():
    """Load capture.py without invoking its CLI or capture loop."""
    sys.path.insert(0, str(HERE))
    try:
        spec = importlib.util.spec_from_file_location("change0519_capture", HERE / "capture.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        sys.path.pop(0)


capture = load_capture()


def sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def require(condition, message):
    if not condition:
        raise RuntimeError(message)


def read_json(path):
    require(path.is_file(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"invalid JSON: {path}: {exc}") from exc


def nearest(values, percentile):
    ordered = sorted(values)
    return ordered[max(0, math.ceil(len(ordered) * percentile / 100) - 1)]


def phase_stats(rows):
    result = {}
    for field in PHASES:
        values = [int(row[field]) for row in rows]
        result[field] = {
            name: (nearest(values, int(name[1:])) if name != "mean" else sum(values) / len(values))
            for name in STATS
        }
    return result


def one_stats(rows):
    """Keep timings and constant per-child gauges for the measured rows."""
    counters = {}
    for field in COUNTERS:
        values = {int(row[field]) for row in rows}
        require(len(values) == 1, f"{field} varies within measured rows")
        counters[field] = values.pop()
    guard_fields = (*capture.prior.BOOLEAN_KEYS, "budget_managed")
    guard_status = {
        field: all(row.get(field) == "true" for row in rows)
        for field in guard_fields
    }
    require(all(guard_status.values()), "a release/output/read guard is false")
    return {"samples": len(rows), "phases_ns": phase_stats(rows),
            "counters": counters, "guard_status": guard_status}


def ratio(numerator, denominator):
    return None if denominator == 0 else numerator / denominator


def pct(value):
    return None if value is None else (value - 1.0) * 100.0


def percentile_float(values, p):
    return nearest(values, p)


def bootstrap_ratio_of_medians(left_groups, right_groups, seed, iterations):
    """Cluster-preserving, unpaired route bootstrap for a median ratio.

    Each route and internal repeat is resampled independently at its original
    size, then the batch median is divided by the repeated median.  Two
    internal repeats are not independent processes, so this is an uncertainty
    display for the measured route distributions, not a host-generalization
    claim.
    """
    require(left_groups and len(left_groups) == len(right_groups),
            "cannot bootstrap route groups with different counts")
    require(all(left_groups) and all(right_groups), "cannot bootstrap empty route groups")
    rng = random.Random(seed)
    medians = []
    for _ in range(iterations):
        left_sample = []
        right_sample = []
        for left_group, right_group in zip(left_groups, right_groups):
            left_sample.extend(left_group[rng.randrange(len(left_group))] for _ in left_group)
            right_sample.extend(right_group[rng.randrange(len(right_group))] for _ in right_group)
        medians.append(ratio(nearest(right_sample, 50), nearest(left_sample, 50)))
    medians.sort()
    return {
        "low": percentile_float(medians, 2.5),
        "high": percentile_float(medians, 97.5),
        "iterations": iterations,
        "seed": seed,
        "unit": "unpaired measured route rows, resampled within each internal repeat; batch median / repeated median",
    }


def stable_seed(seed, *parts):
    payload = ":".join([str(seed), *map(str, parts)]).encode()
    return int.from_bytes(hashlib.sha256(payload).digest()[:8], "big")


def load_case(lane, name):
    folder = HERE / lane
    paths = {suffix: folder / f"{name}.{suffix}" for suffix in ("csv", "json", "stderr", "stdout")}
    for path in paths.values():
        require(path.is_file(), f"{lane}/{name}: missing {path.name}")
    receipt = read_json(paths["json"])
    require(receipt.get("exit_code") == 0, f"{lane}/{name}: command failed")
    require(receipt.get("cleanup_verified") is True, f"{lane}/{name}: scratch cleanup not verified")
    require(receipt.get("source_unchanged") is True, f"{lane}/{name}: source changed")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{lane}/{name}: missing artifact hashes")
    for suffix in ("csv", "stderr", "stdout"):
        path = paths[suffix]
        require(artifacts.get(path.name) == sha256(path), f"{lane}/{name}: {path.name} hash is not receipt-bound")
    require(receipt.get("source_manifest_sha256"), f"{lane}/{name}: source manifest is not bound")
    require(receipt.get("binary_sha256"), f"{lane}/{name}: binary is not bound")
    require(receipt.get("plan_sha256"), f"{lane}/{name}: plan is not bound")
    command = receipt.get("command", [])
    require(isinstance(command, list), f"{lane}/{name}: malformed command")
    command_tokens = {str(token) for token in command}
    require(not ("valgrind" in command_tokens or "perf" in command_tokens
                 or any(token.endswith("/valgrind") or token.endswith("/perf")
                        for token in command_tokens)),
            f"{lane}/{name}: profile/hardware command entered native lane")
    for value in ("--samples", "30", "--warmups", "3", "--repeats", "2"):
        require(value in command, f"{lane}/{name}: command missing {value}")

    with paths["csv"].open(newline="") as stream:
        rows = list(csv.DictReader(stream))
    capture.validate(name, rows, 30, 3, 2)
    measured = [row for row in rows if row["warmup"] == "false"]
    repeat_groups = {
        str(repeat): [row for row in measured if int(row["repeat"]) == repeat]
        for repeat in (0, 1)
    }
    require(all(len(group) == 30 for group in repeat_groups.values()), f"{lane}/{name}: repeat groups")
    match = RSS_RE.search(paths["stderr"].read_text(errors="replace"))
    require(match is not None, f"{lane}/{name}: RSS missing from stderr")
    paragraphs, replacements, source, mode = capture.prior.parse_name(name)
    identity = {key: measured[0][key] for key in capture.prior.IDENTITY_KEYS}
    receipt_binding = {
        "source_manifest_sha256": receipt["source_manifest_sha256"],
        "binary_sha256": receipt["binary_sha256"],
        "plan_sha256": receipt["plan_sha256"],
        "artifacts": artifacts,
    }
    if "candidate_plan_sha256" in receipt:
        receipt_binding["candidate_plan_sha256"] = receipt["candidate_plan_sha256"]
    return {
        "name": name,
        "paragraphs": paragraphs,
        "replacements": replacements,
        "source": source,
        "mode": mode,
        "rows": measured,
        "stats": one_stats(measured),
        "repeat_stats": {key: one_stats(value) for key, value in repeat_groups.items()},
        "rss_kib": int(match.group(1)),
        "identity": identity,
        "receipt": receipt_binding,
    }


def load_campaign(lane):
    folder = HERE / lane
    require(folder.is_dir(), f"missing native campaign directory: {folder}")
    require(not list(folder.glob("*.callgrind")) and not list(folder.glob("*.perf.csv")),
            f"{lane}: profile/hardware artifacts are not native inputs")
    names = {path.stem for path in folder.glob("*.csv")}
    require(names == EXPECTED_NAMES, f"{lane}: expected 24 native CSVs, got {len(names)}")
    for suffix in ("json", "stderr", "stdout"):
        found = {path.name.removesuffix(f".{suffix}") for path in folder.glob(f"*.{suffix}")}
        require(found == EXPECTED_NAMES, f"{lane}: {suffix} inventory")
    cases = {name: load_case(lane, name) for name in sorted(names)}
    bindings = {(case["receipt"]["source_manifest_sha256"],
                 case["receipt"]["binary_sha256"], case["receipt"]["plan_sha256"])
                for case in cases.values()}
    require(len(bindings) == 1, f"{lane}: source/binary/plan binding differs between cases")
    # Each route pair must be the same deterministic output fixture before it
    # is used as an API-choice comparison.
    for base in sorted({name.rsplit("-", 1)[0] for name in names}):
        left = cases[f"{base}-repeated"]["identity"]
        right = cases[f"{base}-batch"]["identity"]
        require(left == right, f"{lane}/{base}: repeated and batch identities differ")
    binding = next(iter(bindings))
    return {"lane": lane, "cases": cases, "binding": {
        "source_manifest_sha256": binding[0],
        "binary_sha256": binding[1],
        "plan_sha256": binding[2],
    }}


def metric_value(case, field, stat):
    if field == "rss_kib":
        return case["rss_kib"]
    return case["stats"]["phases_ns"][field][stat]


def metric_record(left, right, field, stat):
    before = metric_value(left, field, stat)
    after = metric_value(right, field, stat)
    return {"left": before, "right": after,
            "ratio": ratio(after, before), "delta_pct": pct(ratio(after, before))}


def threshold(stat):
    return 5.0 if stat in ("p50", "mean", "rss") else 10.0 if stat == "p95" else 15.0


def flags_for(metrics):
    flags = []
    for value in metrics.values():
        records = [value] if "delta_pct" in value else value.values()
        for record in records:
            change = record.get("delta_pct")
            if change is not None and abs(change) > threshold(record["stat"]):
                flags.append({"metric": record["metric"], "stat": record["stat"],
                              "delta_pct": change, "threshold_pct": threshold(record["stat"])})
    return flags


def api_choice(campaign, seed, bootstrap_iterations):
    result = []
    cases = campaign["cases"]
    bases = sorted({name.rsplit("-", 1)[0] for name in cases})
    for base in bases:
        repeated = cases[f"{base}-repeated"]
        batch = cases[f"{base}-batch"]
        metrics = {}
        for field in (*PHASES, "rss_kib"):
            stats = {}
            stat_names = ("rss",) if field == "rss_kib" else STATS
            for stat in stat_names:
                record = metric_record(repeated, batch, field, "p50" if stat == "rss" else stat)
                record.update(metric=field, stat=stat)
                stats[stat] = record
            if field != "rss_kib":
                left_groups = []
                right_groups = []
                for repeat in ("0", "1"):
                    left_group = [int(row[field]) for row in repeated["rows"]
                                  if row["repeat"] == repeat]
                    right_group = [int(row[field]) for row in batch["rows"]
                                   if row["repeat"] == repeat]
                    require(len(left_group) == len(right_group) == 30,
                            f"{campaign['lane']}/{base}: repeat group count")
                    left_groups.append(left_group)
                    right_groups.append(right_group)
                stats["p50"]["bootstrap95_ci_ratio_of_medians"] = bootstrap_ratio_of_medians(
                    left_groups, right_groups,
                    stable_seed(seed, campaign["lane"], base, field), bootstrap_iterations)
            metrics[field] = stats
        identity_equal = repeated["identity"] == batch["identity"]
        require(identity_equal, f"{campaign['lane']}/{base}: API-choice output identity differs")
        result.append({"case": base, "repeated": repeated["name"], "batch": batch["name"],
                       "metrics": metrics, "flags": flags_for(metrics),
                       "identity": {"repeated": repeated["identity"],
                                    "batch": batch["identity"],
                                    "equal": identity_equal}})
    return result


def compact_case(case):
    """Expose summaries and bindings while leaving raw rows in the CSVs."""
    return {key: value for key, value in case.items() if key != "rows"}


def counter_record(left, right, field):
    before = left["stats"]["counters"][field]
    after = right["stats"]["counters"][field]
    change = ratio(after, before)
    return {
        "left": before,
        "right": after,
        "delta": after - before,
        "ratio": change,
        "delta_pct": pct(change),
        "allowed_difference": field in ALLOWED_COUNTER_DIFFERENCES,
    }


def counter_drift(left, right):
    """Compare gauges while keeping expected Work/live changes visible."""
    counters = {field: counter_record(left, right, field) for field in COUNTERS}
    guard_flags = [
        {"counter": field, **record}
        for field, record in counters.items()
        if field in GUARD_COUNTERS and record["delta"] != 0
    ]
    allowed_differences = [
        {"counter": field, **record}
        for field, record in counters.items()
        if field in ALLOWED_COUNTER_DIFFERENCES and record["delta"] != 0
    ]
    return {"counters": counters, "guard_flags": guard_flags,
            "allowed_differences": allowed_differences}


def cross_campaign(first, second, seed, bootstrap_iterations):
    same_api = []
    api_ratio = []
    for name in sorted(first["cases"]):
        left = first["cases"][name]
        right = second["cases"][name]
        identity_equal = left["identity"] == right["identity"]
        require(identity_equal, f"{name}: same-API output identity differs between campaigns")
        metrics = {}
        for field in (*PHASES, "rss_kib"):
            stat_names = ("rss",) if field == "rss_kib" else STATS
            for stat in stat_names:
                key = f"{field}.{stat}"
                record = metric_record(left, right, field, "p50" if stat == "rss" else stat)
                record.update(metric=field, stat="rss" if field == "rss_kib" else stat)
                metrics[key] = record
        flags = flags_for(metrics)
        same_api.append({"case": name, "metrics": metrics, "flags": flags,
                         "identity": {"left": left["identity"],
                                      "right": right["identity"],
                                      "equal": identity_equal},
                         "counter_drift": counter_drift(left, right)})

    first_choices = {row["case"]: row for row in api_choice(first, seed, bootstrap_iterations)}
    second_choices = {row["case"]: row for row in api_choice(second, seed, bootstrap_iterations)}
    for base in sorted(first_choices):
        left = first_choices[base]
        right = second_choices[base]
        metrics = {}
        for field in (*PHASES, "rss_kib"):
            stat_names = ("rss",) if field == "rss_kib" else STATS
            for stat in stat_names:
                key = f"{field}.{stat}"
                left_record = left["metrics"][field][stat]
                right_record = right["metrics"][field][stat]
                left_ratio = left_record["ratio"]
                right_ratio = right_record["ratio"]
                record = {"left": left_ratio, "right": right_ratio,
                          "ratio": ratio(right_ratio, left_ratio),
                          "delta_pct": pct(ratio(right_ratio, left_ratio)),
                          "metric": field, "stat": "rss" if field == "rss_kib" else stat}
                metrics[key] = record
        api_ratio.append({"case": base, "metrics": metrics, "flags": flags_for(metrics)})
    return {"same_api": same_api, "api_choice_ratio": api_ratio}


def compact_number(value):
    return "—" if value is None else f"{value:.2f}"


def write_markdown(result, path):
    lines = [
        "# 0519 native DOCX analysis", "",
        f"Native campaigns: {', '.join(result['campaigns'])}. Rows are source-bound through each receipt; profile and hardware lanes are excluded.",
        "Measured rows are the 30 `warmup=false` samples in each of two internal repeats. Every case reports nearest-rank p50/p95/p99 and arithmetic mean for elapsed, open, edit, commit, publish, and drop clocks; RSS is the one whole-child maximum from GNU time.",
        f"The unpaired route bootstrap for the p50 ratio uses seed `{result['seed']}` and {result['bootstrap_iterations']} iterations, resampling each route within each internal repeat and taking batch median / repeated median. With only two internal repeats, its interval is descriptive for these row distributions and does not estimate independent-process or host variation.",
        "The historical 0500 validator is imported by `capture.py` and remains the row-level source for release, output, semantic, untouched-member, source-version, and readback guards. Work is expected to remain unchanged in this batch; retained live gauges are reported as descriptive counters and do not relax those guards.",
        "", "## API-choice ratios", "",
        "Ratios are batch / repeated; positive percentages mean batch took longer or used more RSS. The same absolute thresholds are used for route flags and cross-campaign flags: 5% for p50/mean/RSS, 10% for p95, and 15% for p99.", "",
        "| Campaign | Workload | elapsed p50 | edit p50 | publish p50 | RSS | route flags |",
        "| --- | --- | ---: | ---: | ---: | ---: | --- |",
    ]
    for lane in result["campaigns"]:
        for row in result["api_choice"][lane]:
            m = row["metrics"]
            vals = [m[field]["p50"]["delta_pct"] for field in ("elapsed_ns", "edit_ns", "publish_ns")]
            rss = m["rss_kib"]["rss"]["delta_pct"]
            lines.append(f"| {lane} | {row['case']} | {compact_number(vals[0])}% | {compact_number(vals[1])}% | {compact_number(vals[2])}% | {compact_number(rss)}% | {len(row['flags'])} |")
    lines += ["", "## API-choice flags", ""]
    route_flag_rows = [
        (lane, row) for lane in result["campaigns"]
        for row in result["api_choice"][lane] if row["flags"]
    ]
    lines.append(f"{sum(bool(row['flags']) for _, row in route_flag_rows)} route records have one or more absolute threshold flags.")
    if route_flag_rows:
        lines += ["", "| Campaign | Workload | flags |", "| --- | --- | --- |"]
        for lane, row in route_flag_rows:
            flags = "; ".join(
                f"{flag['metric']}.{flag['stat']} {flag['delta_pct']:+.2f}% (>{flag['threshold_pct']:.0f}%)"
                for flag in row["flags"]
            )
            lines.append(f"| {lane} | {row['case']} | {flags} |")
    lines += ["", "## Per-case gauges and guards", "",
              "The table exposes the Work charge, retained live reservations, release/output/input counters, and source reads for every case. Every listed guard is true; Work is expected unchanged and retained live-gauge deltas remain explicit comparison data.", "",
              "| Campaign | Workload | API | after Work | live Work | live memory | live objects | after input | after output | source reads | guards |",
              "| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |"]
    for lane in result["campaigns"]:
        for name, case in sorted(result["campaign_data"][lane]["cases"].items()):
            counters = case["stats"]["counters"]
            guard_ok = all(case["stats"]["guard_status"].values())
            lines.append(
                f"| {lane} | {name} | {case['mode']} | {counters['budget_after_work']} | "
                f"{counters['budget_live_work']} | {counters['budget_live_memory']} | "
                f"{counters['budget_live_objects']} | {counters['budget_after_input']} | "
                f"{counters['budget_after_output']} | {counters['source_read_calls']} | "
                f"{'all true' if guard_ok else 'FAIL'} |"
            )
    lines += ["", "## Cross-campaign flags", ""]
    cross = result["cross_campaign"]
    for title, key in (("Matched same API", "same_api"), ("API-choice ratio", "api_choice_ratio")):
        records = cross[key]
        flagged = [record for record in records if record["flags"]]
        lines += [f"### {title}", "", f"{len(flagged)} of {len(records)} records have one or more absolute threshold flags.", ""]
        if flagged:
            lines += ["| Workload | flags |", "| --- | --- |"]
            for record in flagged:
                flags = "; ".join(f"{flag['metric']}.{flag['stat']} {flag['delta_pct']:+.2f}% (>{flag['threshold_pct']:.0f}%)" for flag in record["flags"])
                lines.append(f"| {record['case']} | {flags} |")
            lines.append("")
        else:
            lines.append("No threshold flags.\n")
    same_api = cross["same_api"]
    guard_drift = [
        (record["case"], record["counter_drift"]["guard_flags"])
        for record in same_api if record["counter_drift"]["guard_flags"]
    ]
    allowed_drift = [
        (record["case"], record["counter_drift"]["allowed_differences"])
        for record in same_api if record["counter_drift"]["allowed_differences"]
    ]
    lines += ["", "## Counter drift", "",
              f"Same-API output identities match for {sum(record['identity']['equal'] for record in same_api)}/{len(same_api)} cases across campaigns.",
              f"Release/output/read guard counter changes: {len(guard_drift)} cases.",
              f"Allowed retained live-gauge changes: {len(allowed_drift)} cases; Work changes would be guard findings. Any values appear in JSON.", ""]
    if guard_drift:
        lines += ["| Workload | guard counter changes |", "| --- | --- |"]
        for name, records in guard_drift:
            details = "; ".join(f"{record['counter']} {record['delta']:+d}" for record in records)
            lines.append(f"| {name} | {details} |")
        lines.append("")
    report_name = path.with_suffix(".json").name
    lines += [f"All case/campaign phase statistics, repeat groups, route ratios, counter drift, identities, guards, and bootstrap intervals are in `{report_name}`.", ""]
    path.write_text("\n".join(lines))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--lanes", nargs="+", default=["r1", "r2"],
                        help="two native campaign directories (default: r1 r2)")
    parser.add_argument("--seed", type=int, default=DEFAULT_SEED)
    parser.add_argument("--bootstrap-iterations", type=int, default=DEFAULT_BOOTSTRAPS)
    parser.add_argument("--json", type=Path)
    parser.add_argument("--markdown", type=Path)
    args = parser.parse_args()
    require(len(args.lanes) == 2, "provide exactly two campaigns for cross-campaign analysis")
    require(args.bootstrap_iterations > 0, "bootstrap iterations must be positive")
    campaigns = {lane: load_campaign(lane) for lane in args.lanes}
    bindings = {tuple(campaign["binding"].values()) for campaign in campaigns.values()}
    require(len(bindings) == 1, "the two campaigns do not share one source/binary/plan binding")
    choices = {lane: api_choice(campaigns[lane], args.seed, args.bootstrap_iterations)
               for lane in args.lanes}
    result = {
        "schema": "managed_paragraph_native_analysis_0519_v1",
        "campaigns": args.lanes,
        "seed": args.seed,
        "bootstrap_iterations": args.bootstrap_iterations,
        "thresholds_percent": {"p50": 5, "mean": 5, "p95": 10, "p99": 15, "rss": 5},
        "scope": "native r1/r2 (or explicitly supplied equivalent lanes), measured warm=false rows only; phase clocks include returned Snapshot drop in publish_ns; RSS is whole-child",
        "binding": campaigns[args.lanes[0]]["binding"],
        "campaign_data": {lane: {"binding": campaigns[lane]["binding"],
                                  "cases": {name: compact_case(case) for name, case in campaigns[lane]["cases"].items()}}
                         for lane in args.lanes},
        "api_choice": choices,
        "cross_campaign": cross_campaign(campaigns[args.lanes[0]], campaigns[args.lanes[1]],
                                           args.seed, args.bootstrap_iterations),
    }
    # Do not include raw rows in the report; retaining the CSVs keeps the
    # source evidence authoritative and makes this summary stable and small.
    default_stem = "candidate-native-analysis" if all(lane.startswith("after-") for lane in args.lanes) else "baseline-native-analysis"
    json_path = args.json or HERE / (default_stem + ".json")
    markdown_path = args.markdown or HERE / (default_stem + ".md")
    json_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    write_markdown(result, markdown_path)
    flag_count = sum(len(row["flags"]) for group in result["cross_campaign"].values() for row in group)
    same_api = result["cross_campaign"]["same_api"]
    guard_drift_count = sum(bool(row["counter_drift"]["guard_flags"]) for row in same_api)
    allowed_drift_count = sum(bool(row["counter_drift"]["allowed_differences"]) for row in same_api)
    print(json.dumps({"campaigns": args.lanes, "cases_per_campaign": 24,
                      "cross_campaign_flag_count": flag_count,
                      "cross_campaign_guard_drift_count": guard_drift_count,
                      "cross_campaign_allowed_work_live_drift_count": allowed_drift_count,
                      "json": str(json_path), "markdown": str(markdown_path)}, sort_keys=True))


if __name__ == "__main__":
    main()
