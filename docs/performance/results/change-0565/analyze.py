#!/usr/bin/env python3
"""Analyze the change-0565 A1/B1/B2/A2 latency matrix (and, optionally, the
before/after deterministic counters and syscall isolation summaries).

Statistics follow the retained 0560/0562 analyzers exactly
(docs/performance/results/change-0562/analyze_latency.py, whose quantile and
mean definitions reproduce docs/performance/results/change-0560/analysis.json):
p50/p95/p99 are linear-interpolated quantiles over the sorted 60 samples at
position fraction * (n - 1); mean is statistics.fmean. A cell improves only
when both directions (B1 vs A1 and B2 vs A2) are lower; a cell adverse in both
directions by more than five percent is a review trigger and is listed with its
absolute nanosecond values.

Usage:
  analyze.py --latency DIR [--output analysis.json] [--matrix matrix.json]
             [--before-summary before/summary.json --after-summary after/summary.json]
             [--perf-abba-summary /path/to/tools/perf_abba_summary.py --strict-dir DIR]
"""

from __future__ import annotations

import argparse
import json
import pathlib
import statistics
import subprocess
import sys

DEFAULT_LEGS = ("A1", "B1", "B2", "A2")
LEGS = DEFAULT_LEGS
STATS = ("p50", "mean", "p95", "p99")
TRIGGER_PERCENT = 5.0


def quantile(values: list[int], fraction: float) -> float:
    ordered = sorted(values)
    if len(ordered) == 1:
        return float(ordered[0])
    position = fraction * (len(ordered) - 1)
    low = int(position)
    high = min(low + 1, len(ordered) - 1)
    weight = position - low
    return ordered[low] * (1.0 - weight) + ordered[high] * weight


def percent(new: float, old: float) -> float:
    return (new - old) / old * 100.0


def load_report(path: pathlib.Path) -> dict:
    report = json.loads(path.read_text())
    cells = {}
    for record in report.get("results", []):
        samples = record.get("elapsed_ns", {}).get("samples")
        if not samples:
            continue
        corpus = (record.get("corpus") or {}).get("name")
        cells[(record["case"], str(corpus))] = {
            "p50": quantile(samples, 0.50),
            "mean": statistics.fmean(samples),
            "p95": quantile(samples, 0.95),
            "p99": quantile(samples, 0.99),
            "samples": len(samples),
            "harness_p50": record["elapsed_ns"].get("p50"),
        }
    identity = {
        "binary_sha256": report["binary_identity"]["binary_sha256"],
        "binary_path": report["binary_identity"]["path"],
        "git_revision": report["environment"].get("git_revision"),
        "git_worktree_dirty": report["environment"].get("git_worktree_dirty"),
        "cpu_affinity": report["environment"].get("cpu_affinity"),
        "logical_cpus_available": report["environment"].get("logical_cpus_available"),
        "samples_per_case": report["configuration"]["samples_per_case"],
        "warmup_iterations_per_case": report["configuration"]["warmup_iterations_per_case"],
        "cases": report["configuration"]["cases"],
    }
    return cells, identity


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--latency", required=True, help="directory holding abba-<leg>-<selector>.json")
    parser.add_argument("--output", required=True)
    parser.add_argument("--prefix", default="abba-")
    parser.add_argument("--legs", default=",".join(DEFAULT_LEGS),
                        help="comma-separated leg names in A1,B1,B2,A2 order "
                             "(change-0560 named its candidate legs E1/E2; pass A1,E1,E2,A2 to replay it)")
    parser.add_argument("--matrix", help="matrix.json written by run_abba.sh (copied into the analysis)")
    parser.add_argument("--before-summary", help="before/summary.json from capture_before.sh")
    parser.add_argument("--after-summary", help="after/summary.json from capture_before.sh run against the candidate")
    parser.add_argument("--perf-abba-summary", help="path to tools/perf_abba_summary.py for a strict per-selector cross-check")
    parser.add_argument("--strict-dir", help="where the strict cross-check summaries are written")
    args = parser.parse_args()
    global LEGS
    LEGS = tuple(part.strip() for part in args.legs.split(",") if part.strip())
    if len(LEGS) != 4:
        raise SystemExit("--legs needs exactly four names in A1,B1,B2,A2 order")
    root = pathlib.Path(args.latency)
    prefix = args.prefix

    first = LEGS[0]
    # run_abba.sh writes a `<name>.receipt.json` sidecar next to every report, so the
    # glob has to drop those explicitly or each one is read as if it were a report.
    names = sorted({p.name[len(prefix) + len(first) + 1:-len(".json")]
                    for p in root.glob(f"{prefix}{first}-*.json")
                    if not p.name.endswith(".receipt.json")})
    rows, triggers, improvements, legs_identity = [], [], [], {}
    incomplete = []
    for name in names:
        legs = {}
        for leg in LEGS:
            path = root / f"{prefix}{leg}-{name}.json"
            if not path.exists():
                legs = {}
                break
            legs[leg], legs_identity[(leg, name)] = load_report(path)
        if not legs:
            incomplete.append(name)
            rows.append({"case": name, "status": "incomplete"})
            continue
        for key in sorted(set.intersection(*(set(v) for v in legs.values()))):
            entry = {"case": key[0], "corpus": key[1], "samples": legs["A1"][key]["samples"], "statistics": {}}
            for stat in STATS:
                a1, b1, b2, a2 = (legs[l][key][stat] for l in LEGS)
                first, second = percent(b1, a1), percent(b2, a2)
                cell = {
                    "first_direction_percent": first,
                    "second_direction_percent": second,
                    "improves_both": first < 0.0 and second < 0.0,
                    "adverse_both": first > 0.0 and second > 0.0,
                    "baseline_ns": a1,
                    "candidate_ns": b1,
                    "second_baseline_ns": a2,
                    "second_candidate_ns": b2,
                    "same_binary_drift_percent": {f"baseline_{LEGS[3]}_vs_{LEGS[0]}": percent(a2, a1),
                                                  f"candidate_{LEGS[2]}_vs_{LEGS[1]}": percent(b2, b1)},
                }
                entry["statistics"][stat] = cell
                if cell["improves_both"]:
                    improvements.append({"case": key[0], "corpus": key[1], "statistic": stat,
                                         "first_direction_percent": first, "second_direction_percent": second,
                                         "baseline_ns": a1, "candidate_ns": b1})
                if first > TRIGGER_PERCENT and second > TRIGGER_PERCENT:
                    triggers.append({"case": key[0], "corpus": key[1], "statistic": stat,
                                     "first_direction_percent": first, "second_direction_percent": second,
                                     "baseline_ns": a1, "candidate_ns": b1,
                                     "second_baseline_ns": a2, "second_candidate_ns": b2})
            rows.append(entry)

    # Identity gates: one binary per stage, stages differ, same configuration and pin everywhere.
    stage_sha = {"baseline": set(), "candidate": set()}
    configs, pins = set(), set()
    for (leg, _name), ident in legs_identity.items():
        stage_sha["baseline" if leg in (LEGS[0], LEGS[3]) else "candidate"].add(ident["binary_sha256"])
        configs.add((ident["samples_per_case"], ident["warmup_iterations_per_case"]))
        pins.add((ident["cpu_affinity"], ident["logical_cpus_available"]))
    gates = {
        "one_binary_per_stage": all(len(v) == 1 for v in stage_sha.values()) if legs_identity else False,
        "stages_differ": stage_sha["baseline"] != stage_sha["candidate"] if legs_identity else False,
        "one_configuration": len(configs) == 1,
        "one_cpu_pin": len(pins) == 1,
        "no_incomplete_selector": not incomplete,
    }
    identity = {
        "baseline_sha256": sorted(stage_sha["baseline"]),
        "candidate_sha256": sorted(stage_sha["candidate"]),
        "configuration": [{"samples": s, "warmup": w} for s, w in sorted(configs)],
        "cpu_pin": [{"cpu_affinity": a, "logical_cpus_available": c} for a, c in sorted(pins, key=str)],
        "git": sorted({(leg[0], i["git_revision"], i["git_worktree_dirty"]) for (leg, _n), i in legs_identity.items()}, key=str),
    }

    complete = [r for r in rows if "statistics" in r]
    result = {
        "schema_version": 1,
        "analysis_kind": "litchi-perf-change-0565-analysis",
        "method": {
            "legs": list(LEGS),
            "leg_roles": {LEGS[0]: "baseline (first)", LEGS[1]: "candidate (first)",
                          LEGS[2]: "candidate (second)", LEGS[3]: "baseline (second)"},
            "statistics": "p50/p95/p99: linear interpolation at fraction*(n-1) over sorted samples; mean: arithmetic mean (same as change-0560/0562 analyzers)",
            "first_direction": f"{LEGS[1]} vs {LEGS[0]}", "second_direction": f"{LEGS[2]} vs {LEGS[3]}",
            "review_trigger": f"adverse in both directions by more than {TRIGGER_PERCENT}%",
        },
        "gates": gates,
        "identity": identity,
        "matrix": json.loads(pathlib.Path(args.matrix).read_text()) if args.matrix else None,
        "deterministic_counters": None,
        "syscalls": None,
        "latency_rows": rows,
        "improve_both_directions_list": improvements,
        "summary": {
            "rows": len(complete),
            "comparisons": len(complete) * len(STATS),
            "improve_both_directions": len(improvements),
            "review_triggers_adverse_both_over_5_percent": len(triggers),
            "incomplete_selectors": incomplete,
        },
        "review_triggers": triggers,
    }

    if args.before_summary and args.after_summary:
        before = json.loads(pathlib.Path(args.before_summary).read_text())
        after = json.loads(pathlib.Path(args.after_summary).read_text())
        counters, syscalls = [], []
        for op_key, b in before["operations"].items():
            a = after["operations"].get(op_key)
            if a is None:
                continue
            bc, ac = b["logical"], a["logical"]
            counters.append({
                "operation": op_key,
                "baseline_read_calls": bc["read_calls"], "candidate_read_calls": ac["read_calls"],
                "read_calls_change_percent": percent(ac["read_calls"], bc["read_calls"]),
                "baseline_read_bytes": bc["read_bytes"], "candidate_read_bytes": ac["read_bytes"],
                "read_bytes_change_percent": percent(ac["read_bytes"], bc["read_bytes"]),
                "baseline_version_calls": bc["version_calls"], "candidate_version_calls": ac["version_calls"],
                "version_calls_change_percent": percent(ac["version_calls"], bc["version_calls"]),
                "baseline_len_calls": bc["len_calls"], "candidate_len_calls": ac["len_calls"],
            })
            bi, ai = b.get("isolation"), a.get("isolation")
            if bi and ai:
                syscalls.append({
                    "operation": op_key,
                    "baseline": {"pread64_per_op": bi["pread64_per_op"], "statx_per_op": bi["statx_per_op"]},
                    "candidate": {"pread64_per_op": ai["pread64_per_op"], "statx_per_op": ai["statx_per_op"]},
                    "pread64_change_percent": percent(ai["pread64_per_op"], bi["pread64_per_op"]) if bi["pread64_per_op"] else None,
                    "statx_change_percent": percent(ai["statx_per_op"], bi["statx_per_op"]) if bi["statx_per_op"] else None,
                    "baseline_four_byte_share_percent": (b.get("trace") or {}).get("four_byte_share_percent"),
                    "candidate_four_byte_share_percent": (a.get("trace") or {}).get("four_byte_share_percent"),
                    "baseline_size_histogram": (b.get("trace") or {}).get("size_histogram"),
                    "candidate_size_histogram": (a.get("trace") or {}).get("size_histogram"),
                })
        result["deterministic_counters"] = counters
        result["syscalls"] = syscalls

    if args.perf_abba_summary:
        strict_dir = pathlib.Path(args.strict_dir or (pathlib.Path(args.output).parent / "strict"))
        strict_dir.mkdir(parents=True, exist_ok=True)
        strict = []
        for name in names:
            out = strict_dir / f"summary-{name}.json"
            argv = [sys.executable, "-B", args.perf_abba_summary,
                    "--a1", str(root / f"{prefix}{LEGS[0]}-{name}.json"), "--b1", str(root / f"{prefix}{LEGS[1]}-{name}.json"),
                    "--b2", str(root / f"{prefix}{LEGS[2]}-{name}.json"), "--a2", str(root / f"{prefix}{LEGS[3]}-{name}.json"),
                    "--json-out", str(out)]
            proc = subprocess.run(argv, capture_output=True, text=True)
            strict.append({"selector": name, "returncode": proc.returncode, "summary": str(out) if out.exists() else None,
                           "stderr": proc.stderr.strip()[-400:]})
        result["strict_cross_check"] = {"tool": args.perf_abba_summary, "results": strict}

    pathlib.Path(args.output).write_text(json.dumps(result, indent=2) + "\n")
    for row in complete:
        p50 = row["statistics"]["p50"]
        flag = "improves" if p50["improves_both"] else ("ADVERSE" if p50["adverse_both"] else "mixed")
        print(f"{row['case']:44s} {row['corpus'][:26]:26s} p50 {p50['first_direction_percent']:+7.2f}% / "
              f"{p50['second_direction_percent']:+7.2f}%  ({p50['baseline_ns']:.0f} -> {p50['candidate_ns']:.0f} ns)  {flag}")
    print(f"\ngates: {gates}")
    print(f"summary: {result['summary']}")
    for t in triggers:
        print(f"  TRIGGER {t['case']} {t['corpus']} {t['statistic']}: {t['first_direction_percent']:+.2f}%/"
              f"{t['second_direction_percent']:+.2f}% ({t['baseline_ns']:.0f} ns -> {t['candidate_ns']:.0f} ns)")
    return 0 if all(gates.values()) else 1


if __name__ == "__main__":
    raise SystemExit(main())
