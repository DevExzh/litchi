#!/usr/bin/env python3
"""Compare the exact shared scalar corpus, retaining every regression trigger."""
import argparse
import csv
import hashlib
import importlib.util
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
PARITY = (
    "repeat", "expected_success", "failure", "successes_p50", "successes_max",
    "refusals_p50", "refusals_max", "checksum_p50", "checksum_max",
    "output_reserved_bytes_p50", "output_reserved_bytes_max",
)
METRICS = (
    "p50_ns", "p95_ns", "p99_ns", "max_rss_kib", "alloc_calls_p50",
    "requested_bytes_p50", "peak_live_delta_p50", "live_after_p50",
)


def corpus():
    path = HERE / "scalar-harness/run.py"
    spec = importlib.util.spec_from_file_location("scalar_harness", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return {(phase, case): repeat for phase in module.PHASES
            for case, repeat in module.COMPARABLE_CASES}


def read(path, expected):
    result = {}
    with path.open(newline="") as stream:
        for row in csv.DictReader(stream):
            key = row["phase"], row["case"]
            if key in result:
                raise ValueError(f"duplicate row in {path}: {key}")
            if row["status"] != "0":
                raise ValueError(f"unsuccessful capture in {path}: {key}")
            if key not in expected or int(row["repeat"]) != expected[key]:
                raise ValueError(f"unexpected case or repeat in {path}: {key}")
            result[key] = row
    if result.keys() != expected.keys():
        raise ValueError(f"incomplete corpus in {path}: {expected.keys() - result.keys()}")
    return result


def compare(baseline, candidate):
    rows = []
    for key in sorted(baseline):
        old, new = baseline[key], candidate[key]
        parity = {field: {"baseline": old[field], "candidate": new[field]}
                  for field in PARITY if old[field] != new[field]}
        changes = {}
        flags = []
        for metric in METRICS:
            before, after = float(old[metric]), float(new[metric])
            delta = (after / before - 1) * 100 if before else None
            changes[metric] = {"baseline": before, "candidate": after,
                               "delta_pct": delta}
            if metric in METRICS[:4] and (after > before * 1.05):
                flags.append(metric)
        rows.append({"phase": key[0], "case": key[1], "changes": changes,
                     "parity_mismatches": parity, "review_triggers": flags})
    return rows


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("baseline", type=Path)
    parser.add_argument("candidate", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    expected = corpus()
    rows = compare(read(args.baseline, expected), read(args.candidate, expected))
    mismatches = sum(bool(row["parity_mismatches"]) for row in rows)
    report = {"baseline_sha256": hashlib.sha256(args.baseline.read_bytes()).hexdigest(),
              "candidate_sha256": hashlib.sha256(args.candidate.read_bytes()).hexdigest(),
              "row_count": len(rows), "parity_mismatch_rows": mismatches,
              "review_trigger_rows": sum(bool(row["review_triggers"]) for row in rows),
              "threshold_pct": 5, "rows": rows}
    args.output.write_text(json.dumps(report, indent=2, allow_nan=False) + "\n")
    print(json.dumps({key: value for key, value in report.items() if key != "rows"}))
    raise SystemExit(1 if mismatches else 0)


if __name__ == "__main__":
    main()
