"""Validate completed baseline repeat-one evidence; no comparison or adoption claim."""
import argparse
import json
from pathlib import Path

import analyze_guards as G
import analyze_metrics as M
import run as R


def validate():
    plan = M.plan_data()
    M.frozen_inputs()
    # The public analyzer initializes its revision binding before its pending check.
    G.analyze()
    M.validate_stage_metadata("baseline", plan)
    _, manifest_sha = G._manifest("baseline")
    receipts = {}
    counts = {}
    for lane in ("preflight", "native", "alloc"):
        kind = M.ALLOC_BINARY if lane == "alloc" else M.NORMAL_BINARY
        binary = M.binary_metadata("baseline", kind, plan)
        count = 0
        for job in M.jobs_for(plan, "baseline", lane):
            if job["repeat"] != 1:
                continue
            _, report, catalog = M.check_receipt(job, binary, plan)
            M.validate_report(report, catalog,
                              R.HERE / "baseline" / (job["name"] + ".stderr"),
                              job, kind, binary, plan)
            name = job["name"] + ".receipt.json"
            receipts[name] = R.sha(R.HERE / "baseline" / name)
            count += 1
        assert count == 16
        counts[lane] = count
    for kind, lane in (("guard-normal", "normal"), ("guard-alloc", "alloc"), ("cap", None)):
        binary = G._binary_descriptor("baseline", kind, manifest_sha)
        G._build_receipt("baseline", kind, manifest_sha)
        jobs = G._cap_jobs() if lane is None else G._guard_jobs(lane)
        count = 0
        for job in jobs:
            if job["repeat"] != 1:
                continue
            if lane is None:
                G._cap_capture("baseline", job, binary, manifest_sha, "")
            else:
                G._guard_capture("baseline", lane, job, binary, manifest_sha, "")
            name = job["name"] + ".receipt.json"
            receipts[name] = R.sha(R.HERE / "baseline" / name)
            count += 1
        assert count == (5 if lane is None else 6)
        counts[kind] = count
    for kind in ("normal", "alloc", "guard-normal", "guard-alloc", "cap"):
        name = "build-" + kind + ".receipt.json"
        receipts[name] = R.sha(R.HERE / "baseline" / name)
    actual = {p.name for p in (R.HERE / "baseline").glob("*.receipt.json")
              if p.name.startswith("build-") or "-r1-" in p.name}
    assert actual == set(receipts) and len(receipts) == 70
    intervals = []
    for name in receipts:
        value = json.loads((R.HERE / "baseline" / name).read_text())
        intervals.append((G.timestamp(value["start_utc"], name),
                          G.timestamp(value["end_utc"], name), name))
    G._check_order(intervals, "baseline repeat-one builds and captures")
    return {
        "schema": "xlsx_0553_baseline_r1_validation_v1", "status": "pass",
        "scope": "Baseline repeat-one job, source, binary, build, receipt, artifact, corpus, report, oracle and serial-interval validation only; no candidate comparison or admission claim.",
        "plan_sha256": R.sha(R.HERE / "plan.json"),
        "source_manifest_sha256": manifest_sha,
        "scripts": {p.name: R.sha(p) for p in (Path(__file__), R.HERE / "analyze_metrics.py", R.HERE / "analyze_guards.py")},
        "counts": counts, "receipts": receipts,
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    result = validate()
    if args.output:
        R.write(args.output, result)
    print("PASS: 70 baseline repeat-one receipts validated; no admission claim")
