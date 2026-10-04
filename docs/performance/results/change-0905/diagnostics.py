"""Descriptive phase, allocator, RSS and schedule summaries; no new gate."""
import hashlib
import json
import statistics as stats
from pathlib import Path

P = Path(__file__).resolve().parent


def desc(p):
    raw = p.read_bytes()
    return {"path": str(p), "bytes": len(raw), "sha256": hashlib.sha256(raw).hexdigest()}


def main():
    a = json.loads((P / "analysis.json").read_text())
    result = {"status": "pass", "analysis": desc(P / "analysis.json"),
              "helper": desc(Path(__file__)), "cases": {}}
    for case in sorted(a["reduction"]["cases"]):
        native = [r for r in a["native"] if r["case_id"] == case]
        alloc = [r for r in a["allocator"] if r["case_id"] == case]
        indexed = {(r["index"], r["role"]): r for r in native}
        row = {"native": {}, "allocator": {}, "phases": {}}
        for metric in ["p50", "p95", "p99"]:
            row["native"][metric] = {
                "role_median_ns": {role: stats.median(r[metric] for r in native if r["role"] == role)
                                   for role in ["baseline", "candidate"]},
                "median_paired_ratio": stats.median(indexed[i, "candidate"][metric] /
                                                     indexed[i, "baseline"][metric] for i in range(6)),
            }
        row["rss_kib"] = {role: [r["resource"]["max_rss_kib"] for r in native if r["role"] == role]
                          for role in ["baseline", "candidate"]}
        for phase in sorted(alloc[0]["allocator_metrics"]):
            roles = {}
            for role in ["baseline", "candidate"]:
                values = [v for r in alloc if r["role"] == role for v in r["allocator_metrics"][phase]]
                assert all(v["status"] == "measured" for v in values)
                roles[role] = {metric: sorted({v[metric] for v in values}) for metric in values[0]
                               if isinstance(values[0][metric], int)}
            row["allocator"][phase] = roles
        for phase in ["open_ns", "plan_ns", "commit_ns", "publication_ns", "reopen_ns"]:
            child_medians = {}
            for r in native:
                raw = json.loads(Path(r["report"]["path"]).read_text())["results"][0]
                values = raw.get("source", {}).get("xlsx_cell_values", {}).get(phase)
                if values is not None:
                    child_medians[r["index"], r["role"]] = stats.median(values)
            if child_medians:
                row["phases"][phase] = {
                    "role_median_ns": {role: stats.median(child_medians[i, role] for i in range(6))
                                       for role in ["baseline", "candidate"]},
                    "median_paired_ratio": stats.median(child_medians[i, "candidate"] /
                                                         child_medians[i, "baseline"] for i in range(6)),
                }
        result["cases"][case] = row
    intervals = {}
    for stage in ["qualification", "native", "allocator"]:
        receipts = [json.loads(Path(r["receipt"]["path"]).read_text()) for r in a[stage]]
        receipts.sort(key=lambda r: r["started_utc"])
        assert all(x["finished_utc"] <= y["started_utc"] for x, y in zip(receipts, receipts[1:]))
        intervals[stage] = {"first_start": receipts[0]["started_utc"],
                            "last_finish": receipts[-1]["finished_utc"]}
    assert intervals["native"]["last_finish"] < intervals["allocator"]["first_start"]
    result["workload_intervals"] = intervals
    result["supervisor_overlap_note"] = (
        "Allocator supervisor/preflight started 2026-10-04T06:01:51.662140+00:00 before the final "
        "native child finished at 06:01:52.158852. The first allocator workload began at "
        "06:01:53.328255; workload children did not overlap. This does not prove host exclusivity.")
    result["scope"] = (
        "Phase child medians and RSS are descriptive. Phase quantiles do not sum to whole-operation "
        "quantiles. RSS is whole-process GNU time high water, affected by executable layout. "
        "Allocator values are separate instrumented runs; commit includes staging and core.")
    with (P / "diagnostics.json").open("x") as output:
        json.dump(result, output, indent=2, sort_keys=True)
        output.write("\n")


if __name__ == "__main__":
    main()
