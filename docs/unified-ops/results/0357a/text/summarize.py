#!/usr/bin/env python3
"""summarize.py ABBA_TSV RANGE_TSV : pooled medians of the record 0357a
text-splice ABBA runs (per-run medians) and the 0627 range-source cases."""
import csv, json, statistics, sys

def rows(path):
    with open(path, newline="") as handle:
        return list(csv.DictReader(handle, delimiter="\t"))

probe = {}
for row in rows(sys.argv[1]):
    if row["exit"] != "0":
        raise SystemExit(f"failed run: {row}")
    key = f'{row["build"]}/{row["mode"]}'
    data = json.loads(row["json"])
    entry = probe.setdefault(key, {"runs": 0, "rss_kb": [], "written_bytes": data["written_bytes"]})
    entry["runs"] += 1
    entry["rss_kb"].append(int(row["max_rss_kb"]))
    for field in ("open_us", "edit_commit_us", "commit_us", "write_to_us", "total_us"):
        entry.setdefault(field, []).append(data[field])
summary = {"probe": {}, "range_source": {}}
for key, entry in sorted(probe.items()):
    out = {"runs": entry["runs"], "written_bytes": entry["written_bytes"],
           "max_rss_kb_median": statistics.median(entry["rss_kb"]),
           "max_rss_kb_range": [min(entry["rss_kb"]), max(entry["rss_kb"])]}
    for field in ("open_us", "edit_commit_us", "commit_us", "write_to_us", "total_us"):
        out[field] = round(statistics.median(entry[field]), 1)
    summary["probe"][key] = out
cases = {}
for row in rows(sys.argv[2]):
    entry = cases.setdefault(row["case"], {}).setdefault(row["build"], {"p50": [], "p95": [], "reads": set(), "bytes": set()})
    entry["p50"].append(int(row["p50_ns"]))
    entry["p95"].append(int(row["p95_ns"]))
    entry["reads"].add(row["read_calls"])
    entry["bytes"].add(row["read_bytes"])
for case, builds in sorted(cases.items()):
    out = {}
    for build, entry in builds.items():
        out[build] = {"runs": len(entry["p50"]), "p50_us": round(statistics.median(entry["p50"]) / 1000, 1),
                      "p95_us": round(statistics.median(entry["p95"]) / 1000, 1),
                      "read_calls": sorted(entry["reads"]), "read_bytes": sorted(entry["bytes"])}
    out["after_over_before_p50"] = round(out["after"]["p50_us"] / out["before"]["p50_us"], 3)
    summary["range_source"][case] = out
print(json.dumps(summary, indent=2))
