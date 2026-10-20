"""Pool the scoped_patch_probe ABBA runs (record 0357a) into summary JSON."""
import json, statistics, sys
rows = []
with open(sys.argv[1]) as handle:
    next(handle)
    for line in handle:
        r, o, build, rss, code, js = line.rstrip("\n").split("\t", 5)
        rows.append((build, int(rss), int(code), json.loads(js)))
out = {"probe": "crates/litchi-ppt/examples/scoped_patch_probe.rs",
       "fixture": "test-data/unified-ops/0294/video.ppt", "patches_per_run": 1000,
       "runs_per_build": sum(1 for b, *_ in rows if b == "before")}
for build in ("before", "after"):
    sel = [(rss, j) for b, rss, c, j in rows if b == build and c == 0]
    med = lambda key: statistics.median(j[key] for _, j in sel)
    out[build] = {
        "max_rss_kb_median": statistics.median(rss for rss, _ in sel),
        "max_rss_kb_range": [min(r for r, _ in sel), max(r for r, _ in sel)],
        "retained_patch_bytes": sel[0][1]["retained_patch_bytes"],
        "retained_per_patch": sel[0][1]["retained_per_patch"],
        "source_len": sel[0][1]["source_len"], "target_len": sel[0][1]["target_len"],
        "commit_median_us": med("commit_median_us"),
        "commit_total_ms_median": med("commit_total_ms"),
        "apply_median_us": med("apply_median_us"),
        "inverse_apply_median_us": med("inverse_apply_median_us"),
    }
b, a = out["before"], out["after"]
out["after_over_before"] = {k: round(a[k] / b[k], 3) for k in (
    "max_rss_kb_median", "retained_patch_bytes", "commit_median_us",
    "commit_total_ms_median", "apply_median_us", "inverse_apply_median_us")}
print(json.dumps(out, indent=2))
