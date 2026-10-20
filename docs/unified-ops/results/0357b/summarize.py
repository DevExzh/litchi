"""Pool the scoped_patch_probe ABC-CBA runs (record 0357b) into summary JSON."""
import json, statistics, sys

BUILDS = ("base", "a", "b")
KEYS = ("commit_median_us", "commit_total_ms", "apply_median_us",
        "inverse_apply_median_us", "first_commit_median_us")
rows = []
with open(sys.argv[1]) as handle:
    next(handle)
    for line in handle:
        r, o, build, rss, code, js = line.rstrip("\n").split("\t", 5)
        rows.append((build, int(rss), int(code), json.loads(js)))
out = {"probe": "crates/litchi-ppt/examples/scoped_patch_probe.rs",
       "fixture": "test-data/unified-ops/0294/video.ppt", "patches_per_run": 1000,
       "first_commits_per_run": 200,
       "builds": {"base": "e673d4addb (pre-0357a, whole-artifact patch)",
                  "a": "7d22d1a18b (0357a tip, one SHA-256 of the target per commit)",
                  "b": "record 0357b (incremental block roots)"},
       "runs_per_build": sum(1 for b, *_ in rows if b == "base")}
for build in BUILDS:
    sel = [(rss, j) for b, rss, c, j in rows if b == build and c == 0]
    entry = {
        "max_rss_kb_median": statistics.median(rss for rss, _ in sel),
        "max_rss_kb_range": [min(r for r, _ in sel), max(r for r, _ in sel)],
        "retained_patch_bytes": sel[0][1]["retained_patch_bytes"],
        "retained_per_patch": sel[0][1]["retained_per_patch"],
        "source_len": sel[0][1]["source_len"], "target_len": sel[0][1]["target_len"],
    }
    for key in KEYS:
        values = [j[key] for _, j in sel]
        name = "commit_total_ms_median" if key == "commit_total_ms" else key
        entry[name] = round(statistics.median(values), 3)
        entry[key + "_range"] = [min(values), max(values)]
    out[build] = entry
ratio_keys = ("max_rss_kb_median", "retained_patch_bytes", "commit_median_us",
              "commit_total_ms_median", "apply_median_us", "inverse_apply_median_us",
              "first_commit_median_us")
for num, den in (("b", "base"), ("b", "a"), ("a", "base")):
    out[f"{num}_over_{den}"] = {k: round(out[num][k] / out[den][k], 3) for k in ratio_keys}
print(json.dumps(out, indent=2))
