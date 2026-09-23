import json, os, statistics, subprocess, sys
ROOT = "/home/zhuhe/code/litchi-worktrees/scratch/0745"
T = "/home/zhuhe/code/litchi-worktrees/0745-ppt-lazy-artifact-digests/test-data/poi/test-data/slideshow"
mode, fixture, step = sys.argv[1], sys.argv[2], int(sys.argv[3])
results = []
for pad in range(0, 4096, step):
    row = []
    for arm in ("AB" if (pad // step) % 2 == 0 else "BA"):
        env = dict(os.environ)
        env["P"] = "x" * pad
        out = subprocess.run(["taskset", "-c", "16", f"{ROOT}/bin/{arm}/probe", "--mode", mode, "--input", f"{T}/{fixture}", "--warmups", "5", "--samples", "30"], capture_output=True, text=True, env=env, check=True).stdout
        row.append((arm, statistics.median(json.loads(out)["elapsed_ns"]) / 1000))
    d = dict(row)
    results.append((pad, d["A"], d["B"]))
changes = [100 * (b / a - 1) for _, a, b in results]
print(mode, fixture, "n", len(results), "median change %+.1f%%" % statistics.median(changes),
      "min %+.1f%% max %+.1f%%" % (min(changes), max(changes)),
      "A range %.0f-%.0f" % (min(a for _, a, _ in results), max(a for _, a, _ in results)),
      "B range %.0f-%.0f" % (min(b for _, _, b in results), max(b for _, _, b in results)))
hist = {}
for c in changes:
    key = int((c // 5) * 5)
    hist[key] = hist.get(key, 0) + 1
print("  histogram(5% bins):", dict(sorted(hist.items())))
json.dump(results, open(f"{ROOT}/prof/padscan3-{mode}-{fixture}.json", "w"))
