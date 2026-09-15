import json, statistics, sys, pathlib
SC = pathlib.Path(sys.argv[1])
TAGS = ["poi", "cond", "syn_small", "syn_large"]
LEGS = ["A1", "B1", "B2", "A2", "A3", "A4"]

def load(leg, tag):
    d = json.load(open(SC / "timing" / f"{leg}-{tag}.json"))
    return {c["case"]: c for c in d["cases"]}, d["corpus"]

def stats(samples):
    s = sorted(samples)
    n = len(s)
    return dict(
        p50=s[n // 2] / 1e6,
        mean=statistics.fmean(s) / 1e6,
        p95=s[min(n - 1, int(round(0.95 * (n - 1))))] / 1e6,
        p99=s[min(n - 1, int(round(0.99 * (n - 1))))] / 1e6,
    )

def pooled(legs, tag, case):
    out = []
    for leg in legs:
        out += load(leg, tag)[0][case]["statistics"]["samples_ns"]
    return out

cases = list(load("A1", "poi")[0].keys())
print("# change 0599 paired timing, 40 samples/leg after 3 warmups, taskset -c 17")
print("# before = A1+A2 pooled, after = B1+B2 pooled, A/A control = A3 vs A4")
for tag in TAGS:
    _, corpus = load("A1", tag)
    print(f"\n## fixture {tag}: {corpus['fixture'].split('/')[-1]} "
          f"{corpus['input_bytes']} B, {corpus['worksheet_count']} sheets, "
          f"{corpus['stored_cell_count']} stored cells on the selected sheet, "
          f"{corpus['package_part_count']} parts")
    print(f"{'case':45s} {'A p50':>9s} {'B p50':>9s} {'B-A%':>8s} {'A mean':>9s} {'B mean':>9s} "
          f"{'A p95':>9s} {'B p95':>9s} {'A p99':>9s} {'B p99':>9s} {'AA12%':>7s} {'AA34%':>7s}")
    for case in cases:
        a = stats(pooled(["A1", "A2"], tag, case))
        b = stats(pooled(["B1", "B2"], tag, case))
        a1 = stats(load("A1", tag)[0][case]["statistics"]["samples_ns"])
        a2 = stats(load("A2", tag)[0][case]["statistics"]["samples_ns"])
        a3 = stats(load("A3", tag)[0][case]["statistics"]["samples_ns"])
        a4 = stats(load("A4", tag)[0][case]["statistics"]["samples_ns"])
        aa12 = 100 * (a2["p50"] - a1["p50"]) / a1["p50"]
        aa = 100 * (a4["p50"] - a3["p50"]) / a3["p50"]
        d = 100 * (b["p50"] - a["p50"]) / a["p50"]
        print(f"{case:45s} {a['p50']:9.3f} {b['p50']:9.3f} {d:+8.2f} {a['mean']:9.3f} "
              f"{b['mean']:9.3f} {a['p95']:9.3f} {b['p95']:9.3f} {a['p99']:9.3f} {b['p99']:9.3f} "
              f"{aa12:+7.2f} {aa:+7.2f}")
    # output identity
    ok = True
    for case in cases:
        sh = {load(leg, tag)[0][case].get("output_sha256") for leg in LEGS}
        ok &= len(sh) == 1
    print(f"output sha256 identical across every leg: {ok}")
