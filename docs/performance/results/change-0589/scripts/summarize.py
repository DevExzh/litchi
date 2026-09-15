import json, pathlib, statistics, sys
base = pathlib.Path("/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0589/out")
def load(name):
    return [json.loads(l) for l in (base / f"perfstat-{name}.jsonl").read_text().splitlines()]
b, a, aa = load("before"), load("after"), load("beforeAA")
paths = [l for l in (base/"doc-fixtures.txt").read_text().split()] + [l for l in (base/"ppt-fixtures.txt").read_text().split()]
assert len(b) == len(a) == len(aa) == len(paths), (len(b), len(a), len(aa), len(paths))
print(f"{'mode':18s} {'fixture':52s} {'bytes':>9s} {'before cyc':>12s} {'after cyc':>12s} {'d%':>7s} {'AA%':>6s} {'before ins':>12s} {'after ins':>12s} {'d%':>7s}")
cd, idl, aad = [], [], []
rows = sorted(zip(b, a, aa, paths), key=lambda t: (t[0]["mode"], -pathlib.Path(t[3]).stat().st_size))
for bd, ad, ad2, path in rows:
    label = str(pathlib.Path(path)).replace("/home/zhuhe/code/litchi/test-data/", "")
    size = pathlib.Path(path).stat().st_size
    dc = (ad["cycles_per_op"] - bd["cycles_per_op"]) / bd["cycles_per_op"] * 100
    di = (ad["instructions_per_op"] - bd["instructions_per_op"]) / bd["instructions_per_op"] * 100
    daa = (ad2["cycles_per_op"] - bd["cycles_per_op"]) / bd["cycles_per_op"] * 100
    cd.append(dc); idl.append(di); aad.append(daa)
    print(f"{bd['mode']:18s} {label[-52:]:52s} {size:9d} {bd['cycles_per_op']:12,.0f} {ad['cycles_per_op']:12,.0f} {dc:+7.1f} {daa:+6.1f} {bd['instructions_per_op']:12,.0f} {ad['instructions_per_op']:12,.0f} {di:+7.1f}")
print()
for name, vals in (("cycles", cd), ("instructions", idl), ("A/A cycles (floor)", aad)):
    print(f"{name:22s} median {statistics.median(vals):+6.1f}%  min {min(vals):+6.1f}%  max {max(vals):+6.1f}%  n={len(vals)}")
