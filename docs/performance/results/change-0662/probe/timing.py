import sys, os
sys.path.insert(0, "/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0662")
from drive import run, stats
SAMPLES, WARMUP = 40, 8
fixtures = [l.strip() for l in open(sys.argv[1]) if l.strip()]
print("fixture\tleg\tn\tp50_ns\tmean_ns\tp95_ns\tp99_ns\tspeedup_vs_w1")
for f in fixtures:
    name = os.path.basename(f)
    legs = {}
    a1,_ = run(["time", f, 64, 1, WARMUP, SAMPLES], SAMPLES)
    for w in (2,4,8):
        legs[f"w{w}"],_ = run(["time", f, 64, w, WARMUP, SAMPLES], SAMPLES)
    for w in (4,8):
        legs[f"w{w}+pool"],_ = run(["time", f, 64, w, WARMUP, SAMPLES, "pool"], SAMPLES)
    a2,_ = run(["time", f, 64, 1, WARMUP, SAMPLES], SAMPLES)
    legs["w1"] = a1 + a2
    opn,_ = run(["open", f, WARMUP, SAMPLES], SAMPLES)
    legs["open-only"] = opn
    base = stats(legs["w1"])["p50"]
    for leg in ("w1","w2","w4","w8","w4+pool","w8+pool","open-only"):
        r = stats(legs[leg])
        print(f"{name}\t{leg}\t{r['n']}\t{r['p50']}\t{r['mean']}\t{r['p95']}\t{r['p99']}\t{base/r['p50']:.3f}")
    x,y = stats(a1)["p50"], stats(a2)["p50"]
    print(f"# aa-floor {name} a1={x} a2={y} spread={abs(x-y)/min(x,y)*100:.2f}%", flush=True)
