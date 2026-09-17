import sys
sys.path.insert(0, "/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0662")
from drive import run, stats
SAMPLES, WARMUP = 40, 8
SHAPES = [(8,1024),(8,4096),(16,4096),(8,16384),(16,16384),(40,16384),(8,65536),(16,65536),(8,262144)]
print("members\tbytes\tremainder\tleg\tn\tp50_ns\tmean_ns\tp95_ns\tp99_ns\tspeedup_vs_w1")
for members, size in SHAPES:
    a1,_ = run(["sweep", members, size, 1, WARMUP, SAMPLES], SAMPLES)
    legs = {}
    for w in (2,4,8):
        legs[f"w{w}"],_ = run(["sweep", members, size, w, WARMUP, SAMPLES], SAMPLES)
    for w in (4,8):
        legs[f"w{w}+pool"],_ = run(["sweep", members, size, w, WARMUP, SAMPLES, "pool"], SAMPLES)
    a2,_ = run(["sweep", members, size, 1, WARMUP, SAMPLES], SAMPLES)
    legs["w1"] = a1+a2
    base = stats(legs["w1"])["p50"]
    rem = (members-1)*size
    for leg in ("w1","w2","w4","w8","w4+pool","w8+pool"):
        r = stats(legs[leg])
        print(f"{members}\t{size}\t{rem}\t{leg}\t{r['n']}\t{r['p50']}\t{r['mean']}\t{r['p95']}\t{r['p99']}\t{base/r['p50']:.3f}")
    x,y = stats(a1)["p50"], stats(a2)["p50"]
    print(f"# aa-floor members={members} bytes={size} a1={x} a2={y} spread={abs(x-y)/min(x,y)*100:.2f}%", flush=True)
