import sys
sys.path.insert(0, "/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0662")
from drive import run, stats
SAMPLES, WARMUP = 80, 15
SHAPES = [(2,64),(2,128),(2,256),(2,512),(2,871),(3,256),(4,128),(4,256),(4,512),(8,64),(8,128),(8,256)]
print("members\tbytes\tremainder\twidth\tn\tp50_ns\tmean_ns\tp95_ns\tp99_ns\tspeedup_vs_seq")
for members, size in SHAPES:
    seq1,_ = run(["zipsweep", members, size, 0, WARMUP, SAMPLES], SAMPLES)
    legs = {}
    for width in (1,2,4):
        legs[width],_ = run(["zipsweep", members, size, width, WARMUP, SAMPLES], SAMPLES)
    seq2,_ = run(["zipsweep", members, size, 0, WARMUP, SAMPLES], SAMPLES)
    legs[0] = seq1+seq2
    base = stats(legs[0])["p50"]
    for width in (0,1,2,4):
        r = stats(legs[width])
        print(f"{members}\t{size}\t{(members-1)*size}\t{width}\t{r['n']}\t{r['p50']}\t{r['mean']}\t{r['p95']}\t{r['p99']}\t{base/r['p50']:.3f}")
    a,b = stats(seq1)["p50"], stats(seq2)["p50"]
    print(f"# aa-floor members={members} bytes={size} a1={a} a2={b} spread={abs(a-b)/min(a,b)*100:.2f}%", flush=True)
