#!/usr/bin/env python3
import sys
sys.path.insert(0, "/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0662")
from drive import run, stats

SAMPLES = 60
WARMUP = 10
SHAPES = [(2, 871), (4, 871), (8, 871), (16, 871), (40, 871), (128, 871),
          (2, 4096), (4, 4096), (8, 4096), (16, 4096), (40, 4096),
          (2, 26698), (4, 26698), (8, 26698), (16, 26698),
          (2, 1893450)]
WIDTHS = [0, 1, 2, 4, 8]

print("members\tbytes\tremainder\twidth\tn\tp50_ns\tmean_ns\tp95_ns\tp99_ns\tspeedup_vs_seq")
for members, size in SHAPES:
    legs = {}
    # A1 (sequential) ... widths ... A2 (sequential), so the floor is measurable.
    seq1, note = run(["zipsweep", members, size, 0, WARMUP, SAMPLES], SAMPLES)
    for width in WIDTHS[1:]:
        legs[width], note = run(["zipsweep", members, size, width, WARMUP, SAMPLES], SAMPLES)
    seq2, _ = run(["zipsweep", members, size, 0, WARMUP, SAMPLES], SAMPLES)
    legs[0] = seq1 + seq2
    remainder = (members - 1) * size
    base = stats(legs[0])["p50"]
    aa = stats(seq1)["p50"], stats(seq2)["p50"]
    for width in WIDTHS:
        row = stats(legs[width])
        print(f"{members}\t{size}\t{remainder}\t{width}\t{row['n']}\t{row['p50']}\t{row['mean']}\t{row['p95']}\t{row['p99']}\t{base/row['p50']:.3f}")
    print(f"# aa-floor members={members} bytes={size} seq_p50_a1={aa[0]} seq_p50_a2={aa[1]} spread={abs(aa[0]-aa[1])/min(aa)*100:.2f}%", flush=True)
