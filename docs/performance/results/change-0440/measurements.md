# Measured allocation result

The change is kept for allocation calls and cumulative requested bytes. Normal elapsed time is not consistently improved. Each main row below contains 30 samples; intervals are deterministic 2,000-resample 95% bootstrap intervals for p50. p95/p99 use nearest rank (p99 is the sample maximum at this sample count).

| Source slides | Phase | p50 ms | p50 95% interval ms | p95 ms | p99 ms |
|---:|---|---:|---|---:|---:|
| 64 | A1 | 2.079 | 2.078–2.081 | 2.093 | 2.094 |
| 4,096 | A1 | 84.444 | 84.315–84.561 | 85.265 | 85.625 |
| 8,192 | A1 | 171.932 | 171.636–172.523 | 173.468 | 173.741 |
| 64 | B1 | 2.076 | 2.069–2.082 | 2.117 | 2.153 |
| 4,096 | B1 | 84.617 | 84.427–84.781 | 85.115 | 85.131 |
| 8,192 | B1 | 169.352 | 169.177–169.636 | 170.408 | 171.054 |
| 8,192 | B2 | 170.960 | 170.684–171.396 | 185.830 | 186.347 |
| 4,096 | B2 | 85.484 | 85.282–85.553 | 85.800 | 85.856 |
| 64 | B2 | 2.074 | 2.072–2.081 | 2.129 | 2.131 |
| 8,192 | A2 | 170.396 | 170.161–170.716 | 171.504 | 172.324 |
| 4,096 | A2 | 84.019 | 83.827–84.115 | 84.342 | 84.495 |
| 64 | A2 | 2.088 | 2.083–2.091 | 2.133 | 2.155 |

All 60 allocator observations per role and shape have identical allocation call and requested-byte values. Peak above entry and retained live delta are unchanged.

| Source slides | Calls before → after | Requested bytes before → after | Peak above entry bytes | Retained live delta bytes |
|---:|---|---|---:|---:|
| 64 | 15,933 → 13,607 | 10,824,868 → 10,702,388 | 812,062 | 73,324 |
| 4,096 | 733,682 → 586,204 | 122,864,214 → 115,097,062 | 19,993,648 | 4,041,249 |
| 8,192 | 1,462,779 → 1,167,845 | 236,704,188 → 221,171,020 | 39,890,548 | 8,074,341 |

Medium/large allocation calls fall 20.101%/20.163%; cumulative requested bytes fall 6.322%/6.562%. Source and commit remain live at the endpoint. No bounded commit-memory or physical-copy claim follows.

Main large R2 p95/p99 flags are +8.353%/+8.138%. The fixed additional confirmation retains all 120 samples: p95 changes −1.410%/−1.525% and p99 −2.613%/−1.868% in the two pairs. The adverse tails did not recur there; the original flags remain part of the decision. No latency or RSS improvement is claimed.

Main before whole-process peak RSS spans 84,647,936–92,905,472 bytes.
Main after whole-process peak RSS spans 84,525,056–86,269,952 bytes.

Whole-process stat instructions change −0.661%, cycles +0.594%, branches −0.993%, and branch misses −1.970%. Profile setup/oracles are included. Zero L1 counters have no cache-miss interpretation. Baseline record conversions retain 15 addr2line warnings each; candidate conversions retain 13 each.

See [summary.json](summary.json) for every main comparison and repeat flag, [supplement.json](supplement.json) for confirmation/profile values, and [decision.json](decision.json) for the scoped decision. The two overlapping initial attempts are retained but excluded by [overlap-exclusion.json](overlap-exclusion.json).

All eight repeat flags remain retained: baseline large normal RSS +9.665%,
baseline large allocator RSS −5.162%, four baseline tiny instrumented timing
statistics about −17.4% to −17.6%, and candidate large normal p95/p99
+9.050%/+8.940%. Allocation counters are identical across repeats. These
variations support neither an instrumented latency nor an RSS benefit claim.
