# Shared staging traversal measurements

The candidate meets the frozen normal-p50 gate in both medium/large repeats.
The matrix retains 24 reports and 720 samples in A1/B1/B2/A2 order.
Each normal row has 30 samples after three warmups. p50 intervals use 2,000
deterministic bootstrap resamples; p95/p99 use nearest rank, making p99 the
sample maximum. These intervals do not eliminate temporal grouping effects.

| Source slides | Phase | p50 ms | p50 95% interval ms | p95 ms | p99 ms |
|---:|---|---:|---|---:|---:|
| 64 | A1 | 2.084 | 2.078–2.113 | 2.142 | 2.151 |
| 4,096 | A1 | 86.552 | 86.358–86.796 | 87.305 | 87.660 |
| 8,192 | A1 | 175.493 | 175.139–175.846 | 176.555 | 176.627 |
| 64 | B1 | 1.938 | 1.933–1.942 | 1.994 | 2.017 |
| 4,096 | B1 | 75.394 | 75.224–75.498 | 75.908 | 76.173 |
| 8,192 | B1 | 152.141 | 151.191–153.670 | 155.312 | 157.059 |
| 8,192 | B2 | 152.777 | 152.491–153.067 | 162.287 | 168.196 |
| 4,096 | B2 | 76.112 | 75.876–76.189 | 77.543 | 78.282 |
| 64 | B2 | 1.930 | 1.921–1.953 | 1.974 | 1.979 |
| 8,192 | A2 | 170.168 | 169.737–170.869 | 172.011 | 172.099 |
| 4,096 | A2 | 84.409 | 84.148–84.479 | 84.899 | 85.334 |
| 64 | A2 | 2.073 | 2.066–2.076 | 2.116 | 2.118 |

Tiny p50 changes R1/R2: -7.020% / -6.923%.
Medium p50 changes R1/R2: -12.891% / -9.830%.
Large p50 changes R1/R2: -13.307% / -10.220%.
The geometric mean of the six explicitly paired p50 before/after ratios is 1.1119×.
Each ratio normalizes one shape and repeat to its own baseline; this is not
a geometric mean across different Office workflows.

All 60 allocator observations per role and shape agree for each counter below.
Peak above entry and retained live delta are unchanged. Source and committed
snapshot remain live at the endpoint; append still materializes the document.

| Source slides | Calls before → after | Requested bytes before → after | Peak above entry unchanged | Retained delta unchanged |
|---:|---|---|---:|---:|
| 64 | 13,478 → 13,440 | 10,671,668 → 10,659,766 | 781,342 | 73,324 |
| 4,096 | 578,011 → 577,973 | 113,130,982 → 113,119,080 | 18,027,568 | 4,041,249 |
| 8,192 | 1,151,460 → 1,151,422 | 217,238,860 → 217,226,958 | 35,958,388 | 8,074,341 |

Each size saves only 38 allocation calls and 11,902 requested bytes. The
practical allocation/peak gate is not met; retention rests on normal latency.
No material allocation, peak/RSS, physical-copy, or bounded append gain is claimed.

Matched adverse >5% flags: 0. Repeat flags: 5.
The before allocator medium elapsed_ns.p50 repeat changes -5.817%.
The before allocator medium elapsed_ns.p95 repeat changes -6.447%.
The before allocator medium elapsed_ns.p99 repeat changes -6.270%.
The before allocator medium elapsed_ns.mean repeat changes -5.854%.
The after normal large elapsed_ns.p99 repeat changes +7.091%.
The large R2 paired p99 comparison is 172.099 → 168.196 ms (-2.268%).
The retained repeat variation and 30-sample tail resolution preclude a general
tail-latency improvement claim. Instrumented timing drift is not used for the speedup claim.
Whole-process before peak RSS spans 84,586,496–84,721,664 bytes.
Whole-process after peak RSS spans 84,508,672–84,721,664 bytes.

Four whole-process profiles include setup, warmups and Rust oracle work.
`cycles:u` changes -10.092%.
`instructions:u` changes -8.824%.
`branches:u` changes -7.863%.
`branch-misses:u` changes -12.188%.
Zero L1 values support no cache-miss interpretation; these are not operation-only causal fractions.
Before symbolization retains 13 report and 13 script addr2line warnings.
After symbolization retains 14 report and 14 script addr2line warnings.

See [summary.json](summary.json), [profile-summary.json](profile-summary.json),
and [decision.json](decision.json). The broader non-iWork goal remains active.
