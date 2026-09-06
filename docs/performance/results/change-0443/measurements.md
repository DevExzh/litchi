# Compact fragment scanner measurements

Frozen acceptance: normal-p50 gate=False; allocation/peak gate=True; practical gate=True.
The matrix retains 24 reports and 720 samples in A1/B1/B2/A2 order.
Each normal row has 30 samples after three warmups. p50 intervals use 2,000
deterministic bootstrap resamples; p95/p99 use nearest rank, making p99 the
sample maximum. These intervals do not eliminate temporal grouping effects.

| Source slides | Phase | p50 ms | p50 95% interval ms | p95 ms | p99 ms |
|---:|---|---:|---|---:|---:|
| 64 | A1 | 1.927 | 1.925–1.932 | 1.977 | 1.996 |
| 4,096 | A1 | 75.258 | 75.190–75.360 | 75.889 | 75.965 |
| 8,192 | A1 | 150.734 | 150.533–150.935 | 151.405 | 151.608 |
| 64 | B1 | 1.932 | 1.929–1.935 | 1.992 | 1.999 |
| 4,096 | B1 | 74.802 | 74.731–74.863 | 75.361 | 75.506 |
| 8,192 | B1 | 151.136 | 150.722–151.456 | 152.252 | 152.446 |
| 8,192 | B2 | 154.084 | 152.959–154.670 | 156.998 | 157.404 |
| 4,096 | B2 | 76.983 | 76.738–77.128 | 78.362 | 78.524 |
| 64 | B2 | 1.930 | 1.926–1.932 | 1.960 | 1.980 |
| 8,192 | A2 | 151.954 | 151.754–152.205 | 153.050 | 153.348 |
| 4,096 | A2 | 75.134 | 75.012–76.171 | 81.725 | 83.024 |
| 64 | A2 | 2.497 | 2.228–2.523 | 2.673 | 2.736 |

Tiny p50 changes R1/R2: +0.213% / -22.723%.
Medium p50 changes R1/R2: -0.606% / +2.460%.
Large p50 changes R1/R2: +0.267% / +1.402%.
The geometric mean of the six explicitly paired p50 before/after ratios is 1.0375×.
Each ratio normalizes one shape and repeat to its own baseline; this is not
a geometric mean across different Office workflows.

All 60 allocator observations per role and shape agree for each counter below.
Source and committed snapshot remain live at the endpoint;
append still materializes the document.

| Source slides | Calls before → after | Requested bytes before → after | Peak above entry before → after | Retained delta before → after |
|---:|---|---|---:|---:|
| 64 | 13,440 → 12,075 | 10,659,766 → 10,613,383 | 781,342 → 781,342 | 73,324 → 73,324 |
| 4,096 | 577,973 → 491,936 | 113,119,080 → 110,226,105 | 18,027,568 → 18,027,568 | 4,041,249 → 4,041,249 |
| 8,192 | 1,151,422 → 979,369 | 217,226,958 → 211,442,207 | 35,958,388 → 35,958,388 | 8,074,341 → 8,074,341 |

Requested bytes are allocation traffic, not physical-copy counts.
No bounded append or general RSS improvement is claimed.

Matched adverse >5% flags: 2. Repeat flags: 8.
Matched allocator medium R1 elapsed_ns.p95 changes +13.310%.
Matched allocator medium R1 elapsed_ns.p99 changes +12.987%.
The before normal tiny elapsed_ns.p50 repeat changes +29.546%.
The before normal tiny elapsed_ns.p95 repeat changes +35.210%.
The before normal tiny elapsed_ns.p99 repeat changes +37.112%.
The before normal tiny elapsed_ns.mean repeat changes +23.544%.
The before normal medium elapsed_ns.p95 repeat changes +7.691%.
The before normal medium elapsed_ns.p99 repeat changes +9.292%.
The after allocator medium elapsed_ns.p95 repeat changes -13.439%.
The after allocator medium elapsed_ns.p99 repeat changes -13.543%.
The large R2 paired p99 comparison is 153.348 → 157.404 ms (+2.645%).
The retained repeat variation and 30-sample tail resolution preclude a general
tail-latency improvement claim. Instrumented timing is not used to establish a normal speedup.
Whole-process before peak RSS spans 84,590,592–84,721,664 bytes.
Whole-process after peak RSS spans 84,578,304–84,721,664 bytes.

Four whole-process profiles include setup, warmups and Rust oracle work.
`cycles:u` changes +1.516%.
`instructions:u` changes -0.739%.
`branches:u` changes -0.803%.
`branch-misses:u` changes -9.036%.
Zero L1 values support no cache-miss interpretation; these are not operation-only causal fractions.
Before symbolization retains 13 report and 13 script addr2line warnings.
After symbolization retains 13 report and 13 script addr2line warnings.

See [summary.json](summary.json), [profile-summary.json](profile-summary.json),
and [decision.json](decision.json). The broader non-iWork goal remains active.
