# Shared source-projection measurements

The original frozen calls/requested-bytes/latency gate failed. The separate
[post-hoc peak-memory review](acceptance-review.md) records the retention
basis. No normal latency, RSS or bounded-memory append improvement is claimed.

Each normal row contains 30 samples after three warmups. p50 intervals use
2,000 deterministic bootstrap resamples; p95/p99 use nearest rank, making
p99 the sample maximum. Full means, intervals and raw values remain in summary.json.

| Source slides | Phase | p50 ms | p50 95% interval ms | p95 ms | p99 ms |
|---:|---|---:|---|---:|---:|
| 64 | A1 | 2.077 | 2.075–2.080 | 2.091 | 2.098 |
| 4,096 | A1 | 83.958 | 83.818–84.049 | 84.543 | 84.993 |
| 8,192 | A1 | 169.397 | 169.205–169.676 | 170.568 | 170.609 |
| 64 | B1 | 2.080 | 2.076–2.084 | 2.102 | 2.105 |
| 4,096 | B1 | 84.516 | 84.430–84.690 | 84.999 | 85.156 |
| 8,192 | B1 | 171.034 | 170.764–171.428 | 172.304 | 172.345 |
| 8,192 | B2 | 170.161 | 169.950–170.352 | 170.754 | 170.897 |
| 4,096 | B2 | 85.514 | 85.394–85.613 | 86.161 | 86.232 |
| 64 | B2 | 2.061 | 2.059–2.067 | 2.099 | 2.129 |
| 8,192 | A2 | 169.039 | 168.827–169.396 | 170.062 | 170.134 |
| 4,096 | A2 | 84.434 | 84.218–84.684 | 85.059 | 85.099 |
| 64 | A2 | 2.104 | 2.089–2.140 | 2.186 | 2.220 |

All 60 allocator observations per role and shape agree for the following
counters. Peak is relative to operation entry; retained delta is the endpoint
minus entry, with source and commit still live.

| Source slides | Calls before → after | Requested bytes before → after | Peak above entry before → after | Retained delta unchanged |
|---:|---|---|---|---:|
| 64 | 13,607 → 13,478 | 10,702,388 → 10,671,668 | 812,062 → 781,342 | 73,324 |
| 4,096 | 586,204 → 578,011 | 115,097,062 → 113,130,982 | 19,993,648 → 18,027,568 | 4,041,249 |
| 8,192 | 1,167,845 → 1,151,460 | 221,171,020 → 217,238,860 | 39,890,548 → 35,958,388 | 8,074,341 |

Medium/large peak falls 9.834%/9.857%. Allocation calls fall only
1.398%/1.403%, and requested bytes 1.708%/1.778%. Normal medium/large p50
is 0.664–1.279% slower. Requested bytes do not measure physical memory copies.

Matched adverse >5% flags: 0. Repeat flags: 1.
The before normal tiny elapsed_ns.p99 repeat changes +5.807%.
Whole-process before peak RSS spans 84,488,192–85,622,784 bytes.
Whole-process after peak RSS spans 84,574,208–84,729,856 bytes.

Four whole-process profiles include setup, warmups and Rust oracle work.
`cycles:u` changes -0.713%.
`instructions:u` changes -0.279%.
`branches:u` changes -0.281%.
`branch-misses:u` changes -6.691%.
Zero L1 values support no cache-miss interpretation.
Before symbolization retains 13 report and 13 script addr2line warnings.
After symbolization retains 11 report and 11 script addr2line warnings.

See [summary.json](summary.json), [profile-summary.json](profile-summary.json),
and [decision.json](decision.json). These observations cover the named owned
append fixtures on this machine and build; the broader non-iWork goal remains open.
