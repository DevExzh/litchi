# Matched resource diagnostic

Normal runs: 100 samples / 10 warmups. Allocator runs: 30 / 3. Each row is one fresh process; allocator elapsed times are excluded. No release latency claim.

## Normal timing and whole-process RSS

| Scenario | Leg | p50 ms | Mean ms | p95 ms | p99 ms | RSS KiB |
|---|---|---:|---:|---:|---:|---:|
| Media-rich | A1 | 710.255527 | 710.289184 | 711.112001 | 711.303500 | 884,904 |
| Media-rich | B1 | 714.916674 | 714.981906 | 716.077350 | 717.612002 | 803,256 |
| Media-rich | B2 | 716.609186 | 716.742979 | 718.294874 | 718.853489 | 803,744 |
| Media-rich | A2 | 727.920450 | 727.713275 | 728.964928 | 729.158231 | 885,900 |
| Plain | A1 | 9.298739 | 9.309594 | 9.398285 | 9.455845 | 82,644 |
| Plain | B1 | 9.455718 | 9.457624 | 9.624733 | 9.681314 | 82,660 |
| Plain | B2 | 9.478239 | 9.501663 | 9.774015 | 9.924835 | 82,736 |
| Plain | A2 | 9.363658 | 9.354683 | 9.404684 | 9.432934 | 82,736 |

## Operation allocation counters

Requested bytes count full realloc requests. Live-after and high-water-after are process snapshots, not operation-local peaks. MB below is decimal; RSS remains KiB.

| Scenario | Leg | Mean allocation calls | Mean requested MB | Live-after MB | High-water-after MB | RSS KiB |
|---|---|---:|---:|---:|---:|---:|
| Media-rich | A1 | 59714.833 | 369.979886 | 287.885879 | 896.603242 | 884,676 |
| Media-rich | B1 | 59714.967 | 369.980306 | 237.452746 | 812.687527 | 803,760 |
| Media-rich | B2 | 59714.833 | 369.979886 | 237.452746 | 812.687527 | 803,508 |
| Media-rich | A2 | 59714.833 | 369.979886 | 287.885879 | 896.603242 | 884,940 |
| Plain | A1 | 49335.000 | 15.664317 | 0.884426 | 3.565760 | 82,740 |
| Plain | B1 | 49335.000 | 15.664317 | 0.796963 | 3.559495 | 82,672 |
| Plain | B2 | 49335.000 | 15.664317 | 0.796963 | 3.559495 | 82,736 |
| Plain | A2 | 49335.000 | 15.664317 | 0.884426 | 3.565760 | 82,608 |

A1/B1 and A2/B2 are the paired comparisons. The raw summary retains both directions, all four same-revision drift checks, and min/max allocation values. Each paired timing/resource direction must be reviewed; no universal speedup is inferred. Shared-host variability and the 100-sample timing scope limit generalization.
