# Matched resource diagnostic

Normal runs: 100 samples / 10 warmups. Allocator runs: 30 / 3. Each row is one fresh process; allocator elapsed times are excluded. No release latency claim.

## Normal timing and whole-process RSS

| Scenario | Leg | p50 ms | Mean ms | p95 ms | p99 ms | RSS KiB |
|---|---|---:|---:|---:|---:|---:|
| Media-rich | A1 | 726.948498 | 727.991433 | 733.523053 | 735.786527 | 885,784 |
| Media-rich | B1 | 715.350345 | 715.520746 | 717.578847 | 718.316162 | 885,152 |
| Media-rich | B2 | 714.987846 | 715.171453 | 717.357228 | 718.074731 | 885,168 |
| Media-rich | A2 | 727.165036 | 728.073434 | 733.484814 | 734.479687 | 885,944 |
| Plain | A1 | 9.350672 | 9.393813 | 9.623393 | 9.792994 | 82,724 |
| Plain | B1 | 9.544329 | 9.558371 | 9.800164 | 9.941464 | 82,620 |
| Plain | B2 | 9.471693 | 9.454236 | 9.636263 | 9.691884 | 82,616 |
| Plain | A2 | 9.435808 | 9.473308 | 9.746165 | 9.863545 | 82,672 |

## Operation allocation counters

Requested bytes count full realloc requests. Live-after and high-water-after are process snapshots, not operation-local peaks. MB below is decimal; RSS remains KiB.

| Scenario | Leg | Mean allocation calls | Mean requested MB | Live-after MB | High-water-after MB | RSS KiB |
|---|---|---:|---:|---:|---:|---:|
| Media-rich | A1 | 61744.867 | 36811.874713 | 287.885879 | 896.603242 | 885,120 |
| Media-rich | B1 | 59714.700 | 369.979465 | 287.885879 | 896.603242 | 884,912 |
| Media-rich | B2 | 59715.000 | 369.980411 | 287.885879 | 896.604810 | 886,136 |
| Media-rich | A2 | 61744.700 | 36811.874187 | 287.885879 | 896.603242 | 886,444 |
| Plain | A1 | 49681.000 | 24.558673 | 0.884426 | 3.565760 | 82,720 |
| Plain | B1 | 49335.000 | 15.664317 | 0.884426 | 3.565760 | 82,588 |
| Plain | B2 | 49335.000 | 15.664317 | 0.884426 | 3.565760 | 82,608 |
| Plain | A2 | 49681.000 | 24.558673 | 0.884426 | 3.565760 | 82,668 |

A1/B1 and A2/B2 are the paired comparisons. The raw summary retains both directions, all four same-revision drift checks, and min/max allocation values. Plain median timing is adverse in both pairs; it must not be described as a universal speedup. Shared-host variability and the 100-sample timing scope limit generalization.
