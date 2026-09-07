# Fixed allocator-policy diagnostic

Controlled whole-process perf stat diagnostic with fixed glibc mmap thresholds; not ordinary timing lanes or API-attributed counters. Original ordinary bytes/media timings remain adverse. Page-fault association is supported; exact historical allocation/map decisions were not traced.

| Threshold | Pair | API p50 change | Instructions change | Cycles change | Page faults A / B |
|---:|---|---:|---:|---:|---:|
| 131072 | 0 / 1 | -9.999% | -1.436% | -2.576% | 1,779,302 / 1,631,486 |
| 131072 | 3 / 2 | -10.266% | -1.390% | -2.573% | 1,779,241 / 1,631,447 |
| 33554432 | 4 / 5 | -2.626% | -3.117% | -4.708% | 712,305 / 442,976 |
| 33554432 | 7 / 6 | -2.989% | -0.059% | -1.042% | 443,008 / 443,201 |

| Ordinary process | Build | API p50 ms | Minor faults |
|---|---|---:|---:|
| runs/1 | baseline | 25.071 | 383,775 |
| runs/5 | candidate | 28.935 | 922,403 |
| runs/10 | candidate | 29.145 | 835,710 |
| runs/14 | baseline | 25.050 | 385,304 |
