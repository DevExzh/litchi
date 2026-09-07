# Fixed allocator-policy diagnostic

Controlled whole-process perf stat diagnostic with fixed glibc mmap thresholds; not ordinary timing lanes or API-attributed counters. Original uncontrolled counters remain adverse. Page-fault association is supported; exact historical allocation/map decisions were not traced.

| Threshold | Pair | API p50 change | Instructions change | Cycles change | Page faults A / B |
|---:|---|---:|---:|---:|---:|
| 131072 | 0 / 1 | -0.281% | -0.057% | -0.171% | 1,779,322 / 1,779,244 |
| 131072 | 3 / 2 | -0.454% | -0.097% | +0.057% | 1,779,232 / 1,779,234 |
| 33554432 | 4 / 5 | -0.714% | -3.151% | -4.028% | 713,336 / 443,020 |
| 33554432 | 7 / 6 | -0.488% | -0.020% | -0.054% | 443,897 / 443,222 |

| Ordinary process | Build | API p50 ms | Minor faults |
|---|---|---:|---:|
| runs/1 | baseline | 25.409 | 383,915 |
| runs/5 | candidate | 25.074 | 385,813 |
| runs/10 | candidate | 25.074 | 384,282 |
| runs/14 | baseline | 30.510 | 942,424 |
| confirmation/runs/0 | baseline | 25.219 | 381,894 |
| confirmation/runs/1 | candidate | 25.157 | 381,238 |
| confirmation/runs/2 | candidate | 25.048 | 384,290 |
| confirmation/runs/3 | baseline | 25.235 | 381,380 |
| confirmation/runs/4 | baseline | 25.199 | 384,415 |
| confirmation/runs/5 | candidate | 25.038 | 385,300 |
| confirmation/runs/6 | candidate | 25.089 | 384,289 |
| confirmation/runs/7 | baseline | 34.827 | 1,221,860 |
