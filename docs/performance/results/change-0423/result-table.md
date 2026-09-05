# 0423 same-revision lifecycle baseline

Normal rows retain timing statistics; allocator rows retain resource vectors and whole-process RSS. No owned/source or normal/allocator elapsed comparison is made.

| Lane | Corpus | Role | Repeat | Samples | p50 ns | Mean ns | p95 ns | p99 ns | RSS KiB | Alloc calls mean | Allocated bytes mean | Region peak mean |
|---|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| normal | plain | owned | R1 | 100 | 9468445 | 9479308.400 | 9692766 | 9805526 | 82,728 | — | — | — |
| normal | plain | source | R1 | 100 | 2164658 | 2166424.000 | 2180128 | 2186088 | 82,620 | — | — | — |
| normal | plain | source | R2 | 100 | 2164877 | 2166287.390 | 2182768 | 2192327 | 82,704 | — | — | — |
| normal | plain | owned | R2 | 100 | 9579184 | 9569306.100 | 9865365 | 9899455 | 82,680 | — | — | — |
| normal | media_rich | owned | R1 | 100 | 718680008 | 718690273.210 | 719358193 | 719579595 | 803,504 | — | — | — |
| normal | media_rich | source | R1 | 100 | 272762485 | 272760130.850 | 273036569 | 273140559 | 804,972 | — | — | — |
| normal | media_rich | source | R2 | 100 | 261778952 | 261831188.250 | 262303934 | 262506945 | 803,684 | — | — | — |
| normal | media_rich | owned | R2 | 100 | 769581184 | 769243195.420 | 771844804 | 772228396 | 804,948 | — | — | — |
| allocator | plain | owned | R1 | 30 | — | — | — | — | 82,736 | 49335.000 | 15664317.000 | 1359983.000 |
| allocator | plain | source | R1 | 30 | — | — | — | — | 82,736 | 10993.000 | 8916495.000 | 1059110.000 |
| allocator | plain | source | R2 | 30 | — | — | — | — | 82,704 | 10993.000 | 8916495.000 | 1059110.000 |
| allocator | plain | owned | R2 | 30 | — | — | — | — | 82,736 | 49335.000 | 15664317.000 | 1359983.000 |
| allocator | media_rich | owned | R1 | 30 | — | — | — | — | 805,020 | 59714.767 | 369979675.533 | 272736283.000 |
| allocator | media_rich | source | R1 | 30 | — | — | — | — | 803,708 | 13314.000 | 117608609.000 | 229300257.000 |
| allocator | media_rich | source | R2 | 30 | — | — | — | — | 804,980 | 13314.000 | 117608609.000 | 229300257.000 |
| allocator | media_rich | owned | R2 | 30 | — | — | — | — | 803,976 | 59715.233 | 369981146.467 | 272736283.000 |

## Within-role repeat drift

| Corpus | Role | p50 | Mean | p95 | p99 | Acceptance-grade all fields |
|---|---|---:|---:|---:|---:|---|
| media_rich | owned | +7.083% | +7.034% | +7.296% | +7.317% | False |
| media_rich | source | -4.027% | -4.007% | -3.931% | -3.893% | True |
| plain | owned | +1.170% | +0.949% | +1.781% | +0.958% | True |
| plain | source | +0.010% | -0.006% | +0.121% | +0.285% | True |

All source/output and semantic/raw-preservation gates are retained per role. Allocator elapsed values are intentionally withheld from comparisons; allocator vectors are process callback-order observations, and RSS is whole-process GNU `time -v` evidence.
