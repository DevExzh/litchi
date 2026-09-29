# 0843 retained diagnostic tables

Native intervals are bootstrap intervals of six block values; time is ns, RSS KiB, faults counts.

| Parts | Mode | Metric | Median | CI low | CI high |
|---:|---|---|---:|---:|---:|
| 4 | publish | p50_ns | 25080.500 | 24960.000 | 25275.000 |
| 4 | publish | p95_ns | 29990.000 | 28010.000 | 36190.000 |
| 4 | publish | p99_ns | 39155.500 | 35800.000 | 51136.000 |
| 4 | publish | mean_ns | 25865.583 | 25700.450 | 26493.267 |
| 4 | publish | rss_kib | 4426.000 | 4348.000 | 4432.000 |
| 4 | publish | minor_faults | 359.000 | 356.000 | 360.000 |
| 4 | publish | major_faults | 0.000 | 0.000 | 0.000 |
| 4 | lifecycle | p50_ns | 39390.000 | 39020.000 | 39940.000 |
| 4 | lifecycle | p95_ns | 50285.000 | 46610.000 | 53220.000 |
| 4 | lifecycle | p99_ns | 52890.000 | 50885.000 | 74380.500 |
| 4 | lifecycle | mean_ns | 40570.667 | 39941.650 | 41666.467 |
| 4 | lifecycle | rss_kib | 4370.000 | 4324.000 | 4534.000 |
| 4 | lifecycle | minor_faults | 359.500 | 357.000 | 361.000 |
| 4 | lifecycle | major_faults | 0.000 | 0.000 | 0.000 |
| 64 | publish | p50_ns | 77120.500 | 76780.500 | 77730.000 |
| 64 | publish | p95_ns | 87730.000 | 86525.500 | 88335.500 |
| 64 | publish | p99_ns | 102505.500 | 87866.000 | 105515.000 |
| 64 | publish | mean_ns | 78555.667 | 78322.800 | 79046.533 |
| 64 | publish | rss_kib | 4488.000 | 4364.000 | 4542.000 |
| 64 | publish | minor_faults | 408.000 | 406.000 | 410.500 |
| 64 | publish | major_faults | 0.000 | 0.000 | 0.000 |
| 64 | lifecycle | p50_ns | 180811.000 | 180471.000 | 181345.500 |
| 64 | lifecycle | p95_ns | 192561.000 | 189785.500 | 195101.000 |
| 64 | lifecycle | p99_ns | 199100.500 | 191216.000 | 209236.000 |
| 64 | lifecycle | mean_ns | 182824.600 | 182445.950 | 183433.733 |
| 64 | lifecycle | rss_kib | 4430.000 | 4356.000 | 4460.000 |
| 64 | lifecycle | minor_faults | 407.000 | 403.500 | 413.000 |
| 64 | lifecycle | major_faults | 0.000 | 0.000 | 0.000 |
| 1024 | publish | p50_ns | 886933.500 | 885173.500 | 888789.000 |
| 1024 | publish | p95_ns | 896113.500 | 894709.000 | 897788.500 |
| 1024 | publish | p99_ns | 901869.000 | 897838.500 | 914263.500 |
| 1024 | publish | mean_ns | 887954.550 | 886412.650 | 889110.550 |
| 1024 | publish | rss_kib | 9068.000 | 9000.000 | 9150.000 |
| 1024 | publish | minor_faults | 1840.000 | 1836.500 | 1850.500 |
| 1024 | publish | major_faults | 0.000 | 0.000 | 0.000 |
| 1024 | lifecycle | p50_ns | 2384535.000 | 2379304.500 | 2392234.500 |
| 1024 | lifecycle | p95_ns | 2402515.000 | 2393040.000 | 2406025.500 |
| 1024 | lifecycle | p99_ns | 2406019.500 | 2397480.000 | 2419445.000 |
| 1024 | lifecycle | mean_ns | 2386367.950 | 2379556.117 | 2392689.817 |
| 1024 | lifecycle | rss_kib | 9114.000 | 9030.000 | 9150.000 |
| 1024 | lifecycle | minor_faults | 1841.500 | 1836.000 | 1846.500 |
| 1024 | lifecycle | major_faults | 0.000 | 0.000 | 0.000 |
| 4096 | publish | p50_ns | 3682305.000 | 3672660.000 | 3714830.000 |
| 4096 | publish | p95_ns | 3752860.500 | 3711460.000 | 3824945.500 |
| 4096 | publish | p99_ns | 3777676.000 | 3719141.000 | 3973816.500 |
| 4096 | publish | mean_ns | 3685291.400 | 3682049.183 | 3726241.583 |
| 4096 | publish | rss_kib | 22460.000 | 22450.000 | 22548.000 |
| 4096 | publish | minor_faults | 7055.000 | 7053.000 | 7055.500 |
| 4096 | publish | major_faults | 0.000 | 0.000 | 0.000 |
| 4096 | lifecycle | p50_ns | 9813840.000 | 9797041.000 | 9855530.500 |
| 4096 | lifecycle | p95_ns | 9899676.000 | 9875850.000 | 9912791.000 |
| 4096 | lifecycle | p99_ns | 9939386.500 | 9907526.000 | 9983446.000 |
| 4096 | lifecycle | mean_ns | 9824009.717 | 9807114.267 | 9859177.583 |
| 4096 | lifecycle | rss_kib | 22490.000 | 22450.000 | 22658.000 |
| 4096 | lifecycle | minor_faults | 7053.000 | 7053.000 | 7090.000 |
| 4096 | lifecycle | major_faults | 0.000 | 0.000 | 0.000 |

Callgrind guest instructions; each repeat is separate.

| Parts | Repeat | Ir |
|---:|---:|---:|
| 4 | 0 | 1081656 |
| 64 | 0 | 2172678 |
| 1024 | 0 | 19786337 |
| 4096 | 0 | 75998450 |
| 4 | 1 | 1080990 |
| 64 | 1 | 2173213 |
| 1024 | 1 | 19891812 |
| 4096 | 1 | 76205742 |

Heaptrack allocations whose full ancestry contains the owner; requested bytes, not peak live memory.

| Parts | Repeat | Allocation events | Requested bytes |
|---:|---:|---:|---:|
| 4 | 0 | 85 | 576175 |
| 64 | 0 | 568 | 704358 |
| 1024 | 0 | 8250 | 3408303 |
| 4096 | 0 | 32826 | 11419801 |
| 4 | 1 | 85 | 576175 |
| 64 | 1 | 568 | 704358 |
| 1024 | 1 | 8250 | 3408303 |
| 4096 | 1 | 32826 | 11419801 |
