### alloc

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich_lifecycle` | 408.473 [407.486–409.900] | 185.546 [177.359–202.612] | 0.4544 [0.4443, 0.4910] | - |
| `pptx_cross_copy_plain_lifecycle` | 8.917 [8.813–9.001] | 8.968 [8.851–9.179] | 1.0077 [0.9969, 1.0352] | - |

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- commit_ns: 89.279 -> 89.443
- plan_ns: 290.630 -> 67.212
- publication_ns: 6.210 -> 6.507

`pptx_cross_copy_media_rich_lifecycle` allocator counters (median of process medians):
- allocation_calls: 56,356 -> 56,407
- deallocation_calls: 46,068 -> 46,100
- reallocation_calls: 6,271 -> 6,274
- allocated_bytes: 277,809,048 -> 272,764,566
- region_peak_live_bytes: 305,068,513 -> 305,071,849

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms):
- commit_ns: 4.055 -> 4.118
- plan_ns: 3.427 -> 3.489
- publication_ns: 0.001 -> 0.001

`pptx_cross_copy_plain_lifecycle` allocator counters (median of process medians):
- allocation_calls: 46,613 -> 46,619
- deallocation_calls: 38,273 -> 38,276
- reallocation_calls: 5,298 -> 5,298
- allocated_bytes: 16,613,707 -> 16,615,405
- region_peak_live_bytes: 1,313,024 -> 1,313,576

### native

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich` | 384.408 [378.707–393.460] | 165.343 [152.640–169.300] | 0.4292 [0.4196, 0.4339] | - |
| `pptx_cross_copy_media_rich_lifecycle` | 405.441 [400.360–413.634] | 185.703 [180.666–197.265] | 0.4609 [0.4460, 0.4739] | - |
| `pptx_cross_copy_plain` | 7.062 [6.954–7.167] | 7.193 [7.106–7.410] | 1.0215 [1.0144, 1.0413] | - |
| `pptx_cross_copy_plain_lifecycle` | 8.292 [8.216–8.487] | 8.359 [8.305–8.542] | 1.0097 [0.9957, 1.0172] | - |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 16.806 [12.030–17.036] | 16.746 [12.038–17.069] | 0.9973 [0.9838, 1.3858] | - |

`pptx_cross_copy_media_rich` phases (median of process medians, ms):
- commit_ns: 88.673 -> 88.736
- plan_ns: 289.123 -> 70.135
- publication_ns: 6.410 -> 6.482

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- commit_ns: 88.585 -> 88.850
- plan_ns: 288.573 -> 68.674
- publication_ns: 6.451 -> 6.336

`pptx_cross_copy_plain` phases (median of process medians, ms):
- commit_ns: 3.786 -> 3.854
- plan_ns: 3.283 -> 3.342
- publication_ns: 0.001 -> 0.001

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms):
- commit_ns: 3.767 -> 3.801
- plan_ns: 3.269 -> 3.300
- publication_ns: 0.001 -> 0.001

`pptx_source_backed_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- open_ns: 0.626 -> 0.619
- plan_ns: 4.114 -> 4.113
- publication_ns: 12.065 -> 12.051
