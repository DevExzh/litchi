### alloc

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich_lifecycle` | 409.232 [407.017–419.433] | 186.962 [175.854–199.118] | 0.4552 [0.4430, 0.4759] | - |
| `pptx_cross_copy_plain_lifecycle` | 8.958 [8.817–9.213] | 9.202 [8.926–9.349] | 1.0237 [1.0029, 1.0502] | - |

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- commit_ns: 89.630 -> 89.329
- plan_ns: 290.899 -> 67.395
- publication_ns: 6.478 -> 6.519

`pptx_cross_copy_media_rich_lifecycle` allocator counters (median of process medians):
- allocation_calls: 56,356 -> 56,383
- deallocation_calls: 46,068 -> 46,086
- reallocation_calls: 6,271 -> 6,272
- allocated_bytes: 277,809,048 -> 272,763,540
- region_peak_live_bytes: 305,068,449 -> 305,071,038

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms):
- commit_ns: 4.070 -> 4.163
- plan_ns: 3.492 -> 3.551
- publication_ns: 0.001 -> 0.002

`pptx_cross_copy_plain_lifecycle` allocator counters (median of process medians):
- allocation_calls: 46,613 -> 46,618
- deallocation_calls: 38,273 -> 38,275
- reallocation_calls: 5,298 -> 5,298
- allocated_bytes: 16,613,707 -> 16,615,187
- region_peak_live_bytes: 1,313,024 -> 1,313,469

### native

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich` | 386.716 [383.267–397.587] | 158.999 [153.792–180.286] | 0.4125 [0.3985, 0.4413] | - |
| `pptx_cross_copy_media_rich_lifecycle` | 410.081 [405.783–416.086] | 183.024 [176.960–194.569] | 0.4439 [0.4406, 0.4585] | - |
| `pptx_cross_copy_plain` | 7.091 [6.962–7.245] | 7.147 [7.054–7.226] | 1.0097 [0.9902, 1.0228] | - |
| `pptx_cross_copy_plain_lifecycle` | 8.360 [8.225–8.485] | 8.401 [8.372–8.576] | 1.0070 [0.9985, 1.0189] | - |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 17.022 [12.103–17.492] | 16.678 [12.028–17.084] | 0.9830 [0.7077, 1.0106] | - |

`pptx_cross_copy_media_rich` phases (median of process medians, ms):
- commit_ns: 88.901 -> 89.157
- plan_ns: 291.186 -> 64.028
- publication_ns: 6.392 -> 6.416

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- commit_ns: 89.158 -> 89.123
- plan_ns: 292.579 -> 65.673
- publication_ns: 6.508 -> 6.298

`pptx_cross_copy_plain` phases (median of process medians, ms):
- commit_ns: 3.786 -> 3.815
- plan_ns: 3.294 -> 3.313
- publication_ns: 0.001 -> 0.001

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms):
- commit_ns: 3.805 -> 3.812
- plan_ns: 3.304 -> 3.323
- publication_ns: 0.001 -> 0.001

`pptx_source_backed_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- open_ns: 0.630 -> 0.635
- plan_ns: 4.101 -> 4.105
- publication_ns: 12.296 -> 12.034
