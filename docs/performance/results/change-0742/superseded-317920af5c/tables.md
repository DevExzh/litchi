### alloc

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich_lifecycle` | 410.032 [405.442–423.347] | 186.648 [177.299–197.484] | 0.4493 [0.4349, 0.4662] | - |
| `pptx_cross_copy_plain_lifecycle` | 8.902 [8.779–9.077] | 8.965 [8.850–9.149] | 1.0075 [0.9972, 1.0199] | - |

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- commit_ns: 89.960 -> 89.974
- plan_ns: 292.334 -> 70.295
- publication_ns: 6.387 -> 3.645

`pptx_cross_copy_media_rich_lifecycle` allocator counters (median of process medians):
- allocation_calls: 56,356 -> 56,404
- deallocation_calls: 46,068 -> 46,100
- reallocation_calls: 6,271 -> 6,274
- allocated_bytes: 277,809,048 -> 272,765,734
- region_peak_live_bytes: 305,068,513 -> 305,071,641

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms):
- commit_ns: 4.044 -> 4.088
- plan_ns: 3.452 -> 3.473
- publication_ns: 0.001 -> 0.001

`pptx_cross_copy_plain_lifecycle` allocator counters (median of process medians):
- allocation_calls: 46,613 -> 46,615
- deallocation_calls: 38,273 -> 38,275
- reallocation_calls: 5,298 -> 5,298
- allocated_bytes: 16,613,707 -> 16,614,997
- region_peak_live_bytes: 1,313,024 -> 1,313,384

### native

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich` | 390.958 [387.192–394.575] | 167.732 [155.223–174.401] | 0.4276 [0.3982, 0.4499] | - |
| `pptx_cross_copy_media_rich_lifecycle` | 408.126 [406.546–412.558] | 194.768 [182.585–201.947] | 0.4742 [0.4611, 0.4837] | - |
| `pptx_cross_copy_plain` | 7.238 [7.162–7.346] | 7.241 [7.067–7.318] | 0.9999 [0.9741, 1.0116] | - |
| `pptx_cross_copy_plain_lifecycle` | 8.426 [8.345–8.477] | 8.448 [8.387–8.483] | 1.0006 [0.9952, 1.0107] | - |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 17.389 [17.108–18.205] | 17.215 [12.204–17.494] | 0.9732 [0.7124, 1.0068] | - |

`pptx_cross_copy_media_rich` phases (median of process medians, ms):
- commit_ns: 89.530 -> 89.896
- plan_ns: 294.535 -> 70.962
- publication_ns: 6.474 -> 6.613

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- commit_ns: 88.922 -> 95.821
- plan_ns: 290.645 -> 70.577
- publication_ns: 6.351 -> 6.476

`pptx_cross_copy_plain` phases (median of process medians, ms):
- commit_ns: 3.860 -> 3.863
- plan_ns: 3.377 -> 3.364
- publication_ns: 0.002 -> 0.002

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms):
- commit_ns: 3.815 -> 3.821
- plan_ns: 3.328 -> 3.332
- publication_ns: 0.001 -> 0.001

`pptx_source_backed_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- open_ns: 0.650 -> 0.651
- plan_ns: 4.226 -> 4.229
- publication_ns: 12.515 -> 12.334
