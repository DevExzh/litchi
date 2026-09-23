### alloc

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich_lifecycle` | 406.561 [405.248–414.672] | 182.356 [176.390–202.327] | 0.4482 [0.4337, 0.4648] | - |
| `pptx_cross_copy_plain_lifecycle` | 8.807 [8.770–8.889] | 8.940 [8.839–9.073] | 1.0170 [1.0040, 1.0286] | - |

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- commit_ns: 89.004 -> 89.057
- plan_ns: 289.374 -> 64.272
- publication_ns: 6.427 -> 6.462

`pptx_cross_copy_media_rich_lifecycle` allocator counters (median of process medians):
- allocation_calls: 56,357 -> 56,411
- deallocation_calls: 46,068 -> 46,104
- reallocation_calls: 6,271 -> 6,274
- allocated_bytes: 277,810,624 -> 272,767,478
- region_peak_live_bytes: 305,068,449 -> 305,072,489

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms):
- commit_ns: 4.000 -> 4.097
- plan_ns: 3.400 -> 3.454
- publication_ns: 0.001 -> 0.001

`pptx_cross_copy_plain_lifecycle` allocator counters (median of process medians):
- allocation_calls: 46,613 -> 46,623
- deallocation_calls: 38,273 -> 38,280
- reallocation_calls: 5,298 -> 5,298
- allocated_bytes: 16,613,707 -> 16,618,317
- region_peak_live_bytes: 1,313,024 -> 1,314,216

### native

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich` | 383.630 [382.137–389.990] | 161.873 [158.354–180.344] | 0.4231 [0.4120, 0.4429] | - |
| `pptx_cross_copy_media_rich_lifecycle` | 411.384 [405.266–418.802] | 186.599 [176.258–201.386] | 0.4522 [0.4339, 0.4802] | - |
| `pptx_cross_copy_plain` | 7.018 [6.981–7.203] | 7.016 [6.979–7.116] | 0.9968 [0.9854, 1.0048] | - |
| `pptx_cross_copy_plain_lifecycle` | 8.361 [8.259–8.396] | 8.291 [8.236–8.347] | 0.9924 [0.9897, 0.9970] | - |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 17.001 [12.104–17.793] | 16.935 [12.060–17.447] | 0.9842 [0.9623, 1.3970] | - |

`pptx_cross_copy_media_rich` phases (median of process medians, ms):
- commit_ns: 88.494 -> 88.891
- plan_ns: 288.997 -> 66.660
- publication_ns: 6.386 -> 6.212

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- commit_ns: 88.968 -> 89.664
- plan_ns: 293.351 -> 68.741
- publication_ns: 6.402 -> 6.513

`pptx_cross_copy_plain` phases (median of process medians, ms):
- commit_ns: 3.766 -> 3.758
- plan_ns: 3.258 -> 3.245
- publication_ns: 0.001 -> 0.001

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms):
- commit_ns: 3.794 -> 3.761
- plan_ns: 3.306 -> 3.267
- publication_ns: 0.001 -> 0.001

`pptx_source_backed_cross_copy_media_rich_lifecycle` phases (median of process medians, ms):
- open_ns: 0.646 -> 0.647
- plan_ns: 4.192 -> 4.186
- publication_ns: 12.153 -> 12.087
