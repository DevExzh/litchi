### alloc

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich_lifecycle` | 183.320 [178.849–196.339] | 105.802 [99.453–114.509] | 0.5722 [0.5456, 0.6161] | - |
| `pptx_cross_copy_plain_lifecycle` | 9.063 [8.837–9.389] | 8.753 [8.618–8.945] | 0.9647 [0.9527, 0.9726] | - |

| case | p95 before → after (ms) | mean before → after (ms) |
|---|---:|---:|
| `pptx_cross_copy_media_rich_lifecycle` | 184.257 → 108.970 | 183.486 → 106.741 |
| `pptx_cross_copy_plain_lifecycle` | 9.121 → 8.835 | 9.049 → 8.770 |

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms; median paired ratio):
- commit_ns: 90.051 -> 26.556 (0.2941 [0.2750, 0.2954])
- plan_ns: 64.868 -> 51.991 (0.7732 [0.6934, 0.8884])
- publication_ns: 6.150 -> 6.144 (1.0071 [0.2404, 3.6840])

`pptx_cross_copy_media_rich_lifecycle` allocator counters (median of process medians):
- allocation_calls: 56,384 -> 56,394
- deallocation_calls: 46,086 -> 46,093
- reallocation_calls: 6,272 -> 6,272
- allocated_bytes: 272,765,116 -> 239,200,195
- region_peak_live_bytes: 305,070,945 -> 288,642,369

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms; median paired ratio):
- commit_ns: 4.127 -> 3.887 (0.9398 [0.9203, 0.9541])
- plan_ns: 3.535 -> 3.454 (0.9736 [0.9539, 0.9922])
- publication_ns: 0.001 -> 0.001 (0.9484 [0.6190, 1.3590])

`pptx_cross_copy_plain_lifecycle` allocator counters (median of process medians):
- allocation_calls: 46,618 -> 46,628
- deallocation_calls: 38,275 -> 38,281
- reallocation_calls: 5,298 -> 5,298
- allocated_bytes: 16,615,187 -> 16,600,474
- region_peak_live_bytes: 1,313,408 -> 1,317,213

### native

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich` | 159.629 [154.153–173.494] | 80.585 [79.910–94.182] | 0.5137 [0.5021, 0.5266] | - |
| `pptx_cross_copy_media_rich_lifecycle` | 184.195 [175.207–202.181] | 109.760 [101.720–117.298] | 0.6083 [0.5228, 0.6438] | - |
| `pptx_cross_copy_plain` | 7.121 [7.094–7.173] | 6.823 [6.789–6.904] | 0.9586 [0.9519, 0.9655] | - |
| `pptx_cross_copy_plain_lifecycle` | 8.454 [8.359–8.818] | 8.166 [8.154–8.175] | 0.9666 [0.9486, 0.9729] | - |
| `pptx_semantic_one_edit_save [pptx-semantic-large]` | 54.667 [53.643–55.520] | 54.492 [54.043–54.886] | 0.9940 [0.9863, 1.0135] | - |
| `pptx_semantic_one_edit_save [pptx-semantic-medium]` | 1.700 [1.685–1.702] | 1.711 [1.707–1.719] | 1.0084 [1.0058, 1.0155] | - |
| `pptx_semantic_one_edit_save [pptx-semantic-tiny]` | 0.907 [0.903–0.913] | 0.919 [0.914–0.922] | 1.0112 [1.0096, 1.0174] | - |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 17.085 [11.972–17.581] | 17.117 [12.518–17.354] | 1.0013 [0.9871, 1.4190] | - |

| case | p95 before → after (ms) | mean before → after (ms) |
|---|---:|---:|
| `pptx_cross_copy_media_rich` | 160.600 → 81.141 | 159.760 → 80.626 |
| `pptx_cross_copy_media_rich_lifecycle` | 184.709 → 110.175 | 184.215 → 109.753 |
| `pptx_cross_copy_plain` | 7.356 → 6.958 | 7.139 → 6.841 |
| `pptx_cross_copy_plain_lifecycle` | 8.637 → 8.269 | 8.453 → 8.177 |
| `pptx_semantic_one_edit_save [pptx-semantic-large]` | 55.676 → 55.247 | 54.656 → 54.507 |
| `pptx_semantic_one_edit_save [pptx-semantic-medium]` | 1.715 → 1.724 | 1.700 → 1.712 |
| `pptx_semantic_one_edit_save [pptx-semantic-tiny]` | 0.917 → 0.927 | 0.907 → 0.920 |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 17.349 → 17.363 | 17.119 → 17.120 |

| case | cycles before → after (whole child) | ratio [95%] | instructions before → after | ratio [95%] |
|---|---:|---:|---:|---:|
| `pptx_cross_copy_media_rich` | 36,524,168,166 → 27,128,040,868 | 0.7462 [0.7335, 0.7709] | 49,813,973,410 → 38,326,967,654 | 0.7715 [0.7589, 0.7913] |
| `pptx_cross_copy_media_rich_lifecycle` | 36,727,150,103 → 28,034,840,054 | 0.7665 [0.7119, 0.7931] | 50,462,174,146 → 39,253,096,289 | 0.7840 [0.7494, 0.8050] |
| `pptx_cross_copy_plain` | 3,660,133,963 → 3,606,602,268 | 0.9826 [0.9714, 0.9984] | 9,509,845,328 → 9,467,648,902 | 0.9955 [0.9954, 0.9957] |
| `pptx_cross_copy_plain_lifecycle` | 3,700,830,598 → 3,635,978,842 | 0.9851 [0.9788, 0.9869] | 9,546,659,658 → 9,504,746,443 | 0.9957 [0.9955, 0.9958] |
| `pptx_semantic_one_edit_save [pptx-semantic-large]` | 200,847,948,858 → 202,570,874,204 | 1.0066 [1.0024, 1.0118] | 832,494,957,823 → 832,125,116,927 | 0.9995 [0.9995, 0.9996] |
| `pptx_semantic_one_edit_save [pptx-semantic-medium]` | 200,847,948,858 → 202,570,874,204 | 1.0066 [1.0024, 1.0118] | 832,494,957,823 → 832,125,116,927 | 0.9995 [0.9995, 0.9996] |
| `pptx_semantic_one_edit_save [pptx-semantic-tiny]` | 200,847,948,858 → 202,570,874,204 | 1.0066 [1.0024, 1.0118] | 832,494,957,823 → 832,125,116,927 | 0.9995 [0.9995, 0.9996] |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 20,465,157,282 → 19,732,981,781 | 0.9611 [0.9509, 1.0306] | 33,666,611,652 → 32,597,664,587 | 0.9715 [0.9546, 1.0263] |

`pptx_cross_copy_media_rich` phases (median of process medians, ms; median paired ratio):
- commit_ns: 89.258 -> 25.870 (0.2906 [0.2877, 0.2922])
- plan_ns: 64.037 -> 48.355 (0.7648 [0.7477, 0.7831])
- publication_ns: 6.413 -> 6.380 (1.0075 [0.9671, 3.8054])

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms; median paired ratio):
- commit_ns: 88.898 -> 25.846 (0.2905 [0.2690, 0.2919])
- plan_ns: 67.027 -> 55.305 (0.8705 [0.6848, 0.9798])
- publication_ns: 6.463 -> 6.371 (0.9855 [0.9357, 4.1468])

`pptx_cross_copy_plain` phases (median of process medians, ms; median paired ratio):
- commit_ns: 3.808 -> 3.592 (0.9426 [0.9364, 0.9481])
- plan_ns: 3.316 -> 3.230 (0.9752 [0.9675, 0.9829])
- publication_ns: 0.001 -> 0.001 (0.8808 [0.7947, 1.2994])

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms; median paired ratio):
- commit_ns: 3.835 -> 3.607 (0.9410 [0.9227, 0.9465])
- plan_ns: 3.354 -> 3.270 (0.9765 [0.9507, 0.9869])
- publication_ns: 0.001 -> 0.001 (1.1088 [0.8683, 1.5120])

`pptx_source_backed_cross_copy_media_rich_lifecycle` phases (median of process medians, ms; median paired ratio):
- open_ns: 0.644 -> 0.655
- plan_ns: 4.190 -> 4.174 (1.0019 [0.9904, 1.0091])
- publication_ns: 12.262 -> 12.275 (0.9985 [0.9855, 1.6876])

## Per-process rows

`alloc` `pptx_cross_copy_media_rich_lifecycle`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 181.421 | 182.371 | 181.433 |
| r0-s1 | after | 106.974 | 107.535 | 107.107 |
| r0-s2 | after | 104.079 | 104.627 | 104.129 |
| r0-s3 | before | 189.699 | 198.863 | 192.328 |
| r1-s0 | before | 178.849 | 179.079 | 178.908 |
| r1-s1 | after | 110.190 | 110.406 | 110.122 |
| r1-s2 | after | 104.629 | 110.822 | 106.375 |
| r1-s3 | before | 196.339 | 197.054 | 196.488 |
| r2-s0 | before | 184.169 | 185.727 | 184.539 |
| r2-s1 | after | 114.509 | 114.807 | 114.321 |
| r2-s2 | after | 110.843 | 110.986 | 110.784 |
| r2-s3 | before | 192.146 | 195.542 | 193.135 |
| r3-s0 | before | 182.471 | 182.787 | 182.433 |
| r3-s1 | after | 103.570 | 103.776 | 103.613 |
| r3-s2 | after | 99.453 | 99.513 | 99.434 |
| r3-s3 | before | 182.291 | 182.640 | 182.397 |

`alloc` `pptx_cross_copy_plain_lifecycle`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 8.995 | 9.079 | 9.016 |
| r0-s1 | after | 8.733 | 8.822 | 8.745 |
| r0-s2 | after | 8.945 | 8.955 | 8.861 |
| r0-s3 | before | 9.389 | 9.590 | 9.408 |
| r1-s0 | before | 9.276 | 9.286 | 9.256 |
| r1-s1 | after | 8.734 | 9.347 | 8.895 |
| r1-s2 | after | 8.662 | 8.787 | 8.690 |
| r1-s3 | before | 8.907 | 8.962 | 8.925 |
| r2-s0 | before | 9.147 | 9.150 | 9.081 |
| r2-s1 | after | 8.808 | 8.818 | 8.771 |
| r2-s2 | after | 8.823 | 8.989 | 8.875 |
| r2-s3 | before | 9.131 | 9.278 | 9.166 |
| r3-s0 | before | 8.837 | 8.880 | 8.848 |
| r3-s1 | after | 8.773 | 8.848 | 8.769 |
| r3-s2 | after | 8.618 | 8.619 | 8.618 |
| r3-s3 | before | 8.992 | 9.092 | 8.993 |

`native` `pptx_cross_copy_media_rich`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 159.152 | 159.957 | 159.256 |
| r0-s1 | after | 79.910 | 80.305 | 79.934 |
| r0-s2 | after | 87.011 | 87.379 | 86.991 |
| r0-s3 | before | 173.494 | 174.099 | 173.504 |
| r1-s0 | before | 154.153 | 155.769 | 154.324 |
| r1-s1 | after | 80.529 | 80.799 | 80.500 |
| r1-s2 | after | 87.470 | 88.445 | 87.506 |
| r1-s3 | before | 166.111 | 166.819 | 166.135 |
| r2-s0 | before | 154.731 | 155.284 | 154.774 |
| r2-s1 | after | 80.642 | 81.298 | 80.686 |
| r2-s2 | after | 80.488 | 80.643 | 80.418 |
| r2-s3 | before | 160.106 | 161.242 | 160.264 |
| r3-s0 | before | 165.180 | 166.518 | 163.183 |
| r3-s1 | after | 94.182 | 94.474 | 94.297 |
| r3-s2 | after | 80.506 | 80.985 | 80.567 |
| r3-s3 | before | 159.026 | 159.568 | 159.061 |

`native` `pptx_cross_copy_media_rich_lifecycle`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 180.141 | 180.873 | 180.202 |
| r0-s1 | after | 115.974 | 116.664 | 116.107 |
| r0-s2 | after | 101.720 | 102.295 | 101.851 |
| r0-s3 | before | 194.577 | 195.108 | 194.614 |
| r1-s0 | before | 187.525 | 188.006 | 187.538 |
| r1-s1 | after | 114.078 | 114.745 | 114.051 |
| r1-s2 | after | 110.018 | 110.419 | 110.040 |
| r1-s3 | before | 180.865 | 181.413 | 180.893 |
| r2-s0 | before | 191.812 | 195.447 | 191.715 |
| r2-s1 | after | 109.471 | 109.894 | 109.466 |
| r2-s2 | after | 117.298 | 117.750 | 117.218 |
| r2-s3 | before | 175.207 | 176.018 | 175.282 |
| r3-s0 | before | 175.852 | 176.400 | 175.879 |
| r3-s1 | after | 109.503 | 109.931 | 109.442 |
| r3-s2 | after | 101.969 | 102.488 | 101.973 |
| r3-s3 | before | 202.181 | 202.933 | 202.275 |

`native` `pptx_cross_copy_plain`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 7.161 | 7.354 | 7.187 |
| r0-s1 | after | 6.817 | 6.898 | 6.828 |
| r0-s2 | after | 6.817 | 6.971 | 6.839 |
| r0-s3 | before | 7.102 | 7.358 | 7.142 |
| r1-s0 | before | 7.152 | 7.421 | 7.190 |
| r1-s1 | after | 6.789 | 6.944 | 6.810 |
| r1-s2 | after | 6.794 | 6.910 | 6.816 |
| r1-s3 | before | 7.097 | 7.309 | 7.120 |
| r2-s0 | before | 7.094 | 7.384 | 7.136 |
| r2-s1 | after | 6.849 | 6.987 | 6.867 |
| r2-s2 | after | 6.828 | 6.916 | 6.843 |
| r2-s3 | before | 7.173 | 7.367 | 7.184 |
| r3-s0 | before | 7.112 | 7.320 | 7.132 |
| r3-s1 | after | 6.850 | 6.972 | 6.868 |
| r3-s2 | after | 6.904 | 7.070 | 6.921 |
| r3-s3 | before | 7.130 | 7.283 | 7.111 |

`native` `pptx_cross_copy_plain_lifecycle`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 8.421 | 8.612 | 8.426 |
| r0-s1 | after | 8.174 | 8.273 | 8.191 |
| r0-s2 | after | 8.164 | 8.369 | 8.192 |
| r0-s3 | before | 8.359 | 8.543 | 8.348 |
| r1-s0 | before | 8.491 | 8.715 | 8.486 |
| r1-s1 | after | 8.155 | 8.283 | 8.179 |
| r1-s2 | after | 8.158 | 8.265 | 8.173 |
| r1-s3 | before | 8.818 | 9.087 | 8.819 |
| r2-s0 | before | 8.396 | 8.561 | 8.401 |
| r2-s1 | after | 8.168 | 8.246 | 8.166 |
| r2-s2 | after | 8.175 | 8.254 | 8.182 |
| r2-s3 | before | 8.455 | 8.611 | 8.448 |
| r3-s0 | before | 8.596 | 8.832 | 8.578 |
| r3-s1 | after | 8.154 | 8.279 | 8.165 |
| r3-s2 | after | 8.169 | 8.257 | 8.174 |
| r3-s3 | before | 8.453 | 8.663 | 8.458 |

`native` `pptx_semantic_one_edit_save [pptx-semantic-large]`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 53.643 | 54.472 | 53.644 |
| r0-s1 | after | 54.369 | 55.073 | 54.382 |
| r0-s2 | after | 54.615 | 55.423 | 54.632 |
| r0-s3 | before | 53.665 | 54.755 | 53.753 |
| r1-s0 | before | 54.373 | 55.558 | 54.448 |
| r1-s1 | after | 54.657 | 55.421 | 54.670 |
| r1-s2 | after | 54.115 | 54.815 | 54.128 |
| r1-s3 | before | 55.139 | 56.419 | 55.224 |
| r2-s0 | before | 54.753 | 55.674 | 54.701 |
| r2-s1 | after | 54.328 | 55.029 | 54.365 |
| r2-s2 | after | 54.043 | 54.674 | 54.065 |
| r2-s3 | before | 54.581 | 55.678 | 54.610 |
| r3-s0 | before | 55.520 | 56.803 | 55.522 |
| r3-s1 | after | 54.762 | 55.869 | 54.818 |
| r3-s2 | after | 54.886 | 55.839 | 54.881 |
| r3-s3 | before | 55.123 | 56.349 | 55.148 |

`native` `pptx_semantic_one_edit_save [pptx-semantic-medium]`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 1.685 | 1.702 | 1.690 |
| r0-s1 | after | 1.707 | 1.721 | 1.708 |
| r0-s2 | after | 1.719 | 1.735 | 1.721 |
| r0-s3 | before | 1.687 | 1.713 | 1.690 |
| r1-s0 | before | 1.688 | 1.702 | 1.689 |
| r1-s1 | after | 1.714 | 1.726 | 1.715 |
| r1-s2 | after | 1.711 | 1.723 | 1.711 |
| r1-s3 | before | 1.699 | 1.718 | 1.700 |
| r2-s0 | before | 1.702 | 1.715 | 1.704 |
| r2-s1 | after | 1.719 | 1.734 | 1.720 |
| r2-s2 | after | 1.711 | 1.722 | 1.712 |
| r2-s3 | before | 1.700 | 1.719 | 1.700 |
| r3-s0 | before | 1.702 | 1.715 | 1.703 |
| r3-s1 | after | 1.710 | 1.721 | 1.711 |
| r3-s2 | after | 1.712 | 1.729 | 1.713 |
| r3-s3 | before | 1.702 | 1.721 | 1.702 |

`native` `pptx_semantic_one_edit_save [pptx-semantic-tiny]`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 0.903 | 0.913 | 0.904 |
| r0-s1 | after | 0.918 | 0.926 | 0.918 |
| r0-s2 | after | 0.922 | 0.932 | 0.923 |
| r0-s3 | before | 0.904 | 0.912 | 0.904 |
| r1-s0 | before | 0.905 | 0.914 | 0.906 |
| r1-s1 | after | 0.920 | 0.927 | 0.920 |
| r1-s2 | after | 0.914 | 0.924 | 0.914 |
| r1-s3 | before | 0.905 | 0.915 | 0.905 |
| r2-s0 | before | 0.909 | 0.920 | 0.909 |
| r2-s1 | after | 0.919 | 0.927 | 0.920 |
| r2-s2 | after | 0.918 | 0.927 | 0.918 |
| r2-s3 | before | 0.908 | 0.918 | 0.908 |
| r3-s0 | before | 0.913 | 0.924 | 0.914 |
| r3-s1 | after | 0.921 | 0.931 | 0.921 |
| r3-s2 | after | 0.919 | 0.928 | 0.920 |
| r3-s3 | before | 0.909 | 0.918 | 0.910 |

`native` `pptx_source_backed_cross_copy_media_rich_lifecycle`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 17.021 | 17.158 | 17.028 |
| r0-s1 | after | 17.088 | 17.280 | 17.062 |
| r0-s2 | after | 12.518 | 12.746 | 12.508 |
| r0-s3 | before | 17.232 | 17.379 | 17.218 |
| r1-s0 | before | 17.581 | 17.704 | 17.597 |
| r1-s1 | after | 17.354 | 17.509 | 17.351 |
| r1-s2 | after | 17.125 | 17.499 | 17.147 |
| r1-s3 | before | 17.148 | 17.414 | 17.183 |
| r2-s0 | before | 17.224 | 17.374 | 17.191 |
| r2-s1 | after | 17.069 | 17.242 | 17.090 |
| r2-s2 | after | 17.110 | 17.278 | 17.094 |
| r2-s3 | before | 11.972 | 12.274 | 12.003 |
| r3-s0 | before | 17.018 | 17.325 | 17.054 |
| r3-s1 | after | 17.271 | 17.446 | 17.227 |
| r3-s2 | after | 17.343 | 17.646 | 17.306 |
| r3-s3 | before | 12.222 | 12.382 | 12.182 |
