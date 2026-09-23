### alloc

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich_lifecycle` | 181.668 [175.220–197.012] | 103.777 [97.151–113.368] | 0.5743 [0.5471, 0.5986] | - |
| `pptx_cross_copy_plain_lifecycle` | 8.978 [8.855–9.362] | 8.666 [8.569–8.758] | 0.9681 [0.9512, 0.9802] | - |

| case | p95 before → after (ms) | mean before → after (ms) |
|---|---:|---:|
| `pptx_cross_copy_media_rich_lifecycle` | 185.041 → 106.901 | 182.575 → 104.278 |
| `pptx_cross_copy_plain_lifecycle` | 9.061 → 8.729 | 8.977 → 8.678 |

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms; median paired ratio):
- commit_ns: 89.120 -> 25.968 (0.2911 [0.2723, 0.2923])
- plan_ns: 63.990 -> 51.760 (0.7704 [0.7009, 0.8556])
- publication_ns: 3.878 -> 6.140 (0.9740 [0.8535, 1.0273])

`pptx_cross_copy_media_rich_lifecycle` allocator counters (median of process medians):
- allocation_calls: 56,384 -> 56,387
- deallocation_calls: 46,087 -> 46,090
- reallocation_calls: 6,272 -> 6,272
- allocated_bytes: 272,766,692 -> 239,185,859
- region_peak_live_bytes: 305,071,007 -> 288,634,911

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms; median paired ratio):
- commit_ns: 4.106 -> 3.853 (0.9379 [0.9277, 0.9512])
- plan_ns: 3.474 -> 3.419 (0.9882 [0.9608, 1.0040])
- publication_ns: 0.001 -> 0.001 (1.0734 [0.9588, 1.1519])

`pptx_cross_copy_plain_lifecycle` allocator counters (median of process medians):
- allocation_calls: 46,618 -> 46,622
- deallocation_calls: 38,275 -> 38,279
- reallocation_calls: 5,298 -> 5,298
- allocated_bytes: 16,615,187 -> 16,594,762
- region_peak_live_bytes: 1,313,406 -> 1,313,403

### native

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [bootstrap 95%] | flag |
|---|---:|---:|---:|---|
| `pptx_cross_copy_media_rich` | 161.920 [153.704–172.583] | 86.380 [75.401–93.649] | 0.5242 [0.4755, 0.5431] | - |
| `pptx_cross_copy_media_rich_lifecycle` | 180.778 [174.664–200.927] | 102.309 [101.536–115.669] | 0.5637 [0.5366, 0.5849] | - |
| `pptx_cross_copy_plain` | 7.079 [6.960–7.225] | 6.883 [6.778–6.935] | 0.9701 [0.9601, 0.9781] | - |
| `pptx_cross_copy_plain_lifecycle` | 8.320 [8.268–8.373] | 8.148 [8.065–8.238] | 0.9768 [0.9713, 0.9891] | - |
| `pptx_semantic_one_edit_save [pptx-semantic-large]` | 53.827 [53.299–55.543] | 53.889 [53.281–54.239] | 0.9952 [0.9897, 1.0061] | - |
| `pptx_semantic_one_edit_save [pptx-semantic-medium]` | 1.692 [1.682–1.701] | 1.699 [1.685–1.833] | 1.0030 [0.9980, 1.0079] | - |
| `pptx_semantic_one_edit_save [pptx-semantic-tiny]` | 0.905 [0.900–0.913] | 0.911 [0.907–0.995] | 1.0065 [1.0039, 1.0147] | - |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 17.053 [16.942–17.126] | 17.063 [12.139–17.137] | 0.9991 [0.9783, 1.0052] | - |

| case | p95 before → after (ms) | mean before → after (ms) |
|---|---:|---:|
| `pptx_cross_copy_media_rich` | 162.391 → 86.770 | 161.946 → 86.417 |
| `pptx_cross_copy_media_rich_lifecycle` | 181.260 → 103.785 | 180.737 → 103.150 |
| `pptx_cross_copy_plain` | 7.253 → 7.035 | 7.080 → 6.877 |
| `pptx_cross_copy_plain_lifecycle` | 8.455 → 8.316 | 8.308 → 8.139 |
| `pptx_semantic_one_edit_save [pptx-semantic-large]` | 54.940 → 54.920 | 53.803 → 53.952 |
| `pptx_semantic_one_edit_save [pptx-semantic-medium]` | 1.709 → 1.713 | 1.693 → 1.701 |
| `pptx_semantic_one_edit_save [pptx-semantic-tiny]` | 0.916 → 0.920 | 0.905 → 0.912 |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 17.243 → 17.231 | 17.033 → 17.036 |

| case | cycles before → after (whole child) | ratio [95%] | instructions before → after | ratio [95%] |
|---|---:|---:|---:|---:|
| `pptx_cross_copy_media_rich` | 36,296,711,364 → 27,210,050,914 | 0.7435 [0.7207, 0.7597] | 49,757,564,398 → 38,524,049,616 | 0.7675 [0.7447, 0.7791] |
| `pptx_cross_copy_media_rich_lifecycle` | 35,951,407,902 → 27,302,436,768 | 0.7645 [0.7261, 0.7969] | 49,573,370,821 → 38,821,542,423 | 0.7835 [0.7513, 0.8208] |
| `pptx_cross_copy_plain` | 3,571,352,235 → 3,564,834,874 | 0.9978 [0.9946, 1.0005] | 9,508,297,982 → 9,452,319,585 | 0.9941 [0.9939, 0.9942] |
| `pptx_cross_copy_plain_lifecycle` | 3,586,349,550 → 3,592,190,202 | 1.0007 [0.9945, 1.0089] | 9,546,184,978 → 9,489,958,746 | 0.9942 [0.9941, 0.9945] |
| `pptx_semantic_one_edit_save [pptx-semantic-large]` | 200,222,877,183 → 200,705,205,632 | 1.0027 [0.9936, 1.0112] | 832,516,364,120 → 832,144,452,582 | 0.9996 [0.9980, 0.9998] |
| `pptx_semantic_one_edit_save [pptx-semantic-medium]` | 200,222,877,183 → 200,705,205,632 | 1.0027 [0.9936, 1.0112] | 832,516,364,120 → 832,144,452,582 | 0.9996 [0.9980, 0.9998] |
| `pptx_semantic_one_edit_save [pptx-semantic-tiny]` | 200,222,877,183 → 200,705,205,632 | 1.0027 [0.9936, 1.0112] | 832,516,364,120 → 832,144,452,582 | 0.9996 [0.9980, 0.9998] |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 20,427,867,406 → 19,675,988,203 | 0.9626 [0.9597, 0.9704] | 33,673,313,906 → 32,677,361,260 | 0.9701 [0.9688, 0.9749] |

`pptx_cross_copy_media_rich` phases (median of process medians, ms; median paired ratio):
- commit_ns: 88.649 -> 25.659 (0.2894 [0.2889, 0.2903])
- plan_ns: 66.747 -> 54.429 (0.7791 [0.7587, 0.8534])
- publication_ns: 6.400 -> 6.297 (0.9699 [0.9574, 1.0046])

`pptx_cross_copy_media_rich_lifecycle` phases (median of process medians, ms; median paired ratio):
- commit_ns: 88.691 -> 25.674 (0.2890 [0.2688, 0.2896])
- plan_ns: 63.885 -> 48.430 (0.7541 [0.7180, 0.7859])
- publication_ns: 6.294 -> 6.218 (0.9929 [0.9332, 4.1112])

`pptx_cross_copy_plain` phases (median of process medians, ms; median paired ratio):
- commit_ns: 3.793 -> 3.619 (0.9534 [0.9429, 0.9609])
- plan_ns: 3.293 -> 3.255 (0.9895 [0.9724, 0.9962])
- publication_ns: 0.001 -> 0.001 (0.9449 [0.9189, 0.9517])

`pptx_cross_copy_plain_lifecycle` phases (median of process medians, ms; median paired ratio):
- commit_ns: 3.787 -> 3.623 (0.9548 [0.9442, 0.9605])
- plan_ns: 3.285 -> 3.274 (0.9922 [0.9793, 1.0078])
- publication_ns: 0.001 -> 0.001 (0.9605 [0.9342, 0.9870])

`pptx_source_backed_cross_copy_media_rich_lifecycle` phases (median of process medians, ms; median paired ratio):
- open_ns: 0.633 -> 0.644
- plan_ns: 4.116 -> 4.119 (1.0003 [0.9950, 1.0060])
- publication_ns: 12.291 -> 12.301 (0.9983 [0.9776, 1.0050])

## Per-process rows

`alloc` `pptx_cross_copy_media_rich_lifecycle`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 196.744 | 197.061 | 196.735 |
| r0-s1 | after | 113.368 | 113.872 | 113.281 |
| r0-s2 | after | 102.862 | 103.229 | 102.910 |
| r0-s3 | before | 188.021 | 194.998 | 190.166 |
| r1-s0 | before | 197.012 | 197.999 | 197.310 |
| r1-s1 | after | 104.151 | 104.574 | 104.278 |
| r1-s2 | after | 102.024 | 109.227 | 104.278 |
| r1-s3 | before | 175.220 | 175.424 | 175.208 |
| r2-s0 | before | 176.525 | 176.617 | 176.542 |
| r2-s1 | after | 97.151 | 101.530 | 98.590 |
| r2-s2 | after | 109.032 | 109.395 | 109.124 |
| r2-s3 | before | 182.143 | 188.181 | 183.835 |
| r3-s0 | before | 181.193 | 181.901 | 181.316 |
| r3-s1 | after | 109.473 | 109.790 | 109.504 |
| r3-s2 | after | 103.402 | 103.607 | 103.458 |
| r3-s3 | before | 180.653 | 181.525 | 180.792 |

`alloc` `pptx_cross_copy_plain_lifecycle`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 9.100 | 9.259 | 9.098 |
| r0-s1 | after | 8.696 | 8.724 | 8.705 |
| r0-s2 | after | 8.731 | 8.734 | 8.708 |
| r0-s3 | before | 8.907 | 8.969 | 8.918 |
| r1-s0 | before | 9.362 | 9.483 | 9.338 |
| r1-s1 | after | 8.569 | 8.582 | 8.571 |
| r1-s2 | after | 8.635 | 8.749 | 8.650 |
| r1-s3 | before | 8.886 | 9.050 | 8.927 |
| r2-s0 | before | 9.060 | 9.194 | 9.085 |
| r2-s1 | after | 8.618 | 8.665 | 8.616 |
| r2-s2 | after | 8.718 | 8.772 | 8.720 |
| r2-s3 | before | 8.855 | 8.881 | 8.855 |
| r3-s0 | before | 8.928 | 9.000 | 8.948 |
| r3-s1 | after | 8.623 | 8.633 | 8.625 |
| r3-s2 | after | 8.758 | 8.814 | 8.730 |
| r3-s3 | before | 9.028 | 9.071 | 9.006 |

`native` `pptx_cross_copy_media_rich`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 164.831 | 165.184 | 164.868 |
| r0-s1 | after | 86.767 | 87.353 | 86.847 |
| r0-s2 | after | 87.685 | 88.282 | 87.685 |
| r0-s3 | before | 165.510 | 166.250 | 165.584 |
| r1-s0 | before | 172.583 | 173.113 | 172.657 |
| r1-s1 | after | 80.359 | 80.815 | 80.320 |
| r1-s2 | after | 79.993 | 80.692 | 80.123 |
| r1-s3 | before | 158.801 | 159.598 | 158.823 |
| r2-s0 | before | 158.566 | 159.082 | 158.570 |
| r2-s1 | after | 75.401 | 76.067 | 75.487 |
| r2-s2 | after | 86.405 | 86.832 | 86.475 |
| r2-s3 | before | 165.535 | 165.923 | 165.541 |
| r3-s0 | before | 159.009 | 159.533 | 159.025 |
| r3-s1 | after | 86.354 | 86.708 | 86.358 |
| r3-s2 | after | 93.649 | 93.862 | 93.643 |
| r3-s3 | before | 153.704 | 154.035 | 153.694 |

`native` `pptx_cross_copy_media_rich_lifecycle`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 200.927 | 201.534 | 200.927 |
| r0-s1 | after | 109.032 | 109.321 | 109.013 |
| r0-s2 | after | 101.536 | 101.837 | 101.549 |
| r0-s3 | before | 180.773 | 181.199 | 180.744 |
| r1-s0 | before | 180.783 | 181.321 | 180.729 |
| r1-s1 | after | 102.266 | 109.200 | 103.931 |
| r1-s2 | after | 104.065 | 104.740 | 104.164 |
| r1-s3 | before | 193.925 | 194.521 | 193.982 |
| r2-s0 | before | 179.860 | 180.498 | 179.942 |
| r2-s1 | after | 115.669 | 116.046 | 115.694 |
| r2-s2 | after | 102.351 | 102.829 | 102.370 |
| r2-s3 | before | 194.087 | 194.761 | 194.240 |
| r3-s0 | before | 175.780 | 176.186 | 175.777 |
| r3-s1 | after | 101.736 | 102.131 | 101.809 |
| r3-s2 | after | 102.169 | 102.640 | 102.240 |
| r3-s3 | before | 174.664 | 175.266 | 174.723 |

`native` `pptx_cross_copy_plain`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 7.060 | 7.246 | 7.068 |
| r0-s1 | after | 6.879 | 7.020 | 6.873 |
| r0-s2 | after | 6.935 | 7.069 | 6.937 |
| r0-s3 | before | 7.090 | 7.282 | 7.117 |
| r1-s0 | before | 7.155 | 7.275 | 7.156 |
| r1-s1 | after | 6.910 | 7.142 | 6.913 |
| r1-s2 | after | 6.887 | 6.997 | 6.860 |
| r1-s3 | before | 7.225 | 7.251 | 7.180 |
| r2-s0 | before | 6.960 | 7.244 | 6.999 |
| r2-s1 | after | 6.858 | 7.102 | 6.881 |
| r2-s2 | after | 6.907 | 7.050 | 6.910 |
| r2-s3 | before | 7.067 | 7.208 | 7.064 |
| r3-s0 | before | 7.103 | 7.308 | 7.093 |
| r3-s1 | after | 6.819 | 6.967 | 6.834 |
| r3-s2 | after | 6.778 | 6.914 | 6.804 |
| r3-s3 | before | 7.032 | 7.255 | 7.055 |

`native` `pptx_cross_copy_plain_lifecycle`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 8.268 | 8.365 | 8.267 |
| r0-s1 | after | 8.178 | 8.398 | 8.192 |
| r0-s2 | after | 8.238 | 8.494 | 8.274 |
| r0-s3 | before | 8.315 | 8.452 | 8.313 |
| r1-s0 | before | 8.359 | 8.557 | 8.368 |
| r1-s1 | after | 8.158 | 8.263 | 8.139 |
| r1-s2 | after | 8.065 | 8.252 | 8.095 |
| r1-s3 | before | 8.303 | 8.458 | 8.297 |
| r2-s0 | before | 8.324 | 8.400 | 8.288 |
| r2-s1 | after | 8.138 | 8.314 | 8.140 |
| r2-s2 | after | 8.186 | 8.346 | 8.153 |
| r2-s3 | before | 8.373 | 8.568 | 8.383 |
| r3-s0 | before | 8.342 | 8.568 | 8.377 |
| r3-s1 | after | 8.080 | 8.208 | 8.079 |
| r3-s2 | after | 8.104 | 8.319 | 8.126 |
| r3-s3 | before | 8.304 | 8.452 | 8.303 |

`native` `pptx_semantic_one_edit_save [pptx-semantic-large]`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 54.489 | 55.173 | 54.467 |
| r0-s1 | after | 53.936 | 57.015 | 54.439 |
| r0-s2 | after | 53.597 | 56.542 | 53.908 |
| r0-s3 | before | 53.735 | 54.508 | 53.688 |
| r1-s0 | before | 53.835 | 54.603 | 53.782 |
| r1-s1 | after | 53.281 | 54.076 | 53.252 |
| r1-s2 | after | 53.581 | 54.554 | 53.609 |
| r1-s3 | before | 53.299 | 54.553 | 53.399 |
| r2-s0 | before | 53.818 | 54.848 | 53.825 |
| r2-s1 | after | 54.146 | 54.923 | 54.167 |
| r2-s2 | after | 54.239 | 55.474 | 54.267 |
| r2-s3 | before | 55.543 | 56.497 | 55.498 |
| r3-s0 | before | 53.616 | 55.033 | 53.733 |
| r3-s1 | after | 53.998 | 54.851 | 53.993 |
| r3-s2 | after | 53.842 | 54.917 | 53.910 |
| r3-s3 | before | 54.228 | 55.354 | 54.280 |

`native` `pptx_semantic_one_edit_save [pptx-semantic-medium]`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 1.701 | 1.718 | 1.703 |
| r0-s1 | after | 1.703 | 1.717 | 1.704 |
| r0-s2 | after | 1.833 | 1.900 | 1.835 |
| r0-s3 | before | 1.683 | 1.709 | 1.686 |
| r1-s0 | before | 1.694 | 1.707 | 1.694 |
| r1-s1 | after | 1.685 | 1.704 | 1.687 |
| r1-s2 | after | 1.691 | 1.706 | 1.692 |
| r1-s3 | before | 1.682 | 1.704 | 1.684 |
| r2-s0 | before | 1.690 | 1.709 | 1.691 |
| r2-s1 | after | 1.703 | 1.760 | 1.708 |
| r2-s2 | after | 1.700 | 1.713 | 1.701 |
| r2-s3 | before | 1.699 | 1.719 | 1.700 |
| r3-s0 | before | 1.696 | 1.743 | 1.700 |
| r3-s1 | after | 1.693 | 1.707 | 1.694 |
| r3-s2 | after | 1.699 | 1.713 | 1.702 |
| r3-s3 | before | 1.691 | 1.705 | 1.692 |

`native` `pptx_semantic_one_edit_save [pptx-semantic-tiny]`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 0.906 | 0.914 | 0.906 |
| r0-s1 | after | 0.911 | 0.920 | 0.911 |
| r0-s2 | after | 0.995 | 1.039 | 1.001 |
| r0-s3 | before | 0.902 | 0.912 | 0.902 |
| r1-s0 | before | 0.904 | 0.917 | 0.904 |
| r1-s1 | after | 0.911 | 0.922 | 0.912 |
| r1-s2 | after | 0.907 | 0.915 | 0.907 |
| r1-s3 | before | 0.900 | 0.910 | 0.901 |
| r2-s0 | before | 0.904 | 0.914 | 0.904 |
| r2-s1 | after | 0.917 | 0.958 | 0.921 |
| r2-s2 | after | 0.912 | 0.920 | 0.913 |
| r2-s3 | before | 0.913 | 0.922 | 0.913 |
| r3-s0 | before | 0.906 | 0.918 | 0.908 |
| r3-s1 | after | 0.911 | 0.921 | 0.912 |
| r3-s2 | after | 0.912 | 0.920 | 0.912 |
| r3-s3 | before | 0.908 | 0.918 | 0.908 |

`native` `pptx_source_backed_cross_copy_media_rich_lifecycle`:

| process | arm | p50 ms | p95 ms | mean ms |
|---|---|---:|---:|---:|
| r0-s0 | before | 16.942 | 17.136 | 16.934 |
| r0-s1 | after | 17.101 | 17.277 | 17.099 |
| r0-s2 | after | 17.137 | 17.364 | 17.167 |
| r0-s3 | before | 17.048 | 17.209 | 17.047 |
| r1-s0 | before | 17.072 | 17.303 | 17.020 |
| r1-s1 | after | 17.089 | 17.227 | 17.085 |
| r1-s2 | after | 16.719 | 16.940 | 16.752 |
| r1-s3 | before | 17.090 | 17.276 | 17.103 |
| r2-s0 | before | 17.059 | 17.296 | 17.076 |
| r2-s1 | after | 12.139 | 12.282 | 12.138 |
| r2-s2 | after | 16.838 | 17.172 | 16.871 |
| r2-s3 | before | 17.003 | 17.140 | 16.994 |
| r3-s0 | before | 16.961 | 17.092 | 16.962 |
| r3-s1 | after | 17.048 | 17.261 | 17.010 |
| r3-s2 | after | 17.078 | 17.236 | 17.061 |
| r3-s3 | before | 17.126 | 17.390 | 17.132 |
