# Formal comparison tables

Values are signed percentage changes, averaged over six paired process blocks. Brackets contain the descriptive 95% block-bootstrap interval for that mean. These are not confidence intervals over independent in-process samples.

| Binary | Route | p50 change [interval] | p95 change [interval] | p99 change [interval] |
| --- | --- | --- | --- | --- |
| normal | deterministic | -18.90% [-25.25, -12.67] | -20.69% [-25.33, -16.45] | -20.82% [-25.86, -16.31] |
| normal | memory_store | -21.37% [-22.26, -20.58] | -19.92% [-22.47, -17.28] | -17.32% [-22.17, -13.00] |
| normal | file_store | -6.73% [-9.71, -4.41] | -12.16% [-35.44, +13.33] | -8.91% [-33.75, +15.40] |
| allocator | deterministic | -16.11% [-23.06, -9.38] | -18.93% [-26.15, -11.70] | -22.04% [-28.75, -13.56] |
| allocator | memory_store | -21.86% [-22.52, -21.28] | -19.41% [-21.33, -17.45] | -17.04% [-19.64, -14.17] |
| allocator | file_store | -2.53% [-5.58, +0.78] | -12.98% [-25.75, +1.10] | -10.86% [-28.31, +10.12] |

## Every file-store block

Values give before → after milliseconds and signed percentage change.

| Binary | Block | p50 | p95 | p99 |
| --- | ---: | --- | --- | --- |
| normal | 1 | 3.634 → 3.472 (-4.45%) | 3.841 → 3.606 (-6.13%) | 3.926 → 3.815 (-2.81%) |
| normal | 2 | 3.849 → 3.536 (-8.13%) | 8.613 → 3.705 (-56.98%) | 9.452 → 3.755 (-60.28%) |
| normal | 3 | 4.107 → 3.543 (-13.72%) | 5.790 → 3.797 (-34.41%) | 7.271 → 5.066 (-30.33%) |
| normal | 4 | 3.646 → 3.482 (-4.51%) | 3.788 → 3.744 (-1.15%) | 3.853 → 5.318 (+38.03%) |
| normal | 5 | 3.629 → 3.407 (-6.12%) | 4.284 → 3.550 (-17.13%) | 4.859 → 4.395 (-9.54%) |
| normal | 6 | 3.573 → 3.449 (-3.47%) | 3.701 → 5.286 (+42.84%) | 4.859 → 5.417 (+11.47%) |
| allocator | 1 | 3.616 → 3.455 (-4.46%) | 3.726 → 3.590 (-3.65%) | 4.354 → 3.840 (-11.81%) |
| allocator | 2 | 3.703 → 3.889 (+5.01%) | 3.803 → 4.370 (+14.89%) | 3.857 → 5.264 (+36.47%) |
| allocator | 3 | 3.786 → 3.464 (-8.50%) | 4.669 → 3.571 (-23.51%) | 4.926 → 3.690 (-25.09%) |
| allocator | 4 | 3.600 → 3.497 (-2.84%) | 3.793 → 3.597 (-5.15%) | 3.893 → 3.703 (-4.88%) |
| allocator | 5 | 3.598 → 3.489 (-3.04%) | 5.400 → 3.565 (-33.98%) | 6.231 → 3.588 (-42.42%) |
| allocator | 6 | 3.564 → 3.517 (-1.34%) | 5.406 → 3.975 (-26.47%) | 5.566 → 4.596 (-17.43%) |

## Operation heap

The median peak increment is constant across all six blocks for each route.

| Route | Before bytes | After bytes | Change |
| --- | ---: | ---: | ---: |
| deterministic | 989,190 | 792,403 | -19.89% |
| memory_store | 9,184,134 | 8,987,347 | -2.14% |
| file_store | 795,721 | 598,933 | -24.73% |

## Every adverse quantile above 5%

The raw `process_rss_bytes` metric is a **saturating before/after RSS delta**, not total resident memory. `process_peak_rss_bytes` is the after-sample VmHWM, not a delta. GNU-time maximum RSS is a separate one-observation-per-child metric; none of its block comparisons crosses +5%. Small interpolated delta baselines can yield very large percentage changes; retain the absolute byte values.

| Binary | Route | Block | Metric | Quantile | Before | After | Change |
| --- | --- | ---: | --- | --- | ---: | ---: | ---: |
| normal | file_store | 4 | elapsed_ns | p99 | 3852528.21 | 5317722.90 | +38.03% |
| normal | file_store | 6 | elapsed_ns | p95 | 3700911.00 | 5286379.50 | +42.84% |
| normal | file_store | 6 | elapsed_ns | p99 | 4859338.97 | 5416906.39 | +11.47% |
| allocator | memory_store | 2 | process_rss_bytes | p95 | 39116.80 | 374988.80 | +858.64% |
| allocator | memory_store | 2 | process_rss_bytes | p99 | 352378.88 | 528384.00 | +49.95% |
| allocator | memory_store | 3 | process_rss_bytes | p99 | 6717.44 | 235970.56 | +3412.80% |
| allocator | memory_store | 5 | process_rss_bytes | p99 | 123781.12 | 198164.48 | +60.09% |
| allocator | memory_store | 6 | process_rss_bytes | p95 | 204.80 | 49561.60 | +24100.00% |
| allocator | memory_store | 6 | process_rss_bytes | p99 | 121364.48 | 530063.36 | +336.75% |
| allocator | file_store | 2 | elapsed_ns | p50 | 3703218.00 | 3888749.00 | +5.01% |
| allocator | file_store | 2 | elapsed_ns | p95 | 3803261.00 | 4369540.55 | +14.89% |
| allocator | file_store | 2 | elapsed_ns | p99 | 3857492.90 | 5264483.51 | +36.47% |
| allocator | file_store | 2 | process_peak_rss_bytes | p50 | 6258688.00 | 6619136.00 | +5.76% |
