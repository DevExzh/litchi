# 0422 allocator correctness diagnostic

Four fresh allocator processes (two selectors, two repeats), 30 samples and 3 warmups each. Elapsed samples are deliberately omitted; this evidence makes no performance claim. Allocation sizes and live/peak columns are bytes.

| Repeat | Selector | Allocation calls mean | Allocated bytes mean | Live-before mean | Live-after mean | Peak-before mean | Peak-after mean | Region peak mean | Whole-process RSS KiB |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|
| R1 | pptx_cross_copy_media_rich_lifecycle | 59714.967 | 369980305.933 | 169821938.000 | 237453206.000 | 812687524.000 | 812687524.000 | 272736303.000 | 803,448 |
| R1 | pptx_cross_copy_plain_lifecycle | 49335.000 | 15664317.000 | 369171.000 | 797423.000 | 3559492.000 | 3559492.000 | 1360003.000 | 82,732 |
| R2 | pptx_cross_copy_media_rich_lifecycle | 59714.800 | 369979780.600 | 169821938.000 | 237453206.000 | 812687524.000 | 812687524.000 | 272736303.000 | 803,448 |
| R2 | pptx_cross_copy_plain_lifecycle | 49335.000 | 15664317.000 | 369171.000 | 797423.000 | 3559492.000 | 3559492.000 | 1360003.000 | 82,736 |

The JSON retains every allocation vector and each vector's count, mean, minimum, and maximum. Peak-before/live-before, peak-after/live-after, and peak-after/peak-before hold for every sample; peak high-water and inter-sample boundaries are checked after restoring chronological sample order. Region peaks include entry live bytes and are bounded by entry/exit live bytes and lifetime peak at exit; they need not be monotonic across operations. RSS is whole-process GNU `time -v` evidence; no elapsed-time or allocator-speed comparison is made.
