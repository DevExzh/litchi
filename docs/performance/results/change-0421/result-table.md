# 0421 allocator correctness diagnostic

Four fresh allocator processes (two selectors, two repeats), 30 samples and 3 warmups each. Elapsed samples are deliberately omitted; this evidence makes no performance claim.

| Repeat | Selector | Allocation calls mean | Allocated bytes mean | Live-before mean | Live-after mean | Peak-before mean | Peak-after mean | Whole-process RSS KiB |
|---|---|---:|---:|---:|---:|---:|---:|---:|
| R1 | pptx_cross_copy_media_rich_lifecycle | 59714.767 | 369979675.533 | 169821458.000 | 237452726.000 | 812687524.000 | 812687524.000 | 803,352 |
| R1 | pptx_cross_copy_plain_lifecycle | 49335.000 | 15664317.000 | 368691.000 | 796943.000 | 3559492.000 | 3559492.000 | 82,684 |
| R2 | pptx_cross_copy_media_rich_lifecycle | 59715.133 | 369980831.267 | 169821458.000 | 237452726.000 | 812687524.000 | 812687524.000 | 805,004 |
| R2 | pptx_cross_copy_plain_lifecycle | 49335.000 | 15664317.000 | 368691.000 | 796943.000 | 3559492.000 | 3559492.000 | 82,608 |

The JSON retains every allocation vector and each vector's count, mean, minimum, and maximum. Peak-before/live-before, peak-after/live-after, and peak-after/peak-before hold for every sample; peak high-water and inter-sample boundaries are checked after restoring chronological sample order. RSS is whole-process GNU `time -v` evidence; no elapsed-time or allocator-speed comparison is made.
