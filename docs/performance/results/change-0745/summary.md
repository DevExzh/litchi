## Timing: medians over 18 paired heap layouts (54 processes per case)

Process p50 in µs is the median over the 18 processes of each arm. Changes are the median of 18 paired per-layout changes, the bootstrap 95% interval of that median, and the min..max over layouts.

| Case | A p50 | B p50 | C p50 | B vs A p50 | C vs B p50 | C vs A p50 |
|---|---:|---:|---:|---|---|---|
| apply-durable-45543 | 1129 | 769 | 692 | -31.9% [-32.2, -31.6] (-37..-18) | -10.1% [-10.6, -9.7] (-24..-9) | -38.6% [-38.9, -38.4] (-43..-38) |
| chain-durable-45543 | 2025 | 2007 | 1689 | -1.6% [-9.9, +0.7] (-16..+14) | -16.2% [-16.8, -14.8] (-19..-7) | -19.9% [-21.2, -15.5] (-25..-5) |
| doc-replace-durable-floating | 1393 | 1395 | 1394 | +0.2% [-3.2, +2.1] (-8..+4) | -0.4% [-1.0, +1.8] (-4..+9) | -0.0% [-1.1, +0.9] (-9..+7) |
| doc-replace-floating | 1065 | 1061 | 1057 | -0.3% [-0.8, +1.5] (-2..+11) | -0.9% [-3.1, +0.3] (-11..+8) | -1.3% [-2.0, -0.2] (-8..+9) |
| harness-ppt-semantic/ppt_semantic_one_edit_save/tiny | 89 | 82 | 81 | -8.8% [-9.3, -8.4] (-11..-7) | -1.2% [-1.5, -0.8] (-2..+1) | -10.1% [-10.4, -9.3] (-11..-8) |
| harness-ppt-semantic/ppt_semantic_noop_edit_save/tiny | 21 | 13 | 13 | -36.5% [-36.7, -35.7] (-38..-35) | -1.0% [-2.7, +0.2] (-3..+3) | -37.0% [-37.4, -36.7] (-38..-35) |
| harness-ppt-semantic/ppt_semantic_one_edit_save/large | 241 | 200 | 184 | -17.0% [-17.4, -16.6] (-21..-14) | -7.3% [-8.2, -5.0] (-10..-4) | -23.6% [-24.1, -21.3] (-26..-19) |
| harness-ppt-semantic/ppt_semantic_noop_edit_save/large | 58 | 20 | 19 | -65.6% [-65.6, -65.4] (-66..-65) | -3.9% [-5.4, -3.2] (-7..-1) | -67.0% [-67.2, -66.8] (-68..-66) |
| hide-45543 | 1030 | 663 | 492 | -35.7% [-35.8, -35.0] (-36..-20) | -25.7% [-26.2, -25.2] (-27..-16) | -52.3% [-52.4, -51.9] (-53..-39) |
| noop-45543 | 421 | 69 | 59 | -83.7% [-83.8, -83.6] (-84..-84) | -13.7% [-14.2, -12.8] (-15..-11) | -85.9% [-86.0, -85.8] (-86..-86) |
| p0734-docfloat | 958 | 1004 | 964 | +0.1% [-0.7, +6.3] (-7..+13) | -1.5% [-5.1, +0.8] (-13..+12) | +0.1% [-2.1, +1.5] (-14..+12) |
| p0734-ppt45543 | 1047 | 666 | 595 | -36.2% [-37.0, -33.8] (-41..-33) | -10.1% [-15.0, -9.4] (-19..-4) | -42.7% [-43.4, -42.0] (-52..-40) |
| remove-41246 | 1355 | 1100 | 1074 | -19.0% [-19.2, -18.7] (-21..-18) | -2.2% [-2.9, -1.9] (-6..-1) | -20.9% [-21.2, -20.4] (-24..-20) |
| remove-45543 | 1079 | 716 | 654 | -33.8% [-42.1, -30.9] (-47..-28) | -8.9% [-10.1, +4.5] (-13..+18) | -39.5% [-39.7, -37.6] (-43..-32) |
| remove-durable-45543 | 1101 | 1101 | 1036 | +0.1% [-0.2, +1.5] (-1..+13) | -5.9% [-7.3, -5.7] (-12..-5) | -5.8% [-6.0, -5.5] (-6..-1) |

### Mean and p95 (median paired change over layouts)

| Case | B vs A mean | B vs A p95 | C vs B mean | C vs B p95 | C vs A mean | C vs A p95 |
|---|---:|---:|---:|---:|---:|---:|
| apply-durable-45543 | -32.0% | -31.5% | -11.6% | -16.3% | -39.6% | -42.0% |
| chain-durable-45543 | -3.8% | -9.8% | -14.0% | -17.2% | -19.1% | -24.7% |
| doc-replace-durable-floating | +0.2% | -0.5% | -0.5% | -0.6% | +0.5% | -0.8% |
| doc-replace-floating | -0.4% | -0.4% | -0.9% | -0.1% | -0.9% | -0.5% |
| harness-ppt-semantic/ppt_semantic_one_edit_save/tiny | -8.6% | -6.5% | -1.5% | -0.3% | -10.1% | -9.8% |
| harness-ppt-semantic/ppt_semantic_noop_edit_save/tiny | -36.4% | -35.2% | -1.9% | -1.8% | -37.4% | -36.3% |
| harness-ppt-semantic/ppt_semantic_one_edit_save/large | -17.2% | -16.4% | -6.9% | -5.6% | -23.4% | -21.8% |
| harness-ppt-semantic/ppt_semantic_noop_edit_save/large | -65.6% | -66.3% | -4.1% | -4.0% | -67.0% | -66.7% |
| hide-45543 | -35.5% | -35.6% | -21.1% | -13.3% | -49.8% | -44.3% |
| noop-45543 | -83.6% | -82.3% | -14.1% | -14.2% | -85.8% | -85.2% |
| p0734-docfloat | +1.3% | +3.5% | -3.1% | -2.2% | -0.8% | +0.9% |
| p0734-ppt45543 | -35.3% | -29.7% | -14.1% | -18.6% | -44.9% | -42.8% |
| remove-41246 | -19.0% | -18.8% | -2.1% | -2.0% | -20.8% | -20.5% |
| remove-45543 | -32.6% | -33.3% | -10.9% | -16.2% | -40.2% | -43.9% |
| remove-durable-45543 | +1.2% | +8.6% | -7.2% | -13.2% | -6.1% | -5.9% |

### Every paired change above +5% (regression flags)

| Case | Comparison | Statistic | Flagged layouts (of 18) | Largest | Median-level flag |
|---|---|---|---:|---:|---|
| apply-durable-45543 | B/A | p95 | 1 | +86.6% | no |
| apply-durable-45543 | C/A | p95 | 1 | +207.0% | no |
| apply-durable-45543 | C/B | mean | 1 | +14.6% | no |
| apply-durable-45543 | C/B | p95 | 1 | +383.5% | no |
| chain-durable-45543 | B/A | mean | 4 | +9.0% | no |
| chain-durable-45543 | B/A | p50 | 4 | +14.0% | no |
| chain-durable-45543 | C/B | p95 | 1 | +18.7% | no |
| doc-replace-durable-floating | B/A | mean | 1 | +46.6% | no |
| doc-replace-durable-floating | B/A | p95 | 1 | +73.4% | no |
| doc-replace-durable-floating | C/A | mean | 3 | +13.1% | no |
| doc-replace-durable-floating | C/A | p50 | 1 | +6.7% | no |
| doc-replace-durable-floating | C/A | p95 | 3 | +97.2% | no |
| doc-replace-durable-floating | C/B | mean | 3 | +16.3% | no |
| doc-replace-durable-floating | C/B | p50 | 2 | +9.3% | no |
| doc-replace-durable-floating | C/B | p95 | 3 | +108.3% | no |
| doc-replace-floating | B/A | mean | 2 | +13.0% | no |
| doc-replace-floating | B/A | p50 | 3 | +10.9% | no |
| doc-replace-floating | B/A | p95 | 2 | +27.4% | no |
| doc-replace-floating | C/A | mean | 2 | +31.3% | no |
| doc-replace-floating | C/A | p50 | 2 | +9.3% | no |
| doc-replace-floating | C/A | p95 | 2 | +180.7% | no |
| doc-replace-floating | C/B | mean | 2 | +25.9% | no |
| doc-replace-floating | C/B | p50 | 1 | +7.7% | no |
| doc-replace-floating | C/B | p95 | 2 | +202.1% | no |
| harness-ppt-semantic/ppt_semantic_noop_edit_save/large | C/B | p95 | 3 | +32.8% | no |
| harness-ppt-semantic/ppt_semantic_one_edit_save/tiny | B/A | mean | 1 | +79.8% | no |
| harness-ppt-semantic/ppt_semantic_one_edit_save/tiny | B/A | p95 | 1 | +6.0% | no |
| harness-ppt-semantic/ppt_semantic_one_edit_save/tiny | C/B | p95 | 2 | +8.6% | no |
| hide-45543 | C/B | mean | 1 | +14.2% | no |
| hide-45543 | C/B | p95 | 1 | +43.6% | no |
| p0734-docfloat | B/A | mean | 6 | +11.8% | no |
| p0734-docfloat | B/A | p50 | 6 | +12.6% | no |
| p0734-docfloat | B/A | p95 | 6 | +28.4% | no |
| p0734-docfloat | C/A | mean | 3 | +10.4% | no |
| p0734-docfloat | C/A | p50 | 2 | +11.8% | no |
| p0734-docfloat | C/A | p95 | 5 | +13.5% | no |
| p0734-docfloat | C/B | mean | 1 | +9.9% | no |
| p0734-docfloat | C/B | p50 | 2 | +12.0% | no |
| p0734-docfloat | C/B | p95 | 3 | +11.7% | no |
| remove-41246 | C/B | mean | 1 | +7.7% | no |
| remove-41246 | C/B | p95 | 1 | +18.0% | no |
| remove-45543 | B/A | p95 | 1 | +221.7% | no |
| remove-45543 | C/B | mean | 2 | +16.2% | no |
| remove-45543 | C/B | p50 | 4 | +18.3% | no |
| remove-45543 | C/B | p95 | 1 | +28.1% | no |
| remove-durable-45543 | B/A | mean | 3 | +13.7% | no |
| remove-durable-45543 | B/A | p50 | 2 | +12.9% | no |
| remove-durable-45543 | B/A | p95 | 11 | +18.6% | +8.6% |

## Per-owner hardware counters (perf stat, 120 minus 20 owners, three layouts)

Median per-owner value over the three layouts; instruction counts are user mode. Page faults are listed per layout.

| Case | Instructions A (M) | B − A (M) | C − B (M) | Page faults A | Page faults B | Page faults C |
|---|---:|---:|---:|---|---|---|
| remove-45543 | 8.19 | -1.99 | -0.50 | 490, 448, 468 | 447, 492, 258 | 413, 443, 442 |
| remove-41246 | 13.54 | -1.43 | -0.26 | 308, 303, 308 | 309, 309, 302 | 310, 310, 306 |
| remove-durable-45543 | 8.54 | +0.01 | -0.50 | 439, 442, 454 | 488, 460, 456 | 429, 437, 441 |
| chain-durable-45543 | 15.15 | -0.98 | -1.01 | 273, 791, 598 | 964, 459, 927 | 584, 631, 448 |
| apply-durable-45543 | 9.56 | -1.99 | -0.56 | 165, 482, 471 | 540, 446, 453 | 422, 451, 449 |
| noop-45543 | 4.74 | -1.99 | -0.06 | -1, -1, -1 | -1, -1, -1 | 0, 0, 0 |
| hide-45543 | 8.22 | -2.04 | -0.39 | 186, 576, 662 | 604, 636, 641 | 439, 484, 483 |
| doc-replace-floating | 11.61 | -0.01 | -0.00 | 404, 355, 401 | 388, 304, 401 | 240, 294, 388 |
| doc-replace-durable-floating | 13.46 | -0.00 | +0.01 | 266, 315, 370 | 266, 387, 382 | 244, 355, 381 |
| p0734-ppt45543 | 18.58 | -1.99 | -0.51 | 383, 341, 352 | 311, 332, 316 | 284, 305, 269 |
| p0734-docfloat | 64.11 | -0.01 | +0.00 | 245, -31, 587 | 50, 564, 635 | 334, 437, 36 |

## Allocation per owner (counting allocator, deterministic across owners)

| Case | Allocated bytes A | B − A | C − B | Calls A | B − A | C − B | Peak live A | B − A | C − B | Retained A | B − A | C − B |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| remove-45543 | 11,391,768 | -32 | -2,349,812 | 5,658 | +0 | -211 | 1,992,885 | +48 | +0 | 390,144 | +0 | +0 |
| remove-41246 | 10,444,631 | -32 | -727,307 | 18,662 | +0 | -208 | 1,976,223 | +48 | +0 | 285,184 | +0 | +0 |
| remove-durable-45543 | 11,601,950 | +96 | -2,349,812 | 5,939 | +2 | -211 | 1,992,885 | +48 | +0 | 398,495 | +0 | +0 |
| chain-durable-45543 | 21,146,898 | +80 | -4,736,435 | 10,625 | +2 | -422 | 2,475,857 | +96 | +0 | 412,478 | +0 | +0 |
| apply-durable-45543 | 10,400,466 | -16 | -2,673,880 | 4,705 | +0 | -228 | 1,599,589 | +64 | +0 | 390,144 | +64 | +0 |
| noop-45543 | 1,348,321 | -128 | -324,068 | 1,373 | -2 | -17 | 703,148 | +0 | -260,671 | 385,024 | +0 | +0 |
| hide-45543 | 10,382,391 | -80 | -2,271,918 | 3,963 | -1 | -104 | 1,555,090 | +0 | +0 | 397,312 | +0 | +0 |
| doc-replace-floating | 14,115,027 | +0 | +0 | 15,141 | +0 | +0 | 3,128,190 | +0 | +0 | 342,528 | +0 | +0 |
| doc-replace-durable-floating | 14,153,167 | +0 | +0 | 15,362 | +0 | +0 | 3,128,190 | +0 | +0 | 343,097 | +0 | +0 |
