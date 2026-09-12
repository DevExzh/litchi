# 0519 bounded native tail guard

This diagnostic captures the two r2 elapsed p99 outliers with retained baseline and candidate binaries in eight fresh processes on CPU 2. The fixed order is A1/B1/B2/A2 over both cases: A1 baseline file+owned, B1 candidate file+owned, B2 candidate owned+file, A2 baseline owned+file. Each case therefore has two independent A/B pairs.

The protocol uses 200 measured samples, 10 warmups, and one internal repeat per process. Phase values are nearest-rank p50/p95/p99 plus arithmetic mean in nanoseconds; RSS is whole-child GNU time maximum in KiB. p50 bootstrap intervals are unpaired and descriptive only.

All original candidate-comparison flags are retained below. This bounded follow-up does not dismiss any original flag as noise or change the release decision by itself.

## Summary

- New adverse flags: 3 (tail p95/p99: 1; short phases: 3; RSS: 0).
- Original adverse flags retained: 107.
- Output identity parity: `True`; oracle/release guards true: `True`; Work parity: `True`; source-counter parity: `True`.

## Case comparison

| Case | Pair | Elapsed p50 (baseline→candidate ns) | Publish p50 (baseline→candidate ns) | RSS (baseline→candidate KiB) | New flags |
| --- | --- | ---: | ---: | ---: | ---: |
| p128-k1-file-batch | A1-B1 | 429802.000→361802.000 (-15.821%) | 135740.000→71180.000 (-47.562%) | 6520→6608 (1.350%) | 0 |
| p128-k1-file-batch | A2-B2 | 428432.000→360282.000 (-15.907%) | 136281.000→70680.000 (-48.137%) | 6584→6612 (0.425%) | 0 |
| p128-k1-owned-batch | A1-B1 | 431362.000→355141.000 (-17.670%) | 135670.000→72140.000 (-46.827%) | 7344→7332 (-0.163%) | 0 |
| p128-k1-owned-batch | A2-B2 | 421542.000→357422.000 (-15.211%) | 136901.000→71350.000 (-47.882%) | 7124→7336 (2.976%) | 3 |

## All phase, tail, and RSS statistics

### p128-k1-file-batch — A1-B1

| Metric | Stat | Baseline | Candidate | Ratio | Δ% | Flag |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| elapsed_ns | p50 | 429802.000 | 361802.000 | 0.842 | -15.821 |  |
| elapsed_ns | mean | 430083.115 | 362450.620 | 0.843 | -15.725 |  |
| elapsed_ns | p95 | 442102.000 | 368281.000 | 0.833 | -16.698 |  |
| elapsed_ns | p99 | 452272.000 | 370532.000 | 0.819 | -18.073 |  |
| open_ns | p50 | 15030.000 | 14900.000 | 0.991 | -0.865 |  |
| open_ns | mean | 15251.885 | 15086.685 | 0.989 | -1.083 |  |
| open_ns | p95 | 15640.000 | 15510.000 | 0.992 | -0.831 |  |
| open_ns | p99 | 22170.000 | 21710.000 | 0.979 | -2.075 |  |
| edit_ns | p50 | 274761.000 | 273031.000 | 0.994 | -0.630 |  |
| edit_ns | mean | 276710.645 | 274284.710 | 0.991 | -0.877 |  |
| edit_ns | p95 | 287911.000 | 280291.000 | 0.974 | -2.647 |  |
| edit_ns | p99 | 292311.000 | 282361.000 | 0.966 | -3.404 |  |
| commit_ns | p50 | 300.000 | 300.000 | 1.000 | 0.000 |  |
| commit_ns | mean | 306.000 | 299.950 | 0.980 | -1.977 |  |
| commit_ns | p95 | 330.000 | 320.000 | 0.970 | -3.030 |  |
| commit_ns | p99 | 340.000 | 340.000 | 1.000 | 0.000 |  |
| publish_ns | p50 | 135740.000 | 71180.000 | 0.524 | -47.562 |  |
| publish_ns | mean | 136849.675 | 71814.425 | 0.525 | -47.523 |  |
| publish_ns | p95 | 143351.000 | 77850.000 | 0.543 | -45.693 |  |
| publish_ns | p99 | 145101.000 | 79071.000 | 0.545 | -45.506 |  |
| drop_ns | p50 | 770.000 | 760.000 | 0.987 | -1.299 |  |
| drop_ns | mean | 769.310 | 764.700 | 0.994 | -0.599 |  |
| drop_ns | p95 | 820.000 | 820.000 | 1.000 | 0.000 |  |
| drop_ns | p99 | 860.000 | 840.000 | 0.977 | -2.326 |  |
| rss_kib | rss | 6520 | 6608 | 1.013 | 1.350 |  |

Bootstrap intervals are attached to each p50 metric in JSON; all phase values above remain the authoritative point statistics.

Flags with absolute values:
- none

Parity:
- Output identity: `True`; oracle fields: `True`; Work fields: `True`; source counters: `True`.

### p128-k1-file-batch — A2-B2

| Metric | Stat | Baseline | Candidate | Ratio | Δ% | Flag |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| elapsed_ns | p50 | 428432.000 | 360282.000 | 0.841 | -15.907 |  |
| elapsed_ns | mean | 428269.655 | 361439.340 | 0.844 | -15.605 |  |
| elapsed_ns | p95 | 439532.000 | 367872.000 | 0.837 | -16.304 |  |
| elapsed_ns | p99 | 442342.000 | 369932.000 | 0.836 | -16.370 |  |
| open_ns | p50 | 15040.000 | 14880.000 | 0.989 | -1.064 |  |
| open_ns | mean | 15184.660 | 15158.000 | 0.998 | -0.176 |  |
| open_ns | p95 | 15450.000 | 15520.000 | 1.005 | 0.453 |  |
| open_ns | p99 | 21920.000 | 22850.000 | 1.042 | 4.243 |  |
| edit_ns | p50 | 272551.000 | 272401.000 | 0.999 | -0.055 |  |
| edit_ns | mean | 274410.745 | 273667.400 | 0.997 | -0.271 |  |
| edit_ns | p95 | 285201.000 | 279901.000 | 0.981 | -1.858 |  |
| edit_ns | p99 | 287831.000 | 282531.000 | 0.982 | -1.841 |  |
| commit_ns | p50 | 310.000 | 320.000 | 1.032 | 3.226 |  |
| commit_ns | mean | 308.600 | 320.700 | 1.039 | 3.921 |  |
| commit_ns | p95 | 340.000 | 360.000 | 1.059 | 5.882 |  |
| commit_ns | p99 | 350.000 | 380.000 | 1.086 | 8.571 |  |
| publish_ns | p50 | 136281.000 | 70680.000 | 0.519 | -48.137 |  |
| publish_ns | mean | 137393.790 | 71296.025 | 0.519 | -48.108 |  |
| publish_ns | p95 | 144160.000 | 78030.000 | 0.541 | -45.873 |  |
| publish_ns | p99 | 145410.000 | 79280.000 | 0.545 | -45.478 |  |
| drop_ns | p50 | 770.000 | 760.000 | 0.987 | -1.299 |  |
| drop_ns | mean | 779.110 | 768.010 | 0.986 | -1.425 |  |
| drop_ns | p95 | 820.000 | 820.000 | 1.000 | 0.000 |  |
| drop_ns | p99 | 850.000 | 840.000 | 0.988 | -1.176 |  |
| rss_kib | rss | 6584 | 6612 | 1.004 | 0.425 |  |

Bootstrap intervals are attached to each p50 metric in JSON; all phase values above remain the authoritative point statistics.

Flags with absolute values:
- none

Parity:
- Output identity: `True`; oracle fields: `True`; Work fields: `True`; source counters: `True`.

### p128-k1-owned-batch — A1-B1

| Metric | Stat | Baseline | Candidate | Ratio | Δ% | Flag |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| elapsed_ns | p50 | 431362.000 | 355141.000 | 0.823 | -17.670 |  |
| elapsed_ns | mean | 431696.120 | 346305.845 | 0.802 | -19.780 |  |
| elapsed_ns | p95 | 437972.000 | 365461.000 | 0.834 | -16.556 |  |
| elapsed_ns | p99 | 441342.000 | 367552.000 | 0.833 | -16.719 |  |
| open_ns | p50 | 13300.000 | 13030.000 | 0.980 | -2.030 |  |
| open_ns | mean | 13747.905 | 13103.965 | 0.953 | -4.684 |  |
| open_ns | p95 | 20230.000 | 13551.000 | 0.670 | -33.015 |  |
| open_ns | p99 | 21350.000 | 20030.000 | 0.938 | -6.183 |  |
| edit_ns | p50 | 278711.000 | 269311.000 | 0.966 | -3.373 |  |
| edit_ns | mean | 279962.625 | 270814.145 | 0.967 | -3.268 |  |
| edit_ns | p95 | 286421.000 | 277811.000 | 0.970 | -3.006 |  |
| edit_ns | p99 | 287651.000 | 279841.000 | 0.973 | -2.715 |  |
| commit_ns | p50 | 320.000 | 320.000 | 1.000 | 0.000 |  |
| commit_ns | mean | 321.505 | 318.450 | 0.990 | -0.950 |  |
| commit_ns | p95 | 350.000 | 360.000 | 1.029 | 2.857 |  |
| commit_ns | p99 | 380.000 | 370.000 | 0.974 | -2.632 |  |
| publish_ns | p50 | 135670.000 | 72140.000 | 0.532 | -46.827 |  |
| publish_ns | mean | 136671.835 | 61096.030 | 0.447 | -55.297 |  |
| publish_ns | p95 | 144440.000 | 77111.000 | 0.534 | -46.614 |  |
| publish_ns | p99 | 146371.000 | 81380.000 | 0.556 | -44.402 |  |
| drop_ns | p50 | 770.000 | 750.000 | 0.974 | -2.597 |  |
| drop_ns | mean | 770.950 | 757.750 | 0.983 | -1.712 |  |
| drop_ns | p95 | 810.000 | 810.000 | 1.000 | 0.000 |  |
| drop_ns | p99 | 840.000 | 830.000 | 0.988 | -1.190 |  |
| rss_kib | rss | 7344 | 7332 | 0.998 | -0.163 |  |

Bootstrap intervals are attached to each p50 metric in JSON; all phase values above remain the authoritative point statistics.

Flags with absolute values:
- none

Parity:
- Output identity: `True`; oracle fields: `True`; Work fields: `True`; source counters: `True`.

### p128-k1-owned-batch — A2-B2

| Metric | Stat | Baseline | Candidate | Ratio | Δ% | Flag |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| elapsed_ns | p50 | 421542.000 | 357422.000 | 0.848 | -15.211 |  |
| elapsed_ns | mean | 422376.490 | 347065.445 | 0.822 | -17.830 |  |
| elapsed_ns | p95 | 430601.000 | 368772.000 | 0.856 | -14.359 |  |
| elapsed_ns | p99 | 433432.000 | 370732.000 | 0.855 | -14.466 |  |
| open_ns | p50 | 13101.000 | 13160.000 | 1.005 | 0.450 |  |
| open_ns | mean | 13311.570 | 13255.660 | 0.996 | -0.420 |  |
| open_ns | p95 | 13510.000 | 13670.000 | 1.012 | 1.184 |  |
| open_ns | p99 | 20570.000 | 22250.000 | 1.082 | 8.167 |  |
| edit_ns | p50 | 268852.000 | 271251.000 | 1.009 | 0.892 |  |
| edit_ns | mean | 270125.460 | 271893.835 | 1.007 | 0.655 |  |
| edit_ns | p95 | 277621.000 | 281412.000 | 1.014 | 1.366 |  |
| edit_ns | p99 | 280481.000 | 283251.000 | 1.010 | 0.988 |  |
| commit_ns | p50 | 310.000 | 340.000 | 1.097 | 9.677 | ADVERSE |
| commit_ns | mean | 310.750 | 342.700 | 1.103 | 10.282 | ADVERSE |
| commit_ns | p95 | 340.000 | 380.000 | 1.118 | 11.765 | ADVERSE |
| commit_ns | p99 | 370.000 | 400.000 | 1.081 | 8.108 |  |
| publish_ns | p50 | 136901.000 | 71350.000 | 0.521 | -47.882 |  |
| publish_ns | mean | 137617.000 | 60559.895 | 0.440 | -55.994 |  |
| publish_ns | p95 | 144411.000 | 76871.000 | 0.532 | -46.769 |  |
| publish_ns | p99 | 146431.000 | 81510.000 | 0.557 | -44.336 |  |
| drop_ns | p50 | 800.000 | 790.000 | 0.988 | -1.250 |  |
| drop_ns | mean | 800.260 | 793.655 | 0.992 | -0.825 |  |
| drop_ns | p95 | 850.000 | 850.000 | 1.000 | 0.000 |  |
| drop_ns | p99 | 900.000 | 860.000 | 0.956 | -4.444 |  |
| rss_kib | rss | 7124 | 7336 | 1.030 | 2.976 |  |

Bootstrap intervals are attached to each p50 metric in JSON; all phase values above remain the authoritative point statistics.

Flags with absolute values:
- `commit_ns.p50` +9.677% (310.000→340.000; threshold 5.0%).
- `commit_ns.mean` +10.282% (310.750→342.700; threshold 5.0%).
- `commit_ns.p95` +11.765% (340.000→380.000; threshold 10.0%).

Parity:
- Output identity: `True`; oracle fields: `True`; Work fields: `True`; source counters: `True`.

## Capture bindings

| Slot | Variant | Case | Binary SHA256 | Source before | Source after | RSS KiB |
| --- | --- | --- | --- | --- | --- | ---: |
| A1 | baseline | p128-k1-file-batch | `75a2679708d4ad9a119a9600683c0adc080b0a6e4a699294360ef4c459e00374` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | 6520 |
| A1 | baseline | p128-k1-owned-batch | `75a2679708d4ad9a119a9600683c0adc080b0a6e4a699294360ef4c459e00374` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | 7344 |
| B1 | candidate | p128-k1-file-batch | `48a09a4668c0e3c8a12ea6d20dd9129330f9e891193967f989f92c26625fc800` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | 6608 |
| B1 | candidate | p128-k1-owned-batch | `48a09a4668c0e3c8a12ea6d20dd9129330f9e891193967f989f92c26625fc800` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | 7332 |
| B2 | candidate | p128-k1-owned-batch | `48a09a4668c0e3c8a12ea6d20dd9129330f9e891193967f989f92c26625fc800` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | 7336 |
| B2 | candidate | p128-k1-file-batch | `48a09a4668c0e3c8a12ea6d20dd9129330f9e891193967f989f92c26625fc800` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | 6612 |
| A2 | baseline | p128-k1-owned-batch | `75a2679708d4ad9a119a9600683c0adc080b0a6e4a699294360ef4c459e00374` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | 7124 |
| A2 | baseline | p128-k1-file-batch | `75a2679708d4ad9a119a9600683c0adc080b0a6e4a699294360ef4c459e00374` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | `5e0b4f9bbe9e5c1c3808b10f5d593a395d639db6eb1023b398bc774d66bd5a93` | 6584 |

Baseline rows are bound to the retained baseline binary/build manifest without requiring the current working tree to equal the old baseline source. Candidate rows bind both current source snapshots to the retained candidate manifest.

## Original comparison flags retained

| Pair | Case | Metric | Stat | Δ% | Baseline | Candidate | Threshold |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: |
| r1 | p128-k1-file-repeated | commit_ns | p99 | 2048.649 | 370 | 7950 | 15.000 |
| r1 | p128-k1-file-repeated | commit_ns | mean | 41.664 | 309.183 | 438.000 | 5.000 |
| r1 | p128-k1-owned-batch | open_ns | p99 | 62.706 | 13171 | 21430 | 15.000 |
| r1 | p128-k1-owned-batch | commit_ns | p50 | 6.897 | 290 | 310 | 5.000 |
| r1 | p128-k1-owned-batch | commit_ns | p95 | 12.903 | 310 | 350 | 10.000 |
| r1 | p128-k1-owned-batch | commit_ns | mean | 9.209 | 293.017 | 320.000 | 5.000 |
| r1 | p128-k1-owned-repeated | commit_ns | p95 | 12.903 | 310 | 350 | 10.000 |
| r1 | p128-k32-file-batch | drop_ns | p99 | 214.213 | 3940 | 12380 | 15.000 |
| r1 | p128-k32-file-batch | drop_ns | mean | 5.492 | 2609.717 | 2753.033 | 5.000 |
| r1 | p128-k32-owned-repeated | commit_ns | p99 | 64.789 | 1420 | 2340 | 15.000 |
| r1 | p128-k32-owned-repeated | commit_ns | mean | 5.957 | 1178.167 | 1248.350 | 5.000 |
| r1 | p128-k8-file-batch | commit_ns | p50 | 5.128 | 390 | 410 | 5.000 |
| r1 | p128-k8-file-repeated | commit_ns | p99 | 1211.538 | 520 | 6820 | 15.000 |
| r1 | p128-k8-file-repeated | commit_ns | mean | 23.054 | 420.683 | 517.667 | 5.000 |
| r1 | p128-k8-owned-batch | open_ns | p99 | 31.306 | 13480 | 17700 | 15.000 |
| r1 | p128-k8-owned-batch | commit_ns | p50 | 15.000 | 400 | 460 | 5.000 |
| r1 | p128-k8-owned-batch | commit_ns | mean | 13.833 | 404.833 | 460.833 | 5.000 |
| r1 | p128-k8-owned-repeated | open_ns | p99 | 99.345 | 13901 | 27711 | 15.000 |
| r1 | p128-k8-owned-repeated | commit_ns | p50 | 9.524 | 420 | 460 | 5.000 |
| r1 | p128-k8-owned-repeated | commit_ns | p95 | 10.870 | 460 | 510 | 10.000 |
| r1 | p128-k8-owned-repeated | commit_ns | p99 | 16.000 | 500 | 580 | 15.000 |
| r1 | p128-k8-owned-repeated | commit_ns | mean | 9.359 | 423.833 | 463.500 | 5.000 |
| r1 | p128-k8-owned-repeated | drop_ns | p99 | 1401.351 | 1480 | 22220 | 15.000 |
| r1 | p128-k8-owned-repeated | drop_ns | mean | 25.035 | 1296.167 | 1620.667 | 5.000 |
| r1 | p512-k1-owned-batch | commit_ns | p50 | 6.250 | 320 | 340 | 5.000 |
| r1 | p512-k1-owned-batch | commit_ns | p95 | 25.000 | 360 | 450 | 10.000 |
| r1 | p512-k1-owned-batch | commit_ns | p99 | 37.500 | 400 | 550 | 15.000 |
| r1 | p512-k1-owned-batch | commit_ns | mean | 8.959 | 321.833 | 350.667 | 5.000 |
| r1 | p512-k1-owned-batch | drop_ns | p99 | 16.667 | 840 | 980 | 15.000 |
| r1 | p512-k1-owned-repeated | commit_ns | p95 | 24.324 | 370 | 460 | 10.000 |
| r1 | p512-k1-owned-repeated | commit_ns | p99 | 17.500 | 400 | 470 | 15.000 |
| r1 | p512-k1-owned-repeated | commit_ns | mean | 7.241 | 340.667 | 365.333 | 5.000 |
| r1 | p512-k32-file-batch | open_ns | p95 | 27.681 | 16040 | 20480 | 10.000 |
| r1 | p512-k32-file-batch | drop_ns | p99 | 51.724 | 6090 | 9240 | 15.000 |
| r1 | p512-k32-file-repeated | commit_ns | p50 | 12.295 | 1220 | 1370 | 5.000 |
| r1 | p512-k32-file-repeated | commit_ns | p95 | 37.956 | 1370 | 1890 | 10.000 |
| r1 | p512-k32-file-repeated | commit_ns | p99 | 30.857 | 1750 | 2290 | 15.000 |
| r1 | p512-k32-file-repeated | commit_ns | mean | 15.659 | 1227.167 | 1419.333 | 5.000 |
| r1 | p512-k32-owned-batch | open_ns | p50 | 21.194 | 13631 | 16520 | 5.000 |
| r1 | p512-k32-owned-repeated | open_ns | p50 | 12.717 | 14390 | 16220 | 5.000 |
| r1 | p512-k32-owned-repeated | commit_ns | p50 | 12.397 | 1210 | 1360 | 5.000 |
| r1 | p512-k32-owned-repeated | commit_ns | p95 | 12.766 | 1410 | 1590 | 10.000 |
| r1 | p512-k32-owned-repeated | commit_ns | p99 | 16.892 | 1480 | 1730 | 15.000 |
| r1 | p512-k32-owned-repeated | commit_ns | mean | 11.374 | 1230.833 | 1370.833 | 5.000 |
| r2 | p128-k1-file-batch | elapsed_ns | p99 | 63.647 | 463512 | 758523 | 15.000 |
| r2 | p128-k1-file-batch | open_ns | p99 | 24.381 | 16160 | 20100 | 15.000 |
| r2 | p128-k1-file-batch | edit_ns | p99 | 124.011 | 285201 | 638882 | 15.000 |
| r2 | p128-k1-file-batch | commit_ns | p95 | 27.273 | 330 | 420 | 10.000 |
| r2 | p128-k1-file-batch | commit_ns | p99 | 102.941 | 340 | 690 | 15.000 |
| r2 | p128-k1-file-batch | commit_ns | mean | 7.965 | 301.333 | 325.333 | 5.000 |
| r2 | p128-k1-file-batch | drop_ns | p99 | 23.256 | 860 | 1060 | 15.000 |
| r2 | p128-k1-file-repeated | commit_ns | p50 | 10.000 | 300 | 330 | 5.000 |
| r2 | p128-k1-file-repeated | commit_ns | p95 | 27.273 | 330 | 420 | 10.000 |
| r2 | p128-k1-file-repeated | commit_ns | p99 | 43.396 | 530 | 760 | 15.000 |
| r2 | p128-k1-file-repeated | commit_ns | mean | 12.473 | 308.667 | 347.167 | 5.000 |
| r2 | p128-k1-owned-batch | elapsed_ns | p99 | 775.442 | 394512 | 3453725 | 15.000 |
| r2 | p128-k1-owned-batch | open_ns | p95 | 72.464 | 13110 | 22610 | 10.000 |
| r2 | p128-k1-owned-batch | open_ns | p99 | 78.474 | 19790 | 35320 | 15.000 |
| r2 | p128-k1-owned-batch | open_ns | mean | 10.299 | 12605.267 | 13903.533 | 5.000 |
| r2 | p128-k1-owned-batch | edit_ns | p95 | 14.913 | 280831 | 322712 | 10.000 |
| r2 | p128-k1-owned-batch | edit_ns | p99 | 1077.920 | 282821 | 3331405 | 15.000 |
| r2 | p128-k1-owned-batch | edit_ns | mean | 24.781 | 275029.083 | 343185.183 | 5.000 |
| r2 | p128-k1-owned-batch | commit_ns | p50 | 16.129 | 310 | 360 | 5.000 |
| r2 | p128-k1-owned-batch | commit_ns | p95 | 388.235 | 340 | 1660 | 10.000 |
| r2 | p128-k1-owned-batch | commit_ns | p99 | 537.143 | 350 | 2230 | 15.000 |
| r2 | p128-k1-owned-batch | commit_ns | mean | 159.989 | 310.333 | 806.833 | 5.000 |
| r2 | p128-k1-owned-batch | drop_ns | p50 | 23.377 | 770 | 950 | 5.000 |
| r2 | p128-k1-owned-batch | drop_ns | p95 | 86.585 | 820 | 1530 | 10.000 |
| r2 | p128-k1-owned-batch | drop_ns | p99 | 172.941 | 850 | 2320 | 15.000 |
| r2 | p128-k1-owned-batch | drop_ns | mean | 31.672 | 775.667 | 1021.333 | 5.000 |
| r2 | p128-k1-owned-repeated | drop_ns | p99 | 51.724 | 870 | 1320 | 15.000 |
| r2 | p128-k32-file-repeated | commit_ns | p99 | 374.286 | 1400 | 6640 | 15.000 |
| r2 | p128-k32-file-repeated | commit_ns | mean | 8.590 | 1204.833 | 1308.333 | 5.000 |
| r2 | p128-k32-file-repeated | drop_ns | p99 | 36.534 | 4270 | 5830 | 15.000 |
| r2 | p128-k32-owned-batch | open_ns | p99 | 15.588 | 27970 | 32330 | 15.000 |
| r2 | p128-k32-owned-batch | open_ns | mean | 13.544 | 14639.067 | 16621.750 | 5.000 |
| r2 | p128-k32-owned-repeated | commit_ns | p95 | 30.657 | 1370 | 1790 | 10.000 |
| r2 | p128-k32-owned-repeated | commit_ns | p99 | 41.379 | 1450 | 2050 | 15.000 |
| r2 | p128-k32-owned-repeated | commit_ns | mean | 6.515 | 1215.167 | 1294.333 | 5.000 |
| r2 | p128-k8-owned-batch | open_ns | p99 | 65.987 | 13730 | 22790 | 15.000 |
| r2 | p512-k1-file-repeated | open_ns | p99 | 29.244 | 31630 | 40880 | 15.000 |
| r2 | p512-k1-file-repeated | commit_ns | p50 | 18.750 | 320 | 380 | 5.000 |
| r2 | p512-k1-file-repeated | commit_ns | p95 | 17.143 | 350 | 410 | 10.000 |
| r2 | p512-k1-file-repeated | commit_ns | p99 | 17.949 | 390 | 460 | 15.000 |
| r2 | p512-k1-file-repeated | commit_ns | mean | 17.601 | 318.167 | 374.167 | 5.000 |
| r2 | p512-k1-owned-batch | commit_ns | p50 | 12.903 | 310 | 350 | 5.000 |
| r2 | p512-k1-owned-batch | commit_ns | p95 | 32.353 | 340 | 450 | 10.000 |
| r2 | p512-k1-owned-batch | commit_ns | p99 | 65.714 | 350 | 580 | 15.000 |
| r2 | p512-k1-owned-batch | commit_ns | mean | 18.636 | 310.333 | 368.167 | 5.000 |
| r2 | p512-k32-file-batch | open_ns | p95 | 34.453 | 16251 | 21850 | 10.000 |
| r2 | p512-k32-file-batch | open_ns | p99 | 54.282 | 16930 | 26120 | 15.000 |
| r2 | p512-k32-file-repeated | open_ns | p95 | 16.538 | 16870 | 19660 | 10.000 |
| r2 | p512-k32-owned-repeated | open_ns | p99 | 20.232 | 30200 | 36310 | 15.000 |
| r2 | p512-k32-owned-repeated | commit_ns | p99 | 186.111 | 1440 | 4120 | 15.000 |
| r2 | p512-k32-owned-repeated | drop_ns | p99 | 255.592 | 3040 | 10810 | 15.000 |
| r2 | p512-k8-file-repeated | commit_ns | p95 | 11.538 | 520 | 580 | 10.000 |
| r2 | p512-k8-file-repeated | commit_ns | mean | 7.964 | 439.500 | 474.500 | 5.000 |
| r2 | p512-k8-owned-batch | commit_ns | p50 | 5.000 | 400 | 420 | 5.000 |
| r2 | p512-k8-owned-batch | commit_ns | p99 | 17.391 | 460 | 540 | 15.000 |
| r2 | p512-k8-owned-repeated | open_ns | p50 | 107.299 | 13290 | 27550 | 5.000 |
| r2 | p512-k8-owned-repeated | open_ns | p95 | 142.012 | 14710 | 35600 | 10.000 |
| r2 | p512-k8-owned-repeated | open_ns | p99 | 37.180 | 28190 | 38671 | 15.000 |
| r2 | p512-k8-owned-repeated | open_ns | mean | 105.005 | 13823.167 | 28338.233 | 5.000 |
| r2 | p512-k8-owned-repeated | commit_ns | p50 | 7.692 | 390 | 420 | 5.000 |
| r2 | p512-k8-owned-repeated | commit_ns | p95 | 11.364 | 440 | 490 | 10.000 |
| r2 | p512-k8-owned-repeated | commit_ns | p99 | 28.261 | 460 | 590 | 15.000 |
| r2 | p512-k8-owned-repeated | commit_ns | mean | 8.923 | 396.000 | 431.333 | 5.000 |

Full rows, receipts, source/binary bindings, guards, parity differences, and bootstrap metadata are retained in `tail-guard.json`.
