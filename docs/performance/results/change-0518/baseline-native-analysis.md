# 0518 native DOCX analysis

Native campaigns: r1, r2. Rows are source-bound through each receipt; profile and hardware lanes are excluded.
Measured rows are the 30 `warmup=false` samples in each of two internal repeats. Every case reports nearest-rank p50/p95/p99 and arithmetic mean for elapsed, open, edit, commit, publish, and drop clocks; RSS is the one whole-child maximum from GNU time.
The unpaired route bootstrap for the p50 ratio uses seed `5343761` and 4000 iterations, resampling each route within each internal repeat and taking batch median / repeated median. With only two internal repeats, its interval is descriptive for these row distributions and does not estimate independent-process or host variation.
The historical 0500 validator is imported by `capture.py` and remains the row-level source for release, output, semantic, untouched-member, source-version, and readback guards. Work and retained live gauges are reported as descriptive counters; they do not relax those guards.

## API-choice ratios

Ratios are batch / repeated; positive percentages mean batch took longer or used more RSS. The same absolute thresholds are used for route flags and cross-campaign flags: 5% for p50/mean/RSS, 10% for p95, and 15% for p99.

| Campaign | Workload | elapsed p50 | edit p50 | publish p50 | RSS | route flags |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| r1 | p128-k1-file | 0.59% | -0.28% | -0.26% | 1.29% | 3 |
| r1 | p128-k1-owned | 0.27% | -0.56% | 0.39% | 1.75% | 1 |
| r1 | p128-k32-file | -86.48% | -92.05% | -1.26% | -0.66% | 12 |
| r1 | p128-k32-owned | -86.72% | -91.98% | -1.01% | -0.84% | 14 |
| r1 | p128-k8-file | -59.34% | -73.62% | -12.43% | 3.58% | 19 |
| r1 | p128-k8-owned | -61.49% | -74.20% | -15.26% | 3.15% | 17 |
| r1 | p512-k1-file | 0.79% | 0.47% | 1.04% | -2.42% | 0 |
| r1 | p512-k1-owned | 2.30% | 4.42% | 0.09% | -2.07% | 5 |
| r1 | p512-k32-file | -87.18% | -92.51% | 1.28% | -0.46% | 10 |
| r1 | p512-k32-owned | -87.15% | -92.51% | 0.17% | -0.70% | 13 |
| r1 | p512-k8-file | -61.57% | -74.62% | -6.01% | -0.82% | 18 |
| r1 | p512-k8-owned | -59.00% | -73.13% | 1.10% | 0.11% | 9 |
| r2 | p128-k1-file | -0.33% | 0.35% | -1.00% | 0.80% | 1 |
| r2 | p128-k1-owned | -1.29% | -1.05% | -1.19% | -0.50% | 0 |
| r2 | p128-k32-file | -86.46% | -92.06% | -0.43% | -1.55% | 13 |
| r2 | p128-k32-owned | -86.72% | -91.99% | -1.80% | 0.89% | 13 |
| r2 | p128-k8-file | -59.73% | -73.71% | -13.32% | 3.62% | 19 |
| r2 | p128-k8-owned | -61.50% | -74.00% | -15.62% | 2.93% | 18 |
| r2 | p512-k1-file | -0.74% | -0.94% | -0.30% | 0.19% | 2 |
| r2 | p512-k1-owned | 0.61% | 0.19% | 0.69% | 1.95% | 2 |
| r2 | p512-k32-file | -87.31% | -92.60% | -0.25% | -0.17% | 12 |
| r2 | p512-k32-owned | -87.37% | -92.62% | -1.78% | -1.13% | 12 |
| r2 | p512-k8-file | -59.93% | -74.10% | 0.32% | 1.82% | 18 |
| r2 | p512-k8-owned | -59.81% | -73.94% | 0.91% | -1.37% | 15 |

## API-choice flags

22 route records have one or more absolute threshold flags.

| Campaign | Workload | flags |
| --- | --- | --- |
| r1 | p128-k1-file | open_ns.p99 -27.99% (>15%); commit_ns.p99 +1315.38% (>15%); commit_ns.mean +32.25% (>5%) |
| r1 | p128-k1-owned | commit_ns.p50 -8.82% (>5%) |
| r1 | p128-k32-file | elapsed_ns.p50 -86.48% (>5%); elapsed_ns.p95 -86.40% (>10%); elapsed_ns.p99 -86.19% (>15%); elapsed_ns.mean -86.51% (>5%); open_ns.p99 +31.94% (>15%); edit_ns.p50 -92.05% (>5%); edit_ns.p95 -91.95% (>10%); edit_ns.p99 -91.94% (>15%); edit_ns.mean -92.04% (>5%); commit_ns.p50 -5.51% (>5%); commit_ns.mean -6.34% (>5%); drop_ns.p99 -28.71% (>15%) |
| r1 | p128-k32-owned | elapsed_ns.p50 -86.72% (>5%); elapsed_ns.p95 -86.50% (>10%); elapsed_ns.p99 -86.66% (>15%); elapsed_ns.mean -86.69% (>5%); open_ns.p95 +93.02% (>10%); open_ns.mean +16.87% (>5%); edit_ns.p50 -91.98% (>5%); edit_ns.p95 -91.88% (>10%); edit_ns.p99 -92.08% (>15%); edit_ns.mean -91.98% (>5%); publish_ns.p99 -21.13% (>15%); drop_ns.p50 -7.38% (>5%); drop_ns.p99 +22.67% (>15%); drop_ns.mean -9.36% (>5%) |
| r1 | p128-k8-file | elapsed_ns.p50 -59.34% (>5%); elapsed_ns.p95 -58.47% (>10%); elapsed_ns.p99 -57.56% (>15%); elapsed_ns.mean -59.16% (>5%); open_ns.p50 -7.21% (>5%); open_ns.p99 +27.67% (>15%); open_ns.mean -6.01% (>5%); edit_ns.p50 -73.62% (>5%); edit_ns.p95 -73.13% (>10%); edit_ns.p99 -73.22% (>15%); edit_ns.mean -73.49% (>5%); commit_ns.p50 -11.11% (>5%); commit_ns.p95 -14.00% (>10%); commit_ns.mean -12.86% (>5%); publish_ns.p50 -12.43% (>5%); publish_ns.mean -10.91% (>5%); drop_ns.p50 -9.92% (>5%); drop_ns.p95 -10.71% (>10%); drop_ns.mean -9.43% (>5%) |
| r1 | p128-k8-owned | elapsed_ns.p50 -61.49% (>5%); elapsed_ns.p95 -61.20% (>10%); elapsed_ns.p99 -60.66% (>15%); elapsed_ns.mean -61.40% (>5%); open_ns.p50 -5.77% (>5%); open_ns.p99 +47.69% (>15%); edit_ns.p50 -74.20% (>5%); edit_ns.p95 -73.64% (>10%); edit_ns.p99 -73.42% (>15%); edit_ns.mean -74.06% (>5%); commit_ns.p50 -7.14% (>5%); commit_ns.mean -5.15% (>5%); publish_ns.p50 -15.26% (>5%); publish_ns.p95 -12.53% (>10%); publish_ns.mean -13.06% (>5%); drop_ns.p50 -5.69% (>5%); drop_ns.mean -5.16% (>5%) |
| r1 | p512-k1-owned | open_ns.p95 -17.92% (>10%); commit_ns.p50 +9.09% (>5%); commit_ns.p95 +13.51% (>10%); commit_ns.p99 +15.79% (>15%); commit_ns.mean +9.66% (>5%) |
| r1 | p512-k32-file | elapsed_ns.p50 -87.18% (>5%); elapsed_ns.p95 -87.06% (>10%); elapsed_ns.p99 -87.08% (>15%); elapsed_ns.mean -87.20% (>5%); edit_ns.p50 -92.51% (>5%); edit_ns.p95 -92.44% (>10%); edit_ns.p99 -92.48% (>15%); edit_ns.mean -92.51% (>5%); drop_ns.p50 -7.37% (>5%); drop_ns.mean -7.17% (>5%) |
| r1 | p512-k32-owned | elapsed_ns.p50 -87.15% (>5%); elapsed_ns.p95 -87.11% (>10%); elapsed_ns.p99 -87.17% (>15%); elapsed_ns.mean -87.16% (>5%); open_ns.p50 +36.30% (>5%); edit_ns.p50 -92.51% (>5%); edit_ns.p95 -92.45% (>10%); edit_ns.p99 -92.51% (>15%); edit_ns.mean -92.50% (>5%); commit_ns.p50 -6.15% (>5%); commit_ns.mean -5.05% (>5%); drop_ns.p50 -8.16% (>5%); drop_ns.mean -8.19% (>5%) |
| r1 | p512-k8-file | elapsed_ns.p50 -61.57% (>5%); elapsed_ns.p95 -61.78% (>10%); elapsed_ns.p99 -62.35% (>15%); elapsed_ns.mean -61.36% (>5%); open_ns.p50 -49.43% (>5%); open_ns.p95 -44.80% (>10%); open_ns.mean -48.04% (>5%); edit_ns.p50 -74.62% (>5%); edit_ns.p95 -74.79% (>10%); edit_ns.p99 -74.88% (>15%); edit_ns.mean -74.51% (>5%); commit_ns.p50 -14.58% (>5%); commit_ns.p95 -14.81% (>10%); commit_ns.mean -14.89% (>5%); publish_ns.p50 -6.01% (>5%); publish_ns.mean -5.97% (>5%); drop_ns.p50 -7.58% (>5%); drop_ns.p99 +397.32% (>15%) |
| r1 | p512-k8-owned | elapsed_ns.p50 -59.00% (>5%); elapsed_ns.p95 -58.91% (>10%); elapsed_ns.p99 -58.95% (>15%); elapsed_ns.mean -59.54% (>5%); edit_ns.p50 -73.13% (>5%); edit_ns.p95 -72.89% (>10%); edit_ns.p99 -73.03% (>15%); edit_ns.mean -73.57% (>5%); commit_ns.p99 -17.81% (>15%) |
| r2 | p128-k1-file | open_ns.p99 -34.56% (>15%) |
| r2 | p128-k32-file | elapsed_ns.p50 -86.46% (>5%); elapsed_ns.p95 -86.32% (>10%); elapsed_ns.p99 -86.18% (>15%); elapsed_ns.mean -86.60% (>5%); open_ns.p99 -31.98% (>15%); open_ns.mean -5.09% (>5%); edit_ns.p50 -92.06% (>5%); edit_ns.p95 -91.90% (>10%); edit_ns.p99 -91.75% (>15%); edit_ns.mean -92.05% (>5%); commit_ns.p50 +5.88% (>5%); commit_ns.mean +5.20% (>5%); drop_ns.p99 +220.47% (>15%) |
| r2 | p128-k32-owned | elapsed_ns.p50 -86.72% (>5%); elapsed_ns.p95 -86.35% (>10%); elapsed_ns.p99 -86.66% (>15%); elapsed_ns.mean -86.68% (>5%); open_ns.p50 -6.02% (>5%); edit_ns.p50 -91.99% (>5%); edit_ns.p95 -91.83% (>10%); edit_ns.p99 -91.83% (>15%); edit_ns.mean -91.98% (>5%); commit_ns.p99 -17.98% (>15%); publish_ns.p99 -24.98% (>15%); drop_ns.p50 -6.99% (>5%); drop_ns.p99 +341.81% (>15%) |
| r2 | p128-k8-file | elapsed_ns.p50 -59.73% (>5%); elapsed_ns.p95 -57.59% (>10%); elapsed_ns.p99 -56.40% (>15%); elapsed_ns.mean -59.45% (>5%); open_ns.p50 -7.97% (>5%); open_ns.mean -7.62% (>5%); edit_ns.p50 -73.71% (>5%); edit_ns.p95 -73.35% (>10%); edit_ns.p99 -72.38% (>15%); edit_ns.mean -73.63% (>5%); commit_ns.p50 -14.89% (>5%); commit_ns.p95 -13.21% (>10%); commit_ns.p99 -18.29% (>15%); commit_ns.mean -14.05% (>5%); publish_ns.p50 -13.32% (>5%); publish_ns.mean -11.65% (>5%); drop_ns.p50 -11.36% (>5%); drop_ns.p95 -10.64% (>10%); drop_ns.mean -10.91% (>5%) |
| r2 | p128-k8-owned | elapsed_ns.p50 -61.50% (>5%); elapsed_ns.p95 -60.94% (>10%); elapsed_ns.p99 -60.45% (>15%); elapsed_ns.mean -61.33% (>5%); open_ns.p50 -7.55% (>5%); open_ns.p99 +108.72% (>15%); open_ns.mean -5.27% (>5%); edit_ns.p50 -74.00% (>5%); edit_ns.p95 -73.32% (>10%); edit_ns.p99 -73.11% (>15%); edit_ns.mean -73.89% (>5%); commit_ns.p50 -9.30% (>5%); commit_ns.mean -9.99% (>5%); publish_ns.p50 -15.62% (>5%); publish_ns.p95 -14.54% (>10%); publish_ns.mean -13.99% (>5%); drop_ns.p50 -9.38% (>5%); drop_ns.mean -8.34% (>5%) |
| r2 | p512-k1-file | commit_ns.p99 +19.05% (>15%); drop_ns.p99 -17.54% (>15%) |
| r2 | p512-k1-owned | open_ns.p95 +25.18% (>10%); open_ns.mean +21.17% (>5%) |
| r2 | p512-k32-file | elapsed_ns.p50 -87.31% (>5%); elapsed_ns.p95 -87.22% (>10%); elapsed_ns.p99 -87.28% (>15%); elapsed_ns.mean -87.37% (>5%); open_ns.p99 +28.82% (>15%); edit_ns.p50 -92.60% (>5%); edit_ns.p95 -92.56% (>10%); edit_ns.p99 -92.57% (>15%); edit_ns.mean -92.60% (>5%); commit_ns.p99 +239.44% (>15%); commit_ns.mean +5.21% (>5%); drop_ns.p99 +19.93% (>15%) |
| r2 | p512-k32-owned | elapsed_ns.p50 -87.37% (>5%); elapsed_ns.p95 -87.06% (>10%); elapsed_ns.p99 -87.01% (>15%); elapsed_ns.mean -87.31% (>5%); open_ns.p50 -21.28% (>5%); open_ns.p99 -25.35% (>15%); edit_ns.p50 -92.62% (>5%); edit_ns.p95 -92.47% (>10%); edit_ns.p99 -92.42% (>15%); edit_ns.mean -92.59% (>5%); drop_ns.p50 -6.71% (>5%); drop_ns.p99 +28.54% (>15%) |
| r2 | p512-k8-file | elapsed_ns.p50 -59.93% (>5%); elapsed_ns.p95 -59.78% (>10%); elapsed_ns.p99 -60.43% (>15%); elapsed_ns.mean -60.05% (>5%); open_ns.p50 -46.44% (>5%); open_ns.p95 -47.32% (>10%); open_ns.p99 -28.50% (>15%); open_ns.mean -45.16% (>5%); edit_ns.p50 -74.10% (>5%); edit_ns.p95 -74.00% (>10%); edit_ns.p99 -74.18% (>15%); edit_ns.mean -74.08% (>5%); commit_ns.p50 -12.24% (>5%); commit_ns.p95 -12.73% (>10%); commit_ns.p99 -29.11% (>15%); commit_ns.mean -13.75% (>5%); drop_ns.p99 -22.60% (>15%); drop_ns.mean -5.13% (>5%) |
| r2 | p512-k8-owned | elapsed_ns.p50 -59.81% (>5%); elapsed_ns.p95 -59.97% (>10%); elapsed_ns.p99 -59.39% (>15%); elapsed_ns.mean -59.71% (>5%); open_ns.mean +28.82% (>5%); edit_ns.p50 -73.94% (>5%); edit_ns.p95 -73.82% (>10%); edit_ns.p99 -73.50% (>15%); edit_ns.mean -73.90% (>5%); commit_ns.p50 -8.51% (>5%); commit_ns.p99 -31.47% (>15%); commit_ns.mean -13.01% (>5%); drop_ns.p95 -11.11% (>10%); drop_ns.p99 -86.17% (>15%); drop_ns.mean -14.20% (>5%) |

## Per-case gauges and guards

The table exposes the Work charge, retained live reservations, release/output/input counters, and source reads for every case. Every listed guard is true; Work and live-gauge deltas remain eligible for explicit candidate comparison.

| Campaign | Workload | API | after Work | live Work | live memory | live objects | after input | after output | source reads | guards |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| r1 | p128-k1-file-batch | batch | 13318855 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| r1 | p128-k1-file-repeated | repeated | 13318855 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| r1 | p128-k1-owned-batch | batch | 13318855 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| r1 | p128-k1-owned-repeated | repeated | 13318855 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| r1 | p128-k32-file-batch | batch | 13773222 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| r1 | p128-k32-file-repeated | repeated | 19737002 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| r1 | p128-k32-owned-batch | batch | 13773222 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| r1 | p128-k32-owned-repeated | repeated | 19737002 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| r1 | p128-k8-file-batch | batch | 13421454 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| r1 | p128-k8-file-repeated | repeated | 14102162 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| r1 | p128-k8-owned-batch | batch | 13421454 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| r1 | p128-k8-owned-repeated | repeated | 14102162 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| r1 | p512-k1-file-batch | batch | 207360199 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| r1 | p512-k1-file-repeated | repeated | 207360199 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| r1 | p512-k1-owned-batch | batch | 207360199 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| r1 | p512-k1-owned-repeated | repeated | 207360199 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| r1 | p512-k32-file-batch | batch | 207814566 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| r1 | p512-k32-file-repeated | repeated | 219777962 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| r1 | p512-k32-owned-batch | batch | 207814566 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| r1 | p512-k32-owned-repeated | repeated | 219777962 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| r1 | p512-k8-file-batch | batch | 207462798 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| r1 | p512-k8-file-repeated | repeated | 209498258 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |
| r1 | p512-k8-owned-batch | batch | 207462798 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| r1 | p512-k8-owned-repeated | repeated | 209498258 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |
| r2 | p128-k1-file-batch | batch | 13318855 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| r2 | p128-k1-file-repeated | repeated | 13318855 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| r2 | p128-k1-owned-batch | batch | 13318855 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| r2 | p128-k1-owned-repeated | repeated | 13318855 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| r2 | p128-k32-file-batch | batch | 13773222 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| r2 | p128-k32-file-repeated | repeated | 19737002 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| r2 | p128-k32-owned-batch | batch | 13773222 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| r2 | p128-k32-owned-repeated | repeated | 19737002 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| r2 | p128-k8-file-batch | batch | 13421454 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| r2 | p128-k8-file-repeated | repeated | 14102162 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| r2 | p128-k8-owned-batch | batch | 13421454 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| r2 | p128-k8-owned-repeated | repeated | 14102162 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| r2 | p512-k1-file-batch | batch | 207360199 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| r2 | p512-k1-file-repeated | repeated | 207360199 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| r2 | p512-k1-owned-batch | batch | 207360199 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| r2 | p512-k1-owned-repeated | repeated | 207360199 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| r2 | p512-k32-file-batch | batch | 207814566 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| r2 | p512-k32-file-repeated | repeated | 219777962 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| r2 | p512-k32-owned-batch | batch | 207814566 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| r2 | p512-k32-owned-repeated | repeated | 219777962 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| r2 | p512-k8-file-batch | batch | 207462798 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| r2 | p512-k8-file-repeated | repeated | 209498258 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |
| r2 | p512-k8-owned-batch | batch | 207462798 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| r2 | p512-k8-owned-repeated | repeated | 209498258 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |

## Cross-campaign flags

### Matched same API

22 of 24 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-file-batch | commit_ns.p99 -95.11% (>15%); commit_ns.mean -27.57% (>5%) |
| p128-k1-file-repeated | commit_ns.p99 -25.00% (>15%) |
| p128-k1-owned-batch | elapsed_ns.p50 -8.52% (>5%); elapsed_ns.mean -7.68% (>5%); publish_ns.p50 -15.63% (>5%); publish_ns.p95 -16.38% (>10%); publish_ns.p99 -16.52% (>15%); publish_ns.mean -14.68% (>5%) |
| p128-k1-owned-repeated | elapsed_ns.p50 -7.08% (>5%); elapsed_ns.mean -6.47% (>5%); commit_ns.p50 -11.76% (>5%); commit_ns.p95 -10.81% (>10%); commit_ns.p99 -15.00% (>15%); commit_ns.mean -9.65% (>5%); publish_ns.p50 -14.27% (>5%); publish_ns.p95 -14.15% (>10%); publish_ns.mean -12.90% (>5%) |
| p128-k32-file-batch | open_ns.p99 -15.67% (>15%); commit_ns.p50 +5.00% (>5%); commit_ns.mean +5.76% (>5%); drop_ns.p99 +261.72% (>15%); drop_ns.mean +5.04% (>5%) |
| p128-k32-file-repeated | open_ns.p99 +63.59% (>15%); commit_ns.p50 -6.30% (>5%); commit_ns.mean -5.84% (>5%); drop_ns.p99 -19.53% (>15%) |
| p128-k32-owned-batch | drop_ns.p99 +25.21% (>15%) |
| p128-k32-owned-repeated | open_ns.p95 +98.05% (>10%); open_ns.mean +20.08% (>5%); commit_ns.p99 +26.24% (>15%); drop_ns.p99 -65.23% (>15%); drop_ns.mean -7.03% (>5%) |
| p128-k8-file-batch | commit_ns.p99 +34.00% (>15%) |
| p128-k8-file-repeated | open_ns.p99 +38.51% (>15%); commit_ns.p99 +49.09% (>15%) |
| p128-k8-owned-batch | open_ns.p99 +46.55% (>15%); commit_ns.p99 -17.86% (>15%) |
| p512-k1-file-repeated | drop_ns.p99 +29.55% (>15%) |
| p512-k1-owned-batch | open_ns.p95 +22.28% (>10%) |
| p512-k1-owned-repeated | open_ns.p95 -19.82% (>10%); open_ns.mean -16.82% (>5%); commit_ns.p50 +9.09% (>5%); commit_ns.p95 +10.81% (>10%); commit_ns.p99 +21.05% (>15%); commit_ns.mean +6.17% (>5%) |
| p512-k32-file-batch | commit_ns.p99 +324.31% (>15%); commit_ns.mean +7.03% (>5%); drop_ns.p99 -60.96% (>15%) |
| p512-k32-file-repeated | open_ns.p99 -29.79% (>15%); commit_ns.p99 +20.00% (>15%); drop_ns.p99 -69.79% (>15%); drop_ns.mean -5.09% (>5%) |
| p512-k32-owned-batch | open_ns.p50 -31.65% (>5%); commit_ns.p95 +10.45% (>10%); drop_ns.p99 +309.44% (>15%); drop_ns.mean +8.01% (>5%) |
| p512-k32-owned-repeated | open_ns.p50 +18.35% (>5%); open_ns.p99 +32.06% (>15%); commit_ns.p99 +20.41% (>15%); drop_ns.p99 +196.74% (>15%) |
| p512-k8-file-batch | open_ns.p50 +5.72% (>5%); commit_ns.mean +5.65% (>5%); drop_ns.p99 -81.51% (>15%); drop_ns.mean -6.95% (>5%) |
| p512-k8-file-repeated | open_ns.p99 +23.81% (>15%); commit_ns.p99 +33.90% (>15%); drop_ns.p99 +18.79% (>15%) |
| p512-k8-owned-batch | open_ns.p99 -22.44% (>15%); commit_ns.p50 -6.52% (>5%); commit_ns.p99 +1025.00% (>15%); commit_ns.mean +18.27% (>5%) |
| p512-k8-owned-repeated | open_ns.p99 -20.67% (>15%); open_ns.mean -22.70% (>5%); commit_ns.p99 +1249.32% (>15%); commit_ns.mean +31.63% (>5%); drop_ns.p99 +587.14% (>15%); drop_ns.mean +10.62% (>5%) |

### API-choice ratio

12 of 12 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-file | commit_ns.p99 -93.48% (>15%); commit_ns.mean -25.21% (>5%) |
| p128-k1-owned | commit_ns.p50 +9.68% (>5%); commit_ns.mean +6.65% (>5%) |
| p128-k32-file | open_ns.p99 -48.45% (>15%); commit_ns.p50 +12.06% (>5%); commit_ns.p95 +10.72% (>10%); commit_ns.p99 +21.27% (>15%); commit_ns.mean +12.32% (>5%); drop_ns.p99 +349.50% (>15%) |
| p128-k32-owned | open_ns.p95 -50.04% (>10%); open_ns.mean -15.78% (>5%); commit_ns.p99 -18.56% (>15%); drop_ns.p99 +260.15% (>15%); drop_ns.mean +9.81% (>5%) |
| p128-k8-file | open_ns.p99 -24.58% (>15%) |
| p128-k8-owned | open_ns.p99 +41.32% (>15%); commit_ns.p99 -17.86% (>15%); commit_ns.mean -5.10% (>5%) |
| p512-k1-file | commit_ns.p99 +19.05% (>15%); drop_ns.p99 -21.13% (>15%) |
| p512-k1-owned | open_ns.p95 +52.50% (>10%); open_ns.mean +23.34% (>5%); commit_ns.p50 -8.33% (>5%); commit_ns.p95 -14.05% (>10%); commit_ns.p99 -23.02% (>15%); commit_ns.mean -7.87% (>5%) |
| p512-k32-file | open_ns.p99 +37.75% (>15%); commit_ns.p99 +253.59% (>15%); commit_ns.mean +8.60% (>5%); drop_ns.p50 +5.26% (>5%); drop_ns.p99 +29.25% (>15%); drop_ns.mean +5.54% (>5%) |
| p512-k32-owned | open_ns.p50 -42.25% (>5%); open_ns.p99 -25.25% (>15%); drop_ns.p99 +37.98% (>15%) |
| p512-k8-file | open_ns.p50 +5.93% (>5%); open_ns.p99 -21.46% (>15%); open_ns.mean +5.54% (>5%); commit_ns.p99 -31.44% (>15%); publish_ns.p50 +6.74% (>5%); publish_ns.mean +5.30% (>5%); drop_ns.p99 -84.44% (>15%) |
| p512-k8-owned | open_ns.mean +27.45% (>5%); commit_ns.p50 -6.52% (>5%); commit_ns.p99 -16.62% (>15%); commit_ns.mean -10.15% (>5%); drop_ns.p95 -10.45% (>10%); drop_ns.p99 -86.83% (>15%); drop_ns.mean -11.29% (>5%) |


## Counter drift

Same-API output identities match for 24/24 cases across campaigns.
Release/output/read guard counter changes: 0 cases.
Allowed Work/live-gauge changes: 0 cases; any values appear in JSON.

All case/campaign phase statistics, repeat groups, route ratios, counter drift, identities, guards, and bootstrap intervals are in `baseline-native-analysis.json`.
