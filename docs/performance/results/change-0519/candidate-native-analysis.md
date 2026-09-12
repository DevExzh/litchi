# 0519 native DOCX analysis

Native campaigns: after-r1, after-r2. Rows are source-bound through each receipt; profile and hardware lanes are excluded.
Measured rows are the 30 `warmup=false` samples in each of two internal repeats. Every case reports nearest-rank p50/p95/p99 and arithmetic mean for elapsed, open, edit, commit, publish, and drop clocks; RSS is the one whole-child maximum from GNU time.
The unpaired route bootstrap for the p50 ratio uses seed `5347857` and 4000 iterations, resampling each route within each internal repeat and taking batch median / repeated median. With only two internal repeats, its interval is descriptive for these row distributions and does not estimate independent-process or host variation.
The historical 0500 validator is imported by `capture.py` and remains the row-level source for release, output, semantic, untouched-member, source-version, and readback guards. Work is expected to remain unchanged in this batch; retained live gauges are reported as descriptive counters and do not relax those guards.

## API-choice ratios

Ratios are batch / repeated; positive percentages mean batch took longer or used more RSS. The same absolute thresholds are used for route flags and cross-campaign flags: 5% for p50/mean/RSS, 10% for p95, and 15% for p99.

| Campaign | Workload | elapsed p50 | edit p50 | publish p50 | RSS | route flags |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| after-r1 | p128-k1-file | 0.62% | 0.61% | -0.78% | -1.57% | 2 |
| after-r1 | p128-k1-owned | -0.01% | 0.14% | -3.80% | -0.11% | 2 |
| after-r1 | p128-k32-file | -89.78% | -92.07% | 53.38% | -3.82% | 16 |
| after-r1 | p128-k32-owned | -90.38% | -92.00% | -4.63% | -0.39% | 14 |
| after-r1 | p128-k8-file | -68.52% | -73.91% | -0.86% | -3.13% | 13 |
| after-r1 | p128-k8-owned | -68.61% | -73.84% | -1.01% | 3.66% | 12 |
| after-r1 | p512-k1-file | -1.01% | -1.05% | 0.68% | 0.99% | 1 |
| after-r1 | p512-k1-owned | 0.32% | 0.47% | -0.61% | 0.11% | 3 |
| after-r1 | p512-k32-file | -91.86% | -92.59% | 3.48% | -0.23% | 20 |
| after-r1 | p512-k32-owned | -91.96% | -92.55% | -12.77% | 0.75% | 15 |
| after-r1 | p512-k8-file | -73.20% | -74.35% | -41.32% | 3.18% | 18 |
| after-r1 | p512-k8-owned | -72.58% | -74.35% | -1.85% | 3.41% | 13 |
| after-r2 | p128-k1-file | 0.01% | -1.66% | 3.17% | 0.31% | 6 |
| after-r2 | p128-k1-owned | 5.29% | 3.32% | 20.36% | -0.33% | 22 |
| after-r2 | p128-k32-file | -89.96% | -92.10% | -4.11% | 0.96% | 14 |
| after-r2 | p128-k32-owned | -90.40% | -92.07% | -8.77% | 0.00% | 18 |
| after-r2 | p128-k8-file | -69.45% | -73.72% | -37.81% | -0.42% | 19 |
| after-r2 | p128-k8-owned | -68.55% | -73.80% | -1.20% | 2.70% | 15 |
| after-r2 | p512-k1-file | -0.44% | -0.39% | -0.26% | -0.12% | 2 |
| after-r2 | p512-k1-owned | -0.64% | -0.73% | 0.92% | -0.17% | 5 |
| after-r2 | p512-k32-file | -91.75% | -92.51% | -5.76% | -0.97% | 14 |
| after-r2 | p512-k32-owned | -91.82% | -92.41% | -3.17% | -0.27% | 15 |
| after-r2 | p512-k8-file | -73.06% | -74.28% | -40.22% | 3.82% | 22 |
| after-r2 | p512-k8-owned | -72.40% | -74.13% | -1.93% | 2.31% | 14 |

## API-choice flags

24 route records have one or more absolute threshold flags.

| Campaign | Workload | flags |
| --- | --- | --- |
| after-r1 | p128-k1-file | commit_ns.p99 -95.22% (>15%); commit_ns.mean -29.57% (>5%) |
| after-r1 | p128-k1-owned | commit_ns.p50 +6.90% (>5%); commit_ns.mean +5.49% (>5%) |
| after-r1 | p128-k32-file | elapsed_ns.p50 -89.78% (>5%); elapsed_ns.p95 -89.64% (>10%); elapsed_ns.p99 -89.88% (>15%); elapsed_ns.mean -89.86% (>5%); open_ns.p99 -26.42% (>15%); edit_ns.p50 -92.07% (>5%); edit_ns.p95 -91.90% (>10%); edit_ns.p99 -92.08% (>15%); edit_ns.mean -92.08% (>5%); commit_ns.p50 +8.85% (>5%); commit_ns.mean +8.39% (>5%); publish_ns.p50 +53.38% (>5%); publish_ns.p95 +50.26% (>10%); publish_ns.p99 +22.14% (>15%); publish_ns.mean +46.11% (>5%); drop_ns.p99 +309.93% (>15%) |
| after-r1 | p128-k32-owned | elapsed_ns.p50 -90.38% (>5%); elapsed_ns.p95 -90.08% (>10%); elapsed_ns.p99 -89.98% (>15%); elapsed_ns.mean -90.35% (>5%); open_ns.p50 -5.22% (>5%); open_ns.p95 +28.87% (>10%); open_ns.mean +12.23% (>5%); edit_ns.p50 -92.00% (>5%); edit_ns.p95 -91.93% (>10%); edit_ns.p99 -91.92% (>15%); edit_ns.mean -92.02% (>5%); commit_ns.p99 -38.89% (>15%); drop_ns.p50 -14.18% (>5%); drop_ns.mean -13.82% (>5%) |
| after-r1 | p128-k8-file | elapsed_ns.p50 -68.52% (>5%); elapsed_ns.p95 -66.57% (>10%); elapsed_ns.p99 -68.96% (>15%); elapsed_ns.mean -68.40% (>5%); open_ns.p99 -16.24% (>15%); edit_ns.p50 -73.91% (>5%); edit_ns.p95 -73.26% (>10%); edit_ns.p99 -74.86% (>15%); edit_ns.mean -73.81% (>5%); commit_ns.p99 -93.11% (>15%); commit_ns.mean -19.90% (>5%); publish_ns.p95 +21.47% (>10%); drop_ns.p99 -29.17% (>15%) |
| after-r1 | p128-k8-owned | elapsed_ns.p50 -68.61% (>5%); elapsed_ns.p95 -67.98% (>10%); elapsed_ns.p99 -68.60% (>15%); elapsed_ns.mean -68.76% (>5%); open_ns.p99 -36.13% (>15%); edit_ns.p50 -73.84% (>5%); edit_ns.p95 -73.39% (>10%); edit_ns.p99 -73.41% (>15%); edit_ns.mean -73.74% (>5%); drop_ns.p50 -7.87% (>5%); drop_ns.p99 -94.19% (>15%); drop_ns.mean -27.78% (>5%) |
| after-r1 | p512-k1-file | publish_ns.p95 +12.12% (>10%) |
| after-r1 | p512-k1-owned | open_ns.p99 -20.42% (>15%); open_ns.mean -5.35% (>5%); commit_ns.p99 +17.02% (>15%) |
| after-r1 | p512-k32-file | elapsed_ns.p50 -91.86% (>5%); elapsed_ns.p95 -91.76% (>10%); elapsed_ns.p99 -91.76% (>15%); elapsed_ns.mean -91.90% (>5%); open_ns.p95 +21.47% (>10%); open_ns.p99 +23.07% (>15%); edit_ns.p50 -92.59% (>5%); edit_ns.p95 -92.46% (>10%); edit_ns.p99 -92.47% (>15%); edit_ns.mean -92.57% (>5%); commit_ns.p50 -12.41% (>5%); commit_ns.p95 -21.16% (>10%); commit_ns.p99 -31.44% (>15%); commit_ns.mean -13.49% (>5%); publish_ns.p95 -17.90% (>10%); publish_ns.p99 -24.73% (>15%); publish_ns.mean -11.50% (>5%); drop_ns.p50 -8.74% (>5%); drop_ns.p95 -11.36% (>10%); drop_ns.p99 +190.57% (>15%) |
| after-r1 | p512-k32-owned | elapsed_ns.p50 -91.96% (>5%); elapsed_ns.p95 -91.86% (>10%); elapsed_ns.p99 -91.91% (>15%); elapsed_ns.mean -91.96% (>5%); edit_ns.p50 -92.55% (>5%); edit_ns.p95 -92.48% (>10%); edit_ns.p99 -92.53% (>15%); edit_ns.mean -92.55% (>5%); commit_ns.p50 -8.82% (>5%); commit_ns.mean -8.36% (>5%); publish_ns.p50 -12.77% (>5%); publish_ns.p95 -14.90% (>10%); publish_ns.mean -11.72% (>5%); drop_ns.p50 -9.29% (>5%); drop_ns.mean -8.70% (>5%) |
| after-r1 | p512-k8-file | elapsed_ns.p50 -73.20% (>5%); elapsed_ns.p95 -72.97% (>10%); elapsed_ns.p99 -73.56% (>15%); elapsed_ns.mean -73.18% (>5%); open_ns.p50 -49.75% (>5%); open_ns.p95 -49.87% (>10%); open_ns.p99 -39.60% (>15%); open_ns.mean -49.39% (>5%); edit_ns.p50 -74.35% (>5%); edit_ns.p95 -74.19% (>10%); edit_ns.p99 -74.90% (>15%); edit_ns.mean -74.36% (>5%); publish_ns.p50 -41.32% (>5%); publish_ns.p95 -40.58% (>10%); publish_ns.p99 -30.18% (>15%); publish_ns.mean -40.36% (>5%); drop_ns.p50 -7.58% (>5%); drop_ns.mean -8.67% (>5%) |
| after-r1 | p512-k8-owned | elapsed_ns.p50 -72.58% (>5%); elapsed_ns.p95 -72.42% (>10%); elapsed_ns.p99 -73.38% (>15%); elapsed_ns.mean -72.58% (>5%); open_ns.mean -8.56% (>5%); edit_ns.p50 -74.35% (>5%); edit_ns.p95 -74.24% (>10%); edit_ns.p99 -75.11% (>15%); edit_ns.mean -74.36% (>5%); commit_ns.p50 -8.51% (>5%); commit_ns.mean -5.15% (>5%); drop_ns.p50 -7.58% (>5%); drop_ns.mean -6.87% (>5%) |
| after-r2 | p128-k1-file | elapsed_ns.p99 +79.78% (>15%); open_ns.p99 +22.56% (>15%); edit_ns.p99 +98.56% (>15%); commit_ns.p50 -6.06% (>5%); commit_ns.mean -6.29% (>5%); drop_ns.p99 +21.84% (>15%) |
| after-r2 | p128-k1-owned | elapsed_ns.p50 +5.29% (>5%); elapsed_ns.p95 +26.93% (>10%); elapsed_ns.p99 +946.83% (>15%); elapsed_ns.mean +26.70% (>5%); open_ns.p95 +74.46% (>10%); open_ns.p99 +111.75% (>15%); open_ns.mean +10.56% (>5%); edit_ns.p95 +15.76% (>10%); edit_ns.p99 +1068.78% (>15%); edit_ns.mean +26.42% (>5%); commit_ns.p50 +20.00% (>5%); commit_ns.p95 +418.75% (>10%); commit_ns.p99 +555.88% (>15%); commit_ns.mean +169.99% (>5%); publish_ns.p50 +20.36% (>5%); publish_ns.p95 +86.83% (>10%); publish_ns.p99 +183.71% (>15%); publish_ns.mean +34.44% (>5%); drop_ns.p50 +23.38% (>5%); drop_ns.p95 +84.34% (>10%); drop_ns.p99 +75.76% (>15%); drop_ns.mean +30.80% (>5%) |
| after-r2 | p128-k32-file | elapsed_ns.p50 -89.96% (>5%); elapsed_ns.p95 -89.80% (>10%); elapsed_ns.p99 -90.01% (>15%); elapsed_ns.mean -90.10% (>5%); open_ns.p99 +23.72% (>15%); edit_ns.p50 -92.10% (>5%); edit_ns.p95 -91.94% (>10%); edit_ns.p99 -92.09% (>15%); edit_ns.mean -92.10% (>5%); commit_ns.p99 -78.92% (>15%); commit_ns.mean -7.41% (>5%); publish_ns.mean -6.11% (>5%); drop_ns.p99 -49.57% (>15%); drop_ns.mean -6.54% (>5%) |
| after-r2 | p128-k32-owned | elapsed_ns.p50 -90.40% (>5%); elapsed_ns.p95 -90.15% (>10%); elapsed_ns.p99 -90.17% (>15%); elapsed_ns.mean -90.40% (>5%); open_ns.p50 -6.45% (>5%); open_ns.mean +10.59% (>5%); edit_ns.p50 -92.07% (>5%); edit_ns.p95 -91.91% (>10%); edit_ns.p99 -91.89% (>15%); edit_ns.mean -92.04% (>5%); commit_ns.p95 -17.32% (>10%); commit_ns.p99 -23.41% (>15%); publish_ns.p50 -8.77% (>5%); publish_ns.p95 -18.43% (>10%); publish_ns.p99 -25.56% (>15%); publish_ns.mean -9.92% (>5%); drop_ns.p50 -13.48% (>5%); drop_ns.mean -11.28% (>5%) |
| after-r2 | p128-k8-file | elapsed_ns.p50 -69.45% (>5%); elapsed_ns.p95 -69.05% (>10%); elapsed_ns.p99 -67.88% (>15%); elapsed_ns.mean -69.26% (>5%); open_ns.p50 -5.75% (>5%); open_ns.mean -6.22% (>5%); edit_ns.p50 -73.72% (>5%); edit_ns.p95 -73.29% (>10%); edit_ns.p99 -73.16% (>15%); edit_ns.mean -73.65% (>5%); commit_ns.p50 -9.30% (>5%); commit_ns.p95 -12.50% (>10%); commit_ns.mean -9.13% (>5%); publish_ns.p50 -37.81% (>5%); publish_ns.p95 -33.42% (>10%); publish_ns.p99 -15.83% (>15%); publish_ns.mean -31.30% (>5%); drop_ns.p50 -8.46% (>5%); drop_ns.mean -8.07% (>5%) |
| after-r2 | p128-k8-owned | elapsed_ns.p50 -68.55% (>5%); elapsed_ns.p95 -68.17% (>10%); elapsed_ns.p99 -68.23% (>15%); elapsed_ns.mean -68.84% (>5%); open_ns.p95 -48.65% (>10%); open_ns.p99 -19.10% (>15%); open_ns.mean -9.45% (>5%); edit_ns.p50 -73.80% (>5%); edit_ns.p95 -73.21% (>10%); edit_ns.p99 -73.20% (>15%); edit_ns.mean -73.65% (>5%); publish_ns.mean -7.98% (>5%); drop_ns.p50 -10.77% (>5%); drop_ns.p95 -12.41% (>10%); drop_ns.mean -10.43% (>5%) |
| after-r2 | p512-k1-file | commit_ns.p50 -13.16% (>5%); commit_ns.mean -9.84% (>5%) |
| after-r2 | p512-k1-owned | open_ns.p99 +30.99% (>15%); commit_ns.p50 +6.06% (>5%); commit_ns.p95 +21.62% (>10%); commit_ns.p99 +45.00% (>15%); commit_ns.mean +9.63% (>5%) |
| after-r2 | p512-k32-file | elapsed_ns.p50 -91.75% (>5%); elapsed_ns.p95 -91.66% (>10%); elapsed_ns.p99 -91.61% (>15%); elapsed_ns.mean -91.80% (>5%); open_ns.p50 -5.57% (>5%); open_ns.p95 +11.14% (>10%); edit_ns.p50 -92.51% (>5%); edit_ns.p95 -92.46% (>10%); edit_ns.p99 -92.48% (>15%); edit_ns.mean -92.51% (>5%); publish_ns.p50 -5.76% (>5%); drop_ns.p50 -8.48% (>5%); drop_ns.p99 +124.21% (>15%); drop_ns.mean -5.34% (>5%) |
| after-r2 | p512-k32-owned | elapsed_ns.p50 -91.82% (>5%); elapsed_ns.p95 -91.72% (>10%); elapsed_ns.p99 -91.63% (>15%); elapsed_ns.mean -91.81% (>5%); open_ns.p50 -6.23% (>5%); open_ns.p99 -20.79% (>15%); edit_ns.p50 -92.41% (>5%); edit_ns.p95 -92.35% (>10%); edit_ns.p99 -92.27% (>15%); edit_ns.mean -92.41% (>5%); commit_ns.p99 -63.59% (>15%); publish_ns.p99 -18.73% (>15%); drop_ns.p50 -12.14% (>5%); drop_ns.p99 -74.28% (>15%); drop_ns.mean -15.86% (>5%) |
| after-r2 | p512-k8-file | elapsed_ns.p50 -73.06% (>5%); elapsed_ns.p95 -72.86% (>10%); elapsed_ns.p99 -72.65% (>15%); elapsed_ns.mean -73.04% (>5%); open_ns.p50 -49.31% (>5%); open_ns.p95 -49.68% (>10%); open_ns.p99 -44.44% (>15%); open_ns.mean -49.24% (>5%); edit_ns.p50 -74.28% (>5%); edit_ns.p95 -74.08% (>10%); edit_ns.p99 -73.82% (>15%); edit_ns.mean -74.22% (>5%); commit_ns.p50 -8.89% (>5%); commit_ns.p95 -22.41% (>10%); commit_ns.p99 -23.81% (>15%); commit_ns.mean -14.01% (>5%); publish_ns.p50 -40.22% (>5%); publish_ns.p95 -40.16% (>10%); publish_ns.p99 -35.70% (>15%); publish_ns.mean -40.03% (>5%); drop_ns.p50 -9.85% (>5%); drop_ns.mean -9.87% (>5%) |
| after-r2 | p512-k8-owned | elapsed_ns.p50 -72.40% (>5%); elapsed_ns.p95 -72.11% (>10%); elapsed_ns.p99 -71.29% (>15%); elapsed_ns.mean -72.32% (>5%); open_ns.p50 -20.29% (>5%); open_ns.p95 -22.75% (>10%); open_ns.p99 -25.32% (>15%); open_ns.mean -16.96% (>5%); edit_ns.p50 -74.13% (>5%); edit_ns.p95 -73.84% (>10%); edit_ns.p99 -72.97% (>15%); edit_ns.mean -74.08% (>5%); drop_ns.p50 -6.82% (>5%); drop_ns.mean -6.77% (>5%) |

## Per-case gauges and guards

The table exposes the Work charge, retained live reservations, release/output/input counters, and source reads for every case. Every listed guard is true; Work is expected unchanged and retained live-gauge deltas remain explicit comparison data.

| Campaign | Workload | API | after Work | live Work | live memory | live objects | after input | after output | source reads | guards |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| after-r1 | p128-k1-file-batch | batch | 6714084 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| after-r1 | p128-k1-file-repeated | repeated | 6714084 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| after-r1 | p128-k1-owned-batch | batch | 6714084 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| after-r1 | p128-k1-owned-repeated | repeated | 6714084 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| after-r1 | p128-k32-file-batch | batch | 7168451 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| after-r1 | p128-k32-file-repeated | repeated | 13132231 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| after-r1 | p128-k32-owned-batch | batch | 7168451 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| after-r1 | p128-k32-owned-repeated | repeated | 13132231 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| after-r1 | p128-k8-file-batch | batch | 6816683 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| after-r1 | p128-k8-file-repeated | repeated | 7497391 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| after-r1 | p128-k8-owned-batch | batch | 6816683 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| after-r1 | p128-k8-owned-repeated | repeated | 7497391 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| after-r1 | p512-k1-file-batch | batch | 103874532 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| after-r1 | p512-k1-file-repeated | repeated | 103874532 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| after-r1 | p512-k1-owned-batch | batch | 103874532 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| after-r1 | p512-k1-owned-repeated | repeated | 103874532 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| after-r1 | p512-k32-file-batch | batch | 104328899 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| after-r1 | p512-k32-file-repeated | repeated | 116292295 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| after-r1 | p512-k32-owned-batch | batch | 104328899 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| after-r1 | p512-k32-owned-repeated | repeated | 116292295 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| after-r1 | p512-k8-file-batch | batch | 103977131 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| after-r1 | p512-k8-file-repeated | repeated | 106012591 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |
| after-r1 | p512-k8-owned-batch | batch | 103977131 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| after-r1 | p512-k8-owned-repeated | repeated | 106012591 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |
| after-r2 | p128-k1-file-batch | batch | 6714084 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| after-r2 | p128-k1-file-repeated | repeated | 6714084 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| after-r2 | p128-k1-owned-batch | batch | 6714084 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| after-r2 | p128-k1-owned-repeated | repeated | 6714084 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| after-r2 | p128-k32-file-batch | batch | 7168451 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| after-r2 | p128-k32-file-repeated | repeated | 13132231 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| after-r2 | p128-k32-owned-batch | batch | 7168451 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| after-r2 | p128-k32-owned-repeated | repeated | 13132231 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| after-r2 | p128-k8-file-batch | batch | 6816683 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| after-r2 | p128-k8-file-repeated | repeated | 7497391 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| after-r2 | p128-k8-owned-batch | batch | 6816683 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| after-r2 | p128-k8-owned-repeated | repeated | 7497391 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| after-r2 | p512-k1-file-batch | batch | 103874532 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| after-r2 | p512-k1-file-repeated | repeated | 103874532 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| after-r2 | p512-k1-owned-batch | batch | 103874532 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| after-r2 | p512-k1-owned-repeated | repeated | 103874532 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| after-r2 | p512-k32-file-batch | batch | 104328899 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| after-r2 | p512-k32-file-repeated | repeated | 116292295 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| after-r2 | p512-k32-owned-batch | batch | 104328899 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| after-r2 | p512-k32-owned-repeated | repeated | 116292295 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| after-r2 | p512-k8-file-batch | batch | 103977131 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| after-r2 | p512-k8-file-repeated | repeated | 106012591 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |
| after-r2 | p512-k8-owned-batch | batch | 103977131 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| after-r2 | p512-k8-owned-repeated | repeated | 106012591 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |

## Cross-campaign flags

### Matched same API

22 of 24 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-file-batch | elapsed_ns.p99 +88.12% (>15%); edit_ns.p99 +122.32% (>15%); commit_ns.p95 +23.53% (>10%); commit_ns.p99 +81.58% (>15%); commit_ns.mean +5.46% (>5%); publish_ns.p50 +5.07% (>5%); publish_ns.p95 +16.03% (>10%); publish_ns.mean +9.21% (>5%); drop_ns.p99 +23.26% (>15%) |
| p128-k1-file-repeated | open_ns.p99 -30.06% (>15%); commit_ns.p50 +6.45% (>5%); commit_ns.p95 +16.67% (>10%); commit_ns.p99 -90.44% (>15%); commit_ns.mean -20.74% (>5%) |
| p128-k1-owned-batch | elapsed_ns.p50 +5.29% (>5%); elapsed_ns.p95 +25.14% (>10%); elapsed_ns.p99 +920.33% (>15%); elapsed_ns.mean +26.62% (>5%); open_ns.p95 +75.41% (>10%); open_ns.p99 +64.82% (>15%); open_ns.mean +9.73% (>5%); edit_ns.p95 +14.05% (>10%); edit_ns.p99 +1035.25% (>15%); edit_ns.mean +26.24% (>5%); commit_ns.p50 +16.13% (>5%); commit_ns.p95 +374.29% (>10%); commit_ns.p99 +471.79% (>15%); commit_ns.mean +152.14% (>5%); publish_ns.p50 +22.71% (>5%); publish_ns.p95 +86.22% (>10%); publish_ns.p99 +175.04% (>15%); publish_ns.mean +35.73% (>5%); drop_ns.p50 +23.38% (>5%); drop_ns.p95 +82.14% (>10%); drop_ns.p99 +169.77% (>15%); drop_ns.mean +32.33% (>5%) |
| p128-k1-owned-repeated | open_ns.p99 -22.09% (>15%); drop_ns.p99 +55.29% (>15%) |
| p128-k32-file-batch | open_ns.p99 +38.38% (>15%); publish_ns.mean -7.59% (>5%); drop_ns.p99 -76.25% (>15%) |
| p128-k32-file-repeated | open_ns.p99 -17.70% (>15%); commit_ns.p50 +6.19% (>5%); commit_ns.p99 +333.99% (>15%); commit_ns.mean +13.32% (>5%); publish_ns.p50 +57.17% (>5%); publish_ns.p95 +52.62% (>10%); publish_ns.p99 +25.40% (>15%); publish_ns.mean +43.82% (>5%); drop_ns.p99 +93.05% (>15%) |
| p128-k32-owned-batch | commit_ns.p95 +12.98% (>10%); commit_ns.mean +5.40% (>5%) |
| p128-k32-owned-repeated | open_ns.p95 +30.10% (>10%); commit_ns.p95 +24.31% (>10%); publish_ns.p95 +15.78% (>10%); publish_ns.p99 +27.72% (>15%); publish_ns.mean +9.54% (>5%) |
| p128-k8-file-batch | open_ns.p99 +44.73% (>15%); commit_ns.p95 -10.64% (>10%); publish_ns.p95 -16.93% (>10%) |
| p128-k8-file-repeated | open_ns.p99 +28.52% (>15%); open_ns.mean +5.30% (>5%); commit_ns.p99 -92.52% (>15%); commit_ns.mean -16.03% (>5%); publish_ns.p50 +58.18% (>5%); publish_ns.p95 +51.55% (>10%); publish_ns.p99 +19.31% (>15%); publish_ns.mean +44.42% (>5%); drop_ns.p99 -26.56% (>15%) |
| p128-k8-owned-batch | open_ns.p99 +28.76% (>15%); commit_ns.p50 -8.70% (>5%); commit_ns.mean -7.78% (>5%) |
| p128-k8-owned-repeated | open_ns.p95 +94.59% (>10%); open_ns.mean +9.39% (>5%); commit_ns.mean -5.65% (>5%); publish_ns.mean +6.54% (>5%); drop_ns.p99 -93.38% (>15%); drop_ns.mean -18.77% (>5%) |
| p512-k1-file-repeated | commit_ns.p50 +11.76% (>5%); commit_ns.mean +7.36% (>5%) |
| p512-k1-owned-batch | open_ns.p99 +40.59% (>15%) |
| p512-k1-owned-repeated | commit_ns.p50 -5.71% (>5%); commit_ns.p95 -19.57% (>10%); commit_ns.mean -8.07% (>5%) |
| p512-k32-file-batch | publish_ns.p50 +6.41% (>5%) |
| p512-k32-file-repeated | open_ns.p95 +16.61% (>10%); open_ns.p99 +17.70% (>15%); commit_ns.p50 -8.03% (>5%); commit_ns.p95 -22.75% (>10%); commit_ns.p99 -24.45% (>15%); commit_ns.mean -9.93% (>5%); publish_ns.p50 +16.84% (>5%); publish_ns.p95 -15.50% (>10%); publish_ns.mean -9.20% (>5%); drop_ns.p99 +28.62% (>15%) |
| p512-k32-owned-batch | open_ns.p50 -18.04% (>5%); open_ns.p99 -17.59% (>15%) |
| p512-k32-owned-repeated | open_ns.p50 -10.97% (>5%); commit_ns.p50 -8.82% (>5%); commit_ns.p95 -10.69% (>10%); commit_ns.p99 +138.15% (>15%); publish_ns.p50 -9.43% (>5%); publish_ns.p95 -14.19% (>10%); publish_ns.mean -9.91% (>5%); drop_ns.p99 +236.76% (>15%) |
| p512-k8-file-repeated | open_ns.p99 +23.75% (>15%); commit_ns.p95 +18.37% (>10%); commit_ns.mean +6.47% (>5%) |
| p512-k8-owned-batch | open_ns.p50 -17.85% (>5%); open_ns.mean -7.13% (>5%) |
| p512-k8-owned-repeated | open_ns.p95 +22.25% (>10%); commit_ns.p50 -10.64% (>5%); commit_ns.mean -7.41% (>5%) |

### API-choice ratio

12 of 12 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-file | elapsed_ns.p99 +79.14% (>15%); open_ns.p99 +29.99% (>15%); edit_ns.p99 +100.60% (>15%); commit_ns.p99 +1799.41% (>15%); commit_ns.mean +33.05% (>5%); drop_ns.p99 +17.59% (>15%) |
| p128-k1-owned | elapsed_ns.p50 +5.30% (>5%); elapsed_ns.p95 +25.54% (>10%); elapsed_ns.p99 +917.88% (>15%); elapsed_ns.mean +26.60% (>5%); open_ns.p95 +74.73% (>10%); open_ns.p99 +111.55% (>15%); open_ns.mean +10.69% (>5%); edit_ns.p95 +13.47% (>10%); edit_ns.p99 +1030.11% (>15%); edit_ns.mean +25.73% (>5%); commit_ns.p50 +12.26% (>5%); commit_ns.p95 +418.75% (>10%); commit_ns.p99 +522.25% (>15%); commit_ns.mean +155.93% (>5%); publish_ns.p50 +25.11% (>5%); publish_ns.p95 +105.03% (>10%); publish_ns.p99 +181.09% (>15%); publish_ns.mean +39.85% (>5%); drop_ns.p50 +24.98% (>5%); drop_ns.p95 +82.14% (>10%); drop_ns.p99 +73.71% (>15%); drop_ns.mean +33.25% (>5%) |
| p128-k32-file | open_ns.p50 -5.32% (>5%); open_ns.p99 +68.15% (>15%); commit_ns.p50 -8.13% (>5%); commit_ns.p99 -78.06% (>15%); commit_ns.mean -14.58% (>5%); publish_ns.p50 -37.48% (>5%); publish_ns.p95 -36.14% (>10%); publish_ns.p99 -21.81% (>15%); publish_ns.mean -35.74% (>5%); drop_ns.p99 -87.70% (>15%); drop_ns.mean -7.79% (>5%) |
| p128-k32-owned | open_ns.p95 -22.97% (>10%); commit_ns.p99 +25.32% (>15%); publish_ns.p95 -16.81% (>10%); publish_ns.p99 -23.99% (>15%); publish_ns.mean -9.71% (>5%) |
| p128-k8-file | commit_ns.p50 -9.30% (>5%); commit_ns.p95 -14.36% (>10%); commit_ns.p99 +1180.35% (>15%); commit_ns.mean +13.45% (>5%); publish_ns.p50 -37.27% (>5%); publish_ns.p95 -45.19% (>10%); publish_ns.p99 -18.66% (>15%); publish_ns.mean -33.97% (>5%); drop_ns.p99 +52.19% (>15%) |
| p128-k8-owned | open_ns.p95 -49.45% (>10%); open_ns.p99 +26.66% (>15%); open_ns.mean -8.94% (>5%); publish_ns.mean -6.68% (>5%); drop_ns.p99 +1610.76% (>15%); drop_ns.mean +24.02% (>5%) |
| p512-k1-file | commit_ns.p50 -10.53% (>5%); commit_ns.mean -6.02% (>5%); drop_ns.p50 +5.13% (>5%); drop_ns.mean +5.35% (>5%) |
| p512-k1-owned | open_ns.p99 +64.59% (>15%); commit_ns.p50 +9.18% (>5%); commit_ns.p95 +24.32% (>10%); commit_ns.p99 +23.91% (>15%); commit_ns.mean +14.21% (>5%) |
| p512-k32-file | commit_ns.p50 +10.54% (>5%); commit_ns.p95 +31.19% (>10%); commit_ns.p99 +37.43% (>15%); commit_ns.mean +13.46% (>5%); publish_ns.p50 -8.93% (>5%); publish_ns.p95 +21.56% (>10%); publish_ns.p99 +16.32% (>15%); publish_ns.mean +15.49% (>5%); drop_ns.p99 -22.84% (>15%) |
| p512-k32-owned | open_ns.p50 -7.94% (>5%); open_ns.p99 -20.84% (>15%); commit_ns.p50 +9.68% (>5%); commit_ns.p99 -59.10% (>15%); publish_ns.p50 +11.00% (>5%); publish_ns.p95 +17.71% (>10%); publish_ns.mean +10.89% (>5%); drop_ns.p99 -71.44% (>15%); drop_ns.mean -7.84% (>5%) |
| p512-k8-file | commit_ns.p95 -22.41% (>10%); commit_ns.mean -10.71% (>5%) |
| p512-k8-owned | open_ns.p50 -17.93% (>5%); open_ns.p95 -19.11% (>10%); open_ns.mean -9.19% (>5%); commit_ns.p50 +9.30% (>5%) |


## Counter drift

Same-API output identities match for 24/24 cases across campaigns.
Release/output/read guard counter changes: 0 cases.
Allowed retained live-gauge changes: 0 cases; Work changes would be guard findings. Any values appear in JSON.

All case/campaign phase statistics, repeat groups, route ratios, counter drift, identities, guards, and bootstrap intervals are in `candidate-native-analysis.json`.
