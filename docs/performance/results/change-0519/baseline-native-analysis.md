# 0519 native DOCX analysis

Native campaigns: r1, r2. Rows are source-bound through each receipt; profile and hardware lanes are excluded.
Measured rows are the 30 `warmup=false` samples in each of two internal repeats. Every case reports nearest-rank p50/p95/p99 and arithmetic mean for elapsed, open, edit, commit, publish, and drop clocks; RSS is the one whole-child maximum from GNU time.
The unpaired route bootstrap for the p50 ratio uses seed `5347857` and 4000 iterations, resampling each route within each internal repeat and taking batch median / repeated median. With only two internal repeats, its interval is descriptive for these row distributions and does not estimate independent-process or host variation.
The historical 0500 validator is imported by `capture.py` and remains the row-level source for release, output, semantic, untouched-member, source-version, and readback guards. Work is expected to remain unchanged in this batch; retained live gauges are reported as descriptive counters and do not relax those guards.

## API-choice ratios

Ratios are batch / repeated; positive percentages mean batch took longer or used more RSS. The same absolute thresholds are used for route flags and cross-campaign flags: 5% for p50/mean/RSS, 10% for p95, and 15% for p99.

| Campaign | Workload | elapsed p50 | edit p50 | publish p50 | RSS | route flags |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| r1 | p128-k1-file | -0.66% | -0.92% | -0.03% | 0.67% | 1 |
| r1 | p128-k1-owned | -1.52% | -2.33% | -1.22% | 0.88% | 2 |
| r1 | p128-k32-file | -88.88% | -92.07% | -2.62% | -1.26% | 15 |
| r1 | p128-k32-owned | -89.22% | -92.01% | -3.12% | 0.73% | 12 |
| r1 | p128-k8-file | -66.09% | -73.71% | -23.61% | 5.28% | 19 |
| r1 | p128-k8-owned | -65.45% | -74.02% | -0.65% | 3.15% | 10 |
| r1 | p512-k1-file | -0.10% | 0.07% | 0.36% | 1.92% | 0 |
| r1 | p512-k1-owned | -1.92% | -2.64% | -0.74% | -2.79% | 2 |
| r1 | p512-k32-file | -90.31% | -92.52% | -0.11% | -0.91% | 11 |
| r1 | p512-k32-owned | -90.40% | -92.51% | -1.11% | 3.16% | 13 |
| r1 | p512-k8-file | -67.83% | -74.06% | -2.11% | 4.42% | 14 |
| r1 | p512-k8-owned | -68.40% | -74.41% | -0.62% | 0.44% | 12 |
| r2 | p128-k1-file | 1.25% | 1.11% | 1.17% | -0.24% | 1 |
| r2 | p128-k1-owned | -9.30% | 0.47% | -30.31% | 2.99% | 6 |
| r2 | p128-k32-file | -88.78% | -92.01% | -1.24% | -3.90% | 10 |
| r2 | p128-k32-owned | -89.18% | -91.94% | -1.75% | 0.45% | 10 |
| r2 | p128-k8-file | -66.40% | -74.06% | -24.05% | -2.22% | 20 |
| r2 | p128-k8-owned | -65.28% | -73.56% | -5.51% | 2.76% | 19 |
| r2 | p512-k1-file | 0.42% | 0.65% | 0.04% | -0.43% | 5 |
| r2 | p512-k1-owned | 0.55% | 0.54% | 0.88% | 0.58% | 4 |
| r2 | p512-k32-file | -90.29% | -92.52% | -1.82% | 0.92% | 13 |
| r2 | p512-k32-owned | -90.49% | -92.58% | -1.30% | 0.16% | 12 |
| r2 | p512-k8-file | -68.80% | -74.11% | -12.83% | 7.45% | 21 |
| r2 | p512-k8-owned | -67.99% | -74.20% | 1.47% | -0.33% | 14 |

## API-choice flags

23 route records have one or more absolute threshold flags.

| Campaign | Workload | flags |
| --- | --- | --- |
| r1 | p128-k1-file | drop_ns.p99 -21.82% (>15%) |
| r1 | p128-k1-owned | open_ns.p99 -38.42% (>15%); drop_ns.p50 +5.26% (>5%) |
| r1 | p128-k32-file | elapsed_ns.p50 -88.88% (>5%); elapsed_ns.p95 -88.79% (>10%); elapsed_ns.p99 -88.88% (>15%); elapsed_ns.mean -89.04% (>5%); open_ns.p99 +15.12% (>15%); edit_ns.p50 -92.07% (>5%); edit_ns.p95 -92.01% (>10%); edit_ns.p99 -92.08% (>15%); edit_ns.mean -92.10% (>5%); commit_ns.p99 -30.96% (>15%); publish_ns.p95 -10.91% (>10%); publish_ns.mean -5.04% (>5%); drop_ns.p50 -8.16% (>5%); drop_ns.p99 +28.76% (>15%); drop_ns.mean -7.94% (>5%) |
| r1 | p128-k32-owned | elapsed_ns.p50 -89.22% (>5%); elapsed_ns.p95 -88.96% (>10%); elapsed_ns.p99 -89.21% (>15%); elapsed_ns.mean -89.19% (>5%); open_ns.p99 +20.34% (>15%); open_ns.mean +9.79% (>5%); edit_ns.p50 -92.01% (>5%); edit_ns.p95 -91.88% (>10%); edit_ns.p99 -92.08% (>15%); edit_ns.mean -91.98% (>5%); drop_ns.p50 -11.79% (>5%); drop_ns.mean -10.25% (>5%) |
| r1 | p128-k8-file | elapsed_ns.p50 -66.09% (>5%); elapsed_ns.p95 -65.88% (>10%); elapsed_ns.p99 -64.22% (>15%); elapsed_ns.mean -65.95% (>5%); open_ns.p50 -7.12% (>5%); open_ns.mean -5.75% (>5%); edit_ns.p50 -73.71% (>5%); edit_ns.p95 -73.33% (>10%); edit_ns.p99 -73.77% (>15%); edit_ns.mean -73.62% (>5%); commit_ns.p50 -7.14% (>5%); commit_ns.mean -6.10% (>5%); publish_ns.p50 -23.61% (>5%); publish_ns.p95 -21.70% (>10%); publish_ns.mean -19.78% (>5%); drop_ns.p50 -9.77% (>5%); drop_ns.p99 -83.59% (>15%); drop_ns.mean -15.86% (>5%); rss_kib.rss +5.28% (>5%) |
| r1 | p128-k8-owned | elapsed_ns.p50 -65.45% (>5%); elapsed_ns.p95 -65.32% (>10%); elapsed_ns.p99 -65.13% (>15%); elapsed_ns.mean -65.46% (>5%); edit_ns.p50 -74.02% (>5%); edit_ns.p95 -73.76% (>10%); edit_ns.p99 -72.77% (>15%); edit_ns.mean -73.96% (>5%); drop_ns.p50 -6.98% (>5%); drop_ns.mean -7.29% (>5%) |
| r1 | p512-k1-owned | commit_ns.p50 -5.88% (>5%); commit_ns.mean -5.53% (>5%) |
| r1 | p512-k32-file | elapsed_ns.p50 -90.31% (>5%); elapsed_ns.p95 -90.24% (>10%); elapsed_ns.p99 -90.24% (>15%); elapsed_ns.mean -90.37% (>5%); edit_ns.p50 -92.52% (>5%); edit_ns.p95 -92.48% (>10%); edit_ns.p99 -92.49% (>15%); edit_ns.mean -92.52% (>5%); drop_ns.p50 -6.36% (>5%); drop_ns.p99 -43.19% (>15%); drop_ns.mean -7.51% (>5%) |
| r1 | p512-k32-owned | elapsed_ns.p50 -90.40% (>5%); elapsed_ns.p95 -90.36% (>10%); elapsed_ns.p99 -90.32% (>15%); elapsed_ns.mean -90.41% (>5%); open_ns.p50 -5.27% (>5%); edit_ns.p50 -92.51% (>5%); edit_ns.p95 -92.42% (>10%); edit_ns.p99 -92.34% (>15%); edit_ns.mean -92.50% (>5%); commit_ns.p99 +477.03% (>15%); commit_ns.mean +11.94% (>5%); drop_ns.p50 -11.54% (>5%); drop_ns.mean -10.79% (>5%) |
| r1 | p512-k8-file | elapsed_ns.p50 -67.83% (>5%); elapsed_ns.p95 -67.59% (>10%); elapsed_ns.p99 -67.38% (>15%); elapsed_ns.mean -67.96% (>5%); open_ns.p50 -46.89% (>5%); open_ns.p95 -46.74% (>10%); open_ns.p99 -43.25% (>15%); open_ns.mean -45.70% (>5%); edit_ns.p50 -74.06% (>5%); edit_ns.p95 -73.71% (>10%); edit_ns.p99 -73.54% (>15%); edit_ns.mean -74.02% (>5%); drop_ns.p50 -9.63% (>5%); drop_ns.mean -8.44% (>5%) |
| r1 | p512-k8-owned | elapsed_ns.p50 -68.40% (>5%); elapsed_ns.p95 -68.33% (>10%); elapsed_ns.p99 -68.31% (>15%); elapsed_ns.mean -68.38% (>5%); edit_ns.p50 -74.41% (>5%); edit_ns.p95 -74.28% (>10%); edit_ns.p99 -74.22% (>15%); edit_ns.mean -74.38% (>5%); commit_ns.p99 -26.76% (>15%); commit_ns.mean -5.14% (>5%); drop_ns.p50 -9.77% (>5%); drop_ns.mean -8.67% (>5%) |
| r2 | p128-k1-file | commit_ns.p99 -35.85% (>15%) |
| r2 | p128-k1-owned | elapsed_ns.p50 -9.30% (>5%); elapsed_ns.mean -8.65% (>5%); publish_ns.p50 -30.31% (>5%); publish_ns.p95 -29.11% (>10%); publish_ns.p99 -29.47% (>15%); publish_ns.mean -27.85% (>5%) |
| r2 | p128-k32-file | elapsed_ns.p50 -88.78% (>5%); elapsed_ns.p95 -88.68% (>10%); elapsed_ns.p99 -88.80% (>15%); elapsed_ns.mean -88.93% (>5%); edit_ns.p50 -92.01% (>5%); edit_ns.p95 -91.96% (>10%); edit_ns.p99 -91.97% (>15%); edit_ns.mean -92.03% (>5%); drop_ns.p50 -6.93% (>5%); drop_ns.p99 +165.81% (>15%) |
| r2 | p128-k32-owned | elapsed_ns.p50 -89.18% (>5%); elapsed_ns.p95 -88.82% (>10%); elapsed_ns.p99 -88.81% (>15%); elapsed_ns.mean -89.15% (>5%); edit_ns.p50 -91.94% (>5%); edit_ns.p95 -91.84% (>10%); edit_ns.p99 -91.86% (>15%); edit_ns.mean -91.94% (>5%); drop_ns.p50 -8.49% (>5%); drop_ns.p99 +228.81% (>15%) |
| r2 | p128-k8-file | elapsed_ns.p50 -66.40% (>5%); elapsed_ns.p95 -66.30% (>10%); elapsed_ns.p99 -64.90% (>15%); elapsed_ns.mean -66.40% (>5%); open_ns.p50 -5.71% (>5%); open_ns.mean -6.57% (>5%); edit_ns.p50 -74.06% (>5%); edit_ns.p95 -73.61% (>10%); edit_ns.p99 -73.57% (>15%); edit_ns.mean -73.92% (>5%); commit_ns.p50 -17.02% (>5%); commit_ns.p95 -37.68% (>10%); commit_ns.p99 -24.18% (>15%); commit_ns.mean -19.07% (>5%); publish_ns.p50 -24.05% (>5%); publish_ns.p95 -21.68% (>10%); publish_ns.p99 -17.54% (>15%); publish_ns.mean -22.75% (>5%); drop_ns.p50 -9.92% (>5%); drop_ns.mean -9.22% (>5%) |
| r2 | p128-k8-owned | elapsed_ns.p50 -65.28% (>5%); elapsed_ns.p95 -65.22% (>10%); elapsed_ns.p99 -65.18% (>15%); elapsed_ns.mean -65.65% (>5%); open_ns.p95 -51.39% (>10%); open_ns.p99 -52.87% (>15%); open_ns.mean -14.51% (>5%); edit_ns.p50 -73.56% (>5%); edit_ns.p95 -73.54% (>10%); edit_ns.p99 -72.89% (>15%); edit_ns.mean -73.56% (>5%); commit_ns.p95 -34.72% (>10%); commit_ns.p99 -39.36% (>15%); commit_ns.mean -10.05% (>5%); publish_ns.p50 -5.51% (>5%); publish_ns.p95 -10.63% (>10%); publish_ns.mean -8.74% (>5%); drop_ns.p50 -9.09% (>5%); drop_ns.mean -8.91% (>5%) |
| r2 | p512-k1-file | open_ns.p99 +26.72% (>15%); commit_ns.p50 +9.38% (>5%); commit_ns.p95 +14.29% (>10%); commit_ns.mean +10.11% (>5%); publish_ns.p99 -22.71% (>15%) |
| r2 | p512-k1-owned | commit_ns.p50 -6.06% (>5%); commit_ns.p95 -10.53% (>10%); commit_ns.mean -7.22% (>5%); drop_ns.p99 +46.15% (>15%) |
| r2 | p512-k32-file | elapsed_ns.p50 -90.29% (>5%); elapsed_ns.p95 -90.24% (>10%); elapsed_ns.p99 -90.26% (>15%); elapsed_ns.mean -90.38% (>5%); open_ns.p99 -32.39% (>15%); edit_ns.p50 -92.52% (>5%); edit_ns.p95 -92.37% (>10%); edit_ns.p99 -92.33% (>15%); edit_ns.mean -92.51% (>5%); commit_ns.p99 +26.83% (>15%); drop_ns.p50 -8.04% (>5%); drop_ns.p95 -10.61% (>10%); drop_ns.p99 +171.19% (>15%) |
| r2 | p512-k32-owned | elapsed_ns.p50 -90.49% (>5%); elapsed_ns.p95 -90.46% (>10%); elapsed_ns.p99 -90.44% (>15%); elapsed_ns.mean -90.50% (>5%); open_ns.p50 -41.16% (>5%); open_ns.mean -5.52% (>5%); edit_ns.p50 -92.58% (>5%); edit_ns.p95 -92.55% (>10%); edit_ns.p99 -92.47% (>15%); edit_ns.mean -92.58% (>5%); drop_ns.p50 -7.45% (>5%); drop_ns.mean -7.36% (>5%) |
| r2 | p512-k8-file | elapsed_ns.p50 -68.80% (>5%); elapsed_ns.p95 -68.57% (>10%); elapsed_ns.p99 -68.09% (>15%); elapsed_ns.mean -68.77% (>5%); open_ns.p50 -49.45% (>5%); open_ns.p95 -57.29% (>10%); open_ns.p99 -43.03% (>15%); open_ns.mean -49.58% (>5%); edit_ns.p50 -74.11% (>5%); edit_ns.p95 -74.03% (>10%); edit_ns.p99 -73.76% (>15%); edit_ns.mean -74.10% (>5%); commit_ns.p95 -17.31% (>10%); commit_ns.p99 -21.43% (>15%); commit_ns.mean -7.85% (>5%); publish_ns.p50 -12.83% (>5%); publish_ns.p95 -13.74% (>10%); publish_ns.mean -12.75% (>5%); drop_ns.p50 -5.38% (>5%); drop_ns.mean -5.43% (>5%); rss_kib.rss +7.45% (>5%) |
| r2 | p512-k8-owned | elapsed_ns.p50 -67.99% (>5%); elapsed_ns.p95 -68.00% (>10%); elapsed_ns.p99 -68.05% (>15%); elapsed_ns.mean -68.00% (>5%); open_ns.p50 +98.72% (>5%); open_ns.p95 +88.38% (>10%); open_ns.p99 +16.46% (>15%); open_ns.mean +76.04% (>5%); edit_ns.p50 -74.20% (>5%); edit_ns.p95 -74.12% (>10%); edit_ns.p99 -74.01% (>15%); edit_ns.mean -74.20% (>5%); publish_ns.mean +6.82% (>5%); drop_ns.mean -5.23% (>5%) |

## Per-case gauges and guards

The table exposes the Work charge, retained live reservations, release/output/input counters, and source reads for every case. Every listed guard is true; Work is expected unchanged and retained live-gauge deltas remain explicit comparison data.

| Campaign | Workload | API | after Work | live Work | live memory | live objects | after input | after output | source reads | guards |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| r1 | p128-k1-file-batch | batch | 6714084 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| r1 | p128-k1-file-repeated | repeated | 6714084 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| r1 | p128-k1-owned-batch | batch | 6714084 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| r1 | p128-k1-owned-repeated | repeated | 6714084 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| r1 | p128-k32-file-batch | batch | 7168451 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| r1 | p128-k32-file-repeated | repeated | 13132231 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| r1 | p128-k32-owned-batch | batch | 7168451 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| r1 | p128-k32-owned-repeated | repeated | 13132231 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| r1 | p128-k8-file-batch | batch | 6816683 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| r1 | p128-k8-file-repeated | repeated | 7497391 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| r1 | p128-k8-owned-batch | batch | 6816683 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| r1 | p128-k8-owned-repeated | repeated | 7497391 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| r1 | p512-k1-file-batch | batch | 103874532 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| r1 | p512-k1-file-repeated | repeated | 103874532 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| r1 | p512-k1-owned-batch | batch | 103874532 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| r1 | p512-k1-owned-repeated | repeated | 103874532 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| r1 | p512-k32-file-batch | batch | 104328899 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| r1 | p512-k32-file-repeated | repeated | 116292295 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| r1 | p512-k32-owned-batch | batch | 104328899 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| r1 | p512-k32-owned-repeated | repeated | 116292295 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| r1 | p512-k8-file-batch | batch | 103977131 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| r1 | p512-k8-file-repeated | repeated | 106012591 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |
| r1 | p512-k8-owned-batch | batch | 103977131 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| r1 | p512-k8-owned-repeated | repeated | 106012591 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |
| r2 | p128-k1-file-batch | batch | 6714084 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| r2 | p128-k1-file-repeated | repeated | 6714084 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| r2 | p128-k1-owned-batch | batch | 6714084 | 6699515 | 367535 | 14586 | 535175 | 532882 | 45 | all true |
| r2 | p128-k1-owned-repeated | repeated | 6714084 | 6699515 | 367517 | 14586 | 535175 | 532882 | 45 | all true |
| r2 | p128-k32-file-batch | batch | 7168451 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| r2 | p128-k32-file-repeated | repeated | 13132231 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| r2 | p128-k32-owned-batch | batch | 7168451 | 7153665 | 387716 | 14865 | 535175 | 533099 | 45 | all true |
| r2 | p128-k32-owned-repeated | repeated | 13132231 | 13117445 | 387140 | 14865 | 535175 | 533099 | 45 | all true |
| r2 | p128-k8-file-batch | batch | 6816683 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| r2 | p128-k8-file-repeated | repeated | 7497391 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| r2 | p128-k8-owned-batch | batch | 6816683 | 6802065 | 372092 | 14649 | 535175 | 532931 | 45 | all true |
| r2 | p128-k8-owned-repeated | repeated | 7497391 | 7482773 | 371948 | 14649 | 535175 | 532931 | 45 | all true |
| r2 | p512-k1-file-batch | batch | 103874532 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| r2 | p512-k1-file-repeated | repeated | 103874532 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| r2 | p512-k1-owned-batch | batch | 103874532 | 103816955 | 1442735 | 57594 | 556679 | 554386 | 47 | all true |
| r2 | p512-k1-owned-repeated | repeated | 103874532 | 103816955 | 1442717 | 57594 | 556679 | 554386 | 47 | all true |
| r2 | p512-k32-file-batch | batch | 104328899 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| r2 | p512-k32-file-repeated | repeated | 116292295 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| r2 | p512-k32-owned-batch | batch | 104328899 | 104271105 | 1462916 | 57873 | 556679 | 554603 | 47 | all true |
| r2 | p512-k32-owned-repeated | repeated | 116292295 | 116234501 | 1462340 | 57873 | 556679 | 554603 | 47 | all true |
| r2 | p512-k8-file-batch | batch | 103977131 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| r2 | p512-k8-file-repeated | repeated | 106012591 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |
| r2 | p512-k8-owned-batch | batch | 103977131 | 103919505 | 1447292 | 57657 | 556679 | 554435 | 47 | all true |
| r2 | p512-k8-owned-repeated | repeated | 106012591 | 105954965 | 1447148 | 57657 | 556679 | 554435 | 47 | all true |

## Cross-campaign flags

### Matched same API

20 of 24 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-file-batch | open_ns.p99 -16.88% (>15%) |
| p128-k1-file-repeated | open_ns.p99 -28.47% (>15%); commit_ns.p99 +43.24% (>15%); drop_ns.p99 -20.91% (>15%) |
| p128-k1-owned-batch | open_ns.p99 +50.25% (>15%); commit_ns.p50 +6.90% (>5%); commit_ns.mean +5.91% (>5%) |
| p128-k1-owned-repeated | elapsed_ns.p50 +9.92% (>5%); elapsed_ns.mean +9.10% (>5%); open_ns.p50 +5.75% (>5%); open_ns.mean +5.04% (>5%); commit_ns.p50 +10.34% (>5%); commit_ns.p95 +16.13% (>10%); commit_ns.p99 +17.65% (>15%); commit_ns.mean +10.11% (>5%); publish_ns.p50 +43.57% (>5%); publish_ns.p95 +42.30% (>10%); publish_ns.p99 +41.09% (>15%); publish_ns.mean +39.04% (>5%) |
| p128-k32-file-batch | drop_ns.p99 +188.07% (>15%); drop_ns.mean +9.28% (>5%) |
| p128-k32-file-repeated | open_ns.p99 +16.40% (>15%); commit_ns.p99 -28.93% (>15%); drop_ns.p99 +39.54% (>15%) |
| p128-k32-owned-batch | open_ns.p99 -20.15% (>15%); open_ns.mean -12.65% (>5%); drop_ns.p99 +226.60% (>15%); drop_ns.mean +5.97% (>5%) |
| p128-k8-file-batch | commit_ns.p99 +50.00% (>15%); rss_kib.rss -7.06% (>5%) |
| p128-k8-file-repeated | commit_ns.p50 +11.90% (>5%); commit_ns.p95 +46.81% (>10%); commit_ns.p99 +75.00% (>15%); commit_ns.mean +16.99% (>5%); drop_ns.p99 -82.54% (>15%); drop_ns.mean -8.52% (>5%) |
| p128-k8-owned-repeated | open_ns.p95 +103.46% (>10%); open_ns.p99 +109.56% (>15%); open_ns.mean +16.18% (>5%); commit_ns.p95 +56.52% (>10%); commit_ns.p99 +88.00% (>15%); commit_ns.mean +10.38% (>5%); publish_ns.p50 +5.40% (>5%); publish_ns.mean +13.57% (>5%) |
| p512-k1-file-repeated | commit_ns.p50 -5.88% (>5%); commit_ns.p99 -18.75% (>15%); commit_ns.mean -7.06% (>5%); publish_ns.p99 +26.27% (>15%) |
| p512-k1-owned-batch | drop_ns.p99 +58.33% (>15%) |
| p512-k32-file-batch | open_ns.p99 -28.78% (>15%); commit_ns.p99 +31.65% (>15%); drop_ns.p99 +60.76% (>15%) |
| p512-k32-file-repeated | commit_ns.p50 +5.74% (>5%); commit_ns.mean +5.15% (>5%); drop_ns.p99 -66.32% (>15%) |
| p512-k32-owned-batch | open_ns.p99 -18.13% (>15%); commit_ns.p99 -81.85% (>15%); commit_ns.mean -8.75% (>5%) |
| p512-k32-owned-repeated | open_ns.p50 +58.03% (>5%) |
| p512-k8-file-batch | commit_ns.p50 -10.87% (>5%); commit_ns.p95 -15.69% (>10%); commit_ns.p99 -18.52% (>15%); commit_ns.mean -11.31% (>5%); publish_ns.p50 -11.78% (>5%); publish_ns.p95 -11.02% (>10%); publish_ns.mean -9.35% (>5%) |
| p512-k8-file-repeated | open_ns.p95 +17.66% (>10%) |
| p512-k8-owned-batch | open_ns.mean -8.87% (>5%); commit_ns.p50 -9.09% (>5%); commit_ns.mean -8.06% (>5%) |
| p512-k8-owned-repeated | open_ns.p50 -51.44% (>5%); open_ns.p95 -49.71% (>10%); open_ns.p99 -23.73% (>15%); open_ns.mean -49.57% (>5%); commit_ns.p50 -15.22% (>5%); commit_ns.p95 -18.52% (>10%); commit_ns.p99 -35.21% (>15%); commit_ns.mean -15.08% (>5%); publish_ns.mean -6.71% (>5%) |

### API-choice ratio

12 of 12 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-file | open_ns.p99 +16.21% (>15%); commit_ns.p99 -34.07% (>15%); drop_ns.p99 +26.44% (>15%) |
| p128-k1-owned | elapsed_ns.p50 -7.90% (>5%); elapsed_ns.mean -7.38% (>5%); open_ns.p50 -6.27% (>5%); open_ns.p99 +56.17% (>15%); publish_ns.p50 -29.45% (>5%); publish_ns.p95 -28.84% (>10%); publish_ns.p99 -28.90% (>15%); publish_ns.mean -27.19% (>5%) |
| p128-k32-file | open_ns.p99 -16.04% (>15%); commit_ns.p99 +46.92% (>15%); publish_ns.p95 +11.64% (>10%); drop_ns.p99 +106.44% (>15%); drop_ns.mean +11.21% (>5%) |
| p128-k32-owned | open_ns.p99 -17.75% (>15%); open_ns.mean -10.59% (>5%); drop_ns.p99 +234.35% (>15%); drop_ns.mean +8.28% (>5%) |
| p128-k8-file | commit_ns.p50 -10.64% (>5%); commit_ns.p95 -33.43% (>10%); commit_ns.mean -13.81% (>5%); drop_ns.p99 +456.42% (>15%); drop_ns.mean +7.89% (>5%); rss_kib.rss -7.12% (>5%) |
| p128-k8-owned | open_ns.p95 -50.81% (>10%); open_ns.p99 -51.40% (>15%); open_ns.mean -13.22% (>5%); commit_ns.p95 -33.27% (>10%); commit_ns.p99 -42.79% (>15%); commit_ns.mean -5.82% (>5%); publish_ns.mean -11.56% (>5%) |
| p512-k1-file | open_ns.p99 +24.60% (>15%); commit_ns.p50 +6.25% (>5%); commit_ns.p95 +11.36% (>10%); commit_ns.p99 +20.21% (>15%); commit_ns.mean +8.16% (>5%); publish_ns.p99 -23.41% (>15%); drop_ns.p99 +23.45% (>15%) |
| p512-k1-owned | drop_ns.p99 +60.07% (>15%) |
| p512-k32-file | open_ns.p99 -30.93% (>15%); commit_ns.p99 +40.48% (>15%); drop_ns.p99 +377.37% (>15%) |
| p512-k32-owned | open_ns.p50 -37.88% (>5%); open_ns.p99 -15.34% (>15%); commit_ns.p50 -5.50% (>5%); commit_ns.p99 -81.35% (>15%); commit_ns.mean -11.35% (>5%) |
| p512-k8-file | open_ns.p95 -19.81% (>10%); open_ns.mean -7.15% (>5%); commit_ns.p50 -8.80% (>5%); commit_ns.p95 -18.93% (>10%); commit_ns.p99 -24.34% (>15%); commit_ns.mean -9.97% (>5%); publish_ns.p50 -10.96% (>5%); publish_ns.p95 -10.21% (>10%); publish_ns.mean -8.44% (>5%) |
| p512-k8-owned | open_ns.p50 +103.94% (>5%); open_ns.p95 +96.16% (>10%); open_ns.p99 +23.97% (>15%); open_ns.mean +80.70% (>5%); commit_ns.p50 +7.23% (>5%); commit_ns.p95 +10.45% (>10%); commit_ns.p99 +36.54% (>15%); commit_ns.mean +8.27% (>5%); publish_ns.mean +7.34% (>5%); drop_ns.p50 +5.68% (>5%) |


## Counter drift

Same-API output identities match for 24/24 cases across campaigns.
Release/output/read guard counter changes: 0 cases.
Allowed retained live-gauge changes: 0 cases; Work changes would be guard findings. Any values appear in JSON.

All case/campaign phase statistics, repeat groups, route ratios, counter drift, identities, guards, and bootstrap intervals are in `baseline-native-analysis.json`.
