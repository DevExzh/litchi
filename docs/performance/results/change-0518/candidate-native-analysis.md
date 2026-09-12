# 0518 native DOCX analysis

Native campaigns: after-r1, after-r2. Rows are source-bound through each receipt; profile and hardware lanes are excluded.
Measured rows are the 30 `warmup=false` samples in each of two internal repeats. Every case reports nearest-rank p50/p95/p99 and arithmetic mean for elapsed, open, edit, commit, publish, and drop clocks; RSS is the one whole-child maximum from GNU time.
The unpaired route bootstrap for the p50 ratio uses seed `5343761` and 4000 iterations, resampling each route within each internal repeat and taking batch median / repeated median. With only two internal repeats, its interval is descriptive for these row distributions and does not estimate independent-process or host variation.
The historical 0500 validator is imported by `capture.py` and remains the row-level source for release, output, semantic, untouched-member, source-version, and readback guards. Work and retained live gauges are reported as descriptive counters; they do not relax those guards.

## API-choice ratios

Ratios are batch / repeated; positive percentages mean batch took longer or used more RSS. The same absolute thresholds are used for route flags and cross-campaign flags: 5% for p50/mean/RSS, 10% for p95, and 15% for p99.

| Campaign | Workload | elapsed p50 | edit p50 | publish p50 | RSS | route flags |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| after-r1 | p128-k1-file | 1.82% | 3.34% | -0.06% | -0.12% | 1 |
| after-r1 | p128-k1-owned | 0.46% | -0.15% | 1.28% | 0.00% | 1 |
| after-r1 | p128-k32-file | -88.77% | -92.05% | -1.53% | -1.20% | 10 |
| after-r1 | p128-k32-owned | -89.23% | -92.01% | -2.13% | -0.39% | 10 |
| after-r1 | p128-k8-file | -66.34% | -73.88% | -24.94% | 0.00% | 21 |
| after-r1 | p128-k8-owned | -68.18% | -73.43% | -31.33% | -0.71% | 19 |
| after-r1 | p512-k1-file | 0.30% | 0.48% | -0.80% | -0.06% | 2 |
| after-r1 | p512-k1-owned | 0.48% | 0.66% | -0.45% | 0.34% | 6 |
| after-r1 | p512-k32-file | -90.29% | -92.52% | -0.52% | 0.23% | 14 |
| after-r1 | p512-k32-owned | -90.31% | -92.41% | -0.06% | 0.22% | 13 |
| after-r1 | p512-k8-file | -68.08% | -74.19% | -6.62% | 5.45% | 18 |
| after-r1 | p512-k8-owned | -68.13% | -74.08% | -0.66% | 4.10% | 14 |
| after-r2 | p128-k1-file | -0.06% | -0.80% | -0.27% | -0.73% | 3 |
| after-r2 | p128-k1-owned | -1.25% | -3.31% | 0.86% | 0.50% | 4 |
| after-r2 | p128-k32-file | -88.97% | -92.15% | -1.67% | 1.33% | 11 |
| after-r2 | p128-k32-owned | -89.18% | -91.95% | -2.11% | -0.11% | 13 |
| after-r2 | p128-k8-file | -66.21% | -73.82% | -24.33% | -0.90% | 17 |
| after-r2 | p128-k8-owned | -68.17% | -73.43% | -32.17% | 3.49% | 18 |
| after-r2 | p512-k1-file | -0.00% | 0.00% | 0.27% | -0.56% | 0 |
| after-r2 | p512-k1-owned | 0.86% | 0.99% | 0.40% | 0.00% | 0 |
| after-r2 | p512-k32-file | -90.38% | -92.52% | -0.94% | 0.58% | 10 |
| after-r2 | p512-k32-owned | -90.33% | -92.43% | -1.27% | 0.87% | 16 |
| after-r2 | p512-k8-file | -69.10% | -74.37% | -12.95% | 3.70% | 20 |
| after-r2 | p512-k8-owned | -68.29% | -74.01% | -1.02% | 0.50% | 15 |

## API-choice flags

22 route records have one or more absolute threshold flags.

| Campaign | Workload | flags |
| --- | --- | --- |
| after-r1 | p128-k1-file | open_ns.p99 -24.22% (>15%) |
| after-r1 | p128-k1-owned | open_ns.p99 +49.85% (>15%) |
| after-r1 | p128-k32-file | elapsed_ns.p50 -88.77% (>5%); elapsed_ns.p95 -88.68% (>10%); elapsed_ns.p99 -88.65% (>15%); elapsed_ns.mean -88.83% (>5%); edit_ns.p50 -92.05% (>5%); edit_ns.p95 -91.95% (>10%); edit_ns.p99 -91.88% (>15%); edit_ns.mean -92.04% (>5%); drop_ns.p50 -7.75% (>5%); drop_ns.mean -6.71% (>5%) |
| after-r1 | p128-k32-owned | elapsed_ns.p50 -89.23% (>5%); elapsed_ns.p95 -89.04% (>10%); elapsed_ns.p99 -88.91% (>15%); elapsed_ns.mean -89.23% (>5%); open_ns.p50 -5.80% (>5%); open_ns.p95 +22.51% (>10%); edit_ns.p50 -92.01% (>5%); edit_ns.p95 -91.92% (>10%); edit_ns.p99 -91.87% (>15%); edit_ns.mean -92.00% (>5%) |
| after-r1 | p128-k8-file | elapsed_ns.p50 -66.34% (>5%); elapsed_ns.p95 -66.09% (>10%); elapsed_ns.p99 -64.96% (>15%); elapsed_ns.mean -66.08% (>5%); open_ns.p50 -7.54% (>5%); open_ns.mean -7.01% (>5%); edit_ns.p50 -73.88% (>5%); edit_ns.p95 -73.41% (>10%); edit_ns.p99 -73.24% (>15%); edit_ns.mean -73.72% (>5%); commit_ns.p50 -11.36% (>5%); commit_ns.p95 -14.29% (>10%); commit_ns.p99 -21.43% (>15%); commit_ns.mean -10.74% (>5%); publish_ns.p50 -24.94% (>5%); publish_ns.p95 -22.91% (>10%); publish_ns.p99 -16.94% (>15%); publish_ns.mean -20.20% (>5%); drop_ns.p50 -11.28% (>5%); drop_ns.p95 -13.19% (>10%); drop_ns.p99 +471.62% (>15%) |
| after-r1 | p128-k8-owned | elapsed_ns.p50 -68.18% (>5%); elapsed_ns.p95 -68.10% (>10%); elapsed_ns.p99 -68.10% (>15%); elapsed_ns.mean -68.24% (>5%); open_ns.p50 -6.99% (>5%); open_ns.p95 -52.08% (>10%); open_ns.p99 -35.20% (>15%); open_ns.mean -14.34% (>5%); edit_ns.p50 -73.43% (>5%); edit_ns.p95 -73.05% (>10%); edit_ns.p99 -73.13% (>15%); edit_ns.mean -73.55% (>5%); publish_ns.p50 -31.33% (>5%); publish_ns.p95 -29.38% (>10%); publish_ns.p99 -31.85% (>15%); publish_ns.mean -31.23% (>5%); drop_ns.p50 -12.98% (>5%); drop_ns.p95 -12.14% (>10%); drop_ns.p99 +435.33% (>15%) |
| after-r1 | p512-k1-file | commit_ns.p50 +6.06% (>5%); commit_ns.mean +5.45% (>5%) |
| after-r1 | p512-k1-owned | open_ns.p95 +23.46% (>10%); open_ns.p99 +21.38% (>15%); commit_ns.p50 +9.68% (>5%); commit_ns.p95 +21.21% (>10%); commit_ns.p99 +32.43% (>15%); commit_ns.mean +10.74% (>5%) |
| after-r1 | p512-k32-file | elapsed_ns.p50 -90.29% (>5%); elapsed_ns.p95 -90.20% (>10%); elapsed_ns.p99 -90.11% (>15%); elapsed_ns.mean -90.36% (>5%); open_ns.p99 +28.94% (>15%); edit_ns.p50 -92.52% (>5%); edit_ns.p95 -92.33% (>10%); edit_ns.p99 -92.22% (>15%); edit_ns.mean -92.48% (>5%); commit_ns.p95 +10.87% (>10%); commit_ns.p99 +354.79% (>15%); commit_ns.mean +10.88% (>5%); drop_ns.p50 -7.45% (>5%); drop_ns.mean -7.81% (>5%) |
| after-r1 | p512-k32-owned | elapsed_ns.p50 -90.31% (>5%); elapsed_ns.p95 -90.24% (>10%); elapsed_ns.p99 -89.32% (>15%); elapsed_ns.mean -90.29% (>5%); edit_ns.p50 -92.41% (>5%); edit_ns.p95 -92.37% (>10%); edit_ns.p99 -91.52% (>15%); edit_ns.mean -92.39% (>5%); commit_ns.p50 -8.40% (>5%); commit_ns.mean -7.52% (>5%); drop_ns.p50 -8.16% (>5%); drop_ns.p99 +35.39% (>15%); drop_ns.mean -7.88% (>5%) |
| after-r1 | p512-k8-file | elapsed_ns.p50 -68.08% (>5%); elapsed_ns.p95 -67.93% (>10%); elapsed_ns.p99 -67.95% (>15%); elapsed_ns.mean -68.33% (>5%); open_ns.p50 -46.81% (>5%); open_ns.p95 -47.18% (>10%); open_ns.p99 -37.52% (>15%); open_ns.mean -46.77% (>5%); edit_ns.p50 -74.19% (>5%); edit_ns.p95 -73.88% (>10%); edit_ns.p99 -73.37% (>15%); edit_ns.mean -74.14% (>5%); commit_ns.p95 +12.50% (>10%); publish_ns.p50 -6.62% (>5%); publish_ns.mean -6.51% (>5%); drop_ns.p50 -6.20% (>5%); drop_ns.mean -6.72% (>5%); rss_kib.rss +5.45% (>5%) |
| after-r1 | p512-k8-owned | elapsed_ns.p50 -68.13% (>5%); elapsed_ns.p95 -68.08% (>10%); elapsed_ns.p99 -68.04% (>15%); elapsed_ns.mean -68.12% (>5%); open_ns.p99 -17.22% (>15%); open_ns.mean -6.47% (>5%); edit_ns.p50 -74.08% (>5%); edit_ns.p95 -73.90% (>10%); edit_ns.p99 -73.88% (>15%); edit_ns.mean -74.05% (>5%); commit_ns.p50 -6.82% (>5%); commit_ns.mean -6.48% (>5%); drop_ns.p50 -6.15% (>5%); drop_ns.mean -6.15% (>5%) |
| after-r2 | p128-k1-file | open_ns.p99 +37.43% (>15%); commit_ns.p99 -91.21% (>15%); commit_ns.mean -14.40% (>5%) |
| after-r2 | p128-k1-owned | open_ns.p99 -22.81% (>15%); commit_ns.p99 -15.56% (>15%); drop_ns.p99 +822.22% (>15%); drop_ns.mean +16.43% (>5%) |
| after-r2 | p128-k32-file | elapsed_ns.p50 -88.97% (>5%); elapsed_ns.p95 -88.84% (>10%); elapsed_ns.p99 -88.79% (>15%); elapsed_ns.mean -89.06% (>5%); edit_ns.p50 -92.15% (>5%); edit_ns.p95 -92.04% (>10%); edit_ns.p99 -92.04% (>15%); edit_ns.mean -92.13% (>5%); drop_ns.p50 -7.09% (>5%); drop_ns.p99 -68.71% (>15%); drop_ns.mean -11.36% (>5%) |
| after-r2 | p128-k32-owned | elapsed_ns.p50 -89.18% (>5%); elapsed_ns.p95 -88.91% (>10%); elapsed_ns.p99 -88.89% (>15%); elapsed_ns.mean -89.15% (>5%); open_ns.p99 +21.68% (>15%); edit_ns.p50 -91.95% (>5%); edit_ns.p95 -91.87% (>10%); edit_ns.p99 -91.83% (>15%); edit_ns.mean -91.96% (>5%); commit_ns.p50 -5.51% (>5%); commit_ns.p99 +278.52% (>15%); drop_ns.p50 -6.88% (>5%); drop_ns.p99 +38.66% (>15%) |
| after-r2 | p128-k8-file | elapsed_ns.p50 -66.21% (>5%); elapsed_ns.p95 -64.87% (>10%); elapsed_ns.p99 -64.25% (>15%); elapsed_ns.mean -66.10% (>5%); open_ns.p50 -8.45% (>5%); open_ns.mean -8.27% (>5%); edit_ns.p50 -73.82% (>5%); edit_ns.p95 -73.33% (>10%); edit_ns.p99 -73.24% (>15%); edit_ns.mean -73.71% (>5%); commit_ns.p50 -9.09% (>5%); commit_ns.mean -7.64% (>5%); publish_ns.p50 -24.33% (>5%); publish_ns.p95 -14.65% (>10%); publish_ns.mean -21.62% (>5%); drop_ns.p50 -7.58% (>5%); drop_ns.p99 +370.52% (>15%) |
| after-r2 | p128-k8-owned | elapsed_ns.p50 -68.17% (>5%); elapsed_ns.p95 -67.89% (>10%); elapsed_ns.p99 -67.44% (>15%); elapsed_ns.mean -68.17% (>5%); open_ns.p50 -6.36% (>5%); open_ns.p95 -51.94% (>10%); open_ns.p99 -51.44% (>15%); open_ns.mean -15.32% (>5%); edit_ns.p50 -73.43% (>5%); edit_ns.p95 -73.03% (>10%); edit_ns.p99 -72.62% (>15%); edit_ns.mean -73.39% (>5%); publish_ns.p50 -32.17% (>5%); publish_ns.p95 -29.80% (>10%); publish_ns.p99 -31.93% (>15%); publish_ns.mean -32.10% (>5%); drop_ns.p50 -8.59% (>5%); drop_ns.mean -8.78% (>5%) |
| after-r2 | p512-k32-file | elapsed_ns.p50 -90.38% (>5%); elapsed_ns.p95 -90.23% (>10%); elapsed_ns.p99 -90.28% (>15%); elapsed_ns.mean -90.39% (>5%); edit_ns.p50 -92.52% (>5%); edit_ns.p95 -92.48% (>10%); edit_ns.p99 -92.50% (>15%); edit_ns.mean -92.52% (>5%); drop_ns.p50 -8.10% (>5%); drop_ns.mean -7.40% (>5%) |
| after-r2 | p512-k32-owned | elapsed_ns.p50 -90.33% (>5%); elapsed_ns.p95 -90.34% (>10%); elapsed_ns.p99 -90.35% (>15%); elapsed_ns.mean -90.34% (>5%); open_ns.p50 -23.28% (>5%); open_ns.p99 -19.17% (>15%); open_ns.mean -23.96% (>5%); edit_ns.p50 -92.43% (>5%); edit_ns.p95 -92.38% (>10%); edit_ns.p99 -92.39% (>15%); edit_ns.mean -92.43% (>5%); commit_ns.p99 -76.38% (>15%); commit_ns.mean -5.93% (>5%); publish_ns.p99 -20.81% (>15%); drop_ns.p50 -6.67% (>5%); drop_ns.p99 +221.24% (>15%) |
| after-r2 | p512-k8-file | elapsed_ns.p50 -69.10% (>5%); elapsed_ns.p95 -68.93% (>10%); elapsed_ns.p99 -68.84% (>15%); elapsed_ns.mean -69.04% (>5%); open_ns.p50 -49.09% (>5%); open_ns.p95 -50.31% (>10%); open_ns.p99 -57.47% (>15%); open_ns.mean -47.37% (>5%); edit_ns.p50 -74.37% (>5%); edit_ns.p95 -74.28% (>10%); edit_ns.p99 -74.20% (>15%); edit_ns.mean -74.35% (>5%); commit_ns.p50 -13.04% (>5%); commit_ns.p95 -16.67% (>10%); commit_ns.mean -12.90% (>5%); publish_ns.p50 -12.95% (>5%); publish_ns.p95 -15.31% (>10%); publish_ns.mean -12.95% (>5%); drop_ns.p50 -5.43% (>5%); drop_ns.p99 +196.64% (>15%) |
| after-r2 | p512-k8-owned | elapsed_ns.p50 -68.29% (>5%); elapsed_ns.p95 -68.14% (>10%); elapsed_ns.p99 -69.00% (>15%); elapsed_ns.mean -68.29% (>5%); open_ns.p50 -36.95% (>5%); open_ns.p95 -24.09% (>10%); open_ns.p99 -38.02% (>15%); open_ns.mean -37.04% (>5%); edit_ns.p50 -74.01% (>5%); edit_ns.p95 -73.90% (>10%); edit_ns.p99 -74.71% (>15%); edit_ns.mean -74.01% (>5%); commit_ns.p99 +30.36% (>15%); drop_ns.p50 -7.69% (>5%); drop_ns.mean -7.59% (>5%) |

## Per-case gauges and guards

The table exposes the Work charge, retained live reservations, release/output/input counters, and source reads for every case. Every listed guard is true; Work and live-gauge deltas remain eligible for explicit candidate comparison.

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

20 of 24 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-file-batch | commit_ns.p50 +6.45% (>5%) |
| p128-k1-file-repeated | open_ns.p99 -36.67% (>15%); commit_ns.p99 +1007.89% (>15%); commit_ns.mean +22.46% (>5%) |
| p128-k1-owned-batch | elapsed_ns.p50 +11.66% (>5%); elapsed_ns.p95 +10.48% (>10%); elapsed_ns.mean +10.27% (>5%); open_ns.p50 +5.13% (>5%); commit_ns.p50 +10.00% (>5%); commit_ns.p95 +12.12% (>10%); commit_ns.mean +8.37% (>5%); publish_ns.p50 +44.39% (>5%); publish_ns.p95 +42.05% (>10%); publish_ns.p99 +35.20% (>15%); publish_ns.mean +39.48% (>5%); drop_ns.p99 +854.02% (>15%); drop_ns.mean +19.74% (>5%) |
| p128-k1-owned-repeated | elapsed_ns.p50 +13.60% (>5%); elapsed_ns.p95 +13.59% (>10%); elapsed_ns.mean +12.52% (>5%); open_ns.p50 +6.05% (>5%); open_ns.p99 +70.79% (>15%); open_ns.mean +6.67% (>5%); commit_ns.p50 +6.67% (>5%); commit_ns.p99 +36.36% (>15%); publish_ns.p50 +44.99% (>5%); publish_ns.p95 +43.77% (>10%); publish_ns.p99 +44.44% (>15%); publish_ns.mean +40.02% (>5%) |
| p128-k32-file-repeated | drop_ns.p99 +201.65% (>15%) |
| p128-k32-owned-batch | open_ns.p99 +27.73% (>15%); open_ns.mean +13.18% (>5%); commit_ns.p99 +294.41% (>15%); commit_ns.mean +6.44% (>5%); drop_ns.p99 +279.42% (>15%) |
| p128-k32-owned-repeated | open_ns.p95 +26.85% (>10%); open_ns.mean +13.00% (>5%); drop_ns.p99 +187.50% (>15%) |
| p128-k8-file-batch | publish_ns.p95 +10.14% (>10%) |
| p128-k8-file-repeated | open_ns.p99 +23.24% (>15%); drop_ns.p99 +16.89% (>15%) |
| p128-k8-owned-batch | open_ns.p99 -34.85% (>15%); commit_ns.p99 -17.86% (>15%); drop_ns.p99 -83.19% (>15%); drop_ns.mean -7.19% (>5%) |
| p512-k1-file-repeated | commit_ns.p50 +6.06% (>5%) |
| p512-k1-owned-batch | open_ns.p95 -17.48% (>10%); commit_ns.p99 -16.33% (>15%) |
| p512-k1-owned-repeated | commit_ns.p50 +6.45% (>5%); commit_ns.p95 +12.12% (>10%); commit_ns.mean +6.49% (>5%) |
| p512-k32-file-batch | commit_ns.p99 -76.96% (>15%); commit_ns.mean -6.69% (>5%) |
| p512-k32-file-repeated | open_ns.p99 +32.43% (>15%); commit_ns.p50 +8.26% (>5%); commit_ns.mean +6.75% (>5%) |
| p512-k32-owned-batch | open_ns.p50 +54.78% (>5%); open_ns.p99 -29.22% (>15%); commit_ns.mean +5.30% (>5%); drop_ns.p99 +135.73% (>15%); drop_ns.mean +6.83% (>5%) |
| p512-k32-owned-repeated | open_ns.p50 +92.22% (>5%); open_ns.mean +27.77% (>5%); commit_ns.p99 +342.31% (>15%); publish_ns.p99 +22.28% (>15%) |
| p512-k8-file-batch | open_ns.p99 -30.62% (>15%); commit_ns.p50 -11.11% (>5%); commit_ns.p95 -16.67% (>10%); commit_ns.mean -10.32% (>5%); publish_ns.p50 -6.68% (>5%); publish_ns.p95 -11.79% (>10%); publish_ns.mean -6.31% (>5%); drop_ns.p99 +217.99% (>15%) |
| p512-k8-file-repeated | commit_ns.p50 +6.98% (>5%); commit_ns.p95 +12.50% (>10%); commit_ns.p99 +19.23% (>15%); commit_ns.mean +7.80% (>5%) |
| p512-k8-owned-batch | open_ns.p50 -36.11% (>5%); open_ns.p95 -22.04% (>10%); open_ns.p99 -24.06% (>15%); open_ns.mean -33.00% (>5%); commit_ns.p50 +9.76% (>5%); commit_ns.p99 +46.00% (>15%); commit_ns.mean +8.92% (>5%) |

### API-choice ratio

11 of 12 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-file | open_ns.p95 +11.49% (>10%); open_ns.p99 +81.35% (>15%); commit_ns.p99 -91.21% (>15%); commit_ns.mean -15.30% (>5%) |
| p128-k1-owned | open_ns.p99 -48.48% (>15%); commit_ns.p99 -18.04% (>15%); drop_ns.p99 +843.42% (>15%); drop_ns.mean +19.00% (>5%) |
| p128-k32-file | drop_ns.p99 -67.42% (>15%) |
| p128-k32-owned | open_ns.p95 -19.89% (>10%); open_ns.p99 +25.71% (>15%); commit_ns.p99 +267.94% (>15%); drop_ns.p99 +31.97% (>15%) |
| p128-k8-file | open_ns.p99 -16.63% (>15%); publish_ns.p95 +10.72% (>10%); drop_ns.p99 -17.69% (>15%) |
| p128-k8-owned | open_ns.p99 -25.06% (>15%); drop_ns.p50 +5.04% (>5%); drop_ns.p99 -82.61% (>15%); drop_ns.mean -5.92% (>5%) |
| p512-k1-owned | open_ns.p95 -18.55% (>10%); open_ns.p99 -23.83% (>15%); commit_ns.p50 -6.06% (>5%); commit_ns.p95 -13.04% (>10%); commit_ns.p99 -24.49% (>15%); commit_ns.mean -7.85% (>5%) |
| p512-k32-file | open_ns.p99 -25.02% (>15%); commit_ns.p50 -6.89% (>5%); commit_ns.p95 -12.87% (>10%); commit_ns.p99 -78.01% (>15%); commit_ns.mean -12.59% (>5%) |
| p512-k32-owned | open_ns.p50 -19.48% (>5%); open_ns.p99 -29.03% (>15%); open_ns.mean -21.91% (>5%); commit_ns.p50 +7.45% (>5%); commit_ns.p99 -76.07% (>15%); publish_ns.p99 -22.15% (>15%); drop_ns.p99 +137.27% (>15%); drop_ns.mean +5.93% (>5%) |
| p512-k8-file | open_ns.p99 -31.93% (>15%); commit_ns.p50 -16.91% (>5%); commit_ns.p95 -25.93% (>10%); commit_ns.mean -16.81% (>5%); publish_ns.p50 -6.78% (>5%); publish_ns.p95 -14.17% (>10%); publish_ns.mean -6.89% (>5%); drop_ns.p99 +213.72% (>15%); drop_ns.mean +5.17% (>5%) |
| p512-k8-owned | open_ns.p50 -36.11% (>5%); open_ns.p95 -22.18% (>10%); open_ns.p99 -25.13% (>15%); open_ns.mean -32.68% (>5%); commit_ns.p50 +7.32% (>5%); commit_ns.p99 +30.36% (>15%); commit_ns.mean +6.42% (>5%) |


## Counter drift

Same-API output identities match for 24/24 cases across campaigns.
Release/output/read guard counter changes: 0 cases.
Allowed Work/live-gauge changes: 0 cases; any values appear in JSON.

All case/campaign phase statistics, repeat groups, route ratios, counter drift, identities, guards, and bootstrap intervals are in `candidate-native-analysis.json`.
