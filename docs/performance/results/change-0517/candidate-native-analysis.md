# 0517 native DOCX analysis

Native campaigns: after-r1, after-r2. Rows are source-bound through each receipt; profile and hardware lanes are excluded.
Measured rows are the 30 `warmup=false` samples in each of two internal repeats. Phase percentiles use nearest rank; means are arithmetic means; RSS is the one whole-child maximum from GNU time.
The unpaired route bootstrap for the p50 ratio uses seed `5339665` and 4000 iterations, resampling each route within each internal repeat and taking batch median / repeated median. With only two internal repeats, its interval is descriptive for these row distributions and does not estimate independent-process or host variation.

## API-choice ratios

Ratios are batch / repeated; positive percentages mean batch took longer or used more RSS. The same absolute thresholds are used for route flags and cross-campaign flags: 5% for p50/mean/RSS, 10% for p95, and 15% for p99.

| Campaign | Workload | elapsed p50 | edit p50 | publish p50 | RSS | route flags |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| after-r1 | p128-k1-file | -0.48% | -1.00% | -0.01% | -4.52% | 3 |
| after-r1 | p128-k1-owned | -7.84% | -0.22% | -15.35% | 4.24% | 6 |
| after-r1 | p128-k32-file | -86.63% | -92.16% | -1.60% | 0.54% | 10 |
| after-r1 | p128-k32-owned | -86.70% | -91.99% | -0.39% | 0.78% | 11 |
| after-r1 | p128-k8-file | -58.67% | -73.78% | -1.58% | 0.12% | 14 |
| after-r1 | p128-k8-owned | -60.20% | -73.75% | -0.58% | -0.55% | 11 |
| after-r1 | p512-k1-file | -0.25% | -0.29% | -0.39% | -2.99% | 5 |
| after-r1 | p512-k1-owned | -0.18% | -0.19% | -0.28% | 3.21% | 4 |
| after-r1 | p512-k32-file | -87.36% | -92.65% | -0.54% | 0.69% | 14 |
| after-r1 | p512-k32-owned | -87.26% | -92.54% | 0.07% | -2.22% | 15 |
| after-r1 | p512-k8-file | -60.96% | -74.17% | -5.17% | 2.67% | 19 |
| after-r1 | p512-k8-owned | -60.24% | -74.12% | -0.54% | 5.01% | 14 |
| after-r2 | p128-k1-file | -0.39% | -0.29% | 0.87% | -0.73% | 0 |
| after-r2 | p128-k1-owned | 8.17% | -0.88% | 16.76% | -0.84% | 10 |
| after-r2 | p128-k32-file | -86.71% | -92.19% | -1.41% | -0.78% | 10 |
| after-r2 | p128-k32-owned | -86.79% | -92.09% | -0.42% | 0.34% | 13 |
| after-r2 | p128-k8-file | -59.86% | -73.80% | -13.21% | 2.23% | 17 |
| after-r2 | p128-k8-owned | -61.30% | -73.52% | -16.41% | 1.33% | 19 |
| after-r2 | p512-k1-file | 0.12% | 0.61% | 0.02% | -2.44% | 2 |
| after-r2 | p512-k1-owned | 0.17% | -0.24% | 0.38% | 2.36% | 5 |
| after-r2 | p512-k32-file | -87.24% | -92.56% | 1.74% | -0.34% | 10 |
| after-r2 | p512-k32-owned | -87.30% | -92.58% | -0.76% | 1.95% | 13 |
| after-r2 | p512-k8-file | -61.05% | -74.16% | -5.80% | 4.54% | 22 |
| after-r2 | p512-k8-owned | -59.59% | -73.33% | 0.66% | 1.76% | 13 |

## Cross-campaign flags

### Matched same API

20 of 24 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-owned-batch | elapsed_ns.p50 +8.63% (>5%); elapsed_ns.mean +7.82% (>5%); open_ns.p99 +62.61% (>15%); commit_ns.p50 +5.88% (>5%); commit_ns.p95 +11.11% (>10%); commit_ns.mean +6.82% (>5%); publish_ns.p50 +17.89% (>5%); publish_ns.p95 +18.60% (>10%); publish_ns.p99 +19.09% (>15%); publish_ns.mean +16.66% (>5%) |
| p128-k1-owned-repeated | elapsed_ns.p50 -7.45% (>5%); elapsed_ns.mean -7.42% (>5%); publish_ns.p50 -14.53% (>5%); publish_ns.p95 -15.14% (>10%); publish_ns.p99 -18.16% (>15%); publish_ns.mean -14.20% (>5%) |
| p128-k32-file-batch | drop_ns.p99 -76.56% (>15%); drop_ns.mean -7.16% (>5%) |
| p128-k32-owned-batch | open_ns.p99 +25.39% (>15%) |
| p128-k32-owned-repeated | open_ns.p95 -46.40% (>10%) |
| p128-k8-file-batch | open_ns.p99 -18.53% (>15%); commit_ns.p50 +5.00% (>5%); commit_ns.p99 -34.25% (>15%) |
| p128-k8-file-repeated | open_ns.p99 -29.51% (>15%); commit_ns.p50 +6.98% (>5%); commit_ns.p99 -90.86% (>15%); commit_ns.mean -11.35% (>5%); publish_ns.p50 +13.80% (>5%); publish_ns.p95 +14.02% (>10%); publish_ns.mean +11.27% (>5%) |
| p128-k8-owned-batch | open_ns.p99 -23.45% (>15%); commit_ns.p50 +10.00% (>5%); commit_ns.mean +8.89% (>5%); drop_ns.p99 -31.58% (>15%) |
| p128-k8-owned-repeated | open_ns.p50 +5.28% (>5%); open_ns.p95 +105.71% (>10%); open_ns.p99 +41.55% (>15%); open_ns.mean +28.72% (>5%); commit_ns.p50 +7.14% (>5%); commit_ns.mean +9.04% (>5%); publish_ns.p50 +18.86% (>5%); publish_ns.p95 +17.95% (>10%); publish_ns.p99 +16.61% (>15%); publish_ns.mean +18.29% (>5%); drop_ns.p50 +5.60% (>5%) |
| p512-k1-file-batch | open_ns.mean +9.47% (>5%); commit_ns.p99 +23.81% (>15%) |
| p512-k1-file-repeated | open_ns.p99 +37.49% (>15%); commit_ns.p50 -12.20% (>5%); commit_ns.p99 -55.45% (>15%); commit_ns.mean -12.01% (>5%) |
| p512-k1-owned-batch | commit_ns.p50 +5.71% (>5%); commit_ns.mean +5.51% (>5%) |
| p512-k1-owned-repeated | open_ns.p99 -18.05% (>15%) |
| p512-k32-file-batch | open_ns.p99 +30.12% (>15%); commit_ns.p99 -16.40% (>15%) |
| p512-k32-file-repeated | commit_ns.p99 +16.89% (>15%) |
| p512-k32-owned-batch | open_ns.p50 +8.52% (>5%); drop_ns.p99 +196.97% (>15%) |
| p512-k32-owned-repeated | open_ns.p50 -13.21% (>5%); commit_ns.p99 +16.13% (>15%); drop_ns.p99 +23.21% (>15%) |
| p512-k8-file-batch | open_ns.p99 -22.82% (>15%); commit_ns.p99 +53.23% (>15%); drop_ns.p99 -23.83% (>15%) |
| p512-k8-file-repeated | open_ns.p99 -16.38% (>15%); open_ns.mean +6.06% (>5%); commit_ns.p95 +12.00% (>10%); commit_ns.p99 +444.23% (>15%); commit_ns.mean +14.05% (>5%); drop_ns.p99 +222.07% (>15%); drop_ns.mean +6.06% (>5%) |
| p512-k8-owned-batch | open_ns.p50 -48.26% (>5%); open_ns.p95 -22.20% (>10%); open_ns.mean -32.20% (>5%); commit_ns.p99 -85.47% (>15%); commit_ns.mean -15.64% (>5%); drop_ns.p99 -21.51% (>15%) |

### API-choice ratio

11 of 12 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-owned | elapsed_ns.p50 +17.38% (>5%); elapsed_ns.p95 +18.12% (>10%); elapsed_ns.p99 +23.88% (>15%); elapsed_ns.mean +16.45% (>5%); open_ns.p50 +7.86% (>5%); open_ns.p99 +58.06% (>15%); open_ns.mean +8.46% (>5%); commit_ns.p50 +9.19% (>5%); commit_ns.p95 +17.28% (>10%); commit_ns.p99 +16.58% (>15%); commit_ns.mean +12.35% (>5%); publish_ns.p50 +37.94% (>5%); publish_ns.p95 +39.75% (>10%); publish_ns.p99 +45.51% (>15%); publish_ns.mean +35.96% (>5%); drop_ns.p50 +5.30% (>5%); drop_ns.mean +5.61% (>5%) |
| p128-k32-file | drop_ns.p99 -77.17% (>15%); drop_ns.mean -7.56% (>5%) |
| p128-k32-owned | open_ns.p95 +87.17% (>10%); open_ns.p99 +28.43% (>15%); open_ns.mean +7.34% (>5%) |
| p128-k8-file | open_ns.p50 -6.00% (>5%); open_ns.p99 +15.57% (>15%); commit_ns.p99 +619.63% (>15%); commit_ns.mean +16.75% (>5%); publish_ns.p50 -11.82% (>5%); publish_ns.p95 -12.32% (>10%); publish_ns.mean -10.03% (>5%) |
| p128-k8-owned | open_ns.p50 -5.47% (>5%); open_ns.p95 -50.94% (>10%); open_ns.p99 -45.92% (>15%); open_ns.mean -22.72% (>5%); publish_ns.p50 -15.93% (>5%); publish_ns.p95 -13.55% (>10%); publish_ns.mean -15.16% (>5%); drop_ns.p50 -6.11% (>5%); drop_ns.p99 -29.76% (>15%); drop_ns.mean -6.07% (>5%) |
| p512-k1-file | open_ns.p99 -32.44% (>15%); open_ns.mean +7.49% (>5%); commit_ns.p50 +17.24% (>5%); commit_ns.p99 +177.88% (>15%); commit_ns.mean +15.86% (>5%) |
| p512-k1-owned | open_ns.p99 +28.31% (>15%); commit_ns.p50 +5.71% (>5%); commit_ns.mean +6.21% (>5%) |
| p512-k32-file | open_ns.p99 +43.75% (>15%); commit_ns.p99 -28.48% (>15%); drop_ns.p99 +21.05% (>15%) |
| p512-k32-owned | open_ns.p50 +25.03% (>5%); drop_ns.p99 +141.02% (>15%) |
| p512-k8-file | open_ns.mean -6.84% (>5%); commit_ns.p95 -10.90% (>10%); commit_ns.p99 -71.85% (>15%); commit_ns.mean -9.03% (>5%); drop_ns.p99 -76.35% (>15%) |
| p512-k8-owned | open_ns.p50 -48.73% (>5%); open_ns.p95 -20.64% (>10%); open_ns.p99 -17.79% (>15%); open_ns.mean -32.41% (>5%); commit_ns.p99 -84.53% (>15%); commit_ns.mean -16.96% (>5%); publish_ns.p99 +22.28% (>15%); drop_ns.p99 -21.51% (>15%) |

All case/campaign phase statistics, repeat groups, route ratios, and bootstrap intervals are in `native-analysis.json`.
