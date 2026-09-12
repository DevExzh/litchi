# 0517 native DOCX analysis

Native campaigns: r1, r2. Rows are source-bound through each receipt; profile and hardware lanes are excluded.
Measured rows are the 30 `warmup=false` samples in each of two internal repeats. Phase percentiles use nearest rank; means are arithmetic means; RSS is the one whole-child maximum from GNU time.
The unpaired route bootstrap for the p50 ratio uses seed `5339665` and 4000 iterations, resampling each route within each internal repeat and taking batch median / repeated median. With only two internal repeats, its interval is descriptive for these row distributions and does not estimate independent-process or host variation.

## API-choice ratios

Ratios are batch / repeated; positive percentages mean batch took longer or used more RSS. The same absolute thresholds are used for route flags and cross-campaign flags: 5% for p50/mean/RSS, 10% for p95, and 15% for p99.

| Campaign | Workload | elapsed p50 | edit p50 | publish p50 | RSS | route flags |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| r1 | p128-k1-file | 0.29% | 0.35% | 0.25% | -1.33% | 1 |
| r1 | p128-k1-owned | 0.17% | 0.01% | -0.09% | -0.05% | 4 |
| r1 | p128-k32-file | -85.65% | -92.01% | -1.44% | -0.48% | 12 |
| r1 | p128-k32-owned | -85.85% | -92.05% | 10.07% | -2.52% | 16 |
| r1 | p128-k8-file | -57.68% | -73.83% | -10.57% | 4.56% | 19 |
| r1 | p128-k8-owned | -59.23% | -74.20% | -12.26% | 1.28% | 18 |
| r1 | p512-k1-file | 0.61% | 0.15% | 0.87% | -2.29% | 4 |
| r1 | p512-k1-owned | 0.25% | 0.73% | -0.30% | 1.26% | 2 |
| r1 | p512-k32-file | -86.08% | -92.67% | 1.06% | -0.86% | 11 |
| r1 | p512-k32-owned | -86.06% | -92.61% | 0.81% | 2.87% | 11 |
| r1 | p512-k8-file | -57.63% | -74.12% | -0.46% | 3.51% | 12 |
| r1 | p512-k8-owned | -57.92% | -74.58% | 0.46% | 2.52% | 12 |
| r2 | p128-k1-file | 0.19% | 0.89% | -0.68% | 0.74% | 1 |
| r2 | p128-k1-owned | -0.44% | -2.32% | -0.95% | 1.17% | 3 |
| r2 | p128-k32-file | -85.62% | -92.10% | -0.20% | 1.55% | 14 |
| r2 | p128-k32-owned | -85.85% | -92.02% | -1.23% | 0.78% | 11 |
| r2 | p128-k8-file | -58.15% | -74.26% | -11.57% | 3.18% | 18 |
| r2 | p128-k8-owned | -58.68% | -74.20% | -3.80% | 1.11% | 19 |
| r2 | p512-k1-file | -2.63% | -4.46% | -0.57% | 0.00% | 4 |
| r2 | p512-k1-owned | -0.46% | -0.99% | -0.09% | 0.69% | 3 |
| r2 | p512-k32-file | -85.74% | -92.51% | 3.90% | 1.15% | 10 |
| r2 | p512-k32-owned | -86.13% | -92.66% | -0.47% | -1.71% | 14 |
| r2 | p512-k8-file | -57.97% | -74.50% | -0.31% | 3.70% | 17 |
| r2 | p512-k8-owned | -58.16% | -74.60% | -0.11% | 0.50% | 11 |

## Cross-campaign flags

### Matched same API

23 of 24 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-file-batch | open_ns.p99 +17.20% (>15%) |
| p128-k1-file-repeated | commit_ns.p99 -18.60% (>15%) |
| p128-k1-owned-batch | open_ns.p99 -25.01% (>15%); commit_ns.p95 -20.41% (>10%); commit_ns.p99 -34.38% (>15%); commit_ns.mean -7.42% (>5%) |
| p128-k1-owned-repeated | open_ns.p99 -23.24% (>15%) |
| p128-k32-file-batch | open_ns.p95 +40.55% (>10%); open_ns.p99 +47.60% (>15%); drop_ns.p99 +139.64% (>15%) |
| p128-k32-file-repeated | open_ns.p99 -16.47% (>15%); commit_ns.p99 +77.61% (>15%) |
| p128-k32-owned-batch | open_ns.mean +13.54% (>5%) |
| p128-k32-owned-repeated | open_ns.p95 +46.20% (>10%); open_ns.p99 +39.86% (>15%); open_ns.mean +6.86% (>5%); publish_ns.p50 +11.83% (>5%); publish_ns.mean +5.80% (>5%) |
| p128-k8-file-batch | open_ns.p99 +30.40% (>15%) |
| p128-k8-file-repeated | open_ns.p99 +26.40% (>15%); drop_ns.p99 -81.85% (>15%); drop_ns.mean -7.56% (>5%) |
| p128-k8-owned-repeated | open_ns.p99 +30.29% (>15%); publish_ns.p50 -8.66% (>5%); publish_ns.p99 +15.72% (>15%) |
| p512-k1-file-batch | commit_ns.p95 +12.20% (>10%); commit_ns.p99 +68.89% (>15%); drop_ns.p99 +38.14% (>15%) |
| p512-k1-file-repeated | open_ns.p99 -20.76% (>15%); commit_ns.p99 -29.23% (>15%) |
| p512-k1-owned-batch | commit_ns.p50 +6.25% (>5%); commit_ns.p99 -35.94% (>15%); commit_ns.mean +5.31% (>5%) |
| p512-k1-owned-repeated | open_ns.p99 +18.29% (>15%) |
| p512-k32-file-batch | open_ns.p99 -34.18% (>15%) |
| p512-k32-file-repeated | open_ns.p99 +28.43% (>15%); drop_ns.p99 -68.90% (>15%) |
| p512-k32-owned-batch | open_ns.p50 -5.77% (>5%); open_ns.p99 -23.81% (>15%); commit_ns.p99 +402.56% (>15%); commit_ns.mean +17.14% (>5%) |
| p512-k32-owned-repeated | open_ns.p99 +19.35% (>15%) |
| p512-k8-file-batch | open_ns.p95 -12.13% (>10%) |
| p512-k8-file-repeated | open_ns.p50 +87.23% (>5%); open_ns.p95 +89.29% (>10%); open_ns.p99 +140.16% (>15%); open_ns.mean +90.80% (>5%); commit_ns.p50 +21.95% (>5%); commit_ns.p95 +17.02% (>10%); commit_ns.p99 +40.82% (>15%); commit_ns.mean +19.31% (>5%); drop_ns.p99 -24.26% (>15%) |
| p512-k8-owned-batch | open_ns.p99 +19.48% (>15%) |
| p512-k8-owned-repeated | open_ns.p95 -15.64% (>10%); commit_ns.p99 -21.62% (>15%); commit_ns.mean +5.00% (>5%) |

### API-choice ratio

12 of 12 records have one or more absolute threshold flags.

| Workload | flags |
| --- | --- |
| p128-k1-file | open_ns.p99 +23.02% (>15%); commit_ns.p95 +12.30% (>10%); commit_ns.p99 +22.86% (>15%); commit_ns.mean +7.31% (>5%) |
| p128-k1-owned | commit_ns.p95 -22.82% (>10%); commit_ns.p99 -28.91% (>15%); commit_ns.mean -7.93% (>5%) |
| p128-k32-file | open_ns.p95 +39.64% (>10%); open_ns.p99 +76.69% (>15%); commit_ns.p99 -43.29% (>15%); drop_ns.p99 +133.58% (>15%) |
| p128-k32-owned | open_ns.p95 -31.10% (>10%); open_ns.p99 -29.32% (>15%); open_ns.mean +6.25% (>5%); publish_ns.p50 -10.27% (>5%) |
| p128-k8-file | drop_ns.p99 +419.56% (>15%); drop_ns.mean +7.84% (>5%) |
| p128-k8-owned | open_ns.p99 -28.89% (>15%); publish_ns.p50 +9.65% (>5%) |
| p512-k1-file | open_ns.p99 +26.61% (>15%); commit_ns.p50 +5.48% (>5%); commit_ns.p99 +138.65% (>15%); commit_ns.mean +6.97% (>5%); drop_ns.p99 +33.64% (>15%) |
| p512-k1-owned | open_ns.p99 -23.89% (>15%); commit_ns.p50 +9.68% (>5%); commit_ns.p99 -35.94% (>15%); commit_ns.mean +5.47% (>5%) |
| p512-k32-file | open_ns.p99 -48.75% (>15%); drop_ns.p99 +205.75% (>15%) |
| p512-k32-owned | open_ns.p50 -5.91% (>5%); open_ns.p99 -36.16% (>15%); commit_ns.p95 +14.55% (>10%); commit_ns.p99 +429.91% (>15%); commit_ns.mean +20.67% (>5%) |
| p512-k8-file | open_ns.p50 -44.86% (>5%); open_ns.p95 -53.58% (>10%); open_ns.p99 -58.74% (>15%); open_ns.mean -47.28% (>5%); commit_ns.p50 -19.95% (>5%); commit_ns.p95 -14.55% (>10%); commit_ns.p99 -28.99% (>15%); commit_ns.mean -18.61% (>5%); drop_ns.p99 +32.02% (>15%) |
| p512-k8-owned | open_ns.p95 +20.62% (>10%); open_ns.p99 +21.66% (>15%); commit_ns.p99 +22.77% (>15%) |

All case/campaign phase statistics, repeat groups, route ratios, and bootstrap intervals are in `native-analysis.json`.
