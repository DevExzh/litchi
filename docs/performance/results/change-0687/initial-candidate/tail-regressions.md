# Additional mean and tail review triggers

Both paired changes exceed 5%. These are descriptive distributions from 100
samples per leg; individual tail samples and control drift can be noisy.
They are retained even when the median improves. Full distributions and A/A
controls remain in `native-comparison.json`.

| Case | Source | Phase | Statistic | B1/A1 % | B2/A2 % | A/A % |
|---|---|---|---|---:|---:|---:|
| 54016-stored-0 | owned | open | mean | +11.81 | +13.41 | -0.49 |
| 54016-stored-0 | owned | open | p95 | +10.58 | +13.28 | -0.33 |
| 54016-stored-0 | owned | open | p99 | +9.47 | +14.27 | -0.23 |
| 54016-stored-0 | file | open | mean | +11.81 | +13.00 | +0.26 |
| 54016-stored-0 | file | open | p95 | +11.36 | +12.04 | +0.29 |
| 54016-stored-0 | file | open | p99 | +9.54 | +11.66 | -0.58 |
| 54016-stored-2097152 | owned | open | mean | +12.46 | +11.91 | +0.16 |
| 54016-stored-2097152 | owned | open | p95 | +12.02 | +11.37 | +0.23 |
| 54016-stored-2097152 | owned | open | p99 | +13.23 | +7.52 | -0.74 |
| 54016-stored-2097152 | file | open | mean | +11.55 | +11.75 | +0.07 |
| 54016-stored-2097152 | file | open | p95 | +10.91 | +11.03 | +0.14 |
| 54016-stored-2097152 | file | open | p99 | +12.19 | +11.39 | +0.27 |
| 54016-missing-1048576 | owned | open | mean | +15.29 | +14.65 | +0.36 |
| 54016-missing-1048576 | owned | open | p95 | +12.69 | +12.26 | +0.08 |
| 54016-missing-1048576 | owned | open | p99 | +10.79 | +13.22 | +0.27 |
| 54016-missing-1048576 | file | open | mean | +14.68 | +14.14 | +0.11 |
| 54016-missing-1048576 | file | open | p95 | +12.47 | +11.91 | -0.37 |
| 54016-missing-1048576 | file | open | p99 | +13.49 | +11.41 | -1.07 |
| 54016-missing-1048576 | file | q3 | p99 | +8.89 | +9.09 | +0.00 |
| Simple-missing-2097152 | file | open-plus-eight | p99 | +33.29 | +34.74 | -24.03 |
| 54016-late | owned | open | mean | +12.80 | +12.65 | -0.05 |
| 54016-late | owned | open | p95 | +12.58 | +12.12 | -0.18 |
| 54016-late | owned | open | p99 | +13.28 | +8.88 | -2.63 |
| 54016-late | file | open | mean | +11.93 | +12.20 | +0.37 |
| 54016-late | file | open | p95 | +11.48 | +11.40 | +0.58 |
| 54016-late | file | open | p99 | +10.89 | +12.55 | +1.86 |
| Plan1-late | owned | open | p99 | +22.35 | +9.69 | -18.45 |
| 45365-first | owned | q2 | mean | +5.88 | +9.02 | +1.95 |
| 45365-first | owned | q2 | p95 | +6.46 | +8.21 | +2.89 |
| 45365-first | file | q2 | mean | +8.30 | +6.49 | -2.53 |
| 45365-first | file | q2 | p95 | +9.13 | +6.45 | -1.44 |
