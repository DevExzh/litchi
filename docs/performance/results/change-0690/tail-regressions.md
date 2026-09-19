# Additional mean and tail review triggers

Both paired changes exceed 5%. These are descriptive distributions from 100
samples per leg; individual tail samples and control drift can be noisy.
They are retained even when the median improves. Full distributions and A/A
controls remain in `native-comparison.json`.

| Case | Source | Phase | Statistic | B1/A1 % | B2/A2 % | A/A % |
|---|---|---|---|---:|---:|---:|
| 54016-missing-1048576 | owned | q3 | p95 | +6.45 | +9.68 | +0.00 |
| 54016-missing-1048576 | owned | q3 | p99 | +8.82 | +6.25 | +6.45 |
| Simple-stored-2097152 | file | q3-to-q8-mean | p99 | +70.54 | +63.55 | +0.85 |
| Simple-missing-2097152 | owned | q3 | mean | +10.52 | +8.66 | -0.78 |
| Simple-missing-2097152 | owned | q8 | p95 | +12.50 | +12.50 | +0.00 |
| Simple-missing-2097152 | file | q1 | mean | +8.94 | +7.66 | -6.38 |
| Plan1-late | file | q1 | p99 | +12.60 | +9.71 | +2.93 |
| 45365-first | file | open | p99 | +15.62 | +14.80 | +13.81 |
| 45365-first | file | q3-to-q8-mean | p99 | +49.67 | +51.92 | +65.99 |
