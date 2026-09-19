# Additional mean and tail review triggers

Both paired changes exceed 5%. These are descriptive distributions from 100
samples per leg; individual tail samples and control drift can be noisy.
They are retained even when the median improves. Full distributions and A/A
controls remain in `native-comparison.json`.

| Case | Source | Phase | Statistic | B1/A1 % | B2/A2 % | A/A % |
|---|---|---|---|---:|---:|---:|
| 54016-missing-1048576 | owned | q8 | p95 | +7.69 | +7.69 | +7.14 |
| 54016-missing-1048576 | owned | q8 | p99 | +7.69 | +7.69 | +6.67 |
| Simple-stored-2097152 | owned | open-plus-eight | p99 | +13.03 | +8.88 | -27.58 |
| synthetic-70000-default | owned | q3-to-q8-mean | p99 | +150.39 | +25.57 | +0.26 |
| 45365-first | file | open | p95 | +17.90 | +10.14 | -2.01 |
| 45365-first | file | q3 | p99 | +52.05 | +53.91 | +4.10 |
| 45365-first | file | q8 | p99 | +54.42 | +52.84 | +0.89 |
| 45365-late | file | q3-to-q8-mean | p99 | +29.66 | +98.95 | +1.93 |
