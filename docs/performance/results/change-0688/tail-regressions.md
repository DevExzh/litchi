# Additional mean and tail review triggers

Both paired changes exceed 5%. These are descriptive distributions from 100
samples per leg; individual tail samples and control drift can be noisy.
They are retained even when the median improves. Full distributions and A/A
controls remain in `native-comparison.json`.

| Case | Source | Phase | Statistic | B1/A1 % | B2/A2 % | A/A % |
|---|---|---|---|---:|---:|---:|
| 54016-missing-1048576 | owned | q8 | p95 | +15.38 | +7.69 | -7.14 |
| 54016-missing-1048576 | owned | q8 | p99 | +23.08 | +7.14 | -6.67 |
| Plan1-stored-2097152 | owned | open | p99 | +6.47 | +10.03 | +21.55 |
| Simple-stored-2097152 | owned | q3-to-q8-mean | mean | +5.34 | +5.36 | +0.29 |
| Simple-stored-2097152 | owned | open-plus-eight | p99 | +18.94 | +35.92 | +19.49 |
| Simple-missing-2097152 | owned | q8 | mean | +11.83 | +6.06 | +0.13 |
| Simple-missing-2097152 | owned | q8 | p99 | +11.11 | +11.11 | -10.00 |
| Simple-missing-2097152 | file | open | p99 | +24.18 | +24.75 | +19.61 |
| Plan1-late | owned | q1 | p99 | +9.51 | +9.52 | -9.01 |
| Plan1-late | file | q1 | p99 | +6.14 | +6.55 | -5.29 |
| 45365-first | file | open | p99 | +7.26 | +26.13 | +29.74 |
