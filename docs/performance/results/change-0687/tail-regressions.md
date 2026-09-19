# Additional mean and tail review triggers

Both paired changes exceed 5%. These are descriptive distributions from 100
samples per leg; individual tail samples and control drift can be noisy.
They are retained even when the median improves. Full distributions and A/A
controls remain in `native-comparison.json`.

| Case | Source | Phase | Statistic | B1/A1 % | B2/A2 % | A/A % |
|---|---|---|---|---:|---:|---:|
| Simple-stored-2097152 | file | open | p99 | +110.16 | +16.46 | -33.35 |
| 45365-first | owned | q2 | mean | +9.38 | +6.81 | +1.95 |
| 45365-first | owned | q2 | p95 | +8.32 | +6.15 | +2.89 |
| 45365-first | owned | q2 | p99 | +11.28 | +6.29 | +3.13 |
| 45365-first | owned | q3-to-q8-mean | p99 | +12.12 | +5.41 | +1.22 |
| 45365-first | file | q2 | mean | +6.73 | +8.23 | -2.53 |
| 45365-first | file | q2 | p95 | +6.41 | +5.25 | -1.44 |
| 45365-first | file | q3-to-q8-mean | p99 | +16.03 | +38.02 | -32.73 |
| 45365-late | owned | q2 | mean | +5.65 | +8.65 | +0.66 |
| 45365-late | owned | q2 | p95 | +5.92 | +8.85 | +0.08 |
