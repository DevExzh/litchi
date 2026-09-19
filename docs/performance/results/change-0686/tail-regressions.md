# Additional mean and tail review triggers

Both paired changes exceed 5%. These are descriptive distributions from 100
samples per leg; individual tail samples and control drift can be noisy.
They are retained even when the median improves. Full distributions and A/A
controls remain in `native-comparison.json`.

| Case | Source | Phase | Statistic | B1/A1 % | B2/A2 % | A/A % |
|---|---|---|---|---:|---:|---:|
| 54016-stored-524288 | owned | q2 | mean | +9.42 | +9.38 | -0.32 |
| 54016-stored-524288 | owned | q2 | p95 | +7.64 | +7.78 | -0.60 |
| 54016-stored-524288 | owned | q2 | p99 | +8.74 | +5.97 | +0.09 |
| 54016-stored-524288 | file | q2 | mean | +7.75 | +7.25 | +0.10 |
| 54016-stored-524288 | file | q2 | p95 | +7.54 | +6.63 | -0.23 |
| 54016-stored-524288 | file | q2 | p99 | +7.92 | +6.61 | +0.26 |
| 54016-stored-1048576 | owned | q2 | p95 | +8.62 | +9.08 | -0.06 |
| 54016-stored-1048576 | owned | q2 | p99 | +9.01 | +8.79 | -0.72 |
| 54016-stored-1048576 | file | q2 | p95 | +7.21 | +8.92 | -0.25 |
| 54016-stored-1048576 | file | q2 | p99 | +7.56 | +9.39 | -0.58 |
| 54016-missing-1048576 | owned | q2 | p95 | +8.69 | +9.56 | +0.10 |
| 54016-missing-1048576 | owned | q2 | p99 | +9.28 | +8.63 | +0.32 |
| 54016-missing-1048576 | file | q2 | p95 | +9.39 | +9.59 | -0.07 |
| 54016-missing-1048576 | file | q2 | p99 | +9.54 | +10.08 | +0.61 |
| Plan1-stored-32768 | owned | q2 | mean | +5.86 | +5.77 | -2.51 |
| Plan1-stored-32768 | owned | q2 | p99 | +5.16 | +11.01 | -2.23 |
| Plan1-stored-2097152 | file | open | p95 | +5.81 | +8.86 | +0.70 |
| Simple-stored-450 | file | q3 | p99 | +6.39 | +5.65 | -4.24 |
| Simple-missing-2097152 | owned | q3 | mean | +10.65 | +11.76 | +3.60 |
| Simple-missing-2097152 | owned | q3 | p95 | +20.00 | +20.00 | +10.00 |
| Simple-missing-2097152 | owned | q3 | p99 | +9.09 | +9.09 | +0.00 |
| formula-refusal-2097152 | owned | q1 | mean | +13.65 | +6.45 | -7.38 |
| formula-refusal-2097152 | owned | q1 | p95 | +8.70 | +7.26 | -1.45 |
| formula-refusal-2097152 | owned | q1 | p99 | +235.26 | +8.53 | -65.22 |
| formula-refusal-2097152 | owned | q2 | mean | +7.29 | +5.12 | +0.45 |
| formula-refusal-2097152 | owned | q2 | p95 | +7.32 | +5.89 | +0.84 |
| formula-refusal-2097152 | owned | q2 | p99 | +7.71 | +6.03 | +0.21 |
| formula-refusal-2097152 | owned | q3 | p95 | +8.61 | +5.68 | +0.22 |
| formula-refusal-2097152 | owned | q3 | p99 | +10.96 | +6.72 | +0.24 |
| formula-refusal-2097152 | owned | q8 | mean | +11.16 | +11.15 | +2.61 |
| formula-refusal-2097152 | owned | q8 | p95 | +9.55 | +9.55 | +0.00 |
| formula-refusal-2097152 | owned | q8 | p99 | +9.93 | +206.68 | +0.22 |
| formula-refusal-2097152 | owned | q3-to-q8-mean | mean | +9.45 | +8.86 | -0.05 |
| formula-refusal-2097152 | owned | q3-to-q8-mean | p95 | +8.59 | +7.81 | +0.19 |
| formula-refusal-2097152 | owned | q3-to-q8-mean | p99 | +34.36 | +6.19 | -0.03 |
| formula-refusal-2097152 | owned | open-plus-eight | p99 | +5.59 | +12.93 | -11.68 |
| formula-refusal-2097152 | file | q1 | mean | +6.78 | +11.35 | -1.32 |
| formula-refusal-2097152 | file | q1 | p95 | +6.49 | +7.41 | -0.35 |
| formula-refusal-2097152 | file | q2 | mean | +7.23 | +5.94 | -1.37 |
| formula-refusal-2097152 | file | q2 | p95 | +5.35 | +5.52 | -0.78 |
| formula-refusal-2097152 | file | q3 | mean | +7.09 | +5.25 | -0.34 |
| formula-refusal-2097152 | file | q3 | p95 | +6.58 | +6.75 | -0.16 |
| formula-refusal-2097152 | file | q3 | p99 | +7.03 | +6.54 | -25.33 |
| formula-refusal-2097152 | file | q8 | mean | +10.61 | +7.41 | -2.04 |
| formula-refusal-2097152 | file | q8 | p95 | +5.85 | +7.24 | -0.66 |
| formula-refusal-2097152 | file | q8 | p99 | +6.48 | +6.66 | -0.65 |
| formula-refusal-2097152 | file | q3-to-q8-mean | mean | +8.05 | +7.41 | +0.09 |
| formula-refusal-2097152 | file | q3-to-q8-mean | p95 | +7.70 | +7.46 | +0.28 |
| synthetic-70000-default | owned | open | p99 | +6.79 | +40.78 | +7.78 |
| synthetic-70000-default | owned | q2 | mean | +6.72 | +6.86 | -0.05 |
| synthetic-70000-default | owned | q2 | p95 | +6.45 | +6.64 | +0.09 |
| synthetic-70000-default | owned | q2 | p99 | +7.64 | +6.88 | +0.18 |
| synthetic-70000-default | file | q2 | mean | +7.20 | +7.19 | -0.02 |
| synthetic-70000-default | file | q2 | p95 | +6.80 | +7.07 | +0.11 |
| synthetic-70000-default | file | q2 | p99 | +7.94 | +6.41 | +0.12 |
| synthetic-100000-default | owned | q2 | mean | +7.06 | +7.06 | +0.17 |
| synthetic-100000-default | owned | q2 | p95 | +6.60 | +6.62 | +0.24 |
| synthetic-100000-default | owned | q2 | p99 | +6.56 | +6.37 | +0.01 |
| synthetic-100000-default | file | q2 | mean | +7.40 | +7.08 | +0.11 |
| synthetic-100000-default | file | q2 | p95 | +7.18 | +5.95 | +0.19 |
| synthetic-100000-default | file | q2 | p99 | +7.45 | +6.23 | -0.05 |
