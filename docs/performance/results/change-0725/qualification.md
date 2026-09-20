# Candidate qualification

All final qualification commands returned zero. Durations are correctness-check receipts, not performance claims.

| Gate | Exit | Seconds |
|---|---:|---:|
| fmt | 0 | 18.19 |
| check | 0 | 0.21 |
| clippy | 0 | 0.13 |
| tests | 0 | 8.33 |
| facade | 0 | 0.18 |
| rustdoc | 0 | 2.87 |

CFB/XLS tests total 1,904 passed and two existing ignored; facade checks add 61 passed. Corpus comparisons cover 126 real XLS fixtures in each of owned and file modes and all 70,001 generated cells, with exact baseline parity.

Final sources are archived in `candidate-source/`. The initial unused-mut build failure and the passing pre-consolidation qualification are retained separately. Removing a duplicate freshness test changed only the test inventory; the existing cache test covers that fence. No main timing captures preceded final qualification.
