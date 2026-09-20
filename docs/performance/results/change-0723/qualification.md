# Candidate qualification

All final qualification commands returned zero. These are correctness and build checks; their durations are not performance claims.

| Gate | Exit | Seconds |
|---|---:|---:|
| fmt | 0 | 18.16 |
| check | 0 | 21.52 |
| clippy | 0 | 11.21 |
| tests | 0 | 68.7 |
| facade | 0 | 38.19 |
| rustdoc | 0 | 2.78 |

CFB/XLS unit, integration and doctest summaries total 1,904 passed and two existing ignored; facade checks add 61 passed. The standalone probes compare 126 real XLS fixtures in each of owned and file modes, plus all 70,001 generated cells, against the frozen baseline. All comparisons pass.

The final candidate source and test files are archived in `candidate-source/`; initial release builds and the pre-review integration test are separately retained. No paired timing was run on that initial test revision.
