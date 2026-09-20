# Candidate qualification

All six final qualification commands returned zero. Durations are correctness-check receipts, not performance claims.

| Gate | Exit | Seconds |
|---|---:|---:|
| fmt | 0 | 18.14 |
| check | 0 | 21.48 |
| clippy | 0 | 11.16 |
| tests | 0 | 67.63 |
| facade | 0 | 38.54 |
| rustdoc | 0 | 2.78 |

CFB/XLS: 1900 passed, 2 existing ignored. Facade: 61 passed. No additional runtime tests were needed: existing warm/cold parity and indexed-missing trailing freshness tests cover this branch. Source/audit ordering assertions cover its execution fence without a timing-dependent test.

All 126 real XLS fixtures in owned/file modes and all 70,001 generated cells preserve exact corpus outcomes and source metrics. Final source is archived in candidate-source/. No prequalification candidate revision or failed build is omitted.
