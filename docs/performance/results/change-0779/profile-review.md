# 0779 Heaptrack profile review

The offline replay in `heaptrack_analysis.py` passed for both retained roles.
It decompressed each interpreted Heaptrack trace, summed every live `+`
allocation event, and matched the result to the corresponding retained
histogram.  The capture and decode receipts, report identity, source identity,
binary receipt, and exact command flags are checked before either trace is
accepted.  No native binary, Cargo command, Heaptrack capture, or
`heaptrack_print` invocation is performed by the replay.

The module exposes a pure `analyze(root)` API for the final validator.  After
relocation it maps only the original `origin.json` owned-worktree prefix for
file checks and keeps the captured absolute strings in the returned result.
If a captured binary has been removed, replay accepts it only with an exact
size/SHA/path witness in a verified `cleanup.json`.

The probe ran one open sample.  Its `run_one` performs the timed
`Workbook::open` and then a second `Workbook::open` for post-clock
verification.  The table therefore reports two open calls inside one sample;
it does not treat them as two samples or as two timing measurements.

| role | open call | allocation events | growth `realloc` events | requested bytes |
| --- | --- | ---: | ---: | ---: |
| before | timed open | 516 | 515 | 1,092,697,469 |
| before | post-clock verification open | 516 | 515 | 1,092,697,469 |
| before | whole process | 1,032 | 1,030 | 2,185,394,938 |
| after | timed open | 44 | 43 | 40,650,096 |
| after | post-clock verification open | 44 | 43 | 40,650,096 |
| after | whole process | 88 | 86 | 81,300,192 |

The pre-change arithmetic sequence is independently reproduced for each
source open: one 8,192-byte initial allocation followed by 515 exact-growth
reallocations.  The candidate changes that sequence to 43 growth reallocations
per open.  Whole-process requested bytes are 3.72% of the baseline
(a 96.28% reduction), and growth reallocations are 8.35% of the baseline (a
91.65% reduction).

The all-process histogram reconciliation is also exact:

| role | histogram rows | allocation events | requested bytes | target share |
| --- | ---: | ---: | ---: | ---: |
| before | 639 | 3,300 | 2,191,954,137 | 99.700760% |
| after | 167 | 2,356 | 87,859,389 | 92.534438% |

These are allocator request totals attributed through the interpreted
`litchi_opc::phys_pkg::read_limited` ancestry.  They do not measure bytes
physically copied by `realloc`, resident memory, RSS, or latency, and the
profile does not establish any of those outcomes.

The source identity is unchanged in both reports:

```text
source bytes:  4,226,429
source SHA256: dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036
```

The captured binary receipts are retained in `heaptrack-analysis.json`: the
before binary is `022031bc489700894de61b72d63347dfd349a0af88f0c10eb783e473ba7e48cd`
and the after binary is
`140773be5783e8cb691eac6cab813451d366bb3acfca8dc1584d37264a58946f`.
