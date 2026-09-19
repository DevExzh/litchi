# Final matched-control regression review

The final capture compares 13 controls in `evaluate` and `parse-evaluate`,
with 15 fresh child processes per phase. Times are p50 nanoseconds per fixed
repeat count, as recorded in `performance-report.md` and
`performance-report.json`.

No matched control has a positive time delta above the 5% review threshold.
The largest positive delta is `scalar-control-arithmetic` in
`parse-evaluate`, 690 ns to 703 ns (+1.88%). The largest negative delta is
`scalar-control-imsum` in `parse-evaluate`, 2,046 ns to 1,935 ns (-5.43%);
that is an improvement and is not a regression.

The threshold is applied to the p50 control comparison rather than to one
child or to candidate-only statistical rows. Fresh child processes, fixed
repeat counts, balanced allocator receipts, stable source and profile-input
hashes, and the independent correctness preflight provide the noise controls;
the result remains a single-host profile with no cross-platform claim.

The generated workload prose was corrected after capture to state the full
nested matrix: all nine reducers have 64-row rows, while `AVERAGE`, `COUNTA`,
and `COUNTBLANK` additionally have 256-row and 1024-row rows. The frozen
harness and `summarize.py` inputs were not changed.
