# ODS order and rank statistics evidence

This batch targets `MEDIAN`, `MODE`, `LARGE`, `SMALL`, `PERCENTILE`,
`PERCENTRANK`, `QUARTILE`, and `RANK` in the explicit read-only formula
evaluator. Formula caches and document CRUD remain inert.

The batch is accepted against the frozen source. It includes bounded numeric
storage and work-charged sorting, streaming scalar RANK/PERCENTRANK queries,
parameter-driven matrix shapes, and explicit LARGE/SMALL array-rank results
that survive scalar publication and transparent wrappers.

The baseline and unrelated tracked edits are recorded in `baseline.json`.
Isolated gates use the retained `gates/Cargo.lock`, inherited from the preceding
dispersion batch, rather than the ambient workspace lockfile.

The [contract](contract.md) records sequence admission, parameter matrix
iteration, ties, interpolation, and the chosen error profile. The independent
[oracle generator](numeric_oracle.py) produces 512 observations across 56
fixtures; regeneration passes. The [native evidence](native/README.md)
contains 45 observations across all eight functions from pinned LibreOffice
FODS files. Its [reproduction receipt](native-reproduction.json) confirms
byte-for-byte regeneration and temporary-tree cleanup. The evaluator passes
all eight oracle tests and both native tests, together with 17 authored
semantic tests and 16 resource tests.

Before final gates, `gates/stage.py` copies and verifies every selected profile
source, including the oracle generator, into the isolated checkout. Source
freeze and gate receipts are retained under `gates/`. The isolated ODS suite
passes 1,421 tests with zero failures or ignored tests. All seven isolated
gates pass with stable source hashes: tests, strict all-target Clippy,
warning-denied rustdoc, crate formatting, explicit batch formatting, crate
boundaries, and diff checks. Independent [semantic](review.md) and
[resource/cache](resource-review.md) reviews are PASS.

The [matched performance report](performance/results/performance-report.md)
retains 630 baseline and 3,660 candidate samples: 21 matched controls and 122
candidate cases in two phases, with three warmups and fifteen fresh process
samples per group. Matched allocation-call counts, work, and resolver reads
are unchanged; requested and peak-live bytes increase by at most 0.31%.
Median elapsed-time changes range from −2.53% to +6.39%, and process RSS from
−6.38% to +1.87%. Candidate-only cases establish absolute costs, not speedups.

The SUMIFS parse-and-evaluate median crosses the 5% review threshold at
+6.39%. Its evaluate-only contrast is +0.53%, and the retained sample
distributions overlap substantially. The [threshold review](performance/results/threshold-review.json)
accepts the batch with that flag retained; it does not establish a causal
regression or claim the flag is cleared. The descriptive bootstrap interval
spans −5.77% to +8.19%. No rerun replaced the flagged samples.

The independent [verifier](verify.py) checks the full raw capture, source and
input hashes, gate logs, native receipts, oracle regeneration, read bounds,
allocation balance, and normalized report values. Its retained
[result](verification.json) and [receipt](verification-receipt.json) pass.
Run `python3 -B docs/report/spec-gap-validation-evidence/ods-formula-order-statistics/verify.py`
from the repository root to reproduce this verification.

Owned temporary worktrees and targets were removed; the
[root cleanup receipt](scratch-cleanup.json) records removal of the isolated
gate checkout and target after verifying all 45 staged source slots, following
the [profiler cleanup](performance/results/scratch-cleanup.json). Shared build
artifacts and unrelated workspace changes were preserved.
