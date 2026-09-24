# Scalar baseline for array and reference evaluation

This is a baseline capture, not evidence that array/reference evaluation is
implemented or that candidate performance meets the regression gate.

The baseline is commit `67360b209c5b428161b27f1f08650169d4997739`, built
offline with the locked dependency graph in a separate sparse worktree. The
four-file `scalar-harness` runs the 53 established scalar controls plus all
41 Roman/Arabic cases on both revisions. Each case measures parsing,
evaluation, and combined parsing/evaluation separately. All 282 captures
completed successfully, including expected resource and capability refusals.

The capture used CPU 6. A repository boundary check ran concurrently on the
host; this capture is not an idle-host measurement. Individual regressions
must receive repeated paired measurements before attributing them to code.
Compiler, host, commands, source hashes, harness hashes, timings, allocation
counters and peak RSS are retained in `baseline/`. The exact executable is
retained outside the repository for subsequent paired comparisons; its hash
and relocation are recorded in `baseline/binary-relocation.json`.

`baseline-root-verification.json` independently checks the exact 282-row
corpus, repeat counts, successful capture statuses and comparison identity.
`compare.py` rejects missing/duplicate cases and semantic/result-reservation
mismatches, reports every individual metric, and flags any median/tail latency
or RSS increase above 5%. A flag is a review trigger, not an automatic failure
or a claim of statistical significance.

The candidate-only array/reference harness, complete candidate gates,
before/after comparison, scaling measurements, and final implementation review
remain outstanding. No workbook recalculation or additional function-family
support is claimed by this baseline.
