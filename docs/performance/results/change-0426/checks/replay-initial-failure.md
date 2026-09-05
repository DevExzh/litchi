# Initial portable replay failure

The first root invocation of `python3 -B
docs/performance/results/change-0426/verify.py` exited 1 with
`FileNotFoundError` for `checks/layout-before.log.gz`. This is a transcription
of the observed tool failure, not a separately captured process-output log.

The verifier incorrectly expected a baseline stdout sidecar. The initial
baseline probe had retained its observed type sizes in `layout-before.json`
and its empty successful compiler diagnostics in `layout-before-build.log`.
The corrected verifier checks those existing artifacts. No baseline observation
or missing raw output was recreated. The candidate probe separately retains
its JSON stdout in `layout-after.log`.

The first staged diff check also found one added blank line at EOF in
`source-review-facade.md`. The initial commit was made despite that check
failure; the final amendment removes the blank line and refreshes the inventory.
The complete batch diff is rechecked before the amendment.
