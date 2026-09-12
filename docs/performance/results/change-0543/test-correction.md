# 0543 pre-capture test correction

The initial baseline attempt passed warning-denied guard Clippy and then
failed the new differential test after its expected raw error and retry checks
had succeeded. Its unaffected-sheet check called `editor.snapshot`, whose
single-sheet workbook restriction correctly refused the two-sheet fixture.
This was test API misuse, not a production regression or a measured candidate.

`failed-baseline-1/` retains the complete source manifest, source snapshots,
partial full-suite stdout/stderr and both receipts. `differential-tests.patch`
and `frozen-inputs.json` retain the original pre-build inputs unchanged.

The corrected test uses the supported `edit_sheets` transaction to check the
unaffected sheet's value, empty patch and unchanged commit. It checks the
original worksheet XML through the no-op publication's archive readback, as
well as byte-exact whole-source publication and the original source version.
The correction keeps the typed error, repeat, source and atomicity assertions.

`differential-tests-final.patch` describes the corrected shared baseline and
candidate test sources. `baseline-frozen-inputs.json` binds those exact files
before the fresh baseline freeze. No release build or benchmark capture was
performed on the failed attempt. The corrected baseline full XLSX suite passes;
subsequent candidate and final quality results are recorded separately.
