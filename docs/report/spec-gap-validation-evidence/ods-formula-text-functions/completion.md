# Ordinary ODF text-function batch completion

Disposition: accepted for the 26-function §6.20 repository profile, with
explicit process-RSS costs. The seven §6.7 byte-position functions remain
open. This is read-only evaluation, not workbook recalculation or formula
cache publication.

The corrected source passes all seven isolated gates and 1,592 tests, with
zero failed or ignored tests. Evidence includes 244 exact independent-oracle
observations, 91 pinned native observations, pinned Unicode 17 regeneration,
and 5,474 exact fraction-helper comparisons. `verification.json` is the
separate root verifier receipt; run `python3 verify.py` to reproduce it.

Independent semantic review found grammar defects in the first freeze.
Those sources and receipts remain in `diagnostics/pre-grammar-fix/`.
The fixes and final five oracle cases were reviewed/completed by root after
review agents reached their usage limit. `grammar-closure.md` distinguishes
that follow-up from the original independent reviews. No second independent
agent approval of those final changes is claimed.

The accepted performance capture contains 840 baseline and 4,830 candidate
samples, with three warmups and fifteen fresh children per case/phase.
All 161 candidate preflights pass. Matched allocation calls, requested and
released bytes, peak live heap, retained budget, work, reads and output bytes
are unchanged. Matched latency shifts range from −3.88% to +3.69%.

There are 42 process-RSS review flags. Across all matched controls the
increases are 40–348 KiB (+1.14% to +10.00%). They are accepted as disclosed
costs of this bounded feature addition, not dismissed as noise or described
as an improvement. Resident code/data growth is a plausible contributor,
not a causal finding. See `performance/results/regression-review.md` for
all flags and bootstrap intervals, and `performance-report.md` for every
case. Unsupported baseline operands, scalar-error consumption and cancellation
repeat-count mistakes caused earlier rejected captures; all are retained
separately and are not used for the accepted comparison.

`README.md` and profile inputs remain at their captured bytes for source
custody, including their capture-time status wording. This completion record
and `validation-status.md` provide the final status. The isolated gate lock,
not the differing ambient Cargo.lock, defines validation dependencies.

Temporary cleanup is recorded in `cleanup.json`; retained raw measurements,
failed-attempt evidence and reconstruction inputs are deliberate deliverables.
