# Byte-position text batch completion

All seven OpenFormula 1.4 section 6.7 functions are implemented: FINDB,
LEFTB, LENB, MIDB, REPLACEB, RIGHTB and SEARCHB. The selected UTF-8 profile,
scalar/matrix behavior, and native compatibility boundary are documented in
[contract.md](contract.md). Native DBCS/code-page emulation, workbook dependency
recalculation, cache publication and other unsupported function families remain
outside this batch.

Independent [semantic review](semantic-review.md) and
[resource/cache review](resource-review.md) are PASS. Review exposed and closed
a projected-cache bug in computed text reducers and the distinction between
malformed and overflowing numeric Text. The permanent regression preserves
coordinate-specific AVERAGE results and SUM's complete matrix argument context.
Conservative computed-text cache refusal can repeat an invariant SUM scan; this
is an explicit correctness tradeoff, not an unchanged-cost claim for that case.

All seven isolated integration gates pass, with 1,605 tests, zero failed or
ignored tests, strict all-target clippy, rustdoc, both formatting checks, crate
boundaries and whitespace validation. The independent boundary oracle covers
1,376 integer/string observations. A separate native fixture contains 19 rows:
seven ASCII matches and ten non-ASCII divergences from LibreOffice. The complete
[independent verifier](verify.py) passed after removing the isolated checkout;
its [receipt](verification.json) verifies frozen source custody, all semantic
fixtures and all raw performance measurements.

The performance capture retains 840 baseline and 2,520 candidate samples, across
both evaluation phases with three warmups and fifteen fresh processes per case.
Matched allocation counts/bytes, peak live bytes, retained budget, work, reads,
output bytes and checksums are unchanged. Control median latency shifts range
from -5.33% to +3.30%, and RSS median shifts from -7.78% to +3.33%. These observed
shifts are not causal speedup claims. The [performance review](performance-review.md)
records host contention and interpretation limits. Cancellation receipts record
one successful read per child across four sticky-cancellation repeats (0.25 per
repeat); the generated integer-normalized table displays zero after flooring.
That display does not establish zero-read cancellation.

The retained frozen lock is `58b4be6c…`; the unrelated ambient root lock remains
`aa945c79…`. Gate and benchmark build directories, the isolated checkout and
owned Python caches were removed; [cleanup.json](cleanup.json) records their
removal. Raw measurements and provenance are retained intentionally. Unrelated
workspace changes are excluded from this batch.
