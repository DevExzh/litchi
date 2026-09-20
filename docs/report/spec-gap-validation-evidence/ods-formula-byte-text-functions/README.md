# ODS byte-position text functions

This batch implements all seven OpenFormula 1.4 section 6.7 functions:
FINDB, LEFTB, LENB, MIDB, REPLACEB, RIGHTB and SEARCHB. The [contract](contract.md)
selects semantic UTF-8 octets, backward snapping of interior starts and complete
Unicode scalar clipping. This is an implementation-defined profile, not native
DBCS/code-page emulation. Evaluation does not publish workbook cached values.

The scalar kernels use borrowed slices and checked output reservations. Byte
boundary clipping takes at most three continuation-byte adjustments. SEARCHB
reuses the pinned Unicode full-fold matcher and maps its match back to original
UTF-8 byte positions. Numeric Text overflow remains #NUM while malformed numeric
Text produces #VALUE. Resource, provider, cancellation and source failures remain
typed evaluation failures.

Review found a projected-cache bug: AVERAGE(LENB(reference)) reused the first
coordinate at later output positions. Computed reducers containing scalar text
functions now conservatively bypass this cache. A permanent regression verifies
coordinate-specific AVERAGE results and SUM's complete matrix-argument behavior.
This can reread invariant computed SUM arguments at each output; direct reference
reducers retain their existing cache classification. Performance evidence must
include this explicit correctness tradeoff.

The [oracle](byte_oracle.py) supplies 1,376 independent integer/string boundary
observations. The [native fixture](native/README.md) separately records seven
ASCII matches and ten non-ASCII differences from LibreOffice; a Rust integration
test checks the selected UTF-8 outputs for those same formulas. Gate source
custody uses [freeze.json](gates/freeze.json) and the retained
[dependency lock](gates/Cargo.lock), not the ambient workspace lock.

Status: frozen candidate with all seven integration gates, independent
resource/cache review and complete evidence verification passing. See
[verification.json](verification.json) and [performance interpretation](performance-review.md).
Independent [semantic review](semantic-review.md) is PASS; owned temporary files
have been removed. See [completion.md](completion.md) for final scope, evidence
and performance limitations.
