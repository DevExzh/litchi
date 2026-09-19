# Independent order-statistics oracle review

Status: PASS for the frozen independent derivation. This review covers the
contract in `contract.md` (SHA-256
`d57a13579f769d35d750c95c3480c382ce6df4888c223bb3fc7a6145bd61611d`) and
does not use the Rust evaluator, LibreOffice, or any production numeric
kernel to produce expected values.

`numeric_oracle.py` generates 512 observations: eight functions over 56
fixtures. Every function has 56 single-reference rows and eight ordered
ReferenceList rows. The retained fixtures include typed reference cells,
formula errors, empty and singleton selections, repeated values and mode ties,
three source-order permutations, a fixed seed (`20260919`), finite extremes,
adjacent large and small values, subnormals, signed zero, a max/−max
interpolation midpoint, a negative-underflow midpoint, and compact 4,096-cell
repetition. ReferenceList rows duplicate occurrences intentionally so rank and
frequency counts remain observable; the MODE and QUARTILE rows record the
contract's zero-read pseudotype refusal.

The reference converts each stored binary64 bit pattern to `Fraction` before
sorting or comparing. Formula Number parameters are first parsed as their
represented binary64 value and then converted to `Fraction`, so rank arithmetic
uses the same represented `X`, `N`, `Quart`, `Order`, and `Significance` values
as the evaluator rather than decimal source text. MEDIAN, MODE, LARGE, SMALL,
integer-rank PERCENTRANK selection, and RANK use exact represented
comparisons. PERCENTILE and QUARTILE use the specification's one-based
`1 + X*(n-1)` rank and exact rational convex interpolation. PERCENTRANK uses
the first occurrence (strictly-less competition rank) for exact ties and exact
fractional bracketing between distinct values, then rounds to the contract's
decimal significance with ties away from zero. An exact rational is converted
to binary64 once. This makes the max/−max midpoint finite and avoids forming an
overflowing difference in the oracle.

Selected zero is required to be canonical `+0`. A computed nonzero negative
fraction that underflows retains Python's IEEE `-0` result; the retained
negative-underflow fixture therefore checks the contract's sign-preserving
midpoint rule. The PERCENTRANK rows include omitted significance, explicit
four-place precision, and a positive tie rounded away from zero. Formula errors
are retained in source order. Generated empty,
domain, mode, missing-rank, and shape failures use `#VALUE!`; non-finite
numeric publication uses `#NUM!`.

The comparison policy is fixed before evaluator execution. Bit equality is
required for direct selections, integer ranks, and either sign of zero.
Interpolated or decimal-rounded finite results allow at most four ULPs because
the production path may round once during safe binary64 interpolation or
decimal rounding; ordinary selected values use a two-ULP metadata bound but are
checked by exact bits. No tolerance is widened after observing evaluator
output. `python3 numeric_oracle.py --check` regenerates the retained JSON byte
for byte and reports 512 observations over 56 fixtures.

The frozen handoff records matching custody for the retained oracle test
(`b94a6f337ab21a1beec3d4f4a78dacefab8f3651333bb3affb42f9283ea7f6f9`),
goldens (`b3b7eb6b53535c1f55c1f9995f47fbc1c5c8d60c0f0cb55a710a47522c3bbce4`),
and staged generator (`88b541a5fbca4d53fb9744cf4155d4b24d1942465c35887fd2ef1662c9d1bb64`).
The frozen receipts report all eight order-oracle tests passing, 43 focused
tests passing (17 semantic, 16 resource, 2 native, and 8 oracle), and 1,421
isolated ODS tests passing with no failures or ignored tests. The prior source
freeze receipts recorded passing format, clippy, rustdoc, boundary, and diff
gates. After this contract-wording refresh and its direct MUNIT assertions,
final gate rerun `33103` passed. Root's final performance verification is
retained in [verification-receipt.json](verification-receipt.json) with status
`ok` and retained-manifest SHA-256
`3bbf194d378a5600328485cabe2f4eb7b94e259e7a244e975c4aea5fbcfb0484`; this
independent numerical review makes no timing claim. It records those receipts
without rebuilding or changing frozen source and inputs.
