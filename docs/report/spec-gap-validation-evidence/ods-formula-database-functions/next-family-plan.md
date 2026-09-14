# Next OpenFormula evaluator family

This is a planning note, not an implementation or a claim of complete
OpenFormula coverage. The normative source is the local ODF 1.4 Part 4 HTML
member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` in
`OpenDocument-v1.4-os.zip` (archive SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`, member
SHA-256 `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`).

## Recommendation: Part 4 §6.17 Rounding Functions

Implement the complete eight-function family as one bounded scalar batch:

`CEILING`, `INT`, `FLOOR`, `MROUND`, `ROUND`, `ROUNDDOWN`, `ROUNDUP`, and
`TRUNC`.

The current function catalog names these functions, but the value evaluator's
scalar admission currently only routes the logical/bitwise, radix, Roman, and
complex families through the scalar bridge. The value-specific dispatch adds
database and matrix handling; there is no §6.17 kernel. This is also consistent
with the current [feature matrix](../../../../crates/litchi-ods/docs/FEATURE_MATRIX.md),
which describes the implemented families without claiming the remaining Part 4
families.

The family is a useful next dependency boundary: it is complete, pure, and
resolver-free, while `INT` and decimal rounding are prerequisites for much of
§6.16 Mathematical and §6.18 Statistical work. It avoids silently introducing
worksheet lookup, date epochs, external services, or locale state in the same
batch. It is a full normative family rather than a hand-picked subset.

The exact rules to preserve from §§6.17.1–6.17.8 are material. `INT` floors;
`ROUND` uses powers of ten and rounds halfway values away from zero;
`ROUNDDOWN` truncates toward zero; `ROUNDUP` rounds away from zero; and `TRUNC`
truncates at the requested decimal position. `MROUND` chooses the nearest
multiple, choosing the greater multiple on a tie. `CEILING` and `FLOOR` require
compatible signs for number and significance, define omitted or Empty
significance from the sign of the number, and distinguish mode zero from a
nonzero mode. Zero number or significance returns zero. Optional-argument
presence must remain distinct from an explicit Empty slot.

No ODF host property is required for this family. The implementation should
explicitly reuse the evaluator's locale-independent finite-`f64` profile and
document its generic Number conversion. It should not infer locale or
`HOST-PRECISION-AS-SHOWN`; those are separate host choices under §3.4.

## Safety and test gates

Digits, significance, and mode need checked finite-to-integer conversion;
NaN, infinity, out-of-range casts, and formula errors must become the existing
typed formula errors without panics. Scaling by powers of ten can overflow or
underflow even when the rounded result is finite, so the kernel needs checked
scaling and a finite-result check rather than an unchecked `pow`/cast sequence.
`MROUND`, `CEILING`, and `FLOOR` likewise need checked quotient and product
paths for very large finite operands. The work is constant per scalar call and
allocates no text or resolver state, but matrix/reference broadcasting must
retain the existing output-shape, work, storage, cancellation, and borrowed
value limits.

Tests should cover positive and negative values, zero and Empty optional
arguments, all CEILING/FLOOR modes and sign combinations, halfway ties,
positive and negative digits, huge magnitudes, underflow/overflow boundaries,
non-finite inputs, propagated formula errors, scalar implicit intersection, and
matrix broadcasting. Independent decimal expected values are preferable to a
checksum-only oracle.

## Sequencing note

The next text-oriented family is §6.7 Byte-position text (`FINDB`, `LEFTB`,
`LENB`, `MIDB`, `REPLACEB`, `RIGHTB`, `SEARCHB`), but it should follow a
deliberate text policy: §6.7.1 leaves byte representation explicitly
implementation-dependent, while the ordinary counterparts are in §6.20 and
are not yet evaluator kernels. A UTF-8 byte profile, interior-code-point
boundary behavior, and the §3.4 search regex/wildcard host properties must be
specified before claiming interoperable results. Date/time, lookup, and
information families have larger epoch, resolver, or worksheet-type
dependencies and are poor candidates for the next isolated scalar batch.
