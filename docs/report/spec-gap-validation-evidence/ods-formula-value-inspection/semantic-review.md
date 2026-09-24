# ODF value-inspection and conversion semantic review

Review status: **SOURCE PASS; focused semantic gate PASS; source frozen;
seven integration gates PASS; performance acceptance pending**. This is an
independent source review of the sixteen-function contract in
[contract.md](contract.md). I made no production or test edits. Final source
hashes are recorded below from the frozen manifest.

The normative archive and Part 4/Part 3 member hashes are recorded in the
contract. The selected profile is the fixed English profile documented there:
1899-12-30 serial epoch without the 1900 leap-day, a 1930 two-digit-year
pivot, raw Empty/formula-error/Complex identity, and no ambient locale,
clock, or provider. The review covers scalar and value dispatch, parser
lexical/domain behavior, streamed references, matrix projection, and lazy
branch context.

## Findings resolved in the current parser and contract

The parser now excludes exponent-marked dash strings from date dispatch, so
signed negative exponents such as VALUE("-1e-3") reach the numeric grammar and
return -0.001. Its fractional-second accumulator now preserves digit order;
the parser tests cover .12, .001, long fractions, and invalid 60.x seconds.
These changes resolve the two parser findings from the provisional review.

The contract gate now requires two-digit years under the fixed 1930 profile and
explicitly records document null-year overrides as future evaluation-context
scope. Its current SHA-256 is
f227e861c55e3360d60d41a42ee7f101ba3c29bbc00d25d56cc01d80c3cd7923.

The scalar inspection bridge now restores AST Missing markers after popping
the stack placeholder, so omitted NUMBERVALUE separators use profile defaults
while explicit empty Text separators remain invalid. The value scheduler now
uses a complete Any-argument context for TYPE and N, including computed Arrays
under projected lazy branches. It preserves the current position for explicit
scalar descendants such as N(reference), and the existing direct MUNIT
first-element behavior remains unchanged. Focused converter and projected
inspection tests pass for these fixes.

## Validation

The focused evaluation target passes 6/6. Its scalar
TYPE(ISNUMBER([.A3:.A4])) case now records the scalar result 4 without an
incorrect read assertion, while the separate TYPE([.A3:.A4]) case resets the
counter and verifies the two-cell scan. The focused scalar converter and
projected-branch cases also pass.

## Frozen source identity

The frozen manifest is
[gates/freeze.json](gates/freeze.json), with 82 selected source/profile files
and status “source frozen; independent semantic/resource reviews and seven
integration gates pass; performance acceptance pending.” The relevant selected
hashes match the frozen tree:

| File | SHA-256 |
| --- | --- |
| evaluation.rs | b88159d778793e14583b5807bfc6ccdff9877eb59a19b6cdbc3514120db80392 |
| inspection.rs | 594fa0a8110296259532e20e7901685eeea9502eb8927f2915d58bfdf52d8918 |
| inspection/parse_value.rs | f29a42e77912a7dd65acc92a82f306fd9875e7d6a3410974029e940276ebf504 |
| value.rs | 9d30e3f77ad8b75a403d126c87259f7637041db52ede7441157c649bdec60e8d |
| value/inspection.rs | 4dfc0cf96e6bcd963906143b0ad07445e6883362073bb557a0fb6a139fff1a14 |
| contract.md | f227e861c55e3360d60d41a42ee7f101ba3c29bbc00d25d56cc01d80c3cd7923 |
| focused evaluation test | 7e882c07b6a6fb1ab1495d3d6ecbb5d23258728e33e6820ccdecb290e03ab8e1 |

## Reviewed and currently consistent

The function catalog and arities cover all sixteen names. Raw predicates,
ERROR.TYPE, and TYPE inspect formula-error values; parity, N, NUMBERVALUE, and
VALUE propagate them. The scalar and value kernels keep Empty, empty Text,
Logical, Number/Complex, and formula Error distinct, and the value mapper
streams admitted references through the existing cell-read and borrowed-text
boundary. TYPE scans complete multi-sheet cuboids for typed failures before
returning 64, while ordinary matrix operations use the documented current-sheet
plane profile and reject known reference lists before cell reads. Direct MUNIT
remains the matrix first-element profile, and the position-sensitive parameter
is represented by MUNIT(N(reference)).

The numeric, grouped, percentage, fraction, date, time, datetime, Complex, and
non-finite paths were reviewed against the contract. The source implementation
and focused semantic gate have no remaining blocker from this review. The
frozen source hashes match the manifest, and all seven integration gates pass;
only the separately tracked performance acceptance remains pending.

Disposition: **semantic source, focused validation, and frozen handoff PASS;
performance acceptance pending**.
