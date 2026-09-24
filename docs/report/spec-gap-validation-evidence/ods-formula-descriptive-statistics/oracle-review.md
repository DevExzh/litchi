# Independent numerical oracle review

Status: **draft, awaiting descriptive evaluator integration**. This review
owns the independent generator, retained goldens, and focused Rust oracle for
`AVEDEV`, `DEVSQ`, `GEOMEAN`, `HARMEAN`, `KURT`, `SKEW`, and `SKEWP`. It does
not infer expected values from the Rust implementation or from a spreadsheet
host.

The semantic boundary is the local ODF 1.4 descriptive-statistics contract,
whose current SHA-256 is
`6d0127bbf1867553860ba20013aab530c9931ea961e022532e9c16df4424965f`.
That contract selects `NumberSequence` for `DEVSQ` and `SKEWP`,
`NumberSequenceList` for the other five functions, `#DIV/0!` for empty
AVEDEV/DEVSQ/GEOMEAN/HARMEAN sequences, `#VALUE!` for moment cardinality or
degenerate constraints, signed odd-root GEOMEAN, and exact reciprocal-domain
rules for HARMEAN. It also requires a ReferenceList shape refusal with zero
resolver cell reads for DEVSQ and SKEWP.

## Corpus and retained observations

`numeric_oracle.py` constructs 53 represented-binary64 datasets and one
virtual `missing-slot` fixture. `numeric-goldens.json` retains 413 rows: 59
rows for each function, consisting of 48 logical References, five selected
ReferenceLists, five direct scalar calls, and one explicit missing-parameter
call. ReferenceList rows duplicate the referenced area in occurrence order;
the two singular-reference functions record a `#VALUE!` result and zero
expected resolver reads before scanning. The generator emits compact
prefix/repeat/suffix fixtures for the 4096-member repeated probes.

The oracle asserts exact zero resolver reads for direct scalar rows and for
the DEVSQ/SKEWP shape refusals. Admitted Reference and ReferenceList rows must
make observable resolver reads, while the row data intentionally does not
freeze an implementation-specific pass count. The resource target owns the
normal centered two-pass receipt; this numerical target permits the contract's
valid early exits for empty, degenerate, formula-error, or otherwise complete
primary scans and HARMEAN's optional accuracy replay.

The fixtures include typed reference cells and source-order formula errors;
singleton, empty, pair, triple, constant, and insufficient-cardinality cases;
signed zero, negative and mixed-sign values; exact HARMEAN cancellation in all
three permutations of `(3, 6, -2)`; a nonzero near-cancellation using the
adjacent representable below `-2`; `(minsub, -minsub, 1e100)` with the finite
`3e100` harmonic result; adjacent large-offset values using both neighboring
representables of `1e16`; near-maximum opposite and adjacent pairs;
subnormal and tiny-spread sequences; the focused `[-1, 1, min_subnormal]`
skew correction and its sign mirror; three order permutations of a large
cancellation sequence; deterministic seeded wide-exponent data; the nine-cell
high-dynamic-range signed moment sequence; repeated large/unit values; and
five harmonic limb regressions. Three limb regressions
contain 100 distinct odd integers near `2^52`, their signed counterparts, and
the residual `1`, so the exact reciprocal denominator is `1` and HARMEAN is
`201`. The distinct prime denominators have an LCM width of 5200 bits. The
large-residual variant uses represented binary64 `1e100`, leaving an exact
nonzero reciprocal denominator and a result of `201e100` at mathematical
scale; the zero-residual variant is exact `#DIV/0!`. The 201-valued cases are
permutation and cancellation regressions; the large-residual case is the
adaptive-exact-cancellation regression and avoids claiming that a rounded
fast path alone proves the fixed-limb boundary.
The typed fixtures exercise omitted reference Text,
Logical, and Empty members, scalar numeric Text and Logical conversion, and
formula Errors at first, middle, last, and multiple positions.

The focused subnormal rows require exact final bits: `SKEW` is
`8000000000000003` for the negative fixture and `0000000000000003` for its
mirror; `SKEWP` is `8000000000000001` and `0000000000000001`, respectively.
This pins the sample-correction-before-final-rounding boundary while retaining
the general eight-ULP policy for other sensitive rows.

The high-dynamic-range fixture retains all seven reducer observations under
the predeclared sensitive eight-ULP policy. Its independent moment references
are `SKEW = 3.0` and `SKEWP = 2.4748737341529163`; the values span roughly 40
decimal orders and include both signs and subnormal-scale terms.

## Independent derivation

The generator decodes each `Number` cell from its retained hexadecimal IEEE
754 binary64 bits. AVEDEV and DEVSQ retain the represented operands as exact
`Fraction` values, form the exact mathematical center, and round only the
published finite result to binary64. This keeps the center residual intact for
adjacent large values and three/four-point close sequences; it does not emulate
a rounded intermediate AVERAGE call. The exact DEVSQ and AVEDEV numerators and
denominators are retained in each gold row.

HARMEAN forms the reciprocal denominator as an exact Fraction before checking
for zero, so an exact `(3, 6, -2)` cancellation is `#DIV/0!` while a valid
nonzero near-cancellation remains numeric. Its finite result is then published
through a 240-digit Decimal context. GEOMEAN tracks zero and negative parity
separately and evaluates the mean of `ln(abs(x))` with 240-digit Decimal
`ln`/`exp`; it never multiplies the operands. A negative product has a real
odd root only when `n` is odd, while a negative product with even `n` is the
contract's `#NUM!` case. Exact zero is emitted as canonical positive zero.

KURT, SKEW, and SKEWP use exact Fraction means and residuals. The central
fourth moment and KURT's normalized expression are retained as an exact
Fraction because the standard-deviation square root cancels from the fourth
power. SKEW and SKEWP retain the central third moment as a Fraction, return
exact `+0` before normalization when it is zero, and use a 240-digit Decimal
square root only for a nonzero final normalization. This preserves symmetric
fixtures such as `[-9,-3,-1,-7,-5]` instead of manufacturing a ~1e-240
residue from Decimal cube cancellation. KURT uses sample deviation and the
bias-corrected excess formula; SKEW uses sample deviation and correction;
SKEWP uses population deviation. The oracle marks insufficient and zero-spread
cases as the contract's generated `#VALUE!` errors.

## Comparison policy fixed before evaluator execution

The retained policy is deliberately part of the JSON and is not adjusted to
match evaluator output:

* exact mathematical zeros must be canonical `+0`;
* `ORDINARY_MAX_ULPS=0`: ordinary exact Fraction rows use exact binary64 bits;
* `SENSITIVE_MAX_ULPS=8`: large-offset, subnormal, seeded, and other Fraction
  cancellation probes use a predeclared maximum of eight ULP;
* `TRANSCENDENTAL_MAX_ULPS=8`: nonzero Decimal logarithm/exponential/square-root
  results use a predeclared maximum of eight ULP, except the focused subnormal
  SKEW/SKEWP correction rows, which require exact bits;
* signs must agree for every nonzero comparison, and all compared values must
  be finite.

The eight-ULP ceiling is reserved before any Rust run for the final binary64
publication plus the bounded compensated/scaled or transcendental path. It is
not a relative tolerance and cannot hide a sign error, an overflow, a zero
versus nonzero result, or a center-collapse error. The corpus retains the
exact Fraction or high-precision Decimal reference alongside each expected
bit pattern so any accepted ULP difference remains auditable.

## Focused test and current evidence

`crates/litchi-ods/tests/ods_formula_descriptive_oracle.rs` reconstructs the
typed resolver from the retained fixture JSON, runs each expression in both
scalar and matrix modes, compares errors or the fixed numerical policy, and
checks exact zero-read refusals plus observable reads for admitted references.
It therefore checks list admission,
zero-read shape refusal, ordered duplicate occurrences, typed-cell filtering,
formula-error precedence, and matrix-mode complete-sequence behavior in the
same run as the numerical values.

The generator and retained bytes currently verify with:

```text
python3 numeric_oracle.py --check
{"observations": 413, "fixtures": 54, "verified": true}
```

Before the focused subnormal additions, the Rust target's first execution ran
all seven function tests with one pass (`HARMEAN`) and six failures; after the
GEOMEAN signed-zero evaluator fix, that pre-expansion execution had two passes
(`GEOMEAN` and `HARMEAN`) and five failures. Those historical failures were
recorded without widening the predeclared policy:

* SKEW and SKEWP `basic-odd` returned negative tiny residues instead of exact
  `+0`;
* AVEDEV `signed-zero` differed by one bit under its exact-bits row;
* DEVSQ `fractional` differed by two bits under its exact-bits row; and
* KURT `basic-odd` differed by three bits under its exact-bits row.

These were evaluator integration findings, not oracle or tolerance changes.
After the exact moment/replay fixes and the focused fixture additions, the
current retained target passes all seven reducer tests:

```text
cargo test --locked --offline -p litchi-ods --test ods_formula_descriptive_oracle -- --quiet
test result: ok. 7 passed; 0 failed
```

This receipt covers both scalar and matrix modes for all 413 observations,
including the exact subnormal SKEW/SKEWP rows and the high-dynamic signed
moment sequence. Freeze the generator, goldens, and this report together only
after the contract hash and retained `--check` output remain unchanged.
