# Independent paired-statistics oracle review

Status: **draft, awaiting paired evaluator integration**. This review owns the
independent generator, retained numeric goldens, and focused Rust oracle for
`CORREL`, `COVAR`, `PEARSON`, `RSQ`, `SLOPE`, `INTERCEPT`, `STEYX`, and
`FORECAST`. It does not derive expected values from the Rust reducer or from a
spreadsheet host, and it makes no production-source change.

The semantic boundary is the local ODF 1.4 paired-statistics contract. The
current contract SHA-256 is
`9ce79b7b5014c6e29c1d30c61a77187d6dc4bce9be7810e1f5006e8ef147aadd`.
That contract selects complete ForceArray pairing, aligned positional
omission of Empty/Text/Logical cells, retained formula-error precedence,
population covariance, ordinary included-constant regression for
`INTERCEPT`, the explicit RSQ `#N/A` cases, the three-pair `STEYX` boundary,
and the finite query bridge for `FORECAST`.

## Corpus and derivation

`numeric_oracle.py` decodes every retained Number from its IEEE-754 binary64
bits. For each admitted pair it forms the means, `Sxx`, `Syy`, and `Sxy` as
exact `Fraction` values. COVAR, RSQ, SLOPE, INTERCEPT, and FORECAST publish the
finite Fraction result after one final binary64 conversion. CORREL and
PEARSON divide by a 300-digit Decimal square root only for that final
transcendental step. STEYX uses the exact centered residual before the same
high-precision square root. Exact algebraic zero is emitted as canonical
positive zero; a signed nonzero result that underflows keeps its sign.

The retained document has 39 paired fixtures and 464 observations: 58 rows
per function covering 39 References, 13 inline Arrays, five direct scalar
data cases, and one rejected ReferenceList. The fixtures include:

* exact, near-perfect, negative, fractional, constant-axis, one-pair,
  two-pair, and empty-pair lines;
* paired Empty, Text, numeric-looking Text, Logical, formula Error, and
  non-finite reference cells, including asymmetric omission positions;
* adjacent representables at large and tiny magnitudes, subnormals, finite
  extremes, signed cancellation, underflowed signed regression results, and a
  256-pair repeated line;
* three deterministic seeded wide-exponent data sets and the same-count
  row-versus-column shape fixture;
* different cell-count and list/pseudotype refusals with zero expected reads;
  and
* `FORECAST` Logical, Empty, decimal Text, malformed Text, a non-finite
  Number query, and query-error-before-later-data-error conversions.

The row-versus-column fixture has equal cell counts but unequal geometry, so it
is distinct from the shorter-array fixture. `RSQ` rows retain `#N/A` for the
different-count and no-admitted-pair cases, while shape and ReferenceList
refusals retain `#VALUE!`. Formula errors are recorded before generated
empty, cardinality, variance, or numeric errors. Reference rows require the
complete two-array read receipt; known shape/list refusals require zero
resolver reads.

## Comparison policy fixed before evaluator execution

The policy is retained in every numeric gold row and was fixed before running
the evaluator target:

* ordinary exact Fraction rows use `ORDINARY_MAX_ULPS = 0`, so their binary64
  bits must match exactly;
* sensitive large-offset, subnormal, cancellation, seeded, and repeated rows
  use a predeclared maximum of eight ULP;
* Decimal square-root rows use the same eight-ULP ceiling, with exact zero and
  exact unit correlations pinned to their exact bits; and
* signs must agree for every nonzero value, all compared values must be
  finite, and every expected zero checks its signed bits.

The eight-ULP ceiling is a finite-publication allowance for the selected
centered/scaled implementation and final square root. It is not a relative
epsilon and cannot hide an incorrect sign, an overflow, a zero-versus-
nonzero result, a domain error, or a shape/read violation. No tolerance was
changed after evaluator results.

## Focused target and current evidence

`crates/litchi-ods/tests/ods_formula_paired_oracle.rs` reconstructs the two
rectangular resolver arguments from the retained fixture JSON, maps the second
argument to a separate column region, evaluates every function in Scalar and
Matrix modes, compares the selected numerical/error policy, and checks the
expected cumulative reads. It therefore exercises complete paired references,
inline arrays, direct scalar data, formula-error completion, non-finite
conversion, shape/list read-free gates, and query conversion in one retained
target.

The generator and retained bytes currently verify with:

```text
python3 numeric_oracle.py --check
{"observations": 464, "fixtures": 39, "verified": true}
```

The focused Rust target is formatted, its no-run compile passes, and the
current integrated receipt is eight passing function tests in both Scalar and
Matrix modes. An earlier six-of-eight run exposed the constant-Y STEYX kernel
case; the production owner corrected that domain path to exact `+0`.

Revision note: the initial oracle version gave `FORECAST("bad";Y-with-#N/A;X)`
the generated query `#VALUE!`. Auditing the frozen contract's retained source
formula-error precedence corrected that row to the data `#N/A`; this was an
oracle semantic correction, not a tolerance change or an adjustment to fit
the evaluator. The retained generator check and focused receipt now agree.
No production or contract edit was made by this oracle review.

At the current handoff, the owned artifact hashes are:

```text
numeric_oracle.py          03ea2bdff5ba6aca8d5d8ed8775dd479e517bd2a73c2d0c617b968d6360c164e
numeric-goldens.json       038d8b16be3fcf8be2a6234a35ac91d5b5a819ba1b512a4dbcbcf7b802f92f2a
ods_formula_paired_oracle.rs  8c5251970f36a94819baf5f714f756c0eb03e8f3d7ee5bd9f3feff09e1ba175a
contract.md                9ce79b7b5014c6e29c1d30c61a77187d6dc4bce9be7810e1f5006e8ef147aadd
```

The goldens, generator, focused target, and this report should be frozen
together after the final source integration and the retained generator check.
