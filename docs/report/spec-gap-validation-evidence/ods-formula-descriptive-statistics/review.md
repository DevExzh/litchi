# Independent descriptive-statistics semantic and numerical review

Status: **PASS** for the frozen implementation and descriptive-statistics
profile. This review covers `AVEDEV`, `DEVSQ`, `GEOMEAN`, `HARMEAN`, `KURT`,
`SKEW`, and `SKEWP`. It compares the frozen source with the repository-local
ODF 1.4 Part 4 source and the profile in [`contract.md`](contract.md). I made
no production or test edits.

## Frozen identity

The implementation freeze is [`gates/freeze.json`](gates/freeze.json), based
on `b8e5d5fe257fd95747c69a3c44a53cedd96f77ed`. The selected normative archive
is the repository-local ODF 1.4 archive
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`; its Part
4 HTML member is
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
The frozen contract is `6d0127bbf1867553860ba20013aab530c9931ea961e022532e9c16df4424965f`.

The selected source, test, oracle, and native-input hashes are recorded in
the freeze manifest. Its isolated `Cargo.lock` is
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`; the
workspace root lockfile is a different pre-existing copy and is not used as
the frozen gate input.

## Semantic findings

The implementation follows the contract's admitted-value rules and formulas:

* `AVEDEV` uses an exact fixed-width first pass and a replay of exact absolute
  residuals. `DEVSQ` forms the exact centered second numerator.
* `GEOMEAN` tracks zero and sign parity separately. It returns signed real
  odd roots for negative products with odd cardinality, rejects negative
  even products with the selected `#NUM!` error, and returns canonical `+0`
  for an exact zero product. No positive-only restriction is added.
* `HARMEAN` uses signed reciprocals. Mixed-sign nonzero inputs are admitted;
  an exact zero reciprocal denominator produces the selected `#DIV/0!`
  result. No extra positive-input restriction is added.
* `KURT` forms the full exact sample excess-kurtosis rational before the
  final binary64 publication. `SKEW` and `SKEWP` retain the exact centered
  third numerator and keep the sample correction distinct from the population
  normalization.
* Empty and cardinality or zero-spread failures use the contract's typed
  profile. Formula errors retain source order, while typed resolver,
  cancellation, source, resource, and allocation failures remain evaluator
  failures.

Reference-list admission matches the contract: `AVEDEV`, `GEOMEAN`,
`HARMEAN`, `KURT`, and `SKEW` admit lists; `DEVSQ` and `SKEWP` refuse them
before scanning. The resolver-backed path retains sequence order and does not
materialize the admitted cells. Scalar and matrix evaluator paths share the
same reducer semantics.

## Numerical review

The independent oracle is [`numeric_oracle.py`](numeric_oracle.py). It
decodes represented binary64 operands from the retained hexadecimal fixtures,
uses exact `Fraction` arithmetic for additive and reciprocal quantities, and
uses a 240-digit Decimal path for logarithm, exponential, and square-root
publication. It does not call the Rust evaluator or a spreadsheet host. The
fixed generator seed is `20260919`.

The retained corpus has 54 fixtures and 413 observations, 59 for each of the
seven functions. Its policy is fixed before evaluator execution: ordinary
exact rows require exact bits; sensitive Fraction and Decimal rows allow at
most eight ULP; exact mathematical zero requires canonical `+0`; a nonzero
result rounded to zero retains its sign. The retained command and result are:

```text
python3 numeric_oracle.py --check
{"observations": 413, "fixtures": 54, "verified": true}
```

I also ran supplementary, unretained diagnostics against the fixed-width
moment helper after the final shifted-borrow repair. They are not additional
freeze inputs: their local generators and scratch harnesses are not retained,
and the seed/category description below is a diagnostic record rather than a
claim of byte-for-byte regeneration. The committed
[`numeric_oracle.py`](numeric_oracle.py), its 413 retained observations, and
the frozen focused tests are the reproducible evidence. The diagnostics used
`random.seed(3)` for 1,000 additive cases and `random.seed(77)` for 800 skew
cases plus four explicit subnormal/sign cases; both decoded inputs with
`Fraction.from_float`, compared ordered binary64 bits, and applied the
predeclared eight-ULP ceiling only where the corpus marks a result sensitive:

* 1,000 seeded exact `Fraction` comparisons for DEVSQ, AVEDEV, and KURT,
  including randomized integer sequences;
* 804 high-precision SKEW/SKEWP comparisons over signed, wide-exponent,
  large-offset, and subnormal cases, with an eight-ULP ceiling for non-exact
  publication;
* the nine-cell high-dynamic-range signed sequence, whose exact references
  are `SKEW = 3.0` and `SKEWP = 2.4748737341529163`;
* the four exact-bit subnormal correction rows. The negative and mirrored
  `SKEW` rows publish `8000000000000003` and `0000000000000003`; the
  corresponding `SKEWP` rows publish `8000000000000001` and
  `0000000000000001`;
* a minimal shifted-subtraction regression representing `2^128 - 2`, run in
  both debug and optimized builds, which confirms borrow propagation through
  a zero high-carry limb.

The helper's fixed state retains no input-cell vector. Its raw moment widths
cover the binary64 exponent and significand bounds, while larger fixed scratch
widths are used for centered products and quotient rounding. The harmonic
path's exact reciprocal replay is bounded by the frozen resource profile and
uses adaptive exact work only when the first pass cannot prove the result.

## Validation receipts

The frozen focused targets are recorded by the root gate run: descriptive
evaluation `12/12`, resource limits `18/18`, native comparisons `2/2`, and
independent oracle `7/7`. The extraction reproduction has its separate
receipt in [`native-reproduction.json`](native-reproduction.json). The final
[`gates/results.json`](gates/results.json) records all seven successful gates,
including boundaries and diff-check; the isolated package run records 1,477
tests with zero failures or ignored tests, and
[`gates/verification.json`](gates/verification.json) confirms stable sources
and required checks. The native corpus is corroborating evidence for host
behavior; the local ODF pseudotype and shape profile remain authoritative.

This PASS is a semantic and numerical disposition for the frozen evidence. It
does not make a universal error-bound claim outside the retained binary64
corpus, and root owns the separate performance receipt and threshold
disposition.
