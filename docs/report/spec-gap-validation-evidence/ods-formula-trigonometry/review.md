# Independent ODF 1.4 trigonometry review

## Disposition

The source and validation review passes for the finite binary64 ODS profile.
The implementation covers all 24 requested functions in both scalar and value
evaluation paths. The root-owned final gate run passed all seven checks, with
1,193 ODS tests passing and zero failures or ignored tests; its source manifest
was stable before and after the run. This review records the independent
semantic and numerical decision.

## Source basis and identity

The normative source is the local ODF 1.4 Part 4 formula document. I verified
the same immutable inputs recorded by `contract.md`:

| Input | SHA-256 |
| --- | --- |
| `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| `.codex-tmp/odf14-part4.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |
| `.codex-tmp/odf14-part4.txt` | `af2f4954409124c1c753ccd4d4d810f2321d43df06c2d455ab8a263a44be1105` |

The reviewed production files are identified by these SHA-256 digests at the
final focused-test pass:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `a20687c9f99820982c1ece5ac0c2e812826c9c5f37235ea42ff5fdc8ce82d71f` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `228800a6276c8dd518672f4e1ad5c7bcd141a8fc0d14b8c6532ac93b27313f08` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs` | `232311fa431f8ea3711609c2262bdb60048094f90eba389f3cff1f6191aad7ba` |
| `crates/litchi-ods/src/codec/formula/evaluation/trigonometry.rs` | `c04bbfd1e128a621f19c760fbb9a4b37602aa11dfccdbe0955d1adf240524f2b` |

The native cache is corroborating evidence only. The retained receipt has 115
selected numeric observations across all 24 functions and one excluded
`ATAN2(0;0)` profile variance. It is not treated as a native acceptance or a
cross-platform bitwise-accuracy claim.

I also reviewed ADR 0001 (correctness-first typed API boundaries), ADR 0004
(focused semantic modules), ADR 0005 (finite budgets, lazy references, and
fallible storage), and ADR 0023 (ODS ownership of formula semantics). The
implementation and evidence stay within those boundaries.

## Semantic checks

The dispatch enum, resolver-free scalar path, and value scalar bridge contain
exactly these 24 names: `ACOS`, `ACOSH`, `ACOT`, `ACOTH`, `ASIN`, `ASINH`,
`ATAN`, `ATAN2`, `ATANH`, `COS`, `COSH`, `COT`, `COTH`, `CSC`, `CSCH`,
`SEC`, `SECH`, `SIN`, `SINH`, `TAN`, `TANH`, `DEGREES`, `RADIANS`, and `PI`.
The catalogue and evaluator dispatch are both wired; catalogue presence alone
is not being counted as coverage.

The domain checks match Part 4: `ACOS` and `ASIN` use `[-1,1]`, `ACOSH` uses
`[1,∞)`, `ACOTH` uses `abs(N)>1`, `ATANH` uses `(-1,1)`, and `COTH` rejects
zero. The inverse functions use the stated real principal branches. `ATAN2`
uses ODF's `(x; y)` order and calls the equivalent `atan2(y, x)` operation;
the profile reports `#NUM!` for its permitted implementation-defined `(0;0)`
case.

The exact negative-x, zero-y `ATAN2` branch maps a returned `−π` to `+π`.
Nonzero lower-quadrant y values retain a rounded `−π`, avoiding a false nearly
`2π` jump. `ACOT` uses the positive principal `atan2(1,N)` branch. The finite
binary64 representation may round mathematical open endpoints for extreme
finite inputs; those rounded results are not domain failures.

Formula errors are retained as values and selected in source order. Domain
violations and non-finite final results use the documented finite-profile
`#NUM!`; invalid text and wrong value types use `#VALUE!`; exact reciprocal
zero uses `#DIV/0!`. No epsilon pole detection is present, so finite `PI()`
rounding does not force `SIN(PI())` to zero. Odd elementary functions preserve
negative zero, while reciprocal poles reject both signed zeros.

## Numerical checks

The implementation avoids avoidable loss in the cases where a direct formula
would corrupt a defined result:

* `ACOTH` uses the sign-symmetric `copysign(0.5*ln1p(2/(abs(N)-1)), N)` form,
  retaining the finite result immediately above and below both unit endpoints
  and the tiny result for very large finite inputs.
* `ACOSH` and `ASINH` factor large arguments before squaring; `CSCH` and
  `SECH` use a split half-exponential so the minimum-subnormal tail remains
  representable around 745.5–745.8. `COTH` uses `1/tanh(N)` and does not
  overflow an intermediate `sinh`/`cosh`.
* `DEGREES` and `RADIANS` use the standard finite factor conversions. Every
  kernel publishes only finite `f64` values; an unrepresentable result becomes
  the profile's Number error.

The independent high-precision vectors in `numeric_oracle.py` use mpmath 1.3.0
at 600 decimal places from exact binary64 inputs. The large ACOTH, ACOSH,
ASINH, SECH, and CSCH cases use a same-sign four-ULP acceptance window. Exact
checks remain for signed zero and the minimum-subnormal tail, whose
representation is the semantic behavior under review.

The profile deliberately observes host libm rounding and does not promise
bitwise equality across platforms. The equations and branches are normative;
the ordinary finite-operation accuracy is a platform property.

## Evaluation boundary and validation

All 24 calls remain eager. `IF`, `IFERROR`, and `IFNA` retain lazy selection,
so an unselected reference is not resolved and a formula-level domain error can
be caught. Cancellation, work, storage, reference-cell, and unsupported
capability failures remain typed evaluation failures. The scalar profile keeps
references and arrays as typed refusals. The value profile applies implicit
intersection in scalar mode, elementwise matrix/reference broadcasting in
matrix mode, per-cell formula errors, and finite shape/reference/storage
limits. The kernels allocate no heap storage.

Independent focused validation observed before this review was written:

```text
cargo test -p litchi-ods --test ods_formula_trigonometry_evaluation -- --nocapture  # 11 passed
cargo test -p litchi-ods --test ods_formula_trigonometry_arrays -- --nocapture      # 8 passed
cargo test -p litchi-ods --test ods_formula_trigonometry_native -- --nocapture      # 1 passed
```

The tests include all 24 scalar vectors, domain and arity boundaries, both
ACOTH signs near `|N|=1`, signed zero, subnormal reciprocal-hyperbolic tails,
the exact and near `ATAN2` branch cut, scalar intersection at a non-origin
cell, lazy unresolved references, matrix broadcasting, per-cell errors,
cancellation, work/reference/shape limits, and storage refund after refusal.

The retained final gate receipt also reports successful locked ODS tests,
strict Clippy, rustdoc, crate formatting, batch formatting, crate-boundary
checks, and diff hygiene. `gates/verification.json` records both stable source
manifests and all required checks passed.

## Final performance review

Independent bounded harness/custody review also passed for the final 300
baseline and 1,500 candidate observations. The
[final review note](performance/results/final-review.md) records the scope and
the planned-but-not-captured baseline refusal receipt.
