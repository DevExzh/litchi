# ODF 1.4 trigonometric evaluator contract

This contract is the normative input and review boundary for the complete
ODF 1.4 Part 4 §6.16 trigonometric, hyperbolic, angle-conversion, and
constant batch. It is deliberately narrower than complete OpenFormula
conformance: it defines the finite `f64` scalar profile already used by the
ODS evaluator, and the value evaluator's existing bounded array/reference
projection around that scalar kernel.

The normative source is the local
`part4-formula/OpenDocument-v1.4-os-part4-formula.html` member of
`3rdparty/specs/OpenDocument-v1.4-os.zip`:

| Source | SHA-256 |
| --- | --- |
| archive `OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| extracted HTML member | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |
| review text extraction | `af2f4954409124c1c753ccd4d4d810f2321d43df06c2d455ab8a263a44be1105` |

The relevant general rules are §3.2.3 (eager argument evaluation and Error
propagation), §3.3 (non-scalar evaluation), §3.6 (numerical model), §4.3
(Number), §5.6 (function arity and syntactically empty parameters), §5.12
(Error values), §6.1 (leftmost Error recommendation), §6.2 (implicit
conversion and constraints), and §6.3.5 (Conversion to Number). The function
definitions are §6.16.3–§6.16.11, §6.16.19–§6.16.25, §6.16.45,
§6.16.49, §6.16.52, §6.16.55–§6.16.57, and §6.16.69–§6.16.70.

## Function set and signatures

The complete batch contains these 24 names. Every listed function has an
exact arity in this profile; an invalid arity is a formula `#VALUE!` result
after all syntactically supplied arguments have been consumed.

| Function | ODF syntax | Return | Domain or additional constraint |
| --- | --- | --- | --- |
| `ACOS` | `ACOS(Number N)` | Number | `-1 ≤ N ≤ 1` |
| `ACOSH` | `ACOSH(Number N)` | Number | `N ≥ 1` |
| `ACOT` | `ACOT(Number N)` | Number | none; principal result `0 < r < π` |
| `ACOTH` | `ACOTH(Number N)` | Number | `abs(N) > 1` |
| `ASIN` | `ASIN(Number N)` | Number | `-1 ≤ N ≤ 1` |
| `ASINH` | `ASINH(Number N)` | Number | none |
| `ATAN` | `ATAN(Number N)` | Number | none; principal result `-π/2 < r < π/2` |
| `ATAN2` | `ATAN2(Number x; Number y)` | Number | `x ≠ 0` or `y ≠ 0`; `(0;0)` is implementation-defined |
| `ATANH` | `ATANH(Number N)` | Number | `-1 < N < 1` |
| `COS` | `COS(Number N)` | Number | none |
| `COSH` | `COSH(Number N)` | Number | none |
| `COT` | `COT(Number N)` | Number | none; reciprocal pole is an Error |
| `COTH` | `COTH(Number N)` | Number | `N ≠ 0` |
| `CSC` | `CSC(Number N)` | Number | none; reciprocal pole is an Error |
| `CSCH` | `CSCH(Number N)` | Number | none; zero denominator is an Error |
| `SEC` | `SEC(Number N)` | Number | none; reciprocal pole is an Error |
| `SECH` | `SECH(Number N)` | Number | none |
| `SIN` | `SIN(Number N)` | Number | none |
| `SINH` | `SINH(Number N)` | Number | none |
| `TAN` | `TAN(Number N)` | Number | none; the finite representation determines pole proximity |
| `TANH` | `TANH(Number N)` | Number | none |
| `DEGREES` | `DEGREES(Number N)` | Number | none |
| `RADIANS` | `RADIANS(Number N)` | Number | none |
| `PI` | `PI()` | Number | none |

The catalogue already recognizes all 24 names. The implementation must wire
the names through the resolver-free scalar evaluator and through the value
evaluator's scalar bridge; adding a name only to the syntax catalogue is not
coverage.

## Normative results and principal branches

The ordinary and hyperbolic functions are real Number functions. They do not
accept the separate Complex value as a Number; a complex value therefore
produces the profile's `#VALUE!` conversion error. The mathematical definitions
are:

* `ACOS(N)` is arc cosine, with `0 ≤ result ≤ π`.
* `ACOSH(N)` is the principal inverse hyperbolic cosine. The displayed
  definition is `ln(N + sqrt(N² − 1))` for the real domain `N ≥ 1`.
* `ACOT(N)` is the principal arc cotangent in `(0, π)`. A stable real
  implementation is `atan2(1, N)`, rather than `π/2 − atan(N)`, because the
  latter can round a large positive input to zero or a large negative input to
  exactly `π`. The mathematical interval is open; a finite binary64 result
  may round an extreme finite input to a represented endpoint without making
  the input invalid.
* `ACOTH(N)` is the real principal inverse hyperbolic cotangent,
  `1/2 ln((N + 1)/(N − 1))`, for `abs(N) > 1`. The stable finite-`f64` form
  is `copysign(0.5 * ln1p(2/(abs(N) − 1)), N)`. Computing `1/N` first and
  calling `atanh` loses relative accuracy near `|N|=1`, even when the reciprocal
  remains strictly inside the `ATANH` domain. The sign-symmetric `ln1p` form
  avoids that loss and overflow/cancellation at large inputs.
* `ASIN(N)` is arc sine, with `−π/2 ≤ result ≤ π/2`.
* `ASINH(N)` is the principal inverse hyperbolic sine, with no real domain
  restriction.
* `ATAN(N)` is arc tangent, with mathematical range `−π/2 < result < π/2`.
  A finite binary64 result may round an extreme finite input to a represented
  endpoint; that is ordinary finite-result rounding, not a domain error.
* `ATAN2(x; y)` uses the coordinates in the ODF order `(x, y)`, so an API
  exposing Rust's `atan2(y, x)` must call `y.atan2(x)`. Its principal range is
  `−π < result ≤ π`; if the platform returns exactly `−π` for a lower-half
  signed-zero point on the negative x-axis, normalize it to `+π`. Preserve a
  rounded `−π` for a nonzero negative y-coordinate near that branch cut: it is
  the finite rounding of a valid angle and mapping it to `+π` would create a
  nearly-`2π` jump. This is a branch-axis rule, not an epsilon or quadrant
  heuristic. `(0;0)` may return zero or an Error; this profile chooses the
  documented formula `#NUM!` (`ScalarError::Number`).
* `COS`, `COSH`, `SIN`, `SINH`, `TAN`, and `TANH` use their real
  trigonometric or hyperbolic definitions in radians. `DEGREES(N)` is
  mathematically `N * 180 / π`; `RADIANS(N)` is `N * PI() / 180`.
* `COT(N) = 1/TAN(N)`, `CSC(N) = 1/SIN(N)`, and `SEC(N) = 1/COS(N)`.
  A denominator that is exactly zero produces the profile's division error.
  Do not classify near-zero results with an epsilon: `PI()` is a finite
  approximation, so `SIN(PI())` is not required to be exactly zero.
* `COTH(N) = 1/TANH(N)` for `N ≠ 0`, and `CSCH(N) = 1/SINH(N)`.
  `SECH(N) = 1/COSH(N)`. For large finite arguments, evaluate these
  reciprocal hyperbolic functions through a scaled exponential, such as
  `exp(LN_2 − abs(N))` (and restore the sign for `CSCH`), rather than first
  materializing `2*exp(-abs(N))`: the separately rounded exponential can lose
  a representable minimum subnormal. `COTH` through `tanh` also avoids a
  `sinh`/`cosh` overflow.
* `PI()` has no arguments and returns the closest representable π available
  to the selected finite Number representation. `std::f64::consts::PI` is
  the profile's value.

The equations define the mathematical result, not a required algorithm (§6.2
and §3.6). A direct libm call is suitable for ordinary finite inputs, but a
naive expansion of `ACOSH`, `ASINH`, `ACOTH`, `CSCH`, or `SECH` must not turn a
representable mathematical result into an avoidable NaN, infinity, or zero.

## Numeric and Error policy

This profile admits finite binary64 Numbers only. Numeric literals that parse
as NaN or infinity are `#NUM!`; Text-to-Number conversion follows the
existing locale-independent decimal policy and rejects non-finite or invalid
text as `#VALUE!`. A finite input whose final result is not representable as a
finite `f64` becomes `#NUM!`. Formula errors remain values and are propagated
before conversion; when several are present the argument in source order is
retained, matching §6.1's leftmost-error recommendation.

Domain violations in the table return `#NUM!`. Exact reciprocal zero uses
`#DIV/0!` in this profile, matching the existing infix-division policy. The
specification permits implementations to choose their non-`#N/A` Error names;
these names are a documented profile choice, not a claim that every evaluator
must serialize the same spelling. Typed cancellation, resource-limit,
allocation, unsupported-capability, and source/provider failures remain
`EvaluationFailure` values and are never caught by `IFERROR` or `IFNA`.

The profile preserves signed zero when the elementary operation is odd:
`ASIN`, `ASINH`, `ATAN`, `ATANH`, `SIN`, `SINH`, `TAN`, `TANH`, `DEGREES`, and
`RADIANS` of `-0` remain `-0`. Even functions do not preserve a negative-zero
sign. Reciprocal poles are errors, including `COT(-0)`, `CSC(-0)`,
`CSCH(-0)`, and `COTH(-0)`.

## Evaluation, arrays, and refusal boundary

All 24 functions are ordinary eager functions. The evaluator computes their
arguments in source order, charges the existing work budget, and then applies
the kernel. Trigonometric functions do not introduce a lazy branch or a
resolver. `IF`, `IFERROR`, and `IFNA` retain their existing lazy semantics
around a trig call: an unselected `SIN([.A1])` is not resolved, and a formula
domain error can be caught, while a cancellation or resource failure cannot.

The resolver-free `evaluate_scalar` profile intentionally refuses references,
arrays, names, labels, and reference operators with typed
`EvaluationFailure::Unsupported` values. Its trig tests must keep this refusal
explicit rather than pretending that a scalar literal test proves worksheet
evaluation.

The value profile uses the existing bounded projection rules. In scalar mode,
a one-cell reference or the implicit intersection of a rectangular reference
is converted to Number; an inline array contributes its `(0,0)` element. In
matrix mode, a trig function maps elementwise over arrays and materialized
rectangular references, broadcasts singleton dimensions, and returns the
element's formula Error at the corresponding position. Reference lists remain
distinct and cannot be silently flattened into a scalar or rectangular array.
Every reference cell, array cell, shape, and temporary argument vector remains
under the existing finite limits; no trig kernel allocates heap storage.

Empty worksheet cells continue to use the value profile's existing Number
conversion (`0`). A syntactically empty parameter is a missing argument and
is not the same as a blank cell; these one-argument functions therefore return
`#VALUE!` for an empty formula slot. Invalid arity must release all popped
values and preserve leftmost formula-error precedence.

## API, production, and validation requirements

The public boundary remains the existing evaluation API. Do not expose a
libm handle, resolver, package identifier, raw AST node, or format-specific
state in ordinary signatures. Keep the numerical family in an outlined module
so the scalar evaluator's hot dispatch stays small. Fixed-size scalar kernels
should use no heap allocation, check the retained `ExecutionContext` through
the existing evaluator operations, and return typed failures before publishing
partial arrays or results.

Tests must cover all 24 names, both signs and principal quadrants, every domain
edge, reciprocal poles, signed zero, large finite values, subnormal values,
near-`|N|=1` ACOTH/ATANH, `ATAN2` argument order and `−π` normalization,
`PI` arity, conversion/error precedence, lazy branches, scalar refusal,
matrix broadcasting, reference intersection, cancellation, work/storage and
reference-cell limits. Independent numerical vectors should be generated from
the ODF equations or an independent high-precision oracle; candidate output
must not generate its own expected values.

Performance evidence may compare scalar, array, and reference entry points,
but it must report the authored workload, toolchain, source identity, and
allocator/RSS limitations. It must not claim complete workbook recalculation,
native LibreOffice acceptance, exact bitwise agreement across libm platforms,
or a language-level resident-memory bound. These boundaries follow ADR 0001's
correctness-first typed API policy, ADR 0004's focused semantic interfaces, and
ADR 0005's finite budgets, fallible allocation, lazy reference state, and
measured-performance requirements.
