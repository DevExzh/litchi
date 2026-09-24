# ODF 1.4 elementary real mathematical evaluator contract

This contract defines the next bounded evaluator batch for eleven
deterministic, scalar real-number functions in ODF 1.4 Part 4 §6.16.  It is
deliberately narrower than complete OpenFormula conformance and than the
whole mathematical-function chapter.  The existing trigonometric and
rounding kernels, value projection rules, budgets, and read-only calculation
boundary remain in force.  The same scalar kernels must be used by the
resolver-free scalar evaluator and by the value evaluator's elementwise
array/reference bridge.

The normative source is the local
`part4-formula/OpenDocument-v1.4-os-part4-formula.html` member of
`3rdparty/specs/OpenDocument-v1.4-os.zip`:

| Source | SHA-256 |
| --- | --- |
| archive `OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| extracted HTML member | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The relevant general rules are §3.2.3 (eager argument evaluation and Error
propagation), §3.3 (non-scalar evaluation), §3.6 (numerical model), §4.3
(Number), §5.6 (function arity and syntactically empty parameters), §5.12
(Error values), §6.1 (leftmost Error recommendation), §6.2 (common function
template and constraints), and §6.3.5 (Conversion to Number).  The function
definitions are §6.16.2, §6.16.31, §6.16.39–§6.16.42, §6.16.46,
§6.16.48, §6.16.54, and §6.16.58–§6.16.59.

## Function set and signatures

The batch contains these eleven names.  Each function has the exact arity
shown below, except that `LOG` may omit its optional `Base`.  Invalid arity is
a formula `#VALUE!` result after the supplied arguments have been consumed.

| Function | ODF syntax | Return | Domain or additional constraint |
| --- | --- | --- | --- |
| `ABS` | `ABS(Number N)` | Number | none |
| `EXP` | `EXP(Number X)` | Number | none |
| `LN` | `LN(Number X)` | Number | `X > 0` |
| `LOG` | `LOG(Number N [; Number Base = 10])` | Number | `N > 0`; this real-number profile additionally requires `Base > 0` and `Base != 1` |
| `LOG10` | `LOG10(Number N)` | Number | `N > 0` |
| `MOD` | `MOD(Number A; Number B)` | Number | `B != 0` |
| `POWER` | `POWER(Number A; Number B)` | Number | special zero and nonpositive-base cases below |
| `QUOTIENT` | `QUOTIENT(Number A; Number B)` | Number | `B != 0` |
| `SIGN` | `SIGN(Number N)` | Number | none |
| `SQRT` | `SQRT(Number N)` | Number | `N >= 0` |
| `SQRTPI` | `SQRTPI(Number N)` | Number | `N >= 0` |

The tokenizer already recognizes these names.  Adding names to that catalogue
alone is not coverage: the names must be admitted by both evaluator dispatch
paths and must use the shared scalar kernels.  After this batch the evaluator
still does not implement the remaining §6.16 families, including Bessel,
combinatoric, conversion, error-function, factorial/gamma, integer
gcd/lcm, parity, sequence/aggregation, random, or sum-of-squares functions.
This batch therefore remains an explicit partial §6.16 implementation.

## Normative results and profile choices

The equations define the mathematical result, not a required algorithm
(§6.2).  This finite binary64 profile admits only finite `f64` Number values.
It returns `#NUM!` for a domain violation or a non-finite result that cannot
be represented as a finite Number.  It retains the existing locale-independent
Text-to-Number conversion, formula-error precedence, and eager source-order
argument evaluation.  A complex value is not a Number and produces the
existing `#VALUE!` conversion result.

* `ABS(N)` returns the nonnegative magnitude.  It must not allocate and should
  canonicalize either signed zero to `+0`.
* `EXP(X)` computes `e^X`, using the platform finite-`f64` exponential.  An
  overflow to infinity is `#NUM!`; ordinary underflow to a representable zero
  remains a Number result.
* `LN(X)` computes the natural logarithm and `LOG10(N)` computes the base-10
  logarithm.  Their strict positive domain check happens before calling libm,
  so `+0`, `-0`, and negative inputs return `#NUM!` rather than leaking an
  infinity or NaN.
* `LOG(N; Base)` defaults `Base` to 10 only when the optional argument is
  omitted.  An explicitly empty parameter is not silently treated as omitted
  under this bounded profile: it is an invalid Number argument and returns
  `#VALUE!`, consistent with §5.6's rule that functions need not accept empty
  parameters unless they say so.  The normative section states only `N > 0`,
  but a real logarithm also requires a finite `Base > 0` other than 1; this
  profile reports `#NUM!` for a zero, negative, or unit base.  The result is
  `ln(N) / ln(Base)`; a direct `log(Base)` call is acceptable after these
  checks.  No epsilon-based rejection or acceptance is used around `Base=1`.
* `POWER(A; B)` follows the existing infix `^` profile so the two spellings
  cannot disagree.  `0^0` returns `1`; `0^B` for negative `B` returns
  `#NUM!`; a negative `A` with a non-integer `B` returns `#NUM!`.  Integer
  exponents for a negative base use the finite `f64` power operation.  Any
  non-finite result, including an overflow or an unrepresentable reciprocal,
  is `#NUM!`.  The integer test is exact in the admitted binary64 domain; no
  epsilon snapping is applied.
* `SQRT(N)` returns the principal real square root.  A negative input returns
  `#NUM!`; zero is a Number zero.  `SQRTPI(N)` is mathematically
  `sqrt(N * PI())` with the same nonnegative domain and error policy.  It must
  avoid forming `N * PI` first: for a large finite `N` that product can
  overflow even though the square root is representable.  A stable profile
  implementation is `sqrt(N) * sqrt(PI)`, followed by the finite-result
  check.  This also avoids unnecessary underflow for tiny positive values.
* `SIGN(N)` returns `-1`, `0`, or `+1` according to the strict comparisons
  in §6.16.54.  Both signed zeros compare as zero and produce numeric zero;
  this function must not use `signum()` in a way that exposes a negative zero
  as a fourth result.
* `MOD(A; B)` returns a remainder with the same sign as `B`, as required by
  §6.16.42.  The recommended finite-`f64` operation is the platform remainder
  `A % B`, followed only when the nonzero remainder has the opposite sign from
  `B` by adding `B`.  Do not compute `A - B * floor(A / B)`: the quotient and
  product can round so that large finite inputs lose the low remainder bits.
  Do not use Rust `rem_euclid` unchanged because it is nonnegative even when
  `B` is negative.  An exact zero remainder may be canonicalized to `+0`.
  The operation is based on the already-evaluated binary64 operands; it does
  not promise exact integer arithmetic for a formula whose intermediate value
  has already rounded.  In particular, an upstream cached value that differs
  for a large quotient is evidence of producer variance, not a reason to
  replace the normative sign-of-divisor operation.
* `QUOTIENT(A; B)` returns the integer portion of `A / B`, using truncation
  toward zero for negative and positive results.  It is independent from
  `MOD`'s sign-of-divisor convention; for example, `QUOTIENT(-5; 2) = -2`
  while `MOD(-5; 2) = 1`.  A zero divisor, including either signed zero,
  returns `#DIV/0!`.  The finite quotient is truncated after division and a
  non-finite quotient/result returns `#NUM!`; no unchecked integer cast or
  unbounded-precision intermediate is introduced.

The profile uses ordinary platform/libm rounding for finite transcendental
results.  It does not claim cross-platform bit-identical output.  Exact
domain boundaries, signed-zero behavior where specified, division errors,
and finite-result checks are stable contract behavior; ordinary last-bit
rounding remains observable.

## Conversion, errors, and optional parameters

Arguments are evaluated eagerly in source order.  Formula Errors propagate
before Number conversion, with the first source-order Error retained according
to the existing evaluator policy.  Text and Logical arguments follow the
existing finite Number conversion.  References and inline arrays are refused
by the resolver-free scalar API and are projected/broadcast by the value API;
an Empty worksheet cell converts to Number zero under the existing value
profile.  A syntactically empty required argument is not an Empty worksheet
cell and returns `#VALUE!`.

Domain violations in the table return `#NUM!`; exact zero divisors return
`#DIV/0!`.  Cancellation, resource exhaustion, allocation failure, provider
failure, and unsupported capability remain typed `EvaluationFailure` values
and are not caught by `IFERROR` or `IFNA`.  No function in this batch is lazy;
the existing lazy behavior of `IF`, `IFERROR`, and `IFNA` continues to govern
unselected branches surrounding these calls.

## Evaluation, resource, and API boundary

The public API remains `evaluate_scalar` and `value::evaluate`; no libm handle,
resolver, AST node, package identifier, or host service is added to a public
signature.  Keep the family in an outlined dispatch module, with fixed-size
scalar state and no heap allocation in the kernels.  Existing evaluator work,
storage, stack, shape, reference-cell, cancellation, and owned-result limits
must remain effective.  The value evaluator applies the same kernel once per
broadcast cell.  A domain violation is a formula `Error` in that cell and is
retained in an otherwise complete array; cancellation, resource exhaustion,
allocation failure, and provider failure are typed evaluation failures and
must publish no partial array or owned result.

The batch does not recalculate worksheet dependencies, update formula caches,
spill results, resolve names or external sources, or execute volatile and host
services.  It does not claim a complete Small Group evaluator or complete
§6.16 coverage.  These boundaries follow ADR 0001's correctness-first typed
API policy, ADR 0004's focused semantic interfaces, and ADR 0005's finite
budgets, fallible allocation, and measured-performance requirements.

## Validation requirements

Tests must cover every listed name, exact arity, text/logical conversion,
formula-error precedence, lazy surrounding branches, scalar reference refusal,
matrix broadcasting, implicit intersection, cancellation, work/storage
refusal, and release of retained reservations.  Numerical vectors must include
both signs, signed zero, subnormal and extreme finite values, all strict domain
edges, `LOG` omitted and invalid bases, negative-base `POWER`, zero and
negative `POWER` exponents, `SQRTPI` values whose `N * PI` would overflow,
both signs of `MOD`'s divisor, large-quotient MOD cases, and truncation of
negative `QUOTIENT` results.  Expected values must come from the normative
equations or an independent high-precision/reference oracle rather than from
the candidate evaluator.  Native cached observations may corroborate ordinary
vectors but cannot redefine the MOD or other numerical profile.

Performance evidence may compare scalar, array, and reference entry points,
but it must retain the workload, toolchain, source identity, lockfile, and
allocator/RSS limitations.  It must not claim whole-workbook recalculation,
native producer acceptance, exact cross-platform agreement, or a language-level
resident-memory bound.
