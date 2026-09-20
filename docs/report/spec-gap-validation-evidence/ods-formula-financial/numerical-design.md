# Financial numerical design review

Status: bounded implementation design for the twenty-function financial
contract. This document is a proposal for semantic and numerical review; it
does not claim that the functions are implemented or that the proposed profile
has been accepted.

The normative inputs are the repository-local ODF 1.4 Part 4 archive and the
[financial contract](contract.md). The archive is
`3rdparty/specs/OpenDocument-v1.4-os.zip`, whose financial definitions are in
`part4-formula/OpenDocument-v1.4-os-part4-formula.html`. The contract records
the archive and extracted HTML SHA-256 values. The direct root-solving rules
are §§6.12.24 (`IRR`), 6.12.42 (`RATE`), and 6.12.51 (`XIRR`); the residual
functions are `NPV` (§6.12.30) and `XNPV` (§6.12.52).

The labels in this note have a precise meaning:

- **Normative** is required by ODF or already fixed by the repository contract.
- **Profile choice** is a concrete proposal needed to make evaluation bounded
  and reproducible. It must be accepted before tests are treated as an oracle.
- **Open** is a question that the semantic or numerical reviewer must resolve;
  an implementation must not silently choose a host-library behavior.

## Findings

ODF defines the equations and permits an approximate iterative result for
`IRR`, `RATE`, and `XIRR`, including an error when an attempt does not
converge. It does not define a tolerance, iteration count, bracketing method,
or choice among multiple roots. The implementation therefore needs a visible
profile for all of these decisions.

The proposed positive-base branch evaluates a rate through

```text
u = ln(1 + rate)
rate = exp(u) - 1
```

and performs the root search in `u`. This keeps rates close to zero
well-conditioned and prevents an iteration from accidentally evaluating a
fractional-power denominator at or below `-1`. This transform is not a
universal domain rule: an integer-period expression can have a real negative
base when `rate < -1`, as described below.
The input sequence for a root function is read once into bounded numeric
scratch after its complete source-order scan. Root evaluations never reread a
resolver range. This is necessary for deterministic results and makes read
and cancellation accounting auditable.

The suggested initial limits are:

| Item | Proposed profile value | Status and reason |
| --- | ---: | --- |
| Default guess | `0.1` | Normative default in ODF. |
| Positive-base root branch | `rate > -1` | Required for `XIRR`/`XNPV` fractional date exponents and proposed for the common `IRR`/`RATE` root branch. `IRR` and integer-period `NPV` do not normatively forbid a negative base; that separate branch is called out below. |
| Transformed lower bound | `u_min = ln1p(r_min)`, where `r_min` is the next representable `f64` above `-1` | Profile choice; avoids evaluating `rate == -1`. |
| Transformed upper bound | `u_max = ln1p(f64::MAX)` | Profile choice; covers every finite positive `f64` rate while keeping bracket exploration finite. The final `expm1` result is checked and can be clipped to the greatest finite rate if rounding reaches infinity. |
| Negative-base branch | `v = ln(-(1 + rate))` | Profile choice for `IRR` and integral-`Nper` `RATE` when the supplied guess is below `-1`; `rate = -1 - exp(v)`. |
| Negative-base bounds | `v_min = ln(-(1 + nextbelow(-1)))`, `v_max = ln(f64::MAX)` | Profile choice; covers representable finite rates below `-1` without treating the branch as available to fractional date/exponent functions. |
| Bracket expansion steps | `64` | Profile choice; each step doubles the distance in the active transformed coordinate (`u` or `v`). |
| Root iterations | `128` | Profile choice; the caller's work budget remains an independent lower bound. |
| Residual/derivative evaluations | `256` total per root call | Profile choice; bracket probes and solve iterations share this cap. |
| Relative residual tolerance | `16 * f64::EPSILON * scale(z)` | Profile choice, where `scale(z)` is the current bounded sum of absolute residual terms in the active transformed coordinate. There is no `max(1, ...)` floor. |
| Transformed-rate step tolerance | `16 * f64::EPSILON * max(1, abs(z))` | Profile choice for active coordinate `z` (`u` or `v`); both residual and step/bracket criteria are required. |
| Log-magnitude span in a scaled sum | `2048` natural-log units | Profile choice; crossing it refuses explicitly rather than discarding a tail that could become visible after cancellation. |

The values above are review inputs, not magic compatibility constants. A
different accepted profile must update the contract, tests, and resource
receipts together.

The evaluation cap is authoritative when it meets the iteration cap first:
the initial residual, bracket probes, and safeguarded solve evaluations all
share the same `256` counter. Thus `128` is the maximum number of solve
iterations, not a promise that 128 iterations remain after every possible
bracket probe.

## Stable scalar kernels

All kernels receive finite `f64` values and return a finite `f64` or the
existing numeric formula error. Every intermediate that can become non-finite
is checked before it is exposed. This follows the existing evaluator practice:
`elementary.rs` maps non-finite numeric results to `ScalarError::Number`, while
`numerics.rs` and `dyadic.rs` use fixed-width state and explicit overflow
refusals instead of silently dropping bits.

When `1 + rate` is positive, let `u = ln1p(rate)`. The common positive-base
annuity helpers should be written as follows:

```text
growth(rate, n) = exp(n * ln1p(rate))
annuity_factor(rate, n) =
    n                                      if rate == 0
    expm1(n * ln1p(rate)) / rate otherwise
due_factor(rate, pay_type) = 1 + pay_type * rate
```

The original input `rate` remains the denominator; reconstructing it through
`expm1(ln1p(rate))` adds an avoidable rounding perturbation. `expm1` is still
important for the numerator when `rate` or `n * rate` is small. The zero branch
is exact and is required by the ODF equations. A non-finite `n * u`,
exponential, quotient, or final product returns the established numeric formula
error.

The positive-base branch is not the only real-valued branch of every financial
equation. For an integer exponent, `(1 + rate)^n` is real for `rate < -1` as
well. The implementation should use a checked signed integer-power path in
that case, preserving the sign parity and refusing only a zero denominator,
non-finite intermediate, or final non-finite result. If an exponent is
non-integer and the base is negative, the real-valued profile returns its
numeric error rather than manufacturing a complex result. This distinction is
required for integer-period `NPV` and is a deliberate guard against applying
`ln1p` to every non-iterative annuity calculation.

`FV`, `PV`, and `PMT` should use the contract's balance equations with
`growth`, `annuity_factor`, and `due_factor`. `IPMT`, `PPMT`, and `CUMIPMT` /
`CUMPRINC` should derive their period balances from those same helpers instead
of maintaining a second power or annuity implementation. `NPER` should use
the source's zero-rate branch and the stable logarithmic form of the nonzero
branch; if a log argument is not positive, it returns the profile's numeric
error rather than a NaN. `ISPMT` is linear in the period and does not need a
root solver.

The other scalar conversions should use the corresponding stable identities:

| Function | Stable evaluation | Required checks |
| --- | --- | --- |
| `EFFECT` | `expm1(payments * ln1p(rate / payments))` | Apply the normative `rate >= 0` and `payments > 0` constraints before arithmetic. |
| `NOMINAL` | `payments * expm1(ln1p(effective_rate) / payments)` | Apply the normative positive constraints and check the final product. |
| `PDURATION` | `(ln(specified) - ln(current)) / ln1p(rate)` | Positive values and `rate > 0`; reject zero/non-finite denominator. |
| `RRI` | `expm1((ln(abs(fv)) - ln(abs(pv))) / nper)` when `fv/pv > 0` | `nper > 0`; same-sign negative `pv`/`fv` is valid for the ratio. Zero or opposite-sign values need the accepted domain/error mapping. |
| `NPV` | Positive base: `value_i * exp(-i * ln1p(rate))`; negative base: checked signed integer power | Preserve one-based indices and argument/row-major source order. ODF prints no rate constraint for NPV, so `rate < -1` cannot be rejected solely because `ln1p` is unavailable. |
| `XNPV` | `value_i * exp(-((date_i-date_0)/365) * ln1p(rate))` | Enforce `rate > -1`, equal element counts, numeric elements, and the contract's date ordering. |

`FVSCHEDULE` has no stated positivity constraint on each schedule element, so
it cannot blindly take `ln1p(schedule_i)`. The proposed implementation reuses
the existing checked `ProductTerm`/`ScaledProductSum` machinery: each factor
`1 + schedule_i` is admitted in source order, decomposed into sign, mantissa,
and exponent, and committed without an overflowing intermediate product. A
zero factor makes the final result zero while the scan still continues for
formula-error and typed-failure precedence. An exponent-span refusal is
explicit; it must not silently lose a factor that could become visible after
cancellation. A factor for which `1 + schedule_i` is non-finite is rejected
before the product state is changed. This preserves the repository's checked
product behavior and avoids the avoidable rounding error of a sum-of-logs
product.

For `MIRR`, keep separate bounded accumulators for positive and negative
cash-flow contributions and use `log1p`/`expm1` for their rate powers. The
explicit requirement for at least one positive and one negative value is
normative. The admission of Logical values, References, and ReferenceLists is
still open in the contract and must be settled before choosing the exact
accumulator shape.

## Root residuals and deterministic solving

### Residuals in transformed rate

For a periodic cash-flow vector `c_i`, define the residual in `u` as

```text
f_irr(u)  = sum_i c_i * exp(-i * u)
f'_irr(u) = -sum_i i * c_i * exp(-i * u)
```

where `i` begins at one, as required by `NPV`. For `XIRR`, replace `i` with
`t_i = (date_i - date_0) / 365`:

```text
f_xirr(u)  = sum_i c_i * exp(-t_i * u)
f'_xirr(u) = -sum_i t_i * c_i * exp(-t_i * u)
```

`RATE` uses the derivative of its contract balance equation in the active
coordinate (`u` or `v`); the same `growth`, annuity, and due-factor helpers
must be used for the residual and the derivative. Derivatives are an
optimization and a conditioning aid, not a change to the equation. If a
derivative is non-finite or too small to produce a candidate inside the current
bracket, the solver falls back to a secant or bisection step.

On the negative-base `IRR` branch, the same residual is evaluated as
`sum(c_i * (-1)^i * exp(-i*v))` and its `v` derivative has the corresponding
`-i` factor. The integral-`Nper` `RATE` branch uses signed powers in the
balance equation and differentiates their positive magnitudes in `v`. These
branches must use the same finite checks and scaled signed sum as the
positive-base branch; they are not permitted to call `ln1p` on a negative
base.

At each residual point `z` (`u` or `v`), calculate two fixed-width quantities:
the signed residual `f(z)` and `scale(z)`, the sum of the absolute magnitudes
of the same terms. Convergence uses
`abs(f(z)) <= 16 * f64::EPSILON * scale(z)`. If `scale(z) == 0`, there is no
nonzero residual input and the sign-variation precondition fails; do not turn
that case into an absolute tolerance of one. If a scaled sum cannot represent
the term span, return the explicit numeric/profile refusal instead of silently
accepting a small residual.

The positive-base coordinate is the natural branch for `XIRR`, whose date
exponents are generally fractional, and for `RATE` when `Nper` is non-integer.
`IRR` has integer period exponents, so its equation can also have real roots
below `-1`. The proposed profile supports that branch when the caller supplies
a guess below `-1`: use `v = ln(-(1 + rate))`, evaluate each periodic term as
`cashflow_i * (-1)^i * exp(-i*v)`, and apply the same safeguarded solver in
`v`. For `RATE`, use the negative-base branch only when `Nper` is a positive
integer, with the signed integer-power balance equation. A guess of exactly
`-1` is handled by the explicit boundary rule below. `XIRR` remains on its
positive-base branch because its date exponents need not be integral. `NPV`
itself still supports the checked signed integer-power path described above.

The residual evaluator must not form a raw `c * exp(...)` when the product may
overflow before cancellation. It should decompose each nonzero term into a
sign and log magnitude, then feed it into a fixed-width, normalized signed
sum. The existing `ScaledProductSum` and fixed-width sum kernels provide the
right resource model; a financial-specific variant may be needed because the
exponent is produced by a real logarithm rather than an exact binary product.
The accumulator must either retain a term or return an explicit numeric/profile
refusal when the `2048`-unit span is exceeded. It must never drop a small term
because it is currently below an `f64` addend: a later cancellation can make
that term determine the sign of the residual.

The absolute term scale used by the residual tolerance is accumulated by the
same bounded scaled state at each candidate coordinate. If that scale itself
cannot be represented, the solver returns a numeric error instead of using an
infinite or unit-sized tolerance.

### Bracketing and iteration

The following is the proposed deterministic profile:

1. Validate the collected values and the guess. A supplied guess must be
   finite. A guess above `-1` selects the positive-base coordinate. A guess
   below `-1` selects the negative-base coordinate only for `IRR` and for
   `RATE` with positive integral `Nper`; otherwise it produces the accepted
   numeric/profile error. The omitted guess is `0.1`. This branch selection is
   a **profile choice** because ODF calls the guess an initial estimate and
   does not specify behavior below the real fractional-power domain.
2. Map a positive-base guess to `u0 = ln1p(guess)`, or a negative-base guess
   to `v0 = ln(-(1 + guess))`, and evaluate the corresponding residual. An
   exact zero is returned after converting the coordinate back to a finite
   rate. A `RATE` guess exactly equal to `-1` first uses the explicit boundary
   evaluation below.
3. Probe outward from the active coordinate with radii `1, 2, 4, ...`, clipped
   to its full representable branch bounds. At each radius, inspect endpoints
   in increasing distance from the guess. If two brackets are equally near,
   choose the one whose resulting rate is numerically lower. Keep the first
   sign-changing bracket found, and return an endpoint immediately if it is
   exactly zero.
4. Solve the selected bracket with a safeguarded Newton/secant/bisection
   hybrid. Accept a Newton or secant candidate only when it is finite and
   strictly inside the bracket. Otherwise bisect. Update the endpoint with
   the same sign as the candidate residual. Check cancellation before every
   residual/derivative evaluation and after it.
5. Stop only when the residual is within `16 * EPSILON * scale(z)` and the
   transformed-rate step or bracket width is within
   `16 * EPSILON * max(1, abs(z))`. Convert with `expm1(u)` on the positive
   branch or `-1 - exp(v)` on the negative branch; reject a non-finite result.
6. If the bracket or evaluation cap is reached without convergence, return
   the accepted numeric error for nonconvergence. Do not call a platform
   financial library as a fallback.

This scan chooses the closest sign-changing bracket to the supplied guess,
with lower-rate tie-breaking. That makes a multiple-root result deterministic
while preserving the purpose of the ODF guess. It does not promise complete
root discovery, but roots bracketed meaningfully around the supplied guess are
handled by the bounded scan. A tangent root with no sign change is reported as
no bracket. Those are intentional consequences of this bounded profile,
subject to review.

The transformed search has two useful safety properties. A negative rate is
allowed whenever it is greater than `-1`, including rates close to `-1`; the
integral-period negative-base branch also covers rates below `-1`; and a step
can never cross a branch boundary merely because a Newton update was large.
`u_max = ln1p(f64::MAX)` and `v_max = ln(f64::MAX)` cover the corresponding
representable finite-rate domains while still bounding search. They are
numerical search limits rather than ODF rate constraints; values whose
residual or final result overflows still produce the accepted numeric error.

### The `RATE` boundary at `rate == -1`

For a positive integral `Nper`, `RATE` has a finite algebraic boundary value
at `rate == -1` even though `ln1p(-1)` is unavailable. In the balance
equation, `(1 + rate)^Nper` is `0`, the annuity factor
`((1 + rate)^Nper - 1) / rate` is `1`, and the due factor is
`1 - PayType`. Evaluate that guarded residual directly, with finite checks and
the same scale-relative tolerance. If it is zero within the accepted profile,
return `-1` exactly. If it is not zero, do not report a domain error merely
because the positive-base transform is undefined: continue with the adjacent
positive and negative integral branches in deterministic distance order (the
negative branch wins an equal-distance tie because it has the lower rate), or
report bounded nonconvergence after their caps are exhausted. For non-integral
`Nper`, `rate == -1` has no real-valued power under this profile and follows
the normal numeric error mapping.

### No-root, multiple-root, and overflow policy

The following are **profile choices**, pending approval:

- no sign-changing bracket, no positive/negative cash-flow variation, an
  invalid transformed domain, non-finite residual state, and nonconvergence
  map to the existing numeric formula error (normally `#NUM!`);
- a multiple-root input returns the first sign-changing bracket in the
  deterministic outward order above, rather than the smallest root globally;
- a root at an endpoint is accepted only when the residual satisfies the same
  finite tolerance, unless it is exactly zero; and
- an intermediate that cannot fit the fixed-width scaled state is a numeric
  error, never an infinity, NaN, or silently rounded cancellation.

The contract leaves the exact error category open. The implementation contract
must record whether these cases are `ScalarError::Number` or another existing
formula-level category before the test oracle is frozen.

## Sequence admission, storage, and resource behavior

Root functions have a two-phase shape: one complete admission scan, followed
by bounded numerical passes over owned finite numbers.

During the scan, `IRR` and `XIRR` must follow the contract's source-order
sequence rules. Charge work and the reference-cell read before each physical
cell, check cancellation before and after the resolver operation, and retain a
formula error while continuing the admitted scan. A later typed
`Unsupported`, `ResourceLimit`, `Allocation`, `Cancelled`, or `SourceChanged`
failure aborts and supersedes a retained formula error. Formula errors are
never converted into typed provider failures, and typed failures are never
converted into formula errors.

After the scan:

- if a formula error was retained, finish the complete scan and return that
  formula error according to the shared precedence rules; do not start a root
  search that cannot change the result;
- if the sequence has no required sign variation, return the accepted numeric
  constraint error;
- reserve checked storage for the exact number of admitted values before
  filling the scratch buffer; and
- for `XIRR`, retain finite date offsets or date serials alongside the values,
  with no borrowed text or provider object in the buffer.

The root solver charges one work unit per residual term and one per derivative
term, plus one unit for each iteration or bracket probe. These are suggested
accounting units; the implementation should use the repository's
`ExecutionContext::consume(Resource::Work, ...)` and its cancellation check so
the parent budget remains authoritative. A root call may use at most 128
iterations and 256 residual/derivative evaluations even when the parent work
budget is larger. A work-limit failure is a typed resource failure. Reaching
the numerical iteration cap is a formula-level nonconvergence result.

`RATE` has no sequence scratch, but it uses the same evaluation, cancellation,
and hard-cap rules. `NPV`, `FVSCHEDULE`, and `XNPV` are one-pass operations and
must remain streaming; they must not materialize a range merely to reuse the
root implementation. Their scalar accumulator state is fixed-size or has an
explicit checked reservation. `CUMIPMT` and `CUMPRINC` charge each requested
period and each arithmetic step; their inclusive period range is checked for
integer overflow before the loop begins.

The value evaluator's source-version and publication fences surround the whole
operation. A cancellation request after the last term but before publication
still wins. A complete sequence descriptor may participate in a demand-cache
key, but a root result is reusable only when the full sequence values,
source/version identity, scalar guess, controls, and numeric profile are in the
key. A projected scalar guess or payment remains position-sensitive.

## Overflow and cancellation-sensitive cases

The implementation and oracle should include these cases before the profile is
accepted:

- zero rate, `+/-1e-12` rates, and rates near the representable value above
  `-1` for `FV`, `PV`, `PMT`, `NPER`, `EFFECT`, and `NOMINAL`;
- large positive `Nper` and rates that make `growth` overflow, plus a
  mathematically cancelling pair of large cash flows to verify the scaled
  residual path;
- subnormal rates and cash flows where direct `(1 + rate)^n - 1` loses all
  meaningful digits;
- negative rates for `NPER`, `RATE`, `IRR`, and `XIRR`, including a rate just
  above `-1`, an integral-period IRR/RATE root below `-1` selected by a guess
  below `-1`, the finite integral-`Nper` RATE boundary at exactly `-1`, and a
  candidate Newton step that would cross a branch boundary;
- IRR and XIRR with one sign only, no bracket from the default guess, two
  sign-changing roots, a repeated/tangent root, and an explicit alternate
  guess;
- XIRR with equal-size values/dates, unsorted dates where the contract allows
  them, and XNPV date-order violations where it does not;
- `FVSCHEDULE` factors of zero, negative factors, and products that overflow
  or underflow; and
- a late formula error, a late typed resolver failure, cancellation before a
  read, cancellation during a root pass, storage refusal while retaining a
  sequence, and source-version change before publication.

For every reference-backed root case, the expected read count is the admitted
cell count, not the count multiplied by root iterations. The expected work
count includes the initial scan and all bounded residual/derivative terms.
Tests should assert that the resolver is not reread by the numerical loop and
that a typed failure remains distinguishable from a generated `#NUM!`.

## Review questions before implementation evidence

The numerical reviewer and `inspection_kernel` should explicitly accept or
change these points:

1. Is the nearest sign-changing bracket to the supplied guess, with lower-rate
   tie-breaking across the positive/negative integral branches, the desired
   multiple-root policy, or should the profile choose a globally ordered root?
2. Are `u_max = ln1p(f64::MAX)`, `v_max = ln(f64::MAX)`, 128 iterations, 256
   evaluations, and the two `16*EPSILON` convergence tests sufficient for the
   expected financial range?
3. Should an unsupported negative-base guess (for example, `XIRR` below
   `-1`) be a numeric error, or should the evaluator use the default as a
   fallback? The latter would hide input data and is not recommended without a
   contract decision.
4. Should the `2048`-unit scaled-sum span refusal be a numeric formula error or
   a typed resource/profile refusal? The choice must not silently discard
   terms.
5. Does the accepted contract enforce XIRR's explanatory first-negative-cash-
   flow wording and XNPV's described sign pattern, or only their explicit
   count/type constraints?
6. Which existing formula error represents root nonconvergence, invalid
   rate-domain input, and overflow? The implementation and all oracle fixtures
   must use that mapping consistently.
7. Which Logical/Reference/ReferenceList forms are admitted by `MIRR(Array
   Values)` before its separate accumulator and resource tests are frozen?

Until these questions are answered, numerical outputs from a host spreadsheet
library are useful for exploration only. They are not a conformance oracle.
