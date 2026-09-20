# ODF 1.4 financial cash-flow and annuity contract

Status: draft implementation contract. This document defines the cash-flow
and annuity part of OpenFormula 1.4 Part 4 §6.12. It is a normative/design
boundary for a future implementation; it is not evidence that the functions
are currently implemented.

This slice contains the following twenty functions:

`CUMIPMT`, `CUMPRINC`, `EFFECT`, `FV`, `FVSCHEDULE`, `IPMT`, `IRR`, `ISPMT`,
`MIRR`, `NOMINAL`, `NPER`, `NPV`, `PDURATION`, `PMT`, `PPMT`, `PV`, `RATE`,
`RRI`, `XIRR`, and `XNPV`.

The other §6.12 functions are deliberately outside this contract. Security and
coupon functions (`ACCRINT`, `ACCRINTM`, `AMORLINC`, `COUPDAYBS`, `COUPDAYS`,
`COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`, `DISC`, `DURATION`, `INTRATE`,
`MDURATION`, `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `PRICE`,
`PRICEDISC`, `PRICEMAT`, `RECEIVED`, `TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`,
`YIELD`, `YIELDDISC`, and `YIELDMAT`) and depreciation/fraction functions
(`DB`, `DDB`, `DOLLARDE`, `DOLLARFR`, `SLN`, `SYD`, and `VDB`) require separate
contracts. This separation does not relax the requirement that the eventual
financial implementation preserve the shared §6.12 conventions.

## Authority

The primary source is the repository-local ODF 1.4 distribution:

| Source | SHA-256 |
| --- | --- |
| `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The direct financial definitions are §§6.12.1, 6.12.11–6.12.12,
6.12.19–6.12.21, 6.12.23–6.12.25, 6.12.27–6.12.30, 6.12.35–6.12.37,
6.12.41–6.12.42, 6.12.44, and 6.12.51–6.12.52. Their common type and
conversion rules are §§4.3.3–4.3.6, 4.11.1–4.11.2, 4.11.5, 4.11.7,
4.11.12–4.11.13, and 6.2–6.3. The existing date/time contract supplies the
repository's Date, DateSequence, serial, and Gregorian profile where XIRR or
XNPV consumes dates.

The archive has exactly 54 function entries in §§6.12.2–6.12.55. This slice
claims the following 20 entries: §§6.12.11–6.12.12 (`CUMIPMT`, `CUMPRINC`),
§§6.12.19–6.12.21 (`EFFECT`, `FV`, `FVSCHEDULE`), §§6.12.23–6.12.25
(`IPMT`, `IRR`, `ISPMT`), §§6.12.27–6.12.30 (`MIRR`, `NOMINAL`, `NPER`,
`NPV`), §§6.12.35–6.12.37 (`PDURATION`, `PMT`, `PPMT`), §§6.12.41–6.12.42
(`PV`, `RATE`), §6.12.44 (`RRI`), and §§6.12.51–6.12.52 (`XIRR`, `XNPV`).
The remaining 34 archive entries are outside this contract and are handled by
separate contracts:
§§6.12.2–6.12.10 (`ACCRINT`, `ACCRINTM`, `AMORLINC`, `COUPDAYBS`,
`COUPDAYS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`), §§6.12.13–6.12.18
(`DB`, `DDB`, `DISC`, `DOLLARDE`, `DOLLARFR`, `DURATION`), §6.12.22
(`INTRATE`), §6.12.26 (`MDURATION`), §§6.12.31–6.12.34 (`ODDFPRICE`,
`ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`), §§6.12.38–6.12.40 (`PRICE`,
`PRICEDISC`, `PRICEMAT`), §6.12.43 (`RECEIVED`), §6.12.45–6.12.46 (`SLN`,
`SYD`), §§6.12.47–6.12.49 (`TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`),
§6.12.50 (`VDB`), and §§6.12.53–6.12.55 (`YIELD`, `YIELDDISC`, `YIELDMAT`).
The seven depreciation/fraction entries in §§6.12.13–6.12.18,
§§6.12.45–6.12.46, and §6.12.50 are specified in the companion
[`depreciation-contract.md`](depreciation-contract.md). The inventory is
exhaustive for §6.12.2–§6.12.55; §6.12.1 is the shared General section and is
not a function entry.

The accepted repository constraints are part of this contract:

- ADR 0001 requires typed failures, no deliberate panics, and typed refusal
  for unsupported capability.
- ADR 0004 requires bounded semantic values and no archive or provider types in
  ordinary evaluator interfaces.
- ADR 0005 requires hierarchical work/storage/reference budgets, cancellation,
  source-version checks, fallible reservations, and no hidden I/O or clock.
- ADR 0006 keeps external providers explicit, preserves formula errors as
  values, and maps malformed or unsupported input to typed outcomes.
- ADR 0008 requires compile-tested positive and negative paths before a family
  is advertised as supported.

## §6.12.1 conventions

An annuity is a recurring series of payments at equal intervals with interest
compounded at those intervals. Payments at the end of an interval form an
ordinary annuity; payments at the beginning form an annuity due. Periods are
numbered from 1. The financial sign convention is normative: outgoing cash
flows are negative and incoming cash flows are positive.

`Currency` and `Percentage` are ODF subtypes of `Number`, not locale or display
objects. A financial result therefore remains a finite numeric value; cell
formatting is outside this evaluator.

`Integer` is a Number with no fractional value. The generic ODF conversion from
a non-integer Number to Integer is implementation-defined unless a function
specifies a rounding operation. This repository profile uses the existing
scalar conversion policy, truncation toward zero, for the `Integer` parameters
in this slice. A future implementation must keep that choice in this contract
and test negative fractional inputs; it must not rely on an unchecked float to
integer cast.

`Basis` is an Integer day-count selector (§4.11.7):

| Basis | Convention | Day-count procedure | Days-in-year procedure |
| ---: | --- | --- | --- |
| 0 | US/NASD 30/360 | A | D |
| 1 | Actual/Actual | B | E |
| 2 | Actual/360 | B | D |
| 3 | Actual/365 | B | F |
| 4 | European 30/360 | C | D |

The functions in this slice do not take `Basis`, but XIRR and XNPV use the
same date serial representation as the completed date/time family. The
security functions outside this slice will use the table directly.

## Exact signatures and explicit constraints

The following table transcribes the Part 4 signatures. Square brackets denote
the optional parameters and defaults printed by §6.2. A constraint column says
`None` when the source explicitly says there is no additional constraint; the
implementation must not silently add a positivity or monotonicity rule that is
absent from the cited section.

| Section/function | Signature | Returns | Explicit constraints and defaults |
| --- | --- | --- | --- |
| §6.12.11 `CUMIPMT` | `CUMIPMT(Number Rate; Number Periods; Number Value; Integer Start; Integer End; Integer Type)` | Currency | `Rate > 0`, `Value > 0`, and `1 <= Start <= End <= Periods`. `Type` is 0 (payment at end) or 1 (payment at beginning). |
| §6.12.12 `CUMPRINC` | `CUMPRINC(Number Rate; Number Periods; Number Value; Integer Start; Integer End; Integer Type)` | Currency | The source states the `Type` table: 0 (payment at end) or 1 (payment at beginning). It does not repeat the outer `CUMIPMT` constraints. Each included `PPMT` term applies the `PPMT` preconditions; an empty reversed integer interval sums to zero after `Type` validation. |
| §6.12.19 `EFFECT` | `EFFECT(Number Rate; Integer Payments)` | Number | `Rate >= 0`; `Payments > 0`. |
| §6.12.20 `FV` | `FV(Number Rate; Number Nper; Number Payment[; [Number Pv = 0][; Number PayType = 0]])` | Currency | No additional constraint. `PayType` 0 means payments at the end; 1 means payments at the beginning. |
| §6.12.21 `FVSCHEDULE` | `FVSCHEDULE(Number Principal; NumberSequence Schedule)` | Currency | None. The schedule is an ordered NumberSequence. |
| §6.12.23 `IPMT` | `IPMT(Number Rate; Number Period; Number Nper; Number PV[; Number FV = 0[; Number Type = 0]])` | Currency | No additional constraint. `Type` 0 means payments at the end; 1 means payments at the beginning. |
| §6.12.24 `IRR` | `IRR(NumberSequence Values[; Number Guess = 0.1])` | Percentage | None. `Guess` is the initial estimate only; failure to converge may return an Error. |
| §6.12.25 `ISPMT` | `ISPMT(Number Rate; Number Period; Number Nper; Number Pv)` | Currency | None. |
| §6.12.27 `MIRR` | `MIRR(Array Values; Number Investment; Number ReinvestRate)` | Percentage | `Values` contains at least one positive and at least one negative value. Text and Empty cells are ignored. |
| §6.12.28 `NOMINAL` | `NOMINAL(Number EffectiveRate; Integer CompoundingPeriods)` | Number | `EffectiveRate > 0`; `CompoundingPeriods > 0`. |
| §6.12.29 `NPER` | `NPER(Number Rate; Number Payment; Number Pv[; [Number Fv = 0][; Number PayType = 0]])` | Number | No additional constraint. `PayType` 0 means payments at the end; 1 means payments at the beginning. The source expressly requires negative rates for Medium/Large evaluators; this profile targets that requirement. |
| §6.12.30 `NPV` | `NPV(Number Rate; {NumberSequenceList Values}+)` | Currency | At least one `NumberSequenceList` argument. Values are evaluated in argument order and ranges/arrays row-wise from top left. No additional rate or sign constraint is printed. |
| §6.12.35 `PDURATION` | `PDURATION(Number Rate; Number CurrentValue; Number SpecifiedValue)` | Number | `Rate > 0`; `CurrentValue > 0`; `SpecifiedValue > 0`. |
| §6.12.36 `PMT` | `PMT(Number Rate; Integer Nper; Number Pv[; [Number Fv = 0][; Number PayType = 0]])` | Currency | `Nper > 0`. `PayType` 0 means payments at the end; 1 means payments at the beginning. |
| §6.12.37 `PPMT` | `PPMT(Number Rate; Integer Period; Integer Nper; Number Present[; Number Future = 0[; Number Type = 0]])` | Number | `Rate > 0`, `Present > 0`, and `0 < Period < Nper`. `Future` defaults to 0; `Type` 0/1 selects end/beginning payment. |
| §6.12.41 `PV` | `PV(Number Rate; Number Nper; Number Payment[; [Number Fv = 0][; Number PayType = 0]])` | Currency | No additional constraint. `PayType` 0 means payments at the end; 1 means payments at the beginning. |
| §6.12.42 `RATE` | `RATE(Number Nper; Number Payment; Number Pv[; [Number Fv = 0][; [Number PayType = 0][; Number Guess = 0.1]]])` | Percentage | `Nper > 0`; `Fv` defaults to 0, `PayType` to 0, and `Guess` to 0.1. Failure to converge may return an Error. |
| §6.12.44 `RRI` | `RRI(Number Nper; Number Pv; Number Fv)` | Percentage | `Nper > 0`. |
| §6.12.51 `XIRR` | `XIRR(NumberSequence Values; DateSequence Dates[; Number Guess = 0.1])` | Number | Values and Dates have equal size; the first admitted cash flow is negative; and at least one admitted cash flow is positive. Violations return `#NUM!`. `Guess` defaults to 0.1. |
| §6.12.52 `XNPV` | `XNPV(Number Rate; Reference | Array Values; Reference | Array Dates)` | Number | Values and Dates have equal element counts, including when their rectangular geometries differ; every element is Number; every date is at least the first date; and `Rate > -1` as stated in the semantics. The first admitted cash flow is negative and at least one admitted cash flow is positive; violations return `#NUM!`. |

The nested optional syntax is significant. An omitted slot uses the listed
default; an explicit missing slot is still a supplied argument and follows the
existing missing-argument conversion/error policy. `Type` and `PayType` are
numeric 0/1 controls, not arbitrary truthiness flags. A Number-declared
`Type` or `PayType` is accepted only when its converted value is exactly 0 or
1. A control slot declared `Integer` first uses the repository's
truncation-toward-zero Integer profile and is then checked for exact 0 or 1;
only declared Integer slots truncate. In this slice that includes `Type` in
`CUMIPMT` and `CUMPRINC`; `PPMT`'s `Type` and the `Type`/`PayType` slots declared Number
in the other annuity signatures use exact comparison without truncation. No
control accepts an arbitrary nonzero Number or Boolean-style coercion.

## Required mathematical semantics

The equations in the cited ODF sections define the result. The implementation
may use a numerically stable algorithm, but it may not substitute an Excel
variant, a host-specific convention, or an uncited approximation.

- `CUMIPMT` is the sum of `IPMT(Rate; p; Periods; Value; 0; Type)` for
  `p = Start..End`, after its printed outer constraints pass. `CUMPRINC` is
  the corresponding sum of `PPMT(Rate; p; Periods; Value; 0; Type)`. For
  `CUMPRINC`, validate `Type` even when `Start > End`; that reversed integer
  interval has an empty sum of zero. For each included period, apply the
  `PPMT` preconditions (`Rate > 0`, `Present > 0`, and `0 < p < Nper`), so an
  included period equal to `Nper` produces `#NUM!`. Do not copy the outer
  `CUMIPMT` constraints onto `CUMPRINC`.
- `EFFECT` is `(1 + Rate / Payments)^Payments - 1`. `NOMINAL` is the inverse
  relation printed in §6.12.28: `Effective = (1 + Nominal / m)^m - 1`.
- `FVSCHEDULE` is `Principal * product(1 + Schedule[i])`, in sequence order.
- `NPV` is `sum(Value[i] / (1 + Rate)^i)` with i beginning at 1. The source's
  explicit order rule is part of the result, including when values are split
  across several sequence-list arguments.
- `PDURATION` is
  `(log(SpecifiedValue) - log(CurrentValue)) / log(Rate + 1)`.
- `RRI` is `(Fv / Pv)^(1 / Nper) - 1`. For a negative ratio, this
  repository selects the real `POWER` profile permitted by §6.16.46: the
  binary64 reciprocal `1 / Nper` must be exactly integral; otherwise return
  `#NUM!`. Evaluate an admitted negative base with checked signed integer
  power. Thus `RRI(0.5;100;-100)` is zero, while `RRI(3;100;-800)` is
  `#NUM!`; there is no additional odd-root extension.
- `XNPV` is
  `sum(Values[i] / (1 + Rate)^((Dates[i] - Dates[1]) / 365))`.
- `FV`, `NPER`, `PMT`, and `PV` use the equations printed in §§6.12.20,
  6.12.29, 6.12.36, and 6.12.41. Their zero-rate branches are explicit in
  the source. The common nonzero-rate expression is the balance equation
  containing `(1 + Rate)^Nper`, the `PayType` factor
  `(1 + Rate * PayType)`, and the annuity factor
  `((1 + Rate)^Nper - 1) / Rate`; do not simplify it through a division by
  `Rate` when Rate is zero.
- `RATE` solves the same balance equation for Rate using the source's
  `Guess`; `IRR` solves the rate for which `NPV` is zero; `XIRR` solves the
  corresponding date-weighted `XNPV` equation. The source explicitly permits
  an approximate iterative result and an Error on nonconvergence.
- `IPMT` and `PPMT` use the following explicit amortization profile because
  their ODF clauses provide no equation. For a nonzero rate, define the
  balance factor `A = (1 + Rate * Type) * ((1 + Rate)^Nper - 1) / Rate` and
  the payment `Payment = -(PV * (1 + Rate)^Nper + FV) / A`. Define the balance
  factor `B` for period `Period` as
  `PV * (1 + Rate)^(Period - 1) + Payment * (1 + Rate * Type) *
  ((1 + Rate)^(Period - 1) - 1) / Rate`. For `Type = 0`,
  `IPMT = -B * Rate`. For `Type = 1`, the first period's interest is zero;
  in subsequent periods `IPMT = -B * Rate / (1 + Rate)`, which charges
  interest on the balance after the preceding beginning-of-period payment.
  Thus `IPMT(0.1;2;2;100;0;1) = -100/21`, not `-110/21`.
  In either case `PPMT = Payment - IPMT`. The zero-rate
  payment branch is `-(PV + FV) / Nper` and its interest component is zero
  where the function's domain admits zero. This is a documented repository
  profile, not an assertion that the ODF prose prints these equations.
- `ISPMT` uses the documented conventional profile
  `ISPMT = Rate * Pv * (Period / Nper - 1)` because §6.12.25 supplies no
  equation. Its source signature and any source constraints still govern
  argument admission.
- `MIRR` uses the §6.12.27 MathML equation:
  `((-NPV(ReinvestRate; PositiveMask) * (1 + ReinvestRate)^n) /
  (NPV(Investment; NegativeMask) * (1 + Investment)))^(1/(n - 1)) - 1`.
  The source writes the masks as `Values > 0` and `Values < 0`; their
  positional interpretation is fixed below. This equation must not be
  replaced with a different host modified-IRR formula.
  Under this profile, first remove Text and Empty elements, retain Logical
  elements as 0/1, and use the retained count `n`. Build the positive and
  negative NPV masks over the same retained positions, contributing zero for
  the opposite sign at each exponent; do not compact the two signs into
  separate period sequences.

The ODF source does not prescribe a convergence tolerance, iteration count,
root-selection method, or treatment of multiple valid roots. Those are
implementation-profile choices and must be stated with deterministic bounded
behavior before implementation evidence is accepted.

## Conversion and sequence profile

The expected pseudotype controls conversion before the function operates:

- Number arguments use the existing evaluator Number bridge: finite Number is
  retained, Logical maps to 0/1, and Text uses the repository's fixed finite
  numeric-text profile. Malformed or non-finite Text produces the existing
  formula error. Complex values are not silently projected to real Number.
- Integer arguments first use Number conversion and then the repository's
  documented truncation-toward-zero profile. Non-finite or out-of-domain
  integer results produce the existing numeric formula error.
- `NumberSequence` accepts scalar Number/Text/Logical as a one-element
  sequence. A single Reference is streamed in source order and admits Number
  and formula Error cells; Empty, Text, and distinguished Logical cells are
  omitted under §6.3.7.
- `NumberSequenceList` has the same element rules and additionally accepts an
  ordered ReferenceList, processing each area in occurrence order. This is why
  NPV can receive one or more sequence-list arguments while IRR and FVSCHEDULE
  use a single NumberSequence.
- `DateSequence` follows the existing date/time profile and §6.3.9. Scalar
  Number, Text, and Logical first use Number conversion and form a one-element
  Date sequence; scalar Text here uses numeric conversion, not DateParam text
  parsing. A single Reference admits Number and formula Error cells, skipping
  Empty, Text, and distinguished Logical cells. Every admitted numeric element
  must be a profile Date serial. The profile does not invent ReferenceList
  flattening for DateSequence. A known ReferenceList mismatch is rejected
  before resolver reads; a computed expression may need to run until its type
  is known.
- `XNPV` names `Reference | Array` directly rather than a sequence pseudotype.
  A rectangular Reference or inline Array is traversed row-wise from the top
  left. A ReferenceList is not silently flattened. Values and Dates are paired
  by row-major element position even when their rectangular geometries differ;
  only their element counts must match. Every non-error admitted element must
  be Number. Empty, Text, and Logical produce generated `#VALUE!` rather than
  becoming a numeric cash flow, while a formula Error remains an error value
  under the common error precedence below and is never coerced to a number.
- `XIRR` uses the complete `NumberSequence` and `DateSequence` in source
  order. Text, Empty, and distinguished Logical cells follow the existing
  sequence admission rules. This profile requires equal original flattened
  slot counts and retains each slot's original index through admission.
  Pair only matching original indices: when both sides skip a slot, drop that
  pair; when exactly one side skips it, retain a generated `#VALUE!` and
  complete both scans. Do not independently compact the arguments and assign
  a later date to an earlier cash flow. Formula Errors occupy admitted slots
  and keep their precedence over a generated mask-mismatch error. Scalar
  sequence inputs have synthetic slot index zero. Different rectangular
  geometries are allowed when the original flattened counts and admission
  masks match. This strict positional rule is a documented profile of the
  source's requirement that dates correspond to values. The first admitted numeric cash flow
  must be negative and at least one admitted numeric cash flow must be
  positive; otherwise the result is `#NUM!`.
- `MIRR` names `Array`, and §6.12.27 explicitly ignores Text and Empty cells.
  In scalar mode, a plain direct Reference or ReferenceList is refused as the
  wrong shape with zero resolver reads; it is admitted only when matrix
  evaluation has already produced an Array (or a computed child explicitly
  produces an Array). Logical elements use the explicit profile conversion to
  0/1. Text and Empty elements are compacted away before the sign masks and do
  not consume a period. Formula Errors remain retained errors and the admitted
  scan continues for typed-failure precedence. For each retained position,
  positive and negative masks share the same position; the opposite mask
  contributes zero at that same NPV exponent. Positive and negative values
  must not be compacted into separate sequences, and `n` is the retained
  count after Text/Empty removal.

In scalar evaluation, a reference or array that is not consumed by one of
these complete-sequence or explicit-array pseudotypes follows the existing
implicit-intersection rules. In value/matrix evaluation, a complete sequence
argument is consumed in full for each scalar result; scalar rate, payment,
guess, and control arguments retain their projected position semantics. A
function returning a scalar is not turned into a result matrix merely because
its sequence input is rectangular.

## Formula errors, typed failures, and source order

Formula errors are values. The evaluator observes arguments in source order and
walks each admitted sequence in its prescribed order. It retains the first
formula Error in conceptual order while continuing the complete admitted scan
when the operation requires it. After typed failures, formula errors outrank
generated formula errors; within each class the first error in source/traversal
order wins. Generated errors from malformed conversion, shape, domain, or
numerical failure are therefore recorded while a required complete scan
continues. A formula Error does not become a typed provider failure and cannot
be hidden by an eager `IFERROR` conversion.

Resolver `Unsupported`, `ResourceLimit`, `Allocation`, `Cancelled`,
`SourceChanged`, and source-version availability failures abort immediately and
remain typed evaluator failures. They supersede any retained formula Error,
including when the failure occurs after an earlier Error cell. `IFERROR` and
`IFNA` may handle formula-error values only; they do not catch these typed
failures.

Before reading a known wrong pseudotype, the value evaluator performs the
shape/type gate and does zero resolver reads. This applies to a ReferenceList
where `NumberSequence` or `DateSequence` admits only a single Reference and to
an invalid explicit XNPV/MIRR shape. A computed child may be evaluated to
establish its resulting type, as in the existing value evaluator contract.

The generated-error profile is fixed for this family. A malformed arity,
pseudotype, non-Number XNPV element, or refused shape is `#VALUE!`; a
numeric/domain/sign/convergence violation is `#NUM!`; and a direct zero
divisor such as the zero denominator in an otherwise admitted arithmetic
branch is `#DIV/0!`, following the existing arithmetic convention. XIRR and
XNPV sign violations are `#NUM!`. An equal-element-count failure for XIRR or
XNPV is a shape failure and is `#VALUE!`. Typed resolver, allocation,
cancellation, and source-version failures never enter this formula-error
mapping.

## Numerical and bounded-resource requirements

The result is finite binary64 under the repository's existing Number profile.
NaN and infinities do not escape. Constraint violations and a non-finite
mathematical result use the numeric formula-error mapping below; malformed
type/arity/shape conversions use the value mapping below. There is no blanket
`Rate > -1` gate for functions whose source has no such constraint. In
particular, `NPV` and other non-iterative rate formulas admit a finite rate
below -1 when their real arithmetic is finite; an exactly zero denominator at
`Rate = -1` is `#DIV/0!`. Source-specific domains such as `PDURATION`'s
positive Rate and XNPV's `Rate > -1` still apply. This contract does not claim
a universal ulp tolerance: every future numeric oracle must state its
precision and rounding profile.

`IRR`, `RATE`, and `XIRR` have no closed form in the specification. Their
accepted bounded solver profile is documented below. It has a checked finite
work bound, periodic cancellation checks, finite intermediate/result checks,
and deterministic convergence criteria. A caller-provided guess cannot
increase the hard iteration or work limit, and the profile makes no universal
root-existence or globally smallest-root promise.

The solver uses a transformed coordinate on each real branch. For a positive
base, let `u = ln1p(rate)` and recover `rate = expm1(u)`. The branch bounds are
`u_min = ln1p(r_min)`, where `r_min` is the next representable `f64` above
`-1`, and `u_max = ln1p(f64::MAX)`. For an integral-period negative base, let
`v = ln(-(1 + rate))` and recover `rate = -1 - exp(v)`. Its bounds are
`v_min = ln(-(1 + nextbelow(-1)))` and `v_max = ln(f64::MAX)`. These bounds
cover the corresponding finite binary64 rates; `u_max` is not an arbitrary
small cap. `XIRR` always uses the positive-base branch because its date
exponents can be fractional. `IRR` may use the negative-base branch, and
`RATE` may use it only when `Nper` is positive and integral. A guess below
`-1` selects that branch only for those functions; an unsupported branch or an
`IRR`/`XIRR` guess exactly at `-1` returns `#NUM!`. A `RATE` guess exactly at
`-1` uses the guarded boundary rule below.

At each residual coordinate, the solver computes the signed residual and a
fixed-width absolute scale equal to the sum of the absolute term magnitudes.
The residual condition is
`abs(residual) <= 16 * f64::EPSILON * scale`, with no `max(1, scale)` floor.
The transformed-coordinate step or bracket-width condition is
`16 * f64::EPSILON * max(1, abs(coordinate))`. A zero scale fails the required
sign-variation precondition. A fixed-width scaled sum that would exceed the
2048-natural-log term span returns `#NUM!`; it never drops a term that might
become visible after cancellation.

The solver probes at most 64 outward bracket radii (`1, 2, 4, ...`) from the
guess, clipped to the active branch bounds. It checks endpoints in increasing
distance from the guess and chooses the first sign-changing bracket. Equally
near brackets are resolved by choosing the lower resulting rate; an exactly
zero endpoint is returned immediately. It then uses
a safeguarded Newton/secant/bisection hybrid: Newton or secant candidates must
be finite and strictly inside the bracket, otherwise the step bisects. The
root-iteration limit is 128, and initial residuals, bracket probes, derivative
evaluations, and solve evaluations share a 256-evaluation cap. The cap is
authoritative if reached first. Every residual/derivative evaluation checks
cancellation before and after it. A root is accepted only when both residual
and step/bracket conditions pass; a tangent root without a sign-changing
bracket is reported as `#NUM!`. No-root, invalid transformed domain,
non-finite residual, scaled-span refusal, and bounded nonconvergence all map to
`#NUM!`.

For `IRR`, the negative-base residual uses
`sum(cashflow_i * (-1)^i * exp(-i * v))`, with `i` beginning at one. For
integral-`Nper` `RATE`, the balance equation uses checked signed integer powers
on that branch. The positive branch uses the corresponding `exp(-i * u)` or
balance equation. Specifically, the periodic residual is
`sum(cashflow_i * exp(-i * u))`, and XIRR's residual replaces `i` with
`(date_i - date_0) / 365`. Derivatives use the same scaled signed sums and
their coordinate derivative; a non-finite or unusable derivative falls back
to the bounded secant/bisection path. The selected bracket is the deterministic
root policy for multiple roots; the solver does not claim to discover every
root.

At the `RATE` boundary `rate = -1` with positive integral `Nper`, evaluate the
balance directly: `(1 + rate)^Nper = 0`, the annuity factor is `1`, and the due
factor is `1 - PayType`. Return `-1` when that guarded residual satisfies the
same relative scale test. Otherwise continue with adjacent positive and
negative integral branches in deterministic distance order, with the negative
branch winning an equal-distance tie, or return bounded `#NUM!`. A non-integral
`Nper` at `rate = -1` follows the numeric error mapping. These solver rules do
not impose `Rate > -1` on non-iterative functions whose ODF definitions have
no such constraint. Conversion from `u` or `v` must yield a finite binary64
rate; an infinite result is `#NUM!` and is never clipped to `f64::MAX`.

`NPV`, `IRR`, `FVSCHEDULE`, `XIRR`, and `XNPV` consume sequence/reference
arguments in order. They must charge work and reference reads before each
physical cell, check cancellation before and after resolver calls, and retain
borrowed text until a conversion decision without cloning it unnecessarily.
`XNPV` scans the complete Values argument first, in row-major source order,
into bounded numeric slots and error/invalid markers, then scans Dates fully
and streams the row-major date side against those slots. It pairs by element
index even when the two rectangular geometries differ. `XIRR` follows the
same Values-then-Dates source order and retains only bounded numeric pairs (or
scalar error/invalid markers) needed by its solver. MIRR's admitted Array is
compacted according to its Text/Empty rule, then retains bounded numeric
values and shared-position sign masks; it does not retain cell objects or
borrowed text. Sequence metadata and any numeric scratch needed for repeated
root evaluation are bounded by the existing reference-cell, array-cell, work,
and storage limits. If an algorithm retains numeric terms for repeated passes,
it reserves the exact checked capacity before allocation, stores no borrowed
cell objects or text, and drops the buffer before releasing its associated
reservation. A one-pass operation must remain streaming and must not
materialize a range merely for convenience.

The root solver charges one work unit per residual term and one per derivative
term, plus one unit for each iteration or bracket probe, through the existing
work budget and cancellation checks. A root call may use at most 128 solve
iterations and 256 residual/derivative evaluations, including initial
residuals and bracket probes. The numerical cap is authoritative if reached
first; a parent work-budget exhaustion remains a typed resource failure, while
reaching the numerical cap is `#NUM!` nonconvergence.

The value evaluator keeps the source-version and cancellation fences around the
whole evaluation and before result publication. A typed read or cancellation
failure is never converted into a formula Error. A demand-cache entry may be
used only after complete sequence descriptors/values and all scalar parameters
are part of the key; projected scalar parameters remain position-sensitive.
Computed sequence expressions are cacheable only when their complete descriptor
and source identity are stable under the existing value-cache rules.

No financial function consults an ambient clock, locale, filesystem, network,
random generator, workbook recalculation service, or external rate provider.

## Validation requirements and profile disposition

Before implementation is advertised, focused validation must cover:

- exact arity, every optional omission/default, explicit missing slots, and
  wrong pseudotypes for all twenty functions;
- every explicit constraint in the signature table, including Type/PayType
  exact 0/1 controls, integer truncation only for declared Integer slots,
  zero-rate branches, negative-rate NPER profile, sequence sign requirements,
  XIRR size/sign rules, and XNPV count/date/sign rules;
- source-order NPV values across multiple ReferenceList areas, row-major Array
  order, formula Error retention, and late typed provider/resource/cancel/source
  failures; XNPV Values-first then Dates scans with equal counts across
  different geometries; and non-Number `#VALUE!` cases;
- scalar, inline-array, single-Reference, 3-D Reference, and projected matrix
  behavior for each sequence-bearing argument, including zero-read refusal for
  known ReferenceList/date-sequence mismatches; direct scalar MIRR
  Reference/ReferenceList refusal; and matrix-produced Array admission;
- exact cancellation and overflow-sensitive cases for NPV, FVSCHEDULE, MIRR,
  XIRR, XNPV, and annuity zero/nonzero-rate branches, including MIRR's
  compacted Text/Empty positions and shared sign-mask exponents;
- CUMPRINC reversed-empty ranges, Type validation on empty ranges, and
  per-term PPMT preconditions including `End = Nper`;
- deterministic IRR/RATE/XIRR convergence, no-root, multiple-root, negative
  rate, guess, iteration-limit, and cancellation cases; and
- reference-cell, work, storage, allocation, and cancellation limits with no
  partial cache publication.

The bounded transformed-coordinate solver profile above resolves the prior
numerical-policy questions: branch bounds, negative-base eligibility,
bracketing order, tie-breaking, residual scale, term-span limit, iteration and
evaluation caps, boundary handling, multiple-root choice, and `#NUM!` failure
mapping are fixed. It makes no universal root-existence guarantee. The
contract still claims no production support or validation disposition until the
implementation, focused tests, resource evidence, and source freeze are
complete.
