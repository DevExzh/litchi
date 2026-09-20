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
The remaining 34 archive entries are reserved for later contracts:
§§6.12.2–6.12.10 (`ACCRINT`, `ACCRINTM`, `AMORLINC`, `COUPDAYBS`,
`COUPDAYS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`), §§6.12.13–6.12.18
(`DB`, `DDB`, `DISC`, `DOLLARDE`, `DOLLARFR`, `DURATION`), §6.12.22
(`INTRATE`), §6.12.26 (`MDURATION`), §§6.12.31–6.12.34 (`ODDFPRICE`,
`ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`), §§6.12.38–6.12.40 (`PRICE`,
`PRICEDISC`, `PRICEMAT`), §6.12.43 (`RECEIVED`), §6.12.45–6.12.46 (`SLN`,
`SYD`), §§6.12.47–6.12.49 (`TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`),
§6.12.50 (`VDB`), and §§6.12.53–6.12.55 (`YIELD`, `YIELDDISC`, `YIELDMAT`).
The inventory is exhaustive for §6.12.2–§6.12.55; §6.12.1 is the shared
General section and is not a function entry.

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
| §6.12.12 `CUMPRINC` | `CUMPRINC(Number Rate; Number Periods; Number Value; Integer Start; Integer End; Integer Type)` | Currency | The source states the `Type` table: 0 (payment at end) or 1 (payment at beginning). It does not repeat the `Rate`, `Value`, or period constraints from `CUMIPMT`; this profile does not invent them. |
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
| §6.12.51 `XIRR` | `XIRR(NumberSequence Values; DateSequence Dates[; Number Guess = 0.1])` | Number | Values and Dates have equal size; Values has at least one positive and one negative cash flow. `Guess` defaults to 0.1; failure to converge may return an Error. |
| §6.12.52 `XNPV` | `XNPV(Number Rate; Reference | Array Values; Reference | Array Dates)` | Number | Values and Dates have equal element counts; every element is Number; every date is at least the first date; and `Rate > -1` as stated in the semantics. The semantics describes a negative initial investment and a positive cash flow; whether that sign pattern is enforced remains open below. |

The nested optional syntax is significant. An omitted slot uses the listed
default; an explicit missing slot is still a supplied argument and follows the
existing missing-argument conversion/error policy. `Type` and `PayType` are
numeric 0/1 controls, not arbitrary truthiness flags.

## Required mathematical semantics

The equations in the cited ODF sections define the result. The implementation
may use a numerically stable algorithm, but it may not substitute an Excel
variant, a host-specific convention, or an uncited approximation.

- `CUMIPMT` is the sum of `IPMT(Rate; p; Periods; Value; 0; Type)` for
  `p = Start..End`. `CUMPRINC` is the corresponding sum of
  `PPMT(Rate; p; Periods; Value; 0; Type)`.
- `EFFECT` is `(1 + Rate / Payments)^Payments - 1`. `NOMINAL` is the inverse
  relation printed in §6.12.28: `Effective = (1 + Nominal / m)^m - 1`.
- `FVSCHEDULE` is `Principal * product(1 + Schedule[i])`, in sequence order.
- `NPV` is `sum(Value[i] / (1 + Rate)^i)` with i beginning at 1. The source's
  explicit order rule is part of the result, including when values are split
  across several sequence-list arguments.
- `PDURATION` is
  `(log(SpecifiedValue) - log(CurrentValue)) / log(Rate + 1)`.
- `RRI` is `(Fv / Pv)^(1 / Nper) - 1`.
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
- `MIRR` uses the §6.12.27 equation over positive and negative Values. The
  source HTML renders the fraction and exponent with MathML; the MathML,
  rather than flattened text extraction, is authoritative. A contract or
  implementation must not replace it with a different modified-IRR formula.

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
- `DateSequence` follows the existing date/time profile and §6.3.9: a single
  Reference or scalar date-number sequence is admitted, with each admitted
  element a Date serial. The current profile does not invent ReferenceList
  flattening for DateSequence. A known ReferenceList mismatch is rejected
  before resolver reads; a computed expression may need to run until its type
  is known.
- `XNPV` names `Reference | Array` directly rather than a sequence pseudotype.
  A rectangular Reference or inline Array is traversed row-wise from the top
  left. A ReferenceList is not silently flattened. Every admitted element must
  be Number; Empty, Text, and Logical do not become a numeric cash flow. A
  formula Error remains an error value under the common error precedence below;
  it is never coerced to a numeric cash flow.
- `MIRR` names `Array`, and §6.12.27 explicitly ignores Text and Empty cells.
  The specification does not define a separate Logical conversion or a
  ReferenceList-to-Array conversion for this parameter. The implementation
  contract must choose and test that boundary before claiming resolver-backed
  MIRR support; it must not silently apply NumberSequenceList rules.

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
when the operation requires it. Generated formula errors from malformed
conversion, shape, domain, or numerical failure are retained according to the
same existing reducer precedence. A formula Error does not become a typed
provider failure and cannot be hidden by an eager `IFERROR` conversion.

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

## Numerical and bounded-resource requirements

The result is finite binary64 under the repository's existing Number profile.
NaN and infinities do not escape. Constraint violations and a non-finite
mathematical result use the evaluator's established numeric formula-error
mapping; malformed type/arity/shape conversions use its established value
mapping. This contract does not claim a universal ulp tolerance: every future
numeric oracle must state its precision and rounding profile.

`IRR`, `RATE`, and `XIRR` have no closed form in the specification. Their
iteration must have a checked finite work bound, periodic cancellation checks,
finite intermediate/result checks, deterministic convergence criteria, and a
documented result when a root is absent or the requested guess does not
converge. A caller-provided guess cannot increase the hard iteration or work
limit. Multiple roots and an iteration that crosses `Rate <= -1` are open
profile decisions; they must not be resolved by an undocumented host-library
fallback.

`NPV`, `IRR`, `FVSCHEDULE`, and `XIRR` consume sequence references in order.
They must charge work and reference reads before each physical cell, check
cancellation before and after resolver calls, and retain borrowed text until a
conversion decision without cloning it unnecessarily. Sequence metadata and
any numeric scratch needed for repeated root evaluation are bounded by the
existing reference-cell, array-cell, work, and storage limits. If an algorithm
retains numeric terms for repeated passes, it reserves the exact checked
capacity before allocation, stores no borrowed cell objects or text, and drops
the buffer before releasing its associated reservation. A one-pass operation
must remain streaming and must not materialize a range merely for convenience.

The value evaluator keeps the source-version and cancellation fences around the
whole evaluation and before result publication. A typed read or cancellation
failure is never converted into a formula Error. A demand-cache entry may be
used only after complete sequence descriptors/values and all scalar parameters
are part of the key; projected scalar parameters remain position-sensitive.
Computed sequence expressions are cacheable only when their complete descriptor
and source identity are stable under the existing value-cache rules.

No financial function consults an ambient clock, locale, filesystem, network,
random generator, workbook recalculation service, or external rate provider.

## Validation requirements and open profile decisions

Before implementation is advertised, focused validation must cover:

- exact arity, every optional omission/default, explicit missing slots, and
  wrong pseudotypes for all twenty functions;
- every explicit constraint in the signature table, including Type/PayType
  values, integer truncation, zero-rate branches, negative-rate NPER profile,
  sequence sign requirements, XIRR size equality, and XNPV date ordering;
- source-order NPV values across multiple ReferenceList areas, row-major Array
  order, formula Error retention, and late typed provider/resource/cancel/source
  failures;
- scalar, inline-array, single-Reference, 3-D Reference, and projected matrix
  behavior for each sequence-bearing argument, including zero-read refusal for
  known ReferenceList/date-sequence mismatches;
- exact cancellation and overflow-sensitive cases for NPV, FVSCHEDULE, MIRR,
  and annuity zero/nonzero-rate branches;
- deterministic IRR/RATE/XIRR convergence, no-root, multiple-root, negative
  rate, guess, iteration-limit, and cancellation cases; and
- reference-cell, work, storage, allocation, and cancellation limits with no
  partial cache publication.

The following choices remain intentionally open pending semantic and numerical
review; they are not permission to guess during implementation:

1. The precise `MIRR(Array Values)` admission of a rectangular Reference,
   inline Array, Logical element, and ReferenceList must be pinned against the
   general array-context rules. The explicit Text/Empty omission is normative;
   other element conversions are not stated by §6.12.27.
2. `CUMPRINC` omits the positivity and period constraints printed for
   `CUMIPMT`; the implementation must preserve that textual distinction or
   obtain an accepted cross-reference decision.
3. Root selection, convergence tolerance, bounded iteration count, and error
   kind for `IRR`, `RATE`, and `XIRR` require an explicit numerical profile.
4. The source says XIRR's first cash flow represents the investment and also
   separately requires only one positive and one negative value; whether the
   first sign is a hard constraint or explanatory convention must be decided
   before tests assert it.
5. XNPV's constraints say every element is Number and dates are ordered, while
   its semantics also describes a negative initial investment and a positive
   later cash flow. The implementation profile must state whether the sign
   requirement is enforced as a constraint or treated as explanatory text.
6. The exact error categories for a violated financial constraint are not
   assigned by the individual §6.12 clauses. The repository-wide mapping must
   be recorded in the implementation contract and applied consistently.

These open points are deliberately visible so a later semantic review can
resolve them without silently importing behavior from another spreadsheet
host. The contract claims no production support or validation disposition until
those decisions, implementation, focused tests, resource evidence, and source
freeze are complete.
