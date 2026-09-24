# ODF 1.4 depreciation and fractional monetary contract

Status: draft implementation contract. This document defines the seven
depreciation and fractional monetary functions in OpenFormula 1.4 Part 4
§6.12. It is a normative/design boundary for a future implementation; it is
not evidence that these functions are currently implemented.

The slice contains `DB`, `DDB`, `DOLLARDE`, `DOLLARFR`, `SLN`, `SYD`, and
`VDB`. The twenty cash-flow and annuity functions have their own
[`contract.md`](contract.md). Together these contracts account for the 27
financial entries not assigned to the security/coupon family in the broader
54-function §6.12 inventory.

## Authority and section scope

The primary source is the repository-local ODF 1.4 distribution:

| Source | SHA-256 |
| --- | --- |
| `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The direct definitions are §6.12.13 `DB`, §6.12.14 `DDB`, §6.12.16
`DOLLARDE`, §6.12.17 `DOLLARFR`, §6.12.45 `SLN`, §6.12.46 `SYD`, and
§6.12.50 `VDB`. `TRUNC` in §6.17.8 is a direct dependency of the two
fractional monetary functions. The shared numeric and argument rules are in
§§4.3.3–4.3.6, 4.11.1–4.11.2, 4.11.5, 4.11.12–4.11.13, and 6.2–6.3.

The financial family has exactly 54 function entries in §§6.12.2–6.12.55.
This document owns only the seven entries above. The cash-flow contract owns
`CUMIPMT`, `CUMPRINC`, `EFFECT`, `FV`, `FVSCHEDULE`, `IPMT`, `IRR`, `ISPMT`,
`MIRR`, `NOMINAL`, `NPER`, `NPV`, `PDURATION`, `PMT`, `PPMT`, `PV`, `RATE`,
`RRI`, `XIRR`, and `XNPV`; the remaining security and coupon entries stay
outside both contracts.

The accepted repository constraints apply here:

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

## Exact signatures and source constraints

The table transcribes the local archive. `Number`, `Integer`, `Logical`, and
`Currency` retain their OpenFormula pseudotypes; they are not host-language
types. The source's `SLN` syntax anomaly is recorded verbatim rather than
silently repaired.

| Section/function | Signature as printed by the archive | Returns | Explicit constraints and defaults |
| --- | --- | --- | --- |
| §6.12.13 `DB` | `DB(Number Cost; Number Salvage; Integer LifeTime; Number Period[; Number Month = 12])` | Currency | `Cost > 0`, `Salvage >= 0`, `LifeTime > 0`, `Period > 0`, and `0 < Month < 13`. `Month` defaults to 12. |
| §6.12.14 `DDB` | `DDB(Number Cost; Number Salvage; Number LifeTime; Number Period[; Number DeclinationFactor = 2])` | Currency | `Cost >= 0`, `Salvage >= 0`, `Salvage <= Cost`, `1 <= Period <= LifeTime`, and `DeclinationFactor > 0`. |
| §6.12.16 `DOLLARDE` | `DOLLARDE(Number Fractional; Integer Denominator)` | Number | `Denominator > 0`. |
| §6.12.17 `DOLLARFR` | `DOLLARFR(Number Decimal; Integer Denominator)` | Number | `Denominator > 0`. |
| §6.12.45 `SLN` | The archive prints `DDB(Number Cost; Number Salvage; Number LifeTime)` under the `SLN` heading. | Currency | The section prints `Constraints: None`; its prose says `LifeTime` is a positive integer. The function-name typo and missing equation are unresolved source issues. |
| §6.12.46 `SYD` | `SYD(Number Cost; Number Salvage; Number LifeTime; Number Period)` | Currency | `Constraints: None` in the source. |
| §6.12.50 `VDB` | `VDB(Number Cost; Number Salvage; Number LifeTime; Number StartPeriod; Number EndPeriod[; Number DepreciationFactor = 2[; Logical NoSwitch = FALSE]])` | Number | `Salvage < Cost`, `LifeTime > 0`, `0 <= StartPeriod <= LifeTime`, `StartPeriod <= EndPeriod <= LifeTime`, and `DepreciationFactor >= 0`. |

`DB` says that when `Month` is specified, `LifeTime` and `Period` are
measured in years. The source distinguishes an omitted optional argument from
an explicitly supplied value in that sentence; the implementation must not
erase that distinction before the profile decision is recorded. `DDB`
explicitly supplies a non-integer `Period` algorithm even though its ordinary
depreciation description speaks of periods. `VDB` likewise describes
fractional periods. The signatures therefore must not be narrowed to integer
periods without an accepted contract amendment.

`VDB` contains a textual tension: its constraint allows
`StartPeriod = EndPeriod`, while its parameter prose says `EndPeriod` is
greater than `StartPeriod`. Both statements are preserved here. The equality
case and its result require a semantic decision before implementation tests
assert behavior.

## Mathematical semantics

The equations and algorithms below are transcribed from the archive's prose
and MathML. A numerically stable implementation may use an equivalent
algorithm after review, but it may not import a spreadsheet-host variant or
silently repair a missing source definition.

### DB

The source first computes and rounds the declining-balance rate to three
decimal places:

```text
rate = 1 - (Salvage / Cost)^(1 / LifeTime)
```

For the first period it defines:

```text
value_1 = Cost * (1 - (Month / 12) * rate)
```

For every `Period <= LifeTime`, the residual value is recursively defined by:

```text
value_Period = value_(Period - 1) * (1 - rate)
```

When `Month` was specified, the source additionally defines:

```text
value_(LifeTime + 1) = value_LifeTime
                       * (1 - (1 - Month / 12) * rate)
```

The allowance is `Cost - value_1` for period 1 and
`value_Period - value_(Period - 1)` for later periods. It is zero for every
period satisfying:

```text
Period > LifeTime + 1 - INT(Month / 12)
```

`INT` refers to §6.17.2. The source does not prescribe a tie rule for “rounded
to 3 decimals,” nor does it state how a fractional `Period` selects a
residual value in the recurrence. Those are explicit profile decisions below;
they must not be hidden in a host-library call.

### DDB

The fixed rate is:

```text
rate = DeclinationFactor / LifeTime
```

For integral periods, the source defines the depreciation for a period as:

```text
book_value_at_start_of_period =
    Cost - sum(DepreciationOfPeriod_i, i = 1 .. Period - 1)
depreciation_of_period =
    MIN(book_value_at_start_of_period * rate,
        book_value_at_start_of_period - Salvage)
```

This stops depreciation at the salvage value. The source then gives the
following algorithm for non-integer `Period`; the branch that sets `rate = 1`
is part of the printed algorithm:

```text
rate = DeclinationFactor / LifeTime
if rate >= 1 then
    rate = 1
    if Period = 1 then
        oldValue = Cost
    else
        oldValue = 0
    endif
else
    oldValue = Cost * (1 - rate)^(Period - 1)
endif

newValue = Cost * (1 - rate)^Period
if newValue < Salvage then
    DDB = oldValue - Salvage
else
    DDB = oldValue - newValue
endif
if DDB < 0 then
    DDB = 0
endif
```

For an integer `Period`, the archive states the relation:

```text
DDB(Cost; Salvage; LifeTime; Period; DeclinationFactor)
  = VDB(Cost; Salvage; LifeTime; Period - 1; Period;
        DeclinationFactor; TRUE)
```

That relation is a validation invariant, not permission to replace one
function's argument or domain rules with the other's.

### DOLLARDE and DOLLARFR

Both functions use §6.17.8 `TRUNC`, whose source syntax is
`TRUNC(Number A; Integer B)` and whose omitted or zero `B` truncates to an
integer. The equations are:

```text
DOLLARDE = TRUNC(Fractional)
            + (Fractional - TRUNC(Fractional)) / Denominator

DOLLARFR = TRUNC(Decimal)
            + (Decimal - TRUNC(Decimal)) * Denominator
```

The denominator is an `Integer` parameter and must be positive. The existing
repository profile converts a non-integer Number to Integer by truncating
toward zero after Number conversion; that profile applies here unless a
function-specific contract amendment says otherwise. `TRUNC`'s treatment of
negative operands comes from §6.17.8 and must be tested directly; do not
replace it with floor.

### SLN

The §6.12.45 heading and summary describe straight-line depreciation, but the
archive prints `DDB(...)` as the syntax, supplies no explicit constraint
beyond the prose statement that `LifeTime` is a positive integer, and prints
no equation or MathML. It also points to `DDB` as the alternative method. This
contract deliberately does not invent an equation or silently change the
signature to `SLN(...)`. A semantic review must pin the intended function name,
parameter type, and straight-line equation before `SLN` can be implemented or
advertised.

### SYD

The source prints this equation:

```text
SYD = (Cost - Salvage) * (LifeTime + 1 - Period) * 2
      / ((LifeTime + 1) * LifeTime)
```

The section declares no constraints. A future implementation must preserve
that absence in the admission layer while still returning the established
numeric formula error if the evaluated expression is non-finite or otherwise
cannot produce a finite Number. It must not silently add a positive-lifetime
or period-range restriction borrowed from `DB`, `DDB`, or `VDB`.

### VDB

The section does not provide a standalone MathML equation. Its normative prose
defines a variable-rate declining-balance calculation over the interval
`StartPeriod` to `EndPeriod`, with these rules:

- `Cost` is greater than `Salvage`; `Salvage` itself may be any value.
- `LifeTime` is positive, and both period arguments lie within the lifetime
  under the stated constraints.
- Fractional `StartPeriod` and `EndPeriod` determine an `initialPeriod`
  option. If both are fractional, the fractional part of `StartPeriod` is
  used.
- The omitted `DepreciationFactor` is 2. It may be zero or any positive value
  under the source constraint.
- When `NoSwitch` is false or omitted, the calculation switches to straight
  line depreciation when that amount exceeds the declining-balance amount.
  When `NoSwitch` is true, it never switches.

The implementation contract must preserve those rules and obtain an accepted
algorithm for the initial-period calculation before claiming VDB support. It
must not infer an Excel-specific switch threshold or fractional-period
rounding rule from the name alone.

## Conversion, shape, and error profile

All seven functions have scalar parameters in the archive. A scalar evaluator
uses the existing Number bridge for `Number` operands: finite Number is
retained, Logical maps to 0/1, and Text uses the repository's fixed finite
numeric-text profile. Malformed or non-finite Text yields the existing formula
error, and Complex values are not silently projected to real Number. `Integer`
operands first use that Number conversion and then the documented
truncation-toward-zero profile. `Logical NoSwitch` uses the existing scalar
Logical conversion; Text is not accepted as a logical spelling in that
profile.

A reference or inline array passed to a scalar parameter follows the existing
implicit-intersection and matrix-selection rules. These functions return one
scalar value; a rectangular input does not turn the result into a result
matrix. The evaluator must not materialize a range merely to convert a scalar
argument. Known shape/type refusals happen before resolver reads, while a
computed child may run until its resulting type is known.

Formula errors are values. Arguments are observed in source order. The first
formula Error encountered in that order is retained as the generated formula
result according to the existing scalar precedence. A typed resolver
`Unsupported`, `ResourceLimit`, `Allocation`, `Cancelled`, `SourceChanged`,
or provider/source-version failure aborts immediately and remains typed; it
supersedes a retained formula Error. `IFERROR` and `IFNA` handle formula-error
values only and do not catch typed failures.

## Bounded-resource and evaluator requirements

These functions do not admit a complete reference sequence. A physical cell
read can occur only while evaluating an argument's existing scalar/reference
intersection, using the evaluator's normal work, read, cancellation, and
source-version checks. There is no financial-specific range scan, cell vector,
or borrowed-cell retention.

The `DB` recurrence, any `VDB` interval walk, and any fallback implementation
of the DDB integral-period definition must have a checked finite work bound.
An implementation may use closed-form powers or bounded segments, but it may
not loop over an unbounded or attacker-controlled lifetime without charging
work and checking cancellation. A budget exhaustion, allocation failure,
non-finite intermediate, or source change remains a typed evaluator failure or
the established numeric formula error as appropriate; it is never silently
clipped to a result.

All temporary numeric state is fixed-size or reserved with checked capacity.
No depreciation function should allocate a vector proportional to `LifeTime`,
`Period`, or a reference's area. Any temporary buffer is dropped before its
associated storage reservation is released. Period, lifetime, factor,
denominator, and month conversions use checked finite arithmetic before
exponentiation, multiplication, rounding, or loop-bound conversion.

The evaluator's source-version and cancellation fences surround the whole
function evaluation and result publication. A demand-cache entry is valid
only when all scalar arguments, projected positions, source identity, and
function-specific optional-presence bits are part of the existing key. A
computed argument is cacheable only under the existing stable source rules.
No function consults an ambient clock, locale, filesystem, network, random
generator, workbook recalculation service, or external financial-rate
provider.

## Open profile decisions and validation

The following points remain open and must be resolved before implementation
evidence is accepted:

1. `SLN` needs a source-corrected function name, parameter typing, and
   straight-line equation. The archive's heading, syntax, and prose do not
   provide a complete executable definition.
2. `DB` does not state the rounding mode for its three-decimal rate or the
   behavior of a fractional `Period`; omitted versus explicitly supplied
   `Month` also affects the source's conditional wording. These must be pinned
   without importing host behavior.
3. `VDB` must resolve the constraint/prose conflict for `StartPeriod =
   EndPeriod` and define the exact finite algorithm for `initialPeriod` and
   fractional interval boundaries.
4. The repository's formula-error category for violated financial domains
   (`DB`/`DDB`/`DOLLAR*`/`VDB`) must be recorded consistently with the shared
   cash-flow contract. Typed resource and provider failures must remain typed.
5. Non-finite powers and results, zero denominators, zero or negative
   lifetimes where the source has no explicit constraint (`SLN`, `SYD`), and
   out-of-range periods need focused tests for the chosen mapping.

Validation must cover exact arity, omitted and explicit missing optional
arguments, every printed domain constraint, integer truncation, negative and
fractional values, the DB three-decimal rate rule, DDB's `rate >= 1` branch,
DDB/VDB integer-period equivalence, DOLLARDE/DOLLARFR round trips and negative
TRUNC cases, SYD's unconstrained inputs, and VDB's no-switch and fractional
period behavior. Resource tests must exercise work/cancellation at the
largest admitted interval, source/provider failure precedence after earlier
formula errors, checked overflow, and absence of lifetime-sized allocation.

Until these decisions and tests are complete, this document records source
scope and implementation boundaries only; it does not claim production
support.
