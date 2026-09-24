# ODF 1.4 security and coupon financial contract

Status: independent source contract.  This document defines the remaining
security, coupon, discount, duration, Treasury-bill, and bond functions in
OpenFormula 1.4 Part 4 §6.12.  It is a semantic and resource boundary for a
future implementation.  It does not claim that any of these functions are
implemented or supported by the production evaluator.

This contract covers the 27 `security-and-coupon` entries in `scope.json`:

`ACCRINT`, `ACCRINTM`, `AMORLINC`, `COUPDAYBS`, `COUPDAYS`, `COUPDAYSNC`,
`COUPNCD`, `COUPNUM`, `COUPPCD`, `DISC`, `DURATION`, `INTRATE`, `MDURATION`,
`ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `PRICE`, `PRICEDISC`,
`PRICEMAT`, `RECEIVED`, `TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `YIELD`,
`YIELDDISC`, and `YIELDMAT`.

The cash-flow and annuity functions in the sibling [`contract.md`](contract.md)
and the depreciation functions in [`depreciation-contract.md`](depreciation-contract.md)
retain their own contracts.  Shared §6.12 conventions remain applicable to
all three documents.

## Authority and accepted boundaries

The normative source is the repository-local ODF distribution.  The archive
and the directly extracted Part 4 member are pinned here so that a later
implementation cannot silently bind this contract to another edition:

| Source | SHA-256 |
| --- | --- |
| `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The function definitions are §§6.12.2–6.12.10, 6.12.15, 6.12.18,
6.12.22, 6.12.26, 6.12.31–6.12.34, 6.12.38–6.12.40, 6.12.43,
6.12.47–6.12.49, and 6.12.53–6.12.55.  Their common semantics use
§§3.2.3, 3.3, 3.6–3.7, 4.3, 4.11.1–4.11.7, 5.6, 6.1–6.3, and the
date-count procedures in §4.11.7.  `DateParam`, date serials, calendar
conversion, and date parsing use the completed date/time contract's selected
profile; this document does not create a second date system.

The design boundary follows the accepted repository decisions:

* [ADR 0001](../../../adr/0001-priorities-and-api-layers.md) requires typed,
  panic-free ordinary APIs and typed refusal for unsupported capability.
* [ADR 0004](../../../adr/0004-semantic-api-design.md) keeps bounded semantic
  values separate from archive and provider types.
* [ADR 0005](../../../adr/0005-io-memory-and-performance.md) requires caller
  budgets, cancellation, source-version checks, fallible storage, and no
  hidden I/O or clock.
* [ADR 0006](../../../adr/0006-validation-security-and-compatibility.md)
  keeps providers explicit, preserves formula Errors as values, and does not
  turn malformed or unsupported input into a silent approximation.
* [ADR 0008](../../../adr/0008-migration-and-verification.md) requires
  compile-tested positive and negative paths before advertising support.

`Currency` and `Percentage` in the signatures are numeric subtypes.  They do
not introduce a formatting or locale dependency into the evaluator.

## Shared scalar, date, and basis profile

The signatures below preserve the Part 4 pseudotypes.  An omitted optional
argument is different from a supplied empty or missing argument.  The
specified default applies only to omission.  A supplied missing slot follows
the evaluator's typed missing-argument conversion rule.

The repository profile for this future family is:

* `Number` accepts finite numeric values, `Logical` as `0` or `1`, and the
  existing fixed numeric-text conversion.  Malformed or non-finite text is a
  formula conversion error.  Complex values are not projected to real
  numbers.
* `Integer` first uses the Number conversion and then truncates toward zero,
  unless an individual function supplies another operation.  The result is
  checked before it is used as a period, frequency, or basis.  This is a
  repository profile for the implementation-defined generic conversion in
  §6.3.6, not a claim that every host uses truncation.
* `DateParam` accepts a Number date serial or Text converted by the selected
  date/time profile.  A date function or coupon schedule truncates fractional
  date/time input where §4.11.7 says `truncate(date)`; it does not silently
  use a host timezone or wall clock.
* `Basis` is the §4.11.7 subtype of Integer.  This profile accepts exactly
  `0..=4` after the checked Integer conversion:

  | Basis | Historical convention | Day-count procedure | Days-in-year procedure |
  | ---: | --- | --- | --- |
  | 0 | US/NASD 30/360 | A | D (360) |
  | 1 | Actual/Actual | B | E (average year procedure) |
  | 2 | Actual/360 | B | D (360) |
  | 3 | Actual/365 | B | F (365) |
  | 4 | European 30/360 | C | D (360) |

  Procedures A–C truncate dates and apply the source's 30/360 day rules;
  procedures B counts actual days from the first date inclusive to the second
  date exclusive; D, E, and F provide the year denominator exactly as §4.11.7
  states.  Counter-factual dates such as February 30 used inside a procedure
  are valid intermediate values and are not calendar parse failures.
* `Frequency` is checked against the function's listed set.  Coupon schedule
  functions accept only 1, 2, or 4.  The bond and irregular-coupon signatures
  declare `Number Frequency` in the source, but their constraints still limit
  the converted value to 1, 2, or 4.  A non-integral or other finite value is
  not silently rounded into a supported frequency.
* A scalar DateParam or Number argument may be supplied by a reference under
  the existing scalar implicit-intersection rules.  A known ReferenceList or
  other non-scalar shape is a pseudotype refusal before resolver reads.  A
  computed expression may be evaluated far enough to establish its resulting
  scalar type.  No function in this slice has a sequence or Array argument in
  its normative signature.

Generated formula errors use the existing evaluator subtype mapping: malformed
arity, scalar conversion, or shape is `#VALUE!`; a violated numeric domain,
non-finite result, invalid date/frequency/basis, or bounded numerical failure
is `#NUM!`; a direct zero denominator is `#DIV/0!`.  The exact subtype for a
source sentence that merely says “Error” must be fixed by the implementation
contract and tests before support is advertised.

Formula Errors already present in an argument remain formula values and are
reported in source/traversal order.  A resolver `Unsupported`, resource,
allocation, cancellation, or source-version failure remains a typed evaluator
failure and supersedes a retained formula Error.  `IFERROR` and `IFNA` handle
formula-error values only; they do not catch provider or resource failures.

## Exact signatures and source constraints

The following table transcribes the archive syntax.  “Should” is retained as
the source's normative strength; the implementation must decide and document
whether that wording is a hard domain gate or a compatibility recommendation
before claiming support.

| Section / function | Part 4 signature | Returns | Explicit source constraints and defaults |
| --- | --- | --- | --- |
| §6.12.2 `ACCRINT` | `ACCRINT(DateParam Issue; DateParam First; DateParam Settlement; Number Coupon; Number Par; Integer Frequency[; Basis B = 0[; Logical CalcMethod = TRUE]])` | Currency | `Issue < First < Settlement`; `Coupon > 0`; `Par > 0`; `Frequency ∈ {1,2,4,12}` (annual, semiannual, quarterly, monthly). |
| §6.12.3 `ACCRINTM` | Intended `ACCRINTM(DateParam Issue; DateParam Settlement; Number Coupon; Number Par[; Basis B = 0])` | Currency | `Coupon > 0`; `Par > 0`. The archive prints `ACCRINT` in this section's syntax; see source anomalies. |
| §6.12.4 `AMORLINC` | `AMORLINC(Number Cost; DateParam PurchaseDate; DateParam FirstPeriodEndDate; Number Salvage; Integer Period; Number Rate[; Basis B = 0])` | Currency | `Cost > 0`; `PurchaseDate ≤ FirstPeriodEndDate`; `Salvage ≥ 0`; `Period ≥ 0`; `Rate > 0`. Equality of the two dates is implementation-defined by the source. |
| §6.12.5 `COUPDAYBS` | `COUPDAYBS(DateParam Settlement; DateParam Maturity; Integer Frequency[; Basis B = 0])` | Number | `Settlement < Maturity`; `Frequency ∈ {1,2,4}`. |
| §6.12.6 `COUPDAYS` | `COUPDAYS(DateParam Settlement; DateParam Maturity; Integer Frequency[; Basis B = 0])` | Number | `Settlement < Maturity`; `Frequency ∈ {1,2,4}`. |
| §6.12.7 `COUPDAYSNC` | Intended `COUPDAYSNC(DateParam Settlement; DateParam Maturity; Integer Frequency[; Basis B = 0])` | Number | `Settlement < Maturity`; `Frequency ∈ {1,2,4}`. The printed syntax spells the function `COUPDAYNC`. |
| §6.12.8 `COUPNCD` | `COUPNCD(DateParam Settlement; DateParam Maturity; Integer Frequency[; Basis B = 0])` | Date | `Settlement < Maturity`; `Frequency ∈ {1,2,4}`. The source summary and semantics say this is the next coupon date after Settlement. |
| §6.12.9 `COUPNUM` | `COUPNUM(DateParam Settlement; DateParam Maturity; Integer Frequency[; Basis B = 0])` | Number | `Frequency ∈ {1,2,4}`. The body does not explicitly repeat `Settlement < Maturity`; do not add that constraint without a profile decision. |
| §6.12.10 `COUPPCD` | `COUPPCD(DateParam Settlement; DateParam Maturity; Integer Frequency[; Basis B = 0])` | Date | `Settlement < Maturity`; `Frequency ∈ {1,2,4}`. The source summary calls this the coupon date prior to Settlement. |
| §6.12.15 `DISC` | `DISC(DateParam Settlement; DateParam Maturity; Number Price; Number Redemption[; Basis B = 0])` | Percentage | `Settlement < Maturity`. No additional Price or Redemption constraint is printed. |
| §6.12.18 `DURATION` | `DURATION(Date Settlement; Date Maturity; Number Coupon; Number Yield; Number Frequency[; Basis B = 0])` | Number | `Yield ≥ 0`; `Coupon ≥ 0`; `Settlement ≤ Maturity`; `Frequency ∈ {1,2,4}`. The source says `Date`, not `DateParam`; this type spelling is an anomaly to resolve against the common date profile. |
| §6.12.22 `INTRATE` | `INTRATE(Date Settlement; Date Maturity; Number Investment; Number Redemption[; Basis Basis = 0])` | Number | `Settlement < Maturity`. The source does not print positivity constraints for Investment or Redemption. The repeated `Basis Basis` spelling is retained as an anomaly. |
| §6.12.26 `MDURATION` | `MDURATION(Date Settlement; Date Maturity; Number Coupon; Number Yield; Number Frequency[; Basis B = 0])` | Number | `Yield ≥ 0`; `Coupon ≥ 0`; `Settlement ≤ Maturity`; `Frequency ∈ {1,2,4}`. The source uses `Date` rather than `DateParam`. |
| §6.12.31 `ODDFPRICE` | `ODDFPRICE(DateParam Settlement; DateParam Maturity; DateParam Issue; DateParam First; Number Rate; Number Yield; Number Redemption; Number Frequency[; Basis B = 0])` | Number | `Rate`, `Yield`, and `Redemption` “should be greater than 0”; `Frequency ∈ {1,2,4}`. No date ordering is printed in this section. |
| §6.12.32 `ODDFYIELD` | `ODDFYIELD(DateParam Settlement; DateParam Maturity; DateParam Issue; DateParam First; Number Rate; Number Price; Number Redemption; Number Frequency[; Basis B = 0])` | Number | `Rate`, `Price`, and `Redemption` “should be greater than 0”; `Maturity > First > Settlement > Issue`; `Frequency ∈ {1,2,4}`. |
| §6.12.33 `ODDLPRICE` | `ODDLPRICE(DateParam Settlement; DateParam Maturity; DateParam Last; Number Rate; Number AnnualYield; Number Redemption; Number Frequency[; Basis B = 0])` | Number | `Rate`, `AnnualYield`, and `Redemption` “should be greater than 0”; `Maturity > Settlement > Last`; `Frequency ∈ {1,2,4}`. |
| §6.12.34 `ODDLYIELD` | `ODDLYIELD(DateParam Settlement; DateParam Maturity; DateParam Last; Number Rate; Number Price; Number Redemption; Number Frequency[; Basis B = 0])` | Number | `Rate`, `Price`, and `Redemption` “should be greater than 0”; no date ordering is printed; `Frequency ∈ {1,2,4}`. |
| §6.12.38 `PRICE` | `PRICE(DateParam Settlement; DateParam Maturity; Number Rate; Number AnnualYield; Number Redemption; Number Frequency[; Basis Bas = 0])` | Number | `Rate`, `AnnualYield`, and `Redemption` “should be greater than 0”; `Frequency ∈ {1,2,4}`. The optional basis parameter is printed as `Bas`. |
| §6.12.39 `PRICEDISC` | `PRICEDISC(DateParam Settlement; DateParam Maturity; Number Discount; Number Redemption[; Basis B = 0])` | Number | `Discount` and `Redemption` “should be greater than 0”. No Settlement/Maturity ordering is printed. |
| §6.12.40 `PRICEMAT` | `PRICEMAT(DateParam Settlement; DateParam Maturity; DateParam Issue; Number Rate; Number AnnualYield[; Basis B = 0])` | Number | `Settlement < Maturity`; `Rate ≥ 0`; `AnnualYield ≥ 0`; if both rates are zero, return `100`. |
| §6.12.43 `RECEIVED` | `RECEIVED(DateParam Settlement; DateParam Maturity; Number Investment; Number Discount[; Basis B = 0])` | Number | `Investment > 0`; `Discount > 0`; `Settlement < Maturity`. |
| §6.12.47 `TBILLEQ` | `TBILLEQ(DateParam Settlement; DateParam Maturity; Number Discount)` | Number | Maturity is less than one year beyond Settlement; Discount is positive. The exact ordering and zero-day behavior are not separately stated. |
| §6.12.48 `TBILLPRICE` | `TBILLPRICE(DateParam Settlement; DateParam Maturity; Number Discount)` | Number | Maturity is less than one year beyond Settlement; Discount is positive. |
| §6.12.49 `TBILLYIELD` | `TBILLYIELD(DateParam Settlement; DateParam Maturity; Number Price)` | Number | Maturity is less than one year beyond Settlement; Price is positive. |
| §6.12.53 `YIELD` | `YIELD(DateParam Settlement; DateParam Maturity; Number Rate; Number Price; Number Redemption; Number Frequency[; Basis B = 0])` | Number | `Rate`, `Price`, and `Redemption` “should be greater than 0”; `Frequency ∈ {1,2,4}`. |
| §6.12.54 `YIELDDISC` | `YIELDDISC(DateParam Settlement; DateParam Maturity; Number Price; Number Redemption[; Basis B = 0])` | Number | `Price > 0`; `Redemption > 0`. No date ordering is printed. |
| §6.12.55 `YIELDMAT` | `YIELDMAT(DateParam Settlement; DateParam Maturity; DateParam Issue; Number Rate; Number Price[; Basis B = 0])` | Number | `Rate > 0`; `Price > 0`. No date ordering is printed. |

## Coupon schedule and accrual semantics

The functions in this group derive coupon periods from Maturity, Settlement,
Frequency, and Basis.  They must use the same checked civil-date and day-count
kernel as `YEARFRAC`; a second schedule algorithm in a matrix bridge would be
a semantic fork.

* `COUPDAYBS` returns the number of days from the beginning of the coupon
  period containing Settlement to Settlement.
* `COUPDAYS` returns the number of days in the coupon period containing
  Settlement.
* `COUPDAYSNC` returns the number of days from Settlement to the next coupon
  date.
* `COUPNCD` returns the next coupon date after Settlement, and `COUPPCD`
  returns the coupon date immediately prior to Settlement.  The source does
  not display a closed equation for either date, so their schedule roll and
  date-boundary profile must still be tested against the shared date-count
  procedures.
* `COUPNUM` returns the number of outstanding coupons in the interval from
  Settlement to Maturity using the coupon schedule.  Because the source does
  not repeat the Settlement ordering constraint in this section, reversed or
  equal dates require a documented profile rather than an invented gate.
* `ACCRINT` supports short, standard, and long first coupon periods.  For
  every coupon period, its interest contribution is

  ```text
  Par * Coupon * YEARFRAC(start_of_period; end_of_period; B)
  ```

  With `CalcMethod = TRUE` (the default), sum the accrued contributions from
  Issue through Settlement.  With `FALSE`, sum from First through Settlement.
  The source describes this as a sum across periods; it does not authorize a
  single whole-interval `YEARFRAC` substitution for irregular first or long
  periods.
* `ACCRINTM` computes accrued interest for a security that pays at maturity.
  The archive gives no displayed equation.  A future implementation must
  choose and document the maturity-accrual equation, date interval, and
  handling of `Issue = Settlement` before it is tested as supported.

The date schedule must not silently change the requested Frequency or Basis.
Date arithmetic that overflows the selected date profile is a numeric formula
error.  Coupon-period iteration is finite from checked date bounds and must
charge work per generated period; it must not allocate a cell-sized schedule
or perform hidden resolver reads.

## Depreciation and simple discount semantics

`AMORLINC` is linear French accounting depreciation.  When `Period = 0`, the
source gives:

```text
Cost * Rate * YEARFRAC(PurchaseDate; FirstPeriodEndDate; B)
```

For a full period after period zero, the depreciation is `Cost * Rate`.  For
the last, possibly partial, period it is:

```text
Cost - Salvage - accumulated_depreciation
```

where accumulated depreciation includes period zero and all earlier full
periods.  Once `Period > (Cost - Salvage) / (Cost * Rate)`, the result is
zero.  The source marks `PurchaseDate = FirstPeriodEndDate` implementation-
defined; this equality must be a named profile case, not an accidental divide
by zero or an arbitrary zero-period result.

`DISC` has the explicit equation:

```text
((Redemption - Price) / Redemption)
    / YEARFRAC(Settlement; Maturity; B)
```

`INTRATE` has the explicit equation:

```text
((Redemption - Investment) / Investment)
    / YEARFRAC(Settlement; Maturity; B)
```

The source uses `rate`-style percentage numbers without a display conversion;
the evaluator returns the finite numeric ratio.  Zero denominators and
non-finite intermediates use the common numeric error mapping.

`RECEIVED` has the explicit equation:

```text
Investment / (1 - Discount * YEARFRAC(Settlement; Maturity; B))
```

The denominator is checked before division.  A zero denominator is a direct
`#DIV/0!` under the repository profile; non-finite or otherwise invalid
intermediate arithmetic is `#NUM!`.

`PRICEDISC`, `PRICEMAT`, and `YIELDMAT` describe prices or yields but the
archive member does not provide a displayed equation in their sections.
They must not receive an uncited Excel formula merely because a host has one.
Their equations, zero-denominator behavior, and inverse relation must be
recorded as an implementation profile before a test or support claim is
accepted.  `PRICEMAT`'s explicit all-zero rule is nevertheless normative:
`PRICEMAT(...; Rate = 0; AnnualYield = 0; ...) = 100`.

## Periodic and irregular bonds

`DURATION` computes Macaulay duration for a fixed-interest security from
Settlement, Maturity, Coupon, Yield, Frequency, and Basis.  `MDURATION` uses
the source's displayed relation:

```text
duration = DURATION(Settlement; Maturity; Coupon; Yield; Frequency; B)
MDURATION = duration / (1 + Yield / Frequency)
```

The source does not display the cash-flow sum used by `DURATION`; the
implementation must document the coupon-date weighting, the final redemption
cash flow, settlement-period fraction, and zero-yield behavior rather than
silently importing a host's convention.

`PRICE` has the one explicit bond-price equation in this batch.  Let:

* `A` be the days from Settlement to the next coupon date;
* `B` be the days in the coupon period containing Settlement;
* `C` be the coupons between Settlement and Maturity/Redemption; and
* `D` be the days from the beginning of that coupon period to Settlement.

Then the source gives:

```text
PRICE = Redemption / (1 + Yield / Frequency)^(C - 1 + A / B)
      + sum(k = 1..C,
            (100 * Rate / Frequency)
              / (1 + Yield / Frequency)^(k - 1 + A / B))
      - 100 * Rate / Frequency * D / B
```

The signature calls the input `AnnualYield`; the equation calls it `Yield`.
The implementation profile must bind those names to the same input and must
define the zero-period, zero-yield, and `B = 0` cases without an unchecked
division.

`ODDFPRICE` and `ODDFYIELD` cover an irregular first interest date (`Issue`,
`First`).  `ODDLPRICE` and `ODDLYIELD` cover an irregular last interest date
(`Last`).  The archive gives no displayed equations for these four functions.
Their coupon schedule construction, stub-period day fraction, compounding,
root selection, and nonconvergence result are therefore explicit open profile
items.  The printed date order and positive-value wording in the signature
table remain the only source constraints until that profile is accepted.

`YIELD` is the periodic-bond yield counterpart to `PRICE`, but the archive
section supplies no displayed inverse equation or solver rule.  A future
implementation must define whether it solves the `PRICE` equation above,
which root is selected when multiple roots exist, the finite iteration/work
bound, and the formula error on nonconvergence.

## Treasury-bill semantics

`TBILLEQ` is the only Treasury-bill function in this batch with a displayed
equation.  Let `DSM` be the number of days between Settlement and Maturity
using the Actual/360 day-count basis (Basis 2).  The source writes:

```text
TBILLEQ = (365 * rate) / (360 - rate * DSM)
```

The parameter is named `Discount` in the signature while the equation uses
`rate`; the implementation profile must treat these as aliases, not two
inputs.  The source says that Maturity is less than one year beyond
Settlement and Discount is positive, but it does not state strict Settlement
< Maturity or the zero-day denominator case.

`TBILLPRICE` computes a Treasury-bill price per 100 face value from Discount;
`TBILLYIELD` computes its yield from Price.  Neither section provides a
displayed equation.  Their 360-day convention, date interval, price/yield
inverse, zero-day behavior, and bounded numerical error policy must be
profiled before implementation evidence is accepted.

## Conversion, errors, and resolver/resource behavior

These functions are scalar financial reducers and schedule calculations, not
sequence consumers.  A direct range is therefore handled through the scalar
implicit-intersection and `DateParam`/`Number` conversion rules.  It must not
be eagerly materialized into a vector merely because it appears as an
argument.  A known wrong shape such as a ReferenceList is rejected before any
resolver cell read; a computed expression can run until its scalar result is
known.  This keeps the behavior distinct from `NPV`/`XIRR`-style complete
sequence consumers in the sibling contract.

Resolver-backed evaluation has the following required shape:

* Charge work and the reference-read budget before each cell read, check
  cancellation before and after the provider call, and charge only a
  successful read.  Do not create a cell vector to convert one scalar
  DateParam or Number.
* Retain borrowed text while converting a cell; do not clone the full source
  text merely to parse a scalar.  Any text-length or intermediate allocation
  limit is a typed resource failure.
* Preserve formula Errors as values at the argument boundary.  Continue only
  where an operation has a complete finite schedule or solver scan to finish;
  a typed provider/resource/cancellation/source failure always returns
  immediately and supersedes retained formula Errors.
* Source-version and cancellation fences surround the entire resolver-backed
  evaluation, including any final publication step.  There is no ambient
  clock for date parameters and no hidden external provider.
* Coupon-period generation, irregular-stub arithmetic, and root-finding use
  checked work and storage reservations.  Bounds are charged for period
  count, solver iterations, and any retained numeric state.  A solver cannot
  turn a caller-supplied guess into an unbounded loop or allocation.
* State is fixed-size or bounded by the checked number of periods; a security
  calculation does not retain all source cells.  If a future implementation
  needs a coupon schedule, it must use a bounded stream or exact checked
  reservation and release entries as soon as they are no longer needed.

An argument formula Error is reported in source argument order.  If a
schedule or numerical operation discovers a generated domain error after an
earlier formula Error, the retained formula Error wins under the common
formula-error precedence.  Typed provider failures remain outside that
precedence and always bubble unchanged.  No `COUNT`-style reducer rule that
ignores errors is imported into this family.

## Source anomalies and open profiles

The following points require an explicit decision and focused tests before
this contract can authorize implementation support:

1. **Printed names and HTML anchors.**  §6.12.3 is headed `ACCRINTM` but
   prints `ACCRINT` in its syntax.  §6.12.7 prints `COUPDAYNC` while the
   heading and inventory name are `COUPDAYSNC`.  The HTML also puts a second
   `id="COUPNCD"` anchor on the §6.12.7 heading, before the actual §6.12.8
   `COUPNCD` heading; several See-also links consequently point to section
   6.12.7.  The §6.12.8 body itself correctly says Date and “next coupon
   date”; the duplicate anchor must not be mistaken for a second function.
2. **Cross-reference defects.**  Several coupon See-also links point to the
   wrong section number (for example the `COUPNCD` link around the coupon
   entries).  Links are navigation only; the numbered heading and body are
   authoritative for this contract.
3. **Type spelling.**  `DURATION` and `MDURATION` use `Date` where adjacent
   financial functions use `DateParam`.  `INTRATE` prints `Basis Basis`.
   `PRICE` prints `Basis Bas`.  These are recorded source spellings; the
   implementation must bind them to the shared DateParam/Basis profile and
   retain the aliases in its compatibility documentation.
4. **Equation omissions.**  The archive has no displayed equation for
   `ACCRINTM`, the coupon schedule functions, `DURATION`, the four irregular
   bond functions, `PRICEDISC`, `PRICEMAT`, the two non-equivalent Treasury
   bill functions, `YIELD`, or `YIELDMAT`.  A host formula or financial-library
   convention is not normative evidence.  Each missing equation needs a
   separately reviewed profile, including date-fraction conventions and
   inverse/root behavior.
5. **Constraint wording.**  `ODDF*`, `ODDL*`, `PRICE`, `PRICEDISC`, `YIELD`,
   and `RECEIVED` use “should be greater than 0” in places where other
   sections use a hard inequality.  The contract preserves that wording and
   leaves the hard-gate/error mapping open.  `COUPNUM`, `PRICEDISC`,
   `YIELDDISC`, and several Treasury/bond sections omit Settlement/Maturity
   ordering constraints; no extra ordering rule is inferred.
6. **AMORLINC equality.**  The source explicitly makes
   `PurchaseDate = FirstPeriodEndDate` implementation-defined.  The zero
   first-period result, denominator behavior, and period numbering need a
   profile test.
7. **Frequency conversion.**  Coupon schedule entries declare Integer
   Frequency, while bond/irregular entries declare Number Frequency but list
   only 1, 2, and 4.  This contract rejects fractional or unsupported values;
   any host that rounds them must be documented as a compatibility profile,
   not presented as direct ODF parity.
8. **Date and schedule boundaries.**  The Treasury-bill “less than one year”
   condition, same-day settlement, maturity equality, leap-day coupon dates,
   end-of-month roll, and negative/fractional date serial handling are not
   fully specified by the individual sections.  They must use the selected
   date/basis procedures and have explicit boundary tests.
9. **Numerical solvers.**  The source does not specify tolerances, iteration
   counts, multiple-root selection, or nonconvergence mapping for yield and
   irregular-price operations.  A future implementation must use finite
   checked work, deterministic root selection, finite intermediates, and a
   typed/formula outcome documented before support is claimed.

These open items are intentional.  This file records what the local normative
source says and where it stops; it is not permission to fill a gap by copying
Excel, LibreOffice, or another host's behavior.

## Required validation before support

Before advertising any function in this batch, independent validation must
cover at least:

* exact arity, omission versus explicit missing optional arguments, every
  listed default, and all source-signature aliases/typos;
* each explicit inequality, Frequency and Basis value, boundary date, leap
  day, end-of-month schedule, zero denominator, and non-finite result;
* scalar literals, single-cell references, multi-cell implicit intersections,
  ReferenceList refusal with zero reads, computed scalar expressions, and
  formula Errors versus typed resolver failures;
* ordinary, short/long/irregular coupon periods, coupon-period ordering, and
  all five day-count bases;
* explicit equations (`ACCRINT`, `AMORLINC`, `DISC`, `INTRATE`, `MDURATION`,
  `PRICE`, `RECEIVED`, `TBILLEQ`, and `YIELDDISC`) against an independent
  oracle, with finite-result and denominator checks; and
* every separately approved profile for functions whose archive section has no
  equation, including bounded solver behavior and native compatibility
  differences.

Until those checks and the unresolved source profiles are complete, this
contract remains a source review artifact and makes no production support
claim.
