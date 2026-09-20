# ODS date and time contract semantic review

Status: **contract semantic PASS; source semantic review provisional PASS
pending the final frozen-source gates**. This review covers the complete
twenty-four-function family in ODF 1.4 Part 4 §6.10. It does not claim final
production support, frozen-gate passage, native parity, or performance.

The reviewed contract is
[`contract.md`](contract.md), SHA-256
`cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`.
The repository-local normative inputs are:

* `3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
  `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`;
* `part4-formula/OpenDocument-v1.4-os-part4-formula.html`, SHA-256
  `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.

The review also checked the existing `EvaluationOptions`,
`EvaluationContext`, value `Context`, calendar helpers, and accepted ADRs
0001, 0004, 0005, 0006, 0008, and 0023. No production, test, Cargo, or
oracle files were changed by this review.

## Normative scope and profile decisions

The contract lists exactly DATE, DATEDIF, DATEVALUE, DAY, DAYS, DAYS360,
EASTERSUNDAY, EDATE, EOMONTH, HOUR, ISOWEEKNUM, MINUTE, MONTH, NETWORKDAYS,
NOW, SECOND, TIME, TIMEVALUE, TODAY, WEEKDAY, WEEKNUM, WORKDAY, YEAR, and
YEARFRAC. Their signatures and result subtypes match §§6.10.2–6.10.25,
including the optional-argument arities and the DateSequence and
LogicalSequence slots.

The following choices are implementation-defined or host-defined by ODF and
are stated explicitly enough to be testable:

* The profile uses proleptic Gregorian dates with epoch 1899-12-30 and no
  synthetic 1900-02-29. The checked family domain is year 0001 through 9999,
  serial `-693,593` through the half-open DateTime bound below `2,958,466`.
  Integer date results stop at `2,958,465`, while a final-day fraction is
  valid. This wider domain is intentionally separate from the existing VALUE
  domain, which remains nonnegative.
* DATE retains ODF month/day rollover, while the profile rejects zero or
  negative Month/Day and rejects a normalized result outside the family
  domain. The wider year profile is a documented extension of DATE's ordinary
  1904–9956 interoperable constraint and is needed to represent
  `EASTERSUNDAY(1583)`.
* DateParam and TimeParam Logical conversion is the selected 1/0 profile.
  Text parsing is fixed en_US with English month names, period/comma numeric
  syntax, and a 1930 two-digit-year pivot. The fallback after DATEVALUE or
  TIMEVALUE is deliberately numeric-only, but includes the fixed VALUE
  numeric branch (exponents, percent, grouping, currency, and valid simple or
  mixed fractions), so a failed date/time parse does not recursively admit the
  full VALUE date/time grammar. DATEVALUE floors a successful numeric
  fallback; TIMEVALUE preserves its raw finite numeric result. The
  contract's examples (`DATEVALUE("123.5")` -> `123`,
  `TIMEVALUE("2.5")` -> `2.5`, and rejection of a clock in DATEVALUE or a
  date in TIMEVALUE) make this boundary executable.
* DATEDIF's reversed-interval #NUM! behavior, NETWORKDAYS's negative reverse
  count, WORKDAY's zero-offset exact serial preservation, and WORKDAY's
  toward-zero Offset truncation are explicit profile choices where ODF does
  not fully specify the result.
* The contract selects the explicit rounded-total-seconds MINUTE formula from
  the two conflicting formulas printed in §6.10.13 and shares that state with
  SECOND. This resolves the source ambiguity rather than silently mixing the
  two formulas. TIME selects the direct fractional formula and deliberately
  does not apply the optional INT preprocessing allowed by §6.10.18.
* YEARFRAC uses the ordered-date procedures delegated by §6.10.25. Procedures
  A, B, and C in §4.11.7 swap reversed dates and return a nonnegative count;
  Procedure E is applied to those chronological dates. The contract does not
  add an unsupported sign restoration.

These are coherent profile decisions, not claims that every ODF host must
make the same implementation-defined choice.

## Function-by-function review

DATE and DATEVALUE preserve the required rollover, ISO acceptance, combined
datetime integer extraction, and explicit domain/error distinction. DAY,
MONTH, YEAR, ISOWEEKNUM, WEEKDAY, and WEEKNUM use the proleptic Gregorian
calendar and state their boundary modes. The WEEKDAY table covers 1, 2, 3,
and 11–17 exactly; WEEKNUM covers 1, 2, 11–17, 21, and 150 exactly.

DAYS correctly retains numeric fractions only for its direct Number/Number
case and otherwise uses the declared DateParam conversions. DAYS360 keeps the
important normative distinction: US/NASD dates are never swapped, while the
European procedure swaps and applies its negative sign. EDATE and EOMONTH
truncate/floor the source date according to the selected civil-date profile,
truncate month addition toward zero, clamp month ends, and use checked
calendar arithmetic.

DATEDIF's `YM` rule is now executable: form the signed total month delta,
decrement it when EndDate's day precedes StartDate's day, then apply
Euclidean modulo 12. Thus the cross-year January example returns 11 rather
than −1. The profile explicitly retains the intended distinction from `M`:
the complete-month `M` calculation uses a clamped anniversary, so
January 31→February 28 is one month, while raw day comparison makes `YM`
zero for that same-month remainder.

EASTERSUNDAY retains the exact explicit-year bound 1583–9956 and the
algorithmic result. Its no-Year form uses one injected timestamp, compares
the current and following eligible Easter, and explicitly handles the
9956-after-Easter and 9957–9999 boundary instead of attempting an
out-of-domain year. HOUR follows the normative INT/floor day-fraction rule;
MINUTE and SECOND share the selected rounded-second profile. The reviewed
source applies that profile literally as `ROUND(T * 86400)` followed by the
day/second modulo operations; a finite input whose multiplication becomes
non-finite therefore reaches the existing `#NUM!` boundary. This means
`MINUTE(-0.5/86400)` and `SECOND(-0.5/86400)` are both 59 under the selected
half-away-from-zero rule. The current oracle row
`second.negative_fraction` still expects 0, which is the rejected
normalized-fraction alternative and must be corrected before oracle evidence
is accepted; this is an evidence-fixture issue, not a source semantic defect.
TIME accepts any finite component values as §6.10.18 permits, with checked
arithmetic, and TIMEVALUE distinguishes clock parsing from its permitted
numeric fallback.

The independent source review of `date_time.rs`, `date_time/kernel.rs`, and
the shared date/time parser found no additional semantic blocker. The recent
WEEKNUM exact-mode validation, EASTERSUNDAY no-Year boundary handling,
fractional-second parsing, and negative-subnormal HOUR decomposition agree
with the contract. Diagnostic focused checks reported by the implementation
owners are evidence for the pending gate run only; they do not replace the
frozen receipts.

NETWORKDAYS and WORKDAY match the §6.10.15 and §6.10.23 signatures and
default workweek. Their sequence rules correctly distinguish scalar sequence
conversion, direct reference filtering, and inline-array profile behavior.
The contract rejects a ReferenceList for DateSequence because §6.3.9 admits a
Reference but does not define the ReferenceList extension reserved for
NumberSequenceList; a single cuboid Reference remains a valid streamed
sequence. The permitted row/column order and sheet order follow §4.11.12.
Formula Errors are retained while an admitted sequence is consumed, whereas
typed resolver, source, cancellation, allocation, and resource failures
remain typed and supersede formula results. This is consistent with the
accepted typed-failure boundary and with eager function-argument evaluation.

YEARFRAC maps bases 0–4 to Procedures A–F exactly: US 30/360, Actual/Actual,
Actual/360, Actual/365, and European 30/360. The contract preserves the
counterfactual February-30 values as internal procedure state rather than
turning them into formula errors.

## Timestamp and API review

The proposed timestamp seam fits the existing API. `EvaluationOptions` is
already a small `Copy + Eq` value with private fields, and both
`EvaluationContext` and the resolver-backed value `Context` copy it by value.
An optional validated `CalculationTimestamp` with consuming builder/accessor
methods therefore preserves the existing construction model and keeps
`ExecutionContext` free of a clock provider. This follows ADR 0006's explicit
provider rule and ADR 0005's runtime-neutral execution boundary.

`CalculationTimestamp::from_serial` must reject non-finite or out-of-domain
values before they enter options. The proposed canonicalized bit-preserving
representation is compatible with exact fractional serials and `Copy + Eq`;
negative zero must be canonicalized. A civil constructor may be provided, but
it must be timezone-free profile civil time and return a typed validation
error. Host code that wants local or UTC wall time must explicitly convert it
to a profile serial before constructing the value. The evaluator must never
consult `SystemTime`, a process timezone, environment state, or workbook
metadata.

NOW, TODAY, and no-Year EASTERSUNDAY use one immutable timestamp snapshot per
evaluation. NOW returns its exact serial, TODAY its civil-date component, and
projected matrix outputs reuse that same snapshot. A missing timestamp is the
typed `Unsupported(CalculationClock)` capability boundary, not a formula
error. Implementing that contract requires adding the dedicated clock
variant to the existing non-exhaustive `UnsupportedKind`; this is an
implementation task, not a semantic objection to the contract.

Timestamp-backed cache entries must include the complete snapshot, and cache
hits must remain inside the existing source-version and cancellation fences.
The timestamp does not authorize a cache hit to bypass source publication or
typed-failure checks.

## Matrix, reference, and resource consequences

The scalar evaluator has no resolver or position, so context-free kernels and
timestamp-backed volatile functions are valid there; scalar references and
sequence reads must remain typed `Unsupported(Reference)`. The value evaluator
may use ordinary scalar intersection/projected elementwise scheduling for
Date and Offset, while NETWORKDAYS and WORKDAY consume holiday/workweek
sequences completely for every projected result. ReferenceList and other
known shape refusals occur before resolver reads; computed expressions may
read while their resulting type is established.

The contract's streaming and bounded-state requirements are consistent with
ADR 0005: reference sequence scans charge work and read limits before every
cell, check cancellation before and after resolver calls, retain borrowed
text, and preserve source/final-cancellation fences. Holiday storage,
parser scratch, sequence metadata, and matrix outputs are explicitly bounded
and fallibly reserved. Formula-level handlers cannot catch typed provider or
resource failures. Demand-cache entries contain complete scalar payloads or
errors only and preserve position, projected shape, source identity,
cancellation identity, and timestamp identity.

## Disposition

No remaining normative, API-design, or source-semantic blocker was found for
contract `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`.
The resource companion review on disk is bound to this same contract hash;
its streaming, bounded-state, typed-precedence, timestamp-cache, and
source-fence findings may be cited with the final frozen receipts. The
source disposition remains provisional until those receipts are complete,
and the stale `second.negative_fraction` oracle expectation must be corrected
before the oracle corpus can serve as final evidence.

Before implementation acceptance, the production work must still prove the
dedicated clock error variant, the separate VALUE/date-family domains,
negative-date floor behavior, numeric-only fallback, the selected MINUTE
formula, ordered YEARFRAC procedures, all sequence/error precedence cases,
and the timestamp-aware cache/source fences. Those are validation gates, not
open contract findings. The historical production baseline at commit
`6fa3b8af6a` lacked date/time dispatch; the current working tree is under
implementation review, so this report makes no production-support claim.
