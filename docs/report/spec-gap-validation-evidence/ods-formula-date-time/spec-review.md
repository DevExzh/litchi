# ODS date and time contract semantic review

Status: **contract semantic PASS; frozen-source and isolated-gate PASS**. This
review covers the complete twenty-four-function family in ODF 1.4 Part 4
§6.10. Performance capture remains separate and pending; native parity is
not claimed.

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

## Frozen source identity

The reviewed candidate is the selected source in
[`gates/freeze.json`](gates/freeze.json), SHA-256
`461a76708e36b2a716cd622df45f014e009186223bd1db511c3e0ff0f7fa3561`, at
candidate commit `d16039ce48cb441c35461318c8634a49ae0b2437`. The manifest
contains 103 selected files and uses isolated `Cargo.lock` SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`.
The Rust implementation and focused test hashes are unchanged from the
reviewed candidate; the corrected freeze adds the final oracle corpus and
provenance bindings.
The directly reviewed implementation and test inputs have these manifest
hashes:

| Input | SHA-256 |
| --- | --- |
| `evaluation/date_time.rs` | `29f1495a55aac7effee9f6509deef38e5d5f747ba2b87c65a188d8759c1265e3` |
| `evaluation/date_time/kernel.rs` | `2bc377db1a01e28dbc4940c99c61fb43a308919b315bfc37a92ce86a67e5d0cc` |
| `evaluation/date_time/parser.rs` | `e71f8cb4075f4815a8c8a43ba5de48e0f5a411143e47c575c4a6de8bc4ac7149` |
| `evaluation/inspection/parse_value.rs` | `7124afd22adbe4f1e3f0182feecbabb228844f5c6b93734761f214675137f2f7` |
| `evaluation/value/date_time.rs` | `1da16aae011dd8c94fafc39fdf81cc5d9f8019024ab3860bdc00b4d28962e159` |
| `evaluation/value.rs` | `e2f9e5511e27e4a165dbeb7b0fba4ea8721231c610773a98f54b846c232c61f3` |
| `tests/ods_formula_date_time_evaluation.rs` | `d9644432fd5dcac9272dcd3f942e4b669865db22fb6fe866a706ebd05bf04ea1` |
| `tests/ods_formula_date_time_limits.rs` | `fbf6ba638d7b80ebcda36d1091641ef45e42be07f9bc20a987583a3987fed411` |
| `tests/ods_formula_date_time_oracle.rs` | `6ab4f5845c58aa88b5890abc766229a4970409cfe581b924228bb4b595e97a05` |
| `oracle-vectors.json` | `b33011089974b18b1fcc0b984adab839600acc98f09dd850e02522348316259f` |

The `evaluation/value/date_time.rs` hash above is the frozen source hash from
the manifest. The current isolated receipt reports all seven required
commands successful, stable sources, and 1,745 package tests with no failures
or ignored tests. Its hashes are:

| Receipt | SHA-256 |
| --- | --- |
| `gates/results.json` | `fa5dfab0bef43d1403b22204185ae7e2e8fffe31e753c07860fb4768c228286a` |
| `gates/verification.json` | `bdeb941ccf550e43a341df081dcd54e1ba9e70779744ecb80e785cca71739243` |
| `gates/source-before.json` | `864ccd3b76cc866e23c7db2cbdc1111702984e225cb7b630adc0befc2bde6a1e` |
| `gates/source-after.json` | `864ccd3b76cc866e23c7db2cbdc1111702984e225cb7b630adc0befc2bde6a1e` |

The superseded provenance-typo receipt remains retained for historical
custody:

| Receipt | SHA-256 |
| --- | --- |
| `gates/history-provenance-typo/results.json` | `0a011cd5c6e6ee18e61c8327a3d27f1d41f9f3f6836d55c6148ebf56164ba2d3` |
| `gates/history-provenance-typo/verification.json` | `bdeb941ccf550e43a341df081dcd54e1ba9e70779744ecb80e785cca71739243` |

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
`second.negative_fraction` now expects 59 with the raw half-away basis. The
earlier expectation of 0 was the rejected normalized-fraction alternative;
that historical finding was corrected before the frozen oracle replay.
The corrected corpus and its 132-vector Rust replay are therefore consistent
with the source profile.
TIME accepts any finite component values as §6.10.18 permits, with checked
arithmetic, and TIMEVALUE distinguishes clock parsing from its permitted
numeric fallback.

The independent source review of `date_time.rs`, `date_time/kernel.rs`, and
the shared date/time parser found no additional semantic blocker. The recent
WEEKNUM exact-mode validation, EASTERSUNDAY no-Year boundary handling,
fractional-second parsing, and negative-subnormal HOUR decomposition agree
with the contract. The frozen value adapter keeps scalar local references in
the scalar-demand path, leaves projected computed and array-valued date/time
parameters position-sensitive, and only treats complete sequence arguments as
invariant. Timestamp functions cache a scalar payload only when an explicit
calculation snapshot exists; the added volatile-cache unit coverage checks
both that reuse and the typed clock refusal. Diagnostic focused checks
reported by the implementation owners (22 semantic and 11 resource cases)
are supporting evidence only; they do not replace the final gate receipts.

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
the frozen candidate under contract
`cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`. The
resource companion review on disk is bound to this same contract hash; its
streaming, bounded-state, typed-precedence, timestamp-cache, and source-fence
findings may be cited with the final gate receipts. The frozen source
disposition is **PASS**, and the current isolated seven-command receipt is
**PASS**, including the corrected 132-vector oracle replay and stable-source
verification. The corrected corpus and provenance bindings supersede the
historical provenance-typo manifest; the prior `second.negative_fraction`
discrepancy is resolved and is no longer a semantic or evidence blocker.

The package and current gate receipts cover the dedicated clock error variant,
separate VALUE/date-family domains, negative-date floor behavior, numeric-only
fallback, the selected MINUTE formula, ordered YEARFRAC procedures,
sequence/error precedence, and timestamp-aware cache/source fences.
Performance capture remains pending; the historical production baseline at
commit `6fa3b8af6a` lacked date/time dispatch.
