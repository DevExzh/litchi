# ODS 1.4 date and time test plan

Status: planning only. The [date/time contract](contract.md) is still under
semantic and resource review. This plan defines the evidence that must be
created after that review; it makes no production-support or passing-gate
claim and does not authorize test or Cargo changes in this planning phase.

The batch covers the complete §6.10 family: `DATE`, `DATEDIF`, `DATEVALUE`,
`DAY`, `DAYS`, `DAYS360`, `EASTERSUNDAY`, `EDATE`, `EOMONTH`, `HOUR`,
`ISOWEEKNUM`, `MINUTE`, `MONTH`, `NETWORKDAYS`, `NOW`, `SECOND`, `TIME`,
`TIMEVALUE`, `TODAY`, `WEEKDAY`, `WEEKNUM`, `WORKDAY`, `YEAR`, and `YEARFRAC`.
The primary normative inputs are the repository-local ODF 1.4 Part 4 archive
and formula member recorded in `contract.md`; the independent oracle and
native capture must pin those same hashes and the final contract hash.

## Evidence ownership and planned artifacts

Root owns the evaluator integration, public timestamp API, source freeze,
isolated gates, and final acceptance. The semantic test owner will later add
the two focused targets below; neither target is created by this plan:

* `crates/litchi-ods/tests/ods_formula_date_time_evaluation.rs` will cover
  finite semantic results, conversions, exact arity, scalar/matrix behavior,
  and differential checks against the independent model.
* `crates/litchi-ods/tests/ods_formula_date_time_limits.rs` will cover
  resolver reads, sequence streaming, text budgets, work/cancellation,
  reservations, source fences, and typed-failure precedence.

The future oracle should live under this evidence directory as an independent
implementation and produce a bounded, contract-hashed JSON corpus. The future
native fixture should retain its FODS input, recalculated output, extracted
results, provenance, and reproduction script in the same directory. The
existing [lookup oracle/native plan](../ods-formula-lookups/oracle-native-plan.md)
is the evidence-layout reference; it is not an oracle for date semantics.
Native spreadsheet output can corroborate ordinary vectors and expose host
differences, but it cannot replace the local ODF profile or the independent
expected values.

No test implementation, oracle generation, native capture, or Cargo command is
part of this document's preparation.

## Arity and function coverage

Every row must have scalar and value-evaluator observations, exact wrong-arity
cases, required-argument missing cases, and the documented optional-slot
cases. A wrong arity is a formula `#VALUE!`; an omitted optional argument is
different from a supplied Empty worksheet value and from a syntactically
empty slot. Formula Errors propagate in source order. Typed resolver,
cancellation, source, allocation, and resource failures remain typed
evaluation failures.

| Function | Exact arity | Required semantic vectors | Formula-error and domain vectors |
| --- | ---: | --- | --- |
| `DATE` | 3 | Truncation toward zero; month/day rollover; leap and non-leap years; serial result and profile bounds | Argument Error precedence; zero/negative month or day, non-finite input, checked rollover, and out-of-domain result are `#NUM!`; malformed conversion is `#VALUE!` |
| `DATEDIF` | 3 | `Y`, `M`, `D`, `MD`, `YM`, `YD`; completed-unit boundaries; end-of-month and leap-day cases | Unknown/empty format is `#VALUE!`; reversed dates and date-domain failures are `#NUM!`; argument Errors retain source order |
| `DATEVALUE` | 1 | ISO date, ISO datetime, en_US numeric forms, English month names, two-digit pivot, integer extraction | Malformed text is `#VALUE!`; admitted non-finite/out-of-domain numeric fallback is `#NUM!`; one Error and wrong arity are checked |
| `DAY` | 1 | DateParam conversion and fractional-time discard; month boundaries | Malformed text is `#VALUE!`; serial domain failure is `#NUM!`; reference-cell Error propagates |
| `DAYS` | 2 | End minus start with retained numeric fractions; mixed DateParams and reversed intervals | Conversion and argument Errors precede arithmetic; non-finite/out-of-domain values are `#NUM!` |
| `DAYS360` | 2–3 | US/NASD default and European Method `TRUE`; 31st and February rules; reversed US interval | Invalid Method conversion is `#VALUE!`; date/domain/overflow failures are `#NUM!`; omitted and explicit-empty Method are compared |
| `EASTERSUNDAY` | 0–1 | Explicit years 1583, leap-century boundaries, 9956; omitted year using injected timestamp and next-Easter selection | No timestamp is typed `Unsupported(CalculationClock)`; unsupported year/domain is `#NUM!`; Error and wrong arity cases are retained |
| `EDATE` | 2 | Positive, negative, and zero month offsets; end-of-month clamping; fractional source truncation | Bad conversion and argument Error precedence; checked month/date overflow is `#NUM!` |
| `EOMONTH` | 2 | Month shift and target-month end for leap/non-leap months | Same conversion, Error, and checked-domain cases as `EDATE`; wrong arity is `#VALUE!` |
| `HOUR` | 1 | DateTime and negative/multi-day Time fractions; 0 and 23 boundaries | Malformed Text is `#VALUE!`; non-finite/out-of-domain Time is `#NUM!`; reference Error propagates |
| `ISOWEEKNUM` | 1 | Monday/first-Thursday rule; ISO week-year boundaries and week 52/53 | Date conversion `#VALUE!`/`#NUM!` and source Error precedence |
| `MINUTE` | 1 | Rounded-total-second profile; boundaries around minute and day rollover | Malformed/non-finite inputs and Error precedence; source conflict with the alternate HTML formula is pinned in review evidence |
| `MONTH` | 1 | DateParam conversion, fractional-time discard, all month boundaries | Same date conversion/domain and Error cases as `DAY` |
| `NETWORKDAYS` | 2–4 | Inclusive forward/reverse counts; default and custom seven-element workweek; holidays, duplicates, empty holiday slot, all-off workweek | Malformed sequences are `#VALUE!`; holiday/date domain and retained sequence Errors are tested; typed failures supersede formula Errors |
| `NOW` | 0 | Exact injected DateTime serial, including fraction; repeated same-snapshot evaluation | Wrong arity is `#VALUE!`; absent timestamp is typed `Unsupported(CalculationClock)`, never a fabricated value or formula Error |
| `SECOND` | 1 | Same rounded-total-second state as `MINUTE`; 0/59 and day rollover | Malformed/non-finite inputs and Error precedence; leap second is `#VALUE!` |
| `TIME` | 3 | Direct finite fractional formula; negative, multi-day, and boundary components; checked arithmetic | Non-finite or overflowing result is `#NUM!`; bad conversion is `#VALUE!`; argument Errors precede arithmetic |
| `TIMEVALUE` | 1 | Clock text, combined datetime fraction, numeric fallback, fractional seconds | Malformed clock/24:00/minute/second/leap-second is `#VALUE!`; non-finite fallback is `#NUM!` |
| `TODAY` | 0 | Integer date portion of injected timestamp; same snapshot across calls | Wrong arity is `#VALUE!`; absent timestamp is typed `Unsupported(CalculationClock)` |
| `WEEKDAY` | 1–2 | Default Type 1; Types 2, 3, 11–17; explicit-empty optional slot; Gregorian boundaries | Invalid Type is `#NUM!`; conversion and argument Errors retain precedence; wrong arity is `#VALUE!` |
| `WEEKNUM` | 1–2 | Modes 1, 2, 11–17, 21, 150; year boundaries and ISO aliases | Invalid or non-integral mode refusal is checked against the frozen conversion profile; ordinary invalid mode is `#NUM!`; Error and arity cases are required |
| `WORKDAY` | 2–4 | Positive/negative/zero offset; exact zero-offset fractional preservation; custom workweek and holidays; boundary stepping | Malformed sequences are `#VALUE!`; non-finite/overflow/no-workday stepping and date-domain failures are `#NUM!`; typed failures supersede retained Errors |
| `YEAR` | 1 | DateParam conversion, fractional discard, two-digit text profile, years 1–9999 | Malformed text is `#VALUE!`; serial/domain failure is `#NUM!`; reference Error propagates |
| `YEARFRAC` | 2–3 | Bases 0–4; 30/360 February and 31st rules; actual/leap spans; ordered reversed dates | Invalid basis is `#NUM!`; conversion and argument Errors precede calculation; wrong arity is `#VALUE!` |

The exact-arity matrix must include too few and too many arguments for every
name, including `NOW(1)`, `TODAY(1)`, `EASTERSUNDAY(1;2)`, and both optional
forms of `WEEKDAY`, `WEEKNUM`, `NETWORKDAYS`, `WORKDAY`, `DAYS360`, and
`YEARFRAC`. Required missing arguments such as `DATE(2020;;1)` are not
zero/default arguments. Optional omitted slots, explicit Empty cells, and
syntactically empty slots are separate rows wherever the contract admits all
three.

## Common semantic corpus

The fixture and independent oracle must share only the declared profile data,
not implementation code. The core corpus includes:

* Epoch `1899-12-30`, proleptic Gregorian conversion, no synthetic
  `1900-02-29`, serial `-693,593` for `0001-01-01`, and the half-open upper
  DateTime boundary below serial `2,958,466`.
* `1900`, `1904`, leap-century, `9999`, and out-of-domain dates; fractional
  DateTime serials; negative finite Time values; and checked integer/month/day
  overflow.
* ISO dates and datetimes, en_US slash and hyphen forms, English short/long
  month names, ASCII surrounding whitespace, the 1930 two-digit-year pivot,
  clock-only forms, malformed fields, 24:00, leap seconds, non-finite text,
  and numeric fallback. Input text must include borrowed resolver strings and
  long strings near the text budget.
* Number, Logical, Text, Empty, Missing, direct single-cell Reference,
  ReferenceList, Array, Complex, and formula Error conversions. The expected
  result records `#VALUE!`, `#NUM!`, or the original formula Error separately.
* Formula Error arguments before and after ordinary invalid controls; one
  formula Error in a holiday/workweek sequence followed by a valid cell; and
  typed provider failure after a retained Error. These cases prove that
  formula-level `IFERROR`/`IFNA` cannot catch typed failures.

The independent expected values must use exact serials and error kinds. They
must not compare formatted display strings, host locale output, or the local
evaluator's date helper.

## Timestamp API and volatile functions

Before semantic vectors are captured, root must expose and document the
timestamp seam described by the contract:

* `CalculationTimestamp::from_serial` validates finite profile-domain input,
  canonicalizes negative zero, preserves valid fractional bits, and has an
  optional civil constructor. Invalid construction returns a typed validation
  error without panicking.
* `EvaluationOptions::with_calculation_timestamp` stores the snapshot and an
  accessor returns it. `EvaluationContext` and the value `Context` forward the
  same immutable snapshot. The public value is `Copy`/`Eq` as required by the
  cache identity.
* `NOW()` returns the exact supplied serial; `TODAY()` returns its date part;
  `EASTERSUNDAY()` without a Year uses that date and the next eligible Easter.
  Repeated scalar, value, and matrix calls with one snapshot must agree.
* Calls without a timestamp return typed `Unsupported(CalculationClock)`.
  They must not read a system clock, timezone, workbook metadata, environment,
  or random source, and `IFERROR`/`IFNA` must not convert the typed refusal.
* Changing only the timestamp invalidates any volatile demand-cache identity;
  a cache hit under one timestamp must never serve another snapshot. Source
  and cancellation fences still run around a cached volatile result.

The timestamp tests need an explicit deterministic `Position` and no ambient
clock. Native captures of `NOW`/`TODAY` are host observations only and must be
excluded from deterministic semantic equality unless the native fixture can
inject the same timestamp.

## Scalar, resolver, and matrix differential coverage

Each context-free formula is evaluated through the scalar evaluator and the
resolver-backed value evaluator with an equivalent context. The comparison
normalizes only the public result subtype and retains exact serial/error
identity. The value evaluator adds a resolver fixture with known sheet order,
extents, Empty/Text/Logical/Number/Error cells, read order, metadata calls, and
source versions.

The matrix schedule must cover:

* scalar date functions under a projected `IF` with a real 2-D condition,
  including selected and unselected references and formula Errors;
* elementwise date/offset/mode arguments with unequal but valid shapes,
  complete output shape, and per-cell formula Errors;
* `DATE`, `TIME`, `WEEKDAY`, `WEEKNUM`, `YEARFRAC`, and date arithmetic in
  array-valued arguments without collapsing the first coordinate into every
  result;
* direct reference arguments, derived references from `INDEX`/`OFFSET`/
  `INDIRECT`, and ReferenceList/3-D refusal before cell reads where the
  pseudotype forbids it;
* `NETWORKDAYS` and `WORKDAY` projected inside `IF`: holiday and workweek
  sequence arguments remain complete descriptors and are consumed in full for
  each selected scalar result. They must not be implicitly intersected with
  the projected output coordinate;
* nested computed scalar arguments containing `MUNIT` (or an equivalent
  position-sensitive matrix producer). Each output coordinate must use its own
  MUNIT scalar value. The demand cache may not reuse a result across positions
  merely because the outer date function is a reducer; and
* typed failure in one matrix coordinate: no partial array is published, and
  no unselected coordinate is read after the failure boundary.

`TIME`/`TIMEVALUE` tests must keep the TimeParam domain separate from the
DateParam domain: every finite numeric time magnitude is admitted by the
selected profile, including negative and multi-day values. Only checked
arithmetic overflow or a non-finite result is `#NUM!`; a date-range refusal
must not be borrowed from DateParam conversion.

For sequence reducers, a matrix shape probe may inspect descriptors and
metadata, but it may not turn a complete holiday/workweek reference into a
single projected cell. A direct `ReferenceList` is refused before cell reads
when the contract says it is not a DateSequence or LogicalSequence.

## Sequence, text, resource, and failure tests

The limits target must use a resolver that can fail at metadata, at a chosen
cell read, after a chosen successful read, on cancellation, and on source
version change. Each test records read order, successful reads, work units,
text bytes, storage reservations, and the final typed result.

Reference-backed holidays and workweeks are streamed in descriptor/sheet/
row-major order. For every cell the evaluator must charge work, check the
reference-cell limit and cancellation, call `read_cell`, run borrowed
`read_to_element` conversion, and check cancellation again. It must retain
only fixed-size reducer state plus bounded holiday/workweek metadata; it must
not materialize the range's raw cells.

The sequence cases must prove all of the following:

1. A formula Error is retained while later sequence members are scanned. A
   later typed resolver, cancellation, source, allocation, or resource failure
   supersedes the retained formula Error.
2. Offset zero does not skip holiday/workweek validation or reads. An all-off
   workweek is valid for `NETWORKDAYS`, while nonzero `WORKDAY` returns its
   documented `#NUM!` after the complete sequence has been consumed.
3. An empty holiday slot means no holidays only where the contract explicitly
   permits it; an explicit Empty worksheet cell remains a value. The fourth
   Workdays argument can be supplied through an empty third slot.
4. Text and Empty cells in a referenced DateSequence follow the selected skip
   rules; malformed inline workweek Text is `#VALUE!`, and the seven-element
   length requirement is checked without flattening an unbounded source.
5. Duplicate holidays, out-of-domain holiday numbers, formula Errors, and
   later valid members preserve the selected error and count semantics.
6. Large date intervals periodically charge work and cancellation even with
   no resolver references. Checked day stepping, month arithmetic, holiday
   count, seven-day cycle state, and output-array growth cannot wrap or loop
   without a budget checkpoint.
7. Long date/clock Text charges borrowed input bytes and uses bounded parser
   scratch. The resolver's text storage remains borrowed through conversion;
   no full cell string is cloned merely to parse it.
8. `max_reference_cells`, `max_array_cells`, stack/scratch, holiday metadata,
   and work budgets fail before exceeding their reservations. A failed
   conversion releases its storage before the enclosing budget lease is
   dropped. The test retains before/after memory and reservation counts.
9. Source-version and final-cancellation fences surround complete evaluation
   and publication. A change after the last sequence read produces a typed
   source/cancellation failure and no partial result.

Zero-cell-read refusals must include wrong static sequence shape,
ReferenceList where not admitted, statically malformed inline workweek shape,
malformed static Text, no timestamp, and out-of-bounds checked geometry.
Reference-backed workweek/holiday lengths and members are eager sequence
inputs: the evaluator must read and scan them before publishing a formula
error or an ordinary result, so an invalid referenced member is not a reason
to assert zero reads. Likewise, a scalar date/mode/basis refusal can be
read-free when all relevant arguments are scalar, but supplied reference
sequences still follow complete-consumption and typed-failure precedence.
Metadata calls needed to establish a sheet extent or canonical name are
allowed only when the contract requires them and are counted separately from
cell reads.

## Existing `VALUE` compatibility

The date/time batch must add no new `VALUE` behavior. Before and after the
date implementation, retain a compatibility set for the existing inspection
`VALUE` parser covering its established 1899-12-30 through 9999-12-31 domain,
negative computed serial refusal, fixed numeric/date/time/datetime forms,
malformed text, and formula-error precedence. The date family may share
calendar parsing helpers only when the helper receives an explicit caller
domain; `DATEVALUE`/DateParam may use the wider profile while `VALUE` retains
its current boundary.

The compatibility receipt must include the existing VALUE-focused tests and a
hash or manifest proving that unrelated VALUE expectations were not rewritten
to make date/time vectors pass. A date/time test that changes `VALUE` output is
a contract/API issue for root, not a date-time acceptance result.

## Independent oracle and native comparison

The oracle corpus should include at least one observation per requirement row,
all 24 names, every exact arity, scalar and matrix mode, explicit position,
timestamp option, expected serial/error subtype, and expected read trace for
resolver cases. It must model sequence consumption, date profile, parser
forms, workweek/holiday order, typed failure precedence, and cache identity
without importing Rust code or calling a spreadsheet engine.

Native evidence should use a bounded local FODS fixture with ordinary dates,
weekends, holidays, month ends, ISO boundaries, and text parsing forms. The
fixture should record unsupported functions, host serial/epoch behavior,
locale differences, and volatile NOW/TODAY behavior as divergences. Native
results are a separate comparison column and cannot supply expected values for
profile decisions such as the 1900 boundary, MINUTE's conflicting formula,
TIME's fractional policy, reversed `DATEDIF`, or the timestamp refusal.

Each retained oracle/native receipt must bind its formula, contract hash,
normative-source hashes, profile, source fixture hash, and result hash. A
missing observation, placeholder JSON object, or native-only expectation is a
failed evidence binding.

## Review gates before implementation and capture

Root and the semantic reviewer must resolve the contract choices that affect
expected values before the focused files or oracle are written: date-domain
width versus `VALUE`, timestamp API names and absent-clock error, DateParam and
TimeParam Logical conversion, `MINUTE`'s conflicting source formula, direct
`TIME` versus optional integer preprocessing, reversed `DATEDIF`, reversed
`NETWORKDAYS`, zero/nonzero `WORKDAY`, ordered `YEARFRAC`, exact non-integral
`WEEKNUM` mode conversion, and complete sequence behavior under projected
matrix evaluation.

After that freeze, the owner will create the two focused test targets, the
independent oracle/native artifacts, and the evidence manifest. The final
validation must run locked isolated ODS tests, strict Clippy/rustdoc/format
checks, resource-boundary checks, and the retained oracle/native comparisons;
this plan records no such run and makes no PASS claim.
