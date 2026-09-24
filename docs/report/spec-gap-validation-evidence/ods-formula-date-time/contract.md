# ODF 1.4 date and time function contract

Status: draft implementation contract; independent semantic and resource
review pending. This document defines the complete ODF 1.4 Part 4 §6.10
family. It does not claim that the production evaluator or its tests already
implement this family.

The batch contains all twenty-four functions in §6.10:

DATE, DATEDIF, DATEVALUE, DAY, DAYS, DAYS360, EASTERSUNDAY, EDATE, EOMONTH,
HOUR, ISOWEEKNUM, MINUTE, MONTH, NETWORKDAYS, NOW, SECOND, TIME, TIMEVALUE,
TODAY, WEEKDAY, WEEKNUM, WORKDAY, YEAR, and YEARFRAC.

The contract keeps date and time values as the evaluator's finite Number
subtypes. A date is an integer serial, a time is a fraction of a day, and a
datetime is their sum. The date kernel must be shared by scalar evaluation and
the resolver-backed value evaluator; a second date implementation in the
matrix bridge is not conforming.

## Authority and accepted design boundaries

The primary normative input is the repository-local ODF distribution:

| Source | SHA-256 |
| --- | --- |
| archive 3rdparty/specs/OpenDocument-v1.4-os.zip | 9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4 |
| member part4-formula/OpenDocument-v1.4-os-part4-formula.html | ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1 |

The function entries are §§6.10.2–6.10.25. General value and conversion rules
come from §§3.2.3, 3.3, 3.4, 4.3, 4.11.3–4.11.4, 4.11.7,
4.11.12, 5.6, 5.12, 6.1, 6.2, 6.3.5–6.3.6, and 6.3.9/6.3.13.

The local contract follows accepted ADRs 0001, 0004, 0005, 0006, 0008, and
0023:

* ordinary formula APIs are typed, panic-free semantic APIs;
* arguments, arrays, reference metadata, scratch, and sequence scans stay
  under explicit finite budgets;
* formula Errors remain formula values, while cancellation, source changes,
  provider failures, allocation failures, and resource limits remain typed
  evaluation failures;
* source-version and cancellation fences surround resolver-backed evaluation;
* no locale, filesystem, wall clock, process identity, or random source is
  consulted implicitly; and
* this family belongs to the ODS formula evaluator and shares its existing
  read-only source boundary.

The PDF table of contents has a stale subsection label in this area. The
normative HTML body and Appendix A identify EASTERSUNDAY as §6.10.8 and EDATE
as §6.10.9; this contract follows the body.

At the current source baseline, the tokenizer knows these names but the ODS
evaluator has no date/time dispatch. Falling through to generic
Unsupported(Function) is therefore a gap, not an implementation of this
contract.

## Function signatures and exact arity

The table records the Part 4 pseudotype, result subtype, and selected profile
for conversion and domain failures. Optional arguments are omitted AST
arguments, not blank worksheet cells. A syntactically empty optional slot is
handled only where this contract says so.

| Function | ODF syntax | Exact arity | Result | Selected conversion/domain rule |
| --- | --- | ---: | --- | --- |
| DATE | DATE(Integer Year; Integer Month; Integer Day) | 3 | Date | Integer conversion truncates toward zero; positive month/day values are normalized with checked rollover |
| DATEDIF | DATEDIF(DateParam StartDate; DateParam EndDate; Text Format) | 3 | Number | Date-only values; Format is one of Y, M, D, MD, YM, YD, case-insensitive |
| DATEVALUE | DATEVALUE(Text D) | 1 | Date | Fixed profile date parser; combined datetime returns its integer date |
| DAY | DAY(DateParam D) | 1 | Number | DateParam conversion, fractional time discarded |
| DAYS | DAYS(DateParam EndDate; DateParam StartDate) | 2 | Number | EndDate minus StartDate; numeric fractions are retained |
| DAYS360 | DAYS360(DateParam StartDate; DateParam EndDate [; Logical Method = FALSE]) | 2–3 | Number | Optional Method false selects US/NASD; true selects European |
| EASTERSUNDAY | EASTERSUNDAY([Integer Year]) | 0–1 | Date | Explicit year is checked; omitted year requires the injected calculation timestamp |
| EDATE | EDATE(DateParam StartDate; Number MonthAdd) | 2 | Number | Floor date and truncate month count; clamp the target day to month end |
| EOMONTH | EOMONTH(DateParam StartDate; Integer MonthAdd) | 2 | Number | Floor date and truncate month count; return target month end |
| HOUR | HOUR(TimeParam T) | 1 | Number | 24-hour component of the day fraction |
| ISOWEEKNUM | ISOWEEKNUM(DateParam D) | 1 | Number | ISO 8601 Monday/first-Thursday week |
| MINUTE | MINUTE(TimeParam T) | 1 | Number | Minute component after the selected total-second rounding |
| MONTH | MONTH(DateParam Date) | 1 | Number | DateParam conversion, fractional time discarded |
| NETWORKDAYS | NETWORKDAYS(DateParam Date1; DateParam Date2 [; [DateSequence Holidays] [; LogicalSequence Workdays]]) | 2–4 | Number | Inclusive workday count; default weekend and holiday rules below |
| NOW | NOW() | 0 | DateTime | Requires an explicit deterministic calculation timestamp |
| SECOND | SECOND(TimeParam T) | 1 | Number | Nearest-second component, 0 through 59 |
| TIME | TIME(Number Hours; Number Minutes; Number Seconds) | 3 | Time | Finite values, direct fractional formula, checked total seconds |
| TIMEVALUE | TIMEVALUE(Text T) | 1 | Time | Fixed profile clock parser; combined datetime returns its fraction |
| TODAY | TODAY() | 0 | Date | Requires an explicit deterministic calculation timestamp |
| WEEKDAY | WEEKDAY(DateParam D [; Integer Type = 1]) | 1–2 | Number | Type must be 1, 2, 3, or 11–17; omitted/empty optional slot uses 1 |
| WEEKNUM | WEEKNUM(DateParam D [; Number Mode = 1]) | 1–2 | Number | Mode must be 1, 2, 11–17, 21, or 150; omitted/empty uses 1 |
| WORKDAY | WORKDAY(DateParam Date; Number Offset [; [DateSequence Holidays] [; LogicalSequence Workdays]]) | 2–4 | DateTime | Checked workday stepping; zero offset returns the exact input serial |
| YEAR | YEAR(DateParam D) | 1 | Number | DateParam conversion, fixed two-digit-year profile for text |
| YEARFRAC | YEARFRAC(DateParam StartDate; DateParam EndDate [; Basis B = 0]) | 2–3 | Number | Basis is 0–4 after integer conversion; day-count table below |

An invalid arity produces formula #VALUE! after the evaluator has consumed the
supplied argument slots according to the existing function boundary. A
missing required argument is not an Empty worksheet value. A formula Error
produced by an argument propagates in source order before conversion. A typed
evaluation failure is never converted into a formula Error and is never caught
by IFERROR or IFNA.

## Deterministic date profile

OpenFormula leaves the epoch and some locale behavior to the host. This
evaluator selects one documented profile rather than consulting document or
process state:

| Profile member | Selected value |
| --- | --- |
| calendar | Proleptic Gregorian |
| epoch | 1899-12-30 |
| synthetic 1900-02-29 | Never present |
| minimum calendar date for this family | 0001-01-01 |
| minimum serial date for this family | -693,593, representing 0001-01-01 |
| maximum calendar date | 9999-12-31 |
| maximum date-part serial | 2,958,465, representing 9999-12-31 |
| maximum DateTime serial | less than 2,958,466, including a final-day fraction |
| locale | en_US |
| decimal separator | period |
| group separator | comma |
| month names | English short and long names |
| two-digit year pivot | 1930 |
| timezone | None; the supplied timestamp is already a serial in this profile |

The profile supports the ODF-required 1904-01-01 through 9999-12-31 range and
the ODF-recommended 1899-12-30 through 9999-12-31 range. It also admits the
proleptic Gregorian extension back to 0001-01-01 because EASTERSUNDAY requires
an explicit Year as early as 1583 and §6.10.8 returns DATE(Year;month;day).
A non-finite Number, a serial below -693,593, or a serial at or above
2,958,466 is a numeric-domain failure (#NUM!) when a DateParam is required.
Fractional serials are valid DateTime values throughout that half-open range;
date-only functions floor them to the represented day. An integer date result
therefore never exceeds 2,958,465, while a final-day DateTime can be any
finite serial below 2,958,466.

The §6.10.2 DATE entry lists 1904–9956 as its ordinary Year constraint while
also defining month/day rollover. This profile intentionally supports the
wider checked serial domain above, including dates in years 1–1903 and
9957–9999. Implementations may expose a stricter profile later, but they must
not silently use a host epoch or the fictitious 1900-02-29.

The existing value-inspection VALUE contract remains narrower: VALUE accepts
its established 1899-12-30 through 9999-12-31 serial domain and continues to
refuse computed negative serials. The date/time implementation may share
calendar parsing and civil-date helpers with VALUE, but the helpers must take
an explicit caller domain. DATEVALUE and DateParam use this family domain;
VALUE retains its existing domain until a separate contract changes it.

### Explicit calculation timestamp

NOW(), TODAY(), and EASTERSUNDAY() without a Year need the current date. They
must work when the caller supplies a deterministic timestamp through the
evaluation options. The proposed API is:

~~~
let stamp = CalculationTimestamp::from_serial(46000.5)?;
let options = EvaluationOptions::default()
    .with_calculation_timestamp(stamp);
let context = EvaluationContext::with_options(&execution, options);
~~~

CalculationTimestamp is a validated, Copy and Eq value whose public meaning is
a finite profile serial with a date part in -693,593..=2,958,465 and a time
part in 0..1 day. To preserve arbitrary valid fractional seconds while
retaining EvaluationOptions' Copy/Eq shape, its private representation may be
an exact validated f64 bit pattern: from_serial rejects non-finite and
out-of-domain values, canonicalizes negative zero to positive zero, and stores
the remaining bits in a u64 field. Its serial accessor reconstructs the
validated f64. It should offer both from_serial and a civil constructor such
as from_ymd_hms, returning a typed validation error. A raw unvalidated f64
must not enter EvaluationOptions.

EvaluationOptions::with_calculation_timestamp stores the snapshot; an
accessor returns Option<CalculationTimestamp>. EvaluationContext and the value
Context forward the same snapshot. No function reads a system clock, timezone,
environment variable, or workbook metadata implicitly. If a volatile function
is called without the option, the result is typed
Unsupported(CalculationClock) rather than a fabricated serial or a formula
Error. This absent-capability case is a valid boundary behavior; an
unconditional NOW/TODAY refusal even when the option is present is not an
implementation of this contract.

NOW returns the supplied serial, including its fractional day. TODAY returns
the integer date component. EASTERSUNDAY with an omitted Year compares the
supplied current date with Easter in the current and following year and
returns the first date not earlier than TODAY.

## Common conversions and errors

### DateParam

For an expected DateParam:

* Number is accepted as a finite serial in the profile domain.
* Text is passed to DATEVALUE. If the date parser rejects it, this profile
  makes the permitted fallback attempt through the fixed numeric-only VALUE
  grammar. Numeric-only means the complete fixed numeric branch, including
  decimal, exponent, percent, currency, valid grouping, and mixed-fraction
  forms, with date/time dispatch disabled. A malformed text returns #VALUE!;
  a syntactically numeric but non-finite or out-of-domain result returns
  #NUM!.
* Logical is converted to 1 or 0. This is the selected implementation-defined
  choice allowed by §6.3.15.
* A single-cell Reference is converted to its scalar value; an Empty cell
  converts to 0. A formula Error in the cell propagates.
* A direct ReferenceList is not a scalar DateParam and returns #VALUE! before
  resolver reads. Matrix/reference scheduling must preserve the ordinary
  scalar intersection or elementwise rules rather than flattening a list.

Date-only consumers remove the Time subtype by flooring the admitted serial,
then operate on the resulting civil date. For the nonnegative serials in the
ODF-required range this is exactly the normative “truncate date” operation
used by DAYS360 and the §4.11.7 procedures. For the explicitly supported
negative extension, the serial is represented as a whole civil-date serial
plus a nonnegative subday fraction, so floor is the corresponding civil-date
decomposition. This is distinct from Integer and WORKDAY Offset conversion,
which remain toward-zero truncations.

### TimeParam

For an expected TimeParam:

* Number is accepted when finite. Its fractional day is the time component;
  DateTime serials are therefore valid.
* Text is passed to TIMEVALUE, with the same permitted numeric-only VALUE
  fallback for a numeric time serial.
* Logical maps to 1 or 0 under the selected implementation-defined choice.
* A Reference is converted to scalar and an Empty cell becomes 0.
* Formula Errors and typed evaluation failures preserve the common rules.

The profile does not invent a timezone, normalize a timestamp through UTC, or
accept leap seconds. TIMEVALUE clock text is restricted to 00:00 through
23:59:59 plus a finite fractional second. TIME itself may produce a negative
or multi-day Number because §6.10.18 permits any finite Number for Hours,
Minutes, and Seconds; downstream component extraction uses the day fraction
defined by the function.

### Other parameter types

* Number accepts finite Number values and the existing Logical-to-0/1 and
  fixed Text-to-Number conversions. Empty worksheet values convert to 0 where
  the existing Number bridge permits it.
* Integer converts a finite Number by truncation toward zero and rejects
  non-finite or unrepresentable results.
* Logical uses the existing distinguished Logical conversion. Number zero is
  false and nonzero is true; malformed Text is #VALUE!. Omitted optional
  Logical slots use their documented defaults. An explicit Empty cell is a
  value and follows the ordinary Empty-to-Logical bridge.
* Text uses the deterministic scalar Text conversion. Number and Logical are
  formatted through the existing profile; an Empty cell is an empty Text only
  where the function's Text conversion admits it. A formula Error propagates.

Domain violations return the function-level formula error stated below,
usually #NUM! for an out-of-range date/mode/basis and #VALUE! for malformed
text, wrong pseudotype, or an invalid format code. A final non-finite
calculation is #NUM!. No unchecked integer cast, saturating date arithmetic,
or wrapping serial calculation is permitted.

## Fixed profile parsing

DATEVALUE and TIMEVALUE use the existing bounded VALUE parser's calendar
helpers and do not call a process locale. The accepted profile is:

* ISO date YYYY-MM-DD, with four-digit year and two-digit month/day;
* ISO datetime with a space or T separator and HH:MM, HH:MM:SS, or
  fractional seconds;
* en_US numeric dates using M/D/YYYY, M/D/YY, and the corresponding
  four-digit-year hyphen form;
* English month names in MMM D, YYYY, MMMM D, YYYY, D MMM YYYY, and
  D MMMM YYYY forms; and
* clock text HH:MM, HH:MM:SS, or HH:MM:SS.f....

After the function-specific date or clock grammar fails, the permitted
numeric-only fallback uses the existing VALUE numeric forms. It accepts an
optional leading sign, an optional dollar sign, decimal digits with an
optional decimal fraction, scientific exponents, and one trailing percent
sign. The en_US grouping form accepts comma groups only when the first group
has the permitted leading width and every later group has exactly three
digits; malformed grouping is #VALUE!. A leading parenthesized numeric or
currency form is the existing negative form. A decimal fraction requires
digits after the decimal point, so 123.5 and .5 are accepted while 123. is
malformed. It also accepts the existing fraction forms: an optional sign, a
numerator, and a one- or two-digit nonzero denominator, with an optional
integer part followed by one required space for a mixed fraction. Slashes are
admitted only in those valid simple or mixed-fraction forms, and date-like
forms with two date separators are not admitted by this fallback. Long
grouped inputs use the same bounded normalization scratch and charging as
VALUE.

ASCII surrounding whitespace is accepted. Month/day fields are validated
against the proleptic Gregorian calendar. Two-digit years use the fixed
1930 pivot: 30–99 map to 1930–1999 and 00–29 map to 2000–2029. A date before
0001-01-01 or after 9999-12-31 is #NUM!. A malformed date or clock,
24:00, an out-of-range minute/second, and a leap second are #VALUE!.

DATEVALUE returns floor(date-plus-time). When its numeric-only fallback
produces a finite serial, DATEVALUE floors it to an integer date serial.
TIMEVALUE returns only the fractional day for a parsed clock or datetime; when
its numeric-only fallback succeeds, it preserves the raw finite numeric serial
as the Time result, including a value greater than one. A date-only string
passed to TIMEVALUE is not a clock; after the permitted fallback it either
yields a numeric time serial or returns #VALUE!. A clock-only string passed to
DATEVALUE is handled analogously. For executable profile cases,
DATEVALUE("123") returns 123, DATEVALUE("123.5") returns 123,
DATEVALUE("1/4") returns 0, TIMEVALUE("0.5") returns 0.5,
TIMEVALUE("2.5") returns 2.5, and TIMEVALUE("1/4") returns 0.25;
DATEVALUE("12:00") returns #VALUE!, and TIMEVALUE("2020-01-01") returns
#VALUE!. This records the implementation-defined fallback permitted by
§6.10.4 and §6.10.19 without letting the full VALUE date/time grammar bypass
the function's own parser.

The parser charges the borrowed input bytes and checks cancellation at its
established scan checkpoints. Long grouped or Unicode input uses caller-owned
bounded scratch with a checked reservation; no date parser clones an entire
reference cell or hides an uncharged allocation.

## Function semantics

### DATE

DATE truncates Year, Month, and Day toward zero. The selected profile accepts
positive checked Month and Day values, including Month greater than 12 and
Day beyond the length of the normalized month, and applies the specified
Gregorian rollover. Zero or negative Month/Day, an unrepresentable
intermediate, and a final date outside the profile domain return #NUM!. The
result is an integer serial. A formula Error in any argument propagates
before conversion.

For example, DATE(2020;13;1) is 2021-01-01 and DATE(2020;2;30) is
2020-03-01 under the proleptic Gregorian profile. The extended year 1–1903
and 9957–9999 support described above applies when the normalized result remains
within the serial domain.

### DATEDIF

DATEDIF first converts and floors both DateParams. The selected profile
returns #NUM! when EndDate is earlier than StartDate; it does not silently
swap the arguments. Format is trimmed ASCII Text and matched
case-insensitively against:

| Format | Result |
| --- | --- |
| Y | completed calendar years |
| M | completed calendar months |
| D | elapsed whole days |
| MD | day difference after ignoring years and months |
| YM | month difference after ignoring years |
| YD | day difference after ignoring years |

Y is the largest nonnegative number of whole calendar years whose
anniversary, with February 29 clamped to the last day of February when
needed, is not after EndDate. M is the analogous complete-month count with
the same end-of-month clamp. D is the serial day difference.

For MD, subtract the StartDate day from the EndDate day; if negative, add
the length of the month preceding EndDate's month. For YM, first compute the
signed month difference, subtract one when EndDate's day is before
StartDate's day, and only then apply Euclidean modulo 12. Thus
DATEDIF(2020-01-31;2021-01-01;"YM") is 11, while
DATEDIF(2020-01-31;2020-02-28;"YM") is 0. M intentionally differs: its
complete-month calculation clamps the January 31 anniversary to February 28,
so the latter interval has M equal to 1. For YD, place StartDate's month/day
in EndDate's year, clamping February 29 to February 28; if that anniversary
is after EndDate, place it in the preceding year, then return the elapsed
days to EndDate. These definitions preserve the complete-unit behavior while
making the otherwise easy-to-misread MD/YM/YD cases executable.

An empty or unknown Format returns #VALUE!. A format Text error or date
conversion error follows the common rules.

### DATEVALUE

DATEVALUE converts one Text argument using the fixed parser. ISO dates are
always accepted independent of locale. A combined date and time returns the
integer date serial. A parser failure may use the fixed VALUE numeric fallback;
if that also fails, the result is #VALUE! or #NUM! according to the common
lexical/domain distinction.

### DAY, MONTH, and YEAR

Each function converts one DateParam, discards its time fraction, and returns
the corresponding Gregorian day, month, or year component. YEAR uses the same
two-digit-year profile when its Text input requires parsing. The functions
return #VALUE! for malformed dates and #NUM! for an admitted but out-of-domain
serial.

### DAYS

DAYS converts EndDate and StartDate in that order and returns EndDate minus
StartDate. When both operands are Numbers, their finite serial fractions are
retained. Text operands use DATEVALUE; mixed operands use their respective
DateParam conversion. No calendar-month normalization is performed.

### DAYS360

DAYS360 first applies the normative date-truncate operation, represented as
floor in this profile, to both DateParams. With Method false (the default),
it uses the US/NASD 30US/360 procedure:

1. If StartDate is the 31st, change its day to 30.
2. Otherwise, if StartDate is the last day of February, change its day to 30.
3. If EndDate is the 31st and the normalized StartDate day is 30, change the
   EndDate day to 30.
4. Return (end.year*360 + end.month*30 + end.day) minus the corresponding
   StartDate expression.

The US procedure does not swap dates, so a reversed interval can be negative.
With Method true, use the European 30E/360 procedure:

1. If StartDate is after EndDate, swap them and retain a sign of -1.
2. Change a 31st StartDate or EndDate to day 30.
3. Return the sign times the same 360-day expression.

February is never changed by the European procedure. Method accepts the
ordinary Logical conversion; omitted and syntactically empty optional slots
select false. A supplied formula Error propagates.

### EASTERSUNDAY

With an explicit Integer Year, require 1583 through 9956 inclusive and apply
the algorithm printed in §6.10.8. Years outside that range or a non-finite
conversion return #NUM!. The result is a profile Date serial.

With no Year argument, require the explicit calculation timestamp. The
timestamp year must be in the supported Easter calculation range 1583–9956.
Compute the Gregorian Easter date for that year and the following year, then
return the smaller date that is not earlier than TODAY. If TODAY is after the
current year's Easter and the following year is outside 1583–9956, return
#NUM!; this includes a timestamp in 9956 after Easter and any timestamp in
9957–9999. Without a timestamp, return typed Unsupported(CalculationClock).
The no-argument form must never read the ambient wall clock.

### EDATE and EOMONTH

EDATE floors StartDate to a date and truncates MonthAdd toward zero, adds the
checked month count, and returns the same day in the target month. If the
target month does not contain that day, clamp to its last day. MonthAdd may
be negative or zero. EOMONTH floors StartDate, performs the same month shift,
ignores the source day, and returns the final day of the target month. Both
return #NUM! for a date or
month arithmetic result outside the profile range and preserve input errors.

### HOUR, MINUTE, and SECOND

HOUR extracts the hour from T minus floor(T), returning 0 through 23. This is
the §6.10.11 formula; using floor also gives a normalized day fraction for a
negative finite Time value.

The HTML body for MINUTE prints both a rounded-total-seconds formula and a
later day-fraction formula. They disagree near second boundaries. This draft
selects the first formula because it is the explicit MOD(ROUND(T*86400))
calculation and is consistent with SECOND:

1. Round the total seconds in T to the nearest second using the evaluator's
   existing half-away-from-zero ROUND profile.
2. Normalize modulo one day.
3. Return (rounded_seconds mod 3600 minus rounded_seconds mod 60) divided by
   60.

SECOND uses the same rounded total-second value and returns rounded_seconds
mod 60, never exposing leap seconds. The source conflict remains an
independent semantic-review item; implementations must not silently mix the
two formulas.

MONTH is listed with these component functions only for dispatch grouping; it
converts DateParam, discards the time fraction, and returns 1 through 12.

### TIME

TIME converts Hours, Minutes, and Seconds to finite Numbers. This profile
selects the direct fractional formula:

~~~
(Hours*3600 + Minutes*60 + Seconds) / 86400
~~~

The source permits evaluators to perform INT() first, but that optional
transformation is not selected here; in particular, negative fractional
inputs are not truncated toward zero. Values outside ordinary clock ranges
remain valid finite Time Numbers. Checked multiplication and final finiteness
checks return #NUM! on overflow.

### TIMEVALUE

TIMEVALUE parses one Text argument using the fixed clock and combined datetime
forms. A combined datetime returns only its fractional day. The permitted
numeric VALUE fallback accepts a finite numeric time serial when no clock form
is present. Malformed clock text, an out-of-range clock component, or an
unconvertible Text returns #VALUE!; a non-finite numeric fallback is #NUM!.

### WEEKDAY

WEEKDAY converts and floors one DateParam. Type is an Integer and defaults
to 1 when omitted or when the optional AST slot is syntactically empty. The
accepted mappings are:

| Type | First weekday | Returned range |
| ---: | --- | --- |
| 1 | Sunday = 1 through Saturday = 7 | 1–7 |
| 2 | Monday = 1 through Sunday = 7 | 1–7 |
| 3 | Monday = 0 through Sunday = 6 | 0–6 |
| 11 | Monday = 1 | 1–7 |
| 12 | Tuesday = 1 | 1–7 |
| 13 | Wednesday = 1 | 1–7 |
| 14 | Thursday = 1 | 1–7 |
| 15 | Friday = 1 | 1–7 |
| 16 | Saturday = 1 | 1–7 |
| 17 | Sunday = 1 | 1–7 |

Any other Type is #NUM!. Dates use the proleptic Gregorian weekday with no
locale or timezone adjustment.

### ISOWEEKNUM and WEEKNUM

ISOWEEKNUM uses ISO 8601: weeks start Monday and week 1 contains the first
Thursday. Dates at the beginning or end of a calendar year can therefore
return week 52 or 53 of the adjacent ISO week-year.

WEEKNUM requires one of the exact modes below. Mode is declared Number, so a
finite non-integer is not truncated into an accepted mode.

| Mode | Week starts | Week 1 |
| ---: | --- | --- |
| 1 | Sunday | week containing January 1 |
| 2 | Monday | week containing January 1 |
| 11 | Monday | week containing January 1 |
| 12 | Tuesday | week containing January 1 |
| 13 | Wednesday | week containing January 1 |
| 14 | Thursday | week containing January 1 |
| 15 | Friday | week containing January 1 |
| 16 | Saturday | week containing January 1 |
| 17 | Sunday | week containing January 1 |
| 21 | Monday | ISO first Thursday |
| 150 | Monday | ISO first Thursday |

Modes 1 and 2 are retained as their historical aliases; 11–17 make the
weekday start explicit. Invalid modes return #NUM!. Mode 150 is the specified
alias for ISO mode 21.

### NETWORKDAYS and WORKDAY

Both functions use DateSequence and LogicalSequence exactly as the ODF
pseudotypes define them:

* A scalar Number, Text, or Logical can form a one-element DateSequence through
  Number conversion. A direct Reference contributes only referenced Number or
  Error cells; Empty and Text cells are skipped. A ReferenceList is not a
  DateSequence under §6.3.9 and is rejected before resolver reads.
* A scalar Number or Logical forms a one-element LogicalSequence. A direct
  Reference contributes Logical or Error cells; because Logical is a
  distinguished type in this profile, referenced Numbers are not silently
  converted. Empty cells are skipped and Text is not admitted.
* Inline arrays follow the existing sequence/matrix traversal rules and stay
  bounded. For a LogicalSequence inline array, Number elements convert to
  false/true by zero/nonzero, Logical elements retain their value, Empty
  elements convert to false, and Text elements are malformed #VALUE! inputs;
  Text is never parsed as a Boolean or Number. The array must contain exactly
  seven elements after traversal. A rectangular reference is streamed in
  row-major order within each sheet, with sheets in descriptor order. The
  selected row-major order is one of the orders permitted by §4.11.12.
  Inline DateSequence elements use the same finite Number conversion and
  date-domain check; an out-of-domain numeric element retains #NUM! while
  later elements are still scanned.
* A formula Error encountered in a sequence is retained as a formula value,
  and the complete admitted sequence is scanned. A typed read, source,
  cancellation, allocation, or resource failure immediately supersedes the
  retained formula Error.

The default Workdays sequence is {1;0;0;0;0;0;1} in Sunday-through-Saturday
order: zero means a workday and nonzero means a non-workday. A supplied
LogicalSequence must contain exactly seven elements. A sequence with no
workday is accepted for NETWORKDAYS, which then returns zero for every
interval; WORKDAY with a nonzero Offset returns #NUM! because it cannot
advance.

Holidays are compared by their floored profile date serial. A referenced
Number outside the date-function serial domain contributes a retained #NUM!
formula result and the scan continues; Empty and Text reference cells are
skipped by DateSequence conversion. Duplicate valid holidays are harmless.
The implementation normalizes each valid holiday to a checked integer day,
charges the retained day buffer, and may sort and deduplicate that bounded
buffer after the complete sequence scan. It must reserve against the
sequence/resource limit before growth and must not allocate an unbounded
range-sized raw-cell vector. Before each inspected reference cell, including
cells later skipped by sequence conversion, the scan charges work, checks the
reference-cell limit and cancellation, reads through the resolver, and checks
cancellation again.

All supplied sequence arguments are consumed and validated even when Offset
is zero or a workweek has no workdays. This preserves eager argument and
sequence semantics: malformed workweek elements, retained formula Errors,
all-off workweeks, and offset-zero results do not suppress later typed read,
source, cancellation, allocation, or resource failures. After the complete
scan, a retained formula Error is published before an ordinary formula result
according to source order; a typed failure always supersedes it. An all-off
workweek is valid for NETWORKDAYS and returns zero, and is #NUM! for WORKDAY
when Offset is nonzero. Offset zero still returns its Date value after the
optional sequences have been consumed.

NETWORKDAYS counts workdays inclusively between Date1 and Date2. A forward
interval includes both endpoints; an interval whose first date is after its
second is evaluated in reverse and returns the negative of the corresponding
forward count. Equal workday endpoints return 1, and equal non-workday
endpoints return 0. Date conversion and holiday sequence errors follow the
common rules.

WORKDAY converts Offset to a finite Number and truncates it toward zero to a
checked whole workday count. A non-finite value, a finite value whose
truncated magnitude cannot be represented in the checked day arithmetic, or
an offset whose result leaves the profile date domain is #NUM!. WORKDAY
starts from Date and advances by the resulting Offset workdays, excluding the
start date for a nonzero offset. A positive Offset moves forward, a negative
Offset moves backward, and zero returns the exact input Date serial, including
its fractional time, even if that date is not a workday. For a nonzero Offset,
the day component is stepped and the input fractional time is reattached to
the resulting date; the result is therefore a profile DateTime serial.

The source explicitly permits the Holidays slot to be empty so that Workdays
can be supplied as the fourth argument, for example
NETWORKDAYS(start;end;;workdays). This empty third slot means no holidays.
An omitted fourth slot means the default workweek. A supplied fourth slot
whose sequence is malformed or not exactly seven values returns #VALUE!.
Explicit Empty worksheet cells remain values and are not treated as omitted
argument slots.

### YEARFRAC

YEARFRAC converts and floors both DateParams. Basis defaults to 0 and is
an Integer conversion toward zero. The §4.11.7 procedures order the dates
before counting, so this profile returns a nonnegative fraction for a
reversed interval as well as a forward interval. This follows the normative
Procedure A/B/C ordering and Procedure E chronological assumptions.

| Basis | Day count | Days in year |
| ---: | --- | --- |
| 0 | US/NASD 30/360, Procedure A | 360 |
| 1 | Actual days, Procedure B | Procedure E |
| 2 | Actual days, Procedure B | 360 |
| 3 | Actual days, Procedure B | 365 |
| 4 | European 30/360, Procedure C | 360 |

Basis 0 uses Procedure A: apply the normative date-truncate operation
(represented as floor in this profile), order the dates, adjust 31sts and
February ends as specified in §4.11.7.3, and divide the resulting day count
by 360. Basis 1 uses actual elapsed days divided by Procedure E's year-length
rule, including the average year length for multi-year intervals and its
specified leap-day cases. Basis 2 and Basis 3 use actual elapsed days divided
by 360 or 365. Basis 4 uses Procedure C, orders dates, adjusts 31sts to 30,
and divides by 360. Intermediate counterfactual February 30 dates in these
procedures are internal values, not formula errors. Invalid basis is #NUM!.

## Matrix, reference, and cache behavior

The scalar evaluator has no Resolver or Position. It can execute context-free
date kernels and timestamp-backed NOW/TODAY/EASTERSUNDAY, but a scalar
Reference or sequence requiring cell reads returns typed
Unsupported(Reference) rather than inventing an origin.

The value evaluator retains the existing argument-shape planner. Date and
Offset arguments use the ordinary scalar intersection or projected
elementwise schedule. NETWORKDAYS and WORKDAY mark DateSequence and
LogicalSequence arguments as complete sequence consumers; their holiday and
workweek references are consumed in full for each projected scalar result and
must not be implicitly intersected or partially selected. ReferenceList
admission follows the pseudotype rules above and is rejected before any read
where the type is known statically.

A date formula Error is a result value. For a sequence reducer, it is retained
while later sequence members are scanned so a typed resolver, cancellation,
resource, allocation, or source-version failure can supersede it. A matrix
result may contain formula Errors in individual cells, but a typed failure
publishes no partial result.

Demand-cache entries may contain only complete scalar payload/error values and
must retain the existing position, projected-shape, source-version, and
cancellation identity. A cached timestamp-backed volatile result is valid only
under the same explicit timestamp snapshot. A cache hit must not bypass a
required source or cancellation fence. This family does not alter the
position-sensitive scalar MUNIT criterion or its existing cache exclusion.

## Resource and safety requirements

The date kernels use fixed-size civil-date state and checked integer
arithmetic. They do not allocate for ordinary scalar operations. Parser
scratch, sequence metadata, holiday membership, matrix arrays, and output
arrays reserve fallibly against the existing finite limits. Where a local
buffer owns a reservation token, declare the reservation before the buffer so
Rust's reverse local drop order drops the buffer first. Where the pair is a
struct, declare the buffer field before the reservation field so field drop
order has the same property. Release each reservation token before its
enclosing budget lease is dropped; do not rely on a blanket scope statement
that hides the ownership order.

Reference-backed holiday/workday sequences are streamed. Before each
inspected reference cell, including cells later skipped by sequence
conversion, the evaluator charges cell work, checks the reference-cell limit
and cancellation, reads through the resolver, and checks cancellation again.
Borrowed text remains borrowed through conversion. Source-version and final
cancellation fences surround the complete value evaluation and publication.
No provider failure is caught by a formula-level error handler.

Date iteration for NETWORKDAYS and WORKDAY uses checked day arithmetic and
periodic work/cancellation checkpoints. A large date interval cannot bypass
the work budget merely because it contains no references. An implementation
may use a bounded seven-day cycle plus holiday handling, but it must preserve
the exact inclusive/offset semantics and must not turn an unbounded holiday
list or day interval into an unchecked allocation or loop.

## Normative and profile decisions still requiring review

This draft makes the following explicit choices so implementation and tests
cannot diverge:

1. Date constructors use the wider year 1–9999 profile domain while retaining
   the §6.10 DATE baseline constraint as the interoperable core. This is
   required for a representable EASTERSUNDAY(1583) result.
2. DateParam and TimeParam Logical conversion is 1/0.
3. Date/time Text parsing uses fixed en_US forms and a 1930 two-digit pivot,
   with the permitted VALUE fallback.
4. DATEDIF rejects reversed intervals with #NUM! and does not swap them.
5. MINUTE uses the rounded-total-seconds formula because the HTML contains a
   conflicting later formula.
6. TIME preserves finite fractional Number inputs and uses the direct
   formula; the optional INT() preprocessing permitted by §6.10.18 is not
   selected.
7. NETWORKDAYS reversed intervals are negative; WORKDAY excludes the start
   date for nonzero offsets, preserves its input fraction, and returns the
   exact input serial for zero offset.
8. WORKDAY truncates a finite fractional Offset toward zero; non-finite and
   checked-day-range overflow are #NUM!.
9. YEARFRAC applies the normative ordered-date procedures and therefore
   returns a nonnegative fraction for reversed dates.
10. The date-function serial domain is wider than VALUE's established
    nonnegative domain; shared helpers must receive an explicit domain rather
    than silently widening VALUE.

## Required validation

The implementation evidence must include:

* exact arity, omitted versus explicit-empty slots, formula-error precedence,
  and typed-failure propagation for every one of the 24 names;
* DATE month/day rollover, leap-year and 1900 non-leap behavior, profile
  bounds, serial fractions, and checked overflow;
* fixed-profile ISO, en_US numeric, English month-name, datetime, clock,
  two-digit-year, malformed-text, and numeric-fallback parsing;
* DATEVALUE integer extraction, TIMEVALUE fraction extraction, HOUR/MINUTE/
  SECOND boundary rounding, TIME negative/multi-day inputs, and leap-second
  rejection;
* DATEDIF all six format codes, complete-unit boundaries, MD/YM/YD
  end-of-month cases, reversed intervals, DAYS fractional subtraction,
  US/European DAYS360 cases, EDATE/EOMONTH month-end clamping, and Easter
  years 1583 and 9956;
* WEEKDAY every accepted Type, WEEKNUM every accepted Mode including 21 and
  150, and ISO week-year boundary dates;
* NETWORKDAYS and WORKDAY default weekends, custom seven-element workweeks,
  empty holiday slots, duplicate holidays, reference Text/Empty skipping,
  formula errors, reversed intervals, negative/zero offsets, and all-workday
  refusal;
* YEARFRAC bases 0–4, leap-year spans, 30/360 February and 31st rules,
  nonnegative reversed-date results, and invalid bases;
* explicit timestamp-backed NOW, TODAY, and no-Year EASTERSUNDAY, plus
  typed refusal when no timestamp is supplied; and
* bounded parser scratch, holiday sequence reservation/drop order, reference
  cell/work/cancellation/source fences, matrix broadcasting/intersection,
  complete sequence consumption, and position-sensitive demand-cache keys.

Native spreadsheet results may corroborate ordinary vectors, but they cannot
replace the local ODF semantics or the explicit profile choices above. This
draft claims no production support or PASS disposition until implementation,
focused tests, resource evidence, and independent review are complete.
