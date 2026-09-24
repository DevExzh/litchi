# Independent date/time oracle plan

This plan owns the resolver-free contract vectors for the 24 ODF 1.4 date and
time functions:

`DATE`, `DATEDIF`, `DATEVALUE`, `DAY`, `DAYS`, `DAYS360`, `EASTERSUNDAY`,
`EDATE`, `EOMONTH`, `HOUR`, `ISOWEEKNUM`, `MINUTE`, `MONTH`, `NETWORKDAYS`,
`NOW`, `SECOND`, `TIME`, `TIMEVALUE`, `TODAY`, `WEEKDAY`, `WEEKNUM`, `WORKDAY`,
`YEAR`, and `YEARFRAC`.

The source contract is currently SHA-256
`cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`.
The local ODF archive is SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`, and its
retained Part 4 HTML member is SHA-256
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
`oracle-vectors.json` is bound to the contract hash and currently contains 132
vectors across all 24 names. Its current SHA-256 is
`b33011089974b18b1fcc0b984adab839600acc98f09dd850e02522348316259f`.

The JSON is an independent expected-value corpus, not a production test
receipt. Each row has an identifier, function, OpenFormula spelling, typed
expected result, profile tags, and a short calculation basis. Date, time and
datetime results retain their serial subtype and ISO annotation. Fractional
values retain an exact decimal string alongside the JSON number. Formula
errors and absent volatile-clock capability are represented as typed outcomes.
Rows that reject a ReferenceList carry a zero-read assertion; no resolver
workbook is embedded in this corpus.

The calculations use only Python standard-library civil-date arithmetic and
exact `Fraction` values during generation. The independent method is:

- serial dates use the proleptic Gregorian calendar with epoch 1899-12-30;
  the accepted domain is `[-693593, 2958466)`, with no synthetic
  1900-02-29;
- DATE rollover, month-end clamping, Easter computus, weekday/week-number
  boundaries, and integer day differences are calculated from civil dates;
- TIME, DAYS, and YEARFRAC fractions are formed as exact rational day values
  before JSON serialization;
- MINUTE and SECOND use the selected half-away-from-zero rounded total-second
  profile, while HOUR uses the normalized fractional day;
- the MINUTE tie vector uses `239.5/86400`: in binary64, this quotient
  multiplies back to the exact `239.5` half-second tie before half-away
  rounding. The superficially smaller `59.5/86400` quotient multiplies back
  to `59.49999999999999`, so it is not used to claim a mathematical tie;
- the SECOND tie vector uses the same exact quotient: `239.5` rounds to `240`
  total seconds, so the seconds component wraps to zero;
- the negative SECOND tie uses the exactly representable `-0.5/86400` value;
  half-away rounding gives `-1` total second, whose day-normalized component
  is `59`;
- 30US/360, 30E/360, actual/360, actual/365, and Procedure-E day-count cases
  are calculated directly from their stated rules; and
- parser rows are hand-authored from the fixed en_US grammar and 1930 pivot,
  without importing the implementation parser. Native spreadsheet behavior
  is reserved for a later compatibility fixture and cannot replace these
  expectations.

Coverage is intentionally boundary-heavy:

| Function | Representative vector coverage |
| --- | --- |
| DATE | month/day rollover, leap/1900 boundary, truncation, serial bounds, overflow |
| DATEDIF | Y/M/D/MD/YM/YD, leap/end-month clamps, reversed interval, bad format |
| DATEVALUE | ISO, datetime floor, en_US numeric, two-digit pivot, English month, numeric and simple-fraction fallback, malformed date |
| DAY | fractional floor, text datetime, upper serial refusal |
| DAYS | retained fractions, reversed subtraction, mixed text/number conversion |
| DAYS360 | US February and reversed interval, European ordered sign, 31st handling |
| EASTERSUNDAY | explicit years 1583/2024/9956, invalid year, timestamp-backed and missing clock |
| EDATE | leap/non-leap month-end clamp, negative and fractional month counts |
| EOMONTH | leap target, negative shift, lower-domain underflow |
| HOUR | midday, negative time normalization, final-day boundary |
| ISOWEEKNUM | 2021 start, 2020 week 53, 2015 first Thursday |
| MINUTE | half-second up/down, negative day, end-of-day wrap |
| MONTH | fractional floor, month-name parse, malformed text |
| NETWORKDAYS | default/reversed/custom/all-off workweeks, duplicate holiday, formula error, list refusal |
| NOW | exact explicit timestamp and missing-clock capability boundary |
| SECOND | half-second boundary, negative fraction, final minute |
| TIME | direct fractional formula, negative/multi-day values, checked overflow |
| TIMEVALUE | clock/datetime fractions, numeric and simple-fraction fallback, 24:00/date-only/leap-second refusals |
| TODAY | explicit timestamp floor and missing-clock capability boundary |
| WEEKDAY | every accepted Type 1, 2, 3, 11–17, default, invalid type |
| WEEKNUM | every accepted Mode 1, 2, 11–17, 21, 150, default, noninteger refusal |
| WORKDAY | forward/backward/default, zero fractional preservation, holiday/custom/all-off/list refusal |
| YEAR | two-digit pivot, fractional date, minimum profile date, malformed text |
| YEARFRAC | bases 0–4, leap year, reversed nonnegative result, invalid basis |

The contract choices requiring an independent semantic second pass are
recorded as vector tags rather than hidden in implementation assumptions:

1. The DATE year 1–9999 profile is wider than the existing VALUE domain.
2. DateParam and TimeParam Logical values map to 1/0.
3. DATEVALUE/TIMEVALUE use fixed en_US forms with a 1930 pivot and the
   permitted finite numeric fallback.
4. DATEDIF rejects reversed intervals; MD/YM/YD use the explicit end-of-month
   procedures in the contract.
5. MINUTE uses the HTML's rounded-total-seconds formula, matching SECOND,
   despite the later conflicting MINUTE formula in Part 4.
6. TIME uses the direct fractional formula and preserves negative fractions;
   it does not apply the optional INT preprocessing.
7. EASTERSUNDAY without a Year, NOW, and TODAY require the explicit
   calculation timestamp; absence is typed `Unsupported(CalculationClock)`.
8. NETWORKDAYS reverses with a negative sign, while WORKDAY excludes the
   start date for nonzero offsets and preserves the input fraction at offset 0.
9. YEARFRAC orders reversed dates before applying the selected day-count
   procedure and therefore remains nonnegative.

The inspection kernel owner should second-pass these exact seams before
implementation acceptance: DATEDIF MD borrow behavior at February
ends; Procedure-E's multi-year denominator and leap-day cases; whether the
selected noninteger WEEKNUM rejection is preserved at the scalar bridge;
optional empty slots versus explicit Empty cells in NETWORKDAYS/WORKDAY;
formula-error publication after complete holiday/workweek scans; timestamp
propagation through matrix evaluation and demand-cache identity; and the
source conflict between the two MINUTE formulas. These are documented profile
decisions in `contract.md`, not values inferred from a native spreadsheet.

Resource and integration work is deliberately outside this artifact. Later
implementation evidence must bind sequence vectors to complete-consumption
receipts, typed read/cancellation/source failures, bounded holiday state,
matrix broadcasting, and volatile timestamp cache identity. The read-only
independent check is reproducible with `python3 oracle_verify.py --check`; it
recomputes every retained vector ID and fails on missing or extra IDs, unknown
outcome kinds, non-finite numbers, or a contract/corpus hash mismatch. Native
conversion is retained separately as compatibility evidence and never supplies
these expected values. No Cargo command or production edit was used to produce
this plan or vector file.
