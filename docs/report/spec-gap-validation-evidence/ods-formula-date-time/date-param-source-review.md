# DATEDIF and ISOWEEKNUM DateParam source review

Status: **PASS** for the frozen date/time source. This supplementary review
binds the DATEDIF DateParam/flooring path and the ISOWEEKNUM DateParam/domain
path to their frozen implementation. It preserves the existing semantic,
resource, WORKDAY, and clock reports.

## Review identity

* Contract: `contract.md`, SHA-256 `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`.
* Freeze: `gates/freeze.json`, SHA-256 `461a76708e36b2a716cd622df45f014e009186223bd1db511c3e0ff0f7fa3561`.
* Frozen conversion source: `crates/litchi-ods/src/codec/formula/evaluation/date_time.rs`, SHA-256 `29f1495a55aac7effee9f6509deef38e5d5f747ba2b87c65a188d8759c1265e3`.
* Frozen date kernel called by that source: `crates/litchi-ods/src/codec/formula/evaluation/date_time/kernel.rs`, SHA-256 `2bc377db1a01e28dbc4940c99c61fb43a308919b315bfc37a92ce86a67e5d0cc`.
* Frozen calendar conversion dependency: `crates/litchi-ods/src/codec/formula/evaluation/calendar.rs`, SHA-256 `8d533dbe4f153997ab4417085d3c7616c0de5a60c8d1dec051350a64ed99bf4d`.
* Frozen focused evaluation tests: `crates/litchi-ods/tests/ods_formula_date_time_evaluation.rs`, SHA-256 `d9644432fd5dcac9272dcd3f942e4b669865db22fb6fe866a706ebd05bf04ea1`.
* Final independent review receipt: `review-receipt.json`, SHA-256
  `af38cc90788e881cc36032c1bd7462f395441e698f02d064d0c7e4e04adbf4e5`.

## DATEDIF conversion and flooring

`Function::valid_arity` requires three DATEDIF arguments, and the dispatcher
routes the function to `apply_datedif`. The function converts both date
operands with `date_only`, which delegates to `date_argument`: finite numeric
serials pass through the profile-domain check, logical values use the defined
numeric conversion, date text uses the fixed parser, and formula errors remain
formula errors. `date_only` then calls the checked `serial_to_date` kernel.

The frozen kernel's `serial_to_parts` delegates to
`calendar::civil_from_serial`. That dependency rejects non-finite and
out-of-profile serials before line 123 applies `serial.floor()`, so the
DateParam date component is deterministic and domain failures remain
`#NUM!`. `apply_datedif` converts the third Text format after both date
arguments, trims ASCII whitespace, accepts only the defined case-insensitive
format names, and delegates the ordered civil-date calculation to
`kernel::datedif`. That kernel returns `#NUM!` when the end date precedes the
start date before performing the unit calculation.

## ISOWEEKNUM conversion and domain handling

The catalog and dispatcher require exactly one ISOWEEKNUM argument and route it
to `apply_weeknum` with the ISO-only flag. The same `date_only` path floors a
finite fractional serial through `civil_from_serial` and rejects values
outside the supported serial domain before constructing a civil date.
Malformed text and formula errors remain the corresponding formula values.
For an admitted date, `iso_week` performs checked civil-day arithmetic and
returns the ISO week number, propagating a domain failure instead of
manufacturing a result.

The frozen evaluation tests cover exact arity, malformed date text, formula
error propagation, and fractional ISO date flooring. They do not add a new
numeric out-of-range vector in this supplementary review; the domain proof is
bound to the unchanged checked `civil_from_serial` conversion and ISO kernel
source above.

## Bound requirements

The receipt splits the exact function-level requirements so metadata can bind
each function's union independently:

### DATEDIF

* `exact arity 3 and DateParam/Text conversion`
* `floored date-only operands and reversed interval #NUM!`

### ISOWEEKNUM

* `exact arity 1 DateParam conversion`
* `profile date flooring and domain errors`
