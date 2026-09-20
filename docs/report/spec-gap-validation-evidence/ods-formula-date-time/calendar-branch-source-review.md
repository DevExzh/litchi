# WEEKNUM and EASTERSUNDAY calendar-branch source review

Status: **PASS for the frozen source proof only**. This supplementary review
closes two function-level evidence bindings through the checked source path.
It does not add or imply new focused-test or oracle-vector coverage. The
existing retained vectors remain exactly the vectors named by the coverage
manifest.

## Review identity

* Contract: `contract.md`, SHA-256
  `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`.
* Freeze: `gates/freeze.json`, SHA-256
  `461a76708e36b2a716cd622df45f014e009186223bd1db511c3e0ff0f7fa3561`.
* Final independent review receipt: `review-receipt.json`, SHA-256
  `af38cc90788e881cc36032c1bd7462f395441e698f02d064d0c7e4e04adbf4e5`.

The primary frozen dispatch source for both proofs is
`crates/litchi-ods/src/codec/formula/evaluation/date_time.rs`, SHA-256
`29f1495a55aac7effee9f6509deef38e5d5f747ba2b87c65a188d8759c1265e3`.
The proof also follows these frozen dependencies:

* `crates/litchi-ods/src/codec/formula/evaluation/calendar.rs`, SHA-256
  `8d533dbe4f153997ab4417085d3c7616c0de5a60c8d1dec051350a64ed99bf4d`;
* `crates/litchi-ods/src/codec/formula/evaluation/date_time/kernel.rs`,
  SHA-256
  `2bc377db1a01e28dbc4940c99c61fb43a308919b315bfc37a92ce86a67e5d0cc`; and
* `crates/litchi-ods/src/codec/formula/evaluation/timestamp.rs`, SHA-256
  `c8ec12c46a116f48a60387e61e339f362770831b07b1c8e03c84958f134c60e7`.

## WEEKNUM DateParam flooring and week-start calculation

`apply_weeknum` obtains its first argument through `date_only` before it
selects a mode. `date_only` calls `date_argument`, whose numeric and parsed
text paths first pass the finite profile-domain check and then call
`serial_to_date`. The shared kernel delegates that conversion to
`calendar::civil_from_serial`; after its finite half-open range check, the
calendar conversion uses `serial.floor()` to obtain the civil date. Thus a
fractional DateParam is admitted as a DateTime serial and its time component
is discarded at the date-only boundary. Invalid or formula-error values remain
the corresponding error result.

For an admitted civil date, `weeknum_result` maps each accepted non-ISO mode to
its defined week-start weekday, computes the day-of-year offset from January
1, and performs the checked week-number calculation. Modes 21 and 150 route
through the shared ISO-week kernel. The source therefore proves both parts of
the bound requirement through one path: DateParam flooring and the
mode-dependent week-start calculation.

There is no dedicated retained fractional `WEEKNUM` vector in this review:
`component_and_month_shift_edges_use_checked_profile_bounds` does not invoke
WEEKNUM, and the retained `weeknum.mode_1` oracle row uses an integer serial.
This receipt records the implementation proof and does not relabel those
rows as fractional-input coverage.

Bound requirement:

* `DateParam flooring and week-start calculation`

## EASTERSUNDAY current/following-year selection and timestamp limits

The omitted-Year branch of `apply_easter` requires the injected calculation
timestamp and immediately converts its serial through `serial_to_date`, so the
comparison uses the timestamp's floored profile date. It computes Easter for
that date's year and compares the date with the current Easter using the
checked civil-date ordering. On or before the current Easter it returns the
current year's Easter; strictly after it, it computes the following year's
Easter. This is exactly the contract's smallest Easter date not earlier than
the supplied TODAY value.

The range proof is composed from the same frozen call chain. The timestamp
constructor admits only finite serials in the checked profile half-open date
range. The shared calendar conversion rejects non-finite and out-of-range
serials before constructing a civil date. Finally,
`kernel::easter_sunday` accepts only years 1583 through 9956 and returns
`#NUM!` outside that interval. Consequently a timestamp in years 1 through
1582 or 9957 through 9999 is rejected by the omitted-Year Easter path, and a
timestamp in 9956 after that year's Easter attempts the following year and is
also rejected. The explicit-year path uses the same bounded Easter kernel.

There is no dedicated retained before/on-current-Easter timestamp vector in
this review. The retained `eastersunday.timestamp_after_current_easter` row
covers only the following-year branch, while `eastersunday.invalid_year`
tests an explicit Year value and is not timestamp-boundary evidence. This
receipt supplies the frozen source proof without fabricating either missing
vector claim.

Bound requirement:

* `current/following-year selection and timestamp range limits`
