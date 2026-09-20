# WORKDAY conversion source review

Status: **PASS** for the frozen date/time source. This supplementary review
binds the WORKDAY conversion requirement to the shared scalar and sequence
kernel. It does not change the existing semantic or resource review reports.

## Review identity

* Contract: `contract.md`, SHA-256 `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`.
* Freeze: `gates/freeze.json`, SHA-256 `461a76708e36b2a716cd622df45f014e009186223bd1db511c3e0ff0f7fa3561`.
* Frozen source: `crates/litchi-ods/src/codec/formula/evaluation/date_time.rs`, SHA-256 `29f1495a55aac7effee9f6509deef38e5d5f747ba2b87c65a188d8759c1265e3`.
* Final independent review receipt: `review-receipt.json`, SHA-256
  `af38cc90788e881cc36032c1bd7462f395441e698f02d064d0c7e4e04adbf4e5`.

## WORKDAY path

`workday_scalar` obtains the first argument through `date_argument`, which
accepts finite in-domain numeric serials through `valid_serial`, converts a
logical value to `1` or `0`, parses date text through the fixed date parser,
and preserves formula errors. It obtains Offset through the shared
`number_argument` conversion, then invokes `workday_loop` after optional
holiday and workweek conversion and source-order formula-error inspection.

`workday_loop` rejects non-finite Offset values, applies `offset.trunc()` for
toward-zero conversion, and checks the truncated value against the signed
`i64` range before casting. A zero truncated Offset returns
`valid_serial(start_value)` directly, preserving the exact valid input serial,
including its fractional part. Nonzero offsets use the sign and absolute
step count, charge each calendar step, use checked civil-date movement, and
reattach the original fractional part after the final workday before the
serial-domain check.

The resolver-backed sequence path reaches the same `workday_loop` through
`finish_sequence`, so reference-derived start dates and offsets use the same
conversion and domain rules after their admitted sequences have been scanned.
The shared loop therefore covers both the scalar façade and the retained
reference path without a second Offset conversion implementation.

## Bound requirement

This report binds exactly the coverage requirement:

* `DateParam and toward-zero finite Offset conversion`
