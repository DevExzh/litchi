# NOW/TODAY clock source review

Status: **PASS** for the frozen date/time source. This supplementary review
binds the function-level NOW/TODAY ambient-state requirement to the explicit
calculation-timestamp path. It preserves the existing semantic, resource, and
WORKDAY supplementary reports.

## Review identity

* Contract: `contract.md`, SHA-256 `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`.
* Freeze: `gates/freeze.json`, SHA-256 `461a76708e36b2a716cd622df45f014e009186223bd1db511c3e0ff0f7fa3561`.
* Frozen source: `crates/litchi-ods/src/codec/formula/evaluation/date_time.rs`, SHA-256 `29f1495a55aac7effee9f6509deef38e5d5f747ba2b87c65a188d8759c1265e3`.
* Final independent review receipt: `review-receipt.json`, SHA-256
  `af38cc90788e881cc36032c1bd7462f395441e698f02d064d0c7e4e04adbf4e5`.

## Clock capability path

`apply_now_today` reads only `evaluator.context.options().calculation_timestamp()`.
When that caller-supplied snapshot is absent, it returns the typed
`Unsupported(CalculationClock)` failure before producing a value. With a
validated snapshot, NOW publishes `stamp.serial()` and TODAY publishes
`stamp.date_serial()`; neither branch reads a system clock, timezone, locale,
filesystem, process environment, or other ambient host state.

The date/time dispatcher validates zero arity before dispatch and routes both
functions directly to `apply_now_today`. The timestamp object is immutable and
constructed through the explicit calculation-timestamp option, so the
function-level path has no fallback clock source or implicit environment
lookup. This review is source-only and does not alter the existing final
review receipt.

## Bound requirement

This report binds exactly the function-level coverage requirement:

* `no ambient wall-clock, timezone, locale, or process state access`
