# ODS date/time evaluation

The evaluator implements the 24 OpenDocument 1.4 Part 4 §6.10 functions within
[the documented deterministic profile](contract.md): DATE, DATEDIF, DATEVALUE,
DAY, DAYS, DAYS360, EASTERSUNDAY, EDATE, EOMONTH, HOUR, ISOWEEKNUM, MINUTE,
MONTH, NETWORKDAYS, NOW, SECOND, TIME, TIMEVALUE, TODAY, WEEKDAY, WEEKNUM,
WORKDAY, YEAR and YEARFRAC.

The implementation uses the Gregorian epoch 1899-12-30 without a synthetic
1900 leap day. NOW, TODAY and omitted-date EASTERSUNDAY require an explicit
validated calculation timestamp; evaluation never reads an ambient clock.
DATEVALUE/TIMEVALUE use the fixed parsing profile and retain the existing VALUE
numeric grammar. The contract describes bounds, rounding, omitted arguments,
reference admission and native compatibility differences precisely.

References are read under work, read, cancellation and source-version checks.
Holiday/workweek sequences are consumed completely before publishing a result;
typed failures supersede retained formula errors. Holiday storage and array
outputs use bounded reservations. Scalar Date/Offset arguments remain
position-sensitive under projected evaluation; complete sequence descriptors
retain their full-reference demand.

## Evidence

- [Batch completion and limitations](completion.md) and
  [strict root verification](root-verification.json).
- [Final freeze](gates/freeze.json) at source candidate `d16039ce48` and
  [seven verified gates](gates/verification.json): 1,745 passing tests, no failures
  or ignored tests.
- [Independent semantic review](spec-review.md), [resource/cache review](contract-resource-review.md),
  and [source proof receipts](review-receipt.json).
- [132 independently recomputed oracle vectors](oracle-vectors.json), their
  [executed Rust replay receipt](oracle-execution.json), and
  [132 native observations](native/native-results.json) with explicit profile
  and host differences.
- [Requirement-level coverage audit](coverage-audit.md) and
  [machine-readable bindings](coverage-requirements.json). These bindings are
  checked separately from the gate and performance receipts.
- [Full performance report](performance/results/performance-report.md),
  [raw capture and metrics](performance/results/performance-report.json), and
  [independent performance review](performance-review.md).

All 4,620 performance samples validate. Across 68 matched control groups, the
allocation, work, read and output accounting sets are unchanged. One median
latency exceeds the 5% review threshold: 4×4 SIN parse/evaluate changes from
11,033.8125 to 11,610.6875 ns (+5.228%; 95% bootstrap interval 0% to +10.909%).
Twenty-two median RSS groups increase by 184–240 KiB. These limitations are
accepted and disclosed for this function-support batch; no overall speedup or
causal explanation is claimed. Tail flags and all uncertainty intervals remain
in the report. Timing includes cached expected-value comparisons, checksum and
drop; independent oracle computation is outside measurement.

## Rechecking the evidence

From the repository root:

```sh
python3 docs/report/spec-gap-validation-evidence/ods-formula-date-time/verify.py
python3 docs/report/spec-gap-validation-evidence/ods-formula-date-time/oracle_verify.py --check
python3 -m unittest discover -s docs/report/spec-gap-validation-evidence/ods-formula-date-time/performance -p 'test_*.py'
python3 docs/report/spec-gap-validation-evidence/ods-formula-date-time/summarize_performance.py
```

The summarizer revalidates the raw capture before regenerating its report.
[Gate preparation](gates/README.md) and the [performance plan](performance/performance-plan.md)
describe reproduction in isolated checkouts. Final raw evidence is retained;
[owned build directories and temporary diagnostics were removed](cleanup.json).
The root Cargo.lock is preserved; gates use their separately retained lock.
