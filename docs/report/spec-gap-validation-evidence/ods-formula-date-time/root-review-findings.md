# Root implementation review findings

These are development findings, not final source or gate dispositions.
Independent contract reviews do not establish that an implementation conforms.

## Timestamp representability

A standalone Rust test probe compiled snapshots of `calendar.rs` and
`timestamp.rs`, without the unfinished evaluator modules. Four of five tests
passed: the two then-present timestamp unit tests, checked rejection of
`i64::MIN/MAX` civil-day inputs, and exact timestamp bit preservation for
negative near-zero inputs.

The fifth test failed: constructing civil time `2020-01-01 23:59:second`, with
`second = f64::from_bits(60.0_f64.to_bits() - 1)`, succeeded with serial
`43832.0` (January 2). This silently advances the caller's civil date because
the fractional-day sum rounds upward. The foundation owner was asked to reject
an unrepresentable same-day timestamp with a typed construction error and add
a permanent regression. This diagnostic does not claim a package test run.

After the fix, the two expanded timestamp unit tests and a standalone
exhaustive calendar test passed (three tests total). The calendar test checked
serial-to-civil-to-serial identity for every integer serial from -693593 through
2958465, plus rejection of `i64::MIN/MAX` before unchecked civil conversion.
The timestamp constructor now rejects a rounded result whose integer day differs
from the requested civil date. The temporary probe sources and binary were
removed; package-level integration remains pending.

## Calendar and financial-date kernels

Early source inspection found an unsigned saturating Thursday offset in ISO
week calculation; Friday through Sunday need negative offsets. It also found
YEARFRAC basis 1 dividing by the number of years times a year length, instead
of Procedure E's average-year denominator, and basis 0 reusing DAYS360 rather
than implementing Procedure A's distinct February and adjustment ordering.
The kernel owner and independent semantic reviewer received these findings.

Shared calendar algorithms must have one owner. The initial kernel duplicated
civil-date conversion and leap/month helpers being added by the foundation
owner. Consolidation is required before the implementation gate.

## Evidence verifier

The initial coverage verifier checked file hashes but did not validate receipt
outcomes or identifier presence. Hash-bound arbitrary placeholders could
therefore have supported an unjustified PASS promotion. The evidence owner
was asked to keep aggregate verification pending until kind-specific receipt
checks exist and to add negative cases for forged receipts and identifiers.

Resolution of these findings requires checking the final source and relevant
tests; this document is not a substitute for that check.

## Integration diagnostics and gate custody

The third isolated `cargo check --locked --offline -p litchi-ods` snapshot
reduced compilation failures to two value-adapter errors: an unwrapped text
argument passed to `to_number`, and an unnecessary mutable binding. After
those fixes, the fourth snapshot reached thirteen unused-helper errors.
Neither run is a passing package gate. The isolated checkout retains lock
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`;
the ambient root lock was not changed.

Gate staging now selects the complete recursive evaluator source closure,
including the shared projected-argument helper in `value/lookup.rs`. Oracle
inputs use an explicit allowlist; future execution receipts cannot enter the
freeze and create a circular dependency. Negative cases also reject arbitrary
performance harness outputs and mutable coverage bindings. Root reran both
gate-closure and coverage negative suites successfully.

The independent oracle checker recomputed all 132 vectors across 24 functions
against contract `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`.
The native reproducer checked 126 deterministic rows and kept six volatile
host observations separate, using a fresh profile that it removed afterwards.
These checks validate expected-value and compatibility artifacts; Rust oracle
execution, final source review, full package gates, and performance capture
remain pending.

## Executed date/time diagnostics

After unused-helper cleanup, the focused targets reached test execution.
The semantic target passed all twelve tests in diagnostic snapshot 12 after
these corrections:

* WEEKNUM rejects non-integer modes rather than truncating them.
* Omitted-year EASTERSUNDAY selects the next eligible Easter after the
  injected current date, including the upper-year refusal.
* TIMEVALUE returns the parsed clock fraction directly for datetime text,
  avoiding precision loss from subtracting a large integer date serial.
* The Procedure E multi-year fixture uses the inclusive-year average
  denominator. An unrepresentable final-day timestamp is expected to fail
  construction rather than round into the following day.

The first executed limits snapshot passed five of six tests. Its remaining
failure was the test observer classifying a legitimate 1-by-1 matrix result
as `Other`; the fixture now checks that exact shape before extracting its
element. Subsequent checks also cover actual formula-error precedence over
generated conversion errors, zero-read list refusals, negative subnormal HOUR,
and signed half-second rounding. Their final rerun is still pending.

MINUTE/SECOND now round signed total seconds before taking the day remainder,
as the contract specifies. The independent oracle's original
`59.5/86400` half-second fixture exposed binary64 precision: multiplying its
represented value by 86400 gives `59.49999999999999`, not exactly 59.5.
The oracle owner is reviewing this fixture independently; no tolerance was
added to production rounding to force agreement.

The completed development rerun passed all 1,724 library and integration
tests in 128 suites, including 13 date/time semantic tests, six limits tests,
and the Rust consumer of all 132 oracle rows. All six documentation tests and
strict Clippy across all targets also passed. The final oracle correction
retains signed rounding for `SECOND(-0.5/86400) = 59`; its independent
checker agrees with corpus hash
`273421f087314733c9b383e6a1b5c1bb1e43de3b65b0be24aeb1945864e6438d`.
Root compared every ODS source and test Rust file with the tested isolated
checkout and found no differences. The feature matrix records implementation
scope while explicitly leaving frozen validation and performance pending.

Independent source reviews found no remaining semantic or resource blocker.
The final freeze, reproducible gate receipts, complete coverage mapping, and
performance capture are still required; these development runs do not replace
those deliverables.

## Pre-freeze harness and coverage checks

The package formatting check and workspace boundary checker passed after a
parser formatting cleanup. The boundary checker reports 65 packages, 241
internal dependency declarations, and 11 already declared migration debts.
The native reproducer again validated 126 deterministic observations and six
host-clock observations using a temporary profile that it removed.

The performance harness now compiles against both the candidate and the
actual production baseline, with the new timestamp API enabled only for the
candidate. Its standalone lock was generated offline from the retained gate
lock; comparing package identities and registry checksums found only the new
harness package, with no new external versions. The harness lock hashes to
`5261723f62e5ebbea6875cb51b29dbfbfdfd6dca1a651febf7779a2d65a70319`.

The first candidate preflight passed 89 of 120 workload cases. The 31 failures
mostly identify malformed range spellings in fixtures, plus drift between
two time formulas and their expectations and a scalar-versus-1-by-1-array
expectation. They are preparation failures, not performance measurements.
Correcting them and replaying all cases on both relevant builds is required
before timing capture.

The separate coverage audit also identified missing assertions beyond the
passing tests: full numeric fallback grammar, broader formula-error
precedence, WORKDAY reference sequence failures, and matrix failure cleanup.
These were assigned for focused test additions before freezing source.
Architectural properties such as absence of ambient I/O require source-review
evidence alongside tests; they must not be claimed from a fabricated mock.

## Projected scalar reference regression and preparation replay

The corrected performance harness passed all 120 candidate preflight cases
and all 34 baseline controls, including exact formula, evaluation path,
output assertions, and resolver-read counts. The retained development receipt
is `performance/development-preflight.json`; it predates the cache fix below
and is not a frozen gate or a timing result.

Strengthening the matrix failure test exposed a production cache defect:
`IF({TRUE()|TRUE()};DATEVALUE([.A1:.A2]);0)` reused the first projected cell
and suppressed the second resolver read. Date/Offset reference arguments now
use position-sensitive scalar reference classification. Complete holiday and
workweek sequence descriptors retain full-reference cache classification.
The success regression requires two distinct dates and exactly two reads;
the failure regression requires a typed provider failure on the second read,
ordered reads of A1 then A2, and zero retained memory after refusal.

After the fix, isolated locked/offline validation passed 19 semantic tests,
10 resource tests, and the 132-vector oracle target. All-target strict Clippy,
Rust 2024 formatting, and diff checks passed. These remain development
checks; final frozen gates and performance capture are still outstanding.
