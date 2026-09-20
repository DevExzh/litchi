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
