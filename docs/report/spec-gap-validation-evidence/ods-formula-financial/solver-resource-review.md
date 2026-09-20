# ODS financial solver resource review

This is a focused review of the resolver-free IRR, RATE, and XIRR solver
kernel. It checks the solver delta against ADR 0005's hierarchical resource
rules, ADR 0006's typed-failure boundary, and the selected financial contract.
It does not review the complete value adapter or claim whole-family support.
No production, test, or Cargo files were changed for this review.

## Reviewed inputs

| Input | SHA-256 |
| --- | --- |
| `evaluation/financial/solver.rs` | `6624fce0af61ed0b6d7dddcad0e7bf0623fe17821e53e236a6d741bc3eb8d4b3` |
| `ods-formula-financial/contract.md` | `fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b` |
| `docs/adr/0005-io-memory-and-performance.md` | `34a6148a8fe77b3e90212810996667b654e25209fb49c56aaecd8d8fbe83f770` |
| `docs/adr/0006-validation-security-and-compatibility.md` | `b686465f342b2e051f856f094e38c2333aa05a2768bb255fc5b223cbba1f3381` |

## Solver disposition

The reviewed solver changes resolve the earlier kernel-level accounting
findings:

* `WORKING_STATE_BYTES` exposes the two-`LogSum` working footprint so the
  caller can reserve it before entering a root call.
* Every nonzero logarithmic-bin shift charges both fixed-bin loops. The
  residual finalization charges the value scale pass and both passes of each
  `finish`, with cancellation checks after each pass.
* RATE's three residual products and derivative terms receive explicit work
  charges and post-operation cancellation checks.
* The numerical evaluation cap checks cancellation before returning its
  generated `#NUM!` refusal.

These changes make the fixed solver state and its bounded scans auditable.
The source-level solver disposition is therefore **PASS for the reviewed
kernel scope**, subject to the adapter reservation below.

## Required adapter handoff

The current value adapter calls `financial::irr_solver`,
`financial::rate_solver`, and `financial::xirr_solver` through a work closure,
but the reviewed tree has no use of `WORKING_STATE_BYTES` outside the solver
definition and its unit test. The adapter must acquire a storage reservation
for that exact working footprint before each solver invocation and hold it
through every result, formula-error, typed-failure, cancellation, and
nonconvergence path. It must release the reservation only after the solver's
stack state and the retained numeric buffers are no longer live. XIRR's value,
date, and offset buffers remain live while the solver runs, so the root-state
reservation is additional to those buffers. RATE should use the same
two-accumulator bound even on its guarded `rate = -1` path.

This is an integration requirement still pending the value-adapter handoff;
it does not invalidate the solver-only PASS above.

## Cross-module follow-up

`ProductSumError::ExponentSpan` remains mapped to `ScalarError::Number` in the
financial reducers. The current contract explicitly maps a fixed scaled-sum
span refusal to `#NUM!`, so this is consistent with the selected semantic
profile. The generic numeric-kernel comment that describes the same payload as
“suitable for a `Resource::Memory` refusal” should be reconciled with that
contract before final support evidence; the pure reducer API cannot propagate a
typed storage failure while it returns only `ScalarError`.

The value adapter also still has an adjacent typed-failure audit item outside
this solver review: its XNPV text path must not discard the result of the text
work/limit charge. That adapter path should propagate the typed failure before
whole-family acceptance.

No Cargo or runtime validation was run in this read-only review. The report is
bound to the source hashes above; any later solver or contract edit requires a
new hash review.

Root handoff note: subsequent reducer edits introduced rich typed failures,
and the XNPV text charge now propagates its failure. These changes still need
adapter integration tests. The solver-specific 2048-natural-log numerical
refusal must remain distinct from the generic accumulator storage-span refusal;
no change to the shared numeric-kernel resource contract is authorized here.
