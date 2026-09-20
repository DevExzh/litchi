# Financial solver development review

Status: **HOLD for resource integration**, with focused numerical probes passing.

The root reviewer compiled an unchanged snapshot of `financial/solver.rs`,
SHA-256 `c967fcd03c7a52f84999206d61b15123d2c128617fb7f1b445b158253a60b7f5`,
in a temporary standalone `rustc --edition=2024 --test` harness. Minimal
formula-error and cancellation types substituted for evaluator types. This
does not exercise public API conversion, resolver reads, storage reservations,
or demand caching. Temporary sources and executable were removed.

All **10 tests passed**: seven existing solver tests and three independent
tests. The independent tests cover:

- IRR two-cash-flow roots at rates `-0.9`, `-0.5`, `0`, `0.1`, `1`, `100`,
  and `1e8`, for cash-flow scales `1e-300`, `1`, and `1e300`.
- XIRR two-cash-flow roots separated by 365 days at the same scales and
  rates through `100`.
- RATE's exact `-1` boundary for `(2, -50, 100, 50, 0, -1)`.

The scale sweeps compare against the analytically constructed rates with
absolute tolerance `1e-10 * (1 + abs(rate))`. They do not establish accuracy
for general multi-period or cancellation-sensitive cash flows.

Independent source/resource review identified outstanding issues:

- Two 1,025-bin accumulator states occupy roughly 65.6 KiB; invocation must
  account for fixed numeric scratch under the storage limit.
- Small bin shifts move the full array but charge only the shift distance.
  Finish/scale charging also omits a possible full pass.
- Cancellation must be checked before refusing at the evaluation cap and
  after residual scans.
- RATE product/term traversal needs explicit work accounting.

The implementation owner is correcting those issues. Numerical accumulator
semantics remain under separate review. No production, native-parity, resource,
or performance PASS is claimed by these standalone numerical results.

## Resource-fix numerical follow-up

Root reran the probes against solver snapshot
`6624fce0af61ed0b6d7dddcad0e7bf0623fe17821e53e236a6d741bc3eb8d4b3`:
**11 passed, 0 failed** (eight embedded tests and three independent tests).
The retained harness is reproducible with:

```sh
python3 -B docs/report/spec-gap-validation-evidence/ods-formula-financial/probe_solver.py CHECKOUT
```

It snapshots the selected source and removes its temporary files on exit.
This follow-up confirms numerical probes still pass after work-accounting
changes; independent resource review and adapter reservations remain pending.

## Boundary regressions found by semantic review

Two further independent probes against the same `6624fce0...` snapshot
compiled successfully and failed at runtime:

| Invocation | Contract result | Observed result |
| --- | --- | --- |
| `IRR([-100, 110], guess=-1)` | `#NUM!` | `0.10000000000000044` |
| `RATE(1.5, -110, 100, 0, 0, guess=-1)` | `#NUM!` | `0.5030684048044026` |

The expanded run had **11 passed, 2 failed**, exit status 101. The contract
explicitly refuses IRR's exact `-1` guess and nonintegral RATE periods at that
boundary. These are implementation failures, not proposed profile changes.
Temporary harness files were removed. The solver owner received both cases
for permanent regression coverage and correction.

Source review additionally identified bracket tie/probe-order, adjacent RATE
branch search, unchanged Newton step, and cross-bin compensation concerns.
Those require targeted fixes and evidence before numerical acceptance; the
earlier passing scale probes do not establish correctness for these cases.

## Boundary-fix follow-up

The retained root probe passes **15 tests, 0 failures** against snapshot
`f8bd6a30140b28cc195d809a2c26af4674db031d2c05f1a134eafa5d2302e405`.
Its embedded tests now include the two failing `-1` cases above, an integral
RATE search on the negative branch, and an underflowed Newton-ratio check.
The temporary harness was removed automatically.

This clears the two reproduced boundary regressions for that snapshot.
It does not clear the remaining source-review findings: a genuine cross-bin
compensation regression and targeted bracket tie/probe-order evidence are
still required, along with final resource and public-evaluator validation.

## Compensation and probe-order follow-up

Root reran `python3 -B docs/report/spec-gap-validation-evidence/ods-formula-financial/probe_solver.py .codex-tmp/ods-financial-development` against solver SHA256
`3b39b8a270f3b5eb4aef4fd1df9d0f67b981e7078d681b8d340f1ca65dccc98d`.
All **18 tests passed**, including the added cross-bucket compensation and
endpoint tie/probe-order regressions. The temporary standalone harness was
removed automatically. This is kernel evidence only; independent source
re-review and public evaluator/resource validation remain open.

## Exact-zero boundary and lint follow-up

The retained root probe passes **19 tests** against solver SHA256
`6b883f407e9c1805f35a48bc427b7d199ec23dc86adea80d91d7522f3f33f889`.
This includes the guarded exact-zero RATE boundary regression, followed by
the Clippy-requested eager midpoint fallback cleanup. The standalone harness
removed its temporary sources and binary. Public roots replay had already
passed before the equivalent fallback cleanup; final frozen gates remain open.
