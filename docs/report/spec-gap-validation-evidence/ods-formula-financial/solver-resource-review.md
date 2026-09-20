# ODS financial solver and reducer resource review

This is a kernel-scope review of the resolver-free financial root solvers and
sequence reducers. It checks the current source against ADR 0005's hierarchical
resource rules, ADR 0006's typed-failure boundary, and the selected financial
contract. It does not claim that the complete value adapter or the whole
financial family is accepted.

The report is bound to the source hashes below. A later edit to either kernel
requires a fresh hash review. This report update made no production or Cargo
change; adapter regressions remain in the isolated financial development
checkout.

## Reviewed inputs

| Input | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation/financial/solver.rs` | `387f9791103bc97fb85fa99e46502e0b5e74f8da59586ba755a8785fd18761c8` |
| `crates/litchi-ods/src/codec/formula/evaluation/financial/reducers.rs` | `ecda9ba493af93ac7fb8f7da24048b3dd6888a52e18ae8b2d45ee3bc8a6c9e27` |
| `ods-formula-financial/contract.md` | `fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b` |
| `ods-formula-financial/numerical-design.md` | `fe91e17caae719784653fce7f74e0bcbd892e2c5749b06abc5bfc4e7d71f8bbe` |
| `docs/adr/0005-io-memory-and-performance.md` | `34a6148a8fe77b3e90212810996667b654e25209fb49c56aaecd8d8fbe83f770` |
| `docs/adr/0006-validation-security-and-compatibility.md` | `b686465f342b2e051f856f094e38c2333aa05a2768bb255fc5b223cbba1f3381` |

## Kernel findings

* `solver.rs` keeps two `LogSum` accumulators with a fixed 1,025-bin array.
  `WORKING_STATE_BYTES` exposes the exact two-accumulator footprint to the
  caller. The root search has explicit 64-step bracketing, 256 evaluations,
  and 128 iterations; there is no input-sized heap state.
* Every cash-flow/date term charges work before it is processed. RATE's fixed
  products and derivative terms have explicit charges. Bracket probes,
  Newton/bisection iterations, validation passes, logarithmic-bin shifts, and
  final scale/finish passes are charged as bounded work. `Budget` checks
  cancellation before the evaluation cap can synthesize `#NUM!`, and the
  residual finalizer checks it after each fixed-bin pass.
* A nonzero `LogSum` shift records both fixed-bin loops and the finalizer charges
  that recorded work. The finalizer also charges the full scale pass and two
  full spans for each reverse finish. This keeps the hidden 1,025-bin scans
  visible to the work budget.
* `reducers.rs` uses fixed `ScaledProductSum` state and fixed `LogSpan` state.
  `NpvReducer`, `XnpvReducer`, `FvScheduleReducer`, and `MirrReducer` retain
  generated formula errors while accepting only finite admitted numbers. The
  XNPV reducer has fixed state for the rate, first date, fractional-date
  discount path, and cancellation-preserving sum; it does not retain resolver
  data or allocate.
* `ReducerKind` exposes conservative per-element and finalization costs. The
  XNPV cost includes the fractional power/log path; MIRR accounts for its two
  accumulators. The value bridge can therefore charge before entering every
  reducer operation and before consuming a fixed-state result.
* `ProductSumError::ExponentSpan` is preserved as
  `ReducerError::ProductSum` rather than being flattened into `#NUM!`. The
  value bridge can map it to a typed `Resource::Memory` refusal. Ordinary
  product failure and the separate 2,048-natural-log `LogSpan` refusal remain
  formula-level numeric errors. This resolves the stale cross-module concern
  recorded by the earlier review.
* The reducers never read a resolver and never materialize a sequence. The
  caller owns ordered admission, formula-error retention, work/cancellation
  callbacks, and typed provider-failure precedence. A callback failure exits
  as an `EvaluationFailure` and cannot be converted by these kernels into a
  formula value.

The source disposition is **PASS for the fixed-state kernel scope**. It is not a
whole-adapter acceptance claim.

## Adapter handoff requirements

The value adapter must reserve `WORKING_STATE_BYTES` before each IRR, RATE, and
XIRR root call and hold that reservation through success, generated formula
errors, typed failures, cancellation, and nonconvergence. It must release the
reservation only after the solver state and any retained numeric buffers are no
longer live. XIRR's value/date scratch remains live during the root call, so the
root-state reservation is additional to those buffers.

The adapter must charge each `ReducerKind` operation before calling `push_rich`
or `finish_rich`, map `ExponentSpan` to the typed memory refusal, and continue a
complete admitted source scan after a generated rate/date/reducer error. The
focused limits coverage should retain regressions for an invalid XNPV rate
followed by a late Dates provider failure and by an earlier Dates formula error:
the provider failure must supersede retained generated errors, while the
formula error must remain the published value. The deferred-rate path must scan
Dates before publishing either result.

One maintenance invariant remains visible at the kernel boundary: if a future
solver change adds an early return after a logarithmic-bin shift, it must retain
the recorded shift work or charge the completed loops before returning. The
current finite/error paths do not create a kernel blocker, but this invariant
must remain part of any subsequent solver change review.

No Cargo or runtime validation was run as part of this read-only kernel review.
The adapter's focused test and gate receipts should be cited separately once
the current isolated source is frozen.
