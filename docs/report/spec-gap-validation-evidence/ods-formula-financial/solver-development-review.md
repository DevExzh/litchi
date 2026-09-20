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
