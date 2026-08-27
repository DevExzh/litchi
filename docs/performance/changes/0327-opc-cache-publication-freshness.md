# Change 0327: OPC cache publication freshness

## Status

Implemented and validated.

## Scope

Change 0327 hardens the `litchi-opc` source-backed cache publication path. A
cache flight is published provisionally before entry construction, while the
final source-version observation and execution fence are performed outside the
cache locks. The publication protocol now preserves exact flight and payload
identity across commit, rollback, and invalidation.

Terminal states are conditional: a result is committed only when the required
freshness and identity conditions still hold. Failure reservations are released
before notifying waiters, so a failed producer cannot leave a stale reservation
behind or wake waiters into an unrecoverable state. If allocation admission is
not available, the fallback is genuinely uncached rather than partially
published. Stable hit and bypass diagnostics remain available for observing the
chosen path, and the borrowed `into_arc` compatibility behavior is retained.

The protocol closes mutations and cancellations observed by the required
freshness fences. It does not make a mutation atomic when it occurs after the
final version observation and before the operation returns. Residual waiter
retry semantics are unchanged.

## Validation

Validation was serialized with `CARGO_BUILD_JOBS=1` and one isolated Cargo
target at a time to avoid concurrent rebuilds and excessive memory use.

- `litchi-opc`: 272 library tests passed.
- Strict `litchi-opc` Clippy passed.
- Strict `litchi-opc` rustdoc passed.
- XLSB source-backed tests: 25 passed.
- XLSX library tests: 808 passed.
- DOCX source-backed tests: 16 passed.
- ODT tests: 14 passed.
- ODP tests: 10 passed.
- Rustfmt check passed.
- Diff check passed.

This change makes no performance, RSS, or OOM claim.
