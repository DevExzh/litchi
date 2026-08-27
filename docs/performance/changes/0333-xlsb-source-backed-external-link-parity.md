# Change 0333: XLSB source-backed external-link parity

## Scope

Source-backed XLSB loading now has a bounded, lazy external-link catalog so
selected-worksheet materialization can expose the same formula/link context as
eager workbook loading. The change closes the previous parity gap without
opening external workbooks or activating DDE, OLE, add-in, or other external
targets. It also corrects the table catalog's worksheet-ordinal mapping so
source-backed table metadata is associated with the owning worksheet.

The catalog is an immutable semantic result shared by concurrent readers and
stored independently of the physical source-part cache. A source-part eviction
therefore does not discard successfully parsed external-link semantics, while a
source mutation or version change invalidates the result and requires a fresh
materialization.

## Implementation

- Source-backed construction and materialization carry the selected
  `ExternalLinkLimits` policy. One operation-scoped aggregate budget is shared
  by every external part reached through the workbook relationship graph;
  independent operations retain independent budgets. The policy remains
  separate from physical `SourceCacheLimits` and cancellation state.
- Relationship and content-type targets are resolved and validated in the
  catalog/open phase. External-part payloads are parsed only in the deferred
  materialization phase, after relationship metadata has passed validation.
  Malformed or mismatched relationships are reported before publication rather
  than being silently omitted.
- Each part performs declared-size preflight before materializing its deferred
  payload and then charges the actual bytes consumed. Parsed record, string,
  token, matrix, cell, and retained-object counts remain governed by the
  external-link policy; limits never truncate or approximate external content.
- Reads and publication are fenced by the source version and cancellation
  state before access, between parts, after parsing, and immediately before
  publishing the immutable cache. Failed, cancelled, or stale work cannot
  publish a partial catalog. Concurrent callers use a single-flight immutable
  cache and share its published result.
- The source-backed formula context now matches the eager path for supporting
  links, external books, and external sheets, including the existing typed
  unresolved/unsupported outcomes. The retained data remains inert and no
  external target is activated, refreshed, evaluated, or recalculated.

## Validation evidence

Evidence was collected without making a performance, RSS, global/process,
concurrency, or absolute OOM claim:

- The focused XLSB library check passed.
- Focused external-link tests: `9/9` passed.
- Source-backed tests: `40/40` passed.
- The all-feature XLSB library suite passed `581/581` tests.
- The all-feature XLSB integration suite passed `135` tests. The exact known
  pre-existing failure
  `checked_in_unique_standard_drawing_corpus_transfers_every_anchor` was the
  only filtered test.
- Strict Clippy passed for all XLSB features, library code, and tests.
- Rustdoc passed with `RUSTDOCFLAGS="-D warnings"`.
- The minimal facade check passed with only the `xlsb` feature enabled.
- Crate-scoped formatting and whitespace checks passed.
- Every Cargo command ran serially with `CARGO_BUILD_JOBS=1` in the single
  isolated `/dev/shm/litchi-0333-target`, which was deleted after validation.

No benchmark or runtime profile was run because this batch establishes typed
correctness and resource-limit semantics, not a quantitative performance
claim.

## Residuals

- A failed `OnceCell` initializer may be retried by later waiters; this is
  bounded by the operation policy but is not a permanently memoized failure.
- Cancellation is checked between external parts, not at every parser token or
  byte, so a single large part may run until its next fence.
- PivotTable external-link semantics remain unsupported.
- The full bounded DDE/OLE cache is retained as inert data; it is not opened or
  activated.
- No external target activation is performed, including workbook opening,
  refresh, formula evaluation, recalculation, DDE, OLE, or add-in execution.
- Exact allocator overhead and caller-owned clones remain outside the
  operation budget.
- The known drawing-corpus integration test remains an existing skipped test
  outside this change's source-backed external-link scope.
