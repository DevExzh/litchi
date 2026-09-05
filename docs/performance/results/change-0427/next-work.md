# 0427 next work: fail-closed PPTX cache phases

0427's frozen production scope has no optimization and
`production_changes=false`. The immediate checkpoint may retain scoped
allocator live-byte observations; it does not authorize cache, managed-budget,
RSS-release, retention, or causal optimization claims. The full distinction is
recorded in the [retention acceptance gate](checks/retention-acceptance.md).

## Smallest reviewed follow-up

The smallest production follow-up is an additive fail-closed diagnostic seam
on both public source-backed PPTX owners:

- `SourceBackedPresentation::cache_diagnostics` currently forwards the
  infallible compatibility snapshot at
  `crates/litchi-pptx/src/presentation/source.rs:1106-1110`.
- `SourceBackedPresentationEditor::cache_diagnostics` does the same at
  `crates/litchi-pptx/src/presentation/source.rs:1391-1398`.
- The OPC owner already provides
  `SourceBackedPackage::try_cache_diagnostics` at
  `crates/litchi-opc/src/source_backed.rs:4235-4251`. It returns
  `SourceCacheDiagnosticsError` for invalid diagnostic state instead of
  converting it into a usable-looking snapshot. Counter intervals must use
  `SourceCacheDiagnostics::checked_counter_delta`, defined at
  `crates/litchi-opc/src/source_backed.rs:1157-1185`.

The next source change should forward that fallible result through the PPTX
owner, preserving the existing infallible compatibility method for ordinary
callers. The benchmark and retention verifier must use only the fallible
method and reject an error; it must not fall back to the compatibility
snapshot. Directly calling the OPC package from the harness would bypass the
PPTX source/ownership lifecycle and is not equivalent evidence.

## What already exists

The follow-up does not need a new budget-configuration API. Both PPTX owners
already expose the managed constructor family:

- `SourceBackedPresentation::from_read_at_with_limits_and_cache_limits_and_execution_context`
  at `source.rs:928-944`;
- `SourceBackedPresentationEditor::from_read_at_with_limits_and_cache_limits_and_execution_context`
  at `source.rs:1346-1363`.

The filesystem aliases forward the same policies at `source.rs:747-761` and
`:1241-1256`. The no-cache-limit execution-context constructors retain the
default finite cache, while the explicit `SourceCacheLimits` constructors let
the capture choose the cache ceiling. `SourceBackedPackage` constructs the
managed cache with the supplied `ExecutionContext` and limits at
`crates/litchi-opc/src/source_backed.rs:3882-3915`. Thus the missing piece is
fail-closed observation through the format owner, not a new way to set the
budget.

## Bounded validation plan

After the additive forwarding seam is available, keep the validation focused:

1. Add one read-only PPTX forwarding test for both public owners. A managed
   open, selected payload read, and drop must return the same content-free
   fields through the fallible method as the owning OPC snapshot. Existing OPC
   tests remain the authority for poisoned-state and counter-overflow error
   construction; the PPTX test proves that those errors are not hidden by the
   facade.
2. Verify each phase with two fallible snapshots and
   `checked_counter_delta`; reject a missing snapshot, a counter regression, or
   any `SourceCacheDiagnosticsError`. Retain gauges such as
   `retained_entries`, `retained_bytes`, `budget_memory_used`,
   `budget_cache_reserved_bytes`, and `budget_objects_used` as point values,
   not interval deltas.
3. Reuse the existing 0424 plain/media source-backed lifecycle and exact
   source/output oracles. Capture baseline, plan, publication/reopen, explicit
   transient drops, and final owner drops only after the fail-closed seam is
   present. Keep allocator callback points, cache gauges, managed budget, and
   RSS in separate scopes.
4. Add the near-limit rows separately: exact admission, one-byte-under refusal
   before payload I/O/output, pinned-handle retention, clean-entry eviction
   after handle drop, oversized bypass, and one repeated publication at the
   configured limit. These rows require exact cache/budget configuration and
   expected source-read, output, refusal, and reservation counters.

For any control/candidate resource comparison, bind the reports to the same
source, corpus, output sink, cache limits, execution-context limits, binary,
observer revision, and protocol. Preserve raw journals and phase ownership
sets. Normal elapsed timing remains separate from allocator and cache vectors;
the forwarding seam and checkpoint do not create a latency claim.

Until this follow-up and its bounded capture pass, 0427 can report only its
scoped allocator callback observations. Cache and managed-budget phase
evidence remain open, and the full non-iWork goal remains incomplete.
