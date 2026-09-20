# 0726 XLS empty-slot setup source review

## Scope and verdict

This is a read-only review of the installed 0726 candidate against baseline
`3ad29e42da`. The live production diff is limited to
`replay_indexed_cell` in `crates/litchi-xls/src/workbook/source.rs`:

1. keep the existing leading execution check and worksheet lookup;
2. call `index.slots_for(row, column)` immediately after that lookup;
3. when the returned slice is empty, perform the existing trailing execution
   check and `owner.ensure_current()` before returning `Ok(None)`; and
4. leave the non-empty resolver, chain hints, selected-frame loop, and trailing
   fences unchanged.

The live source diff is **ACK**: it removes setup that an indexed missing query
cannot use while keeping the cache semantically invisible, and it satisfies the
ordering assertions above. This review is not performance evidence and does
not replace the fresh 0726 build, correctness, allocation, source-I/O, and
timing gates.

The installed source was checked directly against the baseline revision; the
diff contains only the described `source.rs` hunk. No Cargo command, native
probe, or profiler was run.

## Accepted-ADR constraints

The installed shape is compatible with the accepted records relevant to this
path:

- ADR 0001 requires correctness and safety to win performance, typed errors to
  remain typed, and ordinary public APIs to stay free of archive/source
  implementation details. This is a private replay-path change with no public
  signature or dependency change.
- ADR 0005 requires immutable positional-source semantics, stable source
  identity/version fencing, lazy optional cache state, explicit execution
  contexts, and measured evidence for optimization claims. The empty path
  must therefore retain `owner.ensure_current()` even though it reads no
  worksheet bytes, and the source change cannot be retained from this review
  alone.
- ADR 0006 requires validation and preservation behavior to remain unchanged.
  An empty slot range is safe to use only because the published worksheet
  index represents a complete successful scan; the early return must not
  become a partial-scan or decoded-value cache.
- ADR 0008 makes focused correctness and independent evidence gates part of a
  production change. The existing tests below cover the observable contract;
  the 0726 qualification must still establish the isolated performance result.
- ADR 0024 keeps XLS source-backed worksheet ownership in `litchi-xls`; the
  installed change does not move ownership or add a crate edge.
- ADR 0031 keeps cancellation in the caller-supplied `ExecutionContext`. The
  empty branch must retain its final `context.check()` in the same position as
  the existing replay path.

## Required source assertions

The candidate should be rejected from qualification if any of these are not
true:

- The first operation in `replay_indexed_cell` remains the existing optional
  `execution` check. The worksheet lookup remains before the empty-slot return,
  preserving worksheet error precedence and the established replay shape.
- `slots` is computed before the `workbook_path` `Vec` is built. The empty
  branch returns only after both the final optional `context.check()` and
  `owner.ensure_current()` succeed. It must not construct `refs`, a shared
  string resolver, or either CFB chain hint on that branch.
- The non-empty branch iterates over the already computed `slots` slice. It
  retains the existing `SharedStringResolver`, SST checkpoint setup, worksheet
  chain hint selection, per-slot execution checks, selected-frame decoding,
  final execution check, and trailing source fence. It must not change
  duplicate ordering or target-specific error precedence.
- `query_cache.rs`, `INDEX_OVERHEAD`, index admission/publication, scanner
  validation, and the public API remain byte-for-byte unchanged. The only
  intended production source delta is the one replay function described above;
  no source bytes, decoded values, or errors are retained by the cache.

The existing `WorksheetCellIndex::slots_for` implementation is a sorted-slice
range lookup and returns a borrowed slice without allocation
(`crates/litchi-xls/src/workbook/query_cache.rs:51`). The unit test
`cell_slots_are_grouped_by_coordinate` also asserts that an empty index returns
an empty range (`query_cache.rs:1059`).

## Existing tests and what they prove

The following tests are the relevant custody for the installed early return:

- `a_successful_warm_index_shortcuts_later_hits_and_missing_cells`
  (`crates/litchi-xls/tests/xls_query_index_cache.rs:586`) publishes a warm
  index, queries an absent coordinate, and asserts the result is `None` with
  zero source bytes and zero source reads. It also checks a later source change
  is reported as `SourceChanged`.
- `an_indexed_missing_target_still_takes_the_trailing_freshness_fence`
  (`xls_query_index_cache.rs:756`) is the decisive trailing-fence test. It
  arms `bump_before_observation(2)`, queries an indexed-missing coordinate,
  expects `SourceChanged`, and asserts zero reads. The injected change therefore
  cannot be detected by a worksheet scan or only by the leading query fence;
  the early branch must call its trailing `owner.ensure_current()`.
- `a_warm_index_keeps_alternating_first_last_missing_and_random_values_identical`
  (`xls_query_index_cache.rs:1193`) compares an enabled warm index with the
  disabled-cache path over alternating present and absent coordinates. It
  asserts the same values, including `None`, and zero reads for the indexed
  missing query.
- `cloned_handles_can_replay_warm_index_concurrently_with_local_hints`
  (`xls_query_index_cache.rs:1270`) exercises a missing warm query alongside
  present warm queries on cloned handles, preserving the snapshot-local cache
  behavior.
- `disabled_cache_preserves_values_and_repeats_the_complete_scan`
  (`xls_query_index_cache.rs:565`) protects the comparison path: disabling the
  optional cache still repeats the complete validated scan with identical
  source metrics.
- `execution_variants_honor_pre_and_mid_scan_cancellation`
  (`crates/litchi-xls/tests/source_backed.rs:2119`) covers the public execution
  API's pre-operation and source-read-triggered cancellation behavior, while
  `cancellation_during_index_build_drops_the_candidate`
  (`xls_query_index_cache.rs:1340`) covers cancellation during optional index
  construction and confirms that a cancelled candidate is not reused.

The stored-target path has separate coverage for the behavior that this change
must leave untouched, including `formula_string_continuations_have_the_same_value_on_warm_hits`,
`an_earlier_shared_string_read_error_still_refuses_a_warm_duplicate_hit`,
`an_earlier_malformed_formula_string_still_refuses_a_warm_duplicate_hit`,
`a_valid_shared_string_checkpoint_preserves_duplicate_read_error_precedence`,
and `visit_cells_keeps_duplicate_order_after_queries_warm_the_owner`.

## Coverage conclusion

No additional runtime test is required for the source-version behavior: the
indexed-missing trailing-fence test already distinguishes the required final
`ensure_current()` from an unconditional `return Ok(None)`, and the warm/cold
tests cover the result and zero-read behavior. The final execution check in the
empty branch is not separately distinguishable with the current public test
source because that branch intentionally performs no read; a timing-based
cross-thread cancellation test would be nondeterministic. The source review or
independent audit should therefore assert the branch text/order directly:
`context.check()` and then `owner.ensure_current()` must both precede
`return Ok(None)`.

If the candidate changes any non-empty replay code, adds a query-cache layout
change, or drops either empty-branch fence, this ACK no longer applies and a
new focused review is required.
