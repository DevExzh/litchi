# 0820 quality repair review — allocator test gate

This is an unfrozen repair review for the failed 0820 quality attempt. It does
not alter the frozen protocol/source review and does not admit any timing,
build, export, or capture result.

## Failure and root cause

The first 0820 quality attempt passed `cargo fmt` and `cargo check`, then gate
3 (`cargo test --offline --locked --all-features -- --test-threads=2`) failed
only in the `docx_replayable_tail_append` test binary. The failing helper was
`src/bin/support/counting_allocator.rs`:

```text
assertion failed: during.live_bytes >= before.live_bytes + 256
```

The following zeroed-allocation helper then reported `PoisonError` because the
first helper's panic poisoned its local `TEST_LOCK`. The failure is a test
isolation error. The helper's lock serializes the five allocator tests in that
module, but it does not serialize the benchmark tests in the same binary or
allocator activity performed by the test harness. With two test threads, an
unrelated deallocation can occur between the `before` and `during` snapshots.
The process-global live counter is therefore not required to increase by the
test allocation across those boundaries.

The failure does not indicate a production allocator or counter arithmetic
defect. The test directly calls the wrapper, and the cumulative allocation and
deallocation counters remain valid evidence for that delegation check.

## Admissible repair boundary

The repair is limited to `#[cfg(test)]` assertions in
`tools/perf-baseline/src/bin/support/counting_allocator.rs`. It must preserve
the `GlobalAlloc` implementation and every `allocation_metrics` production
counter path. The settled draft removes the fragile global `live_bytes` delta
assertion from the wrapper smoke test and replaces it with a signed
conservation check plus high-water invariants. `snapshot()` and each counter
callback serialize through the same observer mutex, so the conservation check
is stable across unrelated callbacks that occur between the boundaries. The
allocation-call, allocated-byte, deallocation-call, deallocated-byte, and
historical-peak assertions remain in place. The helper is test evidence about
the counters observed by this process; it does not claim complete observation
when the production observer marks a sample unavailable. It must not weaken the
wrapper's pointer ownership, zero-initialization, reallocation,
failed-allocation, or failed-reallocation checks.

`tools/perf-baseline/src/allocation_metrics.rs` already has isolated counter
coverage for exact live/peak behavior: one allocation/deallocation, growth and
shrink reallocation, failed allocation, multi-step live reference,
concurrent overlapping allocations, operation-region entry/exit, retained
entry allocations, cross-thread totals, split regions, overflow, poison, and
reentry. Those tests use fresh `Counters` instances or the library test lock,
so they are the appropriate deterministic coverage for exact live-byte
arithmetic. The test-only conservation helper does not replace that isolated
coverage and makes no production change.

## Deterministic regression outline

The repair review should be closed only after a fresh repair quality run shows:

1. The changed file is limited to the test-only helper, with no production or
   harness source change and no frozen 0820 packet input mutation.
2. The focused `docx_replayable_tail_append` allocator tests pass under the
   same two-thread setting that failed, including all five helper tests and no
   poisoned lock.
3. The full six-gate quality sequence is rerun from fresh repair witnesses;
   the failed attempt remains retained as evidence and is not relabeled as a
   pass.
4. The isolated `allocation_metrics` live/peak tests remain present and pass,
   proving that the repair removed test scheduling sensitivity rather than
   weakening counter semantics.

No Cargo, binary, workload, or repair quality command was run for this review.
