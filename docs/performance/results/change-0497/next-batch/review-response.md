# Response to the 0497 next-batch review

This response addresses `docs/performance/results/change-0497/next-batch/review.md`.
The accompanying `implementation.patch` remains an unapplied proposal. No
production source, Cargo command, benchmark, or protected spec-gap checkout
was touched.

## Dispositions

| Review finding | Disposition in this draft | Evidence and remaining boundary |
| --- | --- | --- |
| Public forwarding seam does not compile | **Fixed in the proposal** | The public method calls `batch::read_parts_ordered(self, partnames)` directly. Compiler and clippy validation remain open until the next-batch owner applies the patch in an isolated review step. |
| Repeated metadata passes and silent `.ok()` policy fallback | **Fixed in the proposal** | `PreparedBatch` resolves each input occurrence once, stores the catalog index and declared size, checks per-Part and aggregate limits in input order, fences after the final record, and feeds both serial and parallel paths. No metadata error is converted into a scheduling choice. |
| Early fences and empty/count/setup errors | **Fixed in the proposal** | Source then context are checked before the count decision; metadata errors are fenced before exposure; empty calls allocate no scheduler state; count, preflight, workspace, and dispatch failures flow through the final source/context precedence function. |
| Aggregate limit and duplicate semantics were unspecified | **Fixed in the proposal** | `ReadLimits::max_total_part_bytes()` is used as a logical request sum. Duplicate occurrences count twice even if `PartCache` shares one payload. This policy needs an acceptance test before implementation. |
| Stop flag did not wake byte-gate waiters | **Fixed in the proposal** | The gate checks stop before admission and after every timed wake. Error/cancellation sets stop and broadcasts the condition variable. A permit is RAII and is dropped before a worker exits. |
| Scheduler counters used saturating arithmetic | **Fixed in the proposal** | Task and byte admission use checked arithmetic; release asserts the established invariant before subtraction. A counter violation is a scheduler failure. |
| Worker/spawn cleanup and panic boundary | **Fixed in the proposal** | The whole worker body is inside `catch_unwind`; spawn failures signal stop and explicitly join already-created handles; join panic is converted to the bounded adapter refusal. A provider that ignores cancellation can still delay join. |
| Scheduler memory proof was incomplete | **Fixed in the proposal** | Descriptor reservations precede descriptor allocation; serial shape excludes slots/handles/stacks; parallel shape charges slots, output, handles, fixed control/error state, worker state, and fixed scoped stacks with checked arithmetic; capacities are checked against charged envelopes. |
| Dynamic `OpcError` exceeded the fixed error envelope | **Fixed in the proposal** | Parallel workers store a fixed-size `StoredFailure` classification. Source/version, cancellation, read limits, resource limits, and I/O kind survive; dynamic provider/ZIP diagnostics become bounded content-free typed errors. Exact dynamic diagnostics across serial and parallel paths remain an explicit open gate. |
| Cache cleanup language implied rollback | **Fixed in the proposal** | Temporary descriptors, slots, permits, flights, and scheduler reservations must be gone on return. Clean payloads already published by successful workers may remain in `PartCache`; retained cache deltas are measured separately and are not promised to return to baseline. Cumulative `InputBytes`/`Work` remain monotonic. |
| Final error precedence was ambiguous | **Fixed in the proposal** | The order is source mutation, context cancellation, then setup/task error. If those fences pass, the lowest observed input ordinal wins. The same function handles preflight, workspace, spawn, provider, and final assembly errors. |
| Gate bounds were overstated | **Fixed in the proposal** | Documentation separates active logical tasks/declared bytes from cache-retained payloads, read-ahead bytes, provider buffers, and explicitly sized worker stacks. No broader RSS bound is claimed. |

## Open gates

The following are deliberately unresolved acceptance gates:

1. Apply the patch in an isolated owner worktree and run the compiler, clippy,
   and documentation checks. `git apply --check` is the only validation run
   for this review artifact.
2. Add executable tests for one prepared plan, source mutation during metadata,
   stop-aware gate waiters, cancellation, spawn/panic cleanup, source/context
   precedence, duplicate aggregate accounting, and retained-cache deltas.
3. Decide whether the public contract requires identical dynamic provider/ZIP
   diagnostics in serial and parallel modes. If it does, design a bounded
   operation-local error transport before implementation; this proposal does
   not make that claim.
4. Decide whether the 0497 brief's “same post-call budget counters” means
   rollback of cumulative `InputBytes`/`Work`. Literal rollback needs a
   child-context design and is outside this scheduler slice.
5. Confirm the fixed worker stack policy and resource receipt on supported
   platforms, and measure provider/read-ahead/cache ownership separately.
6. Establish serial-equivalence and resource/cancellation receipts before any
   benchmark. No throughput, scaling, or Amdahl conclusion is made here.

The next-batch proposal is ready for that review sequence once these gates are
accepted; it is not a production implementation or a performance result.
