# 0513 XLSX metric-enabler source review

Reviewed the current `tools/perf-baseline` diff for the four existing XLSX
update cases: one-cell and one-percent commit, and their commit/save variants.
The implementation is a measurement enabler only; it does not change XLSX
format behavior or production library code. I found no source-level
correctness blocker.

## Region and timer boundaries

`run_xlsx_update_commit` and `run_xlsx_update_commit_save` acquire an
`allocation_metrics::Region` immediately before `Instant::now()`. The elapsed
value is read first, then the region is finished. Workbook opening, edit
preparation, expected-output construction, sink allocation and reservation,
semantic/output oracles, metric-envelope construction, and teardown are
outside the timed operation. Warm-up iterations execute the same operation but
are omitted from both `elapsed` and `observations`.

The save-only profiling boundary is the private `#[inline(never)]`
`xlsx_commit_save_operation`. It performs one `Edit::commit` and one
`Commit::workbook().write_to` call, then returns the successful `Commit` to the
caller. The caller retains that value through sink validation and workbook
readback, black-boxes it after readback, and drops it outside the measured
region. The expected-output commit therefore cannot be mistaken for the
profiled save operation.

On commit or writer failure, `?` propagates the error. A failed helper drops
its local `Commit` while the caller's region guard is still active, and the
guard then releases without publishing a failed sample. A failed direct
commit similarly drops the region guard, so no active-region token or stale
observation remains. Successful paths finish the region before any semantic
checks or caller teardown.

## Availability and alignment

The normal binary does not enable the allocator observer. Its disabled region
therefore returns `None`, which the runners normalize to an explicit
`allocation_metrics::unavailable_sample()`. This preserves unavailable status
and omits numeric vectors instead of reporting fabricated zero allocation
counts.

Only retained samples are inserted into the observation vector. Each retained
observation receives the same checked `elapsed_ns` value pushed into the
caller’s elapsed vector. `from_in_process_observations` applies the existing
`(elapsed_ns, sample_index)` ordering, so `operation_metrics.sample_indices`
matches `elapsed_ns.sample_order`, including ties. Save cases construct their
operation sink view from the already deterministic top-level sink summary;
the later promotion pass is idempotent and preserves the sink compatibility
summary.

## Allocation peak interpretation

`region_peak_live_bytes` is an absolute process live-byte high-water mark. The
allocator implementation checks that it is at least both live-byte endpoint
snapshots and no greater than the process high-water value after the region.
For measured samples only, `region_peak_live_bytes - live_bytes_before` may be
reported as checked incremental live-byte demand above the operation entry
point. It must not be called an operation, process, document, or RSS peak; it
can include other-thread callbacks and retains absolute-counter semantics.
Unavailable or overflow samples must not be subtracted or converted to zero.

The focused regression tests cover all four cases with warm-ups, normal-mode
unavailable allocation status, aligned sample vectors, sink promotion, exact
commit/save output and semantic verification, and propagation of a bounded
sink write failure. The external profile requirement is now resolved by [scope-review.md](scope-review.md):
the raw Callgrind edges show three direct helper-to-`Edit::commit` calls and
three writer calls, with fixture and expected-output work excluded. Allocator elapsed
time and RSS remain descriptive and are excluded from performance claims.
