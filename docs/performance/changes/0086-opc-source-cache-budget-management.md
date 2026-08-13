# Change 0086: OPC source-cache budget management

Date: 2026-08-13

Production revision: `d488ed128`

Status: implemented and correctness-tested; performance evidence pending

## Scope

Source-backed OPC packages can now opt into an explicit `ExecutionContext`.
The managed payload cache charges retained clean entries and active same-Part
load flights to that context's hierarchical memory `Budget`. Compatibility
constructors retain their existing finite `SourceCacheLimits` behavior.

The change preserves the cache's per-entry single-flight contract. A
reservation is acquired before payload I/O, follows the load through
publication, and remains attached to cached or returned managed payload
handles. Clean entries are evictable only when their payload is not externally
pinned. If pinned values prevent retention, a successful load may bypass the
cache without detaching the payload from its reservation.

Content-free diagnostics now identify whether the cache is budget-managed and
report reservation failures, observed budget use, cache/flight reserved bytes,
and the local memory limit. Managed `PartData` cannot be detached into an
unbudgeted raw allocation.

## Correctness evidence

Focused production tests cover:

- reservation acquisition and release on package/handle drop;
- cancellation on hits, waiters, loaders, and allocation-fallback paths;
- eviction of unpinned entries and preservation of externally pinned entries;
- rejection before payload I/O when the memory budget is insufficient;
- hierarchical parent limits and sibling-cache competition;
- same-Part waiter completion, loader failure/retry, and reservation identity;
- source-version changes, diagnostics bounds, and compatibility constructors.

These tests establish accounting, pinning, single-flight and failure behavior.
They are not performance measurements.

## Claim boundary and next evidence gate

No controlled contention, waiter-latency, eviction-rate, allocation, peak-heap,
RSS or throughput artifact has been retained for this implementation. No
latency or memory improvement is claimed. The next gate is a release-build,
CPU-pinned, controlled-contention matrix that records concurrent reads,
single-flight joins, cache hits/misses/bypasses/evictions, budget occupancy,
allocation counts and peak memory on fixed source/corpus revisions.
