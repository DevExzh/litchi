# Change 0235: operation-scoped allocator evidence

Date: 2026-08-20

Status: schema and comparator evidence only; no latency or allocation claim

## Scope

The standalone performance harness now has a second target,
`litchi-perf-baseline-alloc`, enabled by the `allocator-metrics` feature. The
shared harness library and the default `litchi-perf-baseline` entry point both
retain unconditional `#![forbid(unsafe_code)]`; the normal target has no call
site that enables allocation metrics. Only the separate allocator entry-point
crate owns the narrowly scoped benchmark-only wrapper around
`std::alloc::System`, so enabling all package targets cannot instrument the
normal latency binary.

Filesystem child operations begin one non-overlapping measurement region just
before their timed operation and finish it before output verification, hashing,
reopen, or other correctness work. Absolute atomic counters cover successful
allocation, deallocation, reallocation, failed allocation, requested byte,
and live/high-water totals. They include allocations made by operation worker
threads. Regions publish checked differences without resetting any absolute
counter. Overflow and failed region acquisition retain typed status and omit
numeric vectors rather than publishing fabricated values.

The additive `operation_metrics.allocation` vectors are aligned with the
sorted `elapsed_ns.samples` vector. The report's tool identity records both
`binary: litchi-perf-baseline-alloc` and
`instrumentation: system_allocator_operation_scoped`, which prevents normal
and instrumented reports from being treated as the same implementation.
Allocator-instrumented elapsed values are validated for schema integrity but
are withheld from latency comparisons: `perf_abba_summary.py` rejects them for
latency ABBA, and `perf_compare.py` compares only matching allocation metrics.

## Verification boundary

This change does not claim operation peak memory: live and high-water values
are absolute before/after observations, and process-lifetime high-water state
is never reset. It does not claim copied, decompressed, recompressed, or
physical-I/O bytes. Allocation counters are optional and are emitted only by
the companion target; the existing raw filesystem evidence and normal latency
binary remain unchanged.

Focused Rust unit tests cover disabled-mode omission, status serialization,
cross-thread totals, non-overlapping regions, elapsed-vector alignment,
absolute live counters, overflow omission, strict schema cardinality, child
output, and actual allocator success/failure/reallocation paths. The latter
tests live only on the allocator target. Standard-library-only Python tests
cover comparator metric vectors, both normal/allocator policy identities,
refusal to use instrumented elapsed samples for latency ABBA, and the
source-level unsafe target boundary. No Cargo build or test is run as part of
the active benchmark change.
