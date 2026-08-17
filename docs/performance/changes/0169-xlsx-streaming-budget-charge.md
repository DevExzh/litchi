# 0169: Inline hierarchical budget charging for XLSX streaming creation

Date: 2026-08-17

## Decision

Retain commit `b79fd0480`.

`litchi_core::Budget::consume` no longer constructs and drops a temporary
owned node vector for cumulative charges. It walks the immutable child-to-root
chain by reference, performs the same atomic checked additions, and rolls back
the exact successfully charged prefix on failure. Releasable `Reservation`
charges still retain their charged nodes, but use four inline `Arc<Node>` slots
and spill to the heap for deeper caller-defined hierarchies.

The public API, error values, child-to-root charge order, atomic memory
ordering, commit/drop behavior, and unlimited hierarchy depth are unchanged.
The tradeoff is a larger live `Reservation` value (approximately 64 rather
than 40 bytes on this 64-bit build); hierarchies deeper than four nodes still
allocate. This is not a zero-allocation or total-memory claim for every budget
shape.

## Profile that selected the change

The existing `xlsx_streaming_create` large shape writes 131,072 rows and
524,288 cells through `StreamingWorkbookWriter`. Each row charges one Work
unit, each cell charges another, and accepted rows reserve/commit Objects.
The control Heaptrack process (three warmups and twenty samples, including one
untimed artifact build and exhaustive reopen) recorded 38,672,384 allocation
calls. Its dominant production stack was `Budget::charge` under
`StreamingWorkbookWriter::write_row`; the final matched profile removes
18,877,776 process allocation calls.

Clean sampled `perf record` reports attribute 10.52% of control cycles to
`Budget::charge`; 8.92 percentage points were cumulative `consume` calls. In
the candidate, cumulative `ExecutionContext::consume` is 4.62%, and no
`Budget::charge` allocation stack appears among the reported hotspots.
Deflate remains the dominant cost. These percentages are sampled attribution,
not additive stage timings.

## Correctness and boundedness

The focused budget suite covers:

- exact reservation release and partial/exact/failed commit behavior;
- child failure rollback at every charged ancestor;
- four-level inline retention;
- six-level spill, exact commit, cumulative rollback after five successful
  child charges, spilled reservation rollback, and zero consumption;
- concurrent reservation and cumulative-consumption limits.

The full `litchi-core` suite passes 161 unit tests plus eight `FileSource`
integration tests. Strict all-target Clippy passes with
`-D warnings -D deprecated`; rustfmt and diff checks are clean. The unchanged
streaming harness continues to build a complete artifact outside timing,
reopen it through `Workbook`, check every `A:D` cell, and require every timed
sample to reproduce the exact archive hash and sink counters. An independent
adversarial reviewer returned SAFE on the final inline, spill, rollback,
concurrency, zero-charge, and public-type behavior.

## Clean release ABBA

Both binaries were built from clean detached worktrees with locked release
dependencies. Every process was pinned to CPU 2 on the AMD EPYC 9575F host.
The order was `A1, B1, B2, A2`; each leg used twenty warmups and 200 samples
for each shape.

| Shape | Statistic | A1 -> B1 | B2 -> A2 | Accepted scope |
|---|---:|---:|---:|---|
| tiny | p50 | -6.63% | -8.32% | accepted |
| tiny | mean | -6.26% | -6.73% | accepted |
| tiny | p95 | -3.04% | -2.43% | accepted |
| tiny | p99 | **+1.81%** | **+2.75%** | withheld |
| medium | p50 | -8.23% | -7.08% | accepted |
| medium | mean | -7.68% | -6.50% | accepted |
| medium | p95 | -4.86% | -4.66% | accepted |
| medium | p99 | -4.60% | -1.05% | accepted |
| large | p50 | -8.21% | -8.58% | accepted |
| large | mean | -8.47% | -8.70% | accepted |
| large | p95 | -9.76% | -8.33% | accepted |
| large | p99 | -9.08% | -8.74% | accepted |

The predeclared same-implementation drift gates were 5% for p50/mean, 10%
for p95, and 15% for p99. All shapes pass. Tiny p99 is nevertheless withheld
because both paired candidate directions regress. Medium and large improve at
all four statistics in both directions.

All twelve records are clean and revision-exact. Across all legs, the three
archive hashes, decompressed worksheet hashes, row/cell counts, logical sink
write counts, accepted bytes, zero retained output bytes, and 4 KiB retained
row-authoring window remain exact. The timing scope includes deterministic row
text generation as well as writer construction, XML encoding, Deflate, ZIP
finalization, and hashing-sink publication; it excludes expected-artifact
construction, reopen/readback, sink construction, and digest extraction.

## Process resource sidecars

The matched large-shape `perf stat` processes (three warmups, twenty samples)
record:

| Counter | A1 -> B1 | B2 -> A2 |
|---|---:|---:|
| cycles | -8.28% | -7.81% |
| instructions | -6.55% | -6.59% |
| branches | -5.76% | -5.80% |
| branch misses | **+123.10%** | **+82.66%** |

The absolute branch-miss rate remains below 0.25%, but the regression is
disclosed and no branch-prediction improvement is claimed.

Matched Heaptrack processes record 38,672,384 -> 19,794,608 allocation calls
(-48.81%) and 22,545,902 -> 6,815,902 temporary allocations (-69.77%). Peak
heap is unchanged at 225.45M. Heaptrack covers the complete process and adds
profiler overhead, so the allocation totals are not operation-local.

GNU Time maximum RSS is A1/B1/B2/A2 =
252,772/252,128/252,900/252,396 KiB. The paired directions disagree. No RSS,
bounded-total-memory, physical-I/O, cold-cache, decompressed/recompressed-byte,
or real-producer result is accepted.

## Evidence

The compact [summary](../results/xlsx-stream-budget-charge-0169-summary.json),
[primary statistics](../results/xlsx-stream-budget-charge-0169-primary-stats.tsv),
and [comparisons](../results/xlsx-stream-budget-charge-0169-comparisons.tsv)
link the exact revisions, binaries, protocol, corpus/sink hashes, latency
decisions, process counters, and claim boundaries. The adjacent manifest binds
the four compressed raw reports, GNU Time sidecars, `perf stat` CSV files,
Heaptrack captures, sampled profile reports, tables, and summary.

This result is fresh one-sheet inline-scalar XLSX creation through the bounded
forward-only API. It is not existing-document CRUD, general multi-sheet XLSX
creation, shared-string/style/formula/date performance, a filesystem output
result, or evidence for every caller of `Budget`.
