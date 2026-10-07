# 0948 prospective XLSX multi-window qualification protocol

This is the reviewed protocol draft for a prospective comparison of the
current publication path and one candidate that may activate bounded
replacement-window proof. It is not measurement evidence. The packet remains
`status: draft` until the candidate source, binaries, build receipts, semantic
oracles, and exact changed-file manifest are bound by the root coordinator.
No measured workload or profiler process may run while the packet is in this
state; root-authorized controlled builds are governed by the v2 execution
clarification below.

The comparison reuses the four source-backed XLSX rows from 0945:

* medium one-percent edits and dense-sparse one-percent edits are the primary
  rows;
* unmanaged one-edit medium and managed one-edit medium are controls.

The baseline is the current root at `b6c5c80ffb` and its complete source export.
The prior 0945 candidate binary may be reused only after the root records an
exact selected-source, build, lockfile, binary, and workload identity check.
The packet-local binding for that reuse is the versioned
`receipts/baseline-reuse-v2.json`; it retains strict descriptors for both
0945 candidate build chains, logs, wrappers, source manifests, freeze receipt,
and the historical closure receipts. The earlier v1 receipt remains retained
as historical input and is not an admissible capture binding.
The candidate role is a separately bound source/build namespace. A candidate
must preserve the complete original-first audit, namespace and limit checks,
source freshness, cancellation, typed errors, and the full fallback audit.
Window metadata is operation-local and untrusted until bounds, ordering,
disjointness, source/output span equality, and outside-gap identity have been
checked. A diagnostic activation path cannot weaken those requirements.

Qualification has 16 children (native and allocator, both roles, all four
rows), zero warmups, and two samples. Formal native capture has six
alternating paired blocks, ten warmups, and 100 samples per child. Formal
allocator capture has three alternating paired repeats, two warmups, and five
samples per child. Even blocks/repeats run baseline then candidate in forward
case order; odd blocks/repeats reverse both orders. Children run serially on
CPU 12 with `PATH=/usr/bin:/bin`, `LANG=C`, `LC_ALL=C`, `TZ=UTC`, and
`RUST_BACKTRACE=0`; `/usr/bin/time` records maximum RSS. A child is never
retried or replaced.

Each newly built role uses the 0945 durable build contract under a 0948 schema;
the candidate wrappers are `build-tools/build_candidate_native_v3.py` and
`build-tools/build_candidate_allocator_v3.py`, with source rooted at
`candidate-build-v3` and target/log namespaces ending in `candidate-{kind}-v3`.
Cargo must run exactly the locked, offline release command with `-j2`,
toolchain `1.95.0`, release codegen units `1`, debug info `1`, fat LTO, and
optimization level `3`. The receipt binds the complete environment and unset
variable list, packet-relative source manifest and freeze/export descriptors,
runner and Cargo executable digests, stable relative target/binary/log paths,
and the `bindings -> start -> launch -> terminal -> binary/result` chain. The
terminal is written only after process-group reaping and clearing; its logs,
binary digest, source census, and receipt references are checked before a role
can enter capture. Absolute command paths are retained as custody data and
validated by stable suffixes so a copied packet can be replayed.

The native latency gate is fixed before capture. For each primary row, the
candidate/baseline publication p50 paired-median ratio must be at most 0.95,
and its deterministic two-sided 95% paired bootstrap upper endpoint must be
strictly below 1.00 (10,000 draws, seed 948009). Every paired block and every
row must have workflow p50 and p95, and publication p50 and p95, at most 1.05.
These per-block checks include both controls. The aggregate result retains
block medians, raw vectors, p50/p95/p99 quantiles, ratios, bootstrap endpoints,
and spread/regression review flags.

Allocator measurements are safety checks for this parser-work candidate, not
an improvement target. Every repeat and row must provide all five 0945
metrics: `allocated_bytes`, `allocation_calls`, `region_peak_live_bytes`,
`live_bytes_after`, and `max_rss_kib`. Each metric requires matching baseline
and candidate vectors, the expected sample cardinality, nonnegative integer
values, and valid allocator conservation/oracle status. Each per-repeat ratio
and the aggregate paired-median ratio must be at most 1.05. Missing metrics,
unavailable instrumentation, mismatched vectors, or a nonzero candidate over a
zero baseline fail closed. Equal zero pairs are retained as an explicit
degenerate case. Three repeats are not treated as an allocation confidence
interval; all raw values and material flags remain in the reduction.

The formal reducer must reject a changed plan, changed helper, overwritten
output, duplicate child, missing terminal, malformed report, failed semantic
oracle, source/build mismatch, or unexpected artifact. Helpers and outputs are
versioned and create-only. All packet paths are relative and symlink-free;
historical absolute paths may occur only inside retained custody records.
Negative controls cover corrupt XML or report descriptors, duplicate job
identity, traversal/symlink paths, wrong source or binary hashes, terminal
sealing, and output overwrite. Relocation replay must compare a normalized
semantic projection because absolute packet paths are expected to differ.

The optional multi-window activation diagnostic is separate from the primary
qualification. It may report changed-span coverage and proof-choice counts,
but it cannot select rows, alter thresholds, replace the formal workload, or
turn a failed native gate into a pass. No parser-check disable, readback
substitution, planning-parser cache, cross-operation cache, or stopped-observer
shortcut is admissible.


## Versioned premeasurement execution clarification

The draft prohibition applies to measured workloads, profiler runs, and
qualification capture. Root may explicitly authorize controlled baseline or
candidate builds before the final plan freeze when those builds are required
to produce the bindings. Each such build must use its versioned wrapper and
write the complete start, bindings, launch, terminal, binary, result, and log
receipts under create-only paths. The frozen plan binds the resulting source,
binary, helper, and receipt descriptors before any workload or profiler starts;
this clarification changes execution sequencing only and does not change a
case, sample schedule, threshold, or failure rule.
