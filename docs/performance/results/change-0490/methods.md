# 0490 controlled file-store follow-up

0489 retained a small file-store repeat with +79.00/+82.60% normal p95/p99
and +35.24/+51.74% allocator tails. The other normal repeat improved. This
batch resolves whether the observed regression repeats under interleaved
before/after execution and measures synchronization attribution separately.
The file-store implementation is harness-owned and unchanged by 0489.

## Fixed formal comparison

Use the retained 0487 and 0489 normal and allocator executables, whose build
recipes, source manifests and copied binary hashes are authenticated. No
production or benchmark Rust change and no rebuild is planned. The only
workload is 64 source paragraphs and 64 authored paragraphs with the existing
short-text, 64-chunk, owned-source, current-compression, 4,096-byte sink contract.

Six blocks contain both binary versions, both allocator roles, and all three
routes: deterministic, memory store and data-synchronized file store. Odd
blocks execute before then after; even blocks reverse versions, roles and
routes. Each child has 60 measured samples after five warmups. The complete
formal inventory is 72 children and 4,320 measured samples. The process is the
replication unit; the analysis retains each block and uses a seeded block
bootstrap for descriptive uncertainty, rather than treating all samples as
independent replicates. Every adverse latency, heap or RSS observation remains
visible. RSS is one whole-child observation per process.

The frozen 0484 argv and report validators retain source, authored, candidate,
preservation, output, replay, allocation and exact data-sync checks. Replay
scratch receives a fresh per-child directory; successful runs must leave it
empty before removal. The benchmark's cleanup is outside the timed operation,
but may affect the next operation's filesystem state. Both the benchmark and
its preflight still check complete output identity and reversible publication.

## Separate synchronization diagnostics

Four normal-build sync-only traces use two repeats with reversed before/after
order, 30 samples and three warmups. With exactly one file-store preflight,
each must contain exactly 34 successful fdatasync calls on the same caller
replay pathname. The first event is preflight, the next three are warmups,
and the remaining 30 align in order with reported samples. The parser rejects
missing/extra calls, foreign paths, failures, unfinished calls and reordered
events. Each sync duration must fit within its containing timed operation.

Four additional normal-build syscall-summary children use one sample and one
warmup, observing metadata, reads, writes, file lifecycle and synchronization.
These include setup and oracle work. Tracing perturbs timing; these eight
children never enter the formal latency comparison, and their elapsed samples
are not substitutes for untraced results. Correlation or a high synchronized
fraction supports attribution only for these observed diagnostic operations.
It cannot prove the historical 0489 outlier's cause retrospectively.

All benchmark/profiler children pin to CPU 2 and the two drivers share the
existing CPU lock. No unrelated job is stopped and no host cache is dropped.
The machine record is a point-in-time observation of an ordinary shared host,
not evidence of reservation or physical cold reads. This batch makes no cold,
parallel-scaling, atomic-save, or universal file-store performance claim.

## ADR and goal scope

All accepted ADR hashes match the previously read set. ADR 0001/0005 require
reproducible measurements, honest uncertainty and regression review. ADR
0003/0006 retain source authentication, independent initial audits, candidate
proofs, preservation, cancellation and resource bounds. This experiment keeps
`--replay-sync data`; it neither skips synchronization nor weakens freshness.
There is no runtime dependency, public API, unsafe-code or container ownership
change. Helper tests and independent source/measurement review cover this
measurement-only batch; the unchanged Rust source remains bound to 0489's
passing validation evidence.

The full non-iWork goal remains open. The next implementation priority is
verified cold/source-provider coverage, followed by bounded concurrency and
independent native Office cases. This regression follow-up does not substitute
for those completion requirements. Batch-local unused scratch is removed;
retained evidence and the already authenticated executables remain.


## Verification checkout

Both drivers verify the current production source inventory against the retained
0489 manifest before reusing its binaries. Run the evidence verifier from the
recorded 0490 checkout (or an identical source tree). A later production edit
requires checking out the recorded source; it does not make those historical
measurements a result for the new revision. The final seal inventories every
retained file, including unused pre-review helper/protocol artifacts, and
recomputes both formal and diagnostic summaries. Only the seal and its optional
verification receipt are excluded from that inventory.
