# Ordinary ODP append phase attribution

This experiment measures the existing owned Snapshot/Transaction/Commit/Patch
append lifecycle on the deterministic 64-, 4,096-, and 8,192-slide corpus from
0457. It separates snapshot opening, transaction construction, append, commit,
and sequential output at public API boundaries. A separate unsegmented mode
measures the same lifecycle to expose instrumentation overhead. No production
clock, serializer, validation rule, source ownership, or patch behavior changes.

The matrix is two repeats of normal and allocator binaries, each with three
warmups and 30 samples for both modes and all three shapes: 24 reports and 720
retained operations. CPU affinity is 2 and workload concurrency is one. Input
cloning, append strings, sink construction, fixture setup, result checks, and
final drops are outside the lifecycle interval. Phase peaks cannot be added:
each phase starts with the objects retained by earlier phases.

The frozen capture passes all 24 lanes and 720 samples. Two separate large
normal-binary profiles retain another 100 samples each. `summary.json` and
`profile-summary.json` are recomputed from authenticated raw artifacts.
`prior-control-bindings.json` authenticates copied prior reports only for
corpus/output identity; their timings are not this experiment's baseline.

## Results

Normal-binary median phase shares are computed within each sample, then
summarized; independently rounded medians need not sum to 100%.

| Source slides | Snapshot opening R1 / R2 | Transaction R1 / R2 | Commit R1 / R2 | Phase envelope p50 R1 / R2 |
|---:|---:|---:|---:|---:|
| 64 | 19.82 / 19.66% | 20.30 / 20.48% | 58.68 / 58.63% | 1.904 / 1.909 ms |
| 4,096 | 27.09 / 27.07% | 25.88 / 25.99% | 45.50 / 45.43% | 75.984 / 75.649 ms |
| 8,192 | 27.04 / 26.89% | 25.73 / 25.96% | 45.64 / 45.55% | 152.056 / 152.334 ms |

On large sources, append itself accounts for about 1.55% and sequential output
about 0.03%. Commit is the largest individual phase; opening and transaction
construction together consume slightly more than half the measured phase time.
Commit allocates 118,096,476 of 211,442,207 lifecycle bytes (55.85%) on large
sources. Snapshot opening allocates 53,417,100, transaction construction
28,217,057, append 11,711,574, and publication zero bytes. Both repeats agree
exactly. All six allocator comparisons have exact equality between summed
phase allocation volume and direct lifecycle volume. Peak live values have
different entry baselines and must not be added.

Observed phase-envelope versus direct-lifecycle p50 deltas span -0.655% to
+3.448% across all instrumentation/shape/repeat combinations. None of the
p50/p95/p99 or process-lifetime maximum-RSS comparisons exceeds an absolute
5% difference. These are instrumentation comparisons, not a production
optimization. The independent 10,000-resample median-ratio intervals are in
`summary.json`; they do not account for systematic run-order effects or
multiple comparisons. Three warmups and 30 samples do not establish cold-input
or tail-latency guarantees.

## Profiling and next work

All seven whole-process perf counters were available with reported 100%
running time; whole-process IPC is 4.159. The separate 99 Hz cycles:u recording
contains 1,656 parsed samples and no malformed samples. **All sampled periods
remain unattributed to phase markers.** The bound executable contains all five
non-inlined marker symbols, but the captured stacks do not reach them. Retained
stacks frequently end in unknown frames. This recording supports flat symbol
observations only; it cannot rank internal commit stages or attribute hardware
counters to phases. Setup, warmups, checks and reporting are included.

The next measurement target is the ordinary commit path, followed by snapshot
opening and transaction construction. The [source audit](source-audit.md)
identifies repeated content validation/provenance work as one hypothesis;
current data does not establish it as the dominant internal cost. A separately
bound diagnostic build with reliable unwinding or harness-only internal stage
experiments is needed before selecting a production optimization. Candidate
readback, compact XML proof, fallback validation and reversible patch behavior
remain required.

## Validation and replay

The harness suite passes 387 tests with one ignored. Strict Clippy, release
build, formatting, warning-denied rustdoc, and crate boundaries pass. Four
actual-report controls and fourteen oracle mutations pass. Two initial smoke
failures are retained and explained in [integration notes](integration-notes.md).
The [harness review](harness-review.md) found no blocking issues.

`python3 -B verify.py --precleanup` authenticates source and retained executable
custody before cleanup. `python3 -B verify.py --portable` verifies the complete
sealed bundle and recomputes both summaries using only retained files. The
finalization receipts record a fresh-copy replay and inventory-based removal
of the two owned staging executables. Run `seal.py` between receipt-producing
finalization steps; never alter raw capture inputs or reports. The full non-iWork goal remains
open, including selective updates, source/input matrices, native application
roundtrips, and measured bounded-worker scaling.
