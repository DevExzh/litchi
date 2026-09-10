# 0499: operation-local worker reuse for ordered Part batches

The existing `SourceBackedPackage::read_parts_ordered` operation now reuses
one bounded set of operation-local scoped workers when a request spans more
than one admitted wave. Each worker has a one-slot command queue and a
one-slot reply queue; replies are collected in input order, every admitted
worker is shut down and joined, and the lowest input ordinal remains the
typed-error selector. A request whose first wave contains the complete batch
keeps the existing single-wave path, so reuse is only admitted when it can
remove later thread creation. There is no global pool, runtime dependency, or
automatic facade parallelism.

The scheduler preserves the 0498 contract. Source and cancellation fences run
before each wave and at the final boundary; later waves cannot begin until the
current wave has completed successfully. Worker, task, declared-byte, and
request limits remain enforced, output ownership stays with `PartBatch`, and
actual `Work` and `InputBytes` accounting remains monotonic. Scheduler
admission reserves the worker stacks, handle and channel endpoint vectors,
active-ordinal storage, and object state before worker startup. The worker
stack reservation is two MiB per worker; the channel reservation uses a
conservative four KiB per-worker control allowance and is not a portable exact
allocator-size claim.

The loader now has an unwind guard for a provider panic after a cache flight
has been admitted. The guard only holds another reference to the existing
flight, releases failed-flight reservations, removes the pending flight
through the existing cache completion path, and disarms on ordinary return
without an extra success-path cache lock. Worker panic recovery and this guard
are unwind-build behavior; the ordinary release profile uses `panic = abort`.

Focused coverage exercises worker identity reuse across several waves,
operation-local isolation between separate batches, typed provider-panic
termination and retry, and duplicate-waiter recovery after a gated provider
panic. Four new reuse/error tests pass alongside the 14 existing focused
tests. The default OPC gate has 623 passing tests and one ignored external
corpus test; the all-feature gate has 645 passing tests and one ignored test.
Five doctests, warning-denied Clippy and rustdoc, formatting, crate boundaries,
and downstream DOCX/XLSX/PPTX/XLSB checks also pass.

## Measured behavior

The matched capture uses the frozen 0498 executable and unchanged harness:
sixty children, 3,600 measured samples, and 360 warmups across two corpora,
three sources, serial reads, and batch widths 1/2/4/8. Byte, source-counter,
`Work`, `InputBytes`, released-budget, and scratch-cleanup oracles match in
both phases. The table gives nearest-rank aggregate p50 latency in
microseconds for the multi-wave local-read rows:

| Source | Batch 2 before → after | Batch 4 before → after | Batch 8 before → after | After serial p50 |
| --- | ---: | ---: | ---: | ---: |
| many-small owned | 669.84 → 186.51 (-72.16%) | 542.70 → 186.04 (-65.72%) | 475.25 → 228.74 (-51.87%) | 95.86 |
| many-small warm file | 780.59 → 296.04 (-62.07%) | 562.99 → 248.77 (-55.81%) | 496.17 → 273.70 (-44.84%) | 182.04 |

Worker reuse removes most of the matched multi-wave local-read latency, but
the resulting batch route remains slower than the ordinary after serial route
for these small Parts. Few-large owned p50 changes at widths 2/4/8 are
141.13 → 131.29, 89.32 → 87.69, and 85.26 → 88.53 microseconds; warm-file
changes are 152.40 → 142.29, 95.04 → 95.17, and 92.96 → 93.46. The
after-only fixed-delay/instrumented many-small rows scale from the batch-one
control by 1.91x, 3.78x, and 7.27x at widths 2/4/8, but the scaling review
rejects simple Amdahl interpretation for superlinear rows and widths above the
four-Part work cap.

The aggregate comparison retains five adverse flags: three p95 flags and two
p99 flags. They occur on few-large owned batch-8 p95, few-large warm-file
batch-8 p95/p99, and many-small instrumented batch-8 p95/p99. Ten per-repeat
latency/throughput flags remain: one each for mean, p50, and throughput, three
for p95, and four for p99. The largest cluster is the second few-large owned
batch-8 repeat (+6.79% p50, +10.85% p95, +5.13% p99, +6.13% mean, and
−5.78% throughput); the other rows are sparse tail observations. They remain
retained descriptive flags on the shared host, not causal scheduler claims.
The delayed-provider batch-8 aggregate p99 increases 30.79%, despite a 3.67%
median reduction. This is an unresolved tail regression. The change is retained
for the substantial local multi-wave gains, with explicit caller policy and
the serial small-Part alternative still necessary.

No matched whole-child RSS regression exceeds five percent. Many-small
batch-8 RSS falls 12.08% for owned bytes and 10.27% for warm files. These are
process maxima including fixture setup and verification; allocation counts,
exact peak live bytes, and operation-local memory attribution remain unmeasured.

Separate clone tracing records 64 successful `clone3` calls before and four
after for one many-small batch at width four; the few-large width-four trace
is four before and four after. One whole-child profile comparison for
many-small owned width four records −21.17% cycles and −15.67% instructions,
alongside +251.49% context switches and +1,450% CPU migrations. Setup and
verification are included, so these counters show synchronization evidence
at whole-child scope and do not attribute operation-local CPU or RSS. Traced
elapsed times are excluded from the latency evidence.

All individual results and replay commands are retained in the
[evidence directory](../results/change-0499/README.md), including the
[matched comparison](../results/change-0499/comparison.md) and
[after scaling analysis](../results/change-0499/after-scaling.md).

This change remains a low-level OPC capability. It does not close format-level
CRUD selectors, native-producer or controlled-cold behavior, arbitrary
history/composition, or the program-wide tenfold result. The full non-iWork
`docs/GOAL.md` objective remains open.
