# 0492 DOCX bounded read-ahead experiment

This batch pursues source/I/O workstream A of `docs/GOAL.md`. The measured
0491 delayed provider performs 19 nonempty calls, fetching 3,966 bytes, with a
normal median near 20.3 ms. A fresh, finite source window may avoid repeated
one-millisecond transport service costs. This is an unmanaged benchmark pilot,
not a production API or completion of the performance program.

## Policy and evidence boundaries

The candidate retains one fallibly preallocated 4,096-byte window starting at
the requested offset. It fills at most once per miss and returns only a valid
requested prefix. Starting at the request, rather than an aligned earlier
offset, allows even one-byte short fills to make progress without falsely
returning EOF before reaching the requested byte. No retry loop, shared cursor,
network client, runtime, or global executor is introduced. Source identity and
revision remain delegated and checked; one mutex protects each adapter's cache.
The window and counters are fresh per sample. The transport remains the
existing `PptxRangeSource`; short-fill and version failures have focused tests.

The new opt-in v2 reports separate package logical requests, physical adapter
fills, bounded-window counters, and both range traces. Here “physical” means
calls below the read-ahead adapter into the synthetic byte-backed transport,
not disk or network observations. Media overlap is recomputed from the raw
compressed-member ranges without subtracting overfetch. Fetching compressed
media bytes does not prove media decompression. Cache snapshots before semantic
materialization and after text extraction distinguish open work from query
work. The source file/hash and actual returned text must satisfy the existing
0188-media-v1 oracles.

Both new control and candidate use the same v2 timer: package open, two cache
diagnostic snapshots, document materialization, extraction, and package/document
destruction. Text is retained until after the timer; digest/oracle comparison,
range snapshots, serialization and text destruction occur afterward. Source
construction and the candidate's fixed window allocation occur before the
operation allocation region. Report the 4 KiB setup cost separately; allocator
region peak increment and whole-child RSS are different measures. V1 default
reporting remains available for the sealed 0491 protocol; do not compare v1 and
v2 timings as an identical timer.

## Comparison

Use normal and allocator binaries from the same frozen source and toolchain,
three warmups and 30 samples for formal measurements, two repeats with arm and
role order reversed in the second repeat. First run three-sample/one-warmup
pilots. Compare control and 4 KiB candidate at zero transport delay and at
1,000 microseconds fixed service plus 104,857,600 bytes/second minimum-service
pacing, each with a 65,536-byte maximum transport request. CPU 2 affinity and
the established advisory CPU lock serialize this lane; they do not reserve the
CPU against other agents or host activity. Record every latency/tail/allocation/
RSS result and flag adverse changes above approximately 5%, not just medians.
Bootstrap intervals describe within-run sample variation, not independent-host
uncertainty. Captures, failed attempts, build hashes, source manifests and raw
logs are retained with exclusive receipts.

## ADR alignment and production gate

| Constraint | Pilot behavior / remaining production requirement |
| --- | --- |
| 0001 priorities; 0006 preservation/safety | Read-only adapter; bounded fills; actual text oracle; no serializer or edit changes. |
| 0002, 0010, 0011, 0024 ownership | Private benchmark module uses existing `litchi_core::ReadAt`; no facade/archive dependency added. |
| 0003 snapshots/source conflicts | Stable identity/revision delegated; stale cache and mid-fill mutation rejected. |
| 0005 bounded I/O/memory/context | Fixed window allocated explicitly in benchmark setup. Managed OPC currently charges logical request windows, so overfetch MUST NOT be deployed there until each physical fill is charged with cancellation and work checks. |
| 0008 evidence | Before/candidate controls, differential/adversarial/concurrent tests, source-bound captures and warning-denied checks. |

The adapter remains absent from filesystem cold/aligned-tail proofs and splice
publication, which rely on different exact read boundaries. No borrowed-source,
atomic-save, cold-cache, native-producer, cross-format, or multiworker performance
claim follows from this pilot. The broader CRUD taxonomy, production integration,
managed budgets, scaling, native evidence and final report remain active work.
