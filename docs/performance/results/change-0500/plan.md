# 0500 plan: managed DOCX paragraph batches

Base: `c797d04c7`. The previous turn was progress: bounded OPC worker reuse was
committed with verified evidence and cleanup. The non-iWork goal remains open.

The candidate addresses template filling through the existing format-owned
`replace_body_paragraph_texts` method. Managed single replacements already
reconstruct and validate from immutable source XML; the batch method currently
refuses managed snapshots. Earlier unmanaged coalescing (0012) established the
cost of repeated full-document reconstruction. Capture fresh managed repeated
controls before production edits, then compare the same repeated route and the
new admitted batch route using an unchanged harness and equal output oracles.
The batch route is an explicit alternative, not a historical same-API speedup.

Preserve atomic failure, exact no-op/base restoration, staged edits, canonical
selectors, semantic readback, source lineage/version proofs, typed limits,
retained-owner reservations, and monotonic actual work/input charges. A single
batch must reconstruct once rather than delegate to scalar replacements.
No new global scheduler, ambient capability, or package ownership crossing.

Measure managed full open/edit/commit/sequential-publication lifecycle and the
edit phase on deterministic 128/512-paragraph documents with 1/8/32 selected
paragraphs and an opaque payload. Compare owned and warm-file providers, with
three warmups and thirty measured samples in two repeats. Retain every >5%
latency/throughput/RSS adverse flag and supplementary whole-child perf counters.
Byte and semantic preservation, budget release, source counters, and explicit
scratch cleanup are independent measurement gates. Synthetic corpora do not
establish native producer or broad CRUD completion.

Build/check scratch belongs only to `/home/zhuhe/.cache/litchi-goal-0500`.
Preserve the sixteen hashes in protected-work.json and shared workspace target.
Commit the reviewed batch and remove owned scratch after verification, retaining
small frozen executables for replay.
