# Validation history and scope

The unchanged source at `a31fc506e` passed the standalone allocator-feature
release library suite: 326 passed, one ignored. Strict standalone Clippy failed
with 29 rendered diagnostics in 17 distinct path/message groups. Those raw
receipts are retained; the after check must introduce no diagnostic group or
multiplicity, and no diagnostic in the new append module is admitted.

The first coverage-index invocation used a direct script path and failed its
`tools` import. The module invocation then exposed the missing generated-shape
contract for the new selector. Adding the explicit tiny/medium/large contract
made all 35 index tests pass. Both failed invocation receipts remain retained.

This is a harness baseline addition with no production Rust changes. Existing
ODP commit still materializes and validates the full candidate. The measured
sink accepts already-materialized committed bytes; its zero retained-output
counter describes the discard sink alone. The source and commit remain live
at the operation memory endpoint. No cancellation, source-backed, bounded
commit-memory, physical-I/O, native-rendering, or speedup claim follows.

The first focused Rust compile found that data-descriptor metadata belongs to
the archive entry, not its wayfinder. Correcting the accessor made all three
focused tests pass. The final oracle also matches the typed stale-source
`InvalidFormat` message and requires exact-noop shared byte identity. A fourth
test rejects missing append and changed opaque output. The first complete
all-features suite passed 368 tests with one ignored. Strict Clippy retained
exactly the same 29 diagnostics and no diagnostic in the new module; docs pass.

The first external CLI pilot then found a missing exclusion from the generic
OPC dispatch filter. It failed before producing a report. Direct module tests
had bypassed that routing layer. The root added `!case.uses_odp_existing_append()`
to the filter and retained the failed pilot, initial build/source receipts,
binary identities, and original pilot/save driver versions. The corrected
selector routing was independently reviewed source-only before recapture.

All measured workload execution is serialized by the root. Agents supplied
external source drafts and reviews. One initial evidence-driver handoff reported
Python compilation/help/static-protocol checks despite the source-only
instruction; it ran no Cargo, workload or profiler. Those checks are not used
as validation evidence. Root-owned checks and workload receipts are authoritative.

After the routing correction, the full suite again passed 368 tests with one
ignored, and strict lint debt was unchanged. Initial independent Python oracle
adapters had incorrect metadata and shape pins and expected a missing metric
label. Those adapters and failed pilots are retained in draft-history and
checks. All corrections preceded the formal matrix. The initial protocol and
build sidecar are retained alongside the final frozen protocol; no formal
measurement used the initial protocol. Six aligned pilots passed before capture.

The formal matrix completed all 12 reports and 360 samples on CPU 2 with one
worker, 30 retained samples and three warmups per fresh process. The independent
oracle's six controls and 14 corrupted-report probes passed. It independently
regenerates expected semantic content and opaque hashes; Rust gates inspect
the actual archive before reports are accepted.

The first stat profile workload and oracle passed, but its custody check
included artifacts created by the profile itself in full git status. That
failed receipt and all artifacts remain. The separately named refined driver
uses capture's status filter outside the bundle and enforces frozen oracle
hashes. Its amendment binds both driver hashes, unchanged protocol, capture
driver and source. Stat recapture and record then passed. Each record text
conversion retains 13 addr2line warnings. Profiles include untimed setup and
oracle work and cannot attribute costs solely to the operation timer.

The first summary adapter expected configuration.corpus_shapes instead of
semantic_shapes; its failed receipt and original driver are retained. The
corrected summary derives all quantiles, allocation vectors, RSS, repeat checks
and profile summaries from retained reports and was independently rechecked.

The first post-compression summary check exposed storage-dependent artifact
paths: derivation emitted `.gz` paths while the original summary named logical
receipt paths. The corrected adapter keeps receipt paths stable while checking
decompressed bytes against the original hash and size. The original adapter
and failed check remain; the measurement summary itself is unchanged.
