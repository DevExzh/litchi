# Change 0185: shared source-backed OPC overlay ownership

Date: 2026-08-18

## Decision

Retain an additive shared-ownership publication seam in
`litchi-opc::SourceBackedPackage`. The existing `Vec<u8>` single- and
multi-Part overlay methods remain compatibility adapters. New low-level
methods accept `Arc<Vec<u8>>` and move the same immutable allocation into the
bounded OPC regeneration plan.

Before this change, a changed source-backed format snapshot already owned its
selected XML as `Arc<Vec<u8>>`, but the format publisher detached it with
`to_vec()` and OPC immediately wrapped that new vector in another `Arc`. The
changed handoff therefore performed one complete selected-Part
`Arc<Vec<u8>> -> Vec<u8> -> Arc<Vec<u8>>` payload copy. The retained handoff is
`Arc<Vec<u8>> -> Arc<Vec<u8>>`; it is not an end-to-end zero-copy claim.

The shared seam is consumed by the existing same-topology source publishers
for DOCX document variables/main-document/hyperlink sanitization, PPTX
single- and multi-slide shape edits, XLSX scalar cells, and eleven XLSX
metadata/editor families. Topology-changing relationship-removal publishers
remain unchanged.

Exact semantic no-ops pass an empty overlay set and byte-copy the complete
source artifact. They do not detach managed `PartData`; managed bare-Arc escape
remains a typed refusal. Changed paths still sort and deduplicate Part names,
enforce the 64-Part and byte limits, materialize and compare every selected
source Part, validate source and candidate XML, refuse changed signed packages,
build the same raw-copy/regeneration plan, check source versions and
cancellation, and report partial sink progress exactly as before.

## Verification

The integrated production gate passed:

- 196 OPC, 871 DOCX, 520 PPTX, and 769 XLSX library tests;
- 27 focused DOCX source-backed variable/hyperlink/managed tests;
- 17 focused PPTX source-backed edit tests;
- 32 XLSX scalar-cell and 16 row-visibility integration tests;
- production Clippy for all four touched crates with warnings and deprecations
  denied;
- formatting, diff, and crate-boundary checks;
- independent current-tree review: SAFE.

The new OPC tests cover Vec/shared changed parity, signed exact no-op parity,
multi-Part reopen, duplicate/limit refusal before output, and bounded partial
sink failure. New managed XLSX tests prove scalar and multi-sheet exact no-ops
publish the byte-identical source without escaping retained payload
reservations.

Repository-wide all-target Clippy and public rustdoc attempts still expose
unrelated pre-existing test-style and private/broken intra-doc-link findings in
DOCX/XLSX. The touched production libraries themselves pass the strict
warning/deprecation gate.

## Measurement contract

The clean control revision is
`21ab71036682d4be23d5013d0c2b471b85e6975f`, release binary SHA-256
`04408b8fa0c0011db989a3f6bdf7f37512dd8a0744ca460d5e216a9b230bd410`.
The clean candidate revision is
`31855dfd267946bfa59eada1fe8f833fc561ce25`, release binary SHA-256
`69d42504e632c52ff86103e221a7e035c9fe2ccfbe8ccd7ce850dc5cb2c1320b`.

Fresh CPU-2-pinned processes run A1/B1/B2/A2 with 20 warmups and 500 retained
samples for twelve existing source-backed XLSX records:

- scalar one-cell, deterministic 1%, exact-256 batch, and multi-sheet edits on
  medium and dense-sparse corpora;
- row-visibility hide-one and unhide-256 edits on medium and large corpora.

The timer covers editor open, semantic staging, commit, and sequential
publication. The eight scalar-cell records use their existing pre-reserved
`CountingSink`, which retains the complete output; the four row-visibility
records use their existing zero-retention hashing sink with a 64 KiB window.
`publication_ns` is the mechanism-local primary metric; complete `elapsed_ns`
is secondary. The
predeclared p50/mean/p95/p99 same-implementation drift ceilings are
5%/5%/10%/15%. A statistic is accepted only when both paired directions are
lower and both same-implementation drifts pass.

Deleting the revision and all total/phase timing vectors produces the same
canonical projection SHA-256
`2a5edbf015fc0b78e4292b31d4e4b0b1ffc717950ec22ebbc638cee981e9f3e2`
for all four legs. Output and semantic hashes, untouched-member identities,
source/cache/sink counters, refusal gates, corpus identities, and update counts
therefore remain equal.

## Accepted latency scope

The complete lifecycle accepts all four statistics for:

- medium scalar deterministic 1%: p50 2.14%/1.98% lower, mean 2.77%/2.14%,
  p95 4.45%/2.68%, and p99 4.10%/3.95%;
- medium scalar exact-256 batch: p50 2.94%/1.15% lower, mean 3.06%/1.14%,
  p95 2.83%/3.60%, and p99 10.81%/7.29%;
- large row-visibility unhide-256: p50 0.21%/3.13% lower, mean 0.52%/3.46%,
  p95 1.71%/6.08%, and p99 2.37%/7.17%.

Medium scalar one-cell accepts complete p50/mean/p99, medium multi-sheet
accepts complete p50/mean/p95, and dense-sparse scalar one-cell accepts
complete p50/mean. Other complete statistics are withheld.

The primary publication phase accepts p50/mean/p95/p99 only for the large
row-visibility batch, at 0.34%/2.18%, 0.68%/2.40%, 1.84%/4.30%, and
2.35%/3.56% lower. Medium scalar one-cell accepts publication p50/p99;
medium 1% and exact-256 accept publication p99 only. All other publication
statistics are withheld because directions disagree or drift exceeds the
metric-specific ceiling.

This mixed result is expected for one ownership copy inside much larger
parse/validation/compression workflows. It supports retaining the simpler
shared handoff but does not justify an aggregate OOXML latency claim.

## Resource profile and withheld scope

One diagnostic whole-process Heaptrack run uses the large row-visibility
hide-one case with no warmups and one sample. Control/candidate allocation calls
are 64,907,125/64,907,109, temporary allocations 7,512,163/7,512,164, peak
heap 145.19/145.19 MiB, and profiler-inflated peak RSS 172.71/172.65 MiB.
Corpus construction and untimed correctness gates dominate these totals, so no
allocation, peak-heap, RSS, or profiled-runtime claim is accepted.

No physical-I/O, decompression, cold-cache, throughput, scaling, real-producer,
topology-changing publication, broad OOXML CRUD, or iWork claim is made.

Artifacts:

- [summary](../results/opc-shared-overlay-0185-summary.json)
- [manifest](../results/opc-shared-overlay-0185-manifest.json)
- compressed raw A1/B1/B2/A2 reports and diagnostic Heaptrack/profile files
  listed in the manifest
