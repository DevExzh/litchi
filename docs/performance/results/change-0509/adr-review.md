# 0509 ADR review

Reviewed 2026-09-11 against the ADR index and the 29 numbered records. The
mechanically generated [ADR hash manifest](adr-manifest.json) contains 30
entries (the 29 records plus `README.md`); recomputing SHA-256 from the current
file bytes produced zero mismatches. The manifest revision is
`0455d9abe479bef8980a403f0349ae2a95b254e3`.

The candidate in the working tree is limited to the private ODT sink parser in
[`elements/text.rs`](../../../../crates/litchi-odt/src/elements/text.rs). It
adds one operation-local spare `String` to `PendingSinkBlocks`. A completed
value is recycled only after its existing ordered `write_object` call returns
successfully, and a new block may take that spare. `recycle` drops values whose
actual `String::capacity()` exceeds 4,096 bytes, so the retained spare is
bounded to 4,096 bytes. An active block may still grow beyond that bound while
it is being written; that larger allocation is not retained. There is no
global, thread-local, or cross-operation pool.

The source review finds the following invariants preserved:

1. `MAX_TEXT_BLOCKS`, `MAX_TEXT_DEPTH`, `MAX_TEXT_BYTES`, space-count checks,
   fallible stack/frontier reservations, and `SequentialTextWriter` object and
   output limits remain in their existing paths.
2. Attribute validation, XML decoding/precharge, budget charging, slot checks,
   and typed document/sink errors keep their existing order. Recycling adds no
   fallible operation and occurs after the writer call; a writer error returns
   before recycling.
3. Slot reservation, `pending` frontiers, nested block depth, and contiguous
   start-order emission are unchanged. Each emitted value still crosses the
   same `TextObjectKind::Paragraph` writer boundary exactly once.
4. The parser still reads the source and feeds the same decoded text to the
   caller-owned sequential sink. No public API, dependency, manifest, budget,
   ownership, source-preservation, or archive behavior changes.

The [admission record](admission.json) admits only bounded paragraph-buffer
reuse: its control run attributes 200,000 `append_sink_precharged` allocation
calls across 20 exports of 10,000 paragraphs each and sets a 4,096-byte
retention cap. The [profile scope](profile-scope.md) says heaptrack totals are
whole-child measurements and callgrind is an instrumented parser profile; it
does not support a latency, throughput, hardware-counter, cache, or RSS claim.
The fresh hardware probe was denied by permissions. This review therefore
extends no performance claim beyond the admitted retention bound, and it does
not replace the required matched output, progress, and failure-path checks.
The current coder changes also include focused checks for successful reuse,
oversized-capacity drop, nested-frontier behavior, exact source-backed output,
and sink-failure progress; this audit did not execute them.

## Applicability matrix

| ADR | Applicability and disposition |
|---|---|
| 0001 | **Direct.** Correctness, safety, typed failures, and the strict public layers remain primary. The reuse is private, has no raw/public type, panic path, or unsafe code. |
| 0002 | **Direct boundary.** The edit stays in `litchi-odt`; it adds no peer-family, umbrella, or dependency edge and changes no ownership direction. |
| 0003 | **Applicable.** The parser remains read-only: no snapshot, edit, patch, or source mutation changes. The spare is owned by one parser operation and is dropped with it. |
| 0004 | **Direct API constraint.** Semantic text and the existing sequential-writer API are unchanged; no public signature or facade type is added. Plain text remains inert. |
| 0005 | **Direct.** Existing sequential non-seekable output, typed sink progress/errors, text/block/depth ceilings, fallible reservations, and output limits remain. The 4,096 value is a retained-spare capacity cap, not a universal text, allocation, peak-memory, or budget cap. |
| 0006 | **Direct.** XML validation/normalization, suppression rules, decoded bytes, ordering, and typed error boundaries remain. The optimization does not rewrite source markup or activate inert content. |
| 0007 | **Direct ODF model constraint.** Visible paragraph/heading semantics, nested start order, controls, tracked-change exclusion, note/ruby suppression, and paragraph writer kind remain unchanged. |
| 0008 | **Direct evidence gate.** The fresh admission and retained profiles support this bounded candidate only. They do not certify broader format support or a causal timing/RSS improvement; output, progress, malformed-input, and failure checks remain required. |
| 0009 | **No direct change.** ODF detection ownership and its fuzz boundary are untouched. |
| 0010 | **No direct change.** Facade and archive ownership are untouched. |
| 0011 | **No direct change.** OOXML physical package ownership is untouched. |
| 0012 | **No direct change.** BIFF8 formula reference types and encoding are untouched. |
| 0013 | **No direct change.** PPTX notes ownership and deletion are untouched. |
| 0014 | **No direct change.** The amended core-properties reader ownership record is unaffected. |
| 0015 | **No direct change.** Lossless OOXML core-properties CRUD is unaffected. |
| 0016 | **No direct change.** BIFF8 writer-location types are untouched. |
| 0017 | **No direct change.** OOXML producer-template ownership is untouched. |
| 0018 | **No direct change.** XLSX calculation-chain ownership is untouched. |
| 0019 | **No direct change.** DOCX web-settings ownership is untouched. |
| 0020 | **No direct change.** PPTX table-style ownership is untouched. |
| 0021 | **No direct change.** DOCX glossary/building-block ownership is untouched. |
| 0022 | **No direct change.** PPTX embedded-font ownership is untouched. |
| 0023 | **Direct topology constraint.** ODT remains the dedicated family owner; the change imports no concrete family or umbrella crate and moves no shared ODF capability. |
| 0024 | **Direct inventory constraint.** The current package topology and ODT semantic/XML ownership remain unchanged. |
| 0025 | **No direct change.** OGraph chart-area transactions are untouched. |
| 0026 | **No direct change.** Shared OLE directory metadata binding is untouched. |
| 0027 | **No direct change.** XLS sheet-anchor ownership is untouched. |
| 0028 | **Out of scope.** The ordered IWA monolith exit is unaffected; the task excludes iWork work. |
| 0029 | **Out of scope.** The archive-free IWA object-index foundation is unaffected; the task excludes iWork work. |

This is a documentation audit only. This review made no production file edits,
and ran no build, benchmark, test, or commit.
