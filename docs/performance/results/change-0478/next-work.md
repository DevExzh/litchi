# Next bounded append batch after 0478

The next batch should first baseline and phase-profile the existing
source-backed DOCX plain-paragraph tail append path over 64, 8,192, and
131,072 paragraphs. Capture total operation allocator and I/O attribution
before changing production. If that evidence identifies full-document XML or
candidate retention as the dominant avoidable term, and ADR publication and
preservation review permits a narrower path, implement a bounded source-tail
publisher for that exact closure and recapture the same matrix.

This is the OOXML logical-append gap. A bounded source-tail publisher is a
conditional candidate design, not an already-approved replacement for the
general Snapshot/Edit/Commit/Patch path. Any implementation must remain a
specialized, fail-closed capability and must not silently turn that general
path into a lazy transaction.

This recommendation follows the four semantics in [`docs/GOAL.md`](../../../GOAL.md:298)
and the separation required by [`docs/CRUD_Scenario_Checklist.md`](../../../CRUD_Scenario_Checklist.md).
The completion criterion is explicit that streaming creation **and append** use
memory proportional to an explicit window (`docs/GOAL.md:809-810`).

## What the current evidence proves

| Semantic | Current non-iWork evidence | Missing proof or boundary |
| --- | --- | --- |
| Streaming creation from scratch | XLSX has a reusable row scratch buffer and fixed-memory unit coverage in `crates/litchi-xlsx/src/streaming.rs:176-205,2083-2166`; change 0432 measures the public writer. DOCX change 0473 measures `StreamingDocumentWriter` at 64/8,192/131,072 paragraphs with a 414,732-byte operation peak in all 180 allocator samples (`docs/performance/changes/0473-docx-streaming-operation-memory.md:15-38`). PPTX change 0478 now has the checked generated-name, central-directory spool route through `StreamingPresentationWriter` (`docs/performance/results/change-0478/README.md:9-47`). | The final 0478 matrix has 720 samples and a 432,436-byte operation peak in all 180 spool allocator samples at 8/256/8,192 slides. Provider storage grows separately. The validation, cleanup and portable receipts are retained in the bundle. This route is still fresh creation only. |
| Logical append to an existing structure | ODP has a public `SourceBackedTailAppendEdit` / `SourceBackedTailAppendPublicationPlan` (`crates/litchi-odp/src/lib.rs:29-32`) with bounded XML windows and ZIP replay; 0457 retains 64/4,096/8,192 source-slide captures and a flat 620,385-byte source-tail region peak (`docs/performance/changes/0457-odp-bounded-source-tail-publication.md:3-22`). RTF has `rtf_logical_tail_append` and a 16 KiB publication sink, but its validated candidate snapshot is retained (`docs/performance/CRUD_COVERAGE.md:1517-1531`). | The ordinary ODP append remains materialized, and 0457 expressly withholds Commit/Patch equivalence (`docs/performance/results/change-0457/next-work.md:3-27`). No public DOCX or PPTX bounded append writer exists. DOCX's current plain-paragraph append position is only a materialized copy: `paragraph_copy.rs:318-337` permits `before == paragraph_count`, while `paragraph_copy.rs:1348-1368` reserves the complete output and builds a complete candidate snapshot. |
| Adding a new package Part | `opc_part_add_lifecycle` is the existing low-level source-backed OPC baseline (`docs/performance/CRUD_COVERAGE.md:191-198`; selector in `tools/perf-baseline/src/lib.rs:1588-1592,11421-11424`). | It has a synthetic package topology and no semantic DOCX/PPTX/XLSX owner. The 0477 record still lists Part addition as open for public integration (`docs/performance/CRUD_COVERAGE.md:3-10`). It cannot count as logical append. |
| Arbitrary modification followed by repackaging | The semantic DOCX/PPTX/XLSX/ODP edit/save selectors exercise ordinary owned transactions and source-backed overlays in their supported closures. | They retain or regenerate the changed semantic package and do not establish an explicit-window append bound. 0478 itself excludes arbitrary editing, repackaging, append and Part addition (`docs/performance/changes/0478-pptx-generated-metadata-spool.md:73-90`). |

The largest measured existing-append bottleneck remains the ordinary ODP
control: 0457 reports 35,958,388 bytes of operation-region peak above entry at
8,192 slides, versus 620,385 bytes for its specialized source-tail path
(`docs/performance/changes/0457-odp-bounded-source-tail-publication.md:10-16`).
That is a useful control and a separate phase-attribution target. The missing
OOXML append capability is the better independent batch here because the ODP
source-tail path already has bounded evidence while its ordinary Commit/Patch
contract remains distinct.

## First batch: baseline and profile before production

The current existing-document path is the public source-backed
`plain_paragraph_copy` append-at-the-end operation. It accepts
`before == paragraph_count` (`crates/litchi-docx/src/source_backed/paragraph_copy.rs:318-337`),
but the current transaction reserves complete output and builds a complete
candidate snapshot (`crates/litchi-docx/src/source_backed/paragraph_copy.rs:1348-1368`).
Measure this path before changing its production behavior:

1. Add an opt-in `docx_plain_paragraph_tail_append_lifecycle` selector to
   `tools/perf-baseline`. Exercise the current public source-backed
   plain-paragraph append-at-end API against existing DOCX-shaped source
   archives at 64, 8,192, and 131,072 paragraphs, appending fixed 1, 64, and
   256 paragraph requests. Keep the selector out of the default matrix.
2. Bound the timed phases around the operation the public path actually runs:
   source open and snapshot/materialization, edit staging, candidate commit,
   and sequential publication to a 16 KiB non-seek hashing sink. Build any
   semantic or physical oracle outside the timing region. The operation total
   must include every phase, including source and candidate retention that the
   current path owns.
3. Attribute allocator work by phase and for the whole operation: requested,
   allocated, deallocated, reallocated, peak live bytes, live bytes at region
   entry and exit, retained source/candidate/output bytes, and operation-region
   peak above entry. Report the materialized candidate size explicitly rather
   than hiding it behind a single peak number.
4. Attribute logical I/O separately: source `ReadAt` calls and bytes, output
   write calls and accepted sizes, and copied, decompressed, or recompressed
   member bytes where the current instrumentation exposes them. These counters
   must be labeled as logical I/O; they do not establish physical disk I/O.
5. Keep correctness and refusal gates untimed but mandatory: source/output/
   inserted XML bytes and hashes, physical member order, untouched-member byte
   digests, paragraph order and text, final `sectPr` placement, complete
   reopen/readback, stale-source refusal, limits, short/zero/interrupted sink,
   cancellation, malformed or complex-source refusal, and output-limit
   behavior.
6. Run normal and allocator fresh-process captures using the existing
   performance harness conventions, with three warmups and at least 30 samples
   per size/append pair. Record the exact selector, source/build manifests,
   corpus hashes, Rust/toolchain/environment, and a portable post-cleanup seal.

Do not implement a bounded source-tail path until this profile shows that
full-document XML or candidate retention is the dominant avoidable term and an
ADR review confirms the required publication, preservation, stale-source,
failure-atomicity, and public-contract requirements. If that gate does not
clear, retain the measured baseline and document the limiting contract.

## Conditional follow-up: bounded source-tail candidate

If the baseline clears that gate, a candidate implementation would add a
source-backed DOCX tail-append module beside
`crates/litchi-docx/src/source_backed/paragraph_copy.rs`. Its public API and
publication contract require the ADR review; the following is the narrow design
to evaluate, not an authorized architecture:

1. Open through `source_backed::Package::from_read_at` (the constructor defers
   main-document payload materialization;
   `crates/litchi-docx/src/source_backed.rs:309-314`). Accept only the
   canonical main `word/document.xml` story containing direct plain
   `w:p`/`w:r`/`w:t` paragraphs and an optional final body `w:sectPr`.
2. Validate the source with a bounded iterative XML scan, retaining only source
   position/version/fingerprint, a fixed parser window, and one caller-bounded
   encoded paragraph fragment or reusable fragment buffer. Replay the source
   ZIP and insert immediately before the final `w:sectPr` or `</w:body>`.
   Avoid `document_snapshot`, `plain_paragraph_copy_snapshot`, and any path
   that reserves `source.xml.len() + fragment.len()`.
3. Preserve every untouched member and source byte span through the existing
   OPC preservation writer. Refuse MCE branch selection, tables, fields,
   hyperlinks or external relationships, tracked changes, macros, signatures,
   protection, unknown body children, and noncanonical section layouts before
   output. Retain stale-source checks, cancellation, output/input/fragment/
   event limits, and accepted sink progress as typed, failure-atomic gates.
4. Keep this candidate separate from the general reversible patch contract. If
   an inverse or durable patch is added later, it needs its own bounded retained
   artifact design. The initial candidate would need a source-checked
   publication plan with a zero-retained-output sink contract only if the ADR
   approves that contract.

The existing DOCX tests provide the correctness seam to extend: `crates/litchi-docx/tests/streaming.rs:68-240`
checks deterministic three-member creation, short/interrupted sinks, invalid
text/progress/poisoning and cancellation/scratch release;
`crates/litchi-docx/tests/source_backed_paragraph_copy.rs:107-222,261-336,604-720`
checks append-at-the-end semantics, exact untouched-member preservation,
durable/inverse behavior, limits, stale source and partial/zero sinks. The new
tail tests must prove those properties without materializing the complete main
XML.

## Profile acceptance and conditional follow-up evidence

The baseline report is the production gate. It must retain, per normal and
allocator sample, the following evidence for the current materialized path:

* source/output/inserted XML bytes and deterministic source/output hashes;
* source `ReadAt` calls and bytes, output writes and accepted sizes, and any
  available copied/decompressed/recompressed member-byte counters;
* requested, allocated, deallocated and reallocated bytes, peak live bytes,
  live bytes at region entry and exit, retained source/candidate/output bytes,
  and allocator region peak above entry, with vectors for open, snapshot,
  staging, commit, and publication;
* physical member order, untouched-member byte digests, paragraph order/text,
  final `sectPr` placement, complete reopen/readback, and all refusal/sink/
  cancellation/limit gates;
* the exact current-path phase boundaries, source/build manifests, corpus
  hashes, Rust/toolchain/environment, selector implementation, and portable
  post-cleanup seal.

Interpret the result as total operation memory and logical-I/O attribution.
Keep source ownership already live at region entry, cumulative requested
allocations, provider storage, process RSS, and sink storage in separate
fields. Do not call the baseline a constant-RSS or constant-CPU claim.

If a candidate bounded path is approved and implemented, recapture the same
sizes, append requests, fresh-process modes, correctness gates, allocator
fields, and logical-I/O fields. A candidate result is admissible only when its
source-owned data outside the explicit parser/fragment/publication window is
not materialized, the operation-region peak stays flat while output bytes
grow, and the ADR-approved publication and preservation contract remains
failure-atomic. A matched materialized control is useful only when its timing
and retention contract is explicitly comparable; it is not a substitute for
the bounded-path proof.

0478's final PPTX evidence is scoped to fresh creation. Do not
combine its control/spool figures with this append evidence, and do not
use the low-level `opc_part_add_lifecycle` or RTF sink window as substitutes for
an OOXML logical-append memory proof. iWork remains outside this scope.
