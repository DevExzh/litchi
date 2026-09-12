# DOCX managed paragraph publication map

This is a read-only source map at revision `201f65094ff0ca6186fbf1e2ffa6b3641290253b`.
The current priority is OOXML/OLE2; ODF remains deferred and iWork is outside
this slice. I reviewed the accepted ADR index and the relevant accepted ADRs
0001, 0003, 0005, 0006, 0010, 0011, 0017, 0019, 0021, and 0024 before making
this map. No production source was changed, and this map makes no performance
claim by itself.

## Candidate construction

The current managed paragraph example is
[`managed_paragraph_batch_perf.rs`](../../../../crates/litchi-docx/examples/managed_paragraph_batch_perf.rs).
`build_fixture` (454-506) creates the direct-body paragraphs and the untouched
media/opaque members. `build_managed_preflight` (539-580) opens a managed
source-backed package, stages the requested route, commits it, checks forward
and inverse patch application, publishes an expected artifact, and checks
release/exact-output invariants. `stage_edit` (582-598) selects either the
repeated scalar API or `Edit::replace_body_paragraph_texts`.

The production path is:

1. `source_backed::Package::edit_document` ([`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs):940-943) calls
   `main_document_snapshot` (1736-1834). For a managed source it captures
   `PartView::source_xml` -> `SourceBackedPackage::source_xml_part`
   ([`opc/source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs):5881-5911),
   which reads the main member and constructs `SourceXmlPart`; its constructor
   validates the complete source XML (`xml_splice.rs`:83-127). DOCX then runs
   `ensure_source_document_xml` (`source_backed.rs`:2132-2350) and
   `Snapshot::from_source_xml` (`document/transaction.rs`:973-1010), which
   scans/indexes the source document.
2. `Edit::replace_body_paragraph_texts` (`document/transaction.rs`:2459-2465)
   dispatches managed input to `replace_managed_body_paragraph_texts`
   (2552-2931). The batch planner validates/admit strings, scans each selected
   projected paragraph and immutable-base paragraph (2634-2691), checks the
   prior operation ledger (2693-2725), and builds the final operation list.
   Repeated mode enters `replace_managed_paragraph_text` (1870-2168) once per
   replacement and performs the analogous checks each time.
3. Both routes share `reconstruct_managed_paragraph_operations` (2170-2446).
   It rescans every final base paragraph (2223-2270), creates exact
   `SourceXmlPart::checked_range` proofs (`xml_splice.rs`:148-212), audits each
   generated fragment, and calls `XmlSplicePublication::replace`. One
   `XmlSplicePublication::finish` assembles the candidate and validates the
   complete assembled XML (`xml_splice.rs`:510-585). The candidate is then
   reparsed/indexed once by `Snapshot::from_source_xml_with_identity`
   (`document/transaction.rs`:2411-2413), and the selected paragraphs are
   semantically read back (2426-2437) before the projected edit is assigned.
4. `Edit::commit` (`document/transaction.rs`:4365-4405) retains a changed
   source-backed candidate instead of compacting it, rechecks the managed
   context, and builds inverse operation metadata. It does not rebuild the
   managed candidate.

The batch planner's projected scan plus immutable-base scan is required for
selector and stale-source semantics. The shared reconstruction then scans each
final operation's base paragraph again. That is a candidate-construction
repeat-work opportunity to profile separately; it is not publication CPU.
The prior-ledger walk is also bounded but can be O(selected replacements x
prior operations) (`transaction.rs`:2693-2719).

## Public publication and save

The changed managed call is
`source_backed::Package::publish_document_commit_to_stream`
([`docx/source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs):1293-1345):

* It calls `main_document_snapshot("publish_document_commit_to_stream", true)`
  for a changed patch (1298-1301). This is a second source-authorized current
  snapshot: `source_xml_part` reads/validates the original main XML, then
  `ensure_source_document_xml` and `Snapshot::from_source_xml` scan/index it.
  The source cache may make the read warm, but the snapshot and validation work
  remains in the public method.
* `commit.patch().apply(&current)` (1302) checks exact source identity/bytes and
  clones the already-built target; `Patch::apply` is at
  `document/transaction.rs`:4649-4662. Unsupported cross-package operations
  are rejected (1303-1315). For a changed patch,
  `document_policy::validate_changed_document` (1322-1324) checks settings
  relationship/dialect/content type and protection/tracked-revision policy;
  its implementation is `source_backed/document_policy.rs`:40-111.
* A source-backed target is placed in a `SourceTopologyPlan` by
  `SourceTopologyPlan::try_replace_source_xml_part`
  ([`opc/source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs):1511-1547).
  This checks Part identity/XML classification and retains the source proof.
  The method then calls `SourceBackedPackage::write_topology_to_stream`
  (6492-7815). The ordinary authored overlay branch
  `write_part_overlay_shared_to_stream` (7925-7938) is not the managed
  source-proof route.
* `write_topology_to_stream` resolves the replacement against the immutable
  catalog and calls `SourceXmlPart::check_for_replacement`
  (`opc/source_backed.rs`:6605-6613; `xml_splice.rs`:271-298). This checks
  lineage/version/Part identity, exact destination original bytes, and then
  validates the replacement payload. The later source-token loop now calls
  `SourceXmlPart::check_source_state` (`opc/source_backed.rs`:7721-7729),
  retaining the source freshness/context/cancellation fence without scanning
  the immutable candidate XML a second time. Source XML additions receive the
  same treatment: their complete `check_for_publication` proof remains in the
  initial addition-resolution loop (6536-6538), and the append loop performs
  only `check_source_state` (7749-7756).
* The topology path re-reads the original selected member before it builds the
  changed overlay (`opc/source_backed.rs`:7583-7610), rejects byte/source or
  canonical-member mismatches, and refuses encrypted/signature/unsupported
  physical layouts. It then enters
  `write_changed_overlays_with_appended_inner` (9466-9795). This builds the
  ZIP `preservation_index_with_limits` (9477-9507), scans the archive entries
  for replacement counts and canonical names (9513-9629), creates one
  `PreservationPlan` action per source entry (9631-9679), and invokes
  `index.write_to`/`write_to_with_accounting` through source, execution,
  output-budget, and sink guards (9682-9771). The actual copy/regenerate,
  recompression, central-directory serialization, flush, and final
  `finish_source_publication` checks are therefore part of publication CPU.

For one changed main-document replacement, the relationship/addition/removal
vectors are empty, but catalog/topology validation, preservation indexing,
archive-entry scans, plan construction, and complete ZIP output still execute.
The current implementation's broad preservation work is a candidate profile
target only after a function-level publication profile proves its cost.

## 0517 proof matrix and work accounting

The scoped change in `crates/litchi-opc/src/source_backed.rs` replaces only the
second source XML `check_for_publication` calls in the replacement-token and
source-addition transfer loops with the existing `check_source_state` method.
The earlier checks stay in place: replacement resolution still verifies
lineage, source version, Part identity, destination original bytes, content
type, destination limits, and complete XML validity; additions still verify
content type, destination limits, and complete XML validity before topology
planning. The immutable payload, destination limits, and source proof are
therefore the same when transfer begins, while the late source/context fence
still runs before source monitoring and sink transfer.

| Accepted ADRs | Proof applied to this change |
| --- | --- |
| [0001](../../../adr/0001-priorities-and-api-layers.md), [0002](../../../adr/0002-crate-topology.md), [0010](../../../adr/0010-facade-archive-ownership.md), [0011](../../../adr/0011-ooxml-physical-package-ownership.md), [0024](../../../adr/0024-current-topology.md) | The shared OPC publisher remains the physical-package owner; no new public API or format-facade bypass was introduced. |
| [0003](../../../adr/0003-snapshots-edits-and-patches.md), [0004](../../../adr/0004-semantic-api-design.md) | Source proof and fallible checks remain ordered before transfer; failures still return without accepted output. |
| [0005](../../../adr/0005-io-memory-and-performance.md), [0006](../../../adr/0006-validation-security-and-compatibility.md) | The optimization removes only duplicate immutable XML work; source version, cancellation, destination limits, XML validation, and exact-preservation checks remain. |
| [0007](../../../adr/0007-office-object-models.md), [0008](../../../adr/0008-migration-and-verification.md), [0015](../../../adr/0015-lossless-core-properties-crud.md), [0017](../../../adr/0017-ooxml-producer-template-ownership.md), [0019](../../../adr/0019-docx-web-settings-ownership.md), [0021](../../../adr/0021-docx-glossary-ownership.md) | DOCX semantic ownership, lossless properties, and producer-template paths are unchanged; the publication proof remains at the OPC boundary. |
| [0012](../../../adr/0012-biff8-formula-reference-types.md), [0016](../../../adr/0016-biff8-writer-location-types.md), [0026](../../../adr/0026-ole-directory-metadata-binding.md), [0027](../../../adr/0027-xls-sheet-anchor-ownership.md) | OLE2/BIFF paths are untouched; the shared OOXML optimization does not alter their checked behavior. |
| [0013](../../../adr/0013-pptx-notes-deletion.md), [0020](../../../adr/0020-pptx-table-style-ownership.md), [0022](../../../adr/0022-pptx-embedded-font-ownership.md), [0025](../../../adr/0025-ograph-chart-area-transactions.md) | PPTX semantic owners remain unchanged; source XML transfer uses the same shared OPC proof and late source fence. |
| [0018](../../../adr/0018-xlsx-calculation-chain-ownership.md) | XLSX calculation-chain ownership is unchanged. |
| [0009](../../../adr/0009-odf-detection-ownership.md), [0023](../../../adr/0023-odf-family-crate-split.md) | ODF remains deferred until the OLE2/OOXML optimization goal is complete. |
| [0028](../../../adr/0028-iwa-monolith-exit.md), [0029](../../../adr/0029-iwa-index-foundation.md) | iWork/IWA remains outside this task. |

`validate_source_xml` consumes `Resource::Work` equal to the candidate payload
length. Each source XML replacement/addition now performs that full scan once
per `write_topology_to_stream` call, followed by a zero-work
`check_source_state` fence. Relative to the previous path, the cumulative work
charge decreases by exactly one `payload.len()` for each avoided second scan;
the late fence creates no XML-validation memory or object reservation. The
existing managed source payload and metadata reservations remain owned by the
source token/cache and are released by their normal drops.

Focused coverage now checks destination XML depth/event limits, source-version
refusal, cancellation after the initial source proof, one validation-pass work
accounting, empty output on refusal, and managed memory/object release.

Focused validation completed before handoff:

* `cargo fmt --all -- --check`
* `CARGO_TARGET_DIR=/home/zhuhe/litchi-goal-0517-target CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 TMPDIR=/tmp/litchi-goal-0517 cargo test -p litchi-opc --all-features topology_source_xml_addition -- --nocapture` — 3 passed.
* `CARGO_TARGET_DIR=/home/zhuhe/litchi-goal-0517-target CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 TMPDIR=/tmp/litchi-goal-0517 cargo test -p litchi-opc --all-features --test source_xml_publication -- --nocapture` — 8 passed.

The exact test-result lines were `test result: ok. 3 passed; 0 failed; 0
ignored; 0 measured; 408 filtered out` and `test result: ok. 8 passed; 0
failed; 0 ignored; 0 measured; 0 filtered out`.

Source hashes at handoff are `source_backed.rs`
`444ca5d86b0b8f3498d8460c4f77255479c40c3ac730160dbb88ecc0cc3ad543` and
`source_xml_publication.rs`
`d750b56695b6f623d5ca443e50b0fa34baace64ad22d276c11f4d2f0e72872ce`.

## Readback and old evidence

In `run_sample` (635-750), fixture/source/budget/output setup is before the
timed phases. The `publish_ns` timer (673-676) covers the public method and
`drop(published)`; it does not cover the later `drop(commit)` phase. Thus it
does not equal the method body alone: an explicit method-scope profile should
stop when the returned `Snapshot` exists and drop that returned snapshot after
the profile. The consumed `Package` is still dropped as the by-value method
returns.

After all timers, the example computes hashes, reopens the emitted DOCX, and
checks every paragraph in `verify_output` (752-767); it reopens again through
`OpcPackage` for untouched media/opaque payloads (769-776), and compares the
full output to the preflight artifact (688-714). These are required correctness
gates, but they must stay outside a publication-CPU interval. The preflight
forward/inverse and expected-output work is likewise outside samples
(`managed_paragraph_batch_perf.rs`:509-580).

The 0500 tables report descriptive phase timings and whole-child perf counters.
For example, p512 owned batch/repeated publication p50 is approximately
constant across K while edit/lifecycle scales in the historical repeated route
(`docs/performance/results/change-0500/comparison.md`:15-20, 34-39). The 0500
profile scope explicitly includes fixture setup and independent verification
and states that it supplies no operation-local CPU attribution
(`change-0500/README.md` and `profile-summary.json`). Those observations do
not identify `write_topology_to_stream`, `index.write_to`, or any other owner.

## Recommended narrow profile boundary

Use a future dedicated probe with setup and correctness work completed before
the measured region:

1. Build or load the deterministic fixture, create a fresh managed context and
   source-backed package, call `edit_document`, stage one selected batch
   replacement set, and call `commit` before profiling. Keep `Package`, `Commit`,
   and a preallocated output sink alive.
2. Run the required patch/source/semantic preflight once before profiling. Keep
   the source version unchanged and record cache/budget state. Do not call
   `verify_output`, reopen the output, hash it, or drop the commit in the
   measured region. A preallocated `Vec` with the same capacity as the known
   expected output preserves the production write path while keeping capacity
   allocation outside the interval; a counting sink can be a separately named
   serialization-only experiment, not the primary full-publication result.
3. Start counters immediately before the single call to
   `Package::publish_document_commit_to_stream`, and stop after the call returns
   its `Snapshot`. Keep the call itself intact, including the current source
   snapshot/reparse, patch application, policy checks, source proof checks,
   topology validation, preservation planning, ZIP writes, source guards, and
   output/sink checks. Drop the returned snapshot, commit, source, and output
   after counters stop; the existing `publish_ns` placement of
   `drop(published)` must not be used as a method-only claim.
4. Use an `#[inline(never)]` probe wrapper or an equivalent profiler region so
   setup/readback cannot be mistaken for publication. The current release
   baseline retains the symbols
   `litchi_docx::source_backed::Package::publish_document_commit_to_stream`,
   `managed_paragraph_batch_perf::run_sample`, and Snapshot drop symbols in
   `change-0517/baseline/symbols.txt`; this is enough for function-level
   symbol sanity. A formal attribution should use a dedicated region/function
   and suitable line information rather than relying on whole-child counters.
   `samples=1`, `warmups=0`, `repeats=1` is appropriate for checking that the
   symbol/region is retained and the boundary is wired. It is not enough for a
   stable performance conclusion.
5. Retain one unprofiled full-output semantic/untouched-member/exact-output
   gate for the same fixture and route, and report the publication profile as
   warm or cold according to the explicitly prepared cache state. Do not infer
   causality from the old phase timer or from process-wide RSS/counters.

This boundary measures the public changed-document publication contract and
keeps candidate construction, its full candidate reparse/readback, setup, and
post-save verification distinguishable. Any later optimization that fuses or
removes a source proof/reparse must preserve the accepted ADR validation,
source-authority, cancellation, budget, and exact-preservation contracts and
must be re-profiled at this same boundary.
