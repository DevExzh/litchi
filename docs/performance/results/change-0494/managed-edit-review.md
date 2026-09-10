# Managed DOCX ordinary paragraph edit review

Status: read-only design review for change 0494. No production source or test
files were changed while preparing this review, and no Cargo command was run.

## Decision

Keep the current typed refusal from the generic managed `document_snapshot()` /
`edit_document()` surface only as the present compatibility boundary while the
managed ownership path is implemented. The required end state is the ordinary
managed opened-DOCX lifecycle promised by the goal: an isolated paragraph edit,
atomic commit, source-checked reversible patch, forward publication, and
inverse publication through the ordinary facade. A source-backed specialized
transaction may be used as a temporary measured implementation seam while that
ownership path is being built. It cannot become a permanent substitute for the
ordinary facade. The end-state transaction storage must retain an OPC-issued
owner through capture, staging, validation, and publication; it must never
convert `PartData` to an unreserved `Arc<Vec<u8>>`.

The least invasive implementation seam is a deliberately narrow,
source-backed method such as `Package::edit_source_backed_paragraph(position,
text)` or a `Package::source_backed_document_edit()` with only the ordinary
paragraph operation exposed. Use that surface to measure bounded capture,
splicing, reservation lifetime, and reversible publication before wiring the
same owner-carrying storage into the ordinary `Snapshot`/`Edit`/`Patch` path. A
one-way specialized save is only an intermediate investigation result; even a
reversible specialized save does not complete the ordinary-facade requirement.

This is the smallest bounded implementation seam for the requested
provider/file managed one-edit/save path while preserving typed refusal for
operations whose ownership or dependency closure is not proved. It is not a
waiver of the generic snapshot/edit/patch contract. Accepted ADR 0003 and the
goal already authorize hidden owner-carrying snapshot state, persistent/COW
storage, source-owner plus byte-span views, bounded inverse material, and
fallible reservations. The specialized seam can reduce the first measured
change, but the generic ordinary path remains the required integration target.

## Evidence and exact seams

The current source-backed DOCX facade intentionally refuses before reading the
main payload when the cache is budget-managed. `main_document_snapshot()` checks
`cache_diagnostics().budget_managed` and returns `Error::UnsafeEdit` with the
reason that a managed transaction would require an owned snapshot; only after
that check does it call `main.data()` and `PartData::into_arc()`.
See [`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs#L1577-L1610).
The public `document_snapshot()` and `edit_document()` methods are thin wrappers
over that refusal. See
[`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs#L859-L881).
The regression test requires zero payload reads and zero memory use for this
refusal, so changing this branch silently would weaken an existing ownership
contract. See
[`source_backed_managed.rs`](../../../../crates/litchi-docx/tests/source_backed_managed.rs#L340-L367).

`PartData` borrows its payload with `as_bytes()`, but `into_arc()` returns
`ManagedPartDataArcEscape` when the payload carries a reservation. Its clone
keeps the reservation alive. See
[`source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs#L3955-L3997).
The managed read facade already follows this rule: it retains
`DocumentPayload::Managed(PartData)` and runs the bounded source XML check over
borrowed bytes. See
[`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs#L585-L621).

The generic transaction is structurally an `Arc` API. `Snapshot` stores
`Arc<Vec<u8>>`; `paragraph()` and `paragraphs()` clone that Arc into lifetime
free `Paragraph` values; `edit()` clones both base and projected snapshots. See
[`transaction.rs`](../../../../crates/litchi-docx/src/document/transaction.rs#L396-L515).
`replace_paragraph_text()` allocates a complete candidate XML `Vec`, re-parses
it into another `Snapshot`, and stores the candidate in the projected state. See
[`transaction.rs`](../../../../crates/litchi-docx/src/document/transaction.rs#L1006-L1080).
The resulting `Patch` retains before and after `Snapshot`s and can be cloned or
inverted. See
[`transaction.rs`](../../../../crates/litchi-docx/src/document/transaction.rs#L2715-L2778).

Consequently, these seemingly small changes are unsafe for managed input:

* Calling `into_arc()` is explicitly rejected.
* Copying `data.as_bytes()` into a new `Arc<Vec<u8>>` would create an allocation
  whose reservation is no longer attached to the owner returned by the API.
* Adding one reservation field to `Snapshot` is insufficient. Every
  lifetime-free semantic view, every candidate snapshot, every cloned patch,
  durable lineage, and every publication target would have to carry the owner
  token. A missed path would recreate the detached allocation bug.
* Returning a generic `Snapshot` would also expose `shared_xml()` to the
  source-backed publisher, which currently publishes an Arc-backed target rather
  than an OPC source-authorized payload.

The generic paragraph algorithm is still the semantic reference. It scans a
direct paragraph's supported text owner, preserves run boundaries and unknown
run XML, validates authored text, rewrites each text slot, and performs semantic
readback. The replacement helper is private today, so the managed path should
extract a bounded, shared internal helper or duplicate only a carefully reviewed
range-producing portion. See
[`transaction.rs`](../../../../crates/litchi-docx/src/document/transaction.rs#L3411-L3467)
and the fragment sizing/escaping code at
[`transaction.rs`](../../../../crates/litchi-docx/src/document/transaction.rs#L4778-L4887).

The existing OPC source-preserving seam already has the ownership needed for
this design. `PartView::source_xml()` captures a validated XML Part while
retaining `PartData`; it refuses encrypted and signed source infrastructure and
disables read-ahead before publication. See
[`source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs#L3471-L3483)
and the capture path at
[`source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs#L5819-L5856).
`SourceXmlPart` keeps the original `PartData`, its payload, lineage, source
version, and reservations. It issues exact source-range proofs and creates a
one-shot `XmlSplicePublication`. See
[`xml_splice.rs`](../../../../crates/litchi-opc/src/xml_splice.rs#L43-L220).
`replace()` rechecks provenance and stale bytes, while `finish()` reserves the
entire output before copying untouched bytes and validates the assembled XML.
See [`xml_splice.rs`](../../../../crates/litchi-opc/src/xml_splice.rs#L395-L545).
`SourceTopologyPlan::try_replace_source_xml_part()` retains the
`SourceXmlPart` token alongside its payload, so publication cannot outlive the
reservation or source authority. See
[`source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs#L1509-L1545).
The main document Part is an accepted target for this route: the focused OPC
test captures `/word/document.xml`, finishes a checked source splice, stages it
with `try_replace_source_xml_part()`, and publishes it through
`write_topology_to_stream()`. See
[`source_xml_publication.rs`](../../../../crates/litchi-opc/tests/source_xml_publication.rs#L109-L166).
The publisher then rechecks that the destination has the same source lineage,
version, Part identity, content type, and original Part bytes. This is a valid
forward path for `word/document.xml`; it is not by itself an inverse path,
because a derived `SourceXmlPart` cannot be re-edited and its replacement
authority remains tied to the original source bytes.

## Implementable design

Capture one transaction as follows.

1. Check execution and source freshness, resolve the main document Part and
   validate its content type.
2. Obtain `main.source_xml()` (or the equivalent `PartView` call). This keeps
   the managed `PartData` reservation attached to the edit and gives the
   transaction an exact source version, lineage, Part URI, and bounded XML
   payload.
3. Run the existing `ensure_source_document_xml()` check over
   `SourceXmlPart::bytes()`. This must remain a borrowed scan; MCE branch
   selection, processing instructions, DTD/entity constructs, malformed XML,
   and scan-limit failures are typed refusals before staging. The current
   bounded scan is at
   [`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs#L1887-L1901).
4. Scan only direct main-body paragraphs and retain bounded ranges and text-slot
   metadata. Do not expose `Paragraph` values backed by a bare Arc. The public
   managed edit object may expose position, count, and copied text strings, but
   its source owner remains the `SourceXmlPart`.
5. Validate the ordinary paragraph closure before allocating an edit: direct
   `w:p` only; no hyperlink, field, tracked-change, content-control, drawing,
   relationship, external target, or transfer operation unless the existing
   transaction rules prove that exact closure. Refuse unsupported structures
   before output.

Stage one paragraph replacement by source ranges, not by replacing the whole
paragraph with a newly owned buffer:

* Reuse the transaction's text-owner rules and authored-text validation.
* For each supported text slot, compute the same compact replacement fragment
  produced by `try_run_content_fragment()` and issue
  `SourceXmlPart::checked_range(slot_range, original_slot_bytes)`.
* Feed each proof and generated fragment to one
  `XmlSplicePublication::replace()`. This copies every byte outside text slots
  directly from the original source, including paragraph properties, run
  properties, namespace spelling, comments, and unknown run markup.
* `AuthoredXmlFragment::markup()` is appropriate for the generated text-element
  fragment only after its compactness audit succeeds. Do not pass an arbitrary
  source paragraph or a full `rewrite_text_owner()` result as an authored
  fragment: the source-preserving API intentionally distinguishes exact source
  bytes from compact generated markup. If a generated slot fragment needs an
  authoring rule not covered by the current audit, add a small OPC fragment
  constructor with the same bounded audit rather than bypassing it.
* If the requested text equals the observed text, record an exact semantic
  no-op and do not call `finish()` with an edit. The original source remains the
  publication authority.

After staging, call `finish()`. The existing splice implementation checks range
ordering/overlap, output size, work, cancellation, source freshness, and XML
well-formedness, then stores the target bytes with an output reservation in the
returned `SourceXmlPart`. Before putting that token into a topology plan, run
the DOCX source scan and semantic readback against the target bytes through
borrowed slices. Do not call `into_arc()` or manufacture a generic
`Snapshot`. The readback must prove the selected paragraph has exactly the
requested text and that direct paragraph indexing remains stable.

Represent the commit with the semantic operation, exact before/after semantic
state, and the source-authority handles needed for both directions. A commit
may be cloned only if cloning those handles is intended to retain their
reservations. Its forward publication method should:

1. Recheck source version, lineage, cancellation, content type, and the exact
   source bytes before constructing a topology plan.
2. For an exact no-op, use the existing empty source topology publication, which
   copies the complete artifact byte for byte.
3. For a change, call
   `SourceTopologyPlan::try_replace_source_xml_part(main_partname, target)` and
   then `write_topology_to_stream()`. The plan owns the target token through the
   complete write and source checks. `SourceTopologyPlan::new()` is the public
   plan constructor; the source replacement method is already the intended
   source-authorized path.
4. Return a `Commit`/`Publication` that retains the complete original source
   artifact handle and the complete emitted-artifact identity. Do not expose the
   target as a generic unreserved `Arc`.

The reversible end state needs an explicit inverse publication path. Retaining
the original `SourceXmlPart` alone is insufficient: its replacement proof is
bound to the original source bytes, while the inverse destination contains the
derived bytes. Use the existing exact-artifact restoration seam instead:

* Capture `original_artifact = package.source_artifact()` with the forward
  commit. After writing the changed package, fingerprint the complete emitted
  artifact and retain its length and fingerprint in `Publication` and in the
  inverse patch's source precondition.
* On inverse publication, authenticate the current package against that emitted
  length and fingerprint, authenticate the retained/provided original artifact,
  and call `restore_source_artifact_to_stream()` with a
  `SourceArtifactRestoreProof`. This copies the complete original package and
  preserves every untouched ZIP member and physical detail. The OPC primitive
  checks both providers, cancellation, source freshness, output limits, and
  sink progress before and during output. See
  [`artifact_restore.rs`](../../../../crates/litchi-opc/src/source_backed/artifact_restore.rs#L18-L201).
* A rehydrated or durable inverse must carry an explicit original-artifact
  reference/provider and canonical complete-artifact proof. The existing
  `OriginalArtifactProvider`/`apply_exact_inverse` pattern in the tail-append
  patch is the model; a diagnostic hash alone is never an authorization. See
  [`patch.rs`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream/patch.rs#L861-L912).
* `Patch::inverse()` swaps the semantic before/after direction and requires the
  exact emitted artifact as its source. The inverse must restore the accepted
  original artifact byte for byte, or return a typed stale/source mismatch
  before output. Existing plain-paragraph publication tests exercise this
  distinction between a forward patch and the publication-specific inverse. See
  [`paragraph_copy.rs`](../../../../crates/litchi-docx/src/source_backed/paragraph_copy.rs#L1000-L1067)
  and [`source_backed_paragraph_copy.rs`](../../../../crates/litchi-docx/tests/source_backed_paragraph_copy.rs#L158-L190).

This inverse requirement is part of completion, not optional follow-up. A
specialized managed API that only returns a changed Part or a one-way save is an
intermediate seam for implementation work; it does not satisfy the accepted
snapshot/edit/commit/patch contract.

## Ownership and budget requirements

The implementation must keep these invariants visible in the code and tests:

* The original `PartData` reservation stays alive from capture through edit,
  semantic readback, topology planning, and publication. Cache eviction cannot
  evict an active edit owner.
* Each source-range proof retains its own bounded expected-byte reservation until
  the splice plan drops it. Fragment vectors and edit metadata reserve before
  growth.
* `finish()` reserves the exact candidate output length before allocation and
  retains that reservation in the target `SourceXmlPart` until publication or
  drop. A failed allocation, cancellation, stale-source check, or sink error
  drops all tokens.
* The reversible `Commit`/`Publication` retains the original
  `SourceArtifact` (or an explicit bounded provider reference for a durable
  patch) and the complete emitted-artifact fingerprint. Inverse authentication
  must reserve its fixed workspace and charge source reads/output bytes through
  the active execution context.
* Work is charged for source scanning, fragment generation, range proof, splice
  assembly, XML validation, and semantic readback. Cancellation is checked at
  capture, before and after scans, during bounded loops, before finish, and
  immediately before writing.
* The selected Part-byte, event, depth, operation, replacement-text, and output
  ceilings are checked before their corresponding vectors or buffers grow. The
  first API should accept one replacement operation only; a later bounded-N API
  can prove disjoint ranges explicitly.
* No `Arc<Vec<u8>>` derived from managed `PartData` crosses the DOCX public
  boundary. The only Arc inside the OPC source token is private to the
  source-authorized publication value and is paired with its reservation. An
  owner-carrying generic snapshot is equally valid if every exported view,
  candidate, patch, and inverse retains the same owner.
* Existing signed, encrypted, protected, MCE, unsupported relationship, and
  dependency-ambiguous packages continue to return typed refusal. A changed
  signed package must use an explicit signature policy, as required by the
  accepted validation/security ADR.

These constraints follow the accepted snapshot/edit ownership and I/O/memory
ADRs: [ADR 0003](../../../../docs/adr/0003-snapshots-edits-and-patches.md),
[ADR 0005](../../../../docs/adr/0005-io-memory-and-performance.md),
[ADR 0006](../../../../docs/adr/0006-validation-security-and-compatibility.md),
and the accepted lossless/API priorities in
[ADR 0001](../../../../docs/adr/0001-priorities-and-api-layers.md).

## Test matrix

The temporary specialized seam should be tested first because it exposes the
reservation and source-publication boundaries with a small fixture. Final
acceptance must rerun this matrix through ordinary managed
`document_snapshot()`/`edit_document()` and publish APIs. The existing generic
managed refusal test is therefore a temporary boundary regression; it should
be replaced by the owner-carrying ordinary-path assertions when the required
end state lands.

| Case | Required assertion |
| --- | --- |
| Managed provider, cold source, one ordinary paragraph | Text changes, reopened DOCX reads the requested text, and untouched ZIP members remain source-preserved. |
| Managed provider, warm cache, one edit | Same result with no detached allocation; read diagnostics distinguish warm from cold payload work. |
| Managed filesystem source, one edit | Same semantic and physical result through the file-backed `ReadAt` provider. |
| Exact text no-op | No target output allocation or changed Part; saved artifact is byte-for-byte identical to source, including unsupported members and signatures where the empty plan permits it. |
| Exact memory budget | Budget succeeds at the measured reservation envelope; one byte below fails before target allocation and before sink output. Memory returns to baseline after edit/commit/package drop. |
| Long and short replacements | Existing run boundaries and unknown XML remain exact; Unicode and XML-sensitive characters round-trip with semantic readback. |
| Empty replacement | Allowed only under the existing ordinary paragraph policy; generated empty text slot remains valid and the reopened text is empty. |
| Reversible forward/inverse commit | Forward publication returns a source-checked `Commit`/`Publication`; inverse publication against the exact emitted artifact restores the complete original artifact byte for byte. |
| Reopened inverse | A reopened changed package plus the retained or explicitly provided original-artifact reference restores only when complete current and original fingerprints/lengths match; any mismatch emits no output. |
| Structural text | Tabs, line breaks, control characters, invalid Unicode, and malformed authored text follow existing typed refusal rules. |
| Complex paragraph | Hyperlinks, fields, tracked changes, content controls, drawings, unsupported wrappers, and MCE are refused before output. |
| Source mutation | Mutate the provider after capture, during staging, before finish, and before publication. Each path returns `SourceChanged` or the appropriate typed boundary error and writes no sink bytes. |
| Cancellation | Cancel at capture, scan, range proof, finish, semantic readback, and publication. All reservations release and a caller sink remains empty before output begins. |
| Range safety | Stale, foreign, reversed, overlapping, and derived-payload proofs are refused; no second edit can be issued from a derived `SourceXmlPart`. |
| Security boundaries | Signed, encrypted, protected, and unsupported active-content/dependency closures return typed refusal without partial output. |
| Publication failure | A failing sink reports incomplete output according to the OPC contract and releases source, fragment, and target reservations. |
| Ownership lifetime | Cloning and dropping the specialized edit/commit holds/releases the expected memory and object reservations; package/cache drop reaches zero. |
| Regression | Unmanaged generic paragraph edit/save tests retain their current behavior; managed selective reads retain their current zero-copy borrowed behavior. |

The provider/file cold/warm matrix should be measured only after the focused
correctness and reservation tests pass. No claim about throughput or peak RSS
belongs in this change until the existing measurement harness records it.

## Deferred work

Do not make this first slice support generic `Paragraph` views, all transaction
operations, paragraph insertion/removal, cross-package transfer, relationship
rewrites, durable JSON exchange, or source-backed story/section edits. Those
paths need their own ownership and dependency-closure review. The first slice
may defer a durable wire envelope only if it still provides an in-memory
source-checked reversible patch and exact inverse publication; once durable
patch exchange is claimed, its canonical artifact reference/provider rules are
part of that feature. In particular, `source_backed/story_text.rs` has the same
managed `PartData` versus generic `DocumentSnapshot` boundary; this review does
not silently broaden that API.

The current requirements already demand that `Package::edit_document()` itself
work for managed input. The accepted ADR/goal provide the implementation
direction: use an owner-carrying snapshot storage type, propagate ownership
through every semantic view and patch, reserve every candidate allocation, and
define explicit durable-lineage and inverse-publication handling. The storage
owner must be retained by every clone, paragraph view, candidate snapshot,
`Commit`, `Patch`, and inverse; source-authorized forward publication and
complete-artifact inverse restoration must remain typed and source-checked.
This is implementation work under the existing contracts, not a request for a
weaker API or a new ADR. The specialized source-backed seam may remain as a
measured regression fixture, but it cannot replace this ordinary-path end
state. A one-field reservation added to the current `Snapshot` is not enough.
