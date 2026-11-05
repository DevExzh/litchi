# 0011 — DOCX markup-compatibility admission for document transactions

- Milestone: M2 (write kernel). Implements stages 1–3 of record 0008 §5
  (`docs/unified-ops/0008-docx-source-backed-mce-edits.md`).
- Status: implemented on `uo/0011-docx-mce-admission`, based on `uo/m0a` at
  `5f2f36e088`. Paths below are relative to the repository root.
- `performance_claim: none`. Timings are indicative single-host medians, not
  registered claims (§9).

## Answers in brief

1. **What is admitted.** A DOCX document transaction edits and publishes the
   main part's own bytes while readers see the markup-compatibility (MCE)
   projection. It is now admitted whenever one bounded proof shows the two
   agree on everything either reads: the part never mentions the MCE
   namespace, binds it and uses nothing in it (stage 1), or carries valid
   `Ignorable` directives whose projection hides nothing a reader sees
   (stage 2). All 19 `test-data/ooxml/docx` fixtures are admitted on both
   owners and under an execution context; 0 were admitted by the source-backed
   route before.
2. **What stays refused.** `AlternateContent`/`Choice`/`Fallback` anywhere,
   `ProcessContent`, `MustUnderstand`, `PreserveElements`/`PreserveAttributes`,
   unknown MCE attributes, invalid `Ignorable` lists, an `Ignorable` below
   `w:document` that lists a namespace the profile does not understand,
   ignorable markup that holds understood content, and everything else the
   proof cannot establish. Each keeps its historical typed error.
3. **Decided on decoded names.** A namespace declaration spelled with
   character references (`…/200&#54;`, record 0008 W7) now binds the MCE
   namespace for the proof, for the shared codec's fast path and for
   `active_offsets`. Every scanner the audit found that recognized the
   namespace on values as written now refuses such a part (§4).
4. **One policy for both owners (stage 3).** The owning `litchi_docx::Package`
   runs the same admission before any main-document transaction, and refuses a
   changed publication under enforced protection or tracked revisions with the
   source-backed route's exact error. The managed route reads and edits all 19
   fixtures (0 before). Document variables admit 19/19 (15 before).
5. **Cost.** The proof costs 0.26–0.41× of a full MCE pass on the fixtures
   (0.55× on one 3 KB part where setup dominates). An admitted source-backed
   `edit_document` is cheaper than the refusal it replaces (0.51–0.91×), and
   the MCE-free path is unchanged (0.96–1.00×). The owning package pays the
   proof and the protection policy it did not run before (§9).

## 1. The rule

One proof, `litchi_ooxml_common::mce::admission::admit`
(`crates/litchi-ooxml-common/src/mce/admission.rs`), returns one verdict per
part, or a refusal:

| Verdict | Projection relative to the part | Example |
|---|---|---|
| `Unbound` | identical: the codec borrows its input | `listWithLgl.docx` |
| `DeclaredOnly` | spelling only: whitespace outside the root, `<q/>` → `<q></q>`, rebuilt start tags | Word 2007 `xmlns:ve`, 13 fixtures |
| `Ignorable` | spelling, and the `Ignorable` attributes removed | root `mc:Ignorable="w14 wp14 w15"`, 6 fixtures |
| `IgnorableOpaque` | as `Ignorable`, and attributes and subtrees in ignorable namespaces the profile does not understand are dropped; no dropped subtree holds an element or attribute of an understood namespace (`xml:` excepted) | W6; `wp14:anchorId`, `wp14:sizeRelH`; Word 2013 settings leaves |

`Admission::is_lexical()` is true for the first three. Refusals are
`Element`, `Directive`, `NestedIgnorable`, `HiddenContent` and `Unproven`.

The proof admits a part only when the codec would process it successfully:
it repeats the codec's checks with the codec's own namespace scope (`Scope`,
shared with `codec.rs`), name rules and limits — well-formedness, one
document element, UTF-8 names and values, qualified names with bound
prefixes, no attribute named twice under one prefix or two (the lexical and
expanded checks), valid declarations within the binding limit, a valid
`Ignorable` list, only predefined or character references in visible text, no
DTD or processing instruction, no late XML declaration, and the depth,
attribute, token, input and output bounds. It is stricter in two places
(a declaration rebinding `xml` or binding `xmlns`, and profiles with
extension elements are refused). A generated differential checks the
property on 680 inputs (2 namespace spellings × 17 root variants × 20
bodies): whenever the proof admits, processing succeeds and the projection
carries exactly the part's elements, attributes and text, less what the
verdict says is dropped.

**Where each verdict is admitted:**

| Route | Admits | Refusal identity (unchanged) |
|---|---|---|
| Source-backed, unmanaged: `document_snapshot`, `edit_document`, `publish_document_commit_to_stream`, hyperlink sanitization and redaction snapshots | all four (profile: OOXML baseline + `w14` + `w16du`) | the projection's own error (e.g. `Mce(MustUnderstand)`), else `UnsafeEdit { reason: "source-backed document transactions do not support markup-compatibility branch selection" }` |
| Owning package: `document_snapshot`, `edit_document`, `apply_document_patch`, `apply_durable_document_patch`, history, transfers | all four | the projection's error, else `UnsafeEdit { operation: "document_snapshot", reason: "document transactions do not support markup-compatibility branch selection" }` |
| Managed (execution context): `document()`, `write_text_to`, `document_snapshot`, `edit_document` | lexical only — the managed route reads source bytes and builds no projection, so nothing may be dropped from them | `UnsafeEdit { operation: "source-backed document read", reason: "markup-compatibility preprocessing would require an unbudgeted owned payload" }` |
| Source-backed document variables (settings part) | all four (profile: OOXML baseline, as the settings readers use) | the projection's error, else the existing document-variable `UnsafeEdit` |
| Owning package paragraph transfer (donor) | lexical only: dropped markup is ignorable only under the donor's root | `TransferRefusal::MarkupCompatibilityScope` (new) |

A refusal never decides on its own error: when the proof does not hold, the
caller runs the projection exactly as before, so every refused document keeps
today's error.

**Two refinements of 0008 §5.**

- *Nested `Ignorable`.* Litchi's own revision writer emits
  `mc:Ignorable="w16du"` on every `w:ins`/`w:del` it authors with a UTC date.
  Refusing all non-root directives broke the eager owner's existing tests on
  those documents. An `Ignorable` below `w:document` is admitted when every
  namespace it lists is understood by the profile: the codec never drops an
  understood name, so the directive changes nothing but its own presence,
  wherever content moves. One that lists a non-understood namespace stays
  refused (`NestedIgnorable`).
- *Opaque drops.* 0008 admits W6-style drops. They are admitted on the
  unmanaged routes, where readers see the projection and the transaction
  scanners treat foreign markup as opaque. The managed route refuses them,
  because its readers see the source bytes; the same rule keeps managed reads
  equal to unmanaged reads.

## 2. What is admitted, per stage

The 19 fixtures, measured with a census probe (Evidence) at the base commit
and on this branch. The stage columns attribute the result to the verdict
class each stage introduces (0008 §5); both stages land in this change:

| Route | Before | Stage 1 (+13 declared-only) | Stage 2 (+6 root `Ignorable`) | After |
|---|---:|---:|---:|---:|
| Source-backed `edit_document` | 0 | 13 | 19 | **19** |
| Source-backed changed publication | 0 | 10 (3 protected refused) | 16 | **16** |
| Owning-package `edit_document` | 19 | — | — | **19** |
| Owning-package changed publication | 19 | — | — | **16** (stage 3: 3 protected now refused) |
| Managed `document()` read | 0 | 13 | 19 | **19** |
| Managed `document_snapshot` | 0 | 13 | 19 | **19** |
| Document variables | 15 | 15 | 19 | **19** |

Stage 1: `FancyFoot`, `Headers`, `documentProperties`,
`documentProtection_{comments,forms,no_protection,readonly}_no_password`,
`drawing`, `endnotes`, `footnotes`, `headerFooter`, `headerPic`,
`table_footnotes`. Stage 2: `ComplexNumberedLists`, `Hyperlink`, `comment`,
`table-alignment`, `table-indent`, `testComment`. The four document-variable
fixtures refused before (`ComplexNumberedLists`, `Hyperlink`,
`table-alignment`, `testComment`) need the opaque-leaf clause: the baseline
profile drops their `w14:docId`/`w15:*` settings leaves.

The wider corpus (63 DOCX-family files under `test-data/`, 61 openable), same
probe:

| Route | Before | After | Not admitted after |
|---|---:|---:|---|
| Source-backed `edit_document` | 1 | 59 | the run-level `AlternateContent` file (stage 4); `alt-chunk-header.docx` (layout scan: section properties not last, unrelated to MCE) |
| Source-backed changed publication | 1 | 52 | + 4 protection/tracked revisions, 3 signed |
| Owning-package `edit_document` | 59 | 58 | the `AlternateContent` file is now refused |
| Owning-package changed publication | 56 | 51 | + 4 protection/tracked revisions (3 signed refused before and after) |
| Managed read | 1 | 57 | 3 opaque-drop files, 1 `AlternateContent` file |
| Managed `document_snapshot` | 1 | 56 | as above, + `alt-chunk-header.docx` |
| Document variables | 30 | 59 | 2 files without a settings part |

Verdicts over the 61 openable files: 1 `Unbound`, 25 `DeclaredOnly`,
31 `Ignorable`, 3 `IgnorableOpaque`, 1 refused (`Element`). These match
record 0008's census.

## 3. What stays refused

Witnesses (main parts in minimal packages), every route checked by
`crates/litchi-docx/tests/mce_admission.rs`:

| Witness | Source-backed | Owning package | Managed read | Managed snapshot |
|---|---|---|---|---|
| W1 `AlternateContent` in a run | `UnsafeEdit` | `UnsafeEdit` | `UnsafeEdit` | `UnsafeEdit` |
| W2 `AlternateContent` in the body | `UnsafeEdit` | `UnsafeEdit` | `UnsafeEdit` | `UnsafeEdit` |
| W3 `ProcessContent` | `UnsafeEdit` | `UnsafeEdit` | `UnsafeEdit` | `UnsafeEdit` |
| W4 ignorable prefix holding a run | `UnsafeEdit` | `UnsafeEdit` | `UnsafeEdit` | `UnsafeEdit` |
| W5 `MustUnderstand="x"` | `Mce(MustUnderstand("urn:x"))` | `Mce(MustUnderstand("urn:x"))` | `UnsafeEdit` | `UnsafeEdit` |
| W7 body `AlternateContent` via `…/200&#54;` | `UnsafeEdit` | `UnsafeEdit` | `UnsafeEdit` | OPC `SourceBackedOverlayUnavailable` |
| nested `Ignorable` of a foreign namespace | `UnsafeEdit` | `UnsafeEdit` | `UnsafeEdit` | `UnsafeEdit` |
| W6 ignorable attribute and leaf | admitted | admitted | `UnsafeEdit` | `UnsafeEdit` |
| declared only via `…/200&#54;` | admitted | admitted | admitted | OPC `SourceBackedOverlayUnavailable` |

The managed snapshot captures the part through the OPC source-XML seam,
which already refuses any namespace declaration whose written value differs
from its decoded value ("source XML requires literal namespace bindings").
That pre-existing policy is kept; it is the first refusal a managed
transaction meets for an escaped declaration.

The unit matrix in `crates/litchi-ooxml-common/src/mce/admission_tests.rs`
also pins `PreserveElements`, `PreserveAttributes`, an unknown MCE attribute,
an `Ignorable` listing an unbound prefix, a duplicate prefix or the MCE
prefix itself, two `Ignorable` attributes through two prefixes, a dropped
subtree holding `r:id`, a dropped document element, and the malformed and
over-limit inputs the codec rejects.

## 4. Decided on decoded names: the codec and its neighbours

`mce::mentions_namespace(xml)` is now the codec's fast-path test, in
`process_markup_compatibility` and in `active_offsets`: the literal namespace
anywhere (the historical test, kept so a mention in text or in malformed
markup still goes to the parser), or a namespace declaration whose value
holds a reference and decodes to the namespace through quick-xml's own
`unescape`. Only a part holding `&` is searched for declarations. This moves
behaviour for every reader of such parts in every OOXML format: they now
apply MCE instead of reading both branches. The streaming processor already
decoded declarations, so the two processors now agree.

Record 0008 A3 noted this move. A read-only audit of the workspace then found
scanners that recognize the namespace on values as written (literal searches
or quick-xml's resolver) next to readers that use the codec. With the codec
fixed, those would disagree on an escaped part. Each now refuses such a part,
or reports it, through `mce::mentions_namespace` (literal searches) or
`mce::binds_namespace_by_reference` (resolver-based scans), with its existing
refusal:

- XLSX: worksheet shared-traversal admission (`raw/worksheet/mod.rs`), row
  visibility (`row_visibility/rewrite.rs`), `x14ac` capture choice
  (`raw/worksheet/x14ac.rs`), hyperlink editing (`hyperlinks/snapshot.rs`),
  reduced readback (`workbook/edit/semantic/transaction.rs`), connection
  output budget (`connections/codec.rs`), workbook catalog edits
  (`raw/catalog_edit/codec.rs`), tab state (`tab_state/source.rs`), worksheet
  properties (`raw/worksheet_property.rs`), drawing attach
  (`workbook/edit/drawing_transfer.rs`), sheet views (`raw/sheet_view_edit.rs`),
  auto-filter (`auto_filter/package.rs`), calculation properties
  (`calculation_properties/rewriter.rs`), reference renames
  (`raw/reference_edit.rs`), Data Model removal
  (`workbook/data_model/removal.rs`).
- PPTX: transition guard (`presentation/transition.rs`), positional metadata
  (`presentation/order.rs`), embedded OLE slides
  (`presentation/embedded/ole/slide/codec.rs`), animation hosts
  (`animations/package.rs`).
- DrawingML: Theme Family ownership (`theme/family/part.rs`): an escaped
  binding now counts as a hidden owner, so no direct owner is added beside it.
- DOCX: redaction closure (`redact.rs`), story-hyperlink diagnostics
  (`story_hyperlinks.rs`), tracked-revision authoring
  (`revision/authoring.rs`), Ink inventory and placement (`ink/codec.rs`,
  `ink/placement.rs`).

Not changed, recorded as follow-ups: the PPTX validation report's MCE fact
(`validation.rs`, report only); the DOCX footnote-column and content-control
writers, which do not see an escaped `mc:Ignorable` on the element they edit
and append a second one. That second hazard predates this change, and the
codec rejects the duplicate when the result is read back.

## 5. Stage 3: one policy whichever owner holds the document

- **Admission.** `Package::document_snapshot` (owning package) proves the
  main part before building the snapshot, so `edit_document`,
  `apply_document_patch` (rescanned route), `apply_durable_document_patch`,
  undo/redo and transfer planning inherit it. A reused candidate on
  `apply_document_patch`'s exact-source route is held to the same proof before
  publication, however its snapshot was built. The verdict is remembered per
  main-part allocation (a `Weak` reference). The source-backed package
  remembers it per source revision, so its publication does not prove the
  same bytes again.
- **Protection.** `document_policy::validate_changed_owned_document`
  (`crates/litchi-docx/src/source_backed/document_policy.rs`) applies the
  source-backed publisher's settings policy to the owning OPC graph: the same
  relationship, conformance-family and content-type rules, the same bounded
  inspection of the exact settings bytes, and the same error,
  `UnsafeEdit { operation: "changed document publication", reason: "document
  or write protection, or tracked revisions, is enabled" }`. It runs on every
  real change published by `apply_document_patch` (so `publish_document_edit`,
  `publish_document_commit`, `publish_document_commit_with_history`) and by
  `apply_durable_document_patch`. Exact no-ops do not run it. Undo and redo of
  recorded history are inverse restorations and are not checked, as the
  source-backed inverse publication is not.
- **Transfers.** A paragraph moved to another document carries its source
  bytes; markup that the donor's root `Ignorable` drops would lose that scope.
  `plan_paragraph_transfer_from` requires a lexical donor
  (`TransferRefusal::MarkupCompatibilityScope`, a new variant of a
  `#[non_exhaustive]` enum).
- **Managed route.** `ensure_source_document_xml` no longer refuses MCE
  declarations and directives itself; it still refuses MCE elements, then
  runs the shared proof with the visible profile under the scan's existing
  workspace reservation, charging one work unit per byte plus one per event,
  with cancellation checked per event. Only lexical verdicts pass. Managed
  `to_owned_package` is unchanged (still refused).
- **Unified layer.** `litchi-ops` needs no code change: `Docx::edit_document`
  reaches the leaf on both owners, and both now return the same verdicts and
  errors.

## 6. Behaviour changes

1. Source-backed `document_snapshot`, `edit_document`, document-commit
   publication and the hyperlink sanitization/redaction snapshots admit Word
   documents whose MCE use cannot change content (0/19 → 19/19 fixtures).
   Exact no-ops on those documents now republish the source byte for byte
   (they were refused at the guard before).
2. Source-backed document variables admit a root `Ignorable` that drops
   opaque settings leaves (15/19 → 19/19).
3. Managed `document()`, `write_text_to`, `document_snapshot` and
   `edit_document` admit parts whose projection differs in spelling only
   (0/19 → 19/19; corpus 1 → 57 reads).
4. **Owning package, admission (breaking).** Every main-document transaction
   now refuses `AlternateContent`, `ProcessContent`, `Preserve*`, a nested
   foreign `Ignorable`, ignorable markup holding understood content and their
   escaped equivalents with `UnsafeEdit`, and returns the projection's error
   (for example `MustUnderstand`) instead of editing the raw bytes. Corpus:
   1 file newly refused.
5. **Owning package, protection (breaking).** A changed publication refuses
   enforced document or write protection and tracked revisions (3 fixtures;
   corpus 4 files, `redline-range-comment.docx` for tracked revisions).
6. **Owning package, transfers (breaking).** Transfer planning refuses a
   donor whose projection drops opaque markup.
7. **Shared codec.** A namespace declaration spelled with references binds
   the MCE namespace for `process_markup_compatibility`, `process_ooxml`,
   `process_part*`, `process_str` and `active_offsets`, so readers in every
   OOXML format apply MCE to such parts.
8. **Neighbouring scanners.** The XLSX, PPTX, DrawingML and DOCX scanners in
   §4 refuse, or report, a part that binds the namespace through references.
9. `litchi-ops`: `Docx::edit_document` succeeds on Word fixtures from either
   owner; M0a's refusal test is inverted.

## 7. Tests

- `crates/litchi-ooxml-common/src/mce/admission_tests.rs` (8): the predicate on
  literal, escaped, fully-referenced and lookalike declarations; the codec
  applying W7 and `active_offsets` agreeing; the witness matrix with a
  projection check per admitted case; prefix rebinding; malformed and
  over-limit parts never admitted; the generated soundness differential; the
  baseline settings profile; the observer hook.
- `crates/litchi-docx/tests/mce_admission.rs` (11): the 13/6 stage split;
  source-backed edit + publish equals the eager output on 16 fixtures, with the
  root tag and every `mc:`/`ve:`/`Ignorable` token preserved and every other
  member identical in order, bytes, compressed size and method; the 3 protected
  fixtures refused identically by both owners before output; exact no-ops
  republish all 19 byte for byte; the projection oracle
  (`Snapshot::from_xml(raw)` and `(visible)` agree on paragraph, table and
  block-control counts and every paragraph's text); document variables 19/19;
  managed reads equal unmanaged reads and every reservation is released; the
  witness matrix on both owners and the managed route; the proof's work charge
  and cancellation identity; the document-variable rule; the DOCX scanners of
  §4; transfer donors.
- `crates/litchi-docx/tests/source_backed_managed_document_edit.rs`: a managed
  edit and publication of an `Ignorable` document equals the unmanaged one.
- `crates/litchi-docx/tests/source_backed_sections.rs`: the managed escaped
  declaration case now reads (decoded, unused) and its `AlternateContent` use
  is refused.
- `crates/litchi-ops/src/format/docx/tests.rs`: M0a's refusal test inverted to
  `editor_admits_word_markup_compatibility_on_both_owners`; new
  `editor_refuses_alternate_content_on_both_owners`.
- Gate regressions for escaped bindings: XLSX (4), PPTX (4), DrawingML (1).
  The DrawingML one was checked to fail without its guard.

## 8. Gates

Run with `CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0`, offline, with
`3rdparty` symlinked from the main checkout for the DOCX corpus tests and the
`xlsx-form-control-properties` evidence directory added to the sparse
checkout for an XLSX `include_bytes!`:

| Command | Result |
|---|---|
| `cargo fmt --check -p litchi-ooxml-common -p litchi-docx -p litchi-xlsx -p litchi-pptx -p litchi-drawingml -p litchi-ops` | clean |
| `cargo clippy -p <crate> --all-targets --all-features -- -D warnings` for litchi-docx, litchi-ops, litchi-ooxml-common, litchi-xlsx, litchi-pptx, litchi-drawingml | clean |
| `cargo test -p litchi-docx --all-features --no-fail-fast` | 2,022 passed, 32 ignored |
| `cargo test -p litchi-ops --all-features` | 203 passed |
| `cargo test -p litchi-ooxml-common --all-features` | 451 passed |
| `cargo test -p litchi-xlsx --all-features --no-fail-fast` | 2,137 passed |
| `cargo test -p litchi-pptx --all-features --no-fail-fast` | 1,309 passed, 3 ignored |
| `cargo test -p litchi-drawingml --all-features --no-fail-fast` | 369 passed |
| `RUSTDOCFLAGS=-Dwarnings cargo doc -p litchi-docx --no-deps --all-features` | clean |
| `RUSTDOCFLAGS=-Dwarnings cargo doc -p litchi-ooxml-common --no-deps --all-features` | clean |
| `cargo check -p litchi --all-features` | clean |

## 9. Performance evidence

**Method.** One throwaway binary (`uo0011_bench`, see Evidence) built from
the same source at the base commit and on this branch, release profile,
pinned to CPU 21. Each row ran in the order before, after, after, before, and
the whole matrix ran twice. Each run has 3 warm-ups, then samples until
3,000 or 400 ms (at least 71 per run). Cells are the mean of the four run
medians per side, in µs. Inputs are in-memory. The host was shared with other agents
(load average 9–11 on 32 cores). Most rows spread under 4% between runs;
`saut_page` rows reach 11–16%. "Control" copies replace the MCE URI in the
main part with an equal-length unknown URI (0008's marker-stripped control),
so they take the MCE-free path. `listWithLgl.docx` is the corpus's MCE-free
file.

| Leg | FancyFoot (2.9 KB part) | testComment (3.1 KB) | table-indent (72 KB) | drawing (288 KB) | saut_page (2.5 KB part, 2.96 MB file) |
|---|---:|---:|---:|---:|---:|
| source-backed open + `edit_document`: before (refused) | 71.6 | 94.1 | 502.1 | 2,108.2 | 176.5 |
| … after (admitted) | **63.2** (0.88×) | **85.5** (0.91×) | **255.9** (0.51×) | **1,206.0** (0.57×) | **161.4** (0.91×) |
| … control before → after | 55.2 → 53.6 | — | — | 594.6 → 593.7 | 157.2 → 150.9 |
| source-backed lifecycle (open, edit, commit, publish): after | 165.3 | 218.7 | 846.0 | 5,187.0 | 339.2 |
| … control after (same file without MCE) | 154.7 | — | — | 4,663.4 | 329.8 |
| owning package open + `edit_document`: before → after | 75.8 → 84.0 (1.11×) | 96.5 → 115.1 (1.19×) | 146.9 → 270.7 (1.84×) | 683.7 → 1,256.9 (1.84×) | 238.2 → 253.8 (1.07×) |
| … control before → after | 75.4 → 73.6 | — | — | 628.6 → 641.9 | 221.1 → 242.3 (noisy) |
| owning package lifecycle (open, edit, commit, publish, save): before → after | 143.6 → 172.2 (1.20×) | 175.4 → 229.2 (1.31×) | 477.6 → 614.1 (1.29×) | 3,705.0 → 4,311.7 (1.16×) | 400.8 → 419.9 (1.05×) |
| … control before → after | 142.1 → 160.2 (1.13×) | — | — | 3,592.9 → 3,637.3 | 406.3 → 403.3 |

`listWithLgl.docx` (no MCE): source `edit_document` 25.9 → 25.2, source
lifecycle 52.4 → 51.0, owning `edit_document` 30.1 → 29.4, owning lifecycle
57.1 → 56.4.

The proof itself, against the full MCE pass it replaces and the transaction
layout scan it accompanies (separate probes, medians in µs, same host):

| Part | Proof | Full MCE pass | Layout scan | Proof / pass | Proof / layout |
|---|---:|---:|---:|---:|---:|
| FancyFoot | 8.3 | 20.4 | 5.8 | 0.41 | 1.43 |
| testComment | 14.2 | 25.9 | 4.1 | 0.55 | 3.44 |
| table-indent | 116.7 | 441.6 | 80.0 | 0.26 | 1.46 |
| drawing | 531.9 | 1,579.3 | 343.3 | 0.34 | 1.55 |
| saut_page | 8.6 | 22.7 | 4.5 | 0.38 | 1.92 |

Readings, not claims:

- The source-backed MCE-free path pays nothing extra: the new guard checks
  `mentions_namespace` before it builds the MCE profile, and the old guard
  built that profile on every call. On an MCE-free part the codec's own fast
  path now also scans declarations when the part holds `&`: 3.2 → 7.9 µs on
  the 288 KB control part, unchanged (0.2 µs) on the 2.9 KB one.
- An admitted source-backed lifecycle costs about one proof more than the
  same file without MCE (+7% FancyFoot, +11% drawing); the publication reuses
  the per-revision verdict.
- The owning package now runs the proof (stage 3) and, on publication, the
  protection policy (one settings decode and two bounded scans: +18 µs on
  `FancyFoot.control`). On large MCE parts the proof dominates the added cost
  of `edit_document` (+573 µs on drawing).
- On parts of a few KB, setting up the codec's namespace scope for a
  30-declaration Word root dominates (testComment 14 µs).

## 10. What is not claimed

- No timing is a registered performance claim; the method above is
  indicative, on a shared host, with in-memory sources.
- The proof's agreement with the codec is established by construction (the
  codec's own scope and limits), by the generated differential and the
  witness matrix, and on 63 real files; it is not a formal proof.
- Run-level `AlternateContent` (stage 4, 1 corpus file) remains refused.
- The managed route still refuses `IgnorableOpaque` parts (3 corpus files);
  admitting them needs a budgeted projection or a reader-by-reader argument
  that dropped markup is invisible.
- The §4 audit covered `litchi-xlsx`, `litchi-pptx`, `litchi-docx`,
  `litchi-ooxml-common`, `litchi-drawingml` and `litchi-xlsb`; ODF and iWork
  crates do not use the OOXML codec and were not audited.
- No Microsoft Word round trip of any output was performed.

## 11. Follow-ups

- Fold the proof into the transaction layout scan, which tokenizes the same
  part: on drawing that would remove most of the 1.55× of a second pass.
- Stage 4: run-level `AlternateContent` with a mapped edit and readback
  (0008 §5).
- The PPTX validation report and the DOCX footnote-column and content-control
  writers named in §4.
- The owning package's other publishers (story edits, merges) do not run the
  protection policy; this record covers the main-document transaction family.

## Evidence and reproduction

Nothing below is committed: the probes and their outputs lived in the
session scratchpad, and the examples were removed from `crates/` before the
commit.

- **Bench** `uo0011_bench.rs` (SHA-256
  `22872059190b73bc7bfb0c4f1d2193f2f2128a08c834e30af934c13698c47f70`), built as
  a `litchi-docx` example with `cargo build --release -p litchi-docx --example
  uo0011_bench --offline`: before binary (base `5f2f36e088`)
  `48c3ee176c760f718aafd92d3a33d2f6df669d2efda362976d4883ef6d76ac3e`, after
  binary `e4df9a63d1b225c069ab414fb918ae14b1839910c132bc2ca3eee758021b3e45`.
  Modes `source`, `source-edit`, `eager`, `eager-edit`; driver: for each mode
  and file, `taskset -c 21 <bin> <mode> <file> 400` in ABBA order.
- **Control fixtures**: copies whose `word/document.xml` replaces
  `…/markup-compatibility/2006` with `…/markup-compatibility/2999`, every
  other member re-written with its original compression method.
- **Proof probes**: `uo0011_scan` (litchi-ooxml-common example: literal
  search, `mentions_namespace`, `admit`, full codec pass) and
  `uo0011_layout` (litchi-docx example: `Snapshot::from_xml`), 2,000 samples
  after 200 warm-ups, pinned to CPU 21.
- **Census** `uo0011_census.rs` (SHA-256
  `67c8add329838ff9098b00366d100f07016621390fc4101af44761baf3c5b37d`; release
  binary `2b1b6086a534cc619cf9b48f07711abe897015b23e8054e63218877aee3856e4`),
  run on every `*.docx|*.docm|*.dotx|*.dotm` under `test-data/`. The base
  variant replaces the verdict column with a literal-namespace test (the proof
  does not exist at the base) and was built from `git archive 5f2f36e088`
  (binary `c8b501af02eb1c97f037d67a6f8595acb91acf6c88e8aec0bdc2ba5644d40331`).
- Toolchain rustc 1.95.0; AMD EPYC host, Linux 7.0.0-1013-aws.
