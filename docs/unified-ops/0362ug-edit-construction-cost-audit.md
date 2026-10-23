# 0362ug — Edit construction cost audit

Record 0351 section 2.1 item 8 requires `Package::edit*()` to be O(1) in
package size. Section 2.6 item 3 additionally prohibits copying part payloads
at edit entry. Passing the public-signature inventory does not establish
either property. The final sweep found constructors that still materialize
source state before any operation is staged. These are required corrections,
not accepted performance exceptions.

This audit starts from assembled source `7559f44aa975254bd96f6e28eee17366b0ebeb6a`.
The ODG correction and ODB/ODM correction are now assembled at
`8779d059add6c90c74552801372c44f770db9f0f`. Other rows below remain open.

## Confirmed scopes

| Scope | Entry cost found | Correction record and state |
| --- | --- | --- |
| ODG | Primary edit copied content and styles XML, including unused copies for pure joined package replacement. | 0362utf, reviewed and merged at `c3f56b7a9d`; lazy part state retains the source package and materializes on mutation. |
| ODB, ODM | ODB copied content XML; ODM title edits read raw meta.xml and copied projected title strings. | 0362ugb and 0362ugb-a, reviewed and merged at `8779d059ad`; borrowed XML/title state and deferred changed-title publication. |
| ODP | Package edit builds a complete mutable presentation and copies validated slides/source fragments; flat edit clones the slide vector. | 0362uga, correction in progress. The already-lazy tail-append entry is not this finding. |
| ODS | Source cells construct protection state; DDE and sheet-metadata constructors load source projections; package resource-count validation enumerates archive members; flat edits allocate per-sheet staging. | 0362ugc merged at `f144cd6029`; 0362ugca follow-up required to remove the new resource census from open. Other deferred parsing/protection, cheap admission and sparse flat staging remain. |
| DOC | Primary/profiled and scoped families serialize/copy the source at entry; paragraph source-backed entry scans the source prefix. | 0362ugd, correction in progress. |
| XLS | Primary and source-backed cells copy/index source state; scoped capture serializes the archive for comments, visibility, objects, names, charts and pivots. | 0362uge, correction in progress. |
| XLSX | Source-backed worksheet and defined-name/cell constructors load and decode selected parts or workbook state before staging. | 0362ugf, correction in progress. |
| Pages | Body-table constructors rebuild rooted inventories; chart data, drawable order, headers/footers and footnotes decode or copy selected/source graph state. | 0362ugg, correction in progress. |
| Keynote | Text/name selection can decode a whole show; chart data and table/chart constructors decode selected payloads or scan the graph. | 0362ugh, correction in progress. |
| Numbers | Title/image selection scans rooted graph state; chart and merge constructors retain decoded data/regions. | 0362ugi, correction in progress. |
| PPT | Primary and metadata edits serialize and parse the whole CFB. `scoped::Base::capture` copies `to_encrypted_bytes()` into a new shared buffer. | 0362ugj, correction required. |
| PPTX | In-memory `Package::edit` captures all slides, notes and name indexes and fingerprints parts; opened edit construction also clones the slide vector and name index. | 0362ugk, correction required. Source-backed primary and staged part constructors are already cheap. |
| DOCX | Ordinary and source-backed document entry parses/indexes main XML; Ink scans package parts and bindings; OPC edit clones graph metadata; several source-backed document families capture source projections at entry. | 0362ugl, correction in progress. Main-part `blob_arc()` itself is cheap; subsequent snapshot parsing is the cost. |
| XLSB | Primary and cell-value constructors serialize source before-images; several owner families read/clone selected graphs, and a cold `From<OpcPackage>` projection can materialize through generic staged entry. | 0362ugm, correction in progress. |

PPT's accepted 0362j work repairs replay and first-commit digest cost. It
does not authorize serialization at edit construction; the new finding is
separate. Cheap package `Arc` clones and cached immutable handles are not
findings. RTF's opened source tail-append entry remains cheap; its committed
deferred owner rematerialization is explicitly covered by 0359a. Likewise,
0358a's explicit rematerialization of a previously committed deferred ODS
owner is distinct from eager work on pristine source edit entry.

The remaining ODF owner audit was checked directly against the `8779d059ad`
tree. ODC and ODI facade edits clone flat snapshot Arcs; OTH and formula edits
retain a source reference and empty staging vectors. Their package snapshots,
and ODT's archive data and prepared index, are Arc-backed. Common metadata
entry only clones the owner. No source-sized edit-entry scan or payload copy
was found in these routes. This covers the separate ODC/ODI/OTH/formula crates
as well as ODT; the auditor's earlier assumption that they all route through
ODT was corrected before this disposition was recorded.

OLE-common object Package and OGraph Package entry retain shared source state.
The bounded DataSpaces, smart-tag, toolbar and similar codec Snapshot editors
are separately scoped by 0359; model/draft cloning there is not an archive
Package entry or a newly copied archive-part snapshot. Those observations do
not expand the Package O(1) finding to unrelated raw codec drafts.

## Review and validation boundary

Corrections retain immutable source handles, selectors and options until an
operation actually requires the source data. They must preserve source and
signature admission, typed budgets before retention, exact no-ops, patch
replay and inverse behavior, bounded deferred output and every existing
capability. Moving work into `open()` or hiding lazy failure behind a panic
does not satisfy this contract. A necessary fallible-accessor change migrates
every consumer in the same commit.

ODG's final source `05c15847c1a4ffc3f55e50c9ecc7609c9b338e0a` passes 23
library tests, 155 integration tests and one doctest, with strict checks;
0362utf records exact external logs. Independent reviews cover joined
composition, canonical effects, projected style dependencies, replacement
policies, optional styles and lazy XML lifetime. The merged correction
retains the complete return/correction history through receipt `0dca7284f9`.

ODB/ODM source `c44e1c57968cc1ddf227e814c85571b905c422bf` passes 106/58
regular tests respectively, one doctest each, and strict Clippy. Record
0362ugb-a labels its external receipts as reduced summaries of terminal
output, not raw command logs. Review confirms the borrowed state and
semantic behavior. ODB helper paths may materialize before discovering a
self-move or validation failure; construction/read-only/discard stay borrowed.
ODM raw meta.xml failures move from construction to a changed-title commit.

ODS source `2e2f902179d8ea107d05d3805542e89d9a42f27a` is independently
reviewed and passes 1,767 tests across 119 successful result blocks, including
702 library tests. Strict all-target/all-feature Clippy also passes. Both
final commands use the repository's pinned Rust 1.95 and tracked root Clippy
configuration. The raw logs are
`/home/zhuhe/code/litchi-worktrees/.coord/ods-lazy-entry/ods-full-tests-pinned195.log`
(182,062 bytes, SHA-256
`81f35af676d3546b6f94f7b9253b72e3f5be9026677ace21a288855afd3b9bba`)
and `ods-all-targets-clippy-pinned195.log` beside it (729 bytes, SHA-256
`72fd3add599ba9f7e2557891ec1a07f579ac187eeb2ef1d0717681496079782d`).
The coordinator verified both hashes, every result block and actual root
configuration against Git. A preceding fixture expansion had excluded root
configuration and selected Rust 1.98; that run is historical evidence only.
All 37 required fixture paths were already tracked; no fixture was added.

A subsequent coordinator call-path review returned the resource-count cache
for correction as 0362ugca. `Archive::from_family_package` now calls `files()`
and counts filtered names during open, adding the complete name allocation and
census there. Passing tests and a cheap edit constructor do not justify that
cost relocation. The successor must derive the fact during an already-required
traversal or defer admission until first use, preserving limits before
retention. The 0362ugc test receipts remain valid for their exact source; no
native/performance acceptance is claimed for the returned census change.

The reviewed glossary policy still parses as 55 profiles after these merges.
This is neither a fresh rustdoc inventory nor the 53-check gate. Complete
workspace Clippy on earlier source `8225b82dde` remains source-pinned evidence
and does not cover the subsequent constructors. Native qualification must
use the final assembled candidate. Comprehensive redaction audit 22 is the
only deferral; its content-free error and diagnostic rules still apply.
