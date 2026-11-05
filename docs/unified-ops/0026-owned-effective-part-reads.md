# 0026 — Owned reads of retained effective packages

Date: 2026-10-08. Status: effective read adapter foundation; source DOCX owner
migration remains open.

## Change

Retained OPC candidates now expose `part`, `iter_parts`, `rels` and
`main_document_part`. Ordinary reads work for unchanged candidates as well as
changed candidates; commit-time `with_candidate` retains its existing exact
no-op refusal. Root main-document resolution follows the effective relationship
graph and shares the original source resolver's missing/duplicate/external
relationship checks.

`PreparedPartView` owns its candidate. Metadata lookup and iteration do not
load payloads, and the iterator can outlive the original owner. Added Parts,
removals, relationship changes and replacement content come from the effective
catalog. `declared_uncompressed_size` preserves the distinction between an
unverified physical ZIP declaration and a retained replacement's known length.

`PreparedPartData` owns its Part view and any lazily acquired `PartData`.
Prepared payloads share their existing bytes. Clones allocate no new payload
and keep managed admissions alive; no bare Arc escape is exposed. Each payload
handle intentionally pins the complete candidate, including unrelated overlays,
metadata and transfer monitors. Redacted Debug reports no XML content.

Captured metadata and already captured bytes remain readable snapshots after
later physical source changes or cancellation. New view/payload/proof captures
check the base, transfer sources and execution contexts. Trailing source checks
also run before reporting failed captures, including a missing Part lookup.
`source_xml()` returns a fresh logical revision over effective validated XML;
it is not a physical-only hint. Added/unproven authored XML remains readable as
data but cannot acquire fabricated source provenance. `edit_xml` now uses this
same resolver rather than maintain a separate physical/effective branch.

`physical_source_lineage`, `physical_source_version` and
`physical_source_artifact` explicitly expose the original archive. Logical
edits do not change these values. Limits, execution context, source checks and
exact-artifact no-op status are also available directly on the retained owner;
no raw `SourceBackedPackage` accessor was added.

## Verification

- Final full-feature OPC unit/integration run: 1,084 passed; one existing ignored.
- Owned-view/source-XML focused targets: 64 passed before the final full run.
- New coverage verifies lazy metadata while a large cold Part is unreadable,
  iterator/view/payload survival after owner drop, effective additions/removals,
  root relationship redirection, changed relationships and content types, old
  views after later edits, current XML proof capture and stale-proof refusal,
  physical identity stability, Send + Sync, cross-thread reads, clone accounting,
  final-owner release, transfer-source freshness, cancellation and missing-lookup
  source-error precedence.
- The first strengthened lazy-read test used a header walker on a fixture with
  unsuitable member framing. Rewriting that generated fixture with stored
  members made the denied-byte-range assertion valid; the final full run passes.
- Warning-denied OPC library Clippy: passed.
- Warning-denied OPC doctests: ten passed, including owned reads after owner drop.
- Full-feature facade and Python library compilation: passed.
- DOCX managed document/paragraph targets: 52 passed; XLSX form-control
  scalar lifecycle: 32 passed.
- Glossary rustdoc: three profiles, 14 banned names, zero complete transaction
  families; passed.
- Crate-boundary gate: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Focused formatting and `git diff --check`: passed.

Read-only review found no provenance bypass and confirmed the owned/lazy read
behavior. Its requests for explicit physical access, captured-metadata/pinning
documentation and trailing error-path source checks are addressed.

## DOCX integration requirements and remaining work

Inspection confirmed that adding a candidate field beside the existing source
DOCX package would leave stale physical routes in document/text projection,
metadata, embedded Parts, sections, variables, hyperlinks, stories, SVG,
altChunk and the existing stream publishers. The next migration must use one
effective read/publish adapter. A specialized operation that has not migrated
must refuse an edited candidate rather than silently read or publish the base.
`to_owned_package` must materialize the effective owner explicitly, and physical
source access must remain visibly distinct from effective package access.

Main-document snapshots should capture effective source XML without copying
through an unproven byte buffer. Their retained layouts/admission caches need
logical revision keys or invalidation per changed owner: the current physical
`SourceVersion` cache keys are insufficient. Commit publication must preserve the
retained graph and compare logical edit authority independently of physical
archive lineage. The canonical source package-owning Edit/Commit and immutable
apply shape are not yet implemented by this record.

General topology application, added/authored XML provenance, the catalog
handle-copy cost noted in 0025, canonical DOCX family closure and all remaining
original goal milestones remain required. Whole-candidate payload pinning is a
correctness-first lifetime choice, not an allocation/performance improvement
claim. Paired measurements and the changed-closure cost bar remain open.

## Cleanup

Removed build output, root Cargo.lock and temporary check logs after recording
results. No new worktree or fixture file was created. Base checkout, iWork files,
pre-existing worktrees and tracked standalone lockfiles were preserved.
