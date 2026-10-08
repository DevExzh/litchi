# 0029 — DOCX effective owner and physical-operation boundary

Source-backed DOCX owners can adopt a retained `PreparedTopology` from their
physical lineage through `with_prepared_topology`. Adoption checks the effective
main relationship/content type and all source/context fences without reading
ordinary payloads. This is an OPC ingress boundary, not semantic edit validation.
It leaves the original owner and its views unchanged.

`document`, `to_owned_package`, `to_stream`, `check_source` and `source_version`
now operate on the effective owner. Main-document reads follow effective root
redirection and payload replacement. Publication preserves unselected physical
members; explicit owning conversion uses the effective catalog from record 0028.
`source_version` is still a physical version, but now fences donor sources too.
It is explicitly documented as unsuitable for logical cache identity.

Every adoption constructs fresh paragraph and main-admission caches. Clones
share those caches only within their captured owner; the effective topology must
never be replaced while retaining a prior owner's caches. Tests warm earlier
indexes and verify different later paragraph counts without changing old views.

Managed views retain `PreparedPartData` and index admissions. Unmanaged visible
XML currently makes an explicit bounded/fallible payload copy and retains the
original data handle as well. This is not an allocation or throughput improvement
claim. Both paths retain all donor contexts: new OPC `check_execution` methods
check execution without re-observing physical source bytes. DOCX semantic queries
observe donor cancellation while pinned views remain valid after source changes.
Final-owner drop releases the retained donor reservations.

## No stale physical fallback

The old raw `package` field is renamed `physical`. Legacy services obtain it
through checked physical helpers that refuse any effective candidate, including
an unchanged candidate. This covers variables, sections, metadata, embedded
objects, ink, hyperlinks, story/glossary text, SVG, paragraph copy/removal, tail
append/stream and their specialized publishers/inverses. Raw main-document
snapshots/edits also refuse effective owners until semantic commits migrate.

`source_backed_package()` now returns `Result<&SourceBackedPackage>` and refuses
effective owners. Its unified unsign consumer propagates that refusal. The
existing unified owner finish fence calls `source_version`, which now checks the
effective donor boundary. Remaining direct physical accesses are explicit
no-effective branches, physical identity checks and cache/ingress diagnostics.

## Verification

- Full-feature DOCX unit/integration tests: 2,003 passed; one existing ignored.
- Final source-backed semantic target after fence/pressure-path refinements:
  15 passed, including five effective-owner tests.
- Focused OPC retained-topology/shared-clone targets: 55 passed.
- Warning-denied DOCX/OPC doctests: 93 passed, 31 existing ignored.
- Unified operations library: 388 passed.
- Warning-denied DOCX/OPC/operations library Clippy: passed; final DOCX/OPC
  rerun after source-fence and pressure-path changes also passed.
- Full-feature facade and Python library compilation: passed.
- Glossary rustdoc: three profiles, 16 banned names, zero complete families;
  passed.
- Crate boundaries: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Formatting and diff whitespace checks: passed.

Tests cover publication/reopen, owning conversion, root redirection, removal,
old-view survival, foreign-lineage refusal, unmanaged and managed caches,
managed donor cancellation/release, donor staleness, and refusal before sink
output. The full existing suite exercises all migrated physical helper routes.
Temporary SVG fixtures are the two files with sizes/hashes recorded in 0019.

Review identified the donor-context issue; the context-only OPC guard and the
pinned-view regression address it. Cache isolation is per-owner invalidation,
not a claim that physical versions identify logical candidates.

## Remaining work

Canonical source-backed package-owning Edit/Commit and immutable Patch apply
still need integration with effective source XML revisions. The current raw XML
transaction API and specialized operations refuse effective owners; these
refusals are transitional, not completion of the requested CRUD surface.
Metadata/embedded/specialized reads must migrate to the common effective adapter.
The unmanaged visible-XML copy and inline retained-handle size need performance
measurement and refinement. All other format, Python/CLI, glossary-family,
performance and acceptance requirements of the original goal remain active.

## Cleanup

Removed temporary SVG fixture copies, workspace build output, generated root
Cargo.lock and check logs. No new worktree was created. The base checkout,
iWork files and pre-existing worktrees were preserved.
