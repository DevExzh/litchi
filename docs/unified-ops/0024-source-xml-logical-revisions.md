# 0024 — Revision-bound source XML edits

Date: 2026-10-08. Status: per-part revision foundation; source owner migration open.

## Change

`SourceXmlPart::begin_revision()` creates a fresh logical edit base over the
current validated payload. It shares the payload and admissions without ZIP
publication/reopen, XML rescanning or work charges. Independent calls produce
distinct identities even over the same payload; clones share one identity.
Checked ranges bind to this identity, physical source lineage/version, Part URI
and content type. A stale or sibling proof cannot authorize an insertion even
when its expected range is empty or its bytes equal the new base.

Splice publication assembles from the logical edit base and validates the
result through the existing bounded XML validator. A changed result clears the
revision, requiring an explicit new revision before further editing. A no-op
retains its revision. The physical `original` PartData remains unchanged and
continues to authorize replacement against the original archive; source checks
and cancellation still apply. A physical PartView hint does not accept logical
revision tokens, instead capturing the physical source afresh.

Each revision admits its shared holder before allocation and retains the
current payload reservation and original PartData. Range proofs retain the
revision and source metadata admission, including content-type storage.
Dropping the previous edit does not prematurely release a surviving proof's
reservations. Revisions do not retain their parents: a linear edit chain keeps
the physical original and current payload, plus any explicitly retained old
owners/proofs. Existing source token/range size-based admissions include the
new fields. Debug reports revision presence without exposing XML content.

## Verification

New tests cover three chained edits without archive reopening, candidate
readback and final publication, exact-artifact inverse publication, refusal of
foreign physical destinations, sibling/stale/physical proof separation, clone
proof reuse, physical hint fallback, original source changes, content-type
mismatch, exhausted memory, cancellation, 32 bounded sequential revisions and
final proof-owner reservation release. A runnable rustdoc example edits twice.

- Full-feature OPC unit/integration tests: 1,071 passed; one existing ignored.
- Final focused XML publication target, including added inverse assertion:
  19 passed (included in the preceding target count).
- DOCX managed document/paragraph publication targets: 52 passed.
- XLSX form-control scalar lifecycle: 32 passed.
- OPC warning-denied library Clippy: passed.
- OPC warning-denied doctests: eight passed after correcting the example's
  error return type to accommodate PackURI construction.
- Full-feature facade and Python library compilation: passed.
- Glossary rustdoc gate: three profiles, 14 banned names, zero complete
  transaction families; passed.
- Crate-boundary gate: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Focused formatting and `git diff --check`: passed.

Independent read-only review found no correctness or accounting blocker.
Sharing and reservation checks establish ownership behavior, not performance
budgets; paired timing/allocation evidence remains required for the owner wave.

## Remaining work

This is per-part XML chaining, not package-wide revision admission. A retained
package owner must read from its effective catalog and separately compare the
logical revision that an edit targeted. It must not substitute logical bytes
for `check_for_replacement`'s original physical authority. PreparedTopology
still lacks application against a new effective revision, and source DOCX
`Package::edit` still uses the transitional snapshot family from record 0022.
No complete transaction family or W2 completion is claimed.

That owner integration should expose a crate-private logical-identity check
rather than compare payload pointers: independent sibling revisions may share
one payload. It should also enforce consistency between a token's retained
revision and current payload wherever new internal owner code manipulates
those fields. Public APIs cannot construct an inconsistent token today.

## Cleanup

Removed generated build output, root Cargo.lock and this turn's temporary
check logs after recording results. No new worktree or fixture copy was created.
The base checkout, iWork sources, existing worktrees and tracked standalone
lockfiles were left unchanged.
