# 0022 — Eager DOCX package-bound Edit and Commit

Date: 2026-10-08. Status: eager primary commit migrated; full W2 remains open.

## Change

Canonical `document::Edit` now retains the complete eager `Package` from which
it started. It delegates staged mutations to the existing XML engine, then
`commit()` validates and publishes through immutable `Package::apply`. The
result is `document::Commit` with `package()`, `patch()`, `diagnostics()`,
`into_package()` and `into_parts() -> (Package, Patch)`. There is no optional
owner or generic snapshot-as-package parameter. Dropping the original package
or changing its legacy writer after starting an edit does not change the
captured transaction source. A failed commit leaves the original unchanged.

The public eager `publish_document_edit` and `publish_document_commit` methods
are removed. Callers take the committed package; applying a separately prepared
snapshot patch uses `Package::apply`. The retained XML engine lives behind the
explicit `document::snapshot::{Edit, Commit}` public path for standalone XML,
source-backed streams and legacy composition/history. These are distinct types,
not aliases for canonical package commits. Explicit forwarding avoids exposing
raw commit through Deref. Preparing standalone composition remains available
through a snapshot edit, pending package-bound composition ownership.

The canonical Edit/Commit Debug implementations expose no document content;
Commit Debug includes only operation count and change status. Existing raw
snapshot/patch/operation Debug implementations still need the broader glossary
cleanup; this record does not claim all public diagnostics are content-safe.

Updated consumers include the unified DOCX editor's eager/source variants,
standalone benchmark and native-resave tools, managed examples, proof tests,
MCE parity tests and documentation. Eager signature/protection failures now
occur during `Edit::commit`, before it can return a package. Encrypted-source
eager editing remains refused under the existing plaintext publication policy.

## Verification

New tests cover complete related-part retention after source drop, independent
`try_clone` edits, commit cloning, package/patch target agreement, repeated edits
without reopening, inverse restoration, applying the same patch to its source,
no-op graph sharing, failed staging, abandoned-edit rollback, and isolation from
later source writer mutations. Signature tests verify exact signed no-op bytes,
changed commit refusal and explicit unsign; the encryption test verifies edit
refusal retains the encrypted source policy. A compile-fail example distinguishes
standalone XML commits from canonical package commits.

Verification results:

- Full-feature DOCX unit/integration run: 1,992 passed, one existing corpus
  test ignored. Four additional owner/signature tests passed in focused runs
  (1,996 distinct tests total); the updated encryption target also passed.
- Full-feature DOCX library, tests and examples check: passed.
- Warning-denied Clippy for DOCX and litchi-ops: passed.
- Warning-denied DOCX doctests: 83 passed, including the compile-fail boundary;
  31 existing examples ignored.
- `cargo test -p litchi-ops --all-features --lib`: 388 passed.
- `cargo check -p litchi -p litchi-py --all-features --lib`: passed.
- Standalone perf-baseline library and native-resave all-feature checks with
  locked dependencies: passed. The initial DOCX-only native-resave check also
  passed, with its pre-existing unused `std::fs` import warning; the all-feature
  check emitted no warning.
- Glossary rustdoc: 3 profiles, 14 banned-name checks, 0 complete transaction
  families; passed.
- Crate-boundary gate: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Focused formatting and `git diff --check`: passed.

Read-only review confirmed delegation coverage and publication ordering and
requested additional source-mutation/target/security tests; those are included.
It also identified the broader raw Debug and package-lineage follow-ups below.

The full-feature DOCX unit/integration suite passed, followed by focused runs
for the added owner, signature and encryption coverage. Non-ignored SVG cases
used only the two temporary external fixtures identified by path, size and
SHA-256 in record 0019. The base checkout remained read-only.

## Remaining work and limits

Source-backed `Package::edit()` still returns the transitional XML edit; source
stream publication does not yet return a retained overlay owner. That must
converge on the package-bound contract without serializing/reopening between
commits. Legacy history/composition and mutable writer/scoped APIs remain.
`LimitKind`, remaining error vocabulary and all other format waves remain open.
Minimal/full glossary profiles ban the removed publishers and `snapshot` on
canonical Commit, but register no complete transaction family yet.

Ordinary eager patches still authenticate main XML rather than a complete
package lineage; dependency transfer patches additionally check their graph.
Owner-wide lineage must be addressed before claiming full immutable-package
conformance. Equal XML from unrelated packages must not become general package
mutation authority. The existing main-document behavior is unchanged here.

This step adds owner retention and eager publication at commit time. It does
not demonstrate timing or allocation budgets; paired measurements remain
required for the completed ownership wave. No all-format CRUD, M1/M2, Python,
CLI or performance milestone completion is claimed.

## Cleanup

Removed generated build output, root lockfile, temporary external fixture copies
and this turn's check logs after recording results. No new worktree was created;
pre-existing worktrees and tracked standalone lockfiles remain unchanged.
