# 0031 — Source-backed DOCX package-owning Edit and Commit

`source_backed::Package::edit()` now returns
`document::Edit<source_backed::Package>`. Its commit validates the staged patch
through immutable `Package::apply` and returns
`document::Commit<source_backed::Package>`. The commit exposes `package`,
`into_package`, `patch`, `diagnostics` and `(Package, Patch)` through `into_parts`.
It retains the effective owner without writing or reopening a ZIP. Protection,
signature, source and execution failures are now commit-time failures for this
ordinary entry point.

`document::Edit<P = crate::Package>` shares staged methods across both owner
representations. Two specialized commit implementations retain each owner's
publication policy without exposing a public owner trait. Fields and construction
remain private, and Debug never formats the generic owner or document content.
The default generic parameter preserves eager type annotations.

The unified DOCX editor builder now retains the canonical source edit. Its signed
no-op test publishes through the resulting package. Legacy stream/history
publishers still take explicit `document::snapshot::Commit`; their tests and
performance harnesses start with `document_snapshot()?.edit()`. That raw family
remains a distinct type, not an alias for an ordinary package transaction. Tests
of raw publication failures continue to exercise that boundary.

## Preserve existing unmanaged XML compatibility

Moving the MCE witness matrix to canonical commits exposed a gap in record 0030:
source splice proofs reject escaped namespace declarations that the unmanaged
full-document publisher already supports. Refusing those valid producer files
would regress the leaf surface.

OPC `PreparedTopology::apply_xml_bytes` supplies the unmanaged full-document
route. It checks effective bytes, requires a physical XML Part, runs the existing
source/replacement publication audit, and rechecks source fences. Changed
candidates reuse the same flat overlay updater, aggregate limits and exact
inverse handling as typed XML proofs, but carry no positional splice proof.
Managed bases/donors are refused by this raw-byte route; managed canonical edits
retain the existing SourceXmlPart proof and admission route. Added Parts remain
outside this operation. Format semantics and physical identity are checked by
DOCX apply before invoking the low-level byte primitive. The low-level primitive
audits XML even for unchanged bytes; owner-level exact no-ops bypass it.

## Verification

- Full-feature DOCX/OPC/operations unit and integration tests: 3,491 passed,
  two existing ignored (2,008 DOCX, 1,095 OPC, 388 operations).
- Warning-denied DOCX/OPC doctests: 94 passed, 31 existing ignored.
- Warning-denied DOCX/OPC/operations library Clippy: passed.
- Full-feature facade and Python library compilation: passed.
- Standalone performance harness compilation, including final rerun: passed.
- Glossary rustdoc: three profiles, 16 banned names, zero complete families;
  passed.
- Crate boundaries: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Formatting of changed crates/files and diff whitespace checks: passed.

Tests cover both generic owner shapes, managed/unmanaged chains, source-owner
drop, independent staged clones, rollback, package accessors, stale-source
commit refusal, signed exact no-ops and changed refusal, effective protection
policy and graph preservation. The canonical MCE matrix keeps escaped namespace
producer spelling. OPC tests cover full-byte chaining, stale bytes, malformed
XML, managed refusal and exact original-archive restoration.

## Remaining work

This completes the primary source package-owning Edit/Commit step, not W2 or the
unified goal. Specialized effective reads and transactions, history ownership,
legacy stream doors, LimitKind/error vocabulary, full family registration,
mutable-writer removal, remaining formats, Python/CLI, performance and corpus
acceptance remain required. Raw snapshots with independent managed donors still
refuse; unmanaged snapshots still copy/scan. No speed or allocation improvement,
paired benchmark result, or complete transaction family is claimed.

## Cleanup

Removed temporary SVG fixtures (the same files recorded in 0019), workspace and
standalone harness build output, generated root Cargo.lock and check logs.
Preserved tracked standalone lockfiles, base checkout, iWork files and existing
worktrees; no new worktree was created.
