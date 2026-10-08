# 0027 — DOCX patch no-op vocabulary and content-free Debug

The primary main-document `Patch` now exposes `is_noop()` instead of `changed()`.
The predicate retains the existing exact source/target comparison; operation
count alone does not determine whether a patch changes the source. There is no
compatibility alias. `Diagnostics::changed()` remains the commit predicate.
Eager and source-backed publishers, tests, examples and performance-harness
consumers use the new spelling. Both DOCX glossary profiles ban the old method.

Derived Debug previously exposed XML bytes, authored text and transfer metadata
through the main-document transaction graph. Snapshot, raw Edit, Patch,
Operation, replacement inputs, paragraph transfers, transfer graphs, transfer
refusals and transaction errors now format structural information only. Raw
Commit inherits that behavior. Prepared-edit identifiers and composition failure
or conflict payloads are also omitted from their local wrappers' Debug output.
Formatting does not compare source XML or enumerate operations and Parts.
Explicit typed payload access remains available.

Two regression tests cover compact and pretty formatting of snapshots, edits,
commits, inverse patches, operations, prepared identifiers, replacement text,
transfer XML/dependencies and error details. They check both text and byte-vector
representations and bound diagnostic size. Existing no-op, inverse and managed
publication tests exercise the unchanged predicate semantics.

## Scope and remaining work

This is a glossary/default-policy slice, not a complete transaction-family wave.
Error Display, common-core re-exported error/conflict types, and other DOCX
transaction families have not been made content-free here. Typed error/limit
vocabulary and the canonical source-backed package-owning Edit/Commit remain
open. The effective-read adapter and logical cache work in record 0026 are still
required. No allocation or throughput improvement is claimed; all-format CRUD,
Python/CLI completion and the goal's performance/acceptance gates remain open.

## Verification

- Full-feature DOCX unit/integration tests: 1,998 passed, one existing ignored.
- Warning-denied DOCX doctests: 83 passed, 31 existing ignored.
- Warning-denied DOCX and operations library Clippy: passed.
- Full-feature DOCX/facade/Python/operations all-target compilation: passed.
- Standalone perf-baseline locked library and native-resave locked all-feature
  compilation: passed. Initial compile checks identified three remaining test
  and harness call sites; these were migrated before successful reruns.
- Glossary rustdoc: three profiles, 16 banned names, zero complete families;
  passed. All 19 glossary checker tests passed.
- Crate boundaries: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- DOCX formatting and diff whitespace checks: passed.
- Read-only review found no Debug disclosure or missed primary-patch consumer.

Temporary external SVG fixtures are the same two fixtures, with sizes and hashes,
recorded in 0019. They were copied read-only from the base checkout.

## Cleanup

Removed the temporary SVG fixtures, generated root lockfile, check logs and
workspace/standalone build output. No new worktree was created. The base
checkout, iWork files, pre-existing worktrees and tracked standalone lockfiles
were preserved.
