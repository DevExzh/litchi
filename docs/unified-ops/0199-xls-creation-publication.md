# 0199 — XLS minimal creation and unified publication

`Package::new()` produces a strict BIFF8 workbook with one empty `Sheet1`.
It admits the authored bytes through the native owner, so subsequent reads,
primary transactions, inverse patches and publication use the existing package
contracts. The native structural engine now creates the first row block in an
empty indexed worksheet. Admission requires empty INDEX/DIMENSIONS extents,
no block pointers and no orphaned cell/formula records; existing row-block
validation remains in force. Regeneration updates later sheet offsets and the
inverse patch restores the exact empty source. This path still serializes and
parses the fresh image; no performance or zero-reparse claim is made.

Unified `create(Xls)` adopts that owner. Template flavor, OOXML Strict and initial
metadata requests return typed refusals before authoring. The capability report
advertises creation and publication independently of the pending unified XLS
node adapter, and records those three creation-option gaps. Native workbook
reads remain available through the existing facade projection.

Unified memory, stream and atomic file publication delegate to the retained XLS
owner. Path opens retain their origin for `save()`. Atomic publication honors
explicit overwrite and durability policies, validates the destination and checks
source identity before staging and before replacement. Signed sources publish
exactly unless the caller explicitly removes the signature. Unlocked encrypted
owners refuse ordinary output and advertise that refusal to avoid plaintext
publication. Creation has no origin and requires `save_as` for path output.

## Creation option delta

| Option | Current behavior | Follow-up |
|---|---|---|
| Template | Typed unsupported capability | Native XLS template creation |
| Strict conformance | Typed unsupported capability | Use Transitional/default for binary XLS; OOXML Strict is inapplicable |
| Initial metadata | Typed unsupported capability | Native metadata authoring integration |

## Verification

The combined XLS and XLS/encryption facade suite passes 2,229 tests and
doctests, with one ignored. After the empty-table hardening, all 20 native owner
tests pass, including first-cell insertion, two-sheet offset preservation and
exact inverse replay. The unified creator tests cover lazy/eager path origins,
byte/save/reopen, overwrite refusal, changed-source rollback, signature
preservation/removal and encrypted output refusal. A capability golden records
creation/publication support and option gaps without claiming node CRUD.

Warning-denied all-target Clippy and documentation pass; a no-default-feature
facade check and an XLS-only check without encryption also pass. The glossary
audit checks 18 banned names and one transaction family. Independent review
found no remaining blockers after empty-table admission hardening. The boundary audit passes for 66 packages and 268 internal dependencies,
with 11 documented debt items. Builds use an external lockfile and target; repository lockfiles remain
unchanged. No performance benchmark or acceptance claim is made.

## Remaining goal scope

The existing option-bearing `Writer` still needs the canonical Builder migration,
including encrypted publication consumers. Unified XLS node reads, CRUD, batches,
history and full-sheet selectors remain pending. The 4,096-operation primary edit
limit and performance qualification remain open, as do the broader formats,
Python/CLI and full-goal acceptance requirements.
