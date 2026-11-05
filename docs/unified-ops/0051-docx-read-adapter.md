# 0051 — DOCX read adapter integration

## Scope

Integrate the source-only DOCX adapter candidate `6b5e267fb9` over records
0047–0050. No candidate history or large evidence is merged. DOCX now has a
concrete node variant, static query vocabulary, lazy children, direct-property
projection and internal text/annotated/outline/stats rendering. File-level
get/find/query/view dispatch remains the next integration step; other formats
retain their pending node variants. This does not close M1.

Ordinary paths remain positional, including descendants of paragraphs and
content controls that carry native IDs. Explicit raw paths admit paragraph and
content-control IDs; malformed native values are errors, not silent misses.
Native identity fields are omitted from ordinary serialized projections, while
raw query predicates can inspect them. Semantic style and numbering references
retain the candidate's existing property vocabulary pending registry migration.

Reads fence retained source revisions before and after path resolution, each
query yield, child iteration, projection and full document rendering. Errors
fuse iterators. Zero-limit queries read no payload. Typed leaf accessors now
return `Result<Option<_>>` and check the owner; previously extracted leaf values
remain immutable decoded values, as in the direct leaf API. Transaction-position
lookup checks source freshness and preserves the leaf's addressability proof.

Source-backed paragraph inline reads use the effective main-part relationship
table after edits and recheck all retained sources. Managed detached views
continue to use the leaf refusal from record 0049. Invalid direct properties
reported by typed leaf projections cause a content-free corrupt-file error
instead of being silently projected as absent. Native identity, source, and
leaf errors retain typed handling; diagnostics omit authored input.

## Verification

Tests cover both
owning and source-backed owners, effective edits, native raw selection,
positional path round trips, stale cached reads and fused iterators, malformed
properties, corpus reads, query/leaf comparisons, decode counts and both formula
feature variants. These are correctness and I/O-accounting checks, not paired
performance acceptance.

Checks completed:

- `cargo test -p litchi-ops --all-features`: 605 library tests, 17 integration
  tests and 129 doctests passed at the initial integration checkpoint.
- The final DOCX adapter suite: 25 tests passed with all features and with
  `--no-default-features --features docx,test-support`, including effective
  relationship resolution and raw-query-only native identity access.
- `cargo test -p litchi-ops --no-default-features --lib`: 223 passed.
- `cargo test -p litchi-docx --all-features --lib`: 1,397 passed.
- All-feature workspace library check and warning-denied DOCX/ops rustdoc passed.
  Ops doctests were rerun after the typed accessor changes: 129 passed.
- Warning-denied all-target DOCX/ops Clippy passed after the final view-validation
  adjustment.
- API glossary: 11 profiles, 144 banned names, zero transaction families checked.
  Glossary fixture tests: 21 passed. Boundary checker: 66 workspace packages,
  267 internal dependency declarations, 11 existing debt items.
- Final malformed-property regression also checks annotated/outline view refusal.
- Source and snapshot formatting and `git diff --check` passed.

No new Python surface or parser grammar is added here. Pytest, a new adapter
fuzz campaign, generated parity-table closure and paired performance benchmarks
are not claimed by this prerequisite integration; they remain M1/goal gates.
Base `32ad64e3207655b4ba558fc2793d9a7396655cdf` is already an ancestor, so no
additional rebase was necessary.

## Remaining scope

File dispatch, capability/registry tables, all-format writes, Python/CLI,
caller-selected budgets, generated parity deltas and paired performance evidence
remain required by the complete goal. Candidate OfficeCLI-shaped snapshots are
adapted to ADR 0034; this record does not claim complete OfficeCLI parity.

The glossary gate now avoids creating an unused sibling scratch directory when
an explicit target directory is supplied. Its fixture tests use the writable
checkout and clean up afterward. The existing DOCX memo Debug implementation was
moved above its test module to satisfy warning-denied all-target Clippy.

## Hygiene

No temporary worktree was created. Base and iWork sources remain untouched.
Generated build artifacts, lockfiles and scratch logs are removed after checks.
Evidence-size and compiled-binary restrictions apply before committing.

## Commit limitation

The final `git add` attempt failed because the sandbox mounts the worktree Git
metadata read-only: `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`
cannot be created. No commit was made in this turn. Changes remain in the
checkout; imported candidate files were already partially staged before this
restriction, so the index is not the final reviewed result. Staging must be
refreshed before committing once Git metadata is writable.
