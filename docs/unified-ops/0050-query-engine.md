# 0050 — Bounded query compilation and streaming evaluation

## Scope

The concrete DOCX adapter depends on query compilation and traversal that were
absent from this branch. Integrate the code-only engine from `516536a2a2`, adapting
it to records 0047–0049 and ADR 0034. No old history or evidence was merged.

The engine binds selector kinds, property tables and coerced comparisons once
into flat plans, and walks nodes iteratively in document order. It supports
combinators, unions, boolean predicates, positional filters, contains/empty,
has/not, worksheet coordinates, and header-row queries. Early limit/first and
error termination stop further traversal. Plans are format-bound and reusable.

Native IDs are excluded from ordinary predicates. `query::raw::query` opts into
native predicates and raw scope paths; a compiled raw plan retains that choice.
Semantic names keep ordinary access. Unknown keys still produce diagnostics
and follow the design's absent-value semantics (D10); this is deliberately not
changed into a strict unknown-property error. Regex remains a typed refusal
pending the design's regex-engine decision.

Typed ASTs receive preflight checks before recursive compilation: parser-shaped
counts/depths, at most 4,096 syntax nodes, and a 16 KiB canonical text bound.
Empty queries/boolean lists, invalid positions and excessive scope paths refuse.
Walks cap active frames at 256 and visits at one million, including sibling-count
passes and nested descendant tests. Caller-selected semantic budgets and
cancellation integration remain follow-up work.

Compiler errors and diagnostic rendering omit authored content. Plan, QueryArg,
Diagnostic and grammar scope Debug are redacted. Explicit syntax/data accessors
retain their documented content; this is not a prohibition on serialization.

## Verification

Tests use fake Word, presentation and workbook
adapters. They cover reusable plans, scopes, boolean and structural semantics,
early termination, typed coercion, header-row binding, unknown-key behavior,
harvested OfficeCLI inputs, malformed typed ASTs, raw identity admission, deep
traversal refusal and redacted diagnostics. Allocation checks reuse the existing
path-test allocator; no new unsafe code is added. Fuzz support compiles plans
and runs bounded walks across three fake families.

- `cargo test -p litchi-core -p litchi-ops --all-features`: 838 unit/integration
  tests and 181 doctests passed.
- `cargo test -p litchi-ops --no-default-features --lib`: 221 passed.
- `cargo test -p litchi-ops --no-default-features --features test-support`:
  unit/integration, allocation checks and doctests passed. Five existing examples
  now guard their format-specific code so disabled formats do not break doctests.
- `cargo test -p litchi-ops --all-features --doc`: passed after those guards.
- Workspace all-feature library check, Clippy with warnings denied, rustdoc with
  warnings denied: passed.
- API glossary: 11 profiles, 144 banned names, zero complete transaction families.
  Crate boundaries: 66 packages, 267 internal dependencies, 11 existing debt items.
- Query smoke: 792 deterministic inputs, 748 compiled family plans; bounded walks
  completed without internal errors. All fuzz binaries compile. This is smoke,
  not a coverage-guided campaign.

Read-only review identified the typed-AST, raw identity, diagnostic content,
worksheet probe accounting and terminal-row overflow fixes. Regression tests
cover them. Base HEAD `32ad64e3207655b4ba558fc2793d9a7396655cdf` remains an ancestor.

## Remaining scope

No production format query adapter or File get/find/query is registered here.
The next integration remains the concrete DOCX adapter, followed by File dispatch
and capability tables. Other-format reads/writes, registry/write dispatch,
Python/CLI, caller budgets and paired leaf-versus-unified performance evidence
remain required by the complete goal. No performance acceptance is claimed.


## Hygiene

Changed Rust files are formatted. Borrowed fixtures, generated root/fuzz
lockfiles, build output and scratch files are removed before commit. No temporary
worktree was created. Base, iWork and pre-existing worktrees remain untouched.
Evidence-size and binary restrictions are checked before commit.
