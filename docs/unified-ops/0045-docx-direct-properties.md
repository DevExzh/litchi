# 0045 — typed DOCX direct property reads

## Change and provenance

Added local typed property projections for runs, paragraphs, tables, rows and
cells. Optional values distinguish absence from explicit false/zero; invalid
supported property values are listed in `invalid_properties`. These readers do
not resolve style inheritance. Positional `addressable_runs` follows the
OfficeCLI navigation set while the existing `runs` retains its outermost-run
contract. InlineKind classifies inert inline content without rendering it.

The starting leaf source diff was commit
7fa4d17759c7cbda4d7a41b8a650f4f3786da8e6 from the earlier M1 candidate. Only
DOCX source and feature documentation were transplanted and adapted; no branch
ancestry, large evidence or candidate performance claims were imported. The
candidate raw public source-owner relationship accessor was omitted because it
would not account for effective graph edits. Corpus tests use the existing
physical relationship graph only for unchanged source packages.

Retained property, inline and positional-run readers now seed inherited
namespace bindings. Strict and Transitional Word names are accepted, foreign
names remain inert, and new scanners audit lexical/expanded duplicate
attributes and malformed values even on skipped children. DOCTYPE and content
outside the fragment root are refused. New parser errors and touched paragraph,
run, hyperlink, opaque-inline and authored-property Debug output omit authored
content. This does not claim all legacy diagnostics or Debug implementations
have been migrated.

Run readers inspect only their own rPr, excluding revision snapshots and nested
runs. Cell span/merge readers inspect only their own tcPr, preserve inherited
bindings and omit authored values from invalid-value errors. Existing fragment
parser depth/element caps remain; caller-selected semantic limits are still
required work. Namespace discovery for a retained run can scan its owning XML;
no performance improvement is claimed.

## Verification

Regression coverage exercises inherited namespace aliases, foreign names,
explicit false and absent values, nested ownership, revision snapshots,
malformed attributes, DOCTYPE, diagnostic redaction and content-free Debug.
Corpus reads compare supported projections with existing readers.

- Full DOCX/operations all-feature unit/integration tests: 2,493 passed,
  one existing ignored.
- Minimal-feature property suite: 41 passed.
- DOCX/operations doctests: 153 passed, 31 existing ignored; the separate
  warning-denied DOCX doctest run passed 96 tests.
- Warning-denied DOCX/operations library Clippy: passed.
- Workspace all-feature library check, including facade/Python: passed.
- All DOCX fuzz binaries compile; deterministic direct-property smoke: 2,831
  cases, 791 accepted and 2,040 refused. No sustained fuzz campaign is claimed.
- Glossary check: 11 profiles, 144 banned names, zero complete families; passed.
- Boundary check: 66 packages, 267 internal dependencies, 11 debt items; passed.

## Remaining scope and cleanup

This is a read foundation for formatting operations. The M2 formatting engine,
effective source relationships (R8), managed lazy bytes, immutable owning
publication, owner consolidation and full all-format acceptance remain active.
Neither M1 nor M2 is declared complete. No benchmark evidence was imported.

The six borrowed fixture copies listed in 0035, generated lockfiles, build
output and session scratch files are removed before commit. No temporary
worktree was created. Base, iWork and pre-existing worktrees remain untouched.

Rewritten base 32ad64e320 remains an ancestor. Changed Rust formatting and diff
whitespace checks passed; repository artifact restrictions were checked before
commit.
