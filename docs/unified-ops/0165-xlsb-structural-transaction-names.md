# XLSB structural transaction names

The native workbook-structure transaction family now lives in
`litchi_xlsb::cell_values::structure`. Its public names follow ADR 0034:

| Previous name in `cell_values` | Name in `cell_values::structure` |
|---|---|
| `WorkbookEdit` | `Edit` |
| `WorkbookCommit` | `Commit` |
| `WorkbookPatch` | `Patch` |
| `WorkbookHistory` | `History` |
| `WorkbookMergeConflict` | `MergeConflict` |
| `WorkbookMergeOutcome` | `MergeOutcome` |
| `AuthoredStyle` | `AuthoredStyle` |

There are no compatibility aliases. The worksheet-only transaction family
remains in `cell_values`; module qualification distinguishes the two scopes.
Workbook entry points, internal consumers and integration tests use the new
paths. A runnable module example covers edit, commit, durable patch transport,
application and history undo. Both XLSB glossary profiles ban all six old
prefixed names and the former parent-level `AuthoredStyle` export.

This is a naming migration, not completion of the native transaction contract.
`Workbook` and `WorkbookWriter` still need the owner/builder migration; the
structural commit still needs the canonical result/diagnostics/error contract
and immutable publication shape. The implementation retains whole-workbook
before/after images and replay costs. Changed-closure performance and unified
XLSB writes remain prerequisites, as do their Python coverage and paired native
performance evidence. The glossary therefore does not yet register this as a
complete transaction family.

## Verification

- Default-feature native tests and doctests: 1,066 passed (11 doctests ignored).
- Native `sign,vba-inspection` tests and doctests: 1,077 passed.
- XLSB-only unified facade tests and doctests: 489 passed.
- Focused structural/drawing integration tests: 13 passed.
- Native full-feature and XLSB-only facade all-target Clippy passed with warnings denied.
- Minimal and full-feature rustdoc JSON builds passed with warnings denied;
  each glossary profile checked 17 banned names.
- Glossary checker unit tests: 21 passed. OfficeCLI delta regeneration check passed.
- Crate boundaries passed for 66 packages and 268 internal dependency declarations.
- Scoped formatting and whitespace checks passed.

Tests used `CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0
CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true` and a temporary target directory.
The full-feature checks used a temporary copy of the workspace lockfile with
`cc` resolved offline to 1.6.0 through Cargo's `-Z lockfile-path` option; the
inherited workspace lockfile was not modified.

Regeneration commands are `cargo test -p litchi-xlsb`, the same command with
`--features sign,vba-inspection`, and `cargo test -p litchi-ops
--no-default-features --features xlsb`. Run Clippy for those latter two feature
sets with `--all-targets -- -D warnings`. Generate rustdoc JSON for XLSB with
`--no-default-features`, then with `--features sign,vba-inspection`, and run
`tools/check_api_glossary.py` using profiles `xlsb-minimal` and `xlsb-full`.

Comparing the previous `cell_values/root.rs` implementation against
`structure.rs`, after replacing the six identifiers and normalizing whitespace,
produced identical implementation bodies. The only added code is the module
example. Patch wire magic, algorithms, limits and signature policies are
unchanged. No runtime performance improvement or complete performance gate is
claimed by this wave.

Historical report receipts and their frozen reproduction scripts retain the
source paths from their recorded revisions, including `cell_values/root.rs`.
They are not current API examples. Active Rust consumers have migrated; no
Python or performance-harness consumer imports the removed names.

## Remaining work

The full goal remains open. Next XLSB work must resolve the native owner and
commit contract and whole-image staging costs before exposing unified writes.
Tier 2/Tier 3 coverage and all-format Python, fuzz and performance acceptance
remain outstanding.

Temporary targets, the copied lockfile, generated rustdoc JSON, scripts and raw
logs were removed after verification. No worktrees or borrowed fixtures were
created. The inherited workspace lockfile is unchanged, and the rewritten base
remains an ancestor of this branch.
