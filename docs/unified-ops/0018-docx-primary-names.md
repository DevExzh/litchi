# 0018 — DOCX primary entry and focused error names

Date: 2026-10-08. Status: name migration complete; W2 ownership work open.

## Change and compatibility

Both eager and source-backed DOCX packages now expose `edit()` as the primary
main-document transaction entry. `edit_document()` is removed, not deprecated
or retained as an alias. Scoped entries such as `edit_document_variables()`
remain scoped families and are not renamed by a substring replacement.

`document::Error` replaces `document::TransactionError`, and
`document::Result<T>` replaces `document::TransactionResult<T>`. The canonical
result fixes the error to the focused document error; private helpers that
return `Refusal` or other types explicitly use `std::result::Result`. Error
variants, payloads, conversions and display semantics are retained. Some
internal imports keep a private disambiguating local name beside the package's
separate ingress error/result. There are no public compatibility aliases.

Updated the ops layer, typed error adapter, tests, examples, current leaf
feature matrix, active design reference, standalone performance harness and
native-resave tool together. Historical measurement/decision records retain
names used at their recorded revisions. Existing Python read APIs do not call
these leaf mutation entries and their public signatures are unchanged.

The source-backed failure context now says operation `edit` instead of
`edit_document`; the MCE refusal assertion follows the public entry name.
No snapshot, edit, commit, patch, publication, security or resource-budget
algorithm changes. No new unsafe code, runtime or dependency edge.

## Scope within W2

The present document edit owns a main-part `Snapshot`; its `Commit` still
publishes that document snapshot, while separate package publishers perform
package validation and publication. Renaming the entry does not make that
commit package-owning. The larger ADR 0034 contract remains open:

1. Immutable, shareable `Package` with an owned read-only `Document` projection.
2. Package-bound `Edit` and `Commit` carrying the new package, patch and
   diagnostics, including `into_package()` and the canonical `into_parts()`.
3. Canonical immutable `Package::apply`, replacement of the old publication
   helpers, and removal of the mutable writer path after authoring verbs move
   into edits.
4. Focused family `LimitKind`, limit/options renames, metadata unification,
   and remaining policy/error-shape work.

These are implementation requirements, not deferred alternatives. No complete
transaction family is registered as conformant yet. This step establishes
canonical entry/error spellings used by the following ownership migration.

## Enforcement

`tools/api_glossary.json` adds `docx-minimal` and `docx-full`. The latter enables
`automatic-fonts`, `encryption`, `sign` and `vba-inspection`, covering the full
optional DOCX feature closure. Both ban `edit_document` throughout the DOCX
surface and the prefixed error/result names in the document module. Current
consumer compilation complements the rustdoc gate; the gate itself does not
certify documentation prose or unregistered downstream packages.

The glossary updates only the migrated names. Other format families and the
all-format read/write objective remain incomplete.

## Verification

- `cargo check -p litchi-docx --all-features --lib --tests --examples`: passed.
- DOCX full-feature library tests: 1,264 passed.
- Focused integration suites `mce_admission`, `source_backed`,
  `source_backed_file`, `source_backed_managed_document_edit`, and
  `source_backed_managed_paragraph_batch`: 84 passed (11/16/6/43/8).
  They cover byte-preserving no-ops, source freshness, guarded publication,
  MCE admission/refusal, and managed transaction resource behavior.
- `cargo test -p litchi-ops --all-features --lib`: 388 passed.
- Warning-denied DOCX Clippy: passed.
- Warning-denied DOCX doctests: 80 passed; 31 pre-existing ignored examples
  remain ignored. The new runnable example exercises both package owners and
  the canonical result type.
- Warning-denied rustdoc glossary gate: 3 profiles, 6 scoped rule checks,
  0 complete families; passed.
- `python3 -m unittest tools.test_check_api_glossary`: 19 passed.
- Standalone performance harness `cargo check --tests`: passed.
- Native-resave `cargo check --features docx`: passed, with its pre-existing
  unused `std::fs` import warning. Its initial default-feature check did not
  compile the DOCX branch, so the feature-specific check was added.
- Crate-boundary gate: passed for 66 workspace packages and 267 internal
  dependency declarations; 11 existing explicit debt items remain.
- Focused Cargo formatting and `git diff --check`: passed.

The first compile found two test helpers importing the retired result through
`super::*`; those helpers now use the canonical result. Review also caught
an overly general second type parameter on the new result alias; it was
removed before final verification. The reviewer found no unconverted active
Rust callers or public compatibility aliases.

No paired timing is claimed for this naming-only step. Full milestone
feature combinations, Python/fuzz/parity and performance acceptance remain
required as the ownership and operation implementations land.

## Cleanup

No new worktrees. Removed generated root/standalone build artifacts, the
untracked root lockfile, temporary test logs and this session's Python caches.
Preserved checked-in standalone lockfiles and pre-existing independent
worktrees. No artifacts, binaries or large raw evidence committed.
