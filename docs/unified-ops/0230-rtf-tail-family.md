# 0230 — Canonical RTF logical-tail family

The focused module now exports `tail_append::{Edit, Commit, Patch, Diagnostics,
Error, LimitKind}`. The old `TailAppend`-prefixed transaction names are removed
from both the module and crate root, without aliases. Package entry points are
`edit_tail_append` and `edit_tail_append_with_limits`; replay remains
`apply_tail_append` with recipient-policy preservation from record 0229.

Transaction resource refusals now use `Error::Limit { kind, observed, maximum }`.
`LimitKind` distinguishes paragraph/run counts, input/inserted/output/source
bytes, patch bytes/offset, before/after artifact bytes and memory bytes. The
existing count calculations, thresholds and atomic refusal points are retained.
Publication-specific error types and bounds remain a separate pending migration.

Both the retained and source-backed edit paths drop the duplicate
`append_plain_paragraphs`, `append_plain_runs` and four `plan_publication*`
forwarders. Their canonical append and `publication_plan*` methods remain.
Live native tests and benchmark consumers use the new API. Historical evidence
and benchmark measurement-description strings remain unchanged.

## Verification

The complete native suite passes 1,459 tests/doctests. A new doctest covers the
focused Edit, canonical owner entry, commit decomposition, no-op inspection and
package replay. Existing limit tests now match typed dimensions.

Warning-denied native all-target Clippy and rustdoc pass. The glossary verifies
79 banned names and all six members of the focused family; its 21 Python tests
pass. The benchmark library compiles with migrated native calls and facade
dependencies. This scoped native refactor does not rerun the full facade suite.

Validation uses external copied locks and build output. All three repository
lockfile hashes are unchanged. Owned 0230 temporary files are removed after
qualification; no worktree was created and no large evidence is committed.

## Remaining work

Checked scoped limit profiles and options, source/publication family vocabulary,
redaction vocabulary and unified RTF CRUD remain open. Remaining formats, Python,
CLI, corpus and performance acceptance remain required by the active full goal.
No runtime performance acceptance claim is made in this wave.
