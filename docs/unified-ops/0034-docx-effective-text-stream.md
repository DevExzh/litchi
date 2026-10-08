# 0034 — Text streaming from retained DOCX owners

Source-backed `Package::write_text_to()` now resolves the effective main part,
including replacement bytes and root-relationship redirection. Canonical
package-owning commits can stream their current text without serializing or
reopening a ZIP. Older owners continue to stream their original text.

An internal inline payload enum retains either physical PartData or effective
PreparedPartData through parsing, sink writes and the final source fence. The
reader checks declared size before payload capture and keeps the existing clean
paragraph-index eviction retry. Only the selected main payload is decoded;
there is no full-package materialization or accumulated document-text buffer.
The semantic parser retains one bounded paragraph. Unmanaged markup compatibility
processing can still allocate a bounded visible-XML buffer; managed base owners
retain their existing admission and compatibility policy. An unmanaged base
with a managed donor follows the existing effective Document read policy while
retaining and checking the donor context.

The sink now checks the logical owner, so retained donor cancellation and source
changes are checked around underlying writes, at parser boundaries and on error
returns. Source failures take precedence over simultaneous sink failures, with
actual accepted-byte progress preserved. Existing repair-report behavior stays
in place, including reporting repairs made before a failed output call.

## Verification

Four new integration tests cover canonical managed/unmanaged commits, old-owner
stability, partial sink failure, a redirected managed donor main part,
cancellation during output with exact progress, oversized declared XML with no
payload reads, and effective markup compatibility with a simultaneous source
change and sink failure. The prior effective-owner refusal test now checks the
newly supported text output.

- Full-feature DOCX/operations unit and integration tests: 2,407 passed,
  one existing ignored (2,019 DOCX and 388 operations).
- Warning-denied DOCX doctests: 84 passed, 31 existing ignored.
- Warning-denied DOCX/operations library Clippy: passed.
- Full-feature facade and Python library compilation: passed.
- Glossary rustdoc: three profiles, 22 banned names, zero complete families;
  passed.
- Crate boundaries: 66 packages, 267 internal dependency declarations and
  11 existing explicit debt items; passed.
- Formatting and diff whitespace checks: passed.
- Independent read-only review found no blocking lifetime or donor-fence issue.
- Rewritten base remains 32ad64e320 and is an ancestor of this branch.

## Remaining work

This closes one read route on effective source owners. Embedded objects, ink,
variables, layout and other specialized reads and transaction families remain
open. Caller-selected semantic limits, other formats, Python/CLI, performance
and full acceptance remain part of the active goal. No complete glossary family
or measured performance improvement is claimed.

## Cleanup

Removed the two temporary SVG fixtures recorded in 0019, workspace build output,
generated root Cargo.lock and temporary check logs. No new worktree was created;
the base checkout and pre-existing worktrees were preserved.
