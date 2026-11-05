# 0043 — typed owning DOCX publication failures

## Change

The owning DOCX publisher previously formatted every OPC publication error into
Error::Io with ErrorKind::Other. That discarded structured causes, including
IncompleteOutput's accepted-byte count, and prevented unified classification
from seeing OPC limit/capability failures.

write_plain now propagates OPC errors through the existing Error::Opc variant.
Ordinary write_to and explicit to_plain_stream share this path, as do staged
file publications. The package rollback guard is unchanged. Atomic save still
preserves the destination before replacement; a stream's accepted-byte count
refers to the stream sink, including a staging sink when used by atomic save.

This change preserves existing typed sources; it does not normalize every
legacy DOCX error or change errors constructed inside low-level ZIP codecs.

## Verification

The custom-property failure test now requires an IncompleteOutput error with
exactly one accepted byte, then proves rollback and successful property retry.
A new exact-source regression covers zero and 17 accepted bytes, preserves the
underlying OPC IoError, and verifies retry emits the original archive exactly.
These assertions fail with the former string conversion.

- Focused custom-property/publication suite: nine passed.
- Full-feature DOCX/operations unit and integration suites: 2,431 passed,
  one existing ignored (2,042 DOCX, 389 operations).
- Warning-denied DOCX doctests: 85 passed, 31 existing ignored.
- Warning-denied DOCX/operations library Clippy: passed.
- Changed Rust formatting and diff whitespace checks: passed.
- Rewritten base remains 32ad64e320 and is an ancestor of this branch.

Reproduce with `cargo test -p litchi-docx -p litchi-ops --all-features --lib --tests`,
`RUSTDOCFLAGS='-D warnings' cargo test -p litchi-docx --all-features --doc`, and
`cargo clippy -p litchi-docx -p litchi-ops --all-features --lib -- -D warnings`.

## Remaining work and cleanup

Owner consolidation, immutable owning publication, the remaining API families,
all-format operations and final acceptance remain active. No performance gain
or complete milestone is claimed.

The six fixture copies documented in 0035, generated root lockfile, build output
and session logs are removed before commit. No worktree was created. Base,
iWork and pre-existing worktrees remain untouched.
