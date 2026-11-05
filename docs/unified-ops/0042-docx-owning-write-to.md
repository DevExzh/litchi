# 0042 — owning DOCX canonical stream output

## Change

Owning DOCX Package now exposes write_to, matching the lazy owner and ADR 0034.
The old to_stream method is deleted without a compatibility alias. The existing
write_plain implementation is retained: signature policy, property staging,
source-preserving no-ops, sink behavior and retry state use the same code path.
Encrypted ordinary output still refuses; its operation name now says write_to.
Explicit plaintext output remains separately selected by to_plain_stream.

Both DOCX glossary profiles ban to_stream on the owning Package in addition to
the lazy Package. Unit/integration tests, operations signature tests, benchmark
consumers and the fuzz smoke fixture producer use the new method. Low-level OPC
calls keep their existing names. Historical captures and result snapshots were
not rewritten.

The owning method still requires &mut self for its existing staging behavior.
This record does not declare immutable owning publication, owner consolidation,
specialized output methods, budgeted lazy to_bytes or a transaction family done.

## Verification

The existing suites exercise signature-preserving no-ops, signed mutation
refusal, encryption refusal, custom-property staging, retry after sink failure,
source limits and immutable owner isolation through the renamed entry point.
No new tests mirror this name-only implementation change.

- Full-feature DOCX/operations unit and integration suites: 2,430 passed,
  one existing ignored (2,041 DOCX, 389 operations).
- Workspace all-feature library compilation, including facade/Python: passed.
- Active performance harness: all benchmark binaries compile.
- Glossary gate: 11 profiles, 144 banned names, zero complete families; passed.
  All 21 checker unit tests pass.
- Changed Rust formatting and diff whitespace checks: passed.
- Rewritten base remains 32ad64e320 and is an ancestor of this branch.

- Warning-denied DOCX doctests: 85 passed, 31 existing ignored.
- Warning-denied DOCX/operations library Clippy: passed.

- Modified sequential-ingress fuzz smoke binary: compile check passed.

- Crate boundaries: 66 packages, 267 internal dependencies and 11 existing
  explicit debt items; passed.

## Remaining work and cleanup

The complete unified-operations goal remains active, including owner/API
consolidation, broader glossary families, all-format read/write, Python/CLI and
acceptance evidence. No performance gain or sustained fuzz campaign is claimed.

Temporary fixture copies are exactly the six borrowed files listed in 0035.
Those copies, generated root/DOCX-fuzz lockfiles, build output and session logs
are removed before commit. No temporary worktree was created; the base, iWork
and existing worktrees remain untouched.
