# 0039 — DOCX positional OpenOptions

The existing crate-root OpenOptions now also configures DOCX positional and
filesystem ingress. source_backed::Package exposes open/open_with_options and
from_read_at/from_read_at_with_options. Seventeen retired from_path and
combinatorial cache/limits/context constructors are deleted without aliases.
Callers, managed tests, the active performance tools and route descriptions
use the new envelope. Historical evidence sources are unchanged.

OpenOptions adds infallible consuming cache_limits, source_read_policy and
execution_context setters. SourceCacheLimits and SourceReadPolicy are available
at the DOCX root. Defaults preserve finite cache bounds, exact reads and the
absence of an implicit context. Options are Clone, not Copy: clones share the
same budget and cancellation token. Every selected policy reaches the same OPC
admission path as before; deferred semantic reads retain the context. Existing
managed part handles and publication reservations keep their ownership rules.

Configured filesystem ingress checks cancellation before opening the file.
Cancelled preserves the existing typed OPC cancellation variant; other execution
errors retain their typed source. Read-ahead still charges physical overfetch
and retained window memory, and publication disables it through the OPC owner.

## Owning-route boundary

Owning Package payloads cannot currently retain the lazy cache/context policies.
Those routes now reject explicitly supplied lazy settings before input I/O,
including password routes, instead of ignoring the new fields. The typed
UnsupportedOpenPolicy error identifies the first unsupported OpenPolicy;
execution context takes precedence, then cache limits, then source read policy.
Unified classification preserves it as UnsupportedCapability with its leaf
source. from_decrypted also refuses these settings before OPC adoption.
Defaults and source/encryption-only owning options retain their existing behavior.

This refusal is an explicit migration boundary, not a completed capability.
The canonical owner still needs to absorb the owning compatibility layer. Lazy
sequential ingress retains its existing bounded source-limits API; managed
reader ingestion must account for retained input memory before it can join
this envelope. No unaccounted managed reader or silent policy downgrade was
introduced. The crypto raw-policy validation debt from 0038 remains.

## Verification

New tests prove cloned options share budget/cancellation, managed reservations
release after both packages and views drop, owning routes reject each lazy
setting before reader/path/decryption I/O, already-decrypted adoption rejects
those settings, missing-path opens honor prior cancellation, and unified errors
classify the policy refusal. Migrated existing suites cover deferred reads,
exact/one-under budgets, source changes, cache policies, managed edits and
read-ahead publication/release.

- Full-feature DOCX/operations unit and integration suites: 2,414 passed,
  one existing ignored (2,025 DOCX and 389 operations).
- Minimal-feature DOCX owning/managed ingress suites: 11 passed.
- Facade catalog-first detection suite: four passed.
- Warning-denied DOCX doctests: 85 passed, 31 existing ignored.
- Warning-denied DOCX/operations library Clippy: passed.
- Workspace all-feature library compilation, including facade/Python: passed.
- Active perf-baseline bins and all-feature DOCX examples compile.
- Glossary: 11 profiles, 138 banned-name rules, zero complete families; passed.
  Its 21 checker tests also pass. The 17 retired methods are banned in both
  DOCX profiles only at the source_backed::Package scope.
- Crate boundaries: 66 packages, 267 internal dependencies and 11 existing
  explicit debt items; passed.
- Independent read-only review checked options, caller scopes and evidence
  preservation. The low-level OPC donor call and historical snapshots were
  restored after the automated migration; active tool labels were updated.
- Changed Rust formatting and diff whitespace checks: passed.

## Remaining work and cleanup

Sequential-reader options, semantic and crypto validated policies, consolidation
of the canonical owner, specialized DOCX transactions, other formats,
Python/CLI, parity and full acceptance remain active. No performance improvement,
fuzz campaign, Python runtime result or completed unified milestone is claimed.

Temporary test fixtures are the six base-checkout copies listed with hashes in
0035. They, build output, generated root lockfile and temporary logs were removed
before commit. No worktree was created; base, iWork and existing worktrees are
preserved. Rewritten base remains 32ad64e320 and is an ancestor of this branch.
