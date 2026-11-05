# 0038 — Owning DOCX OpenOptions ingress

The owning DOCX Package now accepts the crate-root OpenOptions envelope for
configured path, reader and byte-slice ingress. Consuming source_limits and
feature-gated encryption_limits setters are infallible; passwords stay explicit
arguments and are never retained in options. Defaults preserve the previous
finite OPC and encryption policies.

- open_with_options and from_reader_with_options replace their *_with_limits
  counterparts. from_bytes and from_bytes_with_options add the canonical
  borrowed-slice entry point; they use bounded reader ingestion.
- open_with_password_and_options and from_reader_with_password_and_options
  replace both the ad-hoc *_with and *_with_password_and_limits methods.
  The corresponding byte-slice password methods are also available.
- from_decrypted accepts OpenOptions and retains the original encryption mode.
  It applies the selected clear-archive limits without repeating decryption.
- Ordinary path/reader/password constructors delegate with default options.
  No retired constructor aliases remain on the owning Package.
- Unified decryption adoption, style-effect callers, encryption tests and the
  facade README use the new envelope. Positional-source and cache/context
  constructors on source_backed::Package remain a pending migration.

SourceLimits is already validated by its builder. The existing crypto Limits
owner exposes raw fields and validates them at password ingress. OpenOptions
preserves that behavior; encryption_limits has no effect on plain ingress or
already-decrypted adoption. This is explicit in its rustdoc and tested with an
invalid crypto policy. A validated crypto policy type is still needed to finish
the ADR's prevalidated configuration contract. No new semantic ceiling or
execution-budget support is claimed for the owning DOCX envelope.

## Verification

The exact/one-over input tests now cover path, sequential-reader and byte-slice
routes. Byte slices also reject an excessive ZIP member count. Password tests
exercise all three routes with independently exact outer and clear-input
ceilings, and independently reject either exceeded bound. They check retained
encryption mode and refusal to publish plaintext implicitly. Decrypted adoption
checks the clear limit and retained mode. Existing encryption tests still cover
wrong passwords, explicit plaintext publication and re-encryption.

- Full-feature DOCX/operations unit and integration suites: 2,410 passed,
  one existing ignored (2,022 DOCX and 388 operations).
- Minimal-feature DOCX ingress: two passed; the final full-feature ingress
  suite separately passed all five tests.
- Facade read-limit and catalog-detection suites: 17 passed.
- Warning-denied DOCX doctests: 85 passed, 31 existing ignored, including
  the new OpenOptions example.
- Warning-denied DOCX/operations library Clippy: passed.
- Workspace all-feature library compilation, including facade/Python: passed.
- Glossary: 11 profiles, 104 banned-name rules, zero complete families; passed.
  Its 21 checker tests pass. Constructor bans are scoped to the owning Package.
- Crate boundaries: 66 packages, 267 internal dependencies, 11 explicit existing
  debt items; passed.
- Independent read-only review covered options, encryption and caller migration.
  Findings about documentation and test numeric types were corrected.
- Changed Rust formatting and diff whitespace checks: passed.
- Base remains 32ad64e320 and is an ancestor of this branch.

No new parser or wire format was introduced. Existing DOCX parsers and encryption
owners perform ingress; no fuzz campaign, benchmark or Python runtime coverage
is claimed for this API-shape increment.

## Remaining work

Lazy DOCX cache/read-ahead/execution options must join the same envelope without
silently dropping resource policies. Semantic limits, validated crypto limits,
specialized transactions, other format envelopes, Python/CLI, parity and full
acceptance remain open. This is an ingress refactor, not a completed unified
milestone or a performance claim. No OfficeCLI property coverage changes.

## Cleanup

Temporary fixtures are the six base-checkout copies documented with hashes in
record 0035. Verification artifacts, generated root lockfile and build output
were removed before commit. No worktree was created; the base, iWork ownership
and pre-existing worktrees are preserved.
