# 0044 — owning DOCX bytes and encryption policy

## Change

Owning DOCX Package now provides to_bytes and explicit to_plain_bytes. Both use
its existing publisher, including pending-edit materialization, property staging,
signature checks and source-preserving no-op publication. to_bytes checks the
ordinary-output policy before serialization. to_plain_bytes explicitly permits
plaintext output while retaining the package's encryption provenance.

The private provenance field now uses the shared OOXML PackageEncryption owner,
already used by other OOXML formats. Open/decrypted adoption, clone and successful
encrypted saves retain their modes. Ordinary save, write_to and to_bytes refuse
an encrypted source with Error::EncryptionPolicy and the shared typed cause.
Re-encryption without a retained mode returns the same error family with
NoRetainedMode. Operations classify these refusals as UnsupportedCapability and
retain the typed DOCX source, matching the previous generic refusal's class.
The DOCX encryption feature explicitly enables the shared policy feature.

Encrypted byte production uses to_plain_bytes internally before the explicitly
requested encryption transform. Explicit plaintext output does not turn later
ordinary output into an authorized downgrade. Raw OPC/legacy mutation refusals
remain separate UnsafeEdit errors; this is not a claim that every encryption or
mutation error has been normalized.

The new Vec-returning methods belong to the owning package, which still refuses
managed execution policies at ingress. Managed lazy to_bytes and its returned
buffer accounting remain required work. No managed buffer charge escapes into
an uncharged Vec through this change. Owning publication still takes &mut self.

## Verification

New tests cover pending document/custom-property edits in canonical bytes,
byte-exact reopening and plaintext output, both Standard and Agile encryption,
ordinary output refusal before sink writes, retained provenance after explicit
plaintext output and cloning, missing-mode refusal, and unified classification.
Existing signature tests now exercise to_bytes for signed byte-exact no-ops;
changed-property refusals exercise both byte methods as well as streaming.

- Focused encryption, custom-property and signature suites: 27 passed.
- Full-feature DOCX/operations unit and integration suites: 2,435 passed,
  one existing ignored (2,045 DOCX, 390 operations).
- Minimal-feature custom-property/byte suite: ten passed.
- Encryption-only feature suite: eight passed.
- Warning-denied DOCX doctests: 85 passed, 31 existing ignored.
- Warning-denied DOCX/operations library Clippy: passed.
- Glossary: 11 profiles, 144 banned names, zero complete families; passed.
- Changed Rust formatting and diff whitespace checks: passed.
- Rewritten base remains 32ad64e320 and is an ancestor of this branch.

- Workspace all-feature library compilation, including facade/Python: passed.

- Crate boundaries: 66 packages, 267 internal dependencies and 11 existing
  explicit debt items; passed.

## Remaining work and cleanup

Managed lazy byte output, immutable owning publication, owner consolidation,
semantic/crypto limit validation, other API families and the full all-format
unified goal remain active. No complete milestone or performance gain is claimed.

The six borrowed fixture copies listed in 0035, generated root lockfile, build
output and session logs are removed before commit. No temporary worktree was
created; base, iWork and pre-existing worktrees remain untouched.
