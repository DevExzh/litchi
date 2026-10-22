# 0362w — signed exact no-op replay through `File`

- Date: 2026-10-22.
- Status: implementation complete on `uo2/signed-replay`; coordinator
  integration and the workspace gate remain pending.
- Scope: the unified `litchi_ops::transaction::Patch::apply` admission path
  and the signed replay coverage shared by the format matrix.

The coordinator decision changes the facade rule for an exact no-op. A patch
created from an authorized unsigned view may replay onto a signed leaf or the
outer `File::Signed` wrapper when the recipient has the same format and the
same recorded base bytes. The successful result is a clone of the recipient,
so its signature wrapper and publication bytes remain unchanged. This is the
facade counterpart of the existing format-package no-op contract.

The replay proof still authenticates content. It streams both endpoints into
the existing bounded `(u64 length, SHA-256)` fingerprint and compares the
actual recorded publication. An authorized `RemovalAuthorized` endpoint is
temporarily viewed with its original signed disposition for this proof,
because its ordinary output path intentionally removes signature parts. The
proof therefore does not compare unsigned semantic output, accept a target
because its unwrapped format matches, or rely on signature disposition alone.
Even a shared leaf owner takes this fingerprint path when the recipient is
signed.

The source fence is split only for this no-op replay. DOCX, PPTX and XLSB use
freshness checks that bypass their signed refusal while retaining the existing
encryption and owner checks; the other formats reuse their freshness checks.
Locked recipients and changed signed patches continue through the protected
path. A signed recipient with different recorded bytes returns
`PatchConflict` (`invalid_input`), and a changed retained source returns the
format's source error (`io_error`). The edit admission path is unchanged, so
creating a changed edit from a signed `File` still refuses before the closure
runs.

The tests cover:

- exact no-op replay through both a signed leaf and `File::Signed`, preserving
  the fixture bytes and signature state;
- a different signed DOCX recipient returning `invalid_input`;
- a path-backed signed recipient changed after open returning `io_error`;
- the all-feature signed-input matrix replaying an authorized no-op through
  the unified facade for every signed format route;
- locked and unlocked-encrypted recipient protection remaining intact.

Validation on the capped one-job target used no external fixtures:

```text
cargo fmt --all -- --check
cargo check -p litchi-ops --features docx,pptx,xlsx,xlsb
cargo test -p litchi-ops --test docx_transactions --features docx \
  signature_authorization_is_bound_and_inverse_never_restores_signatures
cargo test -p litchi-ops --test docx_transactions --features docx \
  source_guards_survive_signature_removal_promotion_to_eager_owner
cargo test -p litchi-ops --test docx_transactions --features docx,encryption \
  locked_input_refuses_before_closure_execution
cargo test -p litchi-ops --test docx_transactions --features docx,encryption \
  unlocked_encrypted_input_refuses_without_dropping_publication_protection
cargo test -p litchi-ops --lib --all-features \
  every_signed_input_follows_the_explicit_signature_policy
```

All commands passed. No native timing or performance qualification is claimed
in this record.
