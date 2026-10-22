# 0362w — signed exact no-op replay through `File`

- Date: 2026-10-22.
- Status: the signed replay and DOCX setter slice are implemented on
  `uo2/signed-replay`; the remaining format setter routes are queued for the
  follow-up repair, and coordinator integration and the workspace gate remain
  pending.
- Scope: the unified `litchi_ops::transaction::Patch::apply` admission path
  and the signed replay coverage shared by the format matrix.

The coordinator decision adds a signed-source path for an exact no-op. A patch
created from a signed leaf or the outer `File::Signed` wrapper may replay onto
an independently opened signed recipient when the recipient has the same
format and the same recorded base bytes. The successful result is a clone of
the recipient, so its signature wrapper and publication bytes remain
unchanged. This is the facade counterpart of the existing format-package
no-op contract. A patch created after `Signed::unsign` remains an unsigned
patch and replays only onto an endpoint whose unsigned publication matches it.

The replay proof still authenticates content. It streams both endpoints into
the existing bounded `(u64 length, SHA-256)` fingerprint and compares the
actual recorded publication, including signature parts when the source is
signed. It never re-marks an unsigned endpoint or compares unsigned semantic
output to a signed publication. `RemovalAuthorized` keeps the original owner
for source freshness, but its ordinary writer intentionally emits the
unsigned publication; an authorized no-op therefore matches an independently
reopened unsigned endpoint and conflicts with the signed bytes. The proof
does not accept a target because its unwrapped format matches or rely on
signature disposition alone.

The source fence is split only for this no-op path. A signed source may enter
the transaction closure to establish an exact no-op, while mutation paths
return the existing protected error and a changed signed candidate is rejected
before publication. DOCX, PPTX and XLSB use freshness checks that bypass their
signed refusal while retaining the existing encryption and owner checks; the
other formats reuse their freshness checks. Locked recipients and changed
signed patches continue through the protected path. A signed recipient with
different recorded bytes returns `PatchConflict` (`invalid_input`), and a
changed retained source returns the format's source error (`io_error`).

The first implementation of this record temporarily re-marked every unsigned
endpoint as `Signed` while fingerprinting. Review rejected that approach: a
`RemovalAuthorized` endpoint's ordinary publication is unsigned, and replay
must not make it match physically different signed bytes merely because it
retains the old owner. The current implementation creates signed no-op
patches from signed `File` endpoints and fingerprints every endpoint in its
actual publication state.

The authoring matrix has one explicit logical setter proof in this slice:
DOCX paragraph text can be set to its current value through the signed leaf,
the outer `File::Signed` transaction, and the one-shot `File::set`; all three
retain the original signed bytes, while a changed value returns
`document_protected`. The empty closure remains format-neutral. The same-value
setter routes for PPTX, XLSX, XLSB and the legacy signed leaves still require
the follow-up repair; their existing editor guards are not accepted as
exceptions to the exact no-op rule. No setter is treated as a no-op by
skipping target/property admission, and no unsigned candidate is used to
restore a signature.

The tests cover:

- exact no-op replay through both a signed leaf and `File::Signed`, preserving
  the fixture bytes and signature state;
- a supported DOCX setter through the leaf, outer transaction and one-shot
  facade, with changed setter refusals;
- a different signed DOCX recipient returning `invalid_input`;
- independently reopened unsigned and `RemovalAuthorized` no-op endpoints
  matching each other while both conflict with a signed recipient;
- an inverse of an authorized changed edit that cannot restore signatures;
- a signed encrypted source refusing publication fingerprinting without
  exposing plaintext;
- a path-backed signed recipient changed after open returning `io_error`;
- the all-feature signed-input matrix checking signed empty no-op replay and
  unsigned authorized no-op negatives for every signed format route;
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
