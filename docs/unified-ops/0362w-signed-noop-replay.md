# 0362w — signed exact no-op replay through `File`

- Date: 2026-10-22.
- Status: signed replay and logical setter/cancellation handling are implemented
  on `uo2/signed-replay`; coordinator integration and the workspace gate remain
  pending. The focused validation build is intentionally deferred while the
  coordinator holds new targets for disk pressure.
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

The authoring matrix now retains the original signed bytes for supported
same-value setters and for edits whose final plan cancels its intermediate
changes. DOCX, PPTX, XLSX, XLSB, XLS and the legacy PPT leaf compare the final
native values with the values read from the original signed owner before
collapsing a native rewrite. A final real change, an absent cell populated with
the native default, a formula/error cell without a complete scalar proof, or an
invalid intermediate operation remains a signed refusal. The empty closure
remains format-neutral. No setter is treated as a no-op by skipping
target/property admission, and no unsigned candidate is used to restore a
signature.

The tests cover:

- exact no-op replay through both a signed leaf and `File::Signed`, preserving
  the fixture bytes and signature state;
- supported same-value setters through the leaf, outer transaction and
  one-shot facade, with changed setter refusals;
- cancellation followed by a real change, which must remain refused;
- an absent spreadsheet cell assigned the native default, which must remain a
  signed change rather than matching a missing original;
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

The earlier signed replay slice was validated on the capped one-job target
with no external fixtures. This follow-up's focused build and tests are still
pending the coordinator's disk-pressure hold:

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

No new build or native timing/performance qualification is claimed in this
record.
