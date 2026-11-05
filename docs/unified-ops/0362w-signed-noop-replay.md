# 0362w — signed exact no-op replay through `File`

- Date: 2026-10-23.
- Status: signed replay, logical setter/cancellation handling, and the native
  snapshot-budget audit are implemented on `uo2/signed-replay`; coordinator
  integration and the workspace gate remain pending. The focused validation
  build is intentionally deferred while the coordinator holds new targets for
  disk pressure.
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
discarding a private changed candidate. The returned patch is a genuine native
empty patch rooted in the original signed owner; there is no exact-no-op marker
or synthetic signature state. Authored operation counts and the peak staged
snapshot charge remain in commit diagnostics before canonicalization, so limits
cannot be bypassed by dropping the private candidate. XLSX structural,
rich-run and feature routes all pass the caller's snapshot limit into native
commit adaptation. PPTX signed table and notes plans preflight their captured
bodies against the staged budget, and XLSB cell routes preserve the authored
scalar count while using worksheet transfer limits for the canonical patch.
A final real change, an absent cell populated with the native default, a
formula/error cell without a complete scalar proof, or an invalid intermediate
operation remains a signed refusal. The empty closure remains format-neutral.
No setter is treated as a no-op by skipping target/property admission, and no
unsigned candidate is used to restore a signature.

Legacy XLS signed reads use a bounded read-only OLE probe with a signature-only
allowlist: encryption and DRM components remain refused. The probe compares
the original Workbook stream before opening a mutation-capable editor, and
stale-source/patch-conflict admission still wins over signed refusal.

## Independent review follow-up

The follow-up closes the remaining review gaps without widening the public
commit or signature contracts. DOC, XLS, XLSX, PPTX and PPT coalesced editors
receive the transaction's remaining snapshot budget before they capture a
native before image; the PPT direct planner receives the same budget. PPTX
retained-patch sizing is lazy and cached for unsigned commits, and its
two-phase hyperlink setter carries both native staged charges before dropping
the first proof. Signed PPTX no-op canonicalization keeps the larger staged or
identity charge, and XLSB find no-match proofs select the first actual
worksheet with the same bounded `TransferLimits` as a matched proof.

Signed metadata and feature probes retain their authored no-op semantics and
snapshot charges across XLSX, XLSB and the legacy adapters. XLSB unsigned
metadata and feature history remains lazy (it does not serialize full endpoint
images or walk retained payloads just to account an unsigned no-op). Signed
XLSB metadata and each heterogeneous feature family are re-rooted in their
actual native empty patch, while the authored command count survives the
zero-parts native diagnostic. XLSX metadata/feature wrappers likewise carry
the authored scalar count instead of deriving it from the empty workbook
patch. PPTX metadata no-ops join the same native-empty signed path, and its
direct comment/notes/text adapters preserve one authored command even when
the native patch is empty. Direct PPT/PPTX coalescers report authored command
counts after native expansion and after signed identity replacement.
No synthetic no-op marker, signature re-marking, shared-core commit change, or
unsigned tracking scan was introduced.

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
- locked and unlocked-encrypted recipient protection remaining intact;
- an invalid intermediate target returning its typed path error while both the
  signed leaf and the `File` facade retain their original publication bytes.

Static validation for this follow-up passed on all changed Rust files:

```text
rustfmt --edition 2024 --check <all changed Rust files>
git diff --check
```

The focused build and tests remain pending the coordinator's disk-pressure
hold:

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
