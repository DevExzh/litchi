# 0201 — XLS encryption profile retention

Canonical XLS owners now expose `Package::encryption()` and
`Package::open_with_password(path, password)`. Successful admission retains
`EncryptionProfile`, including CryptoAPI RC4 key length, instead of only a
boolean. This is required groundwork for the ADR 0034 rule that re-encryption
must preserve the source profile rather than choose a default.

The raw reader carries the same profile after decryption. The shared crypto
header exposes its validated, normalized key length; XLS does not parse header
offsets a second time. A wire key length of zero reports the specified 40-bit
default. Unsupported profiles and invalid passwords still fail admission.
Plain eager and positional source owners report `None`. Profile reads do not
read source bytes, derive keys, retain a password, or change publication policy.
Ordinary publication of encrypted owners still returns `EncryptionPolicy`.

## Verification

The combined native XLS, crypto, and XLS/encryption facade suites pass 2,323
tests and doctests, with four ignored. New acceptance tests cover XOR, binary
RC4, every supported CryptoAPI key length, owner cloning, exact encrypted
source retention, ordinary publication refusal, plain eager/source owners,
and password-path admission with typed missing/wrong-password failures. A
shared crypto unit test covers normalized zero and all supported key lengths.

Warning-denied all-target Clippy and documentation pass for XLS, crypto and
the XLS/encryption facade. The XLS glossary audit passes 24 banned names and
one transaction family. Both repository lockfiles are unchanged; builds used
an external lockfile and target. The rewritten base remains an ancestor.
The crate-boundary check passes.

Reproduction: `cargo test -p litchi-xls -p litchi-crypto -p litchi-ops
--no-default-features --features litchi-ops/xls,litchi-ops/encryption`, then
`cargo clippy` with the same packages/features plus `--all-targets -- -D warnings`,
and `RUSTDOCFLAGS='-D warnings' cargo doc` with the same packages/features plus
`--no-deps`. Run `python3 tools/check_api_glossary.py --profile xls-minimal`
and `python3 tools/check_crate_boundaries.py`. Offline verification used a
copied lockfile with locally cached cc 1.6.0 rather than the pinned 1.7.0.
No performance, Python, or CLI acceptance is claimed by this native metadata
change; those full-goal gates remain outstanding.

## Remaining scope

This change supplies the profile projection and password-opening verb. It does
not implement `save_encrypted`, `save_reencrypted`, or the explicit plaintext
conversion of encrypted sources. Those require preserved CFB streams, correct
BIFF offset relocation, signature policy, and atomic publication. Builder
entropy remains the explicit gap recorded in 0200; no salt reuse is introduced.
Unified XLS CRUD, batching/history, remaining format coverage, Python/CLI,
paired performance qualification and full-goal acceptance remain open.
