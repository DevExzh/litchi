# 0202 — Explicit XLS encryption entropy

RC4 Builder authoring now takes caller-owned `EncryptionEntropy` through
`set_encryption_entropy`. `EncryptionEntropy::generate()` is the explicit
ambient-randomness boundary. `from_bytes(salt, verifier)` supports fixed inputs
for reproducible pipelines. The entropy object is move-only and its storage is
zeroized when dropped. It is never cloned with a password configuration.

`build()` does not consult the random generator. A configured RC4 build without
entropy fails with `Error::EncryptionEntropyRequired` before staging a default
sheet. Each RC4 build attempt takes the input, including attempts that fail
later authoring or strict package admission. Repeated or changed encrypted
builds therefore require the caller to supply entropy again. This avoids the
unsafe alternative of caching a salt across different plaintexts. Plain and
explicit weak-XOR builds need no entropy. Callers using fixed bytes must not
reuse a salt/password pair across different plaintexts.

Existing encrypted fixtures now call the entropy generator explicitly. The
Builder/Package ownership and ciphertext retention paths remain unchanged.

## Verification

The combined XLS and XLS/encryption facade suites pass 2,242 tests and
doctests, with one ignored. Tests cover missing entropy before staging, byte
identity with fixed inputs, fresh entropy for changed builds, consumption after
strict admission failure, and preservation of unused entropy across plain/XOR
builds. A compile-fail doctest prevents cloning entropy tokens. The facade maps
missing entropy to `InvalidArgument`.

Warning-denied all-target Clippy and documentation pass. The XLS glossary
audit passes 24 banned names and one transaction family. The boundary check
passes for 66 packages and 268 internal dependencies, with 11 existing debt items. No benchmark or Python/CLI
acceptance is claimed by this native authoring-policy change.

Reproduce with `cargo test -p litchi-xls -p litchi-ops --no-default-features
--features litchi-ops/xls,litchi-ops/encryption`; run Clippy for the same packages
and features with `--all-targets -- -D warnings`, and warning-denied rustdoc.
Offline builds used an external lockfile selecting cached cc 1.6.0 and an
external target. Repository lockfiles remain unchanged.

## Remaining scope

This closes the inherited ambient-RNG call inside XLS Builder build. It does
not implement the pending Package security publication verbs, source-preserving
plaintext conversion, or re-encryption. Full format read/write coverage,
unified XLS CRUD/batching/history, Python/CLI, paired performance acceptance,
and the other full-goal gates remain open.
