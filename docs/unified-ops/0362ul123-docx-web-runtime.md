# 0362ul123 — DOCX durable web target-error consumer

Date: 2026-10-25

## Scope

This source-only follow-up repairs the stale assertion exposed by runtime
qualification 123. The leaf is based on source commit
`a8d32513b8e427c6d234ddd17fb836b0e8d2477b`.

The preserved diagnostic is
`(deleted-coord-artifact)/12-tests-docx-web-extensions-attempt-actual-f5-12.log`
(18,605 bytes, SHA-256
`ec5efd1ebbcd501dc42cd9756e8c26a73c7d1a9e11c99eb78a26f0bad41e44a9`). It
reported 33 passing tests and one failure:
`durable::durable_stale_and_wire_tampering_are_atomic` at
`crates/litchi-docx/tests/web_extensions/durable.rs:523`, where the test
expected the error text to contain `target`.

## Diagnosis and repair

The tampered `target_sha256` is consumed by
`litchi_ooxml_common::web::package::durable::plan_durable`. Its generated
target mismatch returns the common `Error::PatchConflict`; the DOCX conversion
maps that to the typed `litchi_docx::Error::PatchConflict` variant with the
fixed, content-free reason `the selected package graph changed`. The former
string assertion therefore tested an obsolete diagnostic, not the producer's
contract.

The test now uses the existing `assert_patch_conflict` helper. It continues to
run through `assert_rejected_unchanged`, so the target-tamper rejection and
unchanged publication bytes remain required. The operation and intent-blob
tamper checks are unchanged.

## Validation

No Cargo, build, rustdoc, metadata, or runtime command was run for this
source-only repair. With the repository's pinned `rust-toolchain.toml` and
`rustfmt.toml`, both affected Rust test files pass:

```text
rustfmt --edition 2024 --check crates/litchi-docx/tests/web_extensions.rs crates/litchi-docx/tests/web_extensions/durable.rs
git diff --check
```

The runtime diagnostic above remains the pre-fix result; this record makes no
post-fix runtime-pass claim.
