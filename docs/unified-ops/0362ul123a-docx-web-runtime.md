# 0362ul123a — DOCX durable web tamper-error consumers

Date: 2026-10-25

## Scope

This source-only follow-up is based on the `123` repair commit
`04dc8f75bfb68452c8d0fb7a43449c35ed04378c` and audits every error assertion in
`durable_stale_and_wire_tampering_are_atomic`. It addresses the next runtime
review return without changing the producer or publication path.

The preserved diagnostic remains
`(deleted-coord-artifact)/12-tests-docx-web-extensions-attempt-actual-f5-12.log`
(18,605 bytes, SHA-256
`ec5efd1ebbcd501dc42cd9756e8c26a73c7d1a9e11c99eb78a26f0bad41e44a9`): 33
tests passed and `durable::durable_stale_and_wire_tampering_are_atomic` was
the one failure at the pre-123 target-error assertion. No post-repair runtime
claim is made.

## Producer classification

On frozen source `a8d32513b8e427c6d234ddd17fb836b0e8d2477b`:

- The `web.edIt` operation tamper is rejected by
  `litchi_ooxml_common::web::package::durable::plan_durable` as shared
  `Error::Invalid`, which the DOCX adapter exposes as
  `Error::Common(Error::Invalid(_))`.
- A changed `target_sha256` reaches the generated-target comparison in the
  same planner and returns shared `Error::PatchConflict`, which DOCX maps to
  its typed `Error::PatchConflict` variant.
- The altered intent blob has a matching replacement blob digest, so it is
  decoded and replayed. Its changed runtime value makes the generated target
  differ, taking the same shared `Error::PatchConflict` path.

The operation assertion now matches the typed shared invalid variant; the
target and intent assertions use the existing typed DOCX patch-conflict helper.
Every case still goes through `assert_rejected_unchanged`, so rejection and
unchanged publication bytes remain mandatory.

## Validation

No Cargo, build, rustdoc, metadata, or runtime command was run. With the
pinned repository `rust-toolchain.toml` and `rustfmt.toml`, the two affected
Rust test files pass the source-only formatter check, and `git diff --check`
passes. The runtime log above remains the pre-fix evidence.
