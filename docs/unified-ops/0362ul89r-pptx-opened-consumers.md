# 0362ul89r — PPTX opened/source consumer repair

Status: source-ready for independent review. Cargo, rustc, rustdoc, native
execution, and runtime tests were not run in this source-only lane.

Base: `15d8d53a0cf6d27f350b866463b4235bf26b08d9`.

The assigned keep-going diagnostics contained 130 entries. The compiler log is
`(deleted-coord-artifact)/01-pptx-all-features-check-attempt-actual-f5-01.log`
(100,739 bytes, SHA-256
`a31129804840cddda84361dc4201852856d55afc25de7f0d49cc7d5868295084`); its
diagnostics index SHA-256 is
`c25bee0ed0c486cd54e2d132c928509bfb08484991437136c53c403104b5d3be`.

## Repairs

- Migrated opened notes proof tests to
  `load_snapshot_with_slide_root_proofs_and_limits`, deriving the explicit
  notes policy from the package source limits. Limit-pattern tests now bind
  `observed` and retain exact observed/maximum assertions.
- Updated cross-copy test writers and limit patterns for the typed `observed`
  field, including exact candidate/archive observations.
- Updated the source-backed text-structure assertion to the related-notes
  projection's zero-argument `text()` API and fixed the default-profile source
  test for its `u32` shape-id return.
- Added test-module imports for `OnceLock` and `ExecutionError`, removed the
  unused builder/part-edit imports, and removed unnecessary `crate::Error`
  qualifications from source tests.
- Added a narrow `#[cfg(test)] pub(crate)` re-export of the existing
  `fingerprint_log` from `opened`; `part_edit` tests use it without widening
  production visibility, preserving the repeated-replay memo-cost assertions.

Only the assigned consumer/test files, the narrow test re-export, and this
record changed. No default notes wrapper was restored and no selected policy
was replaced by a fallback.

`rustfmt --emit stdout --config skip_children=true` parsed every edited Rust
file, and `git diff --check` passed. No runtime or performance claim is made.
