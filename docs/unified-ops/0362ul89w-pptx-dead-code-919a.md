# 0362ul89w — PPTX 919a dead test-helper repair

Date: 2026-10-25

## Scope

This source-only repair addresses the terminal `cargo check --keep-going --offline --locked -p litchi-pptx --all-targets --all-features`
diagnostics from the PPTX combined candidate at
`919a772ec37db9bd33b71c22d733415ffae08c6d`.  The retained external compiler
log is:

`(deleted-coord-artifact)/01-pptx-all-features-check-attempt-actual-f5-01.log`

It is 27,685 bytes with SHA-256
`c9daebdd4b088e1f89cad4b09b8da3b2ee8f3aa403dce096d57f7a9eb00109bb`.
The capture has no diagnostics-index artifact; the log itself records the
11 terminal `litchi-pptx` dead-code errors.  No Cargo, build, rustdoc, or
native command was run for this repair.  The retained run receipt reports
the 393-file formatting phase PASS in 2.810 seconds, followed by the compiler
phase exiting 101 (FAIL) in 74.007 seconds; its Cargo target occupied
226,040,568 bytes.  The retained `rows.json` is 31,855 bytes with SHA-256
`98f83704acdc5efdf9ec0417c7dc35208fa550daead957b38273c20c7076f841`, and
`result.json` is 33,542 bytes with SHA-256
`d08874ef183058bf31f2556282e84da2fa8c082ffab7c74bfa3ae30d366c5b77`.

## Repair

The source tip is
`6c9a175a0f41cf2337b1ef51372b1f5e36819519` and changes only:

- `crates/litchi-pptx/src/comments/package.rs`
- `crates/litchi-pptx/src/notes/codec.rs`
- `crates/litchi-pptx/src/presentation/source.rs`

The 11 definition-only `#[cfg(test)]` helpers removed are:

- comments: `add_presentation_comment_author`,
  `update_presentation_comment_author`,
  `remove_presentation_comment_author`,
  `reorder_presentation_comment_authors`, `add_presentation_comment`,
  `update_presentation_comment`, `remove_presentation_comment`, and
  `reorder_presentation_comments` (8);
- notes: `validate_resource_xml` and `scan_xml` (2); and
- source: the private `SourceBackedPresentationEditor::from_source_backed_package`
  wrapper (1).

The retained explicit-limit implementations remain unchanged, including all
`*_with_limits` comment operations, `validate_resource_xml_with_limits`,
`scan_xml_with_limits`, and
`from_source_backed_package_with_semantic_limits`.  No public API, default
fallback, limit value, parser, fidelity, or algorithm was changed.

The remaining test-only helpers were audited before deletion.  The comments
`store_presentation_comments` helper is consumed by collaboration/package
tests; notes `root_conformance` and `scan_processed_xml` are consumed by the
in-file tests; and the retained source instrumentation, editor, slide-edit,
batch-edit, snapshot, diagnostics, and publication helpers have test or
internal callsites.  No `dead_code` allowance or alias was added.

## Reproduction handoff

The exact source patch from the frozen base to the source tip is retained
outside the repository at
`(deleted-coord-artifact)/pptx-919-dead-code-89w-6c9a175a0f.patch`.
It is 7,960 bytes with SHA-256
`81a30f8c024fb84da2267bb1242d49e72713ad05664a460e8108d98bef8b3bcd`.

Source-only checks passed: `rustfmt --edition 2024 --check` on all three
changed Rust files, `git diff --check`, changed-path scope verification,
absence of `3rdparty/` references, and repository large-file/binary commit
guards.
