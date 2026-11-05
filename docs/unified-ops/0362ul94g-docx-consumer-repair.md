# 0362ul94g — DOCX consumer compiler repair

This source-only follow-up repairs the DOCX diagnostics from the actual
consumer capture on `5ec78a48589c7ab4396ee90cb1f6d41cf7f1aaa6`. The immutable
capture is retained outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/ooxml-combined-0362ul94-f5-consumers-5ec78a4858`.

The captured DOCX failures were:

- `package/model.rs:808` passed `Arc<Vec<u8>>` to the VBA writer's owned
  `Vec<u8>` parameter. The adapter now moves `Payload::into_bytes()` directly
  into the one-shot semantic edit, avoiding a detached whole-payload copy.
- `package/package/metadata.rs:149`, `:152`, and `:158` lacked conversion from
  the shared properties error. DOCX now flattens `Refused` (retaining its
  `RefusalKind`), `PatchConflict`, `Signed`, `SourceChanged` (retaining both
  versions), and `Internal` explicitly; ordinary invalid, limit, and package
  failures remain represented rather than being silently discarded.
- `metadata.rs:370` now maps the OPC failure through the DOCX error boundary;
  `:459` consumes a named edit instead of moving from a mutable reference;
  `:532` supplies the byte-window length; and the now-unused `Part` import is
  removed.

The repair does not touch the separately assigned DOCX data-store,
bibliography, or mail-merge consumers. The original capture remains a failed
receipt: its 33,227-byte compiler log has SHA-256
`bde73361831156dabd2d8b7032b40ef472506856aa4dda8be40781425eaba220` and its
117-byte cargo-target manifest has SHA-256
`c2dedca2f0195d3c7e6979e469e8013e09162f94fdf29053006c6155d7dbd837`.

No Cargo, rustdoc, runtime, or performance command was run for this repair;
the coordinator must rerun the frozen consumer gate on the composed source.
