# RTF paragraph-frame validation

This evidence records revision v4 of the isolated RTF positioned-paragraph-frame
batch at `/var/tmp/litchi-rtf-paragraph-frames-validation-20260909`. The candidate is based on
`1b3f2c2d0c059e8e59272775a97d0567e14e67f2` and contains 20 source/test paths listed, with
SHA-256 values, in [`source-manifest.json`](source-manifest.json). The evidence
files are docs-only; no source path was staged or committed.

The batch models the RTF 1.9.1 positioned-paragraph controls with bounded values,
style/default inheritance, drop-cap and text-flow ordering, source-aware body
layout publication, and paragraph-frame spans in table cells. Nested and
inherited table-cell frames are retained, and row validation enforces the
specification rule that positioned paragraphs in one table row share the same
controls. The common `\wrapdefault` paragraph default remains line-breaking
state unless other frame controls make a positioned frame present.

Revision v3 preserves the distinction between an in-paragraph `\line` and a
`\par` boundary when cell text is replaced with the same value or when existing
break metadata can be remapped. The frame-aware table writer emits those forms
separately. The external soft-line reproducer passed after save/reopen.

Full v3 root validation used Rust 1.95.0 with the recorded repository dependency graph
(the commands did not pass `--locked` or `--offline`):
1,348 unit/integration tests and 9 doctests passed, with zero failures and zero
ignored tests. Strict all-target/all-feature Clippy, warning-denied rustdoc,
pinned rustfmt, `git diff --check`, and the external soft-line reproducer passed.
Exact commands, environment, `Cargo.lock` hash, source hashes, root receipt hash,
and compressed/raw log hashes are recorded in [`receipt.json`](receipt.json).
Logs are compressed with a zero mtime for reproducibility: [`tests.log.gz`](tests.log.gz),
[`clippy.log.gz`](clippy.log.gz), [`rustdoc.log.gz`](rustdoc.log.gz), and
[`softline-probe.log.gz`](softline-probe.log.gz). The root receipt is retained as
[`root-gates-receipt.json`](root-gates-receipt.json).

The semantic boundary is deliberate. The crate does not render or paginate
frames. Paragraph-layout transactions use direct paragraph properties and refuse
style-dependent effective-property edits, tables, and unsupported dependent
positioned structures before publication. Accepted source-aware rewrites retain
opaque nodes and supported zero-width body metadata; unsupported dependency
closure and arbitrary table mutation remain refusals. The flat table-cell text
API records bounded paragraph spans and frame metadata but does not claim a full
table layout engine or native-producer acceptance. Independent review is clear for v4. The prior full v3 gate is retained separately;
v4 changed only the table-cell invariant/no-op implementation and its tests.
All 32 focused tests, strict Clippy, warning-denied rustdoc and formatting passed.
The v4 row check includes implicit unframed legacy cells, including nested and
empty cells; same-text updates retain spans without allocating a clone.
