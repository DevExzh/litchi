# 0362ul94va — OPC publication Clippy qualification repair

The 94v import repair was rerun from
`f167986c745fdd6c305f8095e311980aee24d799`. Formatting passed, but the
all-features leaf Clippy command across DOCX/PPTX/XLSX/XLSB stopped with 35
`unnecessary qualification` diagnostics in the existing `litchi-opc` package
module. The captured log was
`ooxml-combined-0362ul94-f5-final-f167986c/01-clippy-leaves-attempt-actual-f5-01.log`,
30,940 bytes, SHA-256
`f74cbbe6c4c33edc5c14c47d97a2703fe9c4b1ede69ebee0733137b146f9b45d`.

The broad `BlobPart` and `SourceLimitKind` imports are removed. Only the one
new authored-publication `BlobPart` use and the fourteen new publication
`SourceLimitKind` uses are qualified through the canonical `crate::` paths.
All pre-existing qualified references remain untouched, so this repair does
not broaden lint scope or alter feature/test gating, behavior, limits,
allocation, or the public API.

This was prepared in a fresh sparse worktree from the exact 94v commit.
Rustfmt with `skip_children=true`, static checks, and repository commit guards
were run. No Cargo, compiler, build, rustdoc, metadata, native, or performance
command was run.
