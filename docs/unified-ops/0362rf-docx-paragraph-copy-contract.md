# 0362rf — DOCX plain-paragraph copy and removal contract

Status: complete on `uo2/w8-docx-paragraph-copy`.

This slice closes the remaining DOCX source-backed paragraph-copy and
paragraph-removal API debt. The package-facing operations now use the shared
edit family contract, while detached XML work keeps a focused raw codec.

## Public API map

- `source_backed::plain_paragraph_copy` and
  `source_backed::plain_paragraph_removal` are the package families. Their
  `edit_*` methods stage a private record edit, `commit()` returns
  `litchi_core::edit::Commit<source_backed::Package, Patch, Diagnostics>`, and
  `apply_*` replays a family patch.
- The old public record `Edit`, `Commit`, `RemovalEdit`, and `RemovalCommit`
  paths are private implementation details. There is no compatibility alias.
- `source_backed::paragraph_copy::raw` is the detached XML codec. Its
  `Tree::from_xml` requires `Limits`, keeps staging private, and exposes exact
  copy/removal operations, source identity, durable patch codecs, and
  `&self`-based bounded `to_bytes` and `write_to` outputs.

The raw codec preserves the existing exact XML fragment behavior, refusal and
typed limit errors, no-op and inverse semantics, content freshness checks, and
retained OPC members. Package-family replay retains whole-artifact freshness;
raw tree replay is authorized by the recorded XML bytes. Tree mutation replaces
its model only after the operation succeeds. Output checks both the caller
policy and the admission policy before allocation or writing; `write_to`
streams the admitted bytes without an additional full XML clone. The raw
wrapper adds no retained XML copy beyond the existing Arc-backed record
representation.

All DOCX integration tests and the performance harness consumers were migrated
in this slice. Raw types are used only for detached XML and durable wire
handling; package consumers use the canonical family operations.

## Validation

- `cargo test --offline --locked -p litchi-docx --all-features --test source_backed_paragraph_copy`
  — 13 passed.
- `cargo test --offline --locked -p litchi-docx --all-features --test source_backed_paragraph_removal`
  — 6 passed.
- `cargo check --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets`
  — passed.
- `cargo clippy --offline --locked -p litchi-docx --all-targets --all-features`
  — passed.
- `cargo clippy --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets --all-features`
  — passed with pre-existing warnings outside the migrated DOCX consumers.

No performance measurement or compiled artifact is part of this record.
