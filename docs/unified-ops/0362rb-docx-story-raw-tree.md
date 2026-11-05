# 0362rb — DOCX story text raw codec trees

This slice closes the duplicate public `Edit`/`Commit` surfaces for the
source-backed story-text records. It keeps package publication in the existing
owner families and exposes the detached record capability through focused raw
codec models.

## API mapping

- `story_text::Snapshot::raw_tree(limits)` returns `story_text::raw::Tree`.
- `story_text::GlossaryBatchSnapshot::raw_tree(limits)` returns
  `story_text::raw::GlossaryTree`.
- The raw trees mutate selected paragraph text and return `Patch` or
  `GlossaryBatchPatch`; they never publish a package or a detached Commit.
- The requested `Limits` value is intersected with the admitted snapshot
  policy. A stricter policy re-admits the XML before the tree is created, and
  output continues to enforce the intersected policy.
- Raw trees cache their projection with interior state. Repeated
  `projected()` and `patch()` calls do not clone the complete document edit.
  A successful mutation invalidates that cache. The private owner path keeps
  the consuming `edit_engine().finish(self)` publication path.

The migration is implemented in `dd9481ed9d` on `uo2/w8-docx-story-text`.
Every source-backed story/glossary test consumer was migrated to an explicit
limits argument. No public compatibility alias remains.

## Validation

All commands were run from the branch worktree with the capped environment,
`CARGO_BUILD_JOBS=2`, `RUST_TEST_THREADS=2`, `LITCHI_MEM_MAX=20G`, and the
external target `(deleted-coord-artifact)`.

- `cargo check --offline -p litchi-docx --tests` — passed.
- `cargo test --offline -p litchi-docx --test source_backed_story_text --test source_backed_secondary_story_text --test source_backed_glossary_story_text` — 57 passed. Receipt: `targeted-tests-final.log`, 5,192 bytes, SHA-256
  `d0d12a07037610f1af6a68655dc9a2074dd6b80e8a76f6ee11a4be3da5dc3e96`.
- `cargo test --offline -p litchi-docx --lib` — 1,479 passed. Receipt:
  `lib-tests-final.log`, 132,482 bytes, SHA-256
  `12e327ea9bd94945521d0dc120aa19cf6f441749f697426ef9c8e2b0030c1c07`.
- `cargo clippy --offline -p litchi-docx --all-targets --all-features -- -D warnings` — passed. Receipt: `clippy-final.log`, 223 bytes, SHA-256
  `daccb17dde7935507a97c44b68f38d8fb9f3081f7099da8f119b76d590e8df9c`.
- Targeted rustfmt checks for all changed Rust files — passed.

The receipts remain outside the repository under
`(deleted-coord-artifact)`; no binaries or
large evidence files are committed.
