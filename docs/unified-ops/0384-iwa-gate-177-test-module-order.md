# 0384 — frozen 177 gate test-module ordering repair

- **Status:** source repair committed as
  `1b66a078fc99a5d3f9cb7623c81c5c51ca22af50` on
  `uo2/iwa-gate-177-repair`; receipt follows as a separate commit.
- **Base:** `177ef72ce063cb5d1b8a66a5e731ba03cc82eb25`.
- **Scope:** repair only the row-28 Clippy `items_after_test_module` findings.
  The existing `#[cfg(test)]` modules were moved, unchanged, to the ends of
  `body_drawable_graph.rs` and `sheet_chart_text.rs`, after the production
  helpers. No runtime, public API, ledger, native queue, parity, or
  performance behavior changed.

## Verification

All commands ran from the repair worktree on the source commit above, with
`CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0` where shown:

| Check | Result |
| --- | --- |
| `CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 cargo fmt --all -- --check` | 0 |
| `CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 cargo clippy --offline -p litchi-pages --all-targets --all-features -- -D warnings` | 0 |
| `CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 cargo clippy --offline -p litchi-numbers --all-targets --all-features -- -D warnings` | 0 |
| `CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 cargo test --offline -p litchi-pages --lib body_drawable_graph::tests -- --nocapture` | 0; 2 passed |
| `CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 cargo test --offline -p litchi-numbers --lib package::sheet_chart_text::tests -- --nocapture` | 0; 1 passed |
| `git diff --check` | 0 |

The staged repository-size and binary-path guards were empty before the source
commit. No raw captures or binaries were added. The branch is left clean and
unintegrated for coordinator review; the full gate remains coordinator-owned.
