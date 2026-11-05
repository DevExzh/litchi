# 0398 — W2-D: relocate the Numbers sparse tile-boundary example

- Status: source slice on `uo2/iwa-w2d-sparse-example`, cut from
  frozen coordinator source `e0e297a82b0a534dfd5611ced3a9811851edbf76`.
- Scope: move one source-built command example from the migration host to the
  focused Numbers crate. No runtime, parity, fixture, ledger, native queue,
  or performance changes are included.
- `performance_claim`: none.

## Relocation and behavior

| Before | Destination | Focused route |
|---|---|---|
| `crates/litchi-iwa/examples/create_numbers_sparse_table.rs` | `crates/litchi-numbers/examples/create_numbers_sparse_table.rs` | `Builder`, `Package::edit_table_cells`, `Package::table_cell`, `Package::save` |

The leaf command preserves the original 257 × 2 table, the three values, and
the two batches: one numeric write at row 256 / column 1 followed by text at
row 0 / column 0 and row 256 / column 0. `Identity::BLANK` is passed
explicitly so the source-built output is deterministic. The example asserts
dimensions and semantic values before publication, exact inverse bytes after
each batch, and all three values after reopening the saved output.

Only this command surface relocates. The host differential/oracle and native-
provenance test targets remain unchanged. This relocation changes no runtime
implementation, ledger row, native queue row, fixture, performance receipt,
or native acceptance claim. The host command may be restored if the focused
route refuses the actual 257-row tile-boundary workload; the input will not be
simplified to hide such a refusal.

## Verification

Using the existing repository cache, four build jobs, offline mode, and no
third-party inputs:

- `CARGO_BUILD_JOBS=4 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 cargo
  build --offline -p litchi-numbers --example create_numbers_sparse_table` —
  pass.
- `./target/debug/examples/create_numbers_sparse_table
  (deleted-coord-artifact)`
  — exit 0. The example's three semantic readbacks, both exact inverse
  assertions, dimensions, and reopened readback all passed. Output SHA-256:
  `bf0ea9cbb993c3e236966a65702b633568d073c001048465a84739604f327081`.
- `cargo fmt --all -- --check` — pass.
- Strict example Clippy with `-D warnings` — pass; existing vendor warnings
  only.
- Warning-denied example rustdoc — pass.
- `python3 tools/check_example_targets.py` — pass; its unit suite: 12 passed.
- `python3 -m unittest tools.test_no_3rdparty_references
  tools.test_iwa_exit_ledger` — 17 passed; `iwa_exit_ledger.py --check` —
  pass.
- Numbers glossary `numbers-minimal` and `numbers-full` profiles — pass.
- `python3 tools/check_crate_boundaries.py` — pass: 70 workspace packages,
  313 internal declarations, 12 explicit debt items.
- Staged repository size and binary guards — no output.

No native acceptance, parity-tier, or performance claim is made by this
example relocation. The host differential/oracle and native-provenance tests
remain unchanged.
