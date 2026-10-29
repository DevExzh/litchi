# 0398 — W2-D: relocate the Numbers sparse tile-boundary example

- Status: source slice prepared on `uo2/iwa-w2d-sparse-example`, cut from
  frozen coordinator source `e0e297a82b0a534dfd5611ced3a9811851edbf76`.
- Scope: move one source-built command example from the migration host to the
  focused Numbers crate. Builds, runs, and captures are intentionally pending
  the coordinator's quiet-interval release.
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

The host example and host/oracle tests remain outside this slice's ownership
of parity evidence. This relocation changes no runtime implementation,
ledger row, native queue row, fixture, performance receipt, or native
acceptance claim. The host example may be restored if the focused route
refuses the actual 257-row tile-boundary workload; the input will not be
simplified to hide such a refusal.

## Verification plan

After quiet-interval release, using the existing repository cache and no
third-party inputs:

- build and run only `litchi-numbers --example create_numbers_sparse_table`
  against an external output path;
- verify the three values, 257 × 2 dimensions, both inverse checks, and
  reopened readback;
- run scoped formatting, strict Clippy, rustdoc/glossary as applicable,
  crate-boundary, no-3rdparty, and repository size/binary guards.
