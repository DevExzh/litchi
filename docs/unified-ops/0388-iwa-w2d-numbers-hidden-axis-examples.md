# 0388 — W2-D: relocate Numbers hidden-axis examples

- Status: implementation slice on `uo2/iwa-w2d-hidden-axis-examples`, cut
  from frozen coordinator source `9932c7d0691b2c14bbbd11a2dea7df41f66308b6`.
- Date: 2026-10-28.
- Scope: relocate exactly three Numbers hidden-axis examples from the host
  crate to the focused `litchi-numbers` crate. This record is not a merge or
  W3 deletion approval.

## Scope and API adaptation

| Before | Destination | Leaf route |
|---|---|---|
| `crates/litchi-iwa/examples/create_numbers_table_hidden_axes.rs` | `crates/litchi-numbers/examples/create_numbers_table_hidden_axes.rs` | `Builder`, `Package::edit_table_hidden_axes`, `SheetSelector`, `TableSelector` |
| `crates/litchi-iwa/examples/edit_numbers_hidden_table_topology.rs` | `crates/litchi-numbers/examples/edit_numbers_hidden_table_topology.rs` | `Package::edit_table_topology` with selector-first table coordinates |
| `crates/litchi-iwa/examples/delete_numbers_hidden_table_axes.rs` | `crates/litchi-numbers/examples/delete_numbers_hidden_table_axes.rs` | `Package::edit_table_hidden_axes(...).clear()` |

The create example uses the deterministic focused `Builder` and reopens the
saved candidate to verify the selected hidden axes. The topology example uses
the focused row/column insertion transaction and verifies that hidden-axis
positions shift from row 2/column 1 to row 3/column 2.

The focused topology API intentionally refuses physically deleting a hidden
row or column with typed `Unsupported::HiddenAxis`. The relocated delete
example therefore demonstrates the supported selector-first equivalent:
clearing the selected table's hidden-axis state and saving the reopened
candidate. No compatibility or hidden-axis refusal was weakened.

## Verification

With `CARGO_BUILD_JOBS=2`, `CARGO_INCREMENTAL=0`, and
`CARGO_PROFILE_DEV_DEBUG=0`:

- `cargo build --offline -p litchi-numbers --example create_numbers_table_hidden_axes --example edit_numbers_hidden_table_topology --example delete_numbers_hidden_table_axes` — pass.
- The three compiled examples were run against a generated Numbers source:
  create, topology edit, and hidden-axis clear all completed successfully.
- No performance claim, native acceptance claim, parity claim, ledger update,
  native queue change, runtime change, or fixture change is included.
