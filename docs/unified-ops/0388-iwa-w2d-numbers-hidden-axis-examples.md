# 0388 — W2-D: relocate Numbers hidden-axis examples

- Status: implementation slice on `uo2/iwa-w2d-hidden-axis-examples`, cut
  from frozen coordinator source `9932c7d0691b2c14bbbd11a2dea7df41f66308b6`.
- Date: 2026-10-28.
- Scope: relocate the two Numbers hidden-axis examples whose behavior is
  supported by the focused `litchi-numbers` crate. The physical hidden-axis
  deletion example remains host-only because the focused topology transaction
  intentionally refuses `Unsupported::HiddenAxis`. This record is not a merge
  or W3 deletion approval.

## Scope and API adaptation

| Before | Destination | Leaf route |
|---|---|---|
| `crates/litchi-iwa/examples/create_numbers_table_hidden_axes.rs` | `crates/litchi-numbers/examples/create_numbers_table_hidden_axes.rs` | `Builder`, `Package::edit_table_hidden_axes`, `SheetSelector`, `TableSelector` |
| `crates/litchi-iwa/examples/edit_numbers_hidden_table_topology.rs` | `crates/litchi-numbers/examples/edit_numbers_hidden_table_topology.rs` | `Package::edit_table_topology` with selector-first table coordinates |

`crates/litchi-iwa/examples/delete_numbers_hidden_table_axes.rs` is retained
unchanged as an explicit host-only gap; clearing visibility is not equivalent
to physically deleting the row and column.

The create example uses the deterministic focused `Builder` and reopens the
saved candidate to verify the selected hidden axes. The topology example uses
the focused row/column insertion transaction, verifies the table extent grows
from 6×4 to 7×5, and reopens the saved candidate to verify both the extent and
the shifted hidden-axis positions.

## Verification

With `CARGO_BUILD_JOBS=2`, `CARGO_INCREMENTAL=0`, and
`CARGO_PROFILE_DEV_DEBUG=0`:

- `cargo build --offline -p litchi-numbers --example create_numbers_table_hidden_axes --example edit_numbers_hidden_table_topology` — pass.
- The two compiled leaf examples were run against a generated Numbers source:
  create and topology edit completed successfully.
- The host physical-delete example remains unchanged and is not claimed as a
  leaf relocation; the focused topology refusal is the explicit W2-D gap.
- No performance claim, native acceptance claim, parity claim, ledger update,
  native queue change, runtime change, or fixture change is included.
