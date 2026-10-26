# 0370c — W1-A slice 4: Numbers row and column removal at parity

- Status: landed on branch `uo2/iwa-w1a`, after slice 3
  ([0370b](0370b-iwa-w1a-cell-paragraph-lists.md), merged as
  `3ab7df643a`).
- Date: 2026-10-26.
- `performance_claim`: leaf removal runs 56.6% below the host route (§4).

## 1. What landed

Row and column removal (`Package::edit_table_topology` with
`Edit::{remove_row, remove_column}`) now match the host. Record 0370 §2.1
had pinned the divergence (`numbers_table_removal_divergence_is_reported`);
that pin is replaced by parity cases, and the ledger migrates
`remove_table_row` and `remove_table_column` (tier A).

| Test | Routes | Tier |
|---|---|---|
| `numbers_table_remove_row_parity` | body row 0 of `basic.numbers` | A |
| `numbers_table_remove_column_parity` | body column 0 of `basic.numbers` | A |
| `numbers_table_removal_corpus_parity` | header, body 0 and 1 and footer rows; header, body 0 and 1 columns, on the whole corpus: 120 routes compared, 109 byte-identical, every one tier B; where the host refuses (its table-appearance read, formulas that read the removed axis) the leaf edits or refuses too | B |

## 2. Changes

### Leaf (`litchi-numbers`, `litchi-numbers-wire`)

- A table no formula owner touches (`table_write::table_formula_state` is
  `Inert`) drops the removed cells in the structural pass instead of
  clearing them through the cell transaction first, as the host's axis
  removal does. The root previews stay. A table with formulas keeps the
  record 0329 path (cell clear, cache refresh, then the structural pass).
- Comments on the removed axis are cleared by a comment-family variant for
  axis removal (`clear_table_cell_comment_for_axis_removal`): previews stay
  and a released last comment key is handed back.
- Released list keys hand the last key back one release at a time, in cell
  order, as the host's per-cell decrements do (`reclaim_in_order`).
- `table_topology::storage::tile_columns` drops a row the removal leaves
  without cells (with the tile's `numrows` convention); `header_bucket`
  drops a header left counting no cell and holding nothing else; a BNC row
  the host's column shift rebuilds (no spare slot, or a cell after the
  removed column) states its offset width.
- New `table_topology::storage::pre_bnc_remove_column` for the host.

### Host defects fixed (`litchi-iwa`, mac-queue row)

1. A cell dropped by an axis removal kept its style, text-style,
   conditional-style, control and format list entries counting it
   (`model::release_dropped_cell_metadata`).
2. A header or footer removal left the spanning dependencies' body range at
   the old section edges: `delete_formula_dependencies` now takes the
   before/after extents (`row_insert::removal_extents`), as insertion does.
3. A column removal left the pre-BNC offset table naming the removed column.

`numbers_topology_formula_differential` and the host's own tests still
pass.

## 3. Gaps

Unchanged from 0370 §5 apart from removal: conditional highlights,
sheet/table add/remove, duplicate/resize, sheet duplicate, Builder options,
cell comments, generic data format, fill/border/layout, hidden axes,
pivot, formula AST.

## 4. Performance

Callgrind on the release parity binary, CPU 3 pinned,
`cost_probes::numbers_topology_remove_cost_abba` (body row 0 and body
column 0 of `basic.numbers` and `custom-number-native.numbers`, host,
leaf, leaf, host each).

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| row and column removal | 972,264,278 | 422,337,712 | **−56.6%** |

## 5. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa --all-features` (every target) | 0 |
| `cargo test --all-features -p litchi-numbers`, `-p litchi-numbers-wire` | 0 / 0 |
| `cargo test -p litchi-ops --all-features --test numbers_unified` | 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
