# 0370e — W1-A slice 6: table resize, and the sheet/table lifecycle divergences

- Status: landed on branch `uo2/iwa-w1a`, after slice 5
  ([0370d](0370d-iwa-w1a-conditional-highlights.md)).
- Date: 2026-10-26.
- `performance_claim`: leaf resize runs 60.4% below the host route (§2).

## 1. `litchi_numbers::table::resize`

`Package::{edit_table_resize, apply_table_resize}` with
`Edit::resize(rows, columns)`. Growth adds blank trailing rows and columns.
Shrinking is accepted only when the removed trailing region holds no
stored cell (`Refusal::StoredCells`), and a size that leaves no body row
or column is refused. The route ports the host's `resize_table`:

- the tiles drop the removed rows' records, with `numrows` counted by
  extent as the host does;
- the header buckets drop the headers outside the new size;
- the UID map loses or gains trailing identities through
  `table_topology::storage::UidMap::resize`, which allocates new
  identities above every existing one;
- the stroke sidecar takes the new size;
- the model's rows and columns change;
- the table's own formula owner moves its table and body extents
  (`table_topology::extents::resize_extents`).

**Host defect fixed** (mac-queue row): a resize left that owner's
`total_range_for_table` and `body_range_for_table` at the old size. The
host now applies the same `resize_extents`.

Parity: `numbers_table_resize_parity` (tier A, `basic.numbers`: +3 rows,
+2 columns, +5×+4, −1 row) and `numbers_table_resize_corpus_parity` (tier
B, 109 corpus routes; where the host's table-appearance read refuses, the
leaf edits). One ledger row is migrated.

## 2. Performance

`cost_probes::numbers_table_resize_cost_abba` (+5 rows, +4 columns on
`basic.numbers`): host 157,193,923 Ir, leaf 62,212,874 Ir, **−60.4%**.

## 3. Sheet and table add/remove: findings

`numbers_lifecycle_corpus_parity` (ignored, exploratory) compares
`add_empty_sheet`, `remove_sheet`, `add_empty_table` and `remove_table`
with the leaf's `edit_sheet_lifecycle` / `edit_table_lifecycle` on the
corpus. It compares 80 routes, and none passes tier B:

| Divergence | Host | Leaf | Native-correct |
|---|---|---|---|
| root previews | kept | deleted | either (the app regenerates them); align the leaf to keep them, as the other W1-A families do |
| new sheet record | bare `TN.SheetArchive` (name only) | the first sheet's settings (page setup, style) without its drawables, headers and footers | leaf (the builder's own sheet carries those settings) |
| sidebar order | untouched (a new sheet or table is missing from the sidebar tree) | a `TSK.TreeNode` per new sheet, and per new table under its sheet's node | leaf (the host's own builder writes those nodes): host defect |
| watermark | `add_empty_sheet` leaves `last_object_identifier` | raised | leaf: host defect |
| table identifier order | info, model, owned storage in source-id order, formula owner | info, model, sidebar node, formula owner, storage | host order with the node last (W1-B's Pages comparison prefers the host order too) |
| new formula owner | empty spanning dependencies | the new table's extents | leaf (native owners carry them) |
| table-info clone | clears 6 fields and keeps the parent reference | clears 10 fields | host (W1-B: native keeps that field set) |

Next step, shared with W1-B's Pages `add_table` (coordinator request):

1. Align `litchi_numbers_wire::table_create` and the Numbers leaf `add`
   with the host's identifier order and info clone.
2. Give the host its missing sidebar nodes, watermark and owner extents.
3. Make the leaf keep the previews.

## 4. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa --all-features` | 0 |
| `cargo test --all-features -p litchi-numbers`, `-p litchi-numbers-wire` | 0 / 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal` / `numbers-full` | 0 / 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
