# 0370f — W1-A slice 7: adding a table at host parity

- Status: landed on branch `uo2/iwa-w1a`, after slice 6
  ([0370e](0370e-iwa-w1a-table-resize-and-lifecycle-findings.md)).
- Date: 2026-10-26.
- `performance_claim`: leaf table add runs 91.2% below the host route (§4).

## 1. What changed

Record 0370e §3 listed how the host's `add_empty_table` and the leaf's
`edit_table_lifecycle(..).add(..)` diverge. W1-B's Pages `add_table`
comparison found the same differences, and the coordinator settled them:
the leaf takes the host's identifier order and info clone, and keeps its
random version-4 table UUIDs. Where the native files disagree with the
host, the native files win and the host is fixed.

Leaf (`litchi_numbers::package::table_lifecycle::add`) and the shared
`litchi_numbers_wire::table_create` (W1-B's Pages route):

- **Identifier order.** The order is now the host's: the drawable, the
  model, the owned storage in source-identifier order, the formula owner,
  then the sidebar node. The registry records follow identifier order.
- **Info clone.** It clears the host's six fields: editing state, summary
  model, category order, view UIDs, pivot data model and pivot order. The
  group-by and hidden-state UUIDs, the pasteboard mapper and the pivot flag
  are now kept. Before, the leaf cleared ten fields.
- **Model references.** These are the host's `table_model_references`
  field list, sorted.
- **UID map.** It uses the host's fresh-map identities, seeded with the
  new model's identifier, with unpacked index lists as native maps store
  them. Callers of `table_create::empty_storage` now pass the new model
  identifier as the seed.
- **Stroke sidecar.** It holds the new size only.
- **Previews.** The root previews are kept, as in every other W1-A family.

**Host defects fixed** (mac-queue row):

- `create_empty_table_formula_graph` wrote a new table's formula owner
  with empty spanning dependencies. Native owners record the table and
  body extents (`TableExtents::spanning_dependencies`).
- `add_empty_table` added no sidebar node, and `remove_table` left a
  removed table's node behind. The new `numbers::editor::sidebar` module
  adds and removes the node under the sheet's node, as the builder and the
  native app do.
- `register_cloned_numbers_objects` registered UUIDs and cross-component
  references in hash order, so the same edit produced different metadata
  bytes on each run. The order is now the new identifiers' order.

One ledger row is migrated: `add_empty_table` →
`litchi_numbers::Package::edit_table_lifecycle`.

## 2. Parity

`crates/litchi-iwa/tests/parity/numbers_lifecycle_cases.rs`:

| Test | Routes | Tier |
|---|---|---|
| `numbers_table_add_parity` | a 4×3 table on the first sheet of `basic.numbers` | B |
| `numbers_table_add_corpus_parity` | the same on each corpus fixture: 23 routes; where the host refuses, the leaf edits | B |

Tier B because of three documented policies:

- The table UUID is a fresh version-4 identity on the leaf and
  model-derived on the host. The formula-owner UUID and the owner-map
  entry derive from it, so all three are compared by form.
- On `chart-arrangement-native` and `shared-model-native` the owner map
  holds internal identifiers above every owner object. The host allocates
  a colliding internal identifier there and the leaf allocates one above
  the map (policy 0368a).

## 3. Differential oracle

`numbers_topology_formula_differential` builds its synthetic workbooks
with the leaf's table add. With the new identifier order, the host
editor's known shared-entry defect now shows on `synthetic-incoming`: when
rows are inserted into `Other`, its two relative references into
`Table 1` read `B2` and `B3` on the host, but must both keep `B3`. The
oracle skipped every cross-table text, so it could not classify this.
Now a reference into another table keeps its text at the moved host
cell, and the case is classified as "host formulas wrong, leaf matches the
oracle".

## 4. Performance

`cost_probes::numbers_table_add_cost_abba` (a 4×3 table on
`basic.numbers`): host 1,185,287,142 Ir, leaf 104,632,359 Ir, **−91.2%**.

## 5. Gaps

| Item | Next step |
|---|---|
| sheet add/remove (record 0370e §3: the new sheet record, the sidebar node, the watermark, the removed stylesheet objects) | next slice: host `add_empty_sheet` takes the first sheet's settings and a sidebar node, and raises the watermark |
| `remove_table` parity | same slice |
| no native evidence for an added table | mac queue |

## 6. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa --all-features` | 0 |
| `cargo test --all-features -p litchi-numbers`, `-p litchi-numbers-wire` | 0 / 0 |
| `cargo test --all-features -p litchi-pages`, `-p litchi-keynote` | 0 / 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers -p litchi-numbers-wire` | 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
