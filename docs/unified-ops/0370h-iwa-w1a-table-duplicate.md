# 0370h — W1-A slice 9: Numbers table duplicate

- Status: landed on branch `uo2/iwa-w1a`, after slice 8
  ([0370g](0370g-iwa-w1a-sheet-and-table-lifecycle.md), merged on trunk).
- Date: 2026-10-26.
- `performance_claim`: the leaf duplicate runs 90.8% below the host route (§3).

## 1. What landed

`litchi_numbers::table::lifecycle::Edit::duplicate(selector)` stages a
populated copy of a table on its own sheet. The copy is named
`"<name> copy"`, then `"<name> copy 2"` and so on. It goes at the end of the
sheet's drawables, offset by ten points. The new module
`package/table_lifecycle/duplicate.rs` ports the host editor's
`duplicate_table`, that is `duplicate_attached_table_graph_in_package` and
`clone_table_formula_graph`, onto the structural working set.

**What the copy gets:**

- Fresh identifiers, allocated in the host's order:
  - the drawable;
  - the model;
  - every storage object the model owns, in source-identifier order. This
    includes the tiles, header buckets, data lists, the UID map, the stroke
    sidecar, and every data-list segment those lists chain to.
- The host's deterministic table UUID: `00000000-0000-4000-8000-` followed
  by the new model identifier.
- A clone of the table's CalculationEngine owner family. It includes the
  cell-record and range-precedents tiles and the engine's owner list, owner
  map and formula count. The rewrites are the shared
  `litchi_numbers_wire::formula_graph` ones from record 0368: `remap_owner`,
  `remap_cell_record_tile`, `remap_range_precedents_tile`,
  `append_owners_to_engine` and `next_owner_id`.
- Formula lists in which each formula's host-table UUID, and every AST
  reference to the table or its owner, now name the copy.
- A sidebar node under its sheet's node.
- Registry records: UUIDs for clones of registered objects, and every new
  cross-component reference.

**What the copy keeps:** references to anything outside the table (styles,
other tables' owners, the parent sheet), and every field the clone does
not rewrite, unknown fields included. A formula that reads another table
keeps reading it.

**What is refused:** owner state that cannot be remapped without
ambiguity, as the host refuses it. That is volatile geometry references,
spanning, whole-owner, error or spill state
(`Refusal::UnsupportedDependency`). A table without its own owner, in a
workbook with an engine, is also refused (`Refusal::UnsupportedSource`).

**Host defects fixed** (mac-queue row):

- The copy had no sidebar node. `duplicate_table` now adds one under the
  sheet's node through `numbers::editor::sidebar`.
- `clone_table_formula_graph` registered the CalculationEngine clones'
  UUIDs in hash order. They are now registered in identifier order.

**Documented difference** (policy 0368a): the leaf allocates the cloned
owners' internal identifiers above every owner and every owner-map entry.
The host allocates above the owners only, so on
`chart-arrangement-native.numbers` its identifiers collide with the map.
The corpus case compares those identifiers by form on that fixture.

One ledger row is migrated: `duplicate_table` → `edit_table_lifecycle`.

## 2. Parity

| Test | Routes | Tier |
|---|---|---|
| `numbers_table_duplicate_parity` | the first table of `basic.numbers` | B |
| `numbers_table_duplicate_corpus_parity` | the first table of each corpus fixture the host edits: 23 routes | B |

Tier B because the host draws the clones' object UUIDs at random and the
leaf derives them; this is the documented UUID policy.

## 3. Performance

`cost_probes::numbers_table_duplicate_cost_abba` (the first table of
`basic.numbers`; host, leaf, leaf, host):

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| table duplicate | 1,204,070,144 | 110,727,986 | **−90.8%** |

## 4. Gaps

| Item | Next step |
|---|---|
| sheet duplicate | builds on this route: one table duplicate per table, then a sheet-wide remap of cross-table edges, plus the other drawables (W2-A's engines) |
| no native evidence for a duplicated table | mac queue |

## 5. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa --all-features` | 0 |
| `cargo test --all-features -p litchi-numbers`, `-p litchi-numbers-wire` | 0 / 0 |
| `cargo test --all-features -p litchi-pages`, `-p litchi-keynote` | 0 / 0 |
| `cargo test -p litchi-ops --all-features --test numbers_unified` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal` / `numbers-full` | 0 / 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |

## 6. Lift into `litchi-iwa-table` (2026-10-27)

The format-neutral copy now lives in the shared table engine as
`litchi_iwa_table::structure::duplicate::clone_graph`, so the Pages and
Keynote table owners can wrap it. It covers the drawable, the model, the
owned storage, the CalculationEngine owner family, the formula-list UUID
rewrites and the clones' registry records. Both owners store the same TST
graph:

- A Pages `TST.WPTableInfoArchive` passes `info_path: &[1]`.
- A Numbers or Keynote table info passes `&[]`.

`litchi_numbers::package::table_lifecycle::duplicate` keeps only the
Numbers attachment:

- the sheet's drawable list;
- the sidebar node;
- the sheet's, the engine's and the moved drawable's references.

The parity results are unchanged.
