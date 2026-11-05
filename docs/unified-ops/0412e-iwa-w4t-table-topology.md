# 0412e — W4-T: Pages and Keynote table topology on the shared engine

- Status: merge-ready on branch `uo2/iwa-w4t` after
  [0412d](0412d-iwa-w4t-numbers-comment-threads.md).
- Date: 2026-11-02.
- `performance_claim`: every route uses fewer instructions than its host
  route (§3).

The four Pages and four Keynote `table.topology` rows (insert and remove a
row or a column) are migrated. Every case is tier A: the leaf output is
byte-identical to the host's.

## 1. The shared engine

`litchi_iwa_table::topology` is the Numbers table-topology rewrite
(`litchi-numbers/src/package/table_topology.rs`, records 0309, 0318, 0329,
0370b), lifted so that all three formats run it. It is generic over
`TableLeaf` and runs on the shared structural working set.

- **What moved.** The rewrite covers storage tiles, header buckets, the
  row/column UID map, border layers, the model's counts and the formula
  owner's extents. It also includes the feature refusals (filters,
  categories, pivots, merges, sort rules, hidden or bordered axes) and the
  release of dropped cells' list keys, strings and formulas. The formula
  coordinate shift moved too (`topology/shift.rs`, which now reads a
  `dyn TableSource`). The BNC limits now come from
  `TableSource::archive_limits`, so the Numbers-only `numbers_wire_limits`
  remains only in a test.
- **Entry points.** `resolve`, `read_geometry`, `check` (the
  deletion-safe preflight), `removed_cells`, `table_formula_inert` and
  `publish(source, target, step, before, Publish { drop_cells,
  fit_drawable })`. The engine's typed `Error`/`Refusal` map into each
  leaf's family error. `Unsupported` is shared and re-exported by all
  three leaves.
- **New: `fit_drawable`.** Keynote's host moves the slide drawable bounds
  with the table (`edit_slide_table_topology` sets the `TableInfo`
  geometry size to `table_size_points_in_package`). The engine computes the
  same sum of explicit header sizes and model defaults and patches
  `TSD.DrawableArchive.geometry.size`.
- **Numbers.** Numbers keeps its selectors, the comment and cell clears of
  a removed axis and the cell-by-cell verification, and calls the engine
  for everything else. All `litchi-numbers` tests, the host's
  `numbers_topology_formula_differential`, and the Numbers topology parity
  and corpus cases pass unchanged.

## 2. Facades

`litchi_pages::table::topology` and `litchi_keynote::slide::table::topology`
provide `Edit::{insert_row, insert_column, remove_row, remove_column}` (several
operations per edit, applied in order), `Package::{edit_body_table_topology,
apply_body_table_topology}` and `Package::{edit_slide_table_topology,
apply_slide_table_topology}`, with an exact-source patch and inverse and
typed refusals. Before the row or column goes, its comments are cleared
through the 0412c comment family, threads included. Its cells are dropped
in the structural pass. A formula that reads a dropped cell reads `#REF!`.
Calculation state that the shift does not rewrite is refused as
`Unsupported::Formulas`, and so is any state the host refuses.

## 3. Parity and performance

| Test | Tier |
|---|---|
| `pages_body_table_{insert,remove}_{row,column}_parity` (`body-table-catalog-native.pages`) | A |
| `pages_body_table_remove_commented_row_parity` (the native reply-bearing comment's row) | A |
| `keynote_slide_table_{insert,remove}_{row,column}_parity` (`table-discovery.key`) | A |
| `pages_body_table_topology_formula_state_refused_like_host` (`body-table-cells-native.pages`) | both routes refuse |
| `{pages_body,keynote_slide}_table_topology_refusals_are_typed` | — |

Callgrind, release parity binary SHA-256
`a0d77cdf2fdcce1c3448d943d15ffb2806c01face80e9433490082a7e46e716d`, CPU 5,
ABBA:

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| Pages insert row | 2,578,600,687 | 121,908,739 | 0.047 |
| Pages insert column | 2,578,914,762 | 121,787,809 | 0.047 |
| Pages remove row | 2,595,128,220 | 131,672,240 | 0.051 |
| Pages remove column | 2,608,407,643 | 131,674,723 | 0.050 |
| Keynote insert row | 692,857,715 | 453,090,154 | 0.654 |
| Keynote insert column | 692,977,482 | 453,255,453 | 0.654 |
| Keynote remove row | 708,764,080 | 467,833,367 | 0.660 |
| Keynote remove column | 730,151,142 | 467,972,453 | 0.641 |
| Numbers insert (`numbers_topology_insert_cost_abba`, after the lift) | 752,655,846 | 483,186,522 | 0.642 |
| Numbers remove (`numbers_topology_remove_cost_abba`, after the lift) | 972,291,500 | 557,680,947 | 0.574 |

Ledger: 4 Pages and 4 Keynote rows migrate. Two native queue rows. The
API glossary registers both families. Its Pages and Keynote comment rows
now describe the 0412c thread family.

## 4. Verification

The gate script `(deleted-coord-artifact)` runs on the merge of trunk; its
result is reported in the hand-off.
