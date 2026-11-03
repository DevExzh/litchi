# 0421 — W4-H: table sort and pivot categories

- Status: merge-ready on branch `uo2/iwa-w4h` (cut from trunk
  `d4121f4af7`), slices 1–4 below. The Keynote move onto the shared
  executor is the next slice of this record (0421a).
- Date: 2026-11-02.
- `performance_claim`: every migrated route except the persisted
  sort-order set uses fewer instructions than its host route; the set is
  measured against two host routes (§5).

Six Numbers `table.sort` rows and the Numbers `table.pivot` row migrate.
The Pages body-table sort executor moves into `litchi-iwa-table` as the
shared sort executor, and Numbers gets its **Sort Now** family on it.

## 1. Persisted sort rules (Numbers)

`litchi_numbers::Package::{table_sort_order, edit_table_sort_order}`
already existed, and the host's `table_sort_order`, `set_table_sort_order`
and `clear_table_sort_order` (test-only) delegate to them. A delegation is
no oracle, so two hidden `parity-host-oracle` seams give the harness
independent host routes:

- `__parity_table_sort_order_legacy` reads `TableModelArchive.sort_order`
  through the host package and the host's own codec read;
- `__parity_set_table_sort_order_prost` decodes the model with Prost,
  replaces `sort_order` (a clear keeps Numbers' empty-order marker when the
  table has one) and re-encodes it through the host package writer.

Both seams are ledger rows retired as parity-only oracles. The read is
tier A; set and clear are tier B (the Prost re-encoding moves the
rewritten field, the leaf keeps the model's field order).

## 2. The shared executor: `litchi_iwa_table::sort`

`crates/litchi-pages/src/package/body_table_sort_execution.rs` (record
0371) did the rewrite on Pages' own graph helpers. Its payload rewrites
moved unchanged into `litchi_iwa_table::sort::execute`, now over the
structural working set and generic over `TableLeaf`:

- tile row records move to their destination tile and local index (cells
  keep their bytes); the row header buckets, the row UID map and the
  stroke layers follow; user-hidden row states are re-keyed so hidden rows
  keep their positions;
- the host's scalar subset and feature refusals (formula, error and
  unsupported cells; mixed key kinds; filters, grouping, pivots, spill
  state; conditional styles; no row UID map) are the engine's typed
  `Unsupported`, which Pages and Numbers re-export;
- the leaf resolves the table (`Target`), reads the persisted order and
  checks its scope and the merges; the engine asks for the user-hidden rows
  through a closure, only when the table declares hidden axes and rows
  move, so a leaf whose hidden-axes reader refuses a graph reports the
  executor's content refusal first;
- `Hooks` carry the per-leaf policies: keep or drop the root previews, and
  same-tile moves only (for Keynote, 0421a).

Fixes found by the Numbers cases, all inside the engine:

- Numbers rows keep a fixed-width cell offset table; slots past the last
  column must be empty (the Pages port refused every native Numbers row);
- a row-header bucket holds 65,536 rows, as in the host and in
  `litchi_numbers_wire::table_topology::storage`; the Pages port used 256
  and would refuse a sort that moves a header past row 255;
- `structure::metadata::apply` no longer decodes the package-metadata
  registry for an empty transition with no allocation. Every structural
  family that adds or removes nothing (sort, topology insert) is cheaper;
  outputs are unchanged.

Pages keeps its selector, scope, merge and hidden-axes checks and its
verification; its `Unsupported` is the engine's, and its `LimitKind` is the
working set's (`PackageBytes`, `PayloadBytes`, `Identifiers`). The Pages
sort parity cases stay tier A, unchanged.

## 3. Numbers sort execution

`litchi_numbers::table::sort::execution` provides
`Package::{edit_table_sort_execution, apply_table_sort_execution}` with
`Edit::{apply, apply_to_rows}`, on the structural family macro (exact-source
patch and inverse, no-op when the rows are already in order, typed
refusals). After publication the facade verifies that the persisted order
and the user-hidden axes read back unchanged. Root previews are kept, as
the host keeps them.

## 4. Numbers pivot categories

`litchi_numbers::Package::pivot_categories` returns
`table::pivot::Category { label, group_level, aggregate_type }` in the
host's order. The leaf walk is a bounded wire read of the calculation
engine's group-bys, aggregators and group nodes, with the host's semantics:
enabled group-bys only, one formula owner per owner kind, inline or
referenced aggregators and group nodes, a visited set for child
references, a node ceiling, the label rules (string, number, Boolean,
date, `Grand Total`), and categories that two group-bys describe
differently dropped as ambiguous. Unlike the host's `PivotCategoryInfo`,
`Category` carries no native group-by, column or group UUID.

No tracked native workbook holds a pivot table. The parity case adds
synthetic group-by graphs to `basic.numbers`' calculation engine (inline
and referenced trees with a reference cycle, two aggregated columns, a
disabled group-by, an ambiguous owner kind, an ambiguous category) and
compares both reads, tier A; the route is macOS-queued.

## 5. Parity and performance

| Test (`crates/litchi-iwa/tests/parity/`) | Tier |
|---|---|
| `numbers_table_sort_order_read_parity` (unsorted, three orders, cleared; 3 workbooks) | A |
| `numbers_table_sort_order_set_parity` (3 orders × 3 workbooks) | B |
| `numbers_table_sort_order_clear_parity` | B |
| `numbers_table_sort_apply_parity` (number and text keys on 3 workbooks; host-set borders; a hidden row) | A |
| `numbers_table_sort_apply_to_rows_parity` | A |
| `numbers_table_sort_refusals_match_host` (scope mismatch, no order, formula cell; leaf no-op) | — |
| `numbers_pivot_categories_read_parity` | A |
| `pages_body_table_sort_apply{,_to_rows}_parity` (unchanged) | A |

Callgrind, release parity binary SHA-256
`cf08ad813302fec6847dbcefb2615f1e033989dcb3d31c2456cb03bb939ab0bb`, CPU 5,
ABBA (`--toggle-collect=*probe_host*` / `*probe_leaf*`):

| Route (probe) | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| Numbers Sort Now (`numbers_sort_apply_cost_abba`) | 11,114,228 | 6,237,062 | 0.561 |
| Pages Sort Now after the lift (`table_sort_apply_cost_abba`) | 3,402,212,551 | 66,271,167 | 0.019 |
| Numbers sort-order set, host route step by step (`numbers_sort_order_host_route_cost_abba`) | 115,281,350 | 72,931,531 | 0.633 |
| Numbers sort-order set, against the Prost oracle seam (`numbers_sort_order_cost_abba`) | 42,918,239 | 72,887,996 | 1.698 |
| Numbers pivot categories on opened packages (`numbers_pivot_categories_cost_abba`) | 2,132,883 | 474,956 | 0.223 |
| Pages topology insert row after the registry skip (`pages_body_table_insert_row_cost_abba`) | 2,581,458,648 | 112,458,908 | 0.044 |

The host's `set_table_sort_order` is test-only and delegates to the leaf,
so its real cost is the leaf's plus the host round trip; the step-by-step
probe reproduces that route and the leaf costs 37% less. The Prost seam is
a test oracle, not a host route: it re-encodes one model message without
the leaf's codec preflight, locality verification and reopen, and it is
70% cheaper than the leaf. That miss is recorded, not hidden; the leaf
sort-order family is unchanged by this record.

Pages Sort Now cost 54,905,570 Ir in record 0371 (another binary) on
Pages' own graph helpers; on the shared working set it costs 66,271,167
(+21%), still 1.9% of the host route.

Ledger: `table_sort_order`, `set_table_sort_order`,
`clear_table_sort_order`, `apply_table_sort_order`,
`apply_table_sort_order_to_rows` and `pivot_categories` migrate; the two
seams are retired. Four macOS queue rows. The glossary registers
`litchi_numbers::table::sort::execution` and `table::pivot`, and the Pages
execution row names the shared executor.

## 6. Gaps

- Keynote's `slide_table_physical_sort.rs` still has its own copy of the
  executor (0421a).
- `Category::aggregate_type` is Numbers' summary-function code; a typed
  summary-function enum needs a native mapping of the codes.
- The leaf formula author (`litchi_numbers_wire::formula_cache::author`)
  refuses pivot-category references, so a category cannot yet be used in
  a leaf formula as the host's `FormulaExpression::PivotCategory` can.

## 7. Verification

The gate list of `.coord/iwa-w4h/gates.sh` runs on the merge of trunk; its
result is reported in the hand-off.
