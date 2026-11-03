# 0424a — W4-K: cell edits on pivot and category tables

- Status: merge-ready on branch `uo2/iwa-w4k`, after
  [0424](0424-iwa-w4k-pivot-aggregate-and-formula-references.md) (merged
  at `a198d9b2f5`).
- Date: 2026-11-03.
- `performance_claim`: the leaf cell route on an active-category table
  uses 55.8% fewer instructions than the host route (§3).

## 1. The host allows them; the leaf refused them

`litchi_numbers::Package::edit_table_cells` refused a table that owns
pivot or active category state. The host's Numbers cell writes
(`set_formula`, `set_formula_with_cached_value`, the routes that
`edit_table_cells` replaces) never check those owners. The coordinator
asked for evidence: a host call on the same fixture. I ran a probe on
`basic.numbers` with three fixture variants. Each one rewrites only the
first table's own native state:

| Fixture | Host `set_formula_with_cached_value` | Leaf before | Leaf now |
|---|---|---|---|
| active category: the native group-by `904710` that the table's category owner (model field 86) names is enabled | `Ok` | `Refused(UnsupportedDependency { kind: Category })` | written |
| active deprecated category: the inline group-by of the model's field 81 is enabled | `Ok` | `Refused(UnsupportedDependency { kind: Category })`, from the same rule | written |
| pivot table: the model names a `TST.PivotOwnerArchive` (field 85, declared) | `Ok` | `InvalidSource`: the header dependency check refused every pivot owner, and the cell route mapped that to `InvalidSource` | written |

This was a capability loss, and this slice closes it.

## 2. Change (`litchi-numbers`)

- `table_cell_edit::resolve` no longer refuses a model with a pivot owner
  or an active category owner.
  - The category owner and its group-bys are still resolved and
    validated.
  - A malformed deprecated owner is still refused.
  - A pivot filter set must still be inert (unchanged).
- `table_headers::dependencies::validate_cell_dependencies` is the cell
  route's dependency check. A pivot owner alone does not refuse an edit
  that changes no header or section count. Every count-change rule still
  applies, and the header family keeps `validate_dependencies` unchanged.
- The leaf leaves the group-by and pivot state untouched, as the host
  does. Numbers recomputes it.
- No public API changes. `DependencyKind::{Pivot, Category}` remain for
  the refusals that are left (a non-inert pivot filter set, a malformed
  deprecated owner).

## 3. Parity and performance

The cases are in `crates/litchi-iwa/tests/parity/numbers_pivot_owner_cell_cases.rs`.
Each sets C3 = B3 + 1 with cache 43: host `set_formula_with_cached_value`
against leaf `set_formula_cached` with an inverse readback.

| Test | Tier | Result |
|---|---|---|
| `numbers_active_category_table_cell_parity` | B | pass |
| `numbers_deprecated_category_table_cell_parity` | B | pass |
| `numbers_pivot_owner_table_cell_parity` | B | pass |
| `numbers_owned_table_value_write_is_admitted` (value write on all three; no host value setter remains) | — | pass |

The tier-B differences are the documented formula-authoring ones:
- the leaf invalidates the root previews;
- the leaf's structural metadata;
- the leaf's deterministic UUIDs where the host's are random;
- the ZIP member list.

A macOS queue row (W4-K) covers the route, because no native workbook
holds an active category or pivot table.

Callgrind measured the release parity binary, which was not committed
(SHA-256
`45fbabef95223835b0fad982c946032a0315fd54cf89ec68e6637b73b6c1e19e`). The
runs pinned CPU 5 and used ABBA order, with `--toggle-collect` on
`*owned_host_route*` / `*owned_leaf_route*` in
`numbers_owned_table_cell_cost_abba`.

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| cached formula on the active-category table | 190,165,332 | 84,133,078 | 0.442 (−55.8%) |

## 4. Gaps

- A cell edit leaves the pivot or category aggregates for Numbers to
  recompute, as the host does. Native acceptance is queued.
