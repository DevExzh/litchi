# 0421a — W4-H: Keynote slide-table sort on the shared executor

- Status: merge-ready on branch `uo2/iwa-w4h` after
  [0421](0421-iwa-w4h-table-sort-and-pivot.md) (merged as `d6842ba024`).
- Date: 2026-11-02.
- `performance_claim`: the Keynote route costs 3.6% more instructions than
  its previous private owner, inside the 5% budget (§3).

Keynote's `execute_slide_table_sort_order[_to_rows]` and
`edit_slide_table_physical_sort[_to_rows]` (no host route; Keynote never had
one in `litchi-iwa`) now run the permutation and the row rewrite through
`litchi_iwa_table::sort`, the executor 0421 lifted from Pages. No public API
changes and no ledger row moves.

## 1. What moved, what stays

- **Deleted from `slide_table_physical_sort.rs`** (4,086 → about 3,420
  lines): the private permutation planner, the codec-driven tile, header
  bucket and UID-map rewrite with its staged-component helpers, and the
  helpers only they used.
- **Leaf hook.** `Hooks::invalidating_previews_same_tile()`: Keynote drops
  the three root previews (its natively verified rendering-invalidation
  policy) and admits only same-tile row moves (a cross-tile permutation is
  refused as `UnsupportedTopology`, as before). Pages and Numbers use
  `Hooks::keeping_previews()`.
- **Kept, because Keynote genuinely differs:** selection and physical
  admission through `slide_table_core` (locked tables, comment anchors,
  rich-text cells, formula registries, unknown mutable fields, wide-row and
  pre-BNC consistency), the strict key scan (string-table shape, padded
  offsets, storage version), and the post-publication verification of
  keys, UID maps, headers and member/object locality. The source scan is
  now computed once and handed to the verification instead of being
  rescanned.
- **Errors.** Engine refusals map onto Keynote's typed refusals
  (`UnsupportedFeature::{Formula, FormulaError, UnsupportedCell,
  UnsupportedSortKey, MergedCells, FilteredRows, ConditionalStyles,
  NonPositionalHiddenState}`, `UnsupportedTopology`); a mixed key column is
  located on Keynote's own scan for `MixedSortKeyKinds { row, rule }`.

## 2. Parity

The 26 integration tests of `crates/litchi-keynote/tests/slide_table_physical_sort.rs`
and every `litchi-keynote` test pass unchanged. A temporary digest test
(removed before commit) sorted the entire-table, selected-rows and numeric
fixtures with the previous owner and with the new one: the output packages
and the diagnostics are byte-identical.

## 3. Performance

Callgrind, `--toggle-collect=*execute_slide_table_sort_order*`, CPU 5, the
three digest sorts, release `slide_table_physical_sort` test binaries
(previous owner SHA-256
`56539a0dd043914ae97971f8bcabf033b793b0f09736cfa88000d5db2f582fdf`, new
`648ff61530d605fb01c2b03d60498efff42f0c71ddf87e3bfac1fc3b7d19a77d`):

| Route | Previous Ir | Shared executor Ir | New/previous |
|---|---:|---:|---:|
| Keynote Sort Now (3 sorts) | 13,153,709 | 13,632,341 | 1.036 |

## 4. Verification

The gate list of `.coord/iwa-w4h/gates.sh` runs on the merge of trunk;
its result is reported in the hand-off.
