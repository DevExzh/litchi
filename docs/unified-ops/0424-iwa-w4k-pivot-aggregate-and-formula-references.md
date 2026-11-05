# 0424 — W4-K: typed pivot aggregate and pivot-category formula references

- Status: merge-ready on branch `uo2/iwa-w4k` (cut from trunk
  `cf6bca9146`).
- Date: 2026-11-03.
- `performance_claim`: the leaf formula route with a pivot-category
  reference uses 55.6% fewer instructions than the host route (§4).

Record [0421](0421-iwa-w4h-table-sort-and-pivot.md) left two gaps on
Numbers pivot tables (§6 there). This record closes both.

## 1. Typed summary function: `table::pivot::Aggregate`

`Category::aggregate_type()` returned Numbers' raw summary-function code
(`TST.ColumnAggregateArchive.agg_type`). The 0351 contract forbids native
values in supported facades, so the accessor is deleted (no alias) and
replaced by `Category::aggregate() -> Aggregate`.

The brief asked for a code mapping taken from the host code and the
protobuf definitions, with evidence for each code. None of the sources
below names any code:

| Source | What it says about the code |
|---|---|
| `litchi-iwa` (`numbers/editor/model.rs::formula_pivot_categories`, `numbers/formula.rs`) | copies `agg_type` into `CategoryReferenceArchive.aggregate_type` and compares it for equality; no names |
| `TSTArchives.proto` (`ColumnAggregateArchive.agg_type`), `TSCEArchives.proto` (`CategoryReferenceArchive.aggregate_type`), `TSTCommandArchives.proto` (`CommandCategoryChangeSummaryAggregateType`) | bare `uint32`; no enum |
| every tracked `.numbers`, `.pages` and `.key` fixture | group-bys are all disabled and carry no `column_agg_type` (scan of every `TST.GroupByArchive`, message 6373) |
| git history, records, the host's tests | codes 1, 2 and 4 in synthetic tests, unnamed |

With the coordinator's agreement, no code is named by guesswork:

- `Aggregate` is a `#[non_exhaustive]` enum (`Clone`, `Copy`, `Eq`,
  `Hash`, `Debug`). Today its only variant is
  `Aggregate::Unnamed(UnnamedAggregate)`.
- `UnnamedAggregate` compares by its native code but never exposes it.
  Its `Debug` output prints `UnnamedAggregate { .. }`.
- Two categories with the same summary function compare equal. That is
  what a caller can rely on until the names land.
- The code stays private. A formula author uses it to write a category
  reference byte-exactly.
- Unknown codes need no `Other(u32)` variant: no leaf route writes a
  code that it did not read from the same model.
- **The named variants (`Sum`, `Average`, `Count`, …) are added once the
  macOS evidence lands.** The 0363 mac queue has a W4-K row for it: build
  a native pivot with one value column per "Summarize by" function and
  record each column's `agg_type`. Because the enum is non-exhaustive,
  adding the variants is not a breaking change.

`Category` also keeps its native identity (the group-by, column and
group-node UUIDs) privately, so a formula can reference it. Its `Debug`
output omits that identity.

**Parity: tier A is kept.** `numbers_pivot_categories_read_parity`
compares, in order, the label, the group level and the summary-function
equivalence class (each code numbered by its first occurrence) of the host
and leaf reads. It also asserts that the synthetic codes 1, 2 and 4 form
three distinct classes.

## 2. Pivot-category references in formulas

The host's `set_formula`/`set_formula_with_cached_value` accept
`FormulaExpression::PivotCategory`, and the leaf refused it. This was a
capability loss. Now:

- **Vocabulary.** `litchi_numbers::formula::Expression::pivot_category(&Category)`
  builds an absolute reference from a category read by
  `Package::pivot_categories`. No native identifier crosses the facade.
- **Resolution, with the host's validation.** At commit the author reads
  the edited package's categories through the same bounded walk as
  `pivot_categories`. The walk runs only when the batch holds a
  reference, and its lookups are charged as formula work. Each reference
  then follows the host's `append_pivot_category_reference`:
  - if the category is absent, the edit is refused with
    `Refusal::PivotCategoryNotFound` (`RefusalKind::NotFound`);
  - if the summary function or group level differs, it is refused with
    `Refusal::PivotCategoryChanged` (`InvalidValue`).

  In both cases the host refuses as well.
- **Codec.** `litchi_iwa_protos::numbers_formula_codec::FormulaWriteNode::CategoryReference(FormulaWriteCategoryReference)`
  writes the host's exact node:
  - `CATEGORY_REF_NODE` with the group-by, column and absolute group
    UUIDs, the code and level;
  - preserve flags `(false, true, false, false)`;
  - `agg_index_level = u16::MAX`;
  - the formula's translation flags (only `returns_percent_formatted`
    set).

  It streams two precedents: the grouping-columns cell and the
  aggregate cell of the group-by's formula owner. The strict readback and
  the Buffa cross-check cover both the node and the root flags. The unit
  test `writer_category_reference_matches_prost_host_shape` compares the
  output with the Prost encoding of the host's archive.
- **Readback.** The resolved dependency reader tolerates a strictly
  validated category node as a non-evaluable operand, an
  `UnsupportedLocal` of type 66. Every other cross-owner node is still
  refused.
- **Formula cache** (`litchi-numbers-wire`). The logical graph may name
  precedents in registered owners without table geometry. For a formula
  that holds a category reference, edges from those owners are taken from
  the dependency records, because the AST does not name them. No table
  edit writes those cells, and the formula is never evaluated.
- **Cache policy.** A supplied cache is stored as given. Without one, the
  formula keeps the host's zero placeholder until Numbers recalculates.
  (An evaluable formula without a supplied cache is still evaluated, as in
  record 0376c.)

Keynote and Pages are unchanged. Their host routes share
`set_attached_table_formula`, but Keynote and Pages have no pivot tables,
so the host's category map is always empty there and every reference
fails. `formula_cache::author` keeps refusing the variant, and no
capability is lost.

### Fixture

The 0421 synthetic group-bys named a synthetic formula owner (kind 40)
that the engine's dependency tracker did not register. The leaf, unlike
the host, requires every precedent owner to be registered. The group-bys
now name `basic.numbers`' registered category owner, which is the owner
that its native disabled group-by names: owner kind 8, internal
identifier 13. The ambiguous-owner case (kind 41, two owners) is
unchanged. `pivot_fixture_with(code)` varies the first column's code for
the refusal case.

## 3. Parity (`crates/litchi-iwa/tests/parity/numbers_pivot_formula_cases.rs`)

| Test | Route | Tier | Result |
|---|---|---|---|
| `numbers_pivot_formula_author_parity` | C3 = North / 2, cache 21.5, inline group-by | B | pass; formula archive and C3 dependency edges byte-identical (`internal_owner_id_for_edge: [13, 13]`) |
| `numbers_pivot_formula_author_without_cache_parity` | same, no cache (host zero placeholder) | B | pass; same byte identities |
| `numbers_pivot_formula_author_referenced_tree_parity` | a category of the referenced tree (separate aggregator and group-node objects), another summary function | B | pass; same byte identities |
| `numbers_pivot_formula_missing_category_is_refused_like_host` | a category read from the pivot fixture, authored into `basic.numbers` | — | leaf `PivotCategoryNotFound`; host errors |
| `numbers_pivot_formula_changed_category_is_refused_like_host` | the same category against a model with another code | — | leaf `PivotCategoryChanged`; host errors |
| `numbers_pivot_categories_read_parity` (updated) | read | A | pass |

The tier-B differences are the ones documented for formula authoring:
- the leaf invalidates the root previews;
- the leaf's structural metadata;
- the leaf's deterministic UUIDs where the host's are random;
- the ZIP member list.

No tracked native workbook holds a pivot table. The mac queue therefore
gains two W4-K rows: the formula route, and the naming of the aggregate
codes. The ledger rows are unchanged. `pivot_categories` (tier A) and
`set_formula` (tier B) were already migrated, and the pivot reference is
a capability of the migrated `set_formula` route.
`iwa_exit_ledger.py --check` passes.

## 4. Performance

Callgrind measured the release parity binary, built with
`CARGO_PROFILE_RELEASE_DEBUG=1` from this branch before the record
commit. The binary is `target/release/deps/parity-a7f576456803efd8`, with
SHA-256
`6ee16ad71de27f3498836062104f54590fe6a15a60de69980186ee71789bdd31`. It
was not committed. The runs pinned CPU 5 and used ABBA order, with
`--toggle-collect` on the route symbols.

| Route (probe) | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| Numbers formula with a pivot-category reference, cached (`numbers_pivot_formula_author_cost_abba`, `*pivot_host_route*` / `*pivot_leaf_route*`) | 193,578,119 | 85,893,976 | 0.444 (−55.6%) |
| Numbers formula without a pivot reference, regression check (`numbers_table_formula_author_cost_abba`) | 190,138,577 | 84,167,355 | 0.443 (−55.7%) |

The budget is ≤ +5% over the host; the pivot route is 55.6% under it.
The non-pivot route was measured at 84,009,135 Ir in record 0376c with
another binary. It now measures 84,167,355 Ir (+0.19%). The category scan
added to the writer and the cache engine's tableless-owner index are the
only new work on that route.

Reproduce:

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa \
  --all-features --test parity --no-run
taskset -c 5 valgrind --tool=callgrind --toggle-collect='*pivot_host_route*' \
  target/release/deps/parity-<hash> --exact --ignored --test-threads=1 \
  numbers_pivot_formula_cases::numbers_pivot_formula_author_cost_abba
# and again with '*pivot_leaf_route*'
```

## 5. Gaps

- The `Aggregate` names wait for the macOS run (mac-queue row W4-K).
- A pivot-category formula is never evaluated by the leaf cache engine.
  A supplied cache is kept; otherwise the host's zero placeholder is
  written. That is the host's own behaviour.
- Cell edits on tables that own pivot or category state: the host allows
  them (a host call on the same fixtures succeeds), so record
  [0424a](0424a-iwa-w4k-cell-edits-on-pivot-and-category-tables.md)
  admits them in the leaf.

## 6. Verification

The gate list in `(deleted-coord-artifact)` runs on the merge of trunk.
Its results are reported in the hand-off.
