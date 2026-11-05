# 0432 — W5N: native pivot and volatile evidence

- Status: merge-ready on branch `uo2/iwa-w5n` (cut from trunk `52f3d60714`).
- Date: 2026-11-04.
- `performance_claim`: every measured route stays within the +5% budget against the previous leaf; the largest change is `pivot_categories`, +2.2% (§7).

The owner authored two documents by hand in Numbers 14.4 on macOS 26.5.2. They
are the first native pivot and the first native volatile rule the program has.
They disproved several assumptions of records 0376b, 0421, 0424 and 0424a.

## 1. Fixtures

| Fixture (`test-data/iwork/numbers/`) | SHA-256 | Contents |
|---|---|---|
| `pivot-all-summary-functions-native.numbers` | `e0830bb9f773f755621220263a767af0e864072fe5a37a84502ccaa4900f9baa` | "Table 1 Pivot" pivots "Table 1": rows Region, columns Product, values Amount with 13 functions |
| `volatile-highlight-native.numbers` (authored as `volatile-is-today-native.numbers`) | `ab580f5b4dc7f4977ff03adec4f530352ab9c84d9af72c74157b643cf8df7062` | A2 today, A3 yesterday, A4 tomorrow (2026-11-04, Asia/Shanghai), rule "Date is today" with the preset "Red Fill" on A2:A4 |

Each has a reduced receipt next to it (`*-receipt.json`). Screenshots, the
accessibility dumps and the full receipt, which holds local paths, stay outside
the repository.

## 2. Summary-function codes

| Values position | Numbers function | Header | `agg_type` | `Aggregate` |
|---|---|---|---:|---|
| 1 | Sum | Amount (Sum) | 2 | `Sum` |
| 2 | Average | Amount (Average) | 18 | `Average` |
| 3 | Count (All) | Amount (Count All) | 1 | `CountAll` |
| 4 | Maximum | Amount (Max.) | 5 | `Maximum` |
| 5 | Minimum | Amount (Min.) | 4 | `Minimum` |
| 6 | Count | Amount (Count) | 13 | `Count` |
| 7 | Count (Unique) | Amount (Unique) | 14 | `CountUnique` |
| 8 | Median | Amount (Median) | 19 | `Median` |
| 9 | Product | Amount (Product) | 21 | `Product` |
| 10 | STDEV | Amount (STDEV) | 26 | `StandardDeviation` |
| 11 | STDEVP | Amount (STDEVP) | 28 | `StandardDeviationPopulation` |
| 12 | VAR | Amount (VAR) | 30 | `Variance` |
| 13 | VARP | Amount (VARP) | 32 | `VariancePopulation` |

The Values order is the authority. Three other sources agree with it on the
Widget/North row, whose amounts are 10 and 70: the column's header, the code in
that column's native formula, and the cached result
(`native_pivot_codes_agree_with_headers_formulas_and_cached_results`). The
amounts are all numbers, so the cached results cannot tell `Count` from
`CountAll`, nor `CountUnique` from either; for those three the order, the header
and the formula code are the evidence. Any other code stays
`Aggregate::Unnamed`.

## 3. What the native pivot disproved

| Assumption | Record | What Numbers writes | Fix |
|---|---|---|---|
| A category is one (group node, column) with one summary function | 0421, 0424 | One aggregator per source column and one `column_agg_type` per function, all with the same `column_uid`. The host and the old leaf read only the first, so every category read as Sum: 13 categories instead of 169 | A category is now (group node, column, function), in value order. The host keeps its first-function read |
| The engine's owner map and its owner list have the same length | `table_cell_edit::resolve` | 43 map entries for 36 owner objects (internal 22–26, 49, 50 have no object) | Every registered owner must have its entry; extras pass, as `litchi_numbers_wire::formula_graph` already read them |
| A table's owner has no UUID references | 0424a; resolver and cache engine | The pivot table's owner lists each category formula under its group-by and group UUIDs (field 14) | Admitted when well formed. The formula set is still validated node by node |
| Every other owner's record tile sits at origin (0, 0) | `formula_dependencies` | A dependent owner keeps a tile at column 32 | Only the selected owner's tile origins are checked |
| A foreign table owner with records needs geometry | 0301, cache engine | The pivot's body has records, on another table | Admitted when the edit cannot reach it (§3.1); refused as before when it can |
| Every precedent owner of a workbook with other tables is a table | `extend_formula_referenced_paths` | A category formula's precedent owner is a group-by's | Such an owner adds no table path |
| Group-by owner kinds 8 or 40 | 0421, 0424 synthetic fixtures | Kinds 205 and 206, referenced from a category owner (6372) | No code change: the leaf resolves any kind naming exactly one owner. Only the fixtures assumed it |
| "Fixed-width offset tables" | 0421 | Confirmed: 255 two-byte slots on the source table's and the pivot's rows | None |

The owner-map check ran for every cell edit, so it alone made the whole workbook
uneditable, on any table. The host wrote formulas on both tables.

### 3.1 "Numbers recomputes the state"

Records 0424 and 0424a say a cell edit leaves the pivot's state for Numbers to
recompute. The native graph shows a narrower picture:

- Numbers keeps a hidden model, "Table 1 as Pivot Source Table" (13 × 3, on no
  sheet), as the pivot's snapshot of the source.
- The group-by owners (internal 102, 103) depend on the snapshot's owners (27,
  28, 34). The pivot's owner (37) depends only on 102.
- No dependency path leads from the visible source table's owner (6) to 37. The
  only record naming owner 6 belongs to owner 47, which nothing depends on.

So a source edit dirties nothing in the pivot, which is why the leaf can leave
it alone. **This is read from the structure, not observed.** It suggests that
Numbers shows the old pivot until a Refresh; the Mac rows of §6 check that.

### 3.2 Category formulas compared with Numbers

`native_pivot_category_formula_matches_what_numbers_wrote` writes a category
formula over the native body cell B5. The category reference and the dependency
edges (owner 102, rows 0, 0, the grouping and aggregate columns) equal what
Numbers wrote there.

One difference remains: **Numbers lists each category formula in its owner's
`uuid_references`; the leaf and the host do not add a new formula to it, nor
take a replaced one out.** The leaf keeps the host's behaviour here because no
native run yet shows whether it matters. `W5N-pivot-category-formula-*` (§6)
are the first evidence.

### 3.3 Cell edits

| Edit on the native workbook | Host | Leaf before | Leaf now |
|---|---|---|---|
| value on the source table | no host value setter | `InvalidSource` | written, pivot objects byte-identical, inverse restores the source |
| category formula on the source table, all 13 functions of one group | `Ok` (Sum only) | `InvalidSource` | written |
| category formula on a pivot body cell | `Ok` | `InvalidSource` | written |
| value over a pivot body cell | no host value setter | `InvalidSource` | **refused**: `UnsupportedDependency { kind: FormulaCache }`, nothing published |

The last row is a replacement of a formula by a scalar, which this route's cache
plan refuses (`plan.removals`). I did not check whether the refusal is limited to
pivots, nor whether Numbers' UI lets a user type into a pivot body cell; neither
the receipt nor this record's tests cover that. The host has no scalar setter,
so no host capability is lost.

### 3.4 Read limits left

The leaf reads the stored grid, 5 × 27: Widget before Gadget, no Grand Total
row or columns. Numbers displays 6 × 40 with Gadget first (columns B:N, O:AA,
AB:AN). A second tile holds 13-cell rows at columns past the stored 27; I did
not establish which model owns it. `At` addresses stored coordinates, so a
caller cannot map a displayed pivot address onto a cell. Two categories can also
share a label and level ("North" under each product) with nothing public to tell
them apart.

## 4. What the native volatile file disproved

Record 0376b said every Numbers-saved table lacks the volatile family, and the
mac-queue row repeats it. **The family is not missing.** `basic.numbers` already
registers the same 13 owners, with sizes identical to the native file's except
the table owner and the conditional-style owner. Only the conditional-style
owner has no volatile-time set and no record tile until its first volatile rule.

Fixes:

- `conditional_volatile::ensure` now builds the set and the tile on that first
  rule, as Numbers does, instead of `IncompleteFamily`. All 12 other owners stay
  byte-identical (`first_volatile_rule_on_a_numbers_saved_table_builds_what_numbers_builds`).
  The host still fails on these tables, so the leaf is ahead of it;
  `numbers_native_conditional_highlight_volatile_host_fails_leaf_registers`
  pins both. Keynote and Pages share this planner but keep refusing, as the host
  does (`TableLeaf::CREATES_FIRST_VOLATILE_TILE`): they have no native evidence.
- The reader refused Numbers' own rule, so no leaf route read a Numbers-authored
  highlight. The preset "Red Fill" text style repeats its font colour as a text
  fill (`tsd_fill`, field 46), read as an unknown override. It now reads when the
  fill equals the font colour.

### 4.1 Leaf planner against the native family, field by field

Leaf output: a host `Builder` document, 22 × 7, A2:A4 "is today" set one cell at
a time. Native: the fixture.

| Item | Native | Leaf | Whose |
|---|---|---|---|
| Owner kinds | 1, 3–12, 35, 200 | same | — |
| Rule predicate AST | | identical apart from the cell's owner UUID | — |
| Conditional-style record of each cell | edges `(0, 2, style)`, `(row, column, table)` | same | — |
| Conditional-style tile | 3 records, origin (0, 0) | same | — |
| Kind 200 tile | 2 records | same | — |
| Volatile-time set | one range, rows 1–3 | three single rows | host's merge; removal handles both (`clearing_inside_a_native_range_splits_it`) |
| Kind 8 records (inline and tile) | columns 0, 3, 7 | columns 0, 1, 2, 3, 7, the extra two empty | host's |
| Formula count | 8 = 3 + 3 + 2 records | 10 = 3 + 5 + 2 records | follows from the row above; each side counts its own records |
| Selected table owner's tile | an empty tile | none | Numbers-saved documents only; not family |
| Owner internal ids and object order | 6–16, 20, 21 | another order | host's; ids are local to the file |

The last four are the host's shape. Its packages passed Numbers 14.4 (queue row
W1-F, P0389), so the leaf keeps them for tier-B parity rather than following the
native shape. The first-rule path of §4 matches the native owner shape on a
Numbers-saved table, whose kind 8 records are Numbers' own.

### 4.2 Rule read and edits on the native file

`crates/litchi-numbers/tests/native_volatile_highlight.rs` reads the rule on
A2:A4 with its native fill and text colour, adds a rule on the native family,
clears it, and clears inside Numbers' range.

## 5. Mac queue

- W4-K `Aggregate`: pass, on the native fixture.
- W1-F native-created volatile: **partial**. The first file is delivered, with
  its rule on A2:A4 rather than B2. `volatile-highlight-native-second.numbers`
  and the visual check that the leaf outputs highlight were not delivered.
- No other row changes. Row W1-F's route text still says Numbers-saved tables
  stay refused; §4 supersedes that, and the coordinator owns that row's text.
- **Pages identity row (not changed).** In the owner's manual check, both
  documents opened alone and together with no conflict prompt. The row's "edit
  and save" was not performed. The receipt marks that task `fail` only because a
  screenshot tool errored after both documents were open, not because of a Pages
  dialog or crash.

## 6. Mac kit additions

These fixes need native proof. The packages and lists are in
`.coord/iwa-w5n/mac-additions/` (`additions.tsv` for `control_kit.py add`,
`additions-queue.tsv` in the `queue.tsv` format). None is committed.

| Package | Question for Numbers |
|---|---|
| `W5N-pivot-native-control` | unedited positive control |
| `W5N-pivot-source-cell-edit` | C2 = 55: does it open, and does the pivot wait for a Refresh (§3.1)? |
| `W5N-pivot-category-formula-on-source` | C2 = Grand Total Average / 2: does it recalculate to 32.5 without `uuid_references` (§3.2)? |
| `W5N-pivot-category-formula-on-pivot-body` | B5 = another function's category: does Numbers accept or rebuild it? |
| `W5N-volatile-native-control` | unedited positive control |
| `W5N-volatile-native-add-tomorrow` | A4 "is tomorrow" replaces the native rule |
| `W5N-volatile-native-clear-middle` | A3 cleared inside Numbers' range |
| `W5N-volatile-first-rule-numbers-saved` | `basic.numbers`, B2 "is today": the first rule on a Numbers-saved table (§4) |

## 7. Performance

Callgrind on the release parity binaries of this branch and of `52f3d60714`
(`CARGO_PROFILE_RELEASE_DEBUG=1`, not committed), pinned to CPU 5, in ABBA order,
with `--toggle-collect` on the route. The budget is +5% against the previous leaf.

| Route (probe, toggle) | Base Ir (A, A) | Branch Ir (B, B) | Change |
|---|---:|---:|---:|
| `pivot_categories` (`numbers_pivot_categories_cost_abba`, `*Package*pivot_categories*`) | 413,354; 413,173 | 421,725; 423,403 | +2.2% |
| cached formula on an active-category table (`numbers_owned_table_cell_cost_abba`, `*owned_leaf_route*`) | 84,130,602; 84,126,102 | 84,128,818; 84,118,790 | −0.01% |
| formula with a category reference (`numbers_pivot_formula_author_cost_abba`, `*pivot_leaf_route*`) | 85,888,166; 85,908,577 | 85,898,763; 85,872,417 | −0.01% |
| formula without one (`numbers_table_formula_author_cost_abba`, `*leaf_cost_route*`) | 84,145,166; 84,149,294 | 84,136,900; 84,122,646 | −0.02% |

The volatile first-rule path is new, so it has no earlier cost to compare. An
ordinary edit gains one list reservation in the cache engine and, only with a
foreign owner, the reach check. The probes run on synthetic fixtures; I did not
measure the native workbook, which has no earlier leaf cost either.

## 8. Gaps

- `uuid_references` of new and replaced category formulas (§3.2), pending §6.
- Value over a pivot body cell stays refused (§3.3).
- The pivot read limits of §3.4.
- `Count`, `CountAll` and `CountUnique` were not distinguished by a computed
  result (§2): the source column holds only distinct numbers.
- The second native file and the visual check of row W1-F (§5).
- A synthetic pivot fixture remains structurally unlike Numbers' (see
  `numbers_pivot_cases.rs`); its parity cases prove code paths, not native shape.

## 9. Verification

The two new fixtures join the corpus walks, so these change with no behaviour
change on any existing file: `corpus_fidelity`'s report gains two rows,
`query_corpus` re-blesses its 13 Numbers digests and no other format's, and the
topology corpus test admits the typed `Unsupported::Pivot`. `exit_parity` pins the
volatile fixture as a third file where the host's owner identifier collides with a
map-only entry. `numbers_native_conditional_highlight_volatile_host_fails_leaf_registers`
replaces the test that pinned the refusal.

Gates ran on this branch at trunk `52f3d60714`, which had not moved. The first
pass used the prescribed `TMPDIR`. `/tmp` hit its quota, so its doctest and pyo3
failures (`ld` bus errors, `Disk quota exceeded`) were reruns with `TMPDIR` under
`.coord/iwa-w5n/tmp`.

| Gate | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| clippy, changed crates + litchi-ops, litchi, litchi-py, all targets and features | 0 |
| clippy, litchi-ops no features, numbers, pages, keynote | 0 |
| tests: litchi-iwa-table, litchi-keynote, litchi-iwa-protos, litchi-numbers-wire, litchi-numbers, litchi-pages, litchi-iwa-style | 0 |
| tests: litchi, all features | 0 (rerun) |
| tests: litchi-ops lib, doc (rerun), every `--test` chunk (two after the bless) | 0 |
| `cargo test -p litchi-iwa --all-features --test parity` | 0: 1191 passed, 210 ignored |
| `cargo test -p litchi-iwa --all-features --lib` | 0 (after the `exit_parity` pin) |
| `cargo doc` with `-Dwarnings`, boundaries, `non_iwork_gate`, glossary, ledger, tool unit tests, capability delta, corpus, feature matrix, fuzz locks, perf-baseline | 0 |
| `tools/unified_bench_matrix.py --check` | 0 after regenerating for this record |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (rerun) |
