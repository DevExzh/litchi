# 0376a — W1-F follow-up: batches, cross-table refresh, formula authoring and removal in Keynote and Pages

- Status: landed on branch `uo2/iwa-w1f` (trunk `3f1820147e`, the 0376
  merge, as the base; trunk `d510b2e961` merged).
- Date: 2026-10-26.
- Input: the coordinator's 0376a list. It asked for the batch ledger rows;
  dependents in other tables if the host refreshes them; segmented formula
  lists if the host handles them; unit tests for `formula_cache::refresh`;
  formula-cell overwrite, clear and formula authoring, tested against the
  host where it works and queued for macOS; and measurements of the
  formula-free Keynote and Pages write routes.
- Base record: [0376](0376-iwa-w1f-formula-cache-engine.md).
- `performance_claim`: the formula-free Keynote and Pages writes stay
  within 0.5% of their pre-0376 cost. The authoring routes run 75%
  (Keynote) and 96% (Pages) below the host routes (§5).

## 1. What landed

| Commit | Crate | What |
|---|---|---|
| `518393ef45` | `litchi-numbers-wire` | `refresh`: dependents in every table, segmented formula lists, table UUID references; refresh unit tests |
| `68ee3db6e0` | `litchi-keynote`, `litchi-pages` | Keynote plans each write over the staged state (batches); caches apply in the tile's own table |
| `00d2a8657f` | `litchi-iwa` (tests), ledger | batch and cross-table refresh parity; `set_slide_table_cells` migrated |
| `cef7441e85` | `litchi-numbers-wire` | `formula_cache::author` (authoring and removal), mirrored merge-owner ranges, `table_write::{formula_free, with_formula}`; author unit tests |
| `f4d95be2d4` | `litchi-keynote` | `slide::drawable::Edit::set_table_formula_cached`; a value write over a formula cell removes the formula |
| `ac539e63c7` | `litchi-pages` | `table::values::Edit::set_formula_cached`, `table::formula`; the same removal; formula edits publish through `structure::Working` |
| `13e1c1ba68` | `litchi-iwa` (tests), ledger, queue | authoring, removal and read parity; 6 rows migrated; 2 queue rows |
| `6f182021d7` | — | trunk merge (queue rows from both sides kept) |
| this commit | docs | this record |

## 2. What the host does, and what the leaf now does

| Item | Host (oracle) | Leaf, 0376 → 0376a |
|---|---|---|
| batch writes (`set_slide_table_cells`) | one refresh after all writes | Keynote ran each operation's refresh over the source and ignored earlier writes of the batch; now each refresh reads the workspace's staged state. Pages already planned one refresh per batch |
| dependents in other tables | refreshed: `DependencyGraph` spans every table, and `refresh_formula_hosts` writes the cache in the dependent's table | refused before (`ExternalOwner`); `refresh` now collects the formula cells, list entries and values of every formula-owning table, and `TileCaches` carries its table's column count |
| segmented formula list | followed by the table-data-list loader | refused before (`Unsupported::SegmentedFormulaList`, now removed); segments are resolved and handed to `collect_formula_list` |
| owner UUID references (`uuid_references`, the cross-table reads of `Table 1::B2`) | used | the planner refuses them. `refresh` admits them when every reference names a registered table by owner UUID, and plans over a copy of the owner without that field (it publishes no owner) |
| formula authoring (`set_slide_table_formula`, `set_table_formula`) | works on every fixture | new `set_table_formula_cached` / `set_formula_cached` |
| value write over a formula cell, clear | works on host-authored formulas; on the native fixtures it fails with "expected to remove one protobuf field 11" | now removes the formula, except in a commented cell (the shared writer's existing refusal) and the Pages mirrored-range case (§3) |

`formula_cache::author::plan` follows the host's dependency storage
exactly:

- A record goes inline when the owner has inline records, and also into
  the cell-record tile of its origin (128 rows × 32 columns) when the owner
  has a tiled envelope, the way native files mirror records.
- A missing tile is created with the owner's next identifier. The owner
  references it and the registry gives it a UUID: deterministic in the
  leaf, random in the host, which is the harness's documented UUID policy.
- A removal that empties a tile deletes the tile, its UUID and the owner's
  reference; the watermark stays, as the host keeps it.
- An envelope the shared rewrite empties is kept, empty.
- Releasing the formula-list key just below `nextListID` lowers
  `nextListID`.

Authoring keeps the supplied cache and refreshes nothing, as the host does.
A value write over a formula first removes the formula, then refreshes the
formulas that read the cell over that state.

Public API:

- Keynote: `slide::drawable::Edit::set_table_formula_cached(slide, drawable,
  row, column, FormulaExpression, FormulaCachedValue)`.
- Pages: `table::values::Edit::set_formula_cached(row, column,
  FormulaExpression, FormulaCachedValue)`, and `table::formula` re-exports
  the formula vocabulary.

The name follows Numbers' `set_formula_cached`. A clear is a value write
(`Empty`). The expression may reference only its own table: the
`TableCell`/`TableRange`/`TableRows`/`TableColumns` variants name the other
table by native identifier, so they are refused (§6), and so is a pivot
category. Date and duration caches are refused.

## 3. Parity

All cases are in `crates/litchi-iwa/tests/parity/formula_refresh_cases.rs`.

| Test | Route | Tier | Result |
|---|---|---|---|
| `keynote_slide_table_formula_refresh_batch_parity` | C3 = false, B2 = 30 in one edit, against `set_slide_table_cells` | B | pass; D2's cache changes |
| `pages_body_table_formula_refresh_cross_table_parity` | `body-table-cross-reference-native.pages` table 1 B2 = 41: refreshes table 1 B4 = SUM(B2:B3) and table 2 A3 = Table 1::B2 + 1 | B | pass; both caches change |
| `keynote_slide_table_formula_author_parity` | `table-discovery.key` C3 = B3 + 1 (the owner has no tile at the origin, so a tile is created) | B | pass |
| `keynote_slide_table_formula_author_graph_parity` | `slide-table-read-native.key` C3, beside the native D2 formula and a merge owner with mirrored ranges | B | pass |
| `keynote_slide_table_formula_clear_parity` | clear the host-authored C3 (the tile empties and is deleted) | B | pass |
| `keynote_slide_table_formula_replace_parity` | the host-authored C3 = 5 | B | pass |
| `keynote_slide_table_formula_read_parity` | native D2 (receipt) and host-authored C3 read as their source | A | pass |
| `pages_body_table_formula_author_parity` | `body-table-catalog-native.pages` C3 = B3 + 1 (a new tile) | B | pass |
| `pages_body_table_formula_clear_parity`, `pages_body_table_formula_replace_parity` | the host-authored C3 | B | pass |
| `pages_body_table_formula_read_parity` | native B4 (receipt) and host-authored C3 | A | pass |
| `keynote_native_commented_formula_clear_is_refused` | the native D2 carries a comment: host refuses (field 11), leaf refuses `TableStorage` | — | pinned |
| `pages_body_table_formula_author_graph_parity` | `body-table-cells-native.pages` C3, in the table whose `SUM(B2:B3)` range is mirrored inline and in a range tile (§4) | B | pass |

`known_fields` registers the TSCE schemas 4000, 4008, 4009 and 4010 for
every format: they belong to the calculation engine all three applications
share.

The ledger migrates these rows:

- Keynote `set_slide_table_cells`, `set_slide_table_formula`,
  `clear_slide_table_formula` and `slide_table_formula` (native receipt).
- Pages `set_table_formula`, `clear_table_formula` and `table_formula`
  (native receipt).

The three Pages cell rows W1-B migrated (`set_table_cell`, `set_table_cells`
and `clear_table_cell`, tier A) keep W1-B's base routes, as the coordinator
resolved them.

The macOS queue gains two W1-F rows, one Keynote and one Pages, for
authoring and removal. `iwa_exit_ledger.py --check` passes.

## 4. Numbers

`formula_dependencies` now admits a range that a non-table owner (a
table's merge owner) mirrors inline and in a range-precedents tile: two
different source records, both containers preserved. Numbers refused such
graphs before, so this is a capability the Numbers formula path gains
where a document has such a merge owner. When the selected table mirrors its
range records (the inline copy equals the tiled facts), the inline copy is
set aside at decode and regenerated beside the rewritten tiles; any other
repeated range stays refused. The range-tile verifier now reads an omitted
size dimension as one, as the source decode does. Numbers refused both
shapes before, so these are capabilities the Numbers formula path gains;
every `litchi-numbers` test passes unchanged (1,025).

Follow-up `948d0010e6`: the Pages SUM-table
authoring case, refused in the first 0376a slice, now passes tier B.

## 5. Performance

Callgrind, release with debug information, CPU 3 pinned, Ir.

| Route | Before / host | After / leaf | Δ |
|---|---:|---:|---:|
| formula-free Keynote writes, `table-discovery.key` (B2 = 7, A1 = "probe", clear A2) ×2, ABBA; before = trunk `f8ca3ee82d` (pre-0376) | 131,881,585 | 132,436,328 | +0.42% |
| formula-free Pages writes, `body-table-catalog-native.pages` (D2 = "Zürich", B2 = 7, clear A2) ×2, ABBA | 162,874,476 | 163,235,750 | +0.22% |
| Keynote authoring C3 = B3 + 1, `table-discovery.key` ×2: host `set_slide_table_formula` vs leaf `set_table_formula_cached` | 1,196,140,293 | 302,729,314 | **−74.7%** |
| Pages authoring C3 = B3 + 1, `body-table-catalog-native.pages` ×2: host `set_table_formula` vs leaf `set_formula_cached` | 3,532,195,968 | 134,221,222 | **−96.2%** |

The formula-free cost now includes locating the cell's tile twice and
checking whether the cell holds a formula. The probe that measured it was a
scratch integration test built at both commits; it is not committed. The
authoring probe is committed:

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --no-run
taskset -c 3 valgrind --tool=callgrind \
  '--toggle-collect=*formula_refresh_cases::host_author_keynote*' \
  target/release/deps/parity-<hash> --ignored --exact \
  formula_refresh_cases::formula_author_cost_abba --test-threads=1
# likewise leaf_author_keynote, host_author_pages, leaf_author_pages
```

## 6. Gaps

| Gap | Note |
|---|---|
| Cross-table references in an authored formula | `FormulaExpression::TableCell { table_id: u64 }` names a native object; a selector-first table reference (as in Numbers' `formula::Table`) belongs in the shared vocabulary first (W1-A/W1-D) |
| Commented formula cells | the shared cell writer's existing refusal (`TableStorage`/`UnsupportedStorage`) |
| Date and duration caches of an authored formula | refused at staging (`CellInput` has no date or duration) |
| Dependents of an authored formula | not refreshed, as in the host; the Numbers owner validates and refreshes them |

## 7. Verification

See the final report; gate commands as in record 0376 §7.
